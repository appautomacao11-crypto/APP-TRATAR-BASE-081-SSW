# -*- coding: utf-8 -*-
"""
app.py - ROBO SSW 081 WEB
Interface Streamlit em cima do processa_sswweb.py original (a logica de
classificacao de rota, deduplicacao e geracao de Excel/PDF NAO foi
reescrita - so foi encapsulada pra funcionar com upload/download web em
vez de pastas locais + input() de terminal.
"""

import os
import io
import html
import tempfile
import datetime
import threading
import uuid
from zoneinfo import ZoneInfo
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import streamlit as st

import processa_sswweb as robo

# layout="wide" (era "centered") - o modo centralizado deixava uma faixa
# vazia grande dos dois lados (só mostrando o wallpaper de fundo por trás),
# apertando os cards de resumo no meio. Com "wide" o conteúdo usa a largura
# inteira da tela, sobrando mais espaço pros cards ficarem confortáveis.
# Pedido do Samuel em 28/08/2026 (mandou print marcando as faixas vazias).
st.set_page_config(page_title="TRATAR BASE 081 - SSW", page_icon="📦", layout="wide")

GITHUB_REPO = "appautomacao11-crypto/APP-TRATAR-BASE-081-SSW""

# ------------------------------------------------------------------
# Fuso horário do Brasil (RN) - o servidor do Streamlit Cloud roda em
# UTC, 3 horas à frente do horário de Natal/RN (sem horário de verão).
# Pedido do Samuel em 19/08/2026, depois de notar o horário "Processado
# em..." sempre 3h adiantado.
# ------------------------------------------------------------------
FUSO_RN = ZoneInfo("America/Fortaleza")


def agora_rn():
    return datetime.datetime.now(FUSO_RN)


# ------------------------------------------------------------------
# ESTADO DA SESSÃO
# ------------------------------------------------------------------
if "usuario" not in st.session_state:
    st.session_state.usuario = None
if "parceiro" not in st.session_state:
    st.session_state.parceiro = None
if "resultado" not in st.session_state:
    st.session_state.resultado = None  # guarda os arquivos gerados pra essa sessão
if "base_rotas_bytes" not in st.session_state:
    st.session_state.base_rotas_bytes = None  # bytes da base vinculada ao usuario logado
if "base_rotas_verificada" not in st.session_state:
    st.session_state.base_rotas_verificada = False  # se ja checou no GitHub nessa sessao
if "sessao_id" not in st.session_state:
    st.session_state.sessao_id = str(uuid.uuid4())  # identifica essa aba/navegador especifico


# ------------------------------------------------------------------
# CONTROLE DE SESSÃO ÚNICA POR CÓDIGO - impede que o mesmo código de
# acesso seja usado em duas abas/dispositivos ao mesmo tempo (evita
# sobrecarga de processamento simultâneo e "brigas" pela mesma base de
# rotas). Usa memória compartilhada entre todas as sessões desse
# processo (via st.cache_resource) - guarda só qual sessao esta usando
# cada código e quando foi a última atividade dela. Se uma sessão ficar
# parada por tempo demais (aba fechada sem clicar em Sair, por
# exemplo), a trava expira sozinha e libera o código de novo. Pedido do
# Samuel em 21/08/2026.
# ------------------------------------------------------------------
TIMEOUT_SESSAO_INATIVA_MIN = 20


@st.cache_resource
def _controle_sessoes():
    return {"lock": threading.Lock(), "ativas": {}}


def tenta_reservar_codigo(codigo):
    """Tenta reservar esse código de login pra sessão atual (aba). Se já
    estiver em uso por outra sessão e ainda dentro do prazo de
    atividade, devolve (False, minutos_parado_da_outra_sessao). Se
    conseguiu reservar (livre, ou a trava antiga expirou), devolve
    (True, None)."""
    controle = _controle_sessoes()
    agora = agora_rn()
    with controle["lock"]:
        info = controle["ativas"].get(codigo)
        if info is not None and info["sessao_id"] != st.session_state.sessao_id:
            minutos_parado = (agora - info["ultima_atividade"]).total_seconds() / 60
            if minutos_parado < TIMEOUT_SESSAO_INATIVA_MIN:
                return False, minutos_parado
        controle["ativas"][codigo] = {"sessao_id": st.session_state.sessao_id, "ultima_atividade": agora}
        return True, None


def atualiza_atividade_codigo(codigo):
    """Renova o carimbo de atividade do código pra essa sessão - chama a
    cada tela carregada enquanto o usuário estiver logado, pra a trava
    não expirar sozinha enquanto ele ainda está usando de verdade."""
    controle = _controle_sessoes()
    with controle["lock"]:
        info = controle["ativas"].get(codigo)
        if info and info["sessao_id"] == st.session_state.sessao_id:
            info["ultima_atividade"] = agora_rn()


def libera_codigo(codigo, forcar=False):
    """Libera a trava desse código. Por padrão só libera se for a
    própria sessão dona da trava (logout normal); com forcar=True
    (usado só pelo admin) libera de qualquer jeito."""
    controle = _controle_sessoes()
    with controle["lock"]:
        info = controle["ativas"].get(codigo)
        if info and (forcar or info["sessao_id"] == st.session_state.sessao_id):
            del controle["ativas"][codigo]


def sair():
    """Apaga tudo que estava em memória dessa sessão (arquivos, prévia,
    e qualquer resquicio de sessao de admin), e libera a trava do
    código pra outro dispositivo poder usar."""
    if st.session_state.usuario:
        libera_codigo(st.session_state.usuario)
    st.session_state.usuario = None
    st.session_state.parceiro = None
    st.session_state.resultado = None
    st.session_state.base_rotas_bytes = None
    st.session_state.base_rotas_verificada = False
    st.session_state.pop("senha_admin", None)
    st.session_state.pop("_admin_login_registrado", None)
    st.rerun()


# ------------------------------------------------------------------
# Papéis de parede (marca d'água de fundo). ANTES do login, usa sempre
# o NEUTRO (só marca Dominalog, sem citar parceiro nenhum). DEPOIS do
# login, troca pro wallpaper do parceiro vinculado aquele código de
# acesso especifico (ver usuarios_autorizados.json) - assim quem usa um
# código do parceiro A nunca fica sabendo que existe um parceiro B.
# Pedido do Samuel em 18-20/08/2026.
#
# GENERALIZADO em 01/10/2026: o RN hoje tem 2 parceiros conhecidos
# (AVEX e Te Levo), mas outras unidades (filiais fora do RN, ver
# detecta_uf_predominante) podem vir a ter parceiros completamente
# diferentes, sem dar pra prever o nome deles com antecedência. Em vez
# de exigir uma alteração de CÓDIGO (editar esse dict + redeploy) toda
# vez que uma unidade nova ganha um parceiro, o parceiro agora é texto
# livre digitado pelo admin (ver tela_login, painel do admin) e o
# arquivo de wallpaper correspondente é PROCURADO por convenção de
# nome: "wallpaper_<nome-do-parceiro-em-slug>.png" (ver slug_texto).
# Sem nenhum parceiro configurado ainda pra uma unidade nova = sem
# arquivo correspondente ainda = cai automaticamente no neutro, sem
# erro - é por isso que dá pra digitar o nome do parceiro na hora de
# cadastrar o usuário MESMO ANTES de ter a imagem pronta.
#
# PARCEIROS_LEGADO existe só pra manter os 2 parceiros do RN
# funcionando exatamente como já funcionavam (arquivos
# "wallpaper_rn1_avex.png"/"wallpaper_rn4_televo.png", que não seguem
# o padrão novo de nome) - ninguém precisa renomear nada que já existe
# no repositório. Pra um parceiro NOVO (de outra unidade), o convencional
# e simples é digitar o nome dele (ex: "XPTO") e subir um arquivo
# "wallpaper_xpto.png" - ver o guia "Como adicionar um parceiro novo".
# ------------------------------------------------------------------
WALLPAPER_NEUTRO = "wallpaper_neutro.png"
PARCEIRO_PADRAO = "neutro"
PARCEIROS_LEGADO = {
    "rn1": {"arquivo": "wallpaper_rn1_avex.png", "rotulo": "RN1 · AVEX"},
    "rn4": {"arquivo": "wallpaper_rn4_televo.png", "rotulo": "RN4 · Te Levo"},
}


def slug_texto(texto):
    """Mesma normalização usada em slug_usuario (só letras/números/
    underscore, minúsculo) - reaproveitada aqui pra achar o arquivo de
    wallpaper de um parceiro a partir do nome digitado pelo admin, e
    por slug_usuario (abaixo) pra nome de arquivo de base de rotas."""
    import re as _re
    limpo = _re.sub(r"[^A-Za-z0-9]+", "_", (texto or "").strip())
    return limpo.strip("_").lower()


def nome_arquivo_wallpaper(parceiro):
    """Decide qual arquivo de wallpaper usar pra esse parceiro:
    1. Sem parceiro (ou "neutro") -> sempre o neutro.
    2. Parceiro legado do RN ("rn1"/"rn4") -> arquivo fixo de sempre,
       sem passar pelo slug (nomes antigos não seguem o padrão novo).
    3. Qualquer outro texto -> procura "wallpaper_<slug>.png"; se esse
       arquivo ainda não existir no projeto (parceiro cadastrado mas
       template ainda não subido), cai no neutro sem erro."""
    if not parceiro or parceiro == PARCEIRO_PADRAO:
        return WALLPAPER_NEUTRO
    if parceiro in PARCEIROS_LEGADO:
        return PARCEIROS_LEGADO[parceiro]["arquivo"]
    candidato = f"wallpaper_{slug_texto(parceiro)}.png"
    caminho_candidato = os.path.join(os.path.dirname(__file__), candidato)
    if os.path.isfile(caminho_candidato):
        return candidato
    return WALLPAPER_NEUTRO


def rotulo_parceiro(parceiro):
    """Texto exibido pro admin na lista de usuários cadastrados."""
    if not parceiro or parceiro == PARCEIRO_PADRAO:
        return "Dominalog (neutro)"
    if parceiro in PARCEIROS_LEGADO:
        return PARCEIROS_LEGADO[parceiro]["rotulo"]
    return parceiro


def aplica_wallpaper():
    if st.session_state.usuario:
        nome_arquivo = nome_arquivo_wallpaper(st.session_state.parceiro)
    else:
        nome_arquivo = WALLPAPER_NEUTRO

    caminho = os.path.join(os.path.dirname(__file__), nome_arquivo)
    if not os.path.isfile(caminho):
        return
    import base64
    with open(caminho, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    st.markdown(
        f"""
        <style>
        /* 5a rodada (28/08/2026): o pedido era as telas de DEPOIS do login
        (wallpaper do parceiro, ex: wallpaper_rn4_televo.png) ficarem com a
        MESMA cor "viva" da tela de login (wallpaper NEUTRO,
        wallpaper_neutro.png). Tentei isolar o filtro só no fundo usando um
        .stApp::before com position:fixed + z-index:-1 - resultado: quebrou
        de vez, a caixa de login sumiu (Samuel testou ao vivo em 29/08/2026).
        Motivo provavel: .stApp com position:relative mas SEM z-index proprio
        nao cria um novo contexto de empilhamento, entao o z-index:-1 do
        ::before nao fica garantido "atras só do proprio .stApp" - em vez
        disso ele disputa posição com toda a pagina, e dependendo do
        navegador pode empurrar (ou até sobrepor) conteúdo que deveria estar
        na frente, incluindo a caixa de login. É um problema conhecido desse
        truque (position:fixed + z-index negativo sem contexto de
        empilhamento isolado) - por isso foi revertido.
        6a rodada (atual, 29/08/2026): voltou pro jeito simples e já
        comprovado nas rodadas 1-4 (background-image direto no .stApp, sem
        pseudo-elemento, sem position:fixed, sem z-index). O filtro de
        saturação/brilho pra aproximar a cor do wallpaper do parceiro da cor
        viva do wallpaper neutro foi aplicado direto no .stApp tambem -
        ressalva: como filter numa camada de fundo tambem "vaza" pro
        conteudo em cima dela (texto, caixa), o efeito fica mais sutil aqui
        (valores mais baixos que a tentativa anterior) pra nao lavar o
        texto. Se ficar fraco demais, dá pra subir aos poucos - mas sem
        repetir o truque de isolar via ::before/fixed, que foi o que
        quebrou. */
        .stApp {{
            background-image: url("data:image/png;base64,{b64}");
            background-size: cover;
            background-position: center top;
            background-repeat: no-repeat;
            filter: saturate(1.25) brightness(1.08);
        }}
        [data-testid="stHeader"] {{
            background: rgba(0,0,0,0);
        }}
        /* Preenchimento leve atras do conteudo - com layout="wide" o texto
        passou a ficar em cima de partes mais "cheias" do wallpaper (o
        caminhao/navio), dificultando a leitura. Um fundo navy translucido
        (mesma paleta da marca) da contraste sem esconder o wallpaper -
        ainda da pra ver o papel de parede por tras, só mais suave. Pedido
        do Samuel em 28/08/2026.
        Historico de ajustes (mesmo dia, testado num monitor de verdade a
        cada rodada, nao print de navegador):
          1a versao: rgba(10,0,40,0.35) sem borda.
          2a rodada: subiu pra 0.55 + borda laranja - resolveu a caixa
            sumir no wallpaper NEUTRO (tela de login, sem foto), mas
            escondeu demais o wallpaper com foto nas telas "Enviar
            arquivo"/"Prévia" (essas usam o wallpaper do parceiro, com
            caminhao/navio/mapa).
          3a rodada: desceu pra 0.20 - a tela de login ficou boa (o
            texto/logo dela cai sobre uma parte mais "lisa" do wallpaper
            neutro), mas "Enviar arquivo" e "Prévia" continuaram sem o
            "pop" da login - o conteudo delas cai em cima da parte mais
            cheia da FOTO (caminhao/navio/setas), entao com pouco
            preenchimento a caixa quase nao se destaca do fundo ali.
          4a rodada: 0.20 -> 0.30 + text-shadow no texto - Samuel esclareceu
            que o problema nao era legibilidade, era a COR mais apagada que
            a login.
          5a rodada (atual): saturate/brightness/contrast só no fundo (ver
            acima) pra aproximar a cor do wallpaper do parceiro da cor viva
            do wallpaper neutro; caixa voltou pra 0.22 (entre a 3a e a 4a)
            já que quem carrega o "pop" de cor agora é o filtro no fundo,
            não o preenchimento da caixa. */
        .block-container {{
            background-color: rgba(24, 12, 58, 0.22);
            border: 1px solid rgba(255, 140, 40, 0.35);
            box-shadow: 0 0 28px rgba(255, 140, 40, 0.10);
            border-radius: 16px;
            padding: 2rem 2.5rem;
        }}
        /* Sombra no texto (nao na caixa) pra garantir leitura em cima das
        partes mais "cheias" do wallpaper com foto, sem precisar esconder
        o wallpaper subindo o alpha da caixa. Acrescentado na 4a rodada,
        28/08/2026. */
        .block-container h1, .block-container h2, .block-container h3,
        .block-container p, .block-container span, .block-container label,
        [data-testid="stMetricValue"], [data-testid="stMetricLabel"] {{
            text-shadow: 0 1px 6px rgba(0, 0, 0, 0.65);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

aplica_wallpaper()


def aplica_largura_estreita():
    """Chamada no INÍCIO das telas que devem ficar centralizadas/estreitas
    (login, escolha de base, upload, download) - sobrescreve só a largura
    do .block-container, que o aplica_wallpaper() acima deixa esticada
    (layout="wide", sem limite de largura). st.set_page_config(layout=...)
    só pode ser chamado 1x no script inteiro, então não dá pra alternar
    entre "wide" e "centered" de verdade por tela - o jeito é manter
    "wide" sempre e, nas telas que não precisam da largura toda, encolher
    só o .block-container de volta pro tamanho do modo centralizado
    (736px, é o valor padrão do próprio Streamlit). A tela "2. Prévia
    antes de gerar" (tela_previa) NÃO chama essa função, então continua
    esticada - foi a única tela que o Samuel disse que gostou mais larga.
    Pedido do Samuel em 28/08/2026 (mandou prints das 4 telas: gostou da
    prévia larga, mas achou as outras 3 "muito para cima"/vazias)."""
    st.markdown(
        """
        <style>
        .block-container {
            max-width: 736px;
            margin-left: auto;
            margin-right: auto;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------------
# Caminho da base_rotas.xlsx "padrão" (se você colocar o arquivo dentro
# da pasta do projeto com esse nome, o usuário não precisa subir toda
# vez - só quando ela mudar).
# ------------------------------------------------------------------
CAMINHO_BASE_ROTAS_XLSX = os.path.join(os.path.dirname(__file__), "base_rotas.xlsx")
CAMINHO_BASE_ROTAS_CSV = os.path.join(os.path.dirname(__file__), "base_rotas.csv")


def resolve_base_rotas_padrao(pasta_tmp):
    """Retorna o caminho de um base_rotas.xlsx pronto pra uso. Se só
    existir a versão .csv no projeto (upload de .xlsx deu problema no
    GitHub mobile), converte pra um .xlsx temporário na hora."""
    if os.path.isfile(CAMINHO_BASE_ROTAS_XLSX):
        return CAMINHO_BASE_ROTAS_XLSX
    if os.path.isfile(CAMINHO_BASE_ROTAS_CSV):
        df = pd.read_csv(CAMINHO_BASE_ROTAS_CSV)
        caminho_convertido = os.path.join(pasta_tmp, "base_rotas.xlsx")
        df.to_excel(caminho_convertido, index=False, sheet_name="Sheet1")
        return caminho_convertido
    return None


def obter_bytes_base_padrao():
    """Devolve os BYTES da base_rotas padrão do projeto, prontos pra
    salvar como cópia vinculada a um usuário. Se só existir a versão
    .csv, converte pra .xlsx em memória. Devolve None se não existir
    nenhuma base padrão no projeto (nesse caso a opção 'usar padrão' não
    deve aparecer na tela de escolha)."""
    if os.path.isfile(CAMINHO_BASE_ROTAS_XLSX):
        with open(CAMINHO_BASE_ROTAS_XLSX, "rb") as f:
            return f.read()
    if os.path.isfile(CAMINHO_BASE_ROTAS_CSV):
        df = pd.read_csv(CAMINHO_BASE_ROTAS_CSV)
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, sheet_name="Sheet1")
        return buffer.getvalue()
    return None


# ------------------------------------------------------------------
# MODO "SOMENTE SETOR" - pra quem não tem base de rotas (cidade/bairro/
# CEP) cadastrada. Em vez de inventar um jeito separado de marcar isso
# no session_state, reaproveita o MESMO mecanismo de persistência que já
# existe pra base própria (salva_base_rotas_usuario/
# carrega_base_rotas_usuario, arquivo bases_usuarios/base_rotas_<user>.xlsx
# no GitHub) - só que em vez do conteúdo real da base, salva um arquivo
# "sentinela" (6 colunas vazias, pra não quebrar ler_base_rotas, mais uma
# aba extra só com um marcador). Vantagem: nenhuma das funções de
# persistência/admin (salvar, carregar, desvincular, listar quem tem
# base vinculada) precisa mudar - continuam enxergando "esse usuário tem
# um arquivo vinculado" normalmente, só que processa_sswweb.py recebe
# forcar_modo_setor=True e ignora o conteúdo (vazio mesmo) desse arquivo.
# Pedido do Samuel em 01/10/2026 (repasse do projeto pra um colega sem
# base de rotas pronta).
# ------------------------------------------------------------------
MARCADOR_MODO_SETOR = "MODO_SETOR_SEM_BASE"


def monta_base_setor_sentinela():
    """Monta um 'base_rotas.xlsx' vazio/sentinela pro modo somente Setor
    - ver comentário acima. A 1ª aba tem as 6 colunas que ler_base_rotas
    espera (sem nenhuma linha de dado - nunca é usada de verdade pra
    classificar rota nesse modo); a 2ª aba só existe pra
    eh_modo_setor_ativo() reconhecer esse arquivo depois."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "base_rotas"
    ws.append(["Cidade ou Bairro", "Setor Operacional", "Mesorregião", "UF", "CEP Inicial", "CEP Final"])
    ws_marcador = wb.create_sheet(MARCADOR_MODO_SETOR)
    ws_marcador["A1"] = MARCADOR_MODO_SETOR
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def eh_modo_setor_ativo(base_rotas_bytes):
    """Confere se os bytes vinculados ao usuário são o sentinela do
    'modo somente Setor' (ver monta_base_setor_sentinela), em vez de uma
    base de rotas de verdade."""
    if not base_rotas_bytes:
        return False
    try:
        wb = openpyxl.load_workbook(io.BytesIO(base_rotas_bytes), read_only=True)
        return MARCADOR_MODO_SETOR in wb.sheetnames
    except Exception:
        return False


# ------------------------------------------------------------------
# USUÁRIOS AUTORIZADOS - guardados num arquivo dentro do próprio
# repositório do GitHub (usuarios_autorizados.json), lido e escrito via
# API do GitHub. Cada usuário tem um CÓDIGO e um PARCEIRO vinculado
# (usado pra decidir qual wallpaper mostrar pra ele). Isso permite que o
# painel de admin (dentro do app) adicione/remova gente sem precisar
# editar nada manualmente depois.
#
# Secrets necessários (Settings > Secrets no Streamlit Cloud):
#   admin_senha   = "sua senha de administrador"
#   github_token  = "ghp_xxxxxxxxxxxxxxxxxxxx"   (Personal Access Token,
#                                                  escopo "repo")
# ------------------------------------------------------------------
ARQ_USUARIOS = "usuarios_autorizados.json"


def carrega_usuarios():
    """Lê a lista atual de usuários (código + parceiro vinculado) via
    API autenticada do GitHub. Cada item é {"codigo": ..., "parceiro":
    "neutro"|"rn1"|"rn4"}. Aceita o formato antigo também (lista simples
    de strings em "codigos"), tratando como parceiro padrão - assim não
    quebra pra quem já tinha códigos cadastrados antes dessa mudança."""
    import requests, base64, json as _json

    token = st.secrets.get("github_token")
    if not token:
        return []
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQ_USUARIOS}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    try:
        r = requests.get(url_api, headers=headers, timeout=8)
        if r.status_code == 200:
            conteudo_b64 = r.json()["content"]
            conteudo = base64.b64decode(conteudo_b64).decode("utf-8")
            dados = _json.loads(conteudo)
            if "usuarios" in dados:
                return dados["usuarios"]
            # formato antigo: {"codigos": ["a", "b"]}
            return [{"codigo": c, "parceiro": PARCEIRO_PADRAO} for c in dados.get("codigos", [])]
    except Exception:
        pass
    return []


def salva_usuarios(nova_lista, mensagem_commit):
    """Escreve a lista de usuários (código + parceiro) no arquivo do
    GitHub via API (precisa do github_token nos Secrets, com permissão
    de escrita no repositório)."""
    import requests, base64, json as _json

    token = st.secrets.get("github_token")
    if not token:
        raise RuntimeError("github_token não configurado nos Secrets.")

    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQ_USUARIOS}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}

    sha_atual = None
    r = requests.get(url_api, headers=headers, timeout=8)
    if r.status_code == 200:
        sha_atual = r.json()["sha"]

    conteudo = _json.dumps({"usuarios": nova_lista}, indent=2, ensure_ascii=False)
    payload = {
        "message": mensagem_commit,
        "content": base64.b64encode(conteudo.encode("utf-8")).decode("utf-8"),
    }
    if sha_atual:
        payload["sha"] = sha_atual

    r = requests.put(url_api, headers=headers, json=payload, timeout=10)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"GitHub recusou a atualização ({r.status_code}): {r.text[:200]}")


# ------------------------------------------------------------------
# LOG DE ACESSOS - guarda QUEM entrou (código de usuário ou "ADMIN") e
# QUANDO, num arquivo "log_acessos.jsonl" no repositório (1 linha = 1
# acesso, em JSON). Só serve pra rastreabilidade - não guarda senha,
# nem nenhum dado além do código de login e do horário. Pensado pra dar
# uma base de auditoria simples antes de eventualmente guardar mais
# coisa no sistema. Pedido do Samuel em 21/08/2026.
# ------------------------------------------------------------------
ARQ_LOG_ACESSOS = "log_acessos.jsonl"


def registra_acesso(usuario, tipo="login"):
    """Adiciona uma linha ao log de acessos. Propositalmente NUNCA
    lança erro pra fora - se o log falhar (rede, conflito etc.), o
    login continua funcionando normalmente mesmo assim; só tenta de
    novo algumas vezes se dois acessos colidirem no mesmo instante."""
    import requests, base64, json as _json

    token = st.secrets.get("github_token")
    if not token:
        return
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQ_LOG_ACESSOS}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}

    linha = _json.dumps(
        {"usuario": usuario, "tipo": tipo, "quando": agora_rn().strftime("%Y-%m-%d %H:%M:%S")},
        ensure_ascii=False,
    )

    for _tentativa in range(3):
        try:
            conteudo_atual = ""
            sha_atual = None
            r = requests.get(url_api, headers=headers, timeout=8)
            if r.status_code == 200:
                sha_atual = r.json()["sha"]
                conteudo_atual = base64.b64decode(r.json()["content"]).decode("utf-8")

            novo_conteudo = conteudo_atual
            if novo_conteudo and not novo_conteudo.endswith("\n"):
                novo_conteudo += "\n"
            novo_conteudo += linha + "\n"

            payload = {
                "message": f"Log de acesso: {usuario} ({tipo})",
                "content": base64.b64encode(novo_conteudo.encode("utf-8")).decode("utf-8"),
            }
            if sha_atual:
                payload["sha"] = sha_atual

            r = requests.put(url_api, headers=headers, json=payload, timeout=10)
            if r.status_code in (200, 201):
                return
            if r.status_code == 409:
                continue  # outro acesso salvou entre o GET e o PUT - tenta de novo
            return  # qualquer outro erro - desiste sem travar o login
        except Exception:
            return


def carrega_log_acessos(limite=50):
    """Lê o log de acessos e devolve os últimos 'limite' registros, do
    mais recente pro mais antigo. Devolve lista vazia se ainda não
    existir nenhum acesso registrado ou em caso de erro."""
    import requests, base64, json as _json

    token = st.secrets.get("github_token")
    if not token:
        return []
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{ARQ_LOG_ACESSOS}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    try:
        r = requests.get(url_api, headers=headers, timeout=8)
        if r.status_code != 200:
            return []
        conteudo = base64.b64decode(r.json()["content"]).decode("utf-8")
    except Exception:
        return []

    registros = []
    for linha in conteudo.strip().splitlines():
        linha = linha.strip()
        if not linha:
            continue
        try:
            registros.append(_json.loads(linha))
        except Exception:
            continue
    registros.reverse()
    return registros[:limite]


def codigos_validos():
    return [u["codigo"] for u in carrega_usuarios()]


def parceiro_do_codigo(codigo, usuarios=None):
    for u in (usuarios if usuarios is not None else carrega_usuarios()):
        if u["codigo"] == codigo:
            return u.get("parceiro", PARCEIRO_PADRAO)
    return PARCEIRO_PADRAO


# ------------------------------------------------------------------
# BASE DE ROTAS POR USUÁRIO - guardada dentro do próprio repositório,
# numa subpasta "bases_usuarios/", com nome derivado do código de login
# (ex: 'base_rotas_samuel_dmnz0206.xlsx'). Na primeira vez que um
# usuário loga sem ter uma base vinculada, o app pergunta se ele quer
# anexar uma própria ou usar a padrão - depois disso, nunca mais
# pergunta, usa a vinculada automaticamente. O admin pode desvincular a
# base de um usuário específico pelo painel escondido (?admin=1), o que
# faz a pergunta voltar a aparecer pra ele no próximo login. Pedido do
# Samuel em 19/08/2026.
# ------------------------------------------------------------------
PASTA_BASES_USUARIOS = "bases_usuarios"


def slug_usuario(usuario):
    """Transforma o código de login num nome de arquivo seguro (só
    letras, números e underscore) - ex: 'samuel.dmnz0206' vira
    'samuel_dmnz0206'. Mesma normalização de slug_texto (lá em cima,
    perto do wallpaper) - mantida como função separada só pelo nome,
    pra não mudar nenhuma chamada existente."""
    return slug_texto(usuario)


def carrega_base_rotas_usuario(usuario):
    """Busca no GitHub se existe uma base_rotas.xlsx vinculada a esse
    usuário específico. Retorna os BYTES do arquivo se existir, ou None
    se ainda não tiver (usuário novo, ou nunca vinculou uma base
    própria)."""
    import requests, base64

    token = st.secrets.get("github_token")
    if not token:
        return None
    nome_arquivo = f"base_rotas_{slug_usuario(usuario)}.xlsx"
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PASTA_BASES_USUARIOS}/{nome_arquivo}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    try:
        r = requests.get(url_api, headers=headers, timeout=8)
        if r.status_code == 200:
            conteudo_b64 = r.json()["content"]
            return base64.b64decode(conteudo_b64)
    except Exception:
        pass
    return None


def salva_base_rotas_usuario(usuario, conteudo_bytes, mensagem_commit):
    """Grava (ou sobrescreve) a base_rotas vinculada a esse usuário no
    GitHub, via API (precisa do github_token nos Secrets com permissão
    de escrita)."""
    import requests, base64

    token = st.secrets.get("github_token")
    if not token:
        raise RuntimeError("github_token não configurado nos Secrets.")

    nome_arquivo = f"base_rotas_{slug_usuario(usuario)}.xlsx"
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PASTA_BASES_USUARIOS}/{nome_arquivo}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}

    sha_atual = None
    r = requests.get(url_api, headers=headers, timeout=8)
    if r.status_code == 200:
        sha_atual = r.json()["sha"]

    payload = {
        "message": mensagem_commit,
        "content": base64.b64encode(conteudo_bytes).decode("utf-8"),
    }
    if sha_atual:
        payload["sha"] = sha_atual

    r = requests.put(url_api, headers=headers, json=payload, timeout=15)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"GitHub recusou a atualização ({r.status_code}): {r.text[:200]}")


def lista_usuarios_com_base_vinculada(lista_codigos):
    """Confere, pra cada usuário cadastrado, se ele já tem uma
    base_rotas vinculada - devolve só os que têm. Usado no painel do
    admin pra montar a lista de quem pode ser desvinculado."""
    import requests

    token = st.secrets.get("github_token")
    if not token:
        return []
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PASTA_BASES_USUARIOS}"
    try:
        r = requests.get(url_api, headers=headers, timeout=8)
        if r.status_code != 200:
            return []
        arquivos_existentes = {item["name"] for item in r.json()}
    except Exception:
        return []

    com_base = []
    for codigo in lista_codigos:
        nome_esperado = f"base_rotas_{slug_usuario(codigo)}.xlsx"
        if nome_esperado in arquivos_existentes:
            com_base.append(codigo)
    return com_base


def desvincula_base_rotas_usuario(usuario):
    """Apaga a base_rotas vinculada a esse usuário no GitHub - da
    próxima vez que ele logar, a tela de escolha (padrão vs própria)
    volta a aparecer pra ele. Só o admin tem acesso a essa ação."""
    import requests

    token = st.secrets.get("github_token")
    if not token:
        raise RuntimeError("github_token não configurado nos Secrets.")

    nome_arquivo = f"base_rotas_{slug_usuario(usuario)}.xlsx"
    url_api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PASTA_BASES_USUARIOS}/{nome_arquivo}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}

    r = requests.get(url_api, headers=headers, timeout=8)
    if r.status_code != 200:
        raise RuntimeError(f"Não achei a base desse usuário pra apagar ({r.status_code}).")
    sha_atual = r.json()["sha"]

    payload = {
        "message": f"Desvincula base de rotas do usuario {usuario}",
        "sha": sha_atual,
    }
    r = requests.delete(url_api, headers=headers, json=payload, timeout=10)
    if r.status_code not in (200, 204):
        raise RuntimeError(f"GitHub recusou apagar ({r.status_code}): {r.text[:200]}")


# ==================================================================
# MIGRAÇÃO TEMPORÁRIA (21/08/2026) - remove qualquer menção a "AVEX"
# da base padrão do projeto e das bases já vinculadas a cada usuário
# (aba LEIA-ME), já que hoje tem mais de uma transportadora parecida e
# o texto antigo (feito quando só existia a AVEX) ficou desatualizado.
# É autocontida (não usa nem é usada por mais nada no resto do app) e
# idempotente: se não achar nenhuma menção pra limpar, não faz nada.
# PODE APAGAR ESSA FUNÇÃO INTEIRA (e a linha que a chama lá embaixo, no
# painel do admin) na próxima rodada de atualizações, assim que rodar
# uma vez e confirmar que não sobrou nenhuma base com "AVEX" - depois
# disso ela só fica girando à toa sem achar nada, sem prejuízo nenhum
# em deixar ou tirar.
# ==================================================================
def roda_migracao_temporaria_avex(codigos_atuais, usuarios_com_base):
    import re, requests, base64

    def limpa(conteudo_bytes):
        wb = openpyxl.load_workbook(io.BytesIO(conteudo_bytes))
        mudou = False
        padrao_avex = re.compile(r"\b(?:a\s+)?AVEX\b", re.IGNORECASE)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    valor = cell.value
                    if not isinstance(valor, str) or "avex" not in valor.lower():
                        continue
                    novo_valor = padrao_avex.sub("a transportadora parceira", valor)
                    if novo_valor != valor:
                        cell.value = novo_valor
                        mudou = True
        if not mudou:
            return None
        saida = io.BytesIO()
        wb.save(saida)
        return saida.getvalue()

    def salva_no_github(url_api, conteudo_bytes, mensagem_commit):
        token = st.secrets.get("github_token")
        if not token:
            return
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
        sha_atual = None
        r = requests.get(url_api, headers=headers, timeout=8)
        if r.status_code == 200:
            sha_atual = r.json()["sha"]
        payload = {"message": mensagem_commit, "content": base64.b64encode(conteudo_bytes).decode("utf-8")}
        if sha_atual:
            payload["sha"] = sha_atual
        requests.put(url_api, headers=headers, json=payload, timeout=15)

    corrigidos = []

    bytes_padrao = obter_bytes_base_padrao()
    if bytes_padrao:
        novo = limpa(bytes_padrao)
        if novo:
            salva_no_github(
                f"https://api.github.com/repos/{GITHUB_REPO}/contents/base_rotas.xlsx",
                novo, "Remove mencao a AVEX da base padrao (migracao temporaria)",
            )
            corrigidos.append("base padrão do projeto")

    for codigo in usuarios_com_base:
        bytes_usuario = carrega_base_rotas_usuario(codigo)
        if not bytes_usuario:
            continue
        novo = limpa(bytes_usuario)
        if novo:
            salva_no_github(
                f"https://api.github.com/repos/{GITHUB_REPO}/contents/{PASTA_BASES_USUARIOS}/base_rotas_{slug_usuario(codigo)}.xlsx",
                novo, f"Remove mencao a AVEX da base de {codigo} (migracao temporaria)",
            )
            corrigidos.append(f"usuário {codigo}")

    if corrigidos:
        st.caption(f"🧹 Migração temporária: corrigi a menção à AVEX em {len(corrigidos)} base(s) — {', '.join(corrigidos)}.")


def tela_login():
    aplica_largura_estreita()
    st.title("📦 TRATAR BASE 081 - SSW")
    st.caption("Digite seu código de acesso pra processar o romaneio.")
    codigo = st.text_input("Código de acesso", placeholder="nome.empresaXXXX", type="password")
    if st.button("Entrar", type="primary"):
        usuarios = carrega_usuarios()
        validos = [u["codigo"] for u in usuarios]
        if not validos:
            st.error(
                "Nenhum código de acesso cadastrado ainda. Fale com o administrador."
            )
        elif codigo.strip() not in validos:
            st.error("Código de acesso inválido.")
        else:
            codigo_limpo = codigo.strip()
            ok, minutos_parado = tenta_reservar_codigo(codigo_limpo)
            if not ok:
                st.error(
                    f"Esse código já está em uso em outro dispositivo/aba agora "
                    f"(última atividade há {minutos_parado:.0f} min). "
                    f"Só dá pra usar o mesmo código em um lugar por vez - "
                    f"espere a outra sessão encerrar (ou ficar {TIMEOUT_SESSAO_INATIVA_MIN} min "
                    f"parada) e tente de novo."
                )
            else:
                # limpa qualquer resquicio de sessao de admin antes de logar
                # como usuario normal - garante que login normal NUNCA
                # carrega acesso de administrador junto.
                st.session_state.pop("senha_admin", None)
                st.session_state.usuario = codigo_limpo
                st.session_state.parceiro = parceiro_do_codigo(codigo_limpo, usuarios)
                st.session_state.base_rotas_verificada = False
                registra_acesso(codigo_limpo, tipo="login")
                st.rerun()

    # Painel de administrador FICA ESCONDIDO da tela normal - só aparece
    # pra quem acessar com ?admin=1 no final do link (ex:
    # https://seu-app.streamlit.app/?admin=1). Assim usuarios comuns nem
    # sabem que essa area existe. Pedido do Samuel em 18/08/2026.
    if st.query_params.get("admin") == "1":
        with st.expander("🔐 Área do administrador"):
            senha_admin = st.text_input("Senha de administrador", type="password", key="senha_admin")
            if senha_admin and senha_admin == st.secrets.get("admin_senha"):
                st.success("Acesso de administrador liberado.")
                if not st.session_state.get("_admin_login_registrado"):
                    registra_acesso("ADMIN", tipo="admin")
                    st.session_state["_admin_login_registrado"] = True
                atuais = carrega_usuarios()
                codigos_atuais = [u["codigo"] for u in atuais]

                with st.expander("📜 Log de acessos (últimos 50)"):
                    log = carrega_log_acessos(limite=50)
                    if log:
                        st.dataframe(
                            pd.DataFrame(log)[["quando", "usuario", "tipo"]],
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.write("_nenhum acesso registrado ainda_")

                st.write("**Usuários cadastrados atualmente:**")
                if atuais:
                    for u in atuais:
                        st.write(f"- `{u['codigo']}` — {rotulo_parceiro(u.get('parceiro', PARCEIRO_PADRAO))}")
                else:
                    st.write("_nenhum ainda_")

                st.write("---")
                st.write("**Adicionar novo código:**")
                novo = st.text_input("Novo código (ex: nome.empresaXXXX)", key="novo_codigo")
                # Parceiro agora é texto livre (generalizado em 01/10/2026 -
                # ver comentário de nome_arquivo_wallpaper lá em cima) - pra
                # RN, digitar exatamente "rn1" ou "rn4" continua usando os
                # wallpapers AVEX/Te Levo que já existem; pra qualquer outro
                # nome (parceiro de outra unidade), o wallpaper correspondente
                # é "wallpaper_<nome-em-slug>.png" - sem esse arquivo ainda no
                # projeto, cai automaticamente no Dominalog neutro, sem erro.
                parceiro_digitado = st.text_input(
                    "Parceiro desse código (deixe em branco pra Dominalog neutro)",
                    key="parceiro_novo_codigo",
                    placeholder="em branco = neutro · ou: rn1 · rn4 · nome de um parceiro novo",
                    help=(
                        "RN já tem 2 parceiros prontos - digite exatamente 'rn1' (AVEX) ou "
                        "'rn4' (Te Levo). Pra um parceiro de outra unidade, digite o nome dele "
                        "(ex: 'XPTO') - o wallpaper correspondente precisa se chamar "
                        "'wallpaper_xpto.png' no repositório (ver guia 'Como adicionar um "
                        "parceiro novo'). Sem esse arquivo ainda, fica neutro normalmente."
                    ),
                )
                parceiro_escolhido = parceiro_digitado.strip() or PARCEIRO_PADRAO
                if st.button("Adicionar código"):
                    novo = novo.strip()
                    if not novo:
                        st.warning("Digite um código antes de adicionar.")
                    elif novo in codigos_atuais:
                        st.warning("Esse código já existe.")
                    else:
                        try:
                            novo_usuario = {"codigo": novo, "parceiro": parceiro_escolhido}
                            salva_usuarios(atuais + [novo_usuario], f"Adiciona usuário {novo} ({parceiro_escolhido})")
                            st.success(
                                f"Código '{novo}' adicionado ao parceiro "
                                f"{rotulo_parceiro(parceiro_escolhido)}. Pode levar alguns segundos pra valer."
                            )
                            if parceiro_escolhido != PARCEIRO_PADRAO and parceiro_escolhido not in PARCEIROS_LEGADO:
                                caminho_wp = os.path.join(
                                    os.path.dirname(__file__), f"wallpaper_{slug_texto(parceiro_escolhido)}.png"
                                )
                                if not os.path.isfile(caminho_wp):
                                    st.info(
                                        f"Ainda não existe 'wallpaper_{slug_texto(parceiro_escolhido)}.png' no "
                                        f"projeto - esse usuário vai ver o Dominalog neutro até esse arquivo "
                                        f"ser adicionado ao repositório."
                                    )
                        except Exception as e:
                            st.error(f"Não consegui salvar: {e}")

                st.write("---")
                if atuais:
                    remover = st.selectbox("Remover algum código", options=[""] + codigos_atuais, key="remover_codigo")
                    if st.button("Remover código selecionado") and remover:
                        try:
                            salva_usuarios([u for u in atuais if u["codigo"] != remover], f"Remove usuário {remover}")
                            st.success(f"Código '{remover}' removido.")
                        except Exception as e:
                            st.error(f"Não consegui salvar: {e}")

                st.write("---")
                st.write("**🗂️ Bases de rotas vinculadas**")
                st.caption(
                    "Desvincular a base de um usuário faz a pergunta 'padrão ou "
                    "própria' aparecer de novo pra ele no próximo login."
                )
                usuarios_com_base = lista_usuarios_com_base_vinculada(codigos_atuais)
                if usuarios_com_base:
                    st.write(usuarios_com_base)
                    desvincular = st.selectbox(
                        "Desvincular base de qual usuário?",
                        options=[""] + usuarios_com_base,
                        key="desvincular_base",
                    )
                    if st.button("Desvincular base selecionada") and desvincular:
                        try:
                            desvincula_base_rotas_usuario(desvincular)
                            st.success(
                                f"Base de '{desvincular}' desvinculada. "
                                f"Ele vai ver a tela de escolha de novo no próximo login."
                            )
                        except Exception as e:
                            st.error(f"Não consegui desvincular: {e}")
                else:
                    st.write("_nenhum usuário com base vinculada ainda_")

                st.write("---")
                st.write("**🔒 Sessões ativas agora (trava de uso único por código)**")
                st.caption(
                    "Cada código só pode estar logado em um lugar por vez. Se alguém "
                    "fechou a aba sem clicar em 'Sair', a trava libera sozinha depois "
                    f"de {TIMEOUT_SESSAO_INATIVA_MIN} min parada - ou você pode liberar na mão aqui."
                )
                controle = _controle_sessoes()
                with controle["lock"]:
                    ativas_agora = {
                        cod: (agora_rn() - info["ultima_atividade"]).total_seconds() / 60
                        for cod, info in controle["ativas"].items()
                    }
                if ativas_agora:
                    for cod, minutos in sorted(ativas_agora.items()):
                        st.write(f"- `{cod}` — última atividade há {minutos:.0f} min")
                    liberar = st.selectbox(
                        "Forçar liberação de qual código?",
                        options=[""] + sorted(ativas_agora.keys()),
                        key="liberar_sessao",
                    )
                    if st.button("Forçar liberação selecionada") and liberar:
                        libera_codigo(liberar, forcar=True)
                        st.success(f"Sessão de '{liberar}' liberada. Ele já pode logar de outro lugar.")
                else:
                    st.write("_nenhuma sessão ativa agora_")

                # LINHA TEMPORÁRIA - ver comentário e função lá em cima
                # (roda_migracao_temporaria_avex). Pode apagar essa linha,
                # e a função inteira, na próxima atualização.
                roda_migracao_temporaria_avex(codigos_atuais, usuarios_com_base)
            elif senha_admin:
                st.error("Senha de administrador incorreta.")


# ------------------------------------------------------------------
# TELA 2+ - APÓS LOGIN
# ------------------------------------------------------------------
def cabecalho():
    col1, col2 = st.columns([4, 1])
    with col1:
        st.caption(f"Sessão de **{st.session_state.usuario}**")
    with col2:
        if st.button("Sair"):
            sair()


def monta_tabela_exemplo_base_rotas():
    """Dataframe de exemplo mostrando o formato esperado do
    base_rotas.xlsx - 3 linhas de Natal (Capital/Metropolitana), 3 de
    Parnamirim (incluindo uma rota separada pra bairros de praia) e o
    restante com cidades variadas do Interior. Nome de rota é
    DESCRITIVO, não numerado. Mostrada na tela como referência visual -
    o modelo de DOWNLOAD usa a base padrão real, não esse exemplo (ver
    monta_modelo_download_base_rotas). Pedido do Samuel em 19/08/2026."""
    dados = [
        ["TIROL", "LESTE", "Natal Metropolitana", "RN", "59020-000", "59022-999"],
        ["ALECRIM", "LESTE", "Natal Metropolitana", "RN", "59030-000", "59031-999"],
        ["PONTA NEGRA", "SUL", "Natal Metropolitana", "RN", "59090-000", "59099-999"],
        ["PARNAMIRIM", "PARNAMIRIM", "Parnamirim Metropolitana", "RN", "59140-000", "59149-999"],
        ["COHABINAL", "PARNAMIRIM", "Parnamirim Metropolitana", "RN", "59151-000", "59151-999"],
        ["PIRANGI DO NORTE", "PARNAMIRIM PRAIAS", "Parnamirim Metropolitana", "RN", "59161-000", "59161-999"],
        ["NOVA CRUZ", "SUL INTERIOR", "Interior RN", "RN", "59215-000", "59216-999"],
        ["CAICO", "SERIDÓ", "Interior RN", "RN", "59300-000", "59309-999"],
        ["APODI", "OESTE", "Interior RN", "RN", "59700-000", "59729-999"],
    ]
    colunas = ["Cidade ou Bairro", "Setor Operacional", "Mesorregião", "UF", "CEP Inicial", "CEP Final"]
    return pd.DataFrame(dados, columns=colunas)


TEXTO_ORIENTACOES_BASE_ROTAS = [
    "COMO AJUSTAR ESSA PLANILHA",
    "",
    "Essa é a base de rotas COMPLETA e REAL usada hoje pelo projeto - não "
    "é um exemplo. Ajuste as linhas conforme a divisão de rotas da sua "
    "operação: edite, apague ou adicione linhas na aba 'base_rotas'.",
    "",
    "O robô lê SEMPRE as 6 primeiras colunas, NESSA ORDEM (o nome do "
    "cabeçalho não importa, só a posição da coluna):",
    "1. Cidade ou Bairro — nome do local (um bairro, se for "
    "Capital/Metropolitana; ou o nome da cidade, se for Interior)",
    "2. Setor Operacional — nome da rota. Se tiver mais de uma rota com "
    "a mesma direção, use um qualificador descritivo (ex: 'SUL' e 'SUL "
    "INTERIOR') em vez de numerar (ex: 'SUL 1', 'SUL 2')",
    "3. Mesorregião — precisa conter a palavra 'Metropolitana' pra "
    "contar como Capital. Qualquer outro texto conta como Interior",
    "4. UF",
    "5. CEP Inicial",
    "6. CEP Final",
    "",
    "Colunas extras depois da 6ª são ignoradas pelo robô - pode usar pra "
    "suas próprias anotações, se quiser.",
    "",
    "Se essa planilha tiver uma aba 'Grupos Especiais', ela tem regras "
    "extras pra bairros que não batem exatamente com o nome oficial "
    "cadastrado (ex: vem só 'PIRANGI' em vez de 'PIRANGI DO NORTE') - "
    "ajuste ou remova as linhas que não fizerem sentido pra sua "
    "operação.",
    "",
    "Depois de ajustar, salve o arquivo e suba ele na tela 'Anexar "
    "minha própria base de rotas'.",
]


def monta_modelo_download_base_rotas(bytes_base_padrao):
    """Monta o .xlsx MODELO pronto pra baixar, ajustar e re-subir.

    Se existir uma base padrão configurada no projeto, o modelo é a
    BASE REAL COMPLETA (todas as rotas já cadastradas hoje, com as
    abas 'Grupos Especiais'/'LEIA-ME' que já existirem) - a pessoa parte
    do que já funciona e só ajusta o que for diferente pro caso dela, em
    vez de montar do zero. Se não houver base padrão configurada, cai
    num modelo simplificado só com as 9 linhas de exemplo. Em ambos os
    casos, adiciona (ou substitui) uma aba 'Orientações' com o passo a
    passo. Pedido do Samuel em 19-20/08/2026."""
    if bytes_base_padrao is not None:
        wb = openpyxl.load_workbook(io.BytesIO(bytes_base_padrao))
    else:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            monta_tabela_exemplo_base_rotas().to_excel(writer, index=False, sheet_name="base_rotas")
        buffer.seek(0)
        wb = openpyxl.load_workbook(buffer)

    # se ja existir uma aba "Orientacoes" de uma exportacao anterior,
    # remove antes de recriar - evita duplicar
    if "Orientações" in wb.sheetnames:
        del wb["Orientações"]

    ws = wb.create_sheet("Orientações")
    ws.column_dimensions["A"].width = 110
    for i, linha in enumerate(TEXTO_ORIENTACOES_BASE_ROTAS, start=1):
        cell = ws.cell(row=i, column=1, value=linha)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 1:
            cell.font = Font(bold=True, size=14)

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


def tela_escolha_base_rotas():
    """Tela mostrada só na PRIMEIRA vez que um usuário loga sem ter
    nenhuma base_rotas vinculada. Depois de escolher (anexar própria ou
    usar a padrão), nunca mais aparece pra esse usuário - a base fica
    salva permanentemente vinculada ao login dele. Pedido do Samuel em
    19/08/2026."""
    aplica_largura_estreita()
    cabecalho()
    st.header("Antes de começar: sua base de rotas")
    caixa_info_navy(
        "Essa é a primeira vez que você usa o robô. Escolha se quer usar a "
        "base de rotas padrão do projeto ou anexar a sua própria — essa "
        "escolha fica salva pra você, não vai perguntar de novo."
    )

    bytes_padrao = obter_bytes_base_padrao()

    if bytes_padrao is not None:
        st.subheader("Opção 1 — usar a base padrão do projeto")
        if st.button("✅ Usar a base padrão", type="primary"):
            try:
                salva_base_rotas_usuario(
                    st.session_state.usuario, bytes_padrao,
                    f"Vincula base padrao ao usuario {st.session_state.usuario}",
                )
                st.session_state.base_rotas_bytes = bytes_padrao
                st.session_state.base_rotas_verificada = True
                st.rerun()
            except Exception as e:
                st.error(f"Não consegui vincular a base padrão: {e}")
        st.divider()

    st.subheader("Opção 2 — anexar minha própria base de rotas")

    caixa_atencao_laranja(
        "Se sua roteirização usa como variável CIDADE, BAIRRO e faixa de "
        "CEP, sua planilha precisa ter, nessa ordem, as 6 primeiras "
        "colunas: Cidade ou Bairro (nome do local), Setor Operacional "
        "(nome da rota), Mesorregião (precisa conter a palavra "
        "'Metropolitana' pra contar como Capital — qualquer outro texto "
        "conta como Interior), UF, CEP Inicial e CEP Final. O nome dos "
        "cabeçalhos não importa, o robô lê pela posição da coluna. "
        "Colunas extras depois da 6ª são ignoradas. Dica pro nome da "
        "rota: se tiver mais de uma rota com a mesma direção, prefira um "
        "qualificador descritivo (ex: 'SUL' e 'SUL INTERIOR') em vez de "
        "numerar (ex: 'SUL 1', 'SUL 2').",
        icone="📋",
    )
    st.caption("Exemplo de como fica — 3 linhas de Natal, 3 de Parnamirim (com uma rota separada pra praias) e cidades do Interior:")
    st.dataframe(monta_tabela_exemplo_base_rotas(), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Baixar base completa pra ajustar ao meu uso",
        data=monta_modelo_download_base_rotas(bytes_padrao),
        file_name="base_rotas_modelo.xlsx",
        key="download_modelo_base_rotas",
    )
    st.caption(
        "Baixa a base real completa (com a aba 'Orientações' explicando cada coluna), "
        "ajusta as linhas pro seu caso e sobe abaixo."
    )

    arq_base_propria = st.file_uploader("Selecione seu base_rotas.xlsx", type=["xlsx"], key="upload_base_propria")

    if arq_base_propria is not None:
        conteudo_bytes = arq_base_propria.getvalue()
        try:
            preview_df = pd.read_excel(io.BytesIO(conteudo_bytes), sheet_name=0)
        except Exception as e:
            st.error(f"Não consegui ler esse arquivo como Excel: {e}")
            preview_df = None

        if preview_df is not None:
            st.write(f"**Prévia:** {len(preview_df)} linha(s), {len(preview_df.columns)} coluna(s).")
            st.dataframe(preview_df.head(10), use_container_width=True)

            if st.button("✅ Confirmar e vincular essa base", type="primary"):
                try:
                    salva_base_rotas_usuario(
                        st.session_state.usuario, conteudo_bytes,
                        f"Vincula base propria ao usuario {st.session_state.usuario}",
                    )
                    st.session_state.base_rotas_bytes = conteudo_bytes
                    st.session_state.base_rotas_verificada = True
                    st.rerun()
                except Exception as e:
                    st.error(f"Não consegui vincular sua base: {e}")

    st.divider()
    st.subheader("Opção 3 — usar somente o Setor como roteirização (não tenho base de rotas)")
    caixa_info_navy(
        "Pra quem não tem cidade/bairro/CEP cadastrados: a ROTA de cada CTRC sai direto "
        "da coluna SETOR que já vem pronta no próprio export do SSWWEB, sem passar por "
        "bairro/CEP nenhum. Quem ficar sem SETOR preenchido cai em 'Aguardando Validação'. "
        "Não existe separação Capital/Interior nesse modo — todas as rotas aparecem numa "
        "lista única na hora de gerar os romaneios em PDF."
    )
    if st.button("📍 Usar somente o Setor (não tenho base de rotas)"):
        try:
            bytes_sentinela = monta_base_setor_sentinela()
            salva_base_rotas_usuario(
                st.session_state.usuario, bytes_sentinela,
                f"Ativa modo somente Setor pro usuario {st.session_state.usuario}",
            )
            st.session_state.base_rotas_bytes = bytes_sentinela
            st.session_state.base_rotas_verificada = True
            st.rerun()
        except Exception as e:
            st.error(f"Não consegui ativar o modo Setor: {e}")


def checa_arquivo_desatualizado(df):
    """Sanity-check: olha a data mais recente dentro do PRÓPRIO arquivo
    (coluna DATA ULT. OCORRENCIA) e avisa se está muito antiga em
    relação a hoje - sinal de que o export está desatualizado."""
    if "DATA_ULT_OCORRENCIA_DT" not in df.columns or df["DATA_ULT_OCORRENCIA_DT"].isna().all():
        return None
    mais_recente = df["DATA_ULT_OCORRENCIA_DT"].max()
    if hasattr(mais_recente, "date"):
        mais_recente = mais_recente.date()
    dias_diferenca = (agora_rn().date() - mais_recente).days
    if dias_diferenca > 1:
        return mais_recente, dias_diferenca
    return None


def caixa_info_navy(texto, icone="📍"):
    """Caixa informativa em tom navy/roxo claro - substitui o st.info()
    padrão do Streamlit (azul, fora da paleta da marca) pra mensagens
    neutras/informativas, tipo 'processado às...'. Pedido do Samuel em
    19/08/2026, junto com caixa_atencao_laranja."""
    texto_seguro = html.escape(texto)
    st.markdown(
        f"""
        <div style="
            background-color: rgba(30, 30, 77, 0.85);
            border: 1px solid rgba(120, 110, 200, 0.5);
            border-radius: 6px;
            padding: 12px 16px;
            color: #FFFFFF;
            font-weight: 500;
            font-size: 14px;
            margin-bottom: 10px;
        ">
            {icone} {texto_seguro}
        </div>
        """,
        unsafe_allow_html=True,
    )


def caixa_atencao_laranja(texto, icone="⚠️"):
    """Caixa de atenção em laranja (paleta da marca), no lugar do
    st.warning() padrão do Streamlit (que sai verde-oliva e destoava do
    resto da tela) - usada tanto pro alerta de CTRC fora do RN quanto
    pros avisos de arquivo desatualizado/devolução ignorada/orientação
    de formato da base própria/possível fila de descarregamento. Pedido
    do Samuel em 19-20/08/2026."""
    texto_seguro = html.escape(texto)
    st.markdown(
        f"""
        <div style="
            background-color: rgba(230, 80, 0, 0.88);
            border-radius: 6px;
            padding: 12px 16px;
            color: #FFFFFF;
            font-weight: 600;
            font-size: 14px;
            margin-bottom: 10px;
        ">
            {icone} {texto_seguro}
        </div>
        """,
        unsafe_allow_html=True,
    )


def tela_upload_e_processamento():
    aplica_largura_estreita()
    cabecalho()
    st.header("1. Enviar arquivo SSWWEB")
    st.write("Suba o export do momento. O robô lê, classifica por rota e mostra uma prévia antes de gerar qualquer arquivo.")

    arq_sswweb = st.file_uploader("Arquivo .sswweb", type=["sswweb", "txt", "csv", "xls", "xlsx"])

    # A partir daqui a base ja foi decidida (ou pela tela_escolha_base_rotas,
    # ou o usuario ja tinha uma vinculada de antes) - session_state.base_rotas_bytes
    # sempre tem conteudo nesse ponto do fluxo (base de verdade OU o
    # sentinela do modo somente Setor - ver eh_modo_setor_ativo).
    modo_setor = eh_modo_setor_ativo(st.session_state.base_rotas_bytes)
    if modo_setor:
        st.success("Modo somente Setor ativo — sem base de rotas vinculada, a ROTA vem direto da coluna SETOR do arquivo.")
    else:
        st.success("Usando a base de rotas vinculada à sua conta.")

    if st.button("Processar arquivo", type="primary", disabled=arq_sswweb is None):
        with st.spinner("Lendo e classificando..."):
            horario_processamento = agora_rn()

            # salva os arquivos enviados em local temporário, porque as
            # funções do robô esperam CAMINHO de arquivo, não bytes
            pasta_tmp = tempfile.mkdtemp(prefix="robo_ssw_")
            caminho_sswweb = os.path.join(pasta_tmp, arq_sswweb.name)
            with open(caminho_sswweb, "wb") as f:
                f.write(arq_sswweb.getbuffer())

            caminho_base_rotas = os.path.join(pasta_tmp, "base_rotas.xlsx")
            with open(caminho_base_rotas, "wb") as f:
                f.write(st.session_state.base_rotas_bytes)

            try:
                df, base_rotas, grupos_especiais, rota_para_classificacao, resumo = robo.prepara_dados(
                    caminho_sswweb, caminho_base_rotas
                )
            except Exception as e:
                st.error(f"Não consegui ler o arquivo: {e}")
                return

            aviso_desatualizado = checa_arquivo_desatualizado(df)

            uf_predominante = resumo.get("uf_predominante")
            df = robo.classifica_rotas(df, base_rotas, grupos_especiais, uf_predominante, forcar_modo_setor=modo_setor)
            df_final = robo.monta_colunas_finais(df, uf_predominante)
            df_capital_interior, df_retidos, df_nao_liberadas, df_notas_em_rota, df_a_caminho = robo.separa_liberados_retidos_nao_liberados(
                df_final, rota_para_classificacao, uf_predominante, forcar_modo_setor=modo_setor
            )

            # a lista CODIGOS_LIBERADA_ROTA ja exclui 52/58/85 (nao sao
            # liberado de verdade) - o total_liberado_rota do resumo ja
            # vem certo.
            #
            # Card "Já em rota" (codigo 85) removido da tela em
            # 08/09/2026 a pedido do Samuel - continua fora daqui de
            # propósito mesmo apos a aba "Em rota" do Excel ter voltado
            # em 11/09/2026 (ver EXPORTAR_ABA_EM_ROTA em
            # processa_sswweb.py), porque o numero ainda carrega a mesma
            # limitacao (codigo 85 != confirmado em rota de verdade) - o
            # alerta de validar com BI fica concentrado na tela de
            # download (tela_download), junto com a aba reativada.

            st.session_state.resultado = {
                "horario": horario_processamento,
                "resumo": resumo,
                "uf_predominante": uf_predominante,
                "aviso_desatualizado": aviso_desatualizado,
                "df_capital_interior": df_capital_interior,
                "df_retidos": df_retidos,
                "df_nao_liberadas": df_nao_liberadas,
                "df_notas_em_rota": df_notas_em_rota,
                "df_a_caminho": df_a_caminho,
                "rota_para_classificacao": rota_para_classificacao,
                "pasta_tmp": pasta_tmp,
                "nome_arquivo": arq_sswweb.name,
                "modo_setor": modo_setor,
                "etapa": "previa",
            }
        st.rerun()


def tela_previa():
    cabecalho()
    r = st.session_state.resultado
    st.header("2. Prévia antes de gerar")

    caixa_info_navy(
        f"Processado em {r['horario'].strftime('%d/%m/%Y às %H:%M')} "
        f"a partir de {r['nome_arquivo']}."
    )
    if r["aviso_desatualizado"]:
        data_antiga, dias = r["aviso_desatualizado"]
        caixa_atencao_laranja(
            f"A ocorrência mais recente dentro do arquivo é de "
            f"{data_antiga.strftime('%d/%m/%Y')} ({dias} dia(s) atrás). "
            f"Confira se não subiu um export desatualizado por engano."
        )
    if r["resumo"].get("linhas_devolucao_ignoradas"):
        caixa_atencao_laranja(
            f"{r['resumo']['linhas_devolucao_ignoradas']} CTRC(s) de devolução não "
            f"entraram nessa rodada — esse arquivo é o relatório impresso, que desalinha "
            f"os campos nas linhas de devolução. Se precisar deles, use o export CSV "
            f"ou confira manualmente."
        )

    # UF predominante diferente de RN = arquivo de outra filial - o
    # base_rotas.xlsx (so tem cadastro do RN) e ignorado, mas ROTA sai
    # preenchida com a coluna SETOR que ja vem pronta do SSWWEB (ver
    # classifica_rotas) - REGIÃO continua em branco (nao tem
    # equivalente nesse modo). Excel sai nomeado "Base 081 - <UF>.xlsx"
    # em vez de "Capital e Interior - SSW.xlsx". O romaneio em PDF usa
    # a mesma tela/logica de individual-ou-coletivo por rota que ja
    # existia pro RN (ver gerar_romaneios_pdf) - nao tem mais a opção
    # separada "por cidade" (gerar_romaneio_por_cidade_pdf continua no
    # código, sem uso na tela, caso precise no futuro). Pedido do
    # Samuel em 07/09/2026.
    if r.get("uf_predominante") and r["uf_predominante"] != robo.UF_ESPERADA:
        caixa_atencao_laranja(
            f"UF predominante detectada: {r['uf_predominante']} (arquivo de outra filial, não RN). "
            f"A base de rotas do RN foi ignorada — ROTA sai preenchida com a coluna SETOR do próprio "
            f"arquivo (REGIÃO continua em branco), e o Excel final sai como "
            f"\"Base 081 - {r['uf_predominante']}.xlsx\". CTRC sem SETOR preenchido vai pra "
            f"Aguardando Validação.",
            icone="🗺️",
        )

    # Modo somente Setor ligado manualmente pelo usuário (sem UF
    # diferente detectada - arquivo normal do RN, só que sem base de
    # rotas vinculada). Mesma mecânica de ROTA=SETOR do aviso acima,
    # mas com texto diferente pra não confundir com "arquivo de outra
    # filial". Pedido do Samuel em 01/10/2026.
    if r.get("modo_setor"):
        caixa_info_navy(
            "Modo somente Setor ativo (sem base de rotas vinculada): ROTA sai preenchida "
            "direto com a coluna SETOR do próprio arquivo, sem passar por bairro/CEP. "
            "CTRC sem SETOR preenchido vai pra Aguardando Validação.",
            icone="📍",
        )

    resumo = r["resumo"]
    df_ci = r["df_capital_interior"]

    # Resumo geral e Resumo por prazo lado a lado (antes ficavam um bloco
    # embaixo do outro) - reduz a rolagem pra chegar na parte de gerar
    # romaneio. Pedido do Samuel em 27/08/2026.
    col_resumo1, col_resumo2 = st.columns([4, 5])
    with col_resumo1:
        st.write("**Resumo geral**")
        c1, c2 = st.columns(2)
        c1.metric("Notas lidas", resumo["total_lido"])
        c2.metric("Liberadas", resumo["total_liberado_rota"])

        # Detalhamento de quem NAO esta liberado p/ rota - antes era um
        # unico card "Ainda em transferência" (resumo["total_nao_liberado"])
        # que juntava situacoes bem diferentes numa so conta (retido pela
        # fiscalizacao, pendencia tipo endereco nao localizado/cliente
        # ausente, em transito de verdade), sem contar que "Já em Rota"
        # tambem entrava dentro desse numero apesar da legenda dizer que
        # nao contava em nenhum total. Quebrar nas mesmas 3 categorias que
        # ja viram aba no Excel (Retidos/Nao Liberadas/A Caminho) deixa
        # claro o que cada numero significa. Pedido do Samuel em
        # 28/08/2026 (achou o card unico "esquisito").
        c4, c5, c6 = st.columns(3)
        c4.metric("Retenção fiscal", len(r["df_retidos"]))
        c5.metric("Não liberadas", len(r["df_nao_liberadas"]))
        c6.metric("A caminho", len(r["df_a_caminho"]))
    with col_resumo2:
        # Checkbox pra tirar do "Resumo por prazo" os CTRCs do CARREFOUR
        # parados em "Falta de Documentação" - e uma pendencia cronica
        # desse cliente (nao depende da Dominalog resolver) que inflava o
        # numero de atrasados sem ser algo acionavel no dia a dia. So
        # afeta os cards de prazo aqui embaixo - "Liberadas" acima e o
        # resto da tela continuam mostrando o total real, sem filtro.
        # Pedido do Samuel em 27/08/2026.
        #
        # IMPORTANTE: "do CARREFOUR" e o REMETENTE/PAGADOR (quem despacha/
        # paga o frete) - CLIENTE e o DESTINATARIO, a pessoa que recebe a
        # mercadoria (nunca vai ser "CARREFOUR"). Usar CLIENTE deixava o
        # filtro sem efeito - achado testando com dado real do Samuel em
        # 27/08/2026.
        ocultar_carrefour = st.checkbox("Ocultar \"Falta de Documentação\" do CARREFOUR")
        base_prazo = df_ci
        if ocultar_carrefour:
            eh_carrefour = (
                df_ci["REMETENTE"].apply(robo.normaliza_texto).str.contains("CARREFOUR", na=False)
                | df_ci["PAGADOR"].apply(robo.normaliza_texto).str.contains("CARREFOUR", na=False)
            )
            eh_falta_doc = df_ci["DETALHE ÚLTIMA OCORRÊNCIA"].apply(robo.normaliza_texto).str.contains(
                "FALTA DE DOCUMENTACAO", na=False
            )
            base_prazo = df_ci[~(eh_carrefour & eh_falta_doc)]

        st.write("**Resumo por prazo** (só liberadas p/ rota)")
        ordem = ["ATRASO", "VENCE HOJE", "VENCE AMANHA", "VENCE EM 2 DIAS", "VENCE FUTURO"]
        dist = base_prazo["RESUMO"].value_counts(dropna=False).to_dict()
        cols_resumo = st.columns(len(ordem))
        for col, chave in zip(cols_resumo, ordem):
            col.metric(chave.replace("VENCE ", "").title(), dist.get(chave, 0))

    st.caption(
        "\"Liberadas\" já é o número certo pra entrar no Excel/romaneio agora."
    )

    if "ALERTA" in df_ci.columns:
        com_alerta = df_ci[df_ci["ALERTA"] != ""]
        if len(com_alerta):
            lista_ctrc = ", ".join(com_alerta["CTRC"].head(10).tolist())
            caixa_atencao_laranja(
                f"{len(com_alerta)} CTRC(s) com destino fora do RN — usuário precisa validar: {lista_ctrc}",
                icone="🚩",
            )

    # Aviso de possivel fila de descarregamento (ver
    # calcula_alerta_fila_descarregamento no processa_sswweb.py) - so
    # aparece quando tem algum CTRC flagado nessa rodada. O detalhe de
    # qual manifesto/quantos CTRCs ja sai tambem no topo da aba do
    # Excel (Capital e Interior - SSW.xlsx), essa caixa aqui e so um
    # aviso rapido antes mesmo de baixar o arquivo. Pedido do Samuel em
    # 20/08/2026.
    if "ALERTA_FILA_DESCARREGAMENTO" in df_ci.columns:
        com_fila = df_ci[df_ci["ALERTA_FILA_DESCARREGAMENTO"]]
        if len(com_fila):
            lista_ctrc_fila = ", ".join(com_fila["CTRC"].head(10).tolist())
            caixa_atencao_laranja(
                f"{len(com_fila)} CTRC(s) possivelmente ainda na fila de descarregamento "
                f"(chegaram hoje e o manifesto inteiro ainda não teve nenhum CTRC processado) "
                f"— confira antes de considerar liberado: {lista_ctrc_fila}",
                icone="🕓",
            )

    # Aviso de liberacao pela fiscalizacao (codigo 58) recente (hoje ou
    # ontem - ver calcula_liberado_fiscalizacao_recente no
    # processa_sswweb.py) - chama atencao pro usuario validar e ja
    # colocar em rota, sem precisar reparar sozinho que aquele CTRC saiu
    # da fiscalizacao ha pouco tempo. So NF + cidade (sem CTRC - o
    # Samuel achou desnecessario, ja que a validacao e feita pela nota
    # fiscal mesmo). Pedido do Samuel em 02/09/2026.
    if "LIBERADO_FISCALIZACAO_RECENTE" in df_ci.columns:
        liberadas_fiscalizacao = df_ci[df_ci["LIBERADO_FISCALIZACAO_RECENTE"]]
        if len(liberadas_fiscalizacao):
            lista_liberadas = ", ".join(
                f"NF {r['NOTA FISCAL']} ({r['CIDADE RECEBEDOR']})"
                for _, r in liberadas_fiscalizacao.head(10).iterrows()
            )
            caixa_atencao_laranja(
                f"{len(liberadas_fiscalizacao)} CTRC(s) liberado(s) pela fiscalização "
                f"hoje ou ontem — valide e já coloque em rota: {lista_liberadas}",
                icone="🔔",
            )

    st.divider()

    # ROTA agora sai preenchida tanto pro RN (base_rotas.xlsx) quanto
    # pra outra filial (coluna SETOR do proprio arquivo - ver
    # classifica_rotas) - a mesma tela de multiselect por rota serve
    # pros dois casos. So muda o que cai em "LOCAL" (junta so quem for
    # "Capital" - pra outra filial isso nunca acontece, ROTA la vira
    # sempre "Interior", entao toda rota aparece como opcao separada,
    # e gerar_romaneios_pdf decide sozinho quais rotas pequenas cabem
    # juntas num PDF "Rotas Agrupadas"). Pedido do Samuel em
    # 07/09/2026 (antes disso, tinha uma tela separada "por cidade"
    # so pra quem nao era RN - gerar_romaneio_por_cidade_pdf continua
    # no codigo, sem uso na tela, caso precise no futuro).
    st.write("**Gerar romaneios em PDF de quais rotas?** (Excel sai sempre com tudo)")

    roteirizavel = df_ci[df_ci["ROTA"].notna()].copy()
    roteirizavel["ROTA"] = roteirizavel["ROTA"].astype(str).str.strip()
    rotas_capital = sorted(roteirizavel.loc[roteirizavel["CAPITAL_INTERIOR"] == "Capital", "ROTA"].unique())
    rotas_interior = sorted(roteirizavel.loc[roteirizavel["CAPITAL_INTERIOR"] == "Interior", "ROTA"].unique())

    opcoes = []
    if rotas_capital:
        opcoes.append(("LOCAL - " + ", ".join(rotas_capital), list(rotas_capital)))
    for rota in rotas_interior:
        opcoes.append((rota, [rota]))

    # Rotas / data / botao lado a lado em vez de empilhados - antes cada
    # campo ficava numa linha, empurrando o botao pra bem embaixo da
    # tela e obrigando a rolar bastante pra chegar nele. Pedido do
    # Samuel em 27/08/2026 (mandou print mostrando o quanto tinha que
    # rolar). Em tela estreita/celular o Streamlit empilha essas
    # colunas automaticamente, entao continua funcionando no mobile.
    col_rotas, col_data, col_botao = st.columns([3, 2, 2])
    with col_rotas:
        # placeholder em portugues (era "Choose options", texto padrao
        # do Streamlit em ingles) - pedido do Samuel em 27/08/2026.
        escolhidas_texto = st.multiselect(
            "Rotas", options=[texto for texto, _ in opcoes], default=[],
            placeholder="Selecione as rotas",
        )
    with col_data:
        # format="DD/MM/YYYY" troca o placeholder "yyyy / mm / dd" (e a
        # ordem de digitacao) pro padrao brasileiro - pedido do Samuel
        # em 27/08/2026.
        data_limite_input = st.date_input(
            "Vencimento até (opcional)", value=None, format="DD/MM/YYYY"
        )
    with col_botao:
        # espaco vazio do tamanho de um label, so pra alinhar o botao na
        # mesma altura vertical dos outros dois campos (sem isso ele
        # fica mais alto, porque os outros campos tem rotulo em cima).
        st.write("")
        confirmar = st.button("Confirmar e gerar arquivos", type="primary", use_container_width=True)

    if confirmar:
        rotas_escolhidas = []
        for texto, lista_rotas in opcoes:
            if texto in escolhidas_texto:
                for rt in lista_rotas:
                    if rt not in rotas_escolhidas:
                        rotas_escolhidas.append(rt)

        with st.spinner("Gerando Excel e PDF..."):
            pasta_saida = os.path.join(r["pasta_tmp"], "saida")
            os.makedirs(pasta_saida, exist_ok=True)

            # Único arquivo Excel gerado (desde 21/08/2026, a pedido do
            # Samuel): antes existiam também o Capital - SSW.xlsx e o
            # Interior - SSW.xlsx separados; agora fica só esse aqui,
            # com as abas "Aguardando Validação", "Retenção fiscal", "Não
            # Liberadas", "A Caminho" e "Em rota" entrando junto quando
            # tiver algum CTRC em cada situação. A aba "Em rota" foi
            # desativada em 08/09/2026 e reativada em 11/09/2026 (ver
            # EXPORTAR_ABA_EM_ROTA em gerar_combinado_sswweb) - o numero
            # (codigo 85) continua com a limitação já conhecida, por
            # isso a tela de download (tela_download) mostra um alerta
            # pedindo pra validar com o BI antes de repassar.
            arq_combinado, n_combinado, n_aguardando, n_retidos, n_nao_liberadas, n_notas_em_rota, n_a_caminho = robo.gerar_combinado_sswweb(
                df_ci, r["df_retidos"], r["df_nao_liberadas"], r["df_notas_em_rota"], r["df_a_caminho"], pasta_saida,
                r.get("uf_predominante"), forcar_modo_setor=r.get("modo_setor", False),
            )

            arqs_romaneio = []
            if rotas_escolhidas:
                _pasta_romaneios, arqs_romaneio = robo.gerar_romaneios_pdf(
                    df_ci, pasta_saida, data_limite=data_limite_input, rotas_permitidas=rotas_escolhidas
                )

            r["arquivos_gerados"] = {
                os.path.basename(arq_combinado): arq_combinado,
            }
            for caminho_pdf in arqs_romaneio:
                r["arquivos_gerados"][os.path.basename(caminho_pdf)] = caminho_pdf

            r["n_aguardando"] = n_aguardando
            r["n_combinado"] = n_combinado
            r["n_retidos"] = n_retidos
            r["n_nao_liberadas"] = n_nao_liberadas
            r["n_notas_em_rota"] = n_notas_em_rota
            r["n_a_caminho"] = n_a_caminho
            r["etapa"] = "download"
        st.rerun()


def tela_download():
    aplica_largura_estreita()
    cabecalho()
    r = st.session_state.resultado
    st.header("3. Baixar")
    # "Em rota" fora dessa mensagem desde 08/09/2026 - continua assim
    # mesmo com a aba reativada em 11/09/2026, porque o numero (codigo
    # 85) tem a limitação explicada no alerta logo abaixo. Pedido do
    # Samuel.
    st.success(
        f"Pronto. Liberadas para rota: {r['n_combinado']} · "
        f"Retenção fiscal: {r['n_retidos']} · Não liberadas: {r['n_nao_liberadas']} · "
        f"Aguardando Validação: {r['n_aguardando']} · A caminho: {r['n_a_caminho']}."
    )

    # Aba "Em rota" reativada em 11/09/2026 a pedido do Samuel, com esse
    # alerta junto: o codigo 85 mostra quem já está em posse do parceiro
    # ou já foi manifestado pra ele, não necessariamente quem já saiu
    # fisicamente pra entrega - por isso pode haver diferença em relação
    # ao que o BI mostra (ver histórico em EXPORTAR_ABA_EM_ROTA, em
    # processa_sswweb.py).
    if r.get("n_notas_em_rota", 0) > 0:
        caixa_atencao_laranja(
            f"A aba \"Em rota\" traz {r['n_notas_em_rota']} CTRC(s) com código 85 "
            f"(saída em rota de entrega). Esse número reflete o que já está em "
            f"posse do parceiro/já foi manifestado pra ele — não necessariamente "
            f"quem já saiu pra entrega de fato. Valide com o BI (Regional Status) "
            f"antes de repassar, pois pode haver diferença.",
            icone="🛣️",
        )

    st.caption(f"Referente ao processamento das {r['horario'].strftime('%H:%M')} de {r['horario'].strftime('%d/%m')}.")

    for nome, caminho in r["arquivos_gerados"].items():
        with open(caminho, "rb") as f:
            st.download_button(f"⬇ {nome}", data=f.read(), file_name=nome, key=nome)

    st.divider()
    if st.button("Processar outro arquivo"):
        st.session_state.resultado = None
        st.rerun()


# ------------------------------------------------------------------
# ROTEADOR
# ------------------------------------------------------------------
if st.session_state.usuario is None:
    tela_login()
else:
    # renova o carimbo de atividade da trava de sessao unica a cada
    # tela carregada, pra ela nao expirar enquanto o usuario esta
    # ativo de verdade
    atualiza_atividade_codigo(st.session_state.usuario)

    # so checa no GitHub uma vez por sessao (evita ficar batendo na API
    # a cada rerun do Streamlit, que acontece toda hora)
    if not st.session_state.base_rotas_verificada:
        st.session_state.base_rotas_bytes = carrega_base_rotas_usuario(st.session_state.usuario)
        st.session_state.base_rotas_verificada = True

    if st.session_state.base_rotas_bytes is None:
        tela_escolha_base_rotas()
    elif st.session_state.resultado is None:
        tela_upload_e_processamento()
    elif st.session_state.resultado["etapa"] == "previa":
        tela_previa()
    else:
        tela_download()
