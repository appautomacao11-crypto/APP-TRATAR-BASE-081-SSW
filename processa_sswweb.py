# -*- coding: utf-8 -*-
"""
processa_sswweb.py
ROBO SSW 081 - versao SO-SSWWEB, AUTOSSUFICIENTE (nao depende do processa.py
do ROBO RN4 nem de nenhum outro arquivo .py - essa pasta funciona sozinha).

Nao cruza com o BI. Tudo - classificacao de liberado pra rota, prazo,
cliente, ultima ocorrencia - vem direto do proprio export .sswweb +
base_rotas.xlsx.

Decisoes de mapeamento (confirmadas com o Samuel em 06/08/2026):
  - JA_LIBERADA_ROTA: pelos codigos de ultima ocorrencia (lista validada
    contra o BI, 718/718 de acerto) + excecao do codigo 80 quando o
    proprio CTRC comeca com "RN" (emitido pela unidade RN4)
  - CLIENTE (quem vai receber) = DESTINATARIO
  - PAGADOR = PAGADOR
  - CEP = CEP RECEBEDOR (mesma prioridade que ja usamos pro bairro)
  - RESUMO (5 categorias) = calculado em cima de DATA PREVISTA vs hoje
  - DIAS PARADA EM PISO = dias UTEIS (seg-sex) desde DATA DA ULTIMA
    OCORRENCIA
  - Classificacao de rota (Natal/Parnamirim/SGA): BAIRRO RECEBEDOR tem
    prioridade sobre BAIRRO (CT-e) quando os dois sao validos mas
    diferentes - decisao tomada em 06/08/2026
  - PLACA DO ULTIMO MANIFESTO: nao existe nesse tipo de export do
    SSWWEB, fica de fora
"""

import os
import sys
import io
import csv
import re
import shutil
import unicodedata
import datetime
from zoneinfo import ZoneInfo
import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, PatternFill, Font
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table as PdfTable, TableStyle, Paragraph, Spacer, KeepTogether, PageBreak
from reportlab.pdfgen import canvas as pdfcanvas
from pypdf import PdfReader

# ------------------------------------------------------------------
# LOG DETALHADO
# ------------------------------------------------------------------
# Quando compilado com PyInstaller --onefile, __file__ aponta pra pasta
# temporaria de extracao (_MEIPASS) e nao pra pasta onde o .exe de fato
# esta - por isso, em modo "frozen", usamos sys.executable em vez de
# __file__ (mesma correcao ja aplicada no ROBO RN4).
if getattr(sys, "frozen", False):
    PASTA_BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
# normpath resolve o ".." de uma vez (em vez de deixar cru no meio do
# caminho) - evita erro WinError 3 que aconteceu no .exe compilado ao
# tentar copiar arquivo de Downloads pra dados\ com caminho nao resolvido
PASTA_DADOS = os.path.normpath(os.path.join(PASTA_BASE, "..", "dados"))
PASTA_SAIDA = os.path.normpath(os.path.join(PASTA_BASE, "..", "saida"))
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
ARQ_BASE_ROTAS = os.path.join(PASTA_DADOS, "base_rotas.xlsx")
ARQ_LOG = os.path.join(PASTA_DADOS, "log_processamento_sswweb.txt")
_log_buffer = []

# O servidor do Streamlit Cloud roda no horario UTC, mas Natal/RN e o
# resto do Brasil (exceto ilhas) fica em UTC-3 o ano todo (sem horario
# de verao desde 2019). Sem isso, datetime.now() usa UTC e todo horario
# impresso (romaneio em PDF, logs) sai 3h adiantado em relacao ao
# horario local. Pedido do Samuel em 26/08/2026.
FUSO_BRASIL = ZoneInfo("America/Fortaleza")


def agora_brasil():
    """datetime.now() ja no horario de Natal/RN (UTC-3), independente
    do fuso horario do servidor onde o robo estiver rodando."""
    return datetime.datetime.now(FUSO_BRASIL)


def log_detalhe(msg):
    linha = f"[{agora_brasil().strftime('%H:%M:%S')}] {msg}"
    _log_buffer.append(linha)
    try:
        with open(ARQ_LOG, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except OSError:
        pass


def mostra_log_por_causa_de_erro():
    if not _log_buffer:
        return
    print("\n" + "=" * 70)
    print("Deu um erro - aqui esta o detalhe passo a passo pra ajudar a entender:")
    print("=" * 70)
    for linha in _log_buffer:
        print(linha)
    print("=" * 70)
    print(f"(esse mesmo detalhe tambem fica salvo em: {ARQ_LOG})\n")


CIDADES_CAPITAL_MET = {"NATAL", "PARNAMIRIM", "SAO GONCALO DO AMARANTE"}

# Excecoes de CEP dentro de cidade do interior (comunidade sem status de
# municipio proprio, roteirizada junto com a cidade-sede vizinha).
#
# A excecao de Mossoro (CEP 59649899, comunidade Maisa/Jucuri) que existia
# aqui apontando pra "ROTA BARAÚNA E+" foi REMOVIDA em 02/09/2026: o
# base_rotas.xlsx atual ja consolidou "ROTA MOSSORÓ", "ROTA BARAÚNA E+" e
# "ROTA AREIA BRANCA E+" numa unica rota "10 MOSSORÓ" - ou seja, esse CEP
# ia cair no mesmo lugar do resto de Mossoro de qualquer jeito, a excecao
# virou redundante (dava o mesmo resultado do caminho geral por cidade).
# Se no futuro esse CEP precisar de novo ir pra uma rota DIFERENTE do
# resto de Mossoro, e so acrescentar a entrada de volta aqui.
EXCECOES_CEP_CIDADE_INTERIOR = {}

# ------------------------------------------------------------------
# CODIGOS DE ULTIMA OCORRENCIA que significam "liberado pra rota"
# Lista oficial confirmada pelo Samuel em 18/08/2026 direto contra a
# coluna SITUACAO do BI (Regional Status): 9, 10, 13, 30, 31, 32, 52,
# 54, 58, 74, 75, 84, 98 = "Liberada Para Rota". Codigo 52 e sempre
# "FALTA DE DOCUMENTACAO" (na pratica, quase sempre CARREFOUR) -
# liberado pra rota, mas ganha alerta visivel na coluna ALERTA do Excel
# (ver eh_falta_documentacao). Codigo 58 e "MERCADORIA LIBERADA PELA
# FISCALIZACAO" - liberado normal, sem relacao com falta de
# documentacao.
#
# IMPORTANTE - codigos removidos depois dessa validacao:
#   85 (SAIDA EM ROTA DE ENTREGA), removido em 18/08/2026: o motorista
#     JA SAIU com a carga - isso e "ja esta em rota", nao "liberado
#     pra entrar em rota". Contar como liberado inflava o numero
#     mostrado na tela (ex: 317 quando o real "liberado pra
#     roteirizar" era bem menor).
#   10 (ENDERECO DESTINATARIO NAO LOCALIZADO), removido em 25/08/2026 a
#     pedido do Samuel: apesar do BI marcar como "Liberada Para Rota",
#     esses CTRC passam a cair na aba "Nao Liberadas" em vez de
#     "Capital e Interior", pra ficar mais visivel que o endereco
#     precisa ser confirmado antes de rotear de novo.
#
# Codigo 3 (FALHA MECANICA) ACRESCENTADO em 08/09/2026: inicialmente
# tinha ficado de fora de proposito (ficaria como "Nao Liberadas"), mas
# batendo contra o BI (Regional Status) - 24/24 ocorrencias de codigo 3
# em toda a base regional (nao so RN4) marcadas como "Liberada Para
# Rota" - Samuel decidiu seguir o BI em vez da decisao anterior.
# ------------------------------------------------------------------
CODIGOS_LIBERADA_ROTA = {
    "3", "9", "13", "30", "31", "32", "52", "54", "58", "74", "75", "84", "98"
}
# Excecao: codigo 80 sozinho e ambiguo (so 8% das vezes e liberada) - MAS
# quando o CTRC foi emitido pela propria unidade RN4 (numero comeca com
# "RN"), o codigo 80 bateu 100% (9/9) com liberada para rota.
PREFIXO_CTRC_EMISSAO_PROPRIA = "RN"

# Codigo 56 = "MERCADORIA RETIDA PELA FISCALIZACAO" - mesma constante ja
# validada e confirmada com o Samuel no robo desktop em 21/08/2026. O
# par dele, codigo 58 "MERCADORIA LIBERADA PELA FISCALIZACAO", ja faz
# parte de CODIGOS_LIBERADA_ROTA - ou seja, um CTRC nunca fica em
# RETIDO e LIBERADO ao mesmo tempo. Se aparecer outro tipo de retencao
# no futuro (ex: judicial), e so acrescentar o codigo novo nesse
# conjunto.
CODIGOS_RETIDO = {"56"}

COLUNAS_SSWWEB_COMPLETO = [
    "CTRC", "NFISCAL", "REMETENTE", "PAGADOR", "DESTINATARIO",
    "CIDADE", "BAIRRO", "CEP", "ENDERECO",
    "BAIRRO RECEBEDOR", "CEP RECEBEDOR", "CIDADE RECEBEDOR", "UF RECEBEDOR",
    "COMPLEMENTO RECEBEDOR",
    "DATA PREVISTA",
    "ULT. OCORRENCIA", "DESCRICAO ULT. OCORRENCIA", "DATA ULT. OCORRENCIA",
    # MANIFESTO/END adicionada em 20/08/2026: usada pra estimar se um
    # CTRC em codigo 84 (chegada na unidade) pode ainda estar na fila de
    # descarregamento - ver calcula_alerta_fila_descarregamento.
    "MANIFESTO/END",
    "KG REAL", "M3", "PESO CALCULO",
    # SETOR adicionada em 07/09/2026: pros arquivos de outra filial (UF
    # predominante != RN), onde nao existe base_rotas.xlsx cadastrada,
    # essa coluna (que ja vem pronta do proprio SSWWEB, representando a
    # roteirizacao/regiao daquela filial) passa a ser usada como ROTA -
    # ver classifica_rotas. So nao vem no formato "relatorio impresso"
    # (ler_sswweb_relatorio_impresso nao tem esse campo).
    "SETOR",
]

# UF que o robo espera receber pra entrega - fora dela, o CTRC pode ser
# na verdade uma DEVOLUCAO indo pra outro estado, nao uma entrega normal.
UF_ESPERADA = "RN"


def detecta_uf_predominante(df):
    """UF mais frequente em UF RECEBEDOR. Usada pra saber se esse
    arquivo e de uma base fora do RN (outra filial, ex: PE) - nesse
    caso o robo ignora o base_rotas.xlsx (que so tem cadastro de
    Natal/Parnamirim/SGA/interior do RN) em vez de jogar tudo pra
    Aguardando Validação por falta de rota cadastrada (o que acontecia
    antes: toda linha passava pela cascata de classificar_rota, nao
    batia em nada, e caia em "VERIFIQUE BAIRRO/CIDADE"). Retorna None
    se a coluna nao existir ou vier toda vazia. Pedido do Samuel em
    07/09/2026."""
    uf_receb = df.get("UF RECEBEDOR")
    if uf_receb is None:
        return None
    ufs = uf_receb.fillna("").astype(str).str.strip().str.upper()
    ufs = ufs[ufs != ""]
    if ufs.empty:
        return None
    return ufs.mode().iloc[0]

ORDEM_RESUMO = ["ATRASO", "VENCE HOJE", "VENCE AMANHA", "VENCE EM 2 DIAS", "VENCE FUTURO"]

LINHA_INICIO_TABELA = 5  # tabela comeca nessa linha; 1-2 = resumo, 3-4 = branco


# ------------------------------------------------------------------
# FUNCOES AUXILIARES (normalizacao, CEP, validacao de colunas)
# ------------------------------------------------------------------
def normaliza_texto(txt):
    """Maiuscula, sem acento, sem espaco duplicado/nas pontas."""
    if txt is None:
        return ""
    txt = str(txt).strip().upper()
    txt = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


# Dentro do que sobra em "Nao Liberadas", uma parte e CTRC EM
# TRANSFERENCIA entre unidades - pedido do Samuel em 26/08/2026 pra
# virar aba propria "A Caminho", separada do resto de "Nao Liberadas"
# (que sao problemas de verdade: endereco nao localizado, pedido
# cancelado etc). Comparado contra DETALHE ÚLTIMA OCORRÊNCIA ja
# normalizado (normaliza_texto), pra nao depender de acento/caixa.
#   - "EM TRANSITO - SAIDA UNIDADE TRANSBORDO": saiu de uma unidade de
#     transbordo mas ainda nao chegou na unidade de entrega.
#   - "VEICULO EM POSTO FISCAL PARA SELAGEM DAS NOTAS" (codigo 57,
#     acrescentada em 27/08/2026): veiculo parado num posto de
#     fiscalizacao selando as notas - tambem em transito, so que
#     retido temporariamente na fiscalizacao em vez de na unidade.
# Se aparecer outra variacao de "em transito" no futuro, e so
# acrescentar aqui.
DESCRICOES_A_CAMINHO = {
    normaliza_texto("EM TRANSITO - SAIDA UNIDADE TRANSBORDO"),
    normaliza_texto("VEICULO EM POSTO FISCAL PARA SELAGEM DAS NOTAS"),
}


def cep_para_numero(cep):
    """Converte '59010-120' ou '59010120' em inteiro 59010120."""
    if cep is None:
        return None
    digitos = re.sub(r"\D", "", str(cep))
    if len(digitos) < 8:
        return None
    return int(digitos[:8])


def bairro_valido(txt):
    """Considera 'nao e um bairro de verdade': vazio, so pontuacao, tem
    numero no meio (CEP digitado por engano), ou a palavra 'OUTRO'."""
    if txt is None:
        return False
    limpo = str(txt).strip().strip(".").strip()
    if len(limpo) == 0:
        return False
    if any(ch.isdigit() for ch in limpo):
        return False
    if limpo.upper() == "OUTRO":
        return False
    return True


def texto_descritivo_valido(txt):
    """Versao mais permissiva de bairro_valido, usada so pro COMPLEMENTO
    RECEBEDOR quando ele serve de EXIBICAO (nao classificacao) pra
    cidades do interior: aceita numero no meio (ex: 'ANDAR 44', 'APTO
    201', 'BLOCO 3') porque ali e so um indicativo de endereco, nao
    precisa ser nome de bairro oficial - so rejeita vazio ou so
    pontuacao (ex: '..'). Pedido do Samuel em 12/08/2026."""
    if txt is None:
        return False
    limpo = str(txt).strip().strip(".").strip()
    return len(limpo) > 0


def dias_parados_desde(data_ultima_ocorrencia):
    """Quantos dias UTEIS (seg-sex) a nota esta parada na mesma
    ocorrencia, contando a partir do dia SEGUINTE a data ate hoje."""
    if pd.isna(data_ultima_ocorrencia):
        return None
    hoje = datetime.date.today()
    data = data_ultima_ocorrencia
    if isinstance(data, datetime.datetime):
        data = data.date()
    if not isinstance(data, datetime.date):
        return None
    if data >= hoje:
        return 0
    inicio = data + datetime.timedelta(days=1)
    dias = 0
    d = inicio
    while d <= hoje:
        if d.weekday() < 5:  # seg-sex
            dias += 1
        d += datetime.timedelta(days=1)
    return dias


def valida_colunas(df, colunas_esperadas, nome_arquivo, obrigatorias=None):
    """Confere se as colunas esperadas estao no arquivo lido - avisa (sem
    travar) se mudou, e interrompe so se faltar coluna OBRIGATORIA."""
    faltando = [c for c in colunas_esperadas if c not in df.columns]
    if faltando:
        print(f"\n  [AVISO] {nome_arquivo}: {len(faltando)} coluna(s) esperada(s) nao encontrada(s):")
        for c in faltando:
            print(f"     - {c}")
        print(f"     Isso indica que a estrutura do arquivo mudou. Confira antes de confiar no resultado.\n")
    if obrigatorias:
        faltando_obrig = [c for c in obrigatorias if c not in df.columns]
        if faltando_obrig:
            raise ValueError(
                f"{nome_arquivo}: faltam colunas OBRIGATORIAS pro robo funcionar: {faltando_obrig}. "
                f"O arquivo de origem mudou de estrutura - avise quem gera esse arquivo antes de rodar de novo."
            )
    return faltando


def abre_planilha_flexivel(caminho, **kwargs_read_excel):
    """Abre um Excel independente do formato real por baixo do arquivo."""
    motores = [None, "xlrd", "openpyxl"]
    ultimo_erro = None
    for motor in motores:
        try:
            if motor is None:
                return pd.read_excel(caminho, **kwargs_read_excel)
            return pd.read_excel(caminho, engine=motor, **kwargs_read_excel)
        except Exception as e:
            ultimo_erro = e
            continue
    raise ultimo_erro


def converte_numero_br(txt):
    """Converte texto no formato brasileiro ('108,1500' ou '   31,100')
    para numero de verdade (float). O SSWWEB entrega peso/cubagem assim,
    como texto - sem isso o Excel nao consegue somar/calcular media
    nessas colunas."""
    if txt is None:
        return None
    limpo = str(txt).strip()
    if not limpo:
        return None
    limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return None


def nome_tabela_valido(prefixo, texto):
    """Nome de Tabela do Excel precisa ser unico, sem espaco/acento e nao
    pode comecar com numero."""
    limpo = normaliza_texto(texto)
    limpo = re.sub(r"[^A-Z0-9]+", "_", limpo).strip("_")
    if not limpo:
        limpo = "SEM_NOME"
    return f"{prefixo}_{limpo}"[:250]


# ------------------------------------------------------------------
# LEITURA DA BASE DE ROTAS
# ------------------------------------------------------------------
def ler_base_rotas(caminho):
    log_detalhe("Lendo base de ROTAS (base_rotas.xlsx)...")
    df_bruto = abre_planilha_flexivel(caminho, sheet_name=0)

    if df_bruto.shape[1] < 6:
        raise ValueError(
            f"base_rotas.xlsx: esperava pelo menos 6 colunas, encontrei {df_bruto.shape[1]}. "
            f"A estrutura desse arquivo mudou - confira antes de continuar."
        )
    if df_bruto.shape[1] > 6:
        print(f"  [info] base_rotas.xlsx tem {df_bruto.shape[1]} colunas, mas o robo so usa as 6 primeiras "
              f"(CIDADE_BAIRRO, SETOR_OP, MESORREGIAO, UF, CEP_INI, CEP_FIM). Confira se a ordem continua a mesma.")

    df = df_bruto.iloc[:, :6].copy()
    df.columns = ["CIDADE_BAIRRO", "SETOR_OP", "MESORREGIAO", "UF", "CEP_INI", "CEP_FIM"]

    cep_ini_num = df["CEP_INI"].apply(cep_para_numero)
    pct_cep_valido = cep_ini_num.notna().mean() if len(df) else 0
    uf_parece_sigla = df["UF"].dropna().astype(str).str.strip().str.len().eq(2).mean() if df["UF"].notna().any() else 0

    if pct_cep_valido < 0.7:
        print(
            f"\n  [AVISO IMPORTANTE] base_rotas.xlsx: a coluna 'CEP_INI' (5a coluna) deveria ter "
            f"numero de CEP, mas so {pct_cep_valido:.0%} das linhas parecem numero.\n"
            f"     Forte indicio de que ALGUEM REORDENOU AS COLUNAS - PARE e confira antes de confiar no resultado.\n"
        )
    if uf_parece_sigla < 0.7:
        print(
            f"  [AVISO IMPORTANTE] base_rotas.xlsx: a coluna 'UF' (4a coluna) deveria ter siglas de "
            f"2 letras, mas so {uf_parece_sigla:.0%} das linhas parecem isso.\n"
        )

    df["CHAVE_NORM"] = df["CIDADE_BAIRRO"].apply(normaliza_texto)
    df["CEP_INI_NUM"] = cep_ini_num
    df["CEP_FIM_NUM"] = df["CEP_FIM"].apply(cep_para_numero)
    df["REGIAO"] = df["MESORREGIAO"].apply(
        lambda x: "Capital e Met" if "Metropolitana" in str(x) else "Interior"
    )
    return df


def ler_grupos_especiais(caminho):
    """Le a aba 'Grupos Especiais' do base_rotas.xlsx - regras extras pra
    Parnamirim e Sao Goncalo do Amarante quando o bairro do CT-e nao bate
    exatamente com o nome oficial cadastrado (ex: vem so "PIRANGI" em vez
    de "PIRANGI DO NORTE"). Se a aba nao existir (planilha antiga, ainda
    sem essa aba), o robo continua funcionando - so nao aplica regra
    extra nenhuma pra essas duas cidades (fica so na busca por
    CEP/rota geral)."""
    log_detalhe("Lendo aba 'Grupos Especiais' do base_rotas.xlsx...")
    try:
        df = abre_planilha_flexivel(caminho, sheet_name="Grupos Especiais")
    except Exception:
        print("  [info] base_rotas.xlsx nao tem a aba 'Grupos Especiais' - "
              "seguindo sem regras extras pra Parnamirim/SGA.")
        log_detalhe("[grupos especiais] aba nao encontrada - seguindo sem regras extras")
        return {}

    valida_colunas(
        df, ["CIDADE", "PALAVRA_CHAVE_NO_BAIRRO", "ROTA", "CLASSIFICACAO_CAPITAL_INTERIOR"],
        "Grupos Especiais", obrigatorias=["CIDADE", "PALAVRA_CHAVE_NO_BAIRRO", "ROTA"],
    )

    grupos = {}
    for _, linha in df.iterrows():
        cidade = normaliza_texto(linha.get("CIDADE"))
        palavra = normaliza_texto(linha.get("PALAVRA_CHAVE_NO_BAIRRO"))
        rota = str(linha.get("ROTA") or "").strip()
        classif = str(linha.get("CLASSIFICACAO_CAPITAL_INTERIOR") or "").strip() or None
        if not cidade or not palavra or not rota:
            continue
        grupos.setdefault(cidade, []).append((palavra, rota, classif))
    return grupos


def monta_rota_para_classificacao(base_rotas, grupos_especiais):
    """Pra cada ROTA (setor operacional), decide se ela conta como
    Capital ou Interior na hora de separar os arquivos de exportacao.
    Prioridade: (1) o que estiver escrito na coluna CLASSIFICACAO da aba
    'Grupos Especiais', (2) o valor mais comum de REGIAO (calculado a
    partir da Mesorregiao) pras linhas dessa rota no base_rotas.xlsx
    principal - assim nao precisa mais checar nome de rota fixo
    ("4 SUL" etc.) no meio do codigo."""
    mapa = {}
    for lista in grupos_especiais.values():
        for _palavra, rota, classif in lista:
            if classif and rota not in mapa:
                mapa[rota] = "Interior" if "interior" in classif.lower() else "Capital e Met"

    if len(base_rotas) and "REGIAO" in base_rotas.columns:
        modas = base_rotas.groupby("SETOR_OP")["REGIAO"].agg(lambda s: s.value_counts().idxmax())
        for rota, regiao in modas.items():
            mapa.setdefault(str(rota).strip(), regiao)
    return mapa


# ------------------------------------------------------------------
# CLASSIFICACAO DE ROTA (bairro -> CEP -> cidade -> nao identificado)
# ------------------------------------------------------------------
def busca_por_cep(cep_n, base_rotas):
    """Devolve (setor, regiao, bairro_cadastrado)."""
    if not cep_n:
        return None
    candidatos = base_rotas[
        (base_rotas["CEP_INI_NUM"].notna()) &
        (base_rotas["CEP_FIM_NUM"].notna()) &
        (base_rotas["CEP_INI_NUM"] <= cep_n) &
        (cep_n <= base_rotas["CEP_FIM_NUM"])
    ]
    if len(candidatos) > 0:
        linha = candidatos.iloc[0]
        return linha["SETOR_OP"], linha["REGIAO"], linha["CIDADE_BAIRRO"]
    return None


def busca_por_prefixo(bairro_n, mapa_chave, mapa_nome, tam_minimo=6):
    """Bairro truncado casando com a chave completa (ou vice-versa)."""
    if not bairro_n or len(bairro_n) < tam_minimo:
        return None
    for chave, valor in mapa_chave.items():
        if len(chave) < tam_minimo:
            continue
        if bairro_n.startswith(chave) or chave.startswith(bairro_n):
            setor, regiao = valor
            return setor, regiao, mapa_nome.get(chave, chave)
    return None


def busca_bairro_em_texto(texto, mapa_chave, mapa_nome, tam_minimo=6):
    """ULTIMO RECURSO: procura nome de bairro cadastrado dentro de texto
    livre (ENDERECO ou COMPLEMENTO RECEBEDOR)."""
    texto_n = normaliza_texto(texto)
    if not texto_n:
        return None
    melhor_chave = None
    melhor_valor = None
    melhor_tam = 0
    for chave, valor in mapa_chave.items():
        if len(chave) < tam_minimo:
            continue
        if chave in texto_n and len(chave) > melhor_tam:
            melhor_chave = chave
            melhor_valor = valor
            melhor_tam = len(chave)
    if melhor_chave is None:
        return None
    setor, regiao = melhor_valor
    return setor, regiao, mapa_nome.get(melhor_chave, melhor_chave)


# PARNAMIRIM e SAO GONCALO DO AMARANTE: grupos especiais de bairro que nao
# batem 100% com o nome oficial cadastrado no base_rotas.xlsx (ex: vem so
# "PIRANGI" em vez de "PIRANGI DO NORTE"). Antes ficavam fixos aqui no
# codigo (GRUPO_4_SUL_PARNAMIRIM etc.) - agora vem da aba "Grupos
# Especiais" do proprio base_rotas.xlsx (ver ler_grupos_especiais), pra
# dar pra atualizar so editando a planilha quando o parceiro de entrega
# mudar a divisao de rotas dele. Refatorado a pedido do Samuel em
# 14/08/2026 (troca de parceiro pra AVEX).
ROTA_PADRAO_PARNAMIRIM = "PARNAMIRIM"
REGIAO_PARNAMIRIM = "Capital e Met"
ROTA_PADRAO_SGA_INTERIOR = "INTERIOR NORTE"
REGIAO_SGA_INTERIOR = "Interior"


def classifica_capital_interior(cidade_n, rota, bairro_n, rota_para_classificacao=None):
    """Classificacao Capital/Interior usada pro arquivo de exportacao.
    Pra Parnamirim e SGA, consulta o mapa rota_para_classificacao
    (montado a partir da aba 'Grupos Especiais' + base_rotas.xlsx, ver
    monta_rota_para_classificacao) em vez de checar nome de rota fixo
    ("4 SUL" etc.) - assim continua certo mesmo se o nome da rota mudar."""
    if pd.isna(rota) or not str(rota).strip():
        return "Aguardando Validação"
    if cidade_n == "NATAL":
        return "Capital"
    if cidade_n in ("PARNAMIRIM", "SAO GONCALO DO AMARANTE"):
        classif = (rota_para_classificacao or {}).get(str(rota).strip())
        if classif is not None:
            return "Interior" if classif == "Interior" else "Capital"
        # sem informacao pra essa rota na planilha: mantem o padrao de
        # cada cidade (Parnamirim = Capital, SGA = Interior por padrao)
        return "Capital" if cidade_n == "PARNAMIRIM" else "Interior"
    return "Interior"


def classifica_bairro_parnamirim(bairro, bairro_receb, cep_n=None, cep_receb_n=None,
                                  base_rotas=None, mapa_capital=None, grupo_parnamirim=None):
    """CASCATA: grupos especiais (aba 'Grupos Especiais') > CEP (contra
    TODOS os bairros cadastrados na Regiao Metropolitana - PARQUE DAS
    NACOES, NOVA PARNAMIRIM etc., nao so a linha generica 'PARNAMIRIM')
    > BAIRRO RECEBEDOR > BAIRRO (CT-e) > BAIRRO generico (bloqueado se
    for bairro conhecido de outra cidade) > AGUARDANDO.
    Ate 28/08/2026 o CEP so era comparado contra a linha 'PARNAMIRIM'
    (a rota generica da cidade, CEP 59140-59161) - bairros especificos
    dentro dessa faixa, como PARQUE DAS NACOES (CEP 59158-59159, rota
    'NOVA PARNAMIRIM+'), nunca eram encontrados por CEP, mesmo
    cadastrados no base_rotas.xlsx. Corrigido a pedido do Samuel (achado
    no CTRC SPO065465-5: CEP 59158-256 caia direto na linha generica em
    vez da linha especifica de PARQUE DAS NACOES, que casa certinho).
    Como PARQUE DAS NACOES/NOVA PARNAMIRIM aparecem ANTES da linha
    'PARNAMIRIM' no base_rotas.xlsx, busca_por_cep (que pega a primeira
    linha que bate) ja devolve o bairro mais especifico primeiro -
    'PARNAMIRIM' só aparece se nenhum bairro mais especifico bater."""
    bairro_n = normaliza_texto(bairro)
    bairro_receb_n = normaliza_texto(bairro_receb)

    for palavra, rota, _classif in (grupo_parnamirim or []):
        if palavra in bairro_receb_n:
            return rota, f"OK - Parnamirim, grupo especial '{rota}' (bairro recebedor)", bairro_receb
        if palavra in bairro_n:
            return rota, f"OK - Parnamirim, grupo especial '{rota}' (recebedor nao bateu)", bairro

    if base_rotas is not None:
        base_regiao_metropolitana = base_rotas[base_rotas["REGIAO"] == "Capital e Met"]
        resultado = (busca_por_cep(cep_receb_n, base_regiao_metropolitana)
                     or busca_por_cep(cep_n, base_regiao_metropolitana))
        if resultado:
            setor, _regiao, bairro_cadastrado = resultado
            # So troca o texto do bairro exibido quando o CEP bateu num
            # bairro ESPECIFICO cadastrado (ex: PARQUE DAS NACOES, EMAUS).
            # Quando o unico que bate e a propria linha generica
            # 'PARNAMIRIM' (faixa larga, cobre quase toda a cidade), NAO
            # sobrescreve - mantem o bairro original (RECEBEDOR de
            # preferencia, senao o do CT-e), senao um nome real como
            # 'CAJUPIRANGA' vira o generico 'PARNAMIRIM' no Excel, uma
            # perda de informacao. Achado testando a correcao acima
            # contra os 102 CTRCs de Parnamirim do dia 28/08/2026.
            if normaliza_texto(bairro_cadastrado) != "PARNAMIRIM":
                return setor, "OK - Parnamirim, rota por CEP (bairro corrigido)", bairro_cadastrado
            bairro_exibir = bairro_receb if bairro_valido(bairro_receb) else bairro
            return setor, "OK - Parnamirim, rota geral (por CEP, bairro original mantido)", bairro_exibir

        if bairro_valido(bairro_receb):
            return (ROTA_PADRAO_PARNAMIRIM,
                    "OK - Parnamirim, rota geral (bairro recebedor nao reconhecido, mas valido)",
                    bairro_receb)
        bairro_e_de_outra_cidade = mapa_capital is not None and bairro_n in mapa_capital
        if bairro_valido(bairro) and not bairro_e_de_outra_cidade:
            return (ROTA_PADRAO_PARNAMIRIM,
                    "OK - Parnamirim, rota geral (bairro CT-e nao reconhecido, mas nao e de outra cidade)",
                    bairro)

    return None, "AGUARDANDO VALIDAÇÃO - Parnamirim, bairro fora dos grupos conhecidos, sem CEP valido, ou bairro pertence a outra cidade", bairro


def classifica_bairro_sga(bairro, bairro_receb, mapa_capital, cep_n=None, cep_receb_n=None,
                           base_capital=None, texto_extra="", grupo_sga=None):
    """CASCATA: BAIRRO RECEBEDOR > BAIRRO (CT-e) > texto livre (grupos
    especiais da aba 'Grupos Especiais') > CEP > BAIRRO RECEBEDOR
    generico > BAIRRO (CT-e) generico (bloqueado se for bairro conhecido
    de outra cidade) > texto > AGUARD."""
    bairro_n = normaliza_texto(bairro)
    bairro_receb_n = normaliza_texto(bairro_receb)
    texto_extra_n = normaliza_texto(texto_extra)

    for palavra, rota, classif in (grupo_sga or []):
        if palavra in bairro_receb_n:
            bairro_corrigido = bairro_receb
        elif palavra in bairro_n:
            bairro_corrigido = bairro
        elif texto_extra_n and palavra in texto_extra_n:
            bairro_corrigido = palavra
        else:
            continue
        if palavra in mapa_capital:
            setor, regiao = mapa_capital[palavra]
            return setor, regiao, f"OK - SGA, grupo especial '{rota}' (regra da planilha)", bairro_corrigido
        regiao_grupo = "Interior" if (classif and "interior" in classif.lower()) else "Capital e Met"
        return rota, regiao_grupo, f"OK - SGA, grupo especial '{rota}' (nao achei esse bairro no base_rotas, confira o nome)", bairro_corrigido

    if base_capital is not None:
        resultado = busca_por_cep(cep_receb_n, base_capital) or busca_por_cep(cep_n, base_capital)
        if resultado:
            setor, regiao, bairro_cadastrado = resultado
            return setor, regiao, "OK - SGA, fora dos grupos especiais (achado por CEP)", bairro_cadastrado

    if bairro_valido(bairro_receb):
        return (ROTA_PADRAO_SGA_INTERIOR, REGIAO_SGA_INTERIOR,
                "OK - SGA, fora dos grupos especiais (rota padrao Interior Norte - bairro recebedor)",
                bairro_receb)

    chaves_grupo = {palavra for palavra, _r, _c in (grupo_sga or [])}
    bairro_e_de_outra_cidade = bairro_n in mapa_capital and bairro_n not in chaves_grupo
    if bairro_valido(bairro) and not bairro_e_de_outra_cidade:
        return (ROTA_PADRAO_SGA_INTERIOR, REGIAO_SGA_INTERIOR,
                "OK - SGA, fora dos grupos especiais (rota padrao Interior Norte - bairro CT-e)",
                bairro)

    if texto_extra_n:
        return (ROTA_PADRAO_SGA_INTERIOR, REGIAO_SGA_INTERIOR,
                "OK - SGA, fora dos grupos especiais (rota padrao Interior Norte - so achou pista no endereco)",
                bairro)

    return None, "Capital e Met", "AGUARDANDO VALIDAÇÃO - SGA, bairro nao identificado ou pertence a outra cidade", bairro


def classificar_rota(cidade, bairro, cep, bairro_receb, cep_receb, base_rotas,
                      endereco=None, complemento_receb=None, grupos_especiais=None):
    cidade_n = normaliza_texto(cidade)
    bairro_n = normaliza_texto(bairro)
    bairro_receb_n = normaliza_texto(bairro_receb)
    cep_n = cep_para_numero(cep)
    cep_receb_n = cep_para_numero(cep_receb)

    base_capital = base_rotas[base_rotas["REGIAO"] == "Capital e Met"]
    base_interior = base_rotas[base_rotas["REGIAO"] == "Interior"]
    mapa_capital = dict(zip(base_capital["CHAVE_NORM"],
                             zip(base_capital["SETOR_OP"], base_capital["REGIAO"])))
    mapa_interior = dict(zip(base_interior["CHAVE_NORM"],
                              zip(base_interior["SETOR_OP"], base_interior["REGIAO"])))
    mapa_nome_capital = dict(zip(base_capital["CHAVE_NORM"], base_capital["CIDADE_BAIRRO"]))

    if cidade_n == "PARNAMIRIM":
        grupo_parnamirim = (grupos_especiais or {}).get("PARNAMIRIM", [])
        setor, motivo, bairro_corrigido = classifica_bairro_parnamirim(
            bairro, bairro_receb, cep_n, cep_receb_n, base_rotas, mapa_capital, grupo_parnamirim
        )
        return setor, REGIAO_PARNAMIRIM, motivo, bairro_corrigido

    if cidade_n == "SAO GONCALO DO AMARANTE":
        grupo_sga = (grupos_especiais or {}).get("SAO GONCALO DO AMARANTE", [])
        texto_extra = f"{endereco or ''} {complemento_receb or ''}"
        return classifica_bairro_sga(bairro, bairro_receb, mapa_capital, cep_n, cep_receb_n,
                                      base_capital, texto_extra, grupo_sga)

    if cidade_n not in CIDADES_CAPITAL_MET:
        # Fora da Capital/Met, o campo BAIRRO (CT-e) vem errado com
        # frequencia - as vezes com nome de bairro de outra cidade
        # (ex: "ALECRIM" ou "NOVA PARNAMIRIM" numa cidade do interior).
        # Nao da pra validar contra uma lista de bairros por cidade (a
        # base_rotas so mapeia por CIDADE/CEP nessas regioes), entao a
        # EXIBICAO (nao muda a classificacao de rota, que ja e so por
        # CIDADE/CEP aqui) segue uma cascata de 3 fontes, na mesma
        # ordem de prioridade ja usada no resto do robo pro RECEBEDOR
        # (ex: CEP RECEBEDOR, bairro recebedor em Natal): BAIRRO
        # RECEBEDOR > COMPLEMENTO RECEBEDOR > BAIRRO (CT-e) cru.
        # Antes so tentava COMPLEMENTO e depois BAIRRO (CT-e) - o
        # BAIRRO RECEBEDOR nunca era olhado nessa parte do codigo,
        # mesmo sendo geralmente mais confiavel; corrigido a pedido do
        # Samuel em 02/09/2026 (achado no CTRC PAN081989-1, que tinha
        # BAIRRO RECEBEDOR preenchido e mesmo assim saia vazio).
        if bairro_valido(bairro_receb):
            bairro_exibir = bairro_receb
        elif texto_descritivo_valido(complemento_receb):
            bairro_exibir = complemento_receb
        elif bairro_valido(bairro):
            bairro_exibir = bairro
        else:
            bairro_exibir = ""
        for cep_num in (cep_n, cep_receb_n):
            if not cep_num:
                continue
            for cep_ini, cep_fim, setor_exc, regiao_exc in EXCECOES_CEP_CIDADE_INTERIOR.get(cidade_n, []):
                if cep_ini <= cep_num <= cep_fim:
                    return setor_exc, regiao_exc, "OK - excecao CEP dentro da cidade", bairro_exibir
        if cidade_n in mapa_interior:
            setor, regiao = mapa_interior[cidade_n]
            return setor, regiao, "OK - cidade", bairro_exibir
        resultado = busca_por_cep(cep_n, base_interior) or busca_por_cep(cep_receb_n, base_interior)
        if resultado:
            return resultado[0], resultado[1], "OK - CEP (cidade nao cadastrada)", bairro_exibir
        return None, None, "VERIFIQUE BAIRRO/CIDADE", bairro_exibir

    # Natal: BAIRRO RECEBEDOR primeiro, BAIRRO (CT-e) so se o recebedor nao bateu
    if bairro_receb_n in mapa_capital:
        setor, regiao = mapa_capital[bairro_receb_n]
        return setor, regiao, "OK - bairro recebedor", bairro_receb
    if bairro_n in mapa_capital:
        setor, regiao = mapa_capital[bairro_n]
        return setor, regiao, "OK - bairro (recebedor nao bateu)", bairro

    resultado = busca_por_prefixo(bairro_receb_n, mapa_capital, mapa_nome_capital)
    if resultado:
        setor, regiao, bairro_cadastrado = resultado
        return setor, regiao, "OK - bairro recebedor truncado", bairro_cadastrado
    resultado = busca_por_prefixo(bairro_n, mapa_capital, mapa_nome_capital)
    if resultado:
        setor, regiao, bairro_cadastrado = resultado
        return setor, regiao, "OK - bairro truncado (recebedor nao bateu)", bairro_cadastrado

    # Antes de cair pro CEP (onde faixas sobrepostas de bairros vizinhos
    # podem devolver o bairro errado - ex: BOM PASTOR/QUINTAS/DIX SEPT
    # ROSADO tem CEPs que se cruzam), tenta bater o nome ignorando
    # espacos internos: cobre casos tipo "DIXSEPT ROSADO" (sem espaco,
    # como vem do CT-e) vs "DIX SEPT ROSADO" (como esta cadastrado na
    # base_rotas) - mesmo bairro, so grafia diferente. Pedido do Samuel
    # em 13/08/2026.
    mapa_capital_sem_espaco = {c.replace(" ", ""): v for c, v in mapa_capital.items()}
    mapa_nome_sem_espaco = {c.replace(" ", ""): n for c, n in mapa_nome_capital.items()}
    bairro_receb_se = bairro_receb_n.replace(" ", "")
    if bairro_receb_se in mapa_capital_sem_espaco:
        setor, regiao = mapa_capital_sem_espaco[bairro_receb_se]
        bairro_cadastrado = mapa_nome_sem_espaco.get(bairro_receb_se, bairro_receb)
        return setor, regiao, "OK - bairro recebedor (ignorando espacos)", bairro_cadastrado
    bairro_se = bairro_n.replace(" ", "")
    if bairro_se in mapa_capital_sem_espaco:
        setor, regiao = mapa_capital_sem_espaco[bairro_se]
        bairro_cadastrado = mapa_nome_sem_espaco.get(bairro_se, bairro)
        return setor, regiao, "OK - bairro CT-e (ignorando espacos, recebedor nao bateu)", bairro_cadastrado

    resultado = busca_por_cep(cep_n, base_capital) or busca_por_cep(cep_receb_n, base_capital)
    if resultado:
        setor, regiao, bairro_cadastrado = resultado
        return setor, regiao, "OK - CEP (bairro corrigido pelo CEP)", bairro_cadastrado

    resultado = (busca_bairro_em_texto(endereco, mapa_capital, mapa_nome_capital)
                 or busca_bairro_em_texto(complemento_receb, mapa_capital, mapa_nome_capital))
    if resultado:
        setor, regiao, bairro_cadastrado = resultado
        return setor, regiao, "OK - bairro achado no endereco/complemento (bairro corrigido)", bairro_cadastrado

    if cidade_n in mapa_capital:
        return None, "Capital e Met", "AGUARDANDO VALIDAÇÃO - bairro não identificado", bairro

    return None, None, "VERIFIQUE BAIRRO/CIDADE", bairro


# ------------------------------------------------------------------
# LEITURA E PREPARACAO DO SSWWEB
# ------------------------------------------------------------------
def _eh_linha_de_ruido_relatorio(linha):
    """Linhas que nao sao dado de verdade no relatorio impresso: titulo,
    separador de coluna, cabecalho repetido a cada pagina, linha vazia,
    quebra de pagina (\x0c)."""
    l = linha.strip()
    if not l or l == "\x0c":
        return True
    if set(l) <= {"-", "+"}:
        return True
    if l.startswith(("DOMINALOG EXPRESS", "ssw", "UNIDADE:", "CTRC REV/DEV")):
        return True
    return False


def ler_sswweb_relatorio_impresso(caminho):
    """Le a versao 'relatorio pra imprimir' do SSWWEB (titulo 'CTRCS
    DISPONIVEIS PARA ENTREGA') - formato de colunas alinhadas por
    posicao de texto, SEM ';' separando os campos. E o mesmo dado do
    export CSV, so que resumido/truncado pra caber numa pagina impressa.

    Limitacoes conhecidas desse formato (aceitas a pedido do Samuel em
    18/08/2026, pra nao precisar trocar de opcao no SSWWEB toda vez):
      - BAIRRO vem truncado em ~7 caracteres (ex: 'ALECRIM' vira
        'ALECRI') - o robo ja tem um fallback de bairro truncado
        (busca_por_prefixo) que cobre a maioria dos casos, entao a
        classificacao de rota continua funcionando, so um pouco menos
        precisa que com o CSV completo.
      - Nao existe BAIRRO RECEBEDOR/CEP RECEBEDOR separados - reusa o
        mesmo BAIRRO/CEP do CT-e pros dois.
      - Nao existe COMPLEMENTO RECEBEDOR (fica vazio).
      - Nao vem MANIFESTO/END nesse formato (fica de fora - o alerta de
        possivel fila de descarregamento so funciona com o export CSV).
      - Nao vem a descricao completa da ultima ocorrencia, so o codigo
        numerico - por isso o alerta de Falta de Documentacao (codigo
        52) e preenchido automaticamente com o texto padrao, ja que
        esse codigo especifico sempre significa isso.
    """
    with open(caminho, encoding="ISO-8859-1", newline="") as f:
        texto = f.read()
    linhas = texto.split("\r\n")

    ano_relatorio = agora_brasil().year
    m = re.search(r"\d{2}/\d{2}/(\d{2})\s+\d{2}:\d{2}", texto)
    if m:
        ano_relatorio = 2000 + int(m.group(1))

    limites = None
    for l in linhas:
        if l and "+" in l and set(l.strip()) <= {"-", "+"}:
            limites = [0]
            for i, ch in enumerate(l):
                if ch == "+":
                    limites.append(i + 1)
            limites.append(len(l) + 1)
            break
    if limites is None:
        raise ValueError(
            "Nao consegui identificar as colunas desse relatorio impresso "
            "(nao achei a linha separadora de '-' e '+'). O layout desse "
            "arquivo pode ter mudado - confira antes de rodar de novo."
        )

    def pega(linha, idx):
        ini, fim = limites[idx], limites[idx + 1]
        return linha[ini:fim].strip() if len(linha) > ini else ""

    linhas_dados = []
    linhas_devolucao_ignoradas = 0
    for l in linhas:
        if not l.strip() or l.strip().startswith("SETOR:") or _eh_linha_de_ruido_relatorio(l):
            continue
        # linhas de DEVOLUCAO vem com a palavra 'DEVOL' ocupando espaco
        # extra logo apos o CTRC, empurrando TODOS os campos seguintes -
        # nao da pra so adicionar 2 espacos (testado e ainda fica
        # errado, BAIRRO/CIDADE saem trocados). Mais seguro EXCLUIR
        # essas linhas do que deixar dado errado entrar na classificacao
        # de rota - devolucao nao e entrega normal mesmo. Pedido do
        # Samuel em 18/08/2026 (achado nos CTRC RN4.../DEVOL).
        if not l.startswith("  ") and l[:1].strip() and "-" in l[:14]:
            linhas_devolucao_ignoradas += 1
            continue
        ctrc = pega(l, 0)
        if not ctrc or not re.match(r"^[A-Z0-9][A-Z0-9\-]*$", ctrc):
            continue

        previ = pega(l, 8)
        ultoccor = pega(l, 15)
        codigo_ocorrencia, _, data_ocorrencia = ultoccor.partition("-")

        data_prevista_fmt = f"{previ}/{str(ano_relatorio)[2:]}" if "/" in previ else ""
        data_ocorrencia_fmt = f"{data_ocorrencia}/{str(ano_relatorio)[2:]}" if "/" in data_ocorrencia else ""
        codigo_ocorrencia = codigo_ocorrencia.strip()

        bairro = pega(l, 6)
        cep = pega(l, 7)
        cidade = pega(l, 5)

        linhas_dados.append({
            "CTRC": ctrc,
            "NFISCAL": pega(l, 1),
            "REMETENTE": "",
            "PAGADOR": pega(l, 2),
            "DESTINATARIO": pega(l, 3),
            "CIDADE": cidade,
            "BAIRRO": bairro,
            "CEP": cep,
            "ENDERECO": pega(l, 4),
            "BAIRRO RECEBEDOR": bairro,
            "CEP RECEBEDOR": cep,
            "CIDADE RECEBEDOR": cidade,
            "COMPLEMENTO RECEBEDOR": "",
            "DATA PREVISTA": data_prevista_fmt,
            "ULT. OCORRENCIA": codigo_ocorrencia,
            "DESCRICAO ULT. OCORRENCIA": "FALTA DE DOCUMENTACAO" if codigo_ocorrencia == "52" else "",
            "DATA ULT. OCORRENCIA": data_ocorrencia_fmt,
            "KG REAL": pega(l, 11),
            "M3": pega(l, 12),
            "PESO CALCULO": pega(l, 11),
        })

    df = pd.DataFrame(linhas_dados)
    if len(df):
        df = df[df["CTRC"].notna() & (df["CTRC"].astype(str).str.strip() != "")].copy()
    log_detalhe(f"[relatorio impresso] {len(df)} linha(s) lida(s) do formato 'CTRCs disponiveis pra entrega'")
    if linhas_devolucao_ignoradas:
        log_detalhe(
            f"[relatorio impresso] {linhas_devolucao_ignoradas} linha(s) de DEVOLUCAO "
            f"ignorada(s) - esse formato desalinha os campos nessas linhas, entao "
            f"prefere nao processar a processar com dado errado."
        )
        print(
            f"  [aviso] {linhas_devolucao_ignoradas} CTRC(s) de devolução não foram lidos "
            f"desse relatório (formato desalinha os campos nessas linhas). Se precisar "
            f"deles, use o export CSV ou confira manualmente."
        )
    df.attrs["linhas_devolucao_ignoradas"] = linhas_devolucao_ignoradas
    return df


def ler_sswweb_completo(caminho):
    """Le o .sswweb com TODAS as colunas necessarias pra versao
    so-SSWWEB. Corrige o bug de apostrofo (#39;) que quebra o
    alinhamento das colunas quando o bairro tem apostrofo.

    Aceita DOIS formatos, detectados automaticamente (pedido do Samuel
    em 18/08/2026, pra nao precisar escolher a opcao certa no SSWWEB
    toda vez):
      1. Export CSV padrao (separado por ';') - formato completo,
         preferido, usado ate agora.
      2. Relatorio impresso 'CTRCS DISPONIVEIS PARA ENTREGA' (colunas
         alinhadas por posicao, sem ';') - formato mais resumido, ver
         ler_sswweb_relatorio_impresso() pras limitacoes.
    """
    with open(caminho, encoding="ISO-8859-1", newline="") as f:
        texto = f.read()

    primeira_linha = texto.split("\r\n", 1)[0].split("\n", 1)[0]
    if ";" not in primeira_linha and "CTRCS DISPONIVEIS" in texto[:600].upper():
        log_detalhe("Formato detectado: relatorio impresso (sem ';') - usando leitor alternativo.")
        df = ler_sswweb_relatorio_impresso(caminho)
        valida_colunas(
            df, COLUNAS_SSWWEB_COMPLETO, "SSWWEB (relatorio impresso)",
            obrigatorias=["CTRC", "CIDADE", "BAIRRO", "CEP", "ULT. OCORRENCIA", "DATA PREVISTA"]
        )
        colunas = [c for c in COLUNAS_SSWWEB_COMPLETO if c in df.columns]
        return df[colunas].copy()

    log_detalhe("Lendo SSWWEB completo (versao so-SSWWEB)...")
    n_corrigidos = texto.count("#39;")
    if n_corrigidos:
        texto = texto.replace("#39;", "'")
        log_detalhe(f"Corrigido artefato de apostrofo em {n_corrigidos} lugar(es) antes do parse.")

    linhas = []
    reader = csv.DictReader(io.StringIO(texto), delimiter=";")
    for row in reader:
        linhas.append(row)
    df = pd.DataFrame(linhas)

    valida_colunas(
        df, COLUNAS_SSWWEB_COMPLETO, "SSWWEB completo",
        obrigatorias=["CTRC", "CIDADE", "BAIRRO", "CEP", "ULT. OCORRENCIA", "DATA PREVISTA"]
    )

    colunas = [c for c in COLUNAS_SSWWEB_COMPLETO if c in df.columns]
    df = df[colunas].copy()
    df = df[df["CTRC"].notna() & (df["CTRC"].str.strip() != "")].copy()
    return df


def filtra_movimentacao_interna(df):
    """Exclui movimentacoes internas (remetente Dominalog) - minuta,
    salvado, transferencia entre unidades da propria empresa.

    Usada so no fluxo antigo (funcao main(), CLI - nao e o robo web).
    O robo web (app.py) usa marca_movimentacao_interna, que MANTEM
    essas linhas em vez de excluir (ver essa funcao pro motivo)."""
    if "REMETENTE" not in df.columns:
        return df
    antes = len(df)
    eh_interno = df["REMETENTE"].fillna("").str.upper().str.contains("DOMINALOG")
    df = df[~eh_interno].copy()
    n_removidos = antes - len(df)
    if n_removidos:
        log_detalhe(f"Removidas {n_removidos} movimentacao(oes) interna(s) (remetente Dominalog).")
    return df


def marca_movimentacao_interna(df):
    """Versao do robo web de filtra_movimentacao_interna: em vez de
    EXCLUIR movimentacoes internas (minuta, salvado, transferencia
    entre unidades da propria empresa), so MARCA a coluna
    EH_MOVIMENTACAO_INTERNA e deixa a linha seguir pelo resto do
    pipeline. Antes essas linhas sumiam sem deixar rastro (so contavam
    num numero no resumo da tela); agora aparecem na aba Aguardando
    Validação (ver separa_liberados_retidos_nao_liberados), pra nao
    ficarem escondidas. Pedido do Samuel em 07/09/2026.

    Olha REMETENTE *e* DESTINATARIO conterem "DOMINALOG" - nao so
    REMETENTE. Motivo (confirmado pelo Samuel em 07/09/2026): existe
    o caso inverso da minuta/transferencia normal (saindo da
    Dominalog), que e o "salvado" ENTRANDO - ex: fornecedor manda uma
    peca de reposicao pro cliente sem devolver a mercadoria toda
    (quebrou um espelho, um pe de movel...). Nesse caso quem aparece
    como REMETENTE e o fornecedor/transportadora, e a DOMINALOG e o
    DESTINATARIO. E o mesmo tipo de movimentacao interna, so que na
    direcao contraria - por isso tambem precisa cair em Aguardando
    Validação em vez de seguir a rota normal."""
    tem_remetente = "REMETENTE" in df.columns
    tem_destinatario = "DESTINATARIO" in df.columns
    if not tem_remetente and not tem_destinatario:
        df["EH_MOVIMENTACAO_INTERNA"] = False
        return df

    eh_interna = pd.Series(False, index=df.index)
    if tem_remetente:
        eh_interna = eh_interna | df["REMETENTE"].fillna("").str.upper().str.contains("DOMINALOG")
    if tem_destinatario:
        eh_interna = eh_interna | df["DESTINATARIO"].fillna("").str.upper().str.contains("DOMINALOG")
    df["EH_MOVIMENTACAO_INTERNA"] = eh_interna

    n_internas = int(df["EH_MOVIMENTACAO_INTERNA"].sum())
    if n_internas:
        log_detalhe(f"{n_internas} movimentacao(oes) interna(s) (remetente ou destinatario Dominalog) - vao pra Aguardando Validação.")
    return df


def calcula_liberada_rota(df):
    """JA_LIBERADA_ROTA pelo codigo da ultima ocorrencia + excecao do
    codigo 80 quando o CTRC foi emitido pela propria RN4."""
    codigo_n = df["ULT. OCORRENCIA"].fillna("").astype(str).str.strip().str.split(".").str[0]
    liberada_por_codigo = codigo_n.isin(CODIGOS_LIBERADA_ROTA)
    eh_excecao_80 = (codigo_n == "80") & (
        df["CTRC"].fillna("").str.strip().str.upper().str.startswith(PREFIXO_CTRC_EMISSAO_PROPRIA)
    )
    df["CODIGO_ULT_OCORRENCIA"] = codigo_n
    df["JA_LIBERADA_ROTA"] = liberada_por_codigo | eh_excecao_80
    return df


def calcula_retido(df):
    """RETIDO = codigo de ultima ocorrencia em CODIGOS_RETIDO (56 =
    mercadoria retida pela fiscalizacao). Precisa rodar DEPOIS de
    calcula_liberada_rota (usa a coluna CODIGO_ULT_OCORRENCIA que ela
    calcula). Pedido do Samuel em 21/08/2026, mesma logica do robo
    desktop."""
    df["RETIDO"] = df["CODIGO_ULT_OCORRENCIA"].isin(CODIGOS_RETIDO)
    return df


def calcula_alerta_fila_descarregamento(df, hoje=None):
    """Marca CTRCs em codigo 84 (chegada na unidade) que podem ainda
    estar na fila de descarregamento, em vez de ja terem sido
    descarregados e disponiveis pra rota.

    Nao da pra saber com certeza (o .sswweb nao tem hora nem status de
    descarregamento) - mas 2 sinais juntos, validados contra o BI em
    20/08/2026 (0 erro no teste com dados reais), aproximam bem:
      1. TODOS os CTRCs do mesmo manifesto (MANIFESTO/END) ainda estao
         em codigo 84 - nenhum ja avancou pra outro codigo (85, 10, 31,
         32...). Se algum ja avancou, o manifesto ja foi mexido/aberto,
         entao os que ainda estao em 84 ja devem estar disponiveis
         tambem.
      2. A data da ultima ocorrencia desses CTRCs e HOJE - descarregar
         leva no maximo 1 dia (na pratica), entao codigo 84 datado de
         um dia anterior ja deve ter sido descarregado.

    So dispara quando o manifesto tem 2+ CTRCs (manifesto de 1 CTRC so
    nao da sinal nenhum - testado e dava falso alarme). Pedido do
    Samuel em 20/08/2026, depois de comparar contra o BI e achar 20
    CTRCs que o robo contava como liberado mas ainda estavam em fila."""
    if hoje is None:
        hoje = datetime.date.today()

    df["ALERTA_FILA_DESCARREGAMENTO"] = False
    if "MANIFESTO/END" not in df.columns:
        return df

    manifesto = df["MANIFESTO/END"].fillna("").astype(str).str.strip()
    data_eh_hoje = df["DATA_ULT_OCORRENCIA_DT"].apply(
        lambda d: pd.notna(d) and d.date() == hoje
    )

    for chave, grupo_idx in df.groupby(manifesto).groups.items():
        if not chave or len(grupo_idx) < 2:
            continue
        codigos_grupo = set(df.loc[grupo_idx, "CODIGO_ULT_OCORRENCIA"])
        if codigos_grupo != {"84"}:
            continue
        for idx in grupo_idx:
            if data_eh_hoje.loc[idx]:
                df.loc[idx, "ALERTA_FILA_DESCARREGAMENTO"] = True

    return df


def calcula_liberado_fiscalizacao_recente(df, hoje=None):
    """Marca CTRCs em codigo 58 (MERCADORIA LIBERADA PELA FISCALIZACAO)
    cuja ultima ocorrencia e de HOJE ou ONTEM - fica visivel na tela de
    previa pra usuario validar e ja colocar em rota, em vez de precisar
    reparar sozinho que aquele CTRC saiu da fiscalizacao ha pouco tempo.
    So considera hoje/ontem porque uma liberacao mais antiga ja deveria
    ter sido notada e roteirizada em algum momento anterior - o aviso
    aqui e pra pegar liberacao RECENTE, nao virar uma lista permanente
    de tudo que algum dia passou pela fiscalizacao. Pedido do Samuel em
    02/09/2026.

    Usa agora_brasil() (nao datetime.date.today()) pra "hoje" bater com
    o horario de Natal/RN mesmo o servidor rodando em UTC - mesmo motivo
    documentado em agora_brasil()."""
    if hoje is None:
        hoje = agora_brasil().date()
    ontem = hoje - datetime.timedelta(days=1)

    codigo = df.get("CODIGO_ULT_OCORRENCIA")
    data_ocorrencia = df.get("DATA_ULT_OCORRENCIA_DT")
    if codigo is None or data_ocorrencia is None:
        df["LIBERADO_FISCALIZACAO_RECENTE"] = False
        return df

    eh_codigo_58 = codigo == "58"
    eh_recente = data_ocorrencia.apply(lambda d: pd.notna(d) and d.date() in (hoje, ontem))
    df["LIBERADO_FISCALIZACAO_RECENTE"] = eh_codigo_58 & eh_recente
    return df


def formata_detalhe_ocorrencia(descricao, data_ocorrencia, em_fila_descarregamento):
    """Acrescenta '(DOCA DD/MM)' na descricao da ultima ocorrencia
    quando o CTRC pode ainda estar na fila de descarregamento (ver
    calcula_alerta_fila_descarregamento) - sinal visual direto no texto
    que ja aparece em todo lugar (Excel, romaneio, previa web), sem
    precisar de coluna extra. Pedido do Samuel em 20/08/2026."""
    texto = str(descricao or "").strip()
    if em_fila_descarregamento and pd.notna(data_ocorrencia) and hasattr(data_ocorrencia, "strftime"):
        return f"{texto} (DOCA {data_ocorrencia.strftime('%d/%m')})"
    return texto


def calcula_resumo_prazo(data_prevista, hoje=None):
    """RESUMO com 5 categorias, validado empiricamente contra o BI
    (100% de acerto, 1199 linhas):
      diff < 0 -> ATRASO | diff == 0 -> VENCE HOJE | diff == 1 -> AMANHA
      diff == 2 -> VENCE EM 2 DIAS | diff >= 3 -> VENCE FUTURO"""
    if hoje is None:
        hoje = datetime.date.today()
    if pd.isna(data_prevista):
        return None
    data = data_prevista
    if isinstance(data, datetime.datetime):
        data = data.date()
    if not isinstance(data, datetime.date):
        return None
    diff = (data - hoje).days
    if diff < 0:
        return "ATRASO"
    if diff == 0:
        return "VENCE HOJE"
    if diff == 1:
        return "VENCE AMANHA"
    if diff == 2:
        return "VENCE EM 2 DIAS"
    return "VENCE FUTURO"


def calcula_status_prazo_dias(data_prevista, hoje=None):
    """STATUS PRAZO: dias corridos com sinal, mesmo calculo de base que
    ja alimenta RESUMO (diff = data_prevista - hoje), so com o sinal
    INVERTIDO em relacao a esse - pedido do Samuel em 08/09/2026:
    NEGATIVO = ainda faltam esses dias pra vencer; POSITIVO = ja
    venceu ha esses dias (dias em atraso); 0 = vence hoje.

    Coluna que ja existia como calculo interno so pro romaneio em PDF
    (calcula_status_vencimento, que usa o sinal ao contrario: positivo
    pra 'ainda falta') - agora tambem vira coluna de verdade no Excel."""
    if hoje is None:
        hoje = datetime.date.today()
    if pd.isna(data_prevista):
        return None
    data = data_prevista
    if isinstance(data, datetime.datetime):
        data = data.date()
    if not isinstance(data, datetime.date):
        return None
    return (hoje - data).days


def prepara_dados(caminho_sswweb, caminho_base_rotas):
    df = ler_sswweb_completo(caminho_sswweb)
    linhas_devolucao_ignoradas = df.attrs.get("linhas_devolucao_ignoradas", 0)
    total_lido = len(df)

    uf_predominante = detecta_uf_predominante(df)

    df = marca_movimentacao_interna(df)
    total_apos_filtro_interno = int((~df["EH_MOVIMENTACAO_INTERNA"]).sum())

    df["DATA_PREVISTA_DT"] = pd.to_datetime(df["DATA PREVISTA"], format="%d/%m/%y", errors="coerce")
    df["DATA_ULT_OCORRENCIA_DT"] = pd.to_datetime(df["DATA ULT. OCORRENCIA"], format="%d/%m/%y", errors="coerce")

    df = calcula_liberada_rota(df)
    df = calcula_retido(df)
    df = calcula_alerta_fila_descarregamento(df)
    df = calcula_liberado_fiscalizacao_recente(df)

    df["RESUMO"] = df["DATA_PREVISTA_DT"].apply(calcula_resumo_prazo)
    df["STATUS_PRAZO_DIAS"] = df["DATA_PREVISTA_DT"].apply(calcula_status_prazo_dias)
    df["DIAS_PARADA_PISO"] = df["DATA_ULT_OCORRENCIA_DT"].apply(dias_parados_desde)

    df["CIDADE_NORM"] = df["CIDADE"].apply(normaliza_texto)
    df["EH_CAPITAL_MET"] = df["CIDADE_NORM"].isin(CIDADES_CAPITAL_MET)

    base_rotas = ler_base_rotas(caminho_base_rotas)
    grupos_especiais = ler_grupos_especiais(caminho_base_rotas)
    rota_para_classificacao = monta_rota_para_classificacao(base_rotas, grupos_especiais)

    resumo = {
        "total_lido": total_lido,
        "total_apos_filtro_interno": total_apos_filtro_interno,
        "total_liberado_rota": int(df["JA_LIBERADA_ROTA"].sum()),
        "total_nao_liberado": int((~df["JA_LIBERADA_ROTA"]).sum()),
        "total_retido": int(df["RETIDO"].sum()),
        "total_possivel_fila_descarregamento": int(df["ALERTA_FILA_DESCARREGAMENTO"].sum()),
        "distrib_resumo": df["RESUMO"].value_counts(dropna=False).to_dict(),
        "distrib_capital_interior": df["EH_CAPITAL_MET"].value_counts().to_dict(),
        "linhas_devolucao_ignoradas": linhas_devolucao_ignoradas,
        "uf_predominante": uf_predominante,
    }
    return df, base_rotas, grupos_especiais, rota_para_classificacao, resumo


def classifica_rotas(df, base_rotas, grupos_especiais=None, uf_predominante=None, forcar_modo_setor=False):
    """Se uf_predominante vier preenchida e for diferente de RN, o
    arquivo e de outra filial (ex: PE) - o base_rotas.xlsx so tem
    cadastro do RN, entao NEM TENTA a cascata de classificar_rota (ela
    nunca ia bater mesmo, e so ia jogar tudo pra "VERIFIQUE BAIRRO/
    CIDADE" -> Aguardando Validação). ROTA e REGIAO_ROTA ficam em
    branco (region e rota "vazios", como pedido pelo Samuel em
    07/09/2026); BAIRRO_CORRIGIDO ainda usa a mesma cascata de exibicao
    ja usada pro interior do RN (BAIRRO RECEBEDOR > COMPLEMENTO
    RECEBEDOR > BAIRRO CT-e), que nao depende do base_rotas.

    forcar_modo_setor (pedido do Samuel em 01/10/2026): o mesmo caminho
    acima, so que ligado manualmente pelo usuario na tela de escolha de
    base (app.py, tela_escolha_base_rotas, "Opção 3"), pra quem NAO tem
    base de rotas nenhuma cadastrada - mesmo num arquivo do proprio RN.
    Reusa a mesma cascata "ROTA = coluna SETOR" de baixo, so muda o
    texto do MOTIVO_ROTA pra deixar claro que foi uma escolha do usuario
    (sem base), nao uma deteccao automatica de UF diferente."""
    se_uf_diferente = bool(uf_predominante) and uf_predominante != UF_ESPERADA
    if forcar_modo_setor or se_uf_diferente:
        def _bairro_sem_base_rotas(r):
            bairro_receb = r.get("BAIRRO RECEBEDOR")
            complemento_receb = r.get("COMPLEMENTO RECEBEDOR")
            bairro = r.get("BAIRRO")
            if bairro_valido(bairro_receb):
                return bairro_receb
            if texto_descritivo_valido(complemento_receb):
                return complemento_receb
            if bairro_valido(bairro):
                return bairro
            return ""

        # ROTA = coluna SETOR, direto (sem passar por CEP/bairro/
        # complemento - so cidade e cidade recebedor entram no resto da
        # classificacao, que fica em BAIRRO_CORRIGIDO abaixo, sem
        # nenhuma relacao com ROTA). SETOR ja vem pronto do SSWWEB
        # representando a roteirizacao/regiao daquela filial - testado
        # com dado real do Samuel (arquivo de PE): cidade pequena tem 1
        # setor so, mas cidade grande (RECIFE, JABOATAO DOS GUARARAPES,
        # OLINDA, PAULISTA) tem VARIOS setores diferentes dentro dela -
        # ou seja, SETOR e mais fino que cidade, e usar so 1 valor por
        # cidade (ex: moda) perderia rota de verdade. Por isso e por
        # linha, nao agregado por cidade. Pedido do Samuel em 07/09/2026.
        if "SETOR" in df.columns:
            rota_setor = df["SETOR"].fillna("").astype(str).str.strip()
            df["ROTA"] = rota_setor.where(rota_setor != "", None)
        else:
            df["ROTA"] = None
        df["REGIAO_ROTA"] = None
        if se_uf_diferente:
            df["MOTIVO_ROTA"] = (
                f"UF predominante do arquivo ({uf_predominante}) diferente de RN - base_rotas.xlsx ignorada, "
                f"ROTA = coluna SETOR"
            )
        else:
            df["MOTIVO_ROTA"] = (
                "Modo somente Setor ativado pelo usuário (sem base de rotas vinculada) - "
                "ROTA = coluna SETOR"
            )
        df["BAIRRO_CORRIGIDO"] = df.apply(_bairro_sem_base_rotas, axis=1)
        return df

    resultados = df.apply(
        lambda r: classificar_rota(
            cidade=r["CIDADE"], bairro=r["BAIRRO"], cep=r["CEP"],
            bairro_receb=r.get("BAIRRO RECEBEDOR"), cep_receb=r.get("CEP RECEBEDOR"),
            base_rotas=base_rotas, endereco=r.get("ENDERECO"),
            complemento_receb=r.get("COMPLEMENTO RECEBEDOR"),
            grupos_especiais=grupos_especiais,
        ),
        axis=1, result_type="expand",
    )
    resultados.columns = ["ROTA", "REGIAO_ROTA", "MOTIVO_ROTA", "BAIRRO_CORRIGIDO"]
    for c in resultados.columns:
        df[c] = resultados[c]
    return df


def limpa_nota_fiscal_texto(valor):
    """Nota Fiscal vem do sswweb com espacos de preenchimento (ex:
    '   153573 ') - so tira o espaco e mantem como TEXTO (nao numero),
    porque como numero o Excel formata com separador de milhar
    ('153.573'), que fica estranho pra um numero de identificacao (nao
    e uma quantidade). Pedido do Samuel em 18/08/2026 - reverteu a
    tentativa anterior de deixar como numero de verdade."""
    if valor is None:
        return None
    txt = str(valor).strip()
    return txt if txt else None


def monta_colunas_finais(df, uf_predominante=None):
    detalhe_ocorrencia = df.apply(
        lambda r: formata_detalhe_ocorrencia(
            r.get("DESCRICAO ULT. OCORRENCIA"),
            r.get("DATA_ULT_OCORRENCIA_DT"),
            r.get("ALERTA_FILA_DESCARREGAMENTO", False),
        ),
        axis=1,
    )

    # UF fora do esperado (RN, ou a UF predominante do proprio arquivo
    # quando ele nao e do RN - ver detecta_uf_predominante) = sinal de
    # que pode ser DEVOLUCAO indo pra outro estado, nao entrega normal
    # - o motorista/expedicao precisa confirmar antes de rotear. Pedido
    # do Samuel em 18/08/2026 (achado com um CTRC de devolucao Carrefour
    # indo pra Jaboatao dos Guararapes/PE). So dispara quando a UF vem
    # preenchida (o relatorio impresso nao tem esse campo, entao fica
    # sem alerta nesse formato).
    #
    # Antes de 07/09/2026 comparava sempre contra UF_ESPERADA ("RN")
    # fixo - pra um arquivo de outra filial (ex: PE) isso marcava
    # praticamente TODA linha como "fora do RN", inundando a previa.
    # Agora compara contra uf_predominante quando ela vier preenchida
    # (fica dinamico); sem uf_predominante, cai no comportamento antigo
    # (fixo em RN) pra nao quebrar nenhum uso que ainda nao passa esse
    # argumento.
    uf_esperada_alerta = uf_predominante or UF_ESPERADA
    uf_receb = df.get("UF RECEBEDOR")
    uf_receb_norm = (
        uf_receb.fillna("").astype(str).str.strip().str.upper()
        if uf_receb is not None else pd.Series("", index=df.index)
    )
    if uf_receb is not None:
        alerta_validacao = uf_receb.apply(
            lambda uf: "🚩 USUÁRIO PRECISA VALIDAR O CTRC (destino fora do esperado)"
            if str(uf or "").strip() and normaliza_texto(uf) != uf_esperada_alerta
            else ""
        )
    else:
        alerta_validacao = pd.Series("", index=df.index)

    eh_interna = df.get("EH_MOVIMENTACAO_INTERNA")
    if eh_interna is not None:
        alerta_validacao = alerta_validacao.where(
            ~eh_interna, "🚩 MOVIMENTAÇÃO INTERNA (MINUTA/SALVADO) - CONFIRME SE DEVE ENTRAR NA ROTA"
        )

    final = pd.DataFrame({
        "CTRC": df["CTRC"],
        "NOTA FISCAL": df.get("NFISCAL").apply(limpa_nota_fiscal_texto) if "NFISCAL" in df.columns else None,
        "ALERTA": alerta_validacao,
        "REMETENTE": df.get("REMETENTE"),
        "PAGADOR": df.get("PAGADOR"),
        "CLIENTE": df.get("DESTINATARIO"),
        "CIDADE RECEBEDOR": df.get("CIDADE RECEBEDOR"),
        "BAIRRO RECEBEDOR": df["BAIRRO_CORRIGIDO"],
        "CEP": df.get("CEP RECEBEDOR"),
        "ROTA": df["ROTA"],
        "REGIAO_ROTA": df["REGIAO_ROTA"],
        "PREVISAO ENTREGA": df["DATA_PREVISTA_DT"],
        "RESUMO": df["RESUMO"],
        # STATUS PRAZO: dias corridos com sinal - negativo = ainda
        # faltam esses dias pra vencer, positivo = venceu ha esses dias
        # (0 = vence hoje). Ver calcula_status_prazo_dias. Pedido do
        # Samuel em 08/09/2026.
        "STATUS PRAZO": df["STATUS_PRAZO_DIAS"],
        "DATA DA ÚLTIMA OCORRÊNCIA": df["DATA_ULT_OCORRENCIA_DT"],
        # DETALHE ÚLTIMA OCORRÊNCIA e a UNICA coluna com esse conteudo -
        # ate 25/08/2026 existia tambem uma "ÚLTIMA OCORRÊNCIA" com o
        # mesmo valor duplicado (sobra de quando o nome da coluna mudou
        # e a antiga nao foi removida). Achado pelo Samuel no Excel
        # gerado. aplica_destaque_alerta_offset e a previa web (app.py)
        # foram ajustados pra ler por esse nome.
        "DETALHE ÚLTIMA OCORRÊNCIA": detalhe_ocorrencia,
        "CODIGO ULTIMA OCORRENCIA": df["CODIGO_ULT_OCORRENCIA"],
        "DIAS PARADA EM PISO": df["DIAS_PARADA_PISO"],
        "MANIFESTO/END": df.get("MANIFESTO/END"),
        "ALERTA_FILA_DESCARREGAMENTO": df.get("ALERTA_FILA_DESCARREGAMENTO", False),
        # Sem "_" na frente (nao pode - precisa sobreviver ao
        # colunas_saida de separa_liberados_retidos_nao_liberados pra
        # continuar disponivel em df_capital_interior, que e o que a
        # tela previa do app.py usa pro aviso de liberacao recente).
        # Quem tira essa coluna do Excel final e o colunas_excluir de
        # gerar_combinado_sswweb (mesmo mecanismo ja usado pra
        # ALERTA_FILA_DESCARREGAMENTO/ALERTA/REMETENTE) - a DATA DA
        # ÚLTIMA OCORRÊNCIA acima ja cobre a informacao de data e os
        # romaneios nao usam essa coluna. Pedido do Samuel em
        # 07/09/2026.
        "LIBERADO_FISCALIZACAO_RECENTE": df.get("LIBERADO_FISCALIZACAO_RECENTE", False),
        "KG REAL": df.get("KG REAL").apply(converte_numero_br) if "KG REAL" in df.columns else None,
        "M3": df.get("M3").apply(converte_numero_br) if "M3" in df.columns else None,
        "PESO CALCULO": df.get("PESO CALCULO").apply(converte_numero_br) if "PESO CALCULO" in df.columns else None,
        "_CIDADE_NORM": df["CIDADE_NORM"],
        "_JA_LIBERADA_ROTA": df["JA_LIBERADA_ROTA"],
        "_CODIGO_ULT_OCORRENCIA": df["CODIGO_ULT_OCORRENCIA"],
        "_RETIDO": df["RETIDO"],
        "_UF_RECEBEDOR": uf_receb_norm,
        "_EH_MOVIMENTACAO_INTERNA": df.get("EH_MOVIMENTACAO_INTERNA", False),
    })
    return final


def separa_capital_interior(df_final, rota_para_classificacao=None):
    """Quem entra em Capital/Interior: liberado pra rota E ainda nao saiu
    em rota de entrega (codigo 85)."""
    ja_em_rota = df_final["_CODIGO_ULT_OCORRENCIA"] == "85"
    roteirizavel = df_final[df_final["_JA_LIBERADA_ROTA"] & ~ja_em_rota].copy()
    roteirizavel["CAPITAL_INTERIOR"] = roteirizavel.apply(
        lambda r: classifica_capital_interior(
            r["_CIDADE_NORM"], r["ROTA"], normaliza_texto(r["BAIRRO RECEBEDOR"]),
            rota_para_classificacao,
        ),
        axis=1,
    )
    colunas_saida = [c for c in roteirizavel.columns if not c.startswith("_")]
    return roteirizavel[colunas_saida]


def separa_liberados_retidos_nao_liberados(df_final, rota_para_classificacao=None, uf_predominante=None,
                                            forcar_modo_setor=False):
    """Versao estendida de separa_capital_interior: alem do Capital/
    Interior/Aguardando Validação (calculado igual, exceto pela
    excecao do PAGADOR == CLIENTE - ver _com_capital_interior), agora
    tambem separa RETIDOS, NAO LIBERADAS, NOTAS EM ROTA e A CAMINHO,
    pra virarem abas extras no arquivo unico. Pedido do Samuel em
    21/08/2026 (retidos/nao liberadas), 25/08/2026 (notas em rota) e
    26/08/2026 (a caminho):
      - retidos: RETIDO == True (codigo 56, mercadoria retida pela
        fiscalizacao).
      - a_caminho: dentro do que nao esta liberado/retido/em rota,
        quem tem DETALHE ÚLTIMA OCORRÊNCIA batendo com
        DESCRICOES_A_CAMINHO (ver a constante pra lista completa) -
        CTRC que ja saiu de uma unidade mas ainda nao chegou na unidade
        de entrega, seja em transito entre unidades ou parado num
        posto de fiscalizacao selando notas.
      - nao_liberadas: o resto - nao liberado, nao retido, nao em
        rota e nao em transferencia (codigo 85 continua de fora dessa
        aba, igual sempre foi pro roteirizavel).
      - notas_em_rota: codigo 85 (SAIDA EM ROTA DE ENTREGA) - antes
        ficava de fora do arquivo inteiro, agora ganha aba propria com
        a mesma dinamica do Capital/Interior (mesmas colunas, CAPITAL_
        INTERIOR calculado igual), so que DIAS PARADA EM PISO vira
        DIAS EM ROTA na hora de exportar (ver gerar_combinado_sswweb) -
        pra quem esta em rota, a DATA DA ULTIMA OCORRENCIA e a data que
        SAIU em rota, entao o mesmo calculo de dias uteis passa a
        significar "ha quantos dias esta em rota" em vez de "ha quantos
        dias esta parado".

    uf_predominante (pedido do Samuel em 07/09/2026): quando vem
    preenchida e diferente de RN (arquivo de outra filial, ex: PE),
    duas coisas mudam:
      - CTRCs cuja UF RECEBEDOR e diferente da UF predominante do
        proprio arquivo (ex: 1 CTRC de SP misturado num arquivo
        majoritariamente PE) vao pra "Aguardando Validação" - mesma
        ideia do alerta antigo "fora do RN", so que dinamico em vez de
        travado em RN.
      - Movimentacoes internas (minuta/salvado/transferencia entre
        unidades - EH_MOVIMENTACAO_INTERNA) SEMPRE vao pra "Aguardando
        Validação", tiradas de qualquer outro grupo (retidos/nao
        liberadas/notas em rota/a caminho/roteirizavel) onde cairiam
        naturalmente - antes eram excluidas do arquivo inteiro.

    forcar_modo_setor (pedido do Samuel em 01/10/2026): liga a mesma
    logica de CAPITAL_INTERIOR="Interior" pra todo mundo (ver
    classifica_rotas), mesmo em arquivo do proprio RN sem UF diferente
    detectada - usado quando o usuario escolheu manualmente "nao tenho
    base de rotas" na tela inicial. A checagem de UF RECEBEDOR diferente
    (2 paragrafos acima) continua só rodando quando uf_predominante
    vier preenchida de verdade (guarda "and uf_predominante" abaixo) -
    sem isso, um arquivo sem a coluna UF RECEBEDOR nesse modo forçado
    (uf_predominante=None) marcaria TODO MUNDO como "Aguardando
    Validação" por engano (None comparado com qualquer UF sempre bate
    diferente)."""
    ja_em_rota = df_final["_CODIGO_ULT_OCORRENCIA"] == "85"
    eh_retido = df_final["_RETIDO"]
    eh_a_caminho = df_final["DETALHE ÚLTIMA OCORRÊNCIA"].apply(normaliza_texto).isin(DESCRICOES_A_CAMINHO)
    eh_minuta = df_final.get("_EH_MOVIMENTACAO_INTERNA", pd.Series(False, index=df_final.index)).fillna(False)

    ignora_base_rotas = forcar_modo_setor or (bool(uf_predominante) and uf_predominante != UF_ESPERADA)

    def _com_capital_interior(subset):
        subset = subset.copy()
        if ignora_base_rotas:
            # ROTA agora vem preenchida com a coluna SETOR nesse modo
            # (ver classifica_rotas) - trata todo mundo como "Interior"
            # pra reaproveitar a mesma logica de gerar_romaneios_pdf que
            # ja existia pro RN: rota pequena (cabe numa folha sozinha)
            # entra automatico num PDF agrupado com outras rotas
            # pequenas ("coletivo"), rota grande fica com PDF proprio
            # ("individual") - sem precisar de nocao de capital pra
            # filial de fora do RN (nao daria pra saber qual cidade e
            # "capital" so olhando SETOR). Quem ficar sem SETOR (ROTA
            # vazia) cai em Aguardando Validação, tratado abaixo igual
            # ao caso de UF diferente. Pedido do Samuel em 07/09/2026.
            subset["CAPITAL_INTERIOR"] = "Interior"
        else:
            subset["CAPITAL_INTERIOR"] = subset.apply(
                lambda r: classifica_capital_interior(
                    r["_CIDADE_NORM"], r["ROTA"], normaliza_texto(r["BAIRRO RECEBEDOR"]),
                    rota_para_classificacao,
                ),
                axis=1,
            )
        # PAGADOR igual ao CLIENTE (destinatario) e sinal de devolucao -
        # em entrega normal quem paga o frete costuma ser diferente de
        # quem recebe a mercadoria; quando os dois batem, geralmente e o
        # proprio remetente original recebendo de volta o que enviou.
        # Forca pra "Aguardando Validação" mesmo que o bairro/CEP tenha
        # batido normal, pra nao rotear como entrega comum sem o usuario
        # conferir antes. Pedido do Samuel em 27/08/2026 (achado com
        # casos reais do CARREFOUR onde PAGADOR == CLIENTE).
        pagador_norm = subset["PAGADOR"].apply(normaliza_texto)
        cliente_norm = subset["CLIENTE"].apply(normaliza_texto)
        eh_provavel_devolucao = (pagador_norm == cliente_norm) & (pagador_norm != "")
        subset.loc[eh_provavel_devolucao, "CAPITAL_INTERIOR"] = "Aguardando Validação"

        if ignora_base_rotas and uf_predominante:
            uf_linha = subset.get("_UF_RECEBEDOR", pd.Series("", index=subset.index)).fillna("")
            eh_uf_diferente = (uf_linha != "") & (uf_linha != uf_predominante)
            subset.loc[eh_uf_diferente, "CAPITAL_INTERIOR"] = "Aguardando Validação"

            # Sem SETOR nao da pra montar ROTA nenhuma (ex: CTRC sem
            # setor preenchido, ou arquivo no formato "relatorio
            # impresso" que nem tem essa coluna) - vai pra Aguardando
            # Validação em vez de aparecer em "Interior" sem rota
            # nenhuma pra imprimir. Pedido do Samuel em 07/09/2026.
            rota_vazia = subset["ROTA"].isna() | (subset["ROTA"].astype(str).str.strip() == "")
            subset.loc[rota_vazia, "CAPITAL_INTERIOR"] = "Aguardando Validação"
        return subset

    roteirizavel = _com_capital_interior(df_final[df_final["_JA_LIBERADA_ROTA"] & ~ja_em_rota & ~eh_minuta])
    notas_em_rota = _com_capital_interior(df_final[ja_em_rota & ~eh_minuta])

    if eh_minuta.any():
        minutas = _com_capital_interior(df_final[eh_minuta])
        minutas["CAPITAL_INTERIOR"] = "Aguardando Validação"
        roteirizavel = pd.concat([roteirizavel, minutas])

    retidos = df_final[eh_retido & ~eh_minuta].copy()
    nao_liberado_nem_retido_nem_rota = ~df_final["_JA_LIBERADA_ROTA"] & ~eh_retido & ~ja_em_rota & ~eh_minuta
    a_caminho = df_final[nao_liberado_nem_retido_nem_rota & eh_a_caminho].copy()
    nao_liberadas = df_final[nao_liberado_nem_retido_nem_rota & ~eh_a_caminho].copy()

    colunas_saida = [c for c in df_final.columns if not c.startswith("_")]
    return (
        roteirizavel[colunas_saida + ["CAPITAL_INTERIOR"]],
        retidos[colunas_saida],
        nao_liberadas[colunas_saida],
        notas_em_rota[colunas_saida + ["CAPITAL_INTERIOR"]],
        a_caminho[colunas_saida],
    )


def monta_tabela_retidos(retidos):
    """Monta a aba 'Retidos': CTRC, NOTA FISCAL, REMETENTE/PAGADOR (os
    dois juntos numa coluna so), CIDADE RECEBEDOR, ROTA, PREVISAO
    ENTREGA, RESUMO, DESCRICAO ULTIMA OCORRENCIA (pedido do Samuel em
    21/08/2026, pra dar pra ver o motivo da retencao direto na aba, sem
    abrir o .sswweb original) e DIAS PARADO (dias uteis desde a ultima
    ocorrencia - que pra quem esta retido e a propria data da retencao,
    mesmo calculo que ja existe pra DIAS PARADA EM PISO).
    CIDADE RECEBEDOR/ROTA/PREVISAO ENTREGA/RESUMO acrescentadas em
    28/08/2026 a pedido do Samuel - mesmas colunas ja calculadas em
    monta_colunas_finais pras outras abas, so nao estavam sendo escritas
    aqui ainda. Identica a do robo desktop."""
    colunas = [
        "CTRC", "NOTA FISCAL", "REMETENTE/PAGADOR", "CIDADE RECEBEDOR",
        "ROTA", "PREVISAO ENTREGA", "RESUMO", "STATUS PRAZO",
        "DESCRIÇÃO ÚLTIMA OCORRÊNCIA", "DIAS PARADO",
    ]
    if len(retidos) == 0:
        return pd.DataFrame(columns=colunas)

    remetente = retidos.get("REMETENTE", pd.Series("", index=retidos.index)).fillna("").astype(str).str.strip()
    pagador = retidos.get("PAGADOR", pd.Series("", index=retidos.index)).fillna("").astype(str).str.strip()
    remetente_pagador = (remetente + " / " + pagador).str.strip(" /")

    return pd.DataFrame({
        "CTRC": retidos["CTRC"].values,
        "NOTA FISCAL": retidos["NOTA FISCAL"].values,
        "REMETENTE/PAGADOR": remetente_pagador.values,
        "CIDADE RECEBEDOR": retidos["CIDADE RECEBEDOR"].values,
        "ROTA": retidos["ROTA"].values,
        "PREVISAO ENTREGA": retidos["PREVISAO ENTREGA"].values,
        "RESUMO": retidos["RESUMO"].values,
        "STATUS PRAZO": retidos["STATUS PRAZO"].values,
        "DESCRIÇÃO ÚLTIMA OCORRÊNCIA": retidos["DETALHE ÚLTIMA OCORRÊNCIA"].values,
        "DIAS PARADO": retidos["DIAS PARADA EM PISO"].values,
    })


# ------------------------------------------------------------------
# EXPORTACAO (Capital.xlsx / Interior.xlsx com resumo no topo)
# ------------------------------------------------------------------
def escreve_resumo_topo(ws, df_aba, df_aba_completo=None):
    contagem = df_aba["RESUMO"].value_counts()
    ws.cell(row=1, column=1, value=f"Total: {len(df_aba)}").font = openpyxl.styles.Font(bold=True)
    partes = " | ".join(f"{cat}: {contagem.get(cat, 0)}" for cat in ORDEM_RESUMO)
    ws.cell(row=2, column=1, value=partes)

    # Linha 3 (antes ficava em branco): observacao de possivel fila de
    # descarregamento, agrupada por manifesto - so aparece quando tem
    # algum CTRC flagado nessa aba (ver calcula_alerta_fila_descarregamento).
    # Pedido do Samuel em 20/08/2026.
    if df_aba_completo is not None and "ALERTA_FILA_DESCARREGAMENTO" in df_aba_completo.columns:
        flagados = df_aba_completo[df_aba_completo["ALERTA_FILA_DESCARREGAMENTO"]]
        if len(flagados):
            partes_manifesto = []
            for manifesto, grupo_m in flagados.groupby("MANIFESTO/END"):
                data_ref = grupo_m["DATA DA ÚLTIMA OCORRÊNCIA"].max()
                data_str = data_ref.strftime("%d/%m/%Y") if pd.notna(data_ref) else ""
                partes_manifesto.append(
                    f"{len(grupo_m)} CTRC(s) possível(is) ainda em descarregamento — MANIFESTO {manifesto} ({data_str})"
                )
            cell = ws.cell(row=3, column=1, value="⚠ " + " | ".join(partes_manifesto))
            cell.font = openpyxl.styles.Font(bold=True, color="B36B00")


def monta_linha_resumo_prazo(df_aba):
    """Linha 'ATRASO: X | VENCE HOJE: Y | ...' - mesmo calculo que ja
    existia dentro de escreve_resumo_topo, so que separado numa funcao
    pra poder montar a lista de linhas do resumo ANTES de escrever a
    aba (ver escreve_resumo_topo_mesclado - precisa saber quantas
    linhas o resumo vai ter pra decidir em que linha a tabela comeca)."""
    contagem = df_aba["RESUMO"].value_counts()
    return " | ".join(f"{cat}: {contagem.get(cat, 0)}" for cat in ORDEM_RESUMO)


def monta_linha_fila_descarregamento(df_aba_completo):
    """Linha de alerta de possivel fila de descarregamento, agrupada
    por manifesto - None se nao tiver nenhum CTRC flagado nessa aba.
    Mesmo calculo que ja existia dentro de escreve_resumo_topo."""
    if df_aba_completo is None or "ALERTA_FILA_DESCARREGAMENTO" not in df_aba_completo.columns:
        return None
    flagados = df_aba_completo[df_aba_completo["ALERTA_FILA_DESCARREGAMENTO"]]
    if not len(flagados):
        return None
    partes_manifesto = []
    for manifesto, grupo_m in flagados.groupby("MANIFESTO/END"):
        data_ref = grupo_m["DATA DA ÚLTIMA OCORRÊNCIA"].max()
        data_str = data_ref.strftime("%d/%m/%Y") if pd.notna(data_ref) else ""
        partes_manifesto.append(
            f"{len(grupo_m)} CTRC(s) possível(is) ainda em descarregamento — MANIFESTO {manifesto} ({data_str})"
        )
    return "⚠ " + " | ".join(partes_manifesto)


def monta_linha_resumo_manifestos(df_aba_completo):
    """Linha com quantos CTRC cada manifesto tem na aba 'A Caminho' -
    None se a aba nao tiver a coluna MANIFESTO/END ou vier vazia.
    Pedido do Samuel em 07/09/2026 ('trazer um pequeno resumo tipo,
    quantos ctrc cada manifesto tem e quais são esses')."""
    if df_aba_completo is None or "MANIFESTO/END" not in df_aba_completo.columns:
        return None
    manifesto = df_aba_completo["MANIFESTO/END"].fillna("").astype(str).str.strip()
    com_manifesto = df_aba_completo[manifesto != ""]
    if not len(com_manifesto):
        return None
    partes = []
    contagem = com_manifesto.groupby(com_manifesto["MANIFESTO/END"].astype(str).str.strip()).size()
    for manifesto_nome, qtd in contagem.sort_values(ascending=False).items():
        partes.append(f"{qtd} CTRC(s) — MANIFESTO {manifesto_nome}")
    return f"📦 {len(contagem)} manifesto(s) em trânsito: " + " | ".join(partes)


def escreve_resumo_topo_mesclado(ws, n_colunas, linhas):
    """Escreve as linhas de resumo (Total, RESUMO por prazo, fila de
    descarregamento, manifestos etc.) numa UNICA celula mesclada,
    cobrindo da coluna A ate a ultima coluna da tabela e uma linha de
    planilha pra cada item de 'linhas' (quebra de linha dentro da
    propria celula, com wrap_text). Formato pedido pelo Samuel em
    07/09/2026 (baseado num exemplo que ele montou a mao no Excel -
    antes cada informacao ficava numa linha de celula separada, sem
    mesclar).

    Retorna quantas linhas de planilha a celula mesclada ocupou -
    quem chama usa esse numero pra saber em que linha a tabela de
    dados deve comecar (sempre com 1 linha em branco de respiro depois
    do resumo, igual sempre foi)."""
    linhas = [l for l in linhas if l]
    if not linhas:
        return 0
    n_linhas = len(linhas)
    if n_linhas > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=n_linhas, end_column=max(n_colunas, 1))
    cell = ws.cell(row=1, column=1, value="\n".join(linhas))
    cell.font = openpyxl.styles.Font(bold=True)
    cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top", horizontal="left")
    for i in range(1, n_linhas + 1):
        ws.row_dimensions[i].height = 15.5
    return n_linhas


def formata_como_tabela_offset(caminho, nome_aba, nome_tabela, linha_cabecalho):
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    n_linhas = ws.max_row
    n_colunas = ws.max_column
    if n_linhas < linha_cabecalho + 1 or n_colunas < 1:
        return

    ultima_coluna = openpyxl.utils.get_column_letter(n_colunas)
    ref = f"A{linha_cabecalho}:{ultima_coluna}{n_linhas}"

    tabela = Table(displayName=nome_tabela, ref=ref)
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True,
        showFirstColumn=False, showLastColumn=False, showColumnStripes=False,
    )
    ws.add_table(tabela)

    centralizado = Alignment(horizontal="center", vertical="center")
    for col_cells in ws.columns:
        valores = [c.value for c in col_cells if c.value is not None and c.row >= linha_cabecalho]
        maior = max((len(str(v)) for v in valores), default=8)
        # get_column_letter(coluna) em vez de col_cells[0].column_letter -
        # quando a linha 1 tem celulas mescladas (resumo mesclado, ver
        # escreve_resumo_topo_mesclado), col_cells[0] pode ser um
        # MergedCell, que so tem o atributo .column (numero), sem o
        # ".column_letter" pronto que a celula normal tem. Achado
        # testando o resumo mesclado em 07/09/2026.
        letra = openpyxl.utils.get_column_letter(col_cells[0].column)
        # Teto de largura subido de 45 pra 60 em 25/08/2026 (pedido do
        # Samuel) - colunas com texto longo (ex: BAIRRO RECEBEDOR com
        # observacao de endereco) ficavam cortadas visualmente.
        ws.column_dimensions[letra].width = min(maior + 2, 60)
        for cell in col_cells:
            if cell.row >= linha_cabecalho:
                cell.alignment = centralizado

    # Congela tudo acima (e incluindo) a linha do cabecalho, pra ela
    # continuar visivel rolando a planilha pra baixo - pedido do Samuel
    # em 25/08/2026.
    ws.freeze_panes = f"A{linha_cabecalho + 1}"

    wb.save(caminho)


def aplica_formato_numero_offset(caminho, nome_aba, nomes_colunas, linha_cabecalho):
    """Formata coluna numerica com 2 casas decimais - formato padrao que
    o Excel reconhece como categoria 'Numero' (nao 'Personalizado').
    Usava 4 casas antes, mas esse formato caia em 'Personalizado' no
    Excel por nao bater com nenhum formato embutido - trocado a pedido
    do Samuel em 18/08/2026."""
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    header = [c.value for c in ws[linha_cabecalho]]
    for nome_col in nomes_colunas:
        if nome_col not in header:
            continue
        col_idx = header.index(nome_col) + 1
        for linha in range(linha_cabecalho + 1, ws.max_row + 1):
            cell = ws.cell(row=linha, column=col_idx)
            if cell.value is not None:
                cell.number_format = '#,##0.00'
    wb.save(caminho)


def aplica_formato_texto_offset(caminho, nome_aba, nomes_colunas, linha_cabecalho):
    """Forca formato de TEXTO nessas colunas - evita o Excel 'adivinhar'
    que e numero (CEP, Nota Fiscal) e aplicar formatacao propria por
    conta dele (separador de milhar, ou pior, cortar zero a esquerda).
    Pedido do Samuel em 18/08/2026."""
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    header = [c.value for c in ws[linha_cabecalho]]
    for nome_col in nomes_colunas:
        if nome_col not in header:
            continue
        col_idx = header.index(nome_col) + 1
        for linha in range(linha_cabecalho + 1, ws.max_row + 1):
            cell = ws.cell(row=linha, column=col_idx)
            cell.number_format = "@"
    wb.save(caminho)


def aplica_formato_numero_inteiro_offset(caminho, nome_aba, nomes_colunas, linha_cabecalho):
    """Formata coluna numerica inteira (ex: Nota Fiscal) sem casas
    decimais - so aplica em celulas que ja sao numero de verdade (as
    que ficaram como texto, por nao serem puramente numericas, ficam
    como estavam). Pedido do Samuel em 18/08/2026."""
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    header = [c.value for c in ws[linha_cabecalho]]
    for nome_col in nomes_colunas:
        if nome_col not in header:
            continue
        col_idx = header.index(nome_col) + 1
        for linha in range(linha_cabecalho + 1, ws.max_row + 1):
            cell = ws.cell(row=linha, column=col_idx)
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'
    wb.save(caminho)


def aplica_formato_data_offset(caminho, nome_aba, nomes_colunas, linha_cabecalho):
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    header = [c.value for c in ws[linha_cabecalho]]
    for nome_col in nomes_colunas:
        if nome_col not in header:
            continue
        col_idx = header.index(nome_col) + 1
        for linha in range(linha_cabecalho + 1, ws.max_row + 1):
            cell = ws.cell(row=linha, column=col_idx)
            if cell.value is not None:
                cell.number_format = "DD/MM/YYYY"
    wb.save(caminho)


def aplica_destaque_alerta_offset(caminho, nome_aba, linha_cabecalho):
    """Pinta de vermelho claro a celula da coluna ALERTA (validacao de
    UF fora do RN) quando ela vem preenchida, e de amarelo claro a
    coluna ULTIMA OCORRENCIA quando for especificamente Falta de
    Documentacao - mesmos destaques que ja existiam no romaneio em PDF,
    agora tambem no Excel."""
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    header = [c.value for c in ws[linha_cabecalho]]

    preenchimento_vermelho = PatternFill("solid", fgColor="FBD5D5")
    fonte_vermelha = Font(bold=True, color="B30000")
    preenchimento_amarelo = PatternFill("solid", fgColor="FFF3C4")
    fonte_amarela = Font(bold=True, color="8A6D00")

    if "ALERTA" in header:
        col_idx = header.index("ALERTA") + 1
        for linha in range(linha_cabecalho + 1, ws.max_row + 1):
            cell = ws.cell(row=linha, column=col_idx)
            if cell.value:
                cell.fill = preenchimento_vermelho
                cell.font = fonte_vermelha

    if "DETALHE ÚLTIMA OCORRÊNCIA" in header:
        col_idx = header.index("DETALHE ÚLTIMA OCORRÊNCIA") + 1
        for linha in range(linha_cabecalho + 1, ws.max_row + 1):
            cell = ws.cell(row=linha, column=col_idx)
            if cell.value and eh_falta_documentacao(cell.value):
                cell.fill = preenchimento_amarelo
                cell.font = fonte_amarela

    wb.save(caminho)


def aplica_destaque_dias_em_rota_offset(caminho, nome_aba, linha_cabecalho):
    """So na aba 'Notas em Rota': pinta a celula de DETALHE ÚLTIMA
    OCORRÊNCIA de amarelo/vermelho de acordo com DIAS EM ROTA, com
    limites diferentes pra Capital e Interior (Interior tem prazo de
    entrega maior, entao a tolerancia e maior).

    Capital: amarelo quando DIAS EM ROTA == 1, vermelho quando >= 2.
    Pedido do Samuel em 25/08/2026: 'saiu em rota 21/08 e o arquivo foi
    extraido 22/08 e ainda esta em rota fica amarelo, porem se saiu em
    rota 20/08 ou antes fica vermelho'.

    Interior: amarelo quando DIAS EM ROTA == 2, vermelho quando >= 3.
    Pedido do Samuel em 26/08/2026: limite do Interior um dia mais
    tolerante que o Capital em cada faixa (D+2 amarelo / D+3 vermelho)."""
    wb = openpyxl.load_workbook(caminho)
    if nome_aba not in wb.sheetnames:
        return
    ws = wb[nome_aba]
    header = [c.value for c in ws[linha_cabecalho]]
    if "CAPITAL_INTERIOR" not in header or "DIAS EM ROTA" not in header or "DETALHE ÚLTIMA OCORRÊNCIA" not in header:
        return

    col_capital = header.index("CAPITAL_INTERIOR") + 1
    col_dias = header.index("DIAS EM ROTA") + 1
    col_detalhe = header.index("DETALHE ÚLTIMA OCORRÊNCIA") + 1

    preenchimento_vermelho = PatternFill("solid", fgColor="FBD5D5")
    fonte_vermelha = Font(bold=True, color="B30000")
    preenchimento_amarelo = PatternFill("solid", fgColor="FFF3C4")
    fonte_amarela = Font(bold=True, color="8A6D00")

    LIMITES = {
        "Capital": {"amarelo": 1, "vermelho": 2},
        "Interior": {"amarelo": 2, "vermelho": 3},
    }

    for linha in range(linha_cabecalho + 1, ws.max_row + 1):
        tipo_rota = ws.cell(row=linha, column=col_capital).value
        dias = ws.cell(row=linha, column=col_dias).value
        limites = LIMITES.get(tipo_rota)
        if limites is None or dias is None:
            continue
        cell_detalhe = ws.cell(row=linha, column=col_detalhe)
        if dias == limites["amarelo"]:
            cell_detalhe.fill = preenchimento_amarelo
            cell_detalhe.font = fonte_amarela
        elif dias >= limites["vermelho"]:
            cell_detalhe.fill = preenchimento_vermelho
            cell_detalhe.font = fonte_vermelha

    wb.save(caminho)


def gerar_capital_interior_sswweb(df_capital_interior, pasta_saida):
    df = df_capital_interior
    colunas_excluir = {
        "CAPITAL_INTERIOR", "REGIAO_ROTA", "CODIGO ULTIMA OCORRENCIA",
        "MANIFESTO/END", "ALERTA_FILA_DESCARREGAMENTO",
    }
    colunas_export = [c for c in df.columns if c not in colunas_excluir]

    capital_completo = df[df["CAPITAL_INTERIOR"] == "Capital"]
    interior_completo = df[df["CAPITAL_INTERIOR"] == "Interior"]
    aguardando_completo = df[df["CAPITAL_INTERIOR"] == "Aguardando Validação"]

    capital = capital_completo[colunas_export]
    interior = interior_completo[colunas_export]
    aguardando = aguardando_completo[colunas_export]

    startrow = LINHA_INICIO_TABELA - 1

    arq_capital = os.path.join(pasta_saida, "Capital - SSW.xlsx")
    abas_capital = ["Capital"]
    with pd.ExcelWriter(arq_capital, engine="openpyxl") as writer:
        capital.to_excel(writer, index=False, sheet_name="Capital", startrow=startrow)
        escreve_resumo_topo(writer.sheets["Capital"], capital, capital_completo)
        if len(aguardando) > 0:
            aguardando.to_excel(writer, index=False, sheet_name="Aguardando Validação", startrow=startrow)
            escreve_resumo_topo(writer.sheets["Aguardando Validação"], aguardando, aguardando_completo)
            abas_capital.append("Aguardando Validação")
    for nome_aba in abas_capital:
        aplica_formato_data_offset(arq_capital, nome_aba, ["PREVISAO ENTREGA", "DATA DA ÚLTIMA OCORRÊNCIA"], LINHA_INICIO_TABELA)
        aplica_formato_numero_offset(arq_capital, nome_aba, ["KG REAL", "M3", "PESO CALCULO"], LINHA_INICIO_TABELA)
        aplica_formato_texto_offset(arq_capital, nome_aba, ["NOTA FISCAL", "CEP"], LINHA_INICIO_TABELA)
        formata_como_tabela_offset(arq_capital, nome_aba, nome_tabela_valido("Capital", nome_aba), LINHA_INICIO_TABELA)
        aplica_destaque_alerta_offset(arq_capital, nome_aba, LINHA_INICIO_TABELA)

    arq_interior = os.path.join(pasta_saida, "Interior - SSW.xlsx")
    abas_interior = []
    with pd.ExcelWriter(arq_interior, engine="openpyxl") as writer:
        if len(interior) == 0:
            interior.to_excel(writer, index=False, sheet_name="Interior", startrow=startrow)
            escreve_resumo_topo(writer.sheets["Interior"], interior)
            abas_interior.append("Interior")
        else:
            for rota, grupo_completo in interior_completo.groupby("ROTA", dropna=False):
                nome_rota = str(rota) if pd.notna(rota) and str(rota).strip() else "SEM ROTA"
                nome_aba = nome_rota[:31]
                grupo = grupo_completo[colunas_export]
                grupo.to_excel(writer, index=False, sheet_name=nome_aba, startrow=startrow)
                escreve_resumo_topo(writer.sheets[nome_aba], grupo, grupo_completo)
                abas_interior.append(nome_aba)
    for nome_aba in abas_interior:
        aplica_formato_data_offset(arq_interior, nome_aba, ["PREVISAO ENTREGA", "DATA DA ÚLTIMA OCORRÊNCIA"], LINHA_INICIO_TABELA)
        aplica_formato_numero_offset(arq_interior, nome_aba, ["KG REAL", "M3", "PESO CALCULO"], LINHA_INICIO_TABELA)
        aplica_formato_texto_offset(arq_interior, nome_aba, ["NOTA FISCAL", "CEP"], LINHA_INICIO_TABELA)
        formata_como_tabela_offset(arq_interior, nome_aba, nome_tabela_valido("Interior", nome_aba), LINHA_INICIO_TABELA)
        aplica_destaque_alerta_offset(arq_interior, nome_aba, LINHA_INICIO_TABELA)

    return arq_capital, arq_interior, len(capital), len(interior), len(aguardando)


def gerar_combinado_sswweb(df_capital_interior, retidos, nao_liberadas, notas_em_rota, a_caminho, pasta_saida,
                            uf_predominante=None, forcar_modo_setor=False):
    """Arquivo UNICO de exportacao do robo online (pedido do Samuel em
    20/08/2026, virou o UNICO arquivo Excel gerado em 21/08/2026 - antes
    disso existiam tambem o Capital - SSW.xlsx e o Interior - SSW.xlsx
    separados, gerados por gerar_capital_interior_sswweb; essa funcao
    ainda existe no codigo mas deixou de ser chamada pelo app.py). Ganhou
    as abas Retidos e Nao Liberadas em 21/08/2026, mesma logica ja
    aplicada no robo desktop (ver separa_liberados_retidos_nao_liberados
    e monta_tabela_retidos). Ganhou a aba Notas em Rota em 25/08/2026 e a
    aba A Caminho em 26/08/2026.

    Todas as 6 abas vem ordenadas primeiro por DETALHE ÚLTIMA OCORRÊNCIA
    (ou DESCRIÇÃO ÚLTIMA OCORRÊNCIA na aba Retidos, que usa esse nome de
    coluna) - pedido do Samuel em 26/08/2026, pra CTRCs com a mesma
    ocorrencia ficarem juntos ao abrir o arquivo. As colunas que antes
    eram o criterio principal de ordenacao (CAPITAL_INTERIOR/ROTA/
    CIDADE/BAIRRO) viram desempate dentro de cada grupo de ocorrencia.

    Aba 1 'Capital e Interior': Capital + todas as rotas do Interior
    numa tabela so (mantem a coluna CAPITAL_INTERIOR e ROTA pra dar pra
    filtrar/agrupar dentro do proprio Excel). Diferente de
    gerar_capital_interior_sswweb, aqui a coluna CAPITAL_INTERIOR fica
    visivel de proposito (e o que diferencia as linhas, ja que nao tem
    mais aba separada pra isso). A coluna ALERTA (destino fora do RN)
    nao entra mais no Excel a partir de 25/08/2026 - pedido do Samuel,
    continua calculada por baixo dos panos so pro aviso que aparece na
    tela antes de gerar o arquivo.

    Aba 2 'Aguardando Validação' (so aparece se tiver algum CTRC nessa
    situacao): quem ainda nao tem rota definida (cidade fora do RN ou
    nao cadastrada no base_rotas.xlsx) - antes só existia dentro do
    Capital - SSW.xlsx, agora entra como aba extra desse arquivo unico
    pra nao perder essa informacao.

    Aba 3 'Retidos' (so aparece se tiver algum CTRC retido): codigo 56,
    mercadoria retida pela fiscalizacao - so CTRC/NOTA FISCAL/
    REMETENTE-PAGADOR/DIAS PARADO, igual ao robo desktop.

    Aba 4 'Nao Liberadas' (so aparece se tiver algum CTRC nessa
    situacao): todo o resto que nao esta liberado pra rota e nao esta
    retido (ex: pedido cancelado, destinatario desconhecido, aguardando
    devolucao) - antes esse grupo nao aparecia em nenhum lugar do
    arquivo, so contava no resumo da tela. Codigo 85 (ja saiu com o
    motorista) continua de fora dessa aba, igual sempre foi pro
    roteirizavel.

    Aba 5 'Notas em Rota' (so aparece se tiver algum CTRC em codigo 85):
    mesma dinamica/colunas da aba 'Capital e Interior', so que DIAS
    PARADA EM PISO vira DIAS EM ROTA (mesmo numero, so que aqui
    significa ha quantos dias uteis o CTRC saiu com o motorista, ja que
    pra codigo 85 a DATA DA ULTIMA OCORRENCIA e a data que saiu em
    rota). Chegou a ter o manifesto embutido no DETALHE ÚLTIMA
    OCORRÊNCIA (pedido do Samuel em 26/08/2026), mas foi desfeito em
    27/08/2026 - essa aba ficou so com o texto original da ocorrencia,
    sem manifesto (a aba 'A Caminho' e que ficou com esse tratamento).
    Nas rotas Capital, a celula de DETALHE ÚLTIMA OCORRÊNCIA fica
    amarela quando DIAS EM ROTA == 1 (saiu ontem, ainda nao voltou) e
    vermelha quando DIAS EM ROTA >= 2 (saiu antes de ontem); Interior
    segue a mesma logica com limite D+2 amarelo / D+3 vermelho - ver
    aplica_destaque_dias_em_rota_offset.

    Aba 6 'A Caminho' (so aparece se tiver algum CTRC nessa situacao):
    dentro do que nao esta liberado/retido/em rota, os CTRCs em
    transferencia entre unidades (DETALHE ÚLTIMA OCORRÊNCIA batendo com
    DESCRICOES_A_CAMINHO - 'EM TRANSITO - SAIDA UNIDADE TRANSBORDO' ou
    'VEICULO EM POSTO FISCAL PARA SELAGEM DAS NOTAS', codigo 57
    acrescentado em 27/08/2026) - antes ficavam misturados dentro de
    'Nao Liberadas', pedido do Samuel em 26/08/2026 pra separar numa
    aba propria. Mesmo tratamento pro manifesto: sem coluna nova, ele
    entra dentro da propria celula de DETALHE ÚLTIMA OCORRÊNCIA (ex: 'EM
    TRANSITO - SAIDA UNIDADE TRANSBORDO — Manifesto SPO046018-4'). Sem
    CAPITAL_INTERIOR e sem o destaque amarelo/vermelho de dias em rota
    (isso e especifico de quem ja SAIU com o motorista, codigo 85 -
    aqui o CTRC ainda esta em transferencia, nao em rota de entrega)."""
    df = df_capital_interior
    # LIBERADO_FISCALIZACAO_RECENTE acrescentada em 07/09/2026 - so
    # serve pro aviso na tela previa do app.py, DATA DA ÚLTIMA
    # OCORRÊNCIA ja cobre a informacao de data e os romaneios nao usam
    # essa coluna. CEP/KG REAL/M3/PESO CALCULO tambem tiradas em
    # 07/09/2026 - continuam calculadas normal (o robo online ainda le
    # e monta essas colunas, caso precise voltar a exibir e so tirar
    # daqui), so nao entram mais no Excel final. Pedido do Samuel em
    # 07/09/2026.
    colunas_excluir = {
        "REGIAO_ROTA", "CODIGO ULTIMA OCORRENCIA", "MANIFESTO/END", "ALERTA_FILA_DESCARREGAMENTO",
        "REMETENTE", "ALERTA", "LIBERADO_FISCALIZACAO_RECENTE",
        "CEP", "KG REAL", "M3", "PESO CALCULO",
    }
    colunas_export = [c for c in df.columns if c not in colunas_excluir]
    colunas_export_aguardando = [c for c in colunas_export if c != "CAPITAL_INTERIOR"]
    colunas_export_nao_liberadas = [c for c in nao_liberadas.columns if c not in colunas_excluir]
    colunas_export_notas_em_rota = [c for c in notas_em_rota.columns if c not in colunas_excluir]
    colunas_export_a_caminho = [c for c in a_caminho.columns if c not in colunas_excluir]

    def _com_manifesto_no_detalhe(serie_detalhe, serie_manifesto):
        """DETALHE ÚLTIMA OCORRÊNCIA com ' — Manifesto XXX' no final
        quando tiver manifesto - reusado em Notas em Rota e A Caminho,
        os dois unicos lugares onde o manifesto importa mostrar (pedido
        do Samuel em 26/08/2026)."""
        detalhe = serie_detalhe.fillna("").astype(str).str.strip()
        manifesto = serie_manifesto.fillna("").astype(str).str.strip()
        tem_manifesto = manifesto != ""
        return detalhe.where(~tem_manifesto, detalhe + " — Manifesto " + manifesto)

    # Todas as abas vem ordenadas primeiro por DETALHE/DESCRIÇÃO ÚLTIMA
    # OCORRÊNCIA, pra CTRCs com a mesma ocorrencia ficarem juntos ao
    # abrir o arquivo - pedido do Samuel em 26/08/2026. As colunas
    # antigas de ordenacao (CAPITAL_INTERIOR/ROTA/CIDADE/BAIRRO) viram
    # criterio de desempate dentro de cada grupo de ocorrencia.
    #
    # Quando a UF predominante do arquivo nao e RN, CAPITAL_INTERIOR
    # vem em branco pra quem nao tem rota cadastrada (ver classifica_
    # rotas) - entra no combinado do mesmo jeito, so nao usa mais
    # isin(["Capital", "Interior"]) (que excluiria essas linhas em
    # branco): fica "tudo que nao e Aguardando Validação". Pra arquivo
    # RN normal o resultado e identico ao de antes (Capital/Interior
    # continuam sendo os unicos valores possiveis fora de Aguardando).
    # Pedido do Samuel em 07/09/2026.
    ignora_base_rotas = forcar_modo_setor or (bool(uf_predominante) and uf_predominante != UF_ESPERADA)
    if ignora_base_rotas:
        combinado_completo = df[df["CAPITAL_INTERIOR"] != "Aguardando Validação"].copy()
    else:
        combinado_completo = df[df["CAPITAL_INTERIOR"].isin(["Capital", "Interior"])].copy()
    combinado_completo = combinado_completo.sort_values(
        ["DETALHE ÚLTIMA OCORRÊNCIA", "CAPITAL_INTERIOR", "ROTA", "CIDADE RECEBEDOR", "BAIRRO RECEBEDOR"],
        na_position="last",
    )
    combinado = combinado_completo[colunas_export]

    aguardando_completo = df[df["CAPITAL_INTERIOR"] == "Aguardando Validação"].copy()
    aguardando_completo = aguardando_completo.sort_values(
        ["DETALHE ÚLTIMA OCORRÊNCIA", "ROTA", "CIDADE RECEBEDOR", "BAIRRO RECEBEDOR"],
        na_position="last",
    )
    aguardando = aguardando_completo[colunas_export_aguardando]

    tabela_retidos = monta_tabela_retidos(retidos)
    if len(tabela_retidos) > 0:
        tabela_retidos = tabela_retidos.sort_values(
            ["DESCRIÇÃO ÚLTIMA OCORRÊNCIA", "CTRC"], na_position="last"
        )

    nao_liberadas_completo = nao_liberadas.sort_values(
        ["DETALHE ÚLTIMA OCORRÊNCIA", "ROTA", "CIDADE RECEBEDOR", "BAIRRO RECEBEDOR"],
        na_position="last",
    )
    nao_liberadas_export = nao_liberadas_completo[colunas_export_nao_liberadas]

    # Notas em Rota: SEM manifesto embutido no DETALHE - pedido do
    # Samuel em 27/08/2026 (desfeito o que tinha sido feito em
    # 26/08/2026 nessa aba especifica; a aba 'A Caminho' continua com o
    # manifesto embutido, so essa aqui voltou atras).
    notas_em_rota_completo = notas_em_rota.copy()
    notas_em_rota_completo = notas_em_rota_completo.sort_values(
        ["DETALHE ÚLTIMA OCORRÊNCIA", "CAPITAL_INTERIOR", "ROTA", "CIDADE RECEBEDOR", "BAIRRO RECEBEDOR"],
        na_position="last",
    )

    notas_em_rota_export = notas_em_rota_completo[colunas_export_notas_em_rota].rename(
        columns={"DIAS PARADA EM PISO": "DIAS EM ROTA"}
    )

    # A Caminho: mesmo tratamento (manifesto embutido no DETALHE), mas
    # sem CAPITAL_INTERIOR (nao roda _com_capital_interior nessa
    # entrada) e sem renomear DIAS PARADA EM PISO - pedido do Samuel em
    # 26/08/2026.
    a_caminho_completo = a_caminho.copy()
    a_caminho_completo["DETALHE ÚLTIMA OCORRÊNCIA"] = _com_manifesto_no_detalhe(
        a_caminho_completo["DETALHE ÚLTIMA OCORRÊNCIA"], a_caminho_completo["MANIFESTO/END"]
    )
    a_caminho_completo = a_caminho_completo.sort_values(
        ["DETALHE ÚLTIMA OCORRÊNCIA", "ROTA", "CIDADE RECEBEDOR", "BAIRRO RECEBEDOR"],
        na_position="last",
    )
    a_caminho_export = a_caminho_completo[colunas_export_a_caminho]

    # Resumo do topo de cada aba (Total, RESUMO por prazo, fila de
    # descarregamento, manifestos) numa UNICA celula mesclada, em vez
    # de LINHA_INICIO_TABELA fixa pra todas as abas - como o numero de
    # linhas do resumo varia de aba pra aba (ex: A Caminho ganhou a
    # linha extra de manifestos), a linha onde a tabela de dados
    # comeca tambem passa a variar por aba (guardado em `cabecalhos`).
    # Pedido do Samuel em 07/09/2026 (queria o resumo mesclado igual a
    # um exemplo que ele montou a mao, com RESUMO tambem em Retidos e A
    # Caminho e um resumo extra de manifestos em A Caminho).
    def _monta_linhas_resumo(df_export, df_completo=None, incluir_manifestos=False, apenas_total=False):
        linhas = [f"Total: {len(df_export)}"]
        if apenas_total:
            return linhas
        if "RESUMO" in df_export.columns:
            linhas.append(monta_linha_resumo_prazo(df_export))
        alerta = monta_linha_fila_descarregamento(df_completo)
        if alerta:
            linhas.append(alerta)
        if incluir_manifestos:
            linha_manifestos = monta_linha_resumo_manifestos(df_completo)
            if linha_manifestos:
                linhas.append(linha_manifestos)
        return linhas

    cabecalhos = {}  # nome da aba -> linha (1-indexada) onde comeca o cabecalho da tabela

    def _escreve_aba(writer, nome, df_export, linhas_resumo):
        startrow = len(linhas_resumo) + 1  # pandas e 0-indexado; sobra 1 linha em branco entre o resumo e o cabecalho
        df_export.to_excel(writer, index=False, sheet_name=nome, startrow=startrow)
        n_linhas_usadas = escreve_resumo_topo_mesclado(writer.sheets[nome], len(df_export.columns), linhas_resumo)
        cabecalhos[nome] = n_linhas_usadas + 2

    # Nomes e ordem das abas - pedido do Samuel em 07/09/2026: nome mais
    # direto (ex: "Retidos" -> "Retenção fiscal") e ordem seguindo o
    # fluxo real do CTRC (liberado -> ja em rota -> retido/nao liberado
    # -> precisa validar -> em transito entre unidades), em vez da
    # ordem antiga (que vinha da ordem em que cada aba foi acrescentada
    # ao longo do desenvolvimento). "Aguardando Validação" mantem o
    # mesmo nome de sempre (e tambem o valor usado na coluna
    # CAPITAL_INTERIOR em varios lugares do codigo - NAO mexer nesse
    # texto, so na posicao da aba).
    NOME_ABA_LIBERADAS = "Liberadas para rota"
    NOME_ABA_EM_ROTA = "Em rota"
    # Aba "Em rota" REATIVADA em 11/09/2026 a pedido do Samuel. Historico:
    # foi desligada em 08/09/2026 porque o numero (codigo 85) nao reflete
    # a realidade, ja que a tela do SSW usada pra exportar (081 - Relacao
    # CTRCs Disponiveis para Entrega) mostra o que ja esta em posse do
    # parceiro ou ja manifestado pra ele - nao um retrato confiavel de
    # quem "esta em rota" (confirmado cruzando com o BI - Regional
    # Status, ver conversa de 08/09/2026). Essa limitacao continua
    # valendo - a aba volta, mas acompanhada de um alerta no app.py
    # (tela_download) pedindo pra validar esses numeros com o BI antes
    # de repassar, pois pode haver diferenca. NAO remover esse alerta
    # sem reativar tambem a confianca no dado (ou seja, sem que o
    # levantamento com o BI comprove que o codigo 85 passou a refletir a
    # realidade).
    EXPORTAR_ABA_EM_ROTA = True
    NOME_ABA_RETENCAO_FISCAL = "Retenção fiscal"
    NOME_ABA_NAO_LIBERADAS = "Não liberadas"
    NOME_ABA_AGUARDANDO = "Aguardando Validação"
    NOME_ABA_A_CAMINHO = "A caminho"

    nome_aba = NOME_ABA_LIBERADAS
    # Nome do arquivo segue sempre "Base 081 - <UF>.xlsx", RN incluso -
    # antes o RN saia com nome diferente ("Capital e Interior -
    # SSW.xlsx"), sobra de quando so existia RN. uf_predominante vem
    # vazio (None) so quando a coluna UF RECEBEDOR nao existe/vem toda
    # vazia (ex: formato "relatorio impresso") - cai no UF_ESPERADA
    # (RN) nesse caso, que e a suposicao padrao do robo. Pedido do
    # Samuel em 07/09/2026.
    nome_arquivo = f"Base 081 - {uf_predominante or UF_ESPERADA}.xlsx"
    arq_combinado = os.path.join(pasta_saida, nome_arquivo)
    with pd.ExcelWriter(arq_combinado, engine="openpyxl") as writer:
        _escreve_aba(writer, nome_aba, combinado, _monta_linhas_resumo(combinado, combinado_completo))
        if len(notas_em_rota_export) > 0 and EXPORTAR_ABA_EM_ROTA:
            _escreve_aba(writer, NOME_ABA_EM_ROTA, notas_em_rota_export, _monta_linhas_resumo(notas_em_rota_export, apenas_total=True))
        if len(tabela_retidos) > 0:
            _escreve_aba(writer, NOME_ABA_RETENCAO_FISCAL, tabela_retidos, _monta_linhas_resumo(tabela_retidos))
        if len(nao_liberadas_export) > 0:
            _escreve_aba(writer, NOME_ABA_NAO_LIBERADAS, nao_liberadas_export, _monta_linhas_resumo(nao_liberadas_export, nao_liberadas_completo))
        if len(aguardando) > 0:
            _escreve_aba(writer, NOME_ABA_AGUARDANDO, aguardando, _monta_linhas_resumo(aguardando, aguardando_completo))
        if len(a_caminho_export) > 0:
            _escreve_aba(writer, NOME_ABA_A_CAMINHO, a_caminho_export, _monta_linhas_resumo(a_caminho_export, a_caminho_completo, incluir_manifestos=True))

    abas = [nome_aba]
    if len(notas_em_rota_export) > 0 and EXPORTAR_ABA_EM_ROTA:
        abas.append(NOME_ABA_EM_ROTA)
    if len(nao_liberadas_export) > 0:
        abas.append(NOME_ABA_NAO_LIBERADAS)
    if len(aguardando) > 0:
        abas.append(NOME_ABA_AGUARDANDO)
    if len(a_caminho_export) > 0:
        abas.append(NOME_ABA_A_CAMINHO)
    for aba in abas:
        linha_cabecalho = cabecalhos[aba]
        aplica_formato_data_offset(arq_combinado, aba, ["PREVISAO ENTREGA", "DATA DA ÚLTIMA OCORRÊNCIA"], linha_cabecalho)
        aplica_formato_numero_offset(arq_combinado, aba, ["KG REAL", "M3", "PESO CALCULO"], linha_cabecalho)
        aplica_formato_numero_inteiro_offset(arq_combinado, aba, ["STATUS PRAZO"], linha_cabecalho)
        aplica_formato_texto_offset(arq_combinado, aba, ["NOTA FISCAL", "CEP"], linha_cabecalho)
        formata_como_tabela_offset(arq_combinado, aba, nome_tabela_valido("Combinado", aba), linha_cabecalho)
        aplica_destaque_alerta_offset(arq_combinado, aba, linha_cabecalho)

    if len(tabela_retidos) > 0:
        # PREVISAO ENTREGA entrou nessa aba em 28/08/2026 - "Retenção
        # fiscal" fica fora do loop de "abas" acima (usa formatacao
        # propria, sem CAPITAL_INTERIOR/ALERTA), entao precisa da
        # formatacao de data aqui tambem, senao a coluna nova aparece
        # como numero de serie do Excel em vez de data.
        linha_cabecalho_retidos = cabecalhos[NOME_ABA_RETENCAO_FISCAL]
        aplica_formato_data_offset(arq_combinado, NOME_ABA_RETENCAO_FISCAL, ["PREVISAO ENTREGA"], linha_cabecalho_retidos)
        aplica_formato_numero_inteiro_offset(arq_combinado, NOME_ABA_RETENCAO_FISCAL, ["DIAS PARADO", "STATUS PRAZO"], linha_cabecalho_retidos)
        formata_como_tabela_offset(arq_combinado, NOME_ABA_RETENCAO_FISCAL, "Tab_Retencao_Fiscal", linha_cabecalho_retidos)

    if len(notas_em_rota_export) > 0 and EXPORTAR_ABA_EM_ROTA:
        linha_cabecalho_notas = cabecalhos[NOME_ABA_EM_ROTA]
        aplica_formato_numero_inteiro_offset(arq_combinado, NOME_ABA_EM_ROTA, ["DIAS EM ROTA"], linha_cabecalho_notas)
        aplica_destaque_dias_em_rota_offset(arq_combinado, NOME_ABA_EM_ROTA, linha_cabecalho_notas)

    return (
        arq_combinado, len(combinado), len(aguardando), len(tabela_retidos),
        len(nao_liberadas_export), len(notas_em_rota_export), len(a_caminho_export),
    )


# ------------------------------------------------------------------
# ROMANEIO DE CARGA (PDF, um por rota)
# ------------------------------------------------------------------
COLUNAS_ROMANEIO = ["NOTA FISCAL", "CLIENTE", "CIDADE RECEBEDOR", "BAIRRO RECEBEDOR", "PREVISAO ENTREGA", "ROTA"]


def calcula_status_vencimento(previsao_entrega, hoje=None):
    """Calcula o status de vencimento em dias corridos (nao dia util -
    isso e diferente do AGE STATUS do ROBO RN4, que conta so dia util).
    Retorna o texto pronto pra exibir (VENCIDA / VENCE HOJE / VENCE +N)
    e tambem os dias corridos (usado so pra ordenar, nao aparece no
    PDF) - pedido do Samuel em 10/08/2026."""
    if pd.isna(previsao_entrega):
        return "", None
    if hoje is None:
        hoje = datetime.date.today()
    dias = (previsao_entrega.date() - hoje).days
    if dias < 0:
        texto = f"VENCIDA +{abs(dias)}"
    elif dias == 0:
        texto = "VENCE HOJE"
    else:
        texto = f"VENCE +{dias}"
    return texto, dias


def formata_nota_fiscal(nota_fiscal, pagador):
    """Se o pagador for Amazon, junta o pagador ao lado da nota fiscal
    (ex: '123 - AMAZON'); senao mostra so a nota fiscal normal - pedido
    do Samuel em 09/08/2026, pra identificar de cara no romaneio quais
    CTRCs sao pagos pela Amazon."""
    nf = str(nota_fiscal or "").strip().upper()
    pagador_str = str(pagador or "").strip().upper()
    if "AMAZON" in pagador_str:
        return f"{nf} - AMAZON"
    return nf


def nome_arquivo_valido(texto):
    """Deixa um nome de rota seguro pra usar como nome de arquivo no
    Windows (sem / \\ : * ? " < > |)."""
    limpo = re.sub(r'[\\/:*?"<>|]', "-", str(texto)).strip()
    return limpo if limpo else "SEM_ROTA"


class CanvasComRodapePaginado(pdfcanvas.Canvas):
    """Canvas customizado pra escrever 'pagina/total' no rodape inferior
    direito de cada pagina do PDF. O reportlab so sabe o total de
    paginas depois de montar o documento inteiro - por isso ele guarda
    cada pagina "pausada" em memoria e so desenha o rodape (com o total
    ja certo) na hora de salvar o arquivo."""

    def __init__(self, *args, **kwargs):
        pdfcanvas.Canvas.__init__(self, *args, **kwargs)
        self._paginas_pausadas = []

    def showPage(self):
        self._paginas_pausadas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_paginas = len(self._paginas_pausadas)
        for estado in self._paginas_pausadas:
            self.__dict__.update(estado)
            self._desenha_rodape_pagina(total_paginas)
            pdfcanvas.Canvas.showPage(self)
        pdfcanvas.Canvas.save(self)

    def _desenha_rodape_pagina(self, total_paginas):
        largura_pagina, _ = landscape(A4)
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.grey)
        texto = f"{self._pageNumber}/{total_paginas}"
        self.drawRightString(largura_pagina - 1.0 * cm, 0.5 * cm, texto)


CABECALHO_ROMANEIO = ["NOTA FISCAL", "CLIENTE", "CIDADE", "BAIRRO", "PREVISÃO", "STATUS", "ÚLTIMA OCORRÊNCIA"]
LARGURAS_ROMANEIO = [3.6 * cm, 5.0 * cm, 2.6 * cm, 3.0 * cm, 2.1 * cm, 2.3 * cm, 5.8 * cm]
# Antes o cabecalho tinha fundo azul-marinho solido (#2E2E5C) em toda
# a linha - com muita cidade pequena (romaneio por cidade) ou muita
# rota, esse preenchimento escuro se repete dezenas de vezes no
# documento e gasta bastante tinta na impressao. Trocado por fundo
# branco + texto em negrito na mesma cor + uma linha grossa embaixo
# do cabecalho, que ainda deixa o cabecalho bem distinto da tabela sem
# pintar a celula inteira. O zebrado leve (branco/cinza bem claro) nas
# linhas de dado continua igual - gasta pouca tinta e ajuda a
# acompanhar a linha. Pedido do Samuel em 07/09/2026.
ESTILO_TABELA_ROMANEIO = TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.white),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#2E2E5C")),
    ("FONTSIZE", (0, 0), (-1, 0), 9),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("LINEBELOW", (0, 0), (-1, 0), 1.2, colors.HexColor("#2E2E5C")),
    ("ALIGN", (4, 0), (6, -1), "CENTER"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F7")]),
    ("TOPPADDING", (0, 0), (-1, -1), 2),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
])


def eh_falta_documentacao(detalhe_ocorrencia):
    """Confere se a descricao da ultima ocorrencia e 'FALTA DE
    DOCUMENTACAO' (com/sem acento, maiuscula/minuscula) - usado pra
    marcar o alerta na coluna JUSTIFICATIVA do romaneio - pedido do
    Samuel em 11/08/2026."""
    return "FALTA DE DOCUMENTACAO" in normaliza_texto(detalhe_ocorrencia or "")


def remove_duplicatas_falta_documentacao(df):
    """So pro ROMANEIO EM PDF (nao mexe na base/Excel): quando o mesmo
    CLIENTE tem duas remessas pra MESMA DATA PREVISTA - uma com CTRC
    emitido pela propria unidade RN4 (comeca com 'RN', fica travada no
    codigo 52/Falta de Documentacao) e outra com CTRC de outro deposito
    (nao comeca com 'RN') - remove a que NAO E CTRC-RN, mas SO quando
    essa remessa (a de outro deposito) JA SAIU do codigo 52 (ou seja,
    esta em qualquer outro codigo - 84/Chegada na Unidade, 85/Ja em
    rota, etc.). Enquanto as duas estiverem em 52 (falta de
    documentacao ainda nao resolvida em nenhuma delas), mantem as duas
    no romaneio. Se o CTRC-RN nao tiver par de outro deposito, ele fica
    normalmente. Pedido do Samuel em 12/08/2026 (usa o prefixo do CTRC
    em vez do prefixo da NF, porque a NF e sequencial/progressiva e o
    "15" ia virar "16" mais pra frente - o prefixo RN do CTRC e
    estrutural, nao muda com o tempo)."""
    df = df.copy()
    ctrc_texto = df["CTRC"].fillna("").astype(str).str.strip().str.upper()
    eh_ctrc_proprio_rn4 = ctrc_texto.str.startswith(PREFIXO_CTRC_EMISSAO_PROPRIA)
    codigo_ocorrencia = df.get("CODIGO ULTIMA OCORRENCIA", pd.Series("", index=df.index))
    codigo_ocorrencia = codigo_ocorrencia.fillna("").astype(str).str.strip()
    eh_pendencia_solucionada = (codigo_ocorrencia != "") & (codigo_ocorrencia != "52")

    chave_cliente = df["CLIENTE"].apply(normaliza_texto)
    chave_data = df["PREVISAO ENTREGA"].apply(lambda d: d.date() if pd.notna(d) else None)
    chave_dup = list(zip(chave_cliente, chave_data))

    chaves_com_par_ctrc_rn4 = set(
        chave for chave, eh_rn in zip(chave_dup, eh_ctrc_proprio_rn4) if eh_rn
    )
    remover = [
        (not eh_rn) and eh_resolvida and (chave in chaves_com_par_ctrc_rn4)
        for chave, eh_rn, eh_resolvida in zip(chave_dup, eh_ctrc_proprio_rn4, eh_pendencia_solucionada)
    ]
    n_removidos = sum(remover)
    if n_removidos:
        log_detalhe(
            f"[dedup romaneio] {n_removidos} CTRC(s) de outro deposito (nao-RN) removido(s) "
            f"do romaneio em PDF por duplicidade (mesmo cliente/data ja tem CTRC-RN4 "
            f"e ocorrencia ja saiu do 52/falta de documentacao)"
        )
    return df[[not r for r in remover]]


def monta_tabela_romaneio(grupo, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=False):
    """Monta a tabela (reportlab) de uma rota - mesma logica de linhas
    que ja usavamos, so que separada em funcao pra poder reaproveitar
    tanto no PDF unico da Capital quanto nos PDFs por rota do Interior.

    mostrar_ultima_ocorrencia=True (so nos romaneios normais, Capital e
    Interior - NAO no PDF de rotas escolhidas): em vez da coluna ficar
    em branco esperando preenchimento manual, mostra o detalhe da
    ultima ocorrencia + a data dela (ex: 'CHEGADA NA UNIDADE 25/07',
    'CLIENTE AUSENTE 2 15/08'). O alerta de Falta de Documentacao
    continua tendo prioridade e aparecendo em vermelho igual antes.
    Pedido do Samuel em 14/08/2026."""
    linhas = [CABECALHO_ROMANEIO]
    for _, r in grupo.iterrows():
        previsao = r["PREVISAO ENTREGA"]
        previsao_str = previsao.strftime("%d/%m") if pd.notna(previsao) else ""
        nota_fiscal_fmt = formata_nota_fiscal(r["NOTA FISCAL"], r.get("PAGADOR"))
        status_texto = r.get("_STATUS_VENC_TXT", "") or ""
        if eh_falta_documentacao(r.get("DETALHE ÚLTIMA OCORRÊNCIA")):
            justificativa = Paragraph("[ALERTA] FALTA DE DOCUMENTAÇÃO", estilo_alerta)
        elif mostrar_ultima_ocorrencia:
            detalhe = str(r.get("DETALHE ÚLTIMA OCORRÊNCIA") or "").strip()
            data_ocorr = r.get("DATA DA ÚLTIMA OCORRÊNCIA")
            if detalhe and pd.notna(data_ocorr):
                texto_ocorrencia = f"{detalhe} {data_ocorr.strftime('%d/%m')}"
            else:
                texto_ocorrencia = detalhe
            justificativa = Paragraph(texto_ocorrencia, estilo_celula) if texto_ocorrencia else ""
        else:
            justificativa = ""  # em branco, pra preencher na mao durante o carregamento
        linhas.append([
            Paragraph(nota_fiscal_fmt, estilo_celula),
            Paragraph(str(r["CLIENTE"] or "").upper(), estilo_celula),
            Paragraph(str(r["CIDADE RECEBEDOR"] or "").upper(), estilo_celula),
            Paragraph(str(r["BAIRRO RECEBEDOR"] or "").upper(), estilo_celula),
            previsao_str,
            Paragraph(status_texto, estilo_celula),
            justificativa,
        ])
    tabela = PdfTable(linhas, colWidths=LARGURAS_ROMANEIO, repeatRows=1)
    tabela.setStyle(ESTILO_TABELA_ROMANEIO)
    return tabela


def calcula_data_limite_semana(hoje=None):
    """Calcula a proxima segunda-feira a partir de hoje (dia corrido).
    Se hoje ja for segunda, pega a segunda que vem (nao hoje mesmo) -
    pedido do Samuel em 10/08/2026: seg->seg que vem, ter->seg, qua->seg,
    ... sex->segunda."""
    if hoje is None:
        hoje = datetime.date.today()
    dias_ate_segunda = (7 - hoje.weekday()) % 7
    if dias_ate_segunda == 0:
        dias_ate_segunda = 7
    return hoje + datetime.timedelta(days=dias_ate_segunda)


def resume_status_vencimento(grupos):
    """Conta quantas remessas caem em cada faixa de vencimento, somando
    todas as rotas de 'grupos' (lista de (rota, dataframe)). Usado pra
    montar a linha de resumo no topo do romaneio da Capital - pedido do
    Samuel em 10/08/2026."""
    total_vencidas = 0
    total_hoje_amanha = 0
    total_dois_mais = 0
    for _, grupo in grupos:
        dias = grupo["_STATUS_VENC_DIAS"]
        total_vencidas += int((dias < 0).sum())
        total_hoje_amanha += int(dias.isin([0, 1]).sum())
        total_dois_mais += int((dias >= 2).sum())
    return (
        f"{total_vencidas} VENCIDA(S)  |  "
        f"{total_hoje_amanha} VENCENDO HOJE E AMANHÃ  |  "
        f"{total_dois_mais} VENCENDO A PARTIR DE 2 DIAS"
    )


def gerar_romaneios_pdf(df_capital_interior, pasta_saida, data_limite=None, rotas_permitidas=None):
    """Gera os romaneios de carga em PDF, so pra quem tem rota definida
    (fica de fora quem estiver em 'Aguardando Validação', que ainda nao
    tem rota pra carregar - e tambem quem for CARREFOUR com ocorrencia
    "Falta de Documentacao", que nao tem documento resolvido pra sair).
    Esses dois filtros sao so do PDF - a base/Excel continua saindo com
    tudo. Pedido do Samuel em 27/08/2026 (falta de documentacao do
    CARREFOUR).

    rotas_permitidas: quando informado (lista/set de nomes de rota), so
    gera romaneio pras rotas que o usuario escolheu no menu "GERAR
    ROMANEIO DE QUAIS ROTAS" - os Excel (Capital-SSW.xlsx/Interior-SSW.xlsx)
    continuam saindo com TUDO sempre, independente disso; so os PDFs
    ficam restritos as rotas escolhidas. Quando None, gera romaneio de
    todas as rotas (comportamento antigo). Pedido do Samuel em
    14/08/2026.

    Capital: SAI TUDO EM 1 UNICO ARQUIVO ("ROMANEIO CAPITAL - ROTAS
    X,Y,Z.pdf"), com cada rota em sequencia (um pequeno espaco entre
    uma rota e outra, sem pagina em branco forcada no meio - pra
    imprimir tudo direto, na ordem certa), rodape com "pagina/total" no
    canto inferior direito, e o total que aparece no topo agora e o
    total geral da Capital (nao mais por rota) - confirmado com o
    Samuel em 09/08/2026.

    Interior: continua 1 PDF por rota.

    data_limite: quando informado (datetime.date), filtra CAPITAL E
    INTERIOR pra trazer so as remessas VENCIDAS + as que vencem ate essa
    data (dia corrido, inclusive). Quem nao tem previsao (NaN) fica de
    fora do filtro. Substituiu o antigo menu "1-Todas / 2-so ate proxima
    segunda" por uma data livre digitada pelo usuario - pedido do Samuel
    em 11/08/2026. Quando data_limite=None, sai tudo, sem filtro (igual
    opcao 1 de antes).
    """
    pasta_romaneios = os.path.join(pasta_saida, "romaneios")
    os.makedirs(pasta_romaneios, exist_ok=True)

    # tira os arquivos de romaneio da rodada anterior, pra nao ficar
    # PDF de rota que nao existe mais nessa rodada misturado com os novos
    for f in os.listdir(pasta_romaneios):
        if f.lower().endswith(".pdf"):
            os.remove(os.path.join(pasta_romaneios, f))

    roteirizavel = df_capital_interior[df_capital_interior["ROTA"].notna()].copy()
    roteirizavel = roteirizavel[roteirizavel["ROTA"].astype(str).str.strip() != ""]
    # Filtro explicito pra "Aguardando Validação" (antes ficava de fora so
    # por acaso, porque quem caia la geralmente nao tinha ROTA definida -
    # deixou de ser garantido a partir do caso PAGADOR == CLIENTE, que
    # forca "Aguardando Validação" mesmo com ROTA/bairro OK. Sem esse
    # filtro esses CTRCs de provavel devolucao poderiam voltar a entrar
    # no romaneio se o usuario escolhesse a rota deles. Pedido do Samuel
    # em 27/08/2026.
    if "CAPITAL_INTERIOR" in roteirizavel.columns:
        roteirizavel = roteirizavel[roteirizavel["CAPITAL_INTERIOR"] != "Aguardando Validação"]

    # CARREFOUR com ocorrencia "Falta de Documentacao" nao sai no romaneio -
    # a nota ainda nao tem documento resolvido pra ser carregada, entao nao
    # agrega nada aparecer junto com o resto ja pronto pra entrega. So tira
    # do PDF (a base/Excel continua saindo com tudo, sem filtro). Pedido do
    # Samuel em 27/08/2026.
    #
    # IMPORTANTE: quem e "do CARREFOUR" aqui e o REMETENTE/PAGADOR (quem
    # despachou/paga o frete) - a coluna CLIENTE e o DESTINATARIO, ou seja,
    # a pessoa que vai RECEBER a mercadoria (ex: "ELIZABETH S SOUZA..."),
    # que nunca vai ser "CARREFOUR". Usar CLIENTE aqui deixava o filtro sem
    # efeito nenhum - achado testando com um .sswweb real do Samuel em
    # 27/08/2026 (17 CTRCs do CARREFOUR com Falta de Documentacao no
    # arquivo, 0 pegos pelo filtro errado).
    eh_carrefour = (
        roteirizavel["REMETENTE"].apply(normaliza_texto).str.contains("CARREFOUR", na=False)
        | roteirizavel["PAGADOR"].apply(normaliza_texto).str.contains("CARREFOUR", na=False)
    )
    eh_falta_doc = roteirizavel["DETALHE ÚLTIMA OCORRÊNCIA"].apply(normaliza_texto).str.contains(
        "FALTA DE DOCUMENTACAO", na=False
    )
    roteirizavel = roteirizavel[~(eh_carrefour & eh_falta_doc)]

    roteirizavel = remove_duplicatas_falta_documentacao(roteirizavel)
    if rotas_permitidas is not None:
        roteirizavel = roteirizavel[roteirizavel["ROTA"].astype(str).str.strip().isin(rotas_permitidas)]

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloRomaneio", parent=styles["Heading1"], fontSize=13, spaceAfter=2,
    )
    estilo_subtitulo = ParagraphStyle(
        "SubtituloRomaneio", parent=styles["Normal"], fontSize=8, textColor=colors.grey, spaceAfter=6,
    )
    estilo_rota = ParagraphStyle(
        "RotaRomaneio", parent=styles["Heading2"], fontSize=10, spaceBefore=10, spaceAfter=2,
        textColor=colors.HexColor("#2E2E5C"),
    )
    estilo_celula = ParagraphStyle(
        "CelulaRomaneio", parent=styles["Normal"], fontSize=7, leading=8.5,
    )
    estilo_alerta = ParagraphStyle(
        "AlertaRomaneio", parent=styles["Normal"], fontSize=7, leading=8.5,
        textColor=colors.HexColor("#C0392B"), fontName="Helvetica-Bold",
    )

    # separa os grupos em Capital e Interior - toda rota cai sempre na
    # mesma classificacao Capital/Interior (e uma funcao da rota), pega
    # o valor mais comum do grupo so por seguranca, caso algum dia mude
    hoje = datetime.date.today()
    status_calculado = roteirizavel["PREVISAO ENTREGA"].apply(lambda p: calcula_status_vencimento(p, hoje))
    roteirizavel["_STATUS_VENC_TXT"] = status_calculado.apply(lambda t: t[0])
    roteirizavel["_STATUS_VENC_DIAS"] = status_calculado.apply(lambda t: t[1])

    grupos_capital = []
    grupos_interior = []
    for rota, grupo in roteirizavel.groupby("ROTA"):
        # ordem pedida: CIDADE, BAIRRO, e dentro disso por data de
        # vencimento (mais vencida primeiro) - quem nao tem previsao
        # (NaN) vai pro fim
        grupo = grupo.sort_values(
            ["CIDADE RECEBEDOR", "BAIRRO RECEBEDOR", "_STATUS_VENC_DIAS"],
            ascending=True, na_position="last",
        )
        eh_capital = (grupo["CAPITAL_INTERIOR"] == "Capital").mean() >= 0.5
        if eh_capital:
            grupos_capital.append((rota, grupo))
        else:
            grupos_interior.append((rota, grupo))

    # filtro por data (quando informado) vale pra Capital E Interior -
    # antes so valia pra Capital, mudou a pedido do Samuel em 11/08/2026
    dias_limite = None
    if data_limite is not None:
        dias_limite = (data_limite - hoje).days

        def _filtra_por_data(grupos):
            filtrados = []
            for rota, grupo in grupos:
                grupo_filtrado = grupo[
                    grupo["_STATUS_VENC_DIAS"].notna()
                    & (grupo["_STATUS_VENC_DIAS"] <= dias_limite)
                ]
                if len(grupo_filtrado) > 0:
                    filtrados.append((rota, grupo_filtrado))
            return filtrados

        grupos_capital = _filtra_por_data(grupos_capital)
        grupos_interior = _filtra_por_data(grupos_interior)

    arquivos_gerados = []
    agora = agora_brasil().strftime("%d/%m/%Y %H:%M")

    # ---------------------------------------------------------------
    # CAPITAL - um unico PDF com todas as rotas em sequencia
    # ---------------------------------------------------------------
    if grupos_capital:
        total_capital = sum(len(grupo) for _, grupo in grupos_capital)
        nomes_rotas = ", ".join(str(rota) for rota, _ in grupos_capital)
        sufixo_data = f" - ATE {data_limite.strftime('%d-%m')}" if data_limite is not None else ""
        # ":" nao pode em nome de arquivo no Windows - vira " -" no nome
        # do arquivo, mas o titulo escrito dentro do PDF continua com ":"
        nome_arquivo = nome_arquivo_valido(f"ROMANEIO CAPITAL - ROTAS {nomes_rotas}{sufixo_data}") + ".pdf"
        caminho_pdf = os.path.join(pasta_romaneios, nome_arquivo)

        doc = SimpleDocTemplate(
            caminho_pdf, pagesize=landscape(A4),
            leftMargin=1.0 * cm, rightMargin=1.0 * cm, topMargin=0.9 * cm, bottomMargin=1.2 * cm,
        )
        elementos = []
        elementos.append(Paragraph(f"ROMANEIO CAPITAL - ROTAS: {nomes_rotas}", estilo_titulo))
        elementos.append(Paragraph(
            f"Gerado em {agora}  |  Total Capital: {total_capital} REMESSA(S)", estilo_subtitulo
        ))
        if data_limite is not None:
            elementos.append(Paragraph(
                f"Filtro aplicado: VENCIDAS + vencendo até {data_limite.strftime('%d/%m')}",
                estilo_subtitulo,
            ))
        elementos.append(Paragraph(resume_status_vencimento(grupos_capital), estilo_subtitulo))

        for i, (rota, grupo) in enumerate(grupos_capital):
            if i > 0:
                elementos.append(Spacer(1, 8))  # pequeno espaco entre rotas, sem forcar pagina nova
            # KeepTogether: titulo da rota e a tabela pulam de pagina JUNTOS
            # quando nao cabem no espaco que sobrou - evita titulo orfao no
            # topo de uma pagina quase vazia, ou tabela pequena cortada no
            # meio sem necessidade (pedido do Samuel em 11/08/2026). Se a
            # rota tiver mais remessas do que cabe numa pagina inteira, ela
            # ainda quebra normalmente entre paginas.
            elementos.append(KeepTogether([
                Paragraph(f"Rota {rota}  ({len(grupo)} remessa(s))", estilo_rota),
                monta_tabela_romaneio(grupo, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=True),
            ]))

        doc.build(elementos, canvasmaker=CanvasComRodapePaginado)
        arquivos_gerados.append(caminho_pdf)

    # ---------------------------------------------------------------
    # INTERIOR - o proprio robo decide quais rotas cabem juntas num
    # unico PDF: renderiza um PDF de teste (em memoria, nao grava em
    # disco) e CONTA as paginas de verdade, em vez de usar um numero
    # fixo de remessas "no chute". Pedido do Samuel em 11/08/2026.
    #
    # Regra: uma rota so entra no "pool" de possiveis combinadas se,
    # SOZINHA, ja cabe em 1 pagina. Rotas que sozinhas passam de 1
    # pagina continuam com PDF proprio, como sempre foi. Dentro do
    # pool, agrupa rotas (da menor pra maior) num PDF combinado ate o
    # limite de PAGINAS_ALVO_AGRUPADO paginas; se sobrar mais gente
    # pequena depois de bater o alvo, comeca um novo PDF combinado
    # (Agrupadas 2, 3...) em vez de deixar de fora.
    # ---------------------------------------------------------------
    PAGINAS_ALVO_AGRUPADO = 3  # quantas paginas cada PDF combinado pode ter, no maximo

    def _monta_elementos_combinado(grupos_bloco, sufixo_titulo=""):
        els = [
            Paragraph(f"ROMANEIO INTERIOR - ROTAS AGRUPADAS{sufixo_titulo}", estilo_titulo),
            Paragraph(
                f"Gerado em {agora}  |  Total: {sum(len(g) for _, g in grupos_bloco)} REMESSA(S)",
                estilo_subtitulo,
            ),
        ]
        if data_limite is not None:
            els.append(Paragraph(
                f"Filtro aplicado: VENCIDAS + vencendo até {data_limite.strftime('%d/%m')}",
                estilo_subtitulo,
            ))
        els.append(Paragraph(resume_status_vencimento(grupos_bloco), estilo_subtitulo))
        for i, (rota, grupo) in enumerate(grupos_bloco):
            if i > 0:
                els.append(Spacer(1, 8))
            els.append(KeepTogether([
                Paragraph(f"Rota {rota}  ({len(grupo)} remessa(s))", estilo_rota),
                monta_tabela_romaneio(grupo, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=True),
            ]))
        return els

    def _conta_paginas(elementos_teste):
        """Renderiza em memoria (nao grava arquivo) e conta as paginas
        de verdade - assim a decisao de agrupar usa o tamanho REAL da
        tabela (nomes de cliente/bairro longos etc.), nao uma estimativa."""
        buffer = io.BytesIO()
        doc_teste = SimpleDocTemplate(
            buffer, pagesize=landscape(A4),
            leftMargin=1.0 * cm, rightMargin=1.0 * cm, topMargin=0.9 * cm, bottomMargin=1.2 * cm,
        )
        doc_teste.build(elementos_teste)
        buffer.seek(0)
        return len(PdfReader(buffer).pages)

    # 1) separa quem cabe sozinho numa unica pagina (candidatas a
    #    agrupar) de quem ja precisa de PDF proprio
    candidatas_agrupar = []
    grupos_interior_grandes = []
    for rota, grupo in grupos_interior:
        paginas_sozinha = _conta_paginas(_monta_elementos_combinado([(rota, grupo)]))
        if paginas_sozinha <= 1:
            candidatas_agrupar.append((rota, grupo))
        else:
            grupos_interior_grandes.append((rota, grupo))

    # 2) PDF proprio pras rotas grandes - igual sempre foi
    for rota, grupo in grupos_interior_grandes:
        sufixo_data = f" - ATE {data_limite.strftime('%d-%m')}" if data_limite is not None else ""
        nome_arquivo = f"Interior - Romaneio - {nome_arquivo_valido(rota)}{sufixo_data}.pdf"
        caminho_pdf = os.path.join(pasta_romaneios, nome_arquivo)

        doc = SimpleDocTemplate(
            caminho_pdf, pagesize=landscape(A4),
            leftMargin=1.0 * cm, rightMargin=1.0 * cm, topMargin=0.9 * cm, bottomMargin=0.9 * cm,
        )
        elementos = []
        elementos.append(Paragraph(f"Romaneio de Carga - Rota {rota}", estilo_titulo))
        elementos.append(Paragraph(f"Gerado em {agora}  |  Total: {len(grupo)} REMESSA(S)", estilo_subtitulo))
        if data_limite is not None:
            elementos.append(Paragraph(
                f"Filtro aplicado: VENCIDAS + vencendo até {data_limite.strftime('%d/%m')}",
                estilo_subtitulo,
            ))
        elementos.append(Paragraph(resume_status_vencimento([(rota, grupo)]), estilo_subtitulo))
        elementos.append(monta_tabela_romaneio(grupo, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=True))
        doc.build(elementos)
        arquivos_gerados.append(caminho_pdf)

    # 3) empacota as candidatas (da menor pra maior) em lotes que cabem
    #    ate PAGINAS_ALVO_AGRUPADO paginas cada, testando de verdade a
    #    cada tentativa de adicionar mais uma rota ao lote
    candidatas_agrupar.sort(key=lambda item: len(item[1]))
    lotes = []
    lote_atual = []
    for rota, grupo in candidatas_agrupar:
        tentativa = lote_atual + [(rota, grupo)]
        if _conta_paginas(_monta_elementos_combinado(tentativa)) <= PAGINAS_ALVO_AGRUPADO:
            lote_atual = tentativa
        else:
            if lote_atual:
                lotes.append(lote_atual)
            lote_atual = [(rota, grupo)]
    if lote_atual:
        lotes.append(lote_atual)

    for idx, lote in enumerate(lotes, start=1):
        sufixo_lote = f" {idx}" if len(lotes) > 1 else ""
        sufixo_data = f" - ATE {data_limite.strftime('%d-%m')}" if data_limite is not None else ""
        nome_arquivo = f"Interior - Rotas Agrupadas{sufixo_lote}{sufixo_data}.pdf"
        caminho_pdf = os.path.join(pasta_romaneios, nome_arquivo)

        doc = SimpleDocTemplate(
            caminho_pdf, pagesize=landscape(A4),
            leftMargin=1.0 * cm, rightMargin=1.0 * cm, topMargin=0.9 * cm, bottomMargin=1.2 * cm,
        )
        elementos = _monta_elementos_combinado(lote, sufixo_titulo=(f" - PARTE {idx}" if len(lotes) > 1 else ""))
        doc.build(elementos, canvasmaker=CanvasComRodapePaginado)
        arquivos_gerados.append(caminho_pdf)

    return pasta_romaneios, arquivos_gerados


def gerar_romaneio_por_cidade_pdf(df_capital_interior, pasta_saida, uf_predominante=None, data_limite=None):
    """Romaneio em PDF pra arquivo de OUTRA FILIAL (UF predominante
    diferente de RN - ver detecta_uf_predominante/classifica_rotas),
    onde o base_rotas.xlsx foi ignorado e ROTA/REGIAO_ROTA ficam em
    branco (sem cadastro de rota pra essas cidades - so tem do RN).
    Sem rota/regional pra agrupar, gerar_romaneios_pdf nao serve (ele
    filtra por ROTA preenchida, que aqui sempre vem vazia).

    Em vez disso: 1 PDF UNICO com todas as remessas liberadas, em
    ordem alfabetica de CIDADE RECEBEDOR e, dentro de cada cidade, por
    previsao de entrega crescente (mais vencida primeiro - mesma
    logica de ordenacao que ja usavamos dentro de cada rota).

    Cidade PEQUENA (cabe sozinha em 1 pagina): so a tabela, sem
    titulo (a coluna CIDADE ja repete em cada linha - pedido do Samuel
    em 07/09/2026), dentro de um KeepTogether pra tentar ficar inteira
    na mesma folha que a cidade anterior, sem forcar quebra de pagina
    a toa.

    Cidade GRANDE (nao cabe sozinha em 1 pagina, tipo OLINDA com
    dezenas de remessas): comeca numa folha nova (assim a paginacao
    fica limpa) e e dividida em pedacos que cabem 1 por folha, cada um
    com o titulo "FOLHA X DE Y - CIDADE" antes da tabela - pra quem
    esta com o romaneio impresso em maos saber, so de olhar o topo da
    folha, que ainda esta na mesma cidade e em qual pedaco dela.
    Achar esse ponto de corte e feito testando de verdade (mesmo
    metodo ja usado no agrupamento do Interior em gerar_romaneios_pdf:
    renderiza em memoria e conta pagina), so que por busca binaria
    dentro da propria cidade em vez de cidade inteira. Pedido do
    Samuel em 07/09/2026 (viu bastante espaco em branco desperdicado
    antes de uma cidade grande, e quis a legenda de folha pra cidade
    grande).

    Mesmos filtros de sempre: exclui Aguardando Validação, CARREFOUR
    com Falta de Documentacao, duplicatas de Falta de Documentacao, e
    o mesmo filtro por data_limite (VENCIDAS + vencendo ate a data
    escolhida, quando informado)."""
    pasta_romaneios = os.path.join(pasta_saida, "romaneios")
    os.makedirs(pasta_romaneios, exist_ok=True)
    for f in os.listdir(pasta_romaneios):
        if f.lower().endswith(".pdf"):
            os.remove(os.path.join(pasta_romaneios, f))

    validas = df_capital_interior.copy()
    if "CAPITAL_INTERIOR" in validas.columns:
        validas = validas[validas["CAPITAL_INTERIOR"] != "Aguardando Validação"]

    eh_carrefour = (
        validas["REMETENTE"].apply(normaliza_texto).str.contains("CARREFOUR", na=False)
        | validas["PAGADOR"].apply(normaliza_texto).str.contains("CARREFOUR", na=False)
    )
    eh_falta_doc = validas["DETALHE ÚLTIMA OCORRÊNCIA"].apply(normaliza_texto).str.contains(
        "FALTA DE DOCUMENTACAO", na=False
    )
    validas = validas[~(eh_carrefour & eh_falta_doc)]
    validas = remove_duplicatas_falta_documentacao(validas)

    if len(validas) == 0:
        return pasta_romaneios, []

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloRomaneioCidade", parent=styles["Heading1"], fontSize=13, spaceAfter=2,
    )
    estilo_subtitulo = ParagraphStyle(
        "SubtituloRomaneioCidade", parent=styles["Normal"], fontSize=8, textColor=colors.grey, spaceAfter=6,
    )
    estilo_celula = ParagraphStyle(
        "CelulaRomaneioCidade", parent=styles["Normal"], fontSize=7, leading=8.5,
    )
    estilo_alerta = ParagraphStyle(
        "AlertaRomaneioCidade", parent=styles["Normal"], fontSize=7, leading=8.5,
        textColor=colors.HexColor("#C0392B"), fontName="Helvetica-Bold",
    )
    # so aparece em cidade GRANDE (mais de 1 folha) - ver _divide_em_folhas
    estilo_folha = ParagraphStyle(
        "FolhaRomaneioCidade", parent=styles["Heading2"], fontSize=10, spaceBefore=0, spaceAfter=2,
        textColor=colors.HexColor("#2E2E5C"),
    )

    hoje = datetime.date.today()
    status_calculado = validas["PREVISAO ENTREGA"].apply(lambda p: calcula_status_vencimento(p, hoje))
    validas["_STATUS_VENC_TXT"] = status_calculado.apply(lambda t: t[0])
    validas["_STATUS_VENC_DIAS"] = status_calculado.apply(lambda t: t[1])

    if data_limite is not None:
        dias_limite = (data_limite - hoje).days
        validas = validas[
            validas["_STATUS_VENC_DIAS"].notna() & (validas["_STATUS_VENC_DIAS"] <= dias_limite)
        ]
        if len(validas) == 0:
            return pasta_romaneios, []

    # agrupa por cidade normalizada (maiuscula/sem acento) pra "Natal"
    # e "NATAL" nao virarem grupos diferentes, mas exibe o nome como
    # veio no arquivo (so limpo de espaco extra)
    validas["_CIDADE_ROMANEIO"] = validas["CIDADE RECEBEDOR"].fillna("").astype(str).str.strip()
    validas.loc[validas["_CIDADE_ROMANEIO"] == "", "_CIDADE_ROMANEIO"] = "(SEM CIDADE)"
    validas["_CIDADE_ORDENACAO"] = validas["_CIDADE_ROMANEIO"].apply(normaliza_texto)

    grupos = []
    for _, grupo in validas.groupby("_CIDADE_ORDENACAO"):
        cidade_exibir = grupo["_CIDADE_ROMANEIO"].iloc[0]
        grupo = grupo.sort_values(["_STATUS_VENC_DIAS"], ascending=True, na_position="last")
        grupos.append((cidade_exibir, grupo))
    grupos.sort(key=lambda item: normaliza_texto(item[0]))

    total_geral = sum(len(g) for _, g in grupos)
    sufixo_uf = f" - {uf_predominante}" if uf_predominante else ""
    sufixo_data = f" - ATE {data_limite.strftime('%d-%m')}" if data_limite is not None else ""
    nome_arquivo = nome_arquivo_valido(f"ROMANEIO POR CIDADE{sufixo_uf}{sufixo_data}") + ".pdf"
    caminho_pdf = os.path.join(pasta_romaneios, nome_arquivo)

    doc = SimpleDocTemplate(
        caminho_pdf, pagesize=landscape(A4),
        leftMargin=1.0 * cm, rightMargin=1.0 * cm, topMargin=0.9 * cm, bottomMargin=1.2 * cm,
    )
    elementos = []
    titulo_uf = f" ({uf_predominante})" if uf_predominante else ""
    elementos.append(Paragraph(f"ROMANEIO POR CIDADE{titulo_uf}", estilo_titulo))
    agora = agora_brasil().strftime("%d/%m/%Y %H:%M")
    elementos.append(Paragraph(
        f"Gerado em {agora}  |  Total: {total_geral} REMESSA(S)  |  {len(grupos)} cidade(s)", estilo_subtitulo
    ))
    if data_limite is not None:
        elementos.append(Paragraph(
            f"Filtro aplicado: VENCIDAS + vencendo até {data_limite.strftime('%d/%m')}",
            estilo_subtitulo,
        ))
    elementos.append(Paragraph(resume_status_vencimento(grupos), estilo_subtitulo))

    def _conta_paginas(elementos_teste):
        """Renderiza em memoria (nao grava arquivo) e conta as paginas
        de verdade - mesmo metodo usado no agrupamento do Interior em
        gerar_romaneios_pdf, so que aqui tambem serve pra testar um
        PEDACO de uma cidade grande, nao so a cidade inteira."""
        buffer = io.BytesIO()
        doc_teste = SimpleDocTemplate(
            buffer, pagesize=landscape(A4),
            leftMargin=1.0 * cm, rightMargin=1.0 * cm, topMargin=0.9 * cm, bottomMargin=1.2 * cm,
        )
        doc_teste.build(elementos_teste)
        buffer.seek(0)
        return len(PdfReader(buffer).pages)

    def _cabe_em_uma_folha(subgrupo):
        elementos_teste = [
            Paragraph("FOLHA 99 DE 99 - (TESTE)", estilo_folha),
            monta_tabela_romaneio(subgrupo, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=True),
        ]
        return _conta_paginas(elementos_teste) <= 1

    def _divide_em_folhas(grupo):
        """Corta o dataframe da cidade em pedacos que cabem 1 por
        folha, usando busca binaria (poucos testes de render mesmo pra
        cidade com muita linha) pra achar o maior pedaco que ainda
        cabe numa unica pagina a partir de cada ponto de corte."""
        linhas = grupo.reset_index(drop=True)
        pedacos = []
        inicio = 0
        total = len(linhas)
        while inicio < total:
            lo, hi = inicio + 1, total
            melhor = inicio + 1  # pelo menos 1 linha por pedaco, mesmo que estoure sozinha
            while lo <= hi:
                meio = (lo + hi) // 2
                if _cabe_em_uma_folha(linhas.iloc[inicio:meio]):
                    melhor = meio
                    lo = meio + 1
                else:
                    hi = meio - 1
            pedacos.append(linhas.iloc[inicio:melhor])
            inicio = melhor
        return pedacos

    primeira_cidade = True
    for cidade, grupo in grupos:
        tabela_cidade_inteira = monta_tabela_romaneio(grupo, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=True)
        cabe_sozinha = _conta_paginas([tabela_cidade_inteira]) <= 1

        if cabe_sozinha:
            # cidade pequena: so a tabela (sem titulo - a coluna CIDADE
            # ja repete em cada linha), dentro de um KeepTogether pra
            # tentar ficar inteira na mesma folha que a cidade anterior
            # em vez de forcar quebra de pagina a toa. Pedido do Samuel
            # em 07/09/2026.
            if not primeira_cidade:
                elementos.append(Spacer(1, 10))
            elementos.append(KeepTogether([tabela_cidade_inteira]))
        else:
            # cidade grande: comeca numa folha nova (pra paginacao ficar
            # limpa) e sai em pedacos "FOLHA X DE Y - CIDADE", um por
            # folha. So forca a quebra de pagina aqui - antes disso,
            # cidades pequenas continuam se empacotando livremente.
            if not primeira_cidade:
                elementos.append(PageBreak())
            pedacos = _divide_em_folhas(grupo)
            total_folhas = len(pedacos)
            for i, pedaco in enumerate(pedacos, start=1):
                if i > 1:
                    elementos.append(PageBreak())
                elementos.append(Paragraph(f"FOLHA {i} DE {total_folhas} - {cidade}", estilo_folha))
                elementos.append(monta_tabela_romaneio(pedaco, estilo_celula, estilo_alerta, mostrar_ultima_ocorrencia=True))

        primeira_cidade = False

    doc.build(elementos, canvasmaker=CanvasComRodapePaginado)
    return pasta_romaneios, [caminho_pdf]


def localiza_arquivo_sswweb(pasta):
    candidatos = [
        f for f in os.listdir(pasta)
        if not f.lower().endswith((".xlsx", ".xls"))
        and not f.startswith("~$")
        and not f.startswith(".")
        and not f.lower().startswith("log_processamento")
    ]
    if not candidatos:
        raise FileNotFoundError(f"Nao encontrei nenhum arquivo .sswweb em {pasta}")
    candidatos_completos = [os.path.join(pasta, f) for f in candidatos]
    mais_recente = max(candidatos_completos, key=os.path.getmtime)
    for caminho in candidatos_completos:
        if caminho != mais_recente:
            os.remove(caminho)
            log_detalhe(f"[limpeza] SSWWEB antigo removido: {os.path.basename(caminho)}")
    return mais_recente


def busca_sswweb_da_downloads(pasta_downloads, pasta_dados):
    if not os.path.isdir(pasta_downloads):
        print(f"  [aviso] pasta Downloads nao encontrada em {pasta_downloads} - pulando busca automatica.")
        return
    try:
        arquivos = os.listdir(pasta_downloads)
    except OSError as e:
        print(f"  [aviso] nao consegui ler a pasta Downloads ({e}) - pulando busca automatica.")
        return
    candidatos = [f for f in arquivos if f.lower().endswith(".sswweb") and not f.startswith("~$")]
    if not candidatos:
        print("  [info] nenhum .sswweb encontrado em Downloads - vou tentar usar o que ja estiver na pasta 'dados'.")
        return
    caminhos_completos = [os.path.join(pasta_downloads, f) for f in candidatos]
    mais_recente = max(caminhos_completos, key=os.path.getmtime)
    destino = os.path.join(pasta_dados, os.path.basename(mais_recente))
    shutil.copy2(mais_recente, destino)
    print(f"  [ok] SSWWEB: trouxe '{os.path.basename(mais_recente)}' de Downloads.")
    log_detalhe(f"[downloads] copiado para dados: {os.path.basename(mais_recente)}")


def copia_saida_para_desktop(pasta_saida, arq_capital, arq_interior):
    pasta_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(pasta_desktop):
        print(f"  [aviso] Area de Trabalho nao encontrada em {pasta_desktop} - pulando copia.")
        return
    for origem in [arq_capital, arq_interior]:
        if not os.path.isfile(origem):
            continue
        destino = os.path.join(pasta_desktop, os.path.basename(origem))
        try:
            shutil.copy2(origem, destino)
        except OSError as e:
            print(f"  [aviso] nao consegui copiar {os.path.basename(origem)} pra Area de Trabalho ({e})")
    print("  Copia tambem deixada na Area de Trabalho (Capital e Interior).")


def main():
    os.makedirs(PASTA_SAIDA, exist_ok=True)

    print(f"[{agora_brasil()}] Iniciando processamento ROBO SSW 081...")
    print("Buscando arquivo mais recente na pasta Downloads...")
    busca_sswweb_da_downloads(PASTA_DOWNLOADS, PASTA_DADOS)

    arq_sswweb = localiza_arquivo_sswweb(PASTA_DADOS)
    log_detalhe(f"SSWWEB encontrado: {os.path.basename(arq_sswweb)}")

    if not os.path.isfile(ARQ_BASE_ROTAS):
        raise FileNotFoundError(f"Nao encontrei base_rotas.xlsx em {PASTA_DADOS}")

    df, base_rotas, grupos_especiais, rota_para_classificacao, resumo = prepara_dados(arq_sswweb, ARQ_BASE_ROTAS)
    log_detalhe(f"SSWWEB: {resumo['total_lido']} linha(s) lida(s), {resumo['total_apos_filtro_interno']} apos filtro interno")

    df = classifica_rotas(df, base_rotas, grupos_especiais)
    df_final = monta_colunas_finais(df)
    df_capital_interior = separa_capital_interior(df_final, rota_para_classificacao)

    arq_capital, arq_interior, n_capital, n_interior, n_aguardando = gerar_capital_interior_sswweb(
        df_capital_interior, PASTA_SAIDA
    )

    print()
    resposta_data = input(
        "Quer gerar os romaneios (Capital e Interior) com vencimento ate qual dia? "
        "Digite DD/MM (ou so ENTER pra gerar com todas as remessas): "
    ).strip()

    data_limite = None
    if resposta_data:
        try:
            dia_str, mes_str = resposta_data.split("/")
            ano_atual = agora_brasil().year
            data_limite = datetime.date(ano_atual, int(mes_str), int(dia_str))
            print(f"  [ok] Filtro aplicado: vencidas + vencendo ate {data_limite.strftime('%d/%m/%Y')} (Capital e Interior).")
            log_detalhe(f"[filtro data romaneio] usuario digitou '{resposta_data}' -> ate {data_limite.strftime('%d/%m/%Y')}")
        except (ValueError, IndexError):
            print("  [aviso] Data invalida (formato esperado DD/MM). Gerando romaneios com todas as remessas.")
            log_detalhe(f"[filtro data romaneio] entrada invalida: '{resposta_data}' - seguiu sem filtro")

    pasta_romaneios = os.path.join(PASTA_SAIDA, "romaneios")

    print()
    roteirizavel_para_menu = df_capital_interior[df_capital_interior["ROTA"].notna()].copy()
    roteirizavel_para_menu["ROTA"] = roteirizavel_para_menu["ROTA"].astype(str).str.strip()
    rotas_capital = sorted(
        roteirizavel_para_menu.loc[roteirizavel_para_menu["CAPITAL_INTERIOR"] == "Capital", "ROTA"].unique()
    )
    rotas_interior = sorted(
        roteirizavel_para_menu.loc[roteirizavel_para_menu["CAPITAL_INTERIOR"] == "Interior", "ROTA"].unique()
    )

    def _cidades_exemplo(rota):
        """Pras rotas do Interior, mostra ate 3 cidades mais frequentes
        que essa rota cobre, pra ajudar a escolher sem precisar decorar
        qual rota e qual cidade."""
        cidades = (
            roteirizavel_para_menu.loc[roteirizavel_para_menu["ROTA"] == rota, "CIDADE RECEBEDOR"]
            .dropna().astype(str).str.strip()
        )
        cidades = cidades[cidades != ""]
        if cidades.empty:
            return ""
        top = cidades.value_counts().index[:3].tolist()
        return "  (" + ", ".join(top) + ")"

    # Capital nao aparece mais rota por rota - vira uma UNICA opcao
    # "LOCAL" que já cobre todas as rotas de Capital de uma vez (a
    # maioria das vezes o Samuel quer o romaneio da Capital inteiro, nao
    # rota por rota). Interior continua uma opcao por rota. Pedido do
    # Samuel em 13/08/2026. A partir de 14/08/2026 o texto e montado
    # dinamicamente com o nome das rotas encontradas na rodada (em vez
    # de nomes fixos escritos na mao) - assim continua certo mesmo se o
    # parceiro de entrega mudar os nomes das rotas de novo.
    opcoes_menu = []
    if rotas_capital:
        descricao_local = "LOCAL - " + ", ".join(rotas_capital)
        opcoes_menu.append((descricao_local, list(rotas_capital)))
    for rota in rotas_interior:
        opcoes_menu.append((f"{rota}{_cidades_exemplo(rota)}", [rota]))

    # A partir de 14/08/2026: essa escolha passa a CONTROLAR quais
    # romaneios em PDF sao gerados (antes gerava tudo automatico +
    # ainda um PDF extra so com as rotas escolhidas - virou confuso,
    # pedido do Samuel pra simplificar). Os Excel (Capital-SSW.xlsx /
    # Interior-SSW.xlsx) continuam saindo com TUDO sempre, independente
    # da escolha - so os PDFs que ficam restritos.
    rotas_escolhidas = []
    if opcoes_menu:
        print("GERAR ROMANEIO DE QUAIS ROTAS:")
        for i, (texto, _rotas) in enumerate(opcoes_menu, start=1):
            print(f"  ({i}) {texto}")
        resposta_rotas = input(
            "Digite os numeros separados por virgula (ex: 1,3,5), ou so ENTER pra nao gerar nenhum romaneio em PDF: "
        ).strip()

        if resposta_rotas:
            indices_validos = []
            for parte in resposta_rotas.split(","):
                parte = parte.strip()
                if parte.isdigit():
                    idx = int(parte)
                    if 1 <= idx <= len(opcoes_menu):
                        indices_validos.append(idx)
            for i in sorted(set(indices_validos)):
                for r in opcoes_menu[i - 1][1]:
                    if r not in rotas_escolhidas:
                        rotas_escolhidas.append(r)
            if not rotas_escolhidas:
                print("  [aviso] Nenhum numero valido reconhecido - nenhum romaneio em PDF sera gerado.")

    arqs_romaneio = []
    if rotas_escolhidas:
        pasta_romaneios, arqs_romaneio = gerar_romaneios_pdf(
            df_capital_interior, PASTA_SAIDA, data_limite=data_limite, rotas_permitidas=rotas_escolhidas
        )
        log_detalhe(f"[rotas escolhidas] usuario escolheu: {rotas_escolhidas}")
    else:
        print("  [info] Nenhuma rota escolhida - so os Excel (Capital/Interior) foram gerados, sem romaneio em PDF.")
        log_detalhe("[rotas escolhidas] nenhuma rota escolhida - PDF nao foi gerado")

    print()

    print()
    print(f"  Capital - SSW.xlsx: {n_capital} REMESSA(S)")
    print(f"  Interior - SSW.xlsx: {n_interior} REMESSA(S)")
    if n_aguardando:
        print(f"  Aguardando Validação (dentro do Capital - SSW.xlsx): {n_aguardando} REMESSA(S)")
    print(f"  Liberado pra rota: {resumo['total_liberado_rota']} | Ainda em transferencia: {resumo['total_nao_liberado']}")
    if resumo.get("total_possivel_fila_descarregamento"):
        print(f"  Possivel fila de descarregamento (dentro dos liberados): {resumo['total_possivel_fila_descarregamento']}")
    print(f"  Romaneios em PDF: {len(arqs_romaneio)} arquivo(s) em {pasta_romaneios}")

    copia_saida_para_desktop(PASTA_SAIDA, arq_capital, arq_interior)
    print()
    print(f"[{agora_brasil()}] Processamento concluido com sucesso!")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        mostra_log_por_causa_de_erro()
        print(f"[ERRO] {e}")
        raise
