# Robô SSW 081 — TRATAR BASE 081

App web (Streamlit) que lê o export do SSWWEB, classifica por rota (base de rotas
ou coluna Setor) e gera o Excel + romaneios em PDF prontos pra expedição.

**Antes de mexer em qualquer coisa**, siga o guia completo de implantação
(GitHub → Streamlit Cloud → Secrets → primeiro acesso → como adicionar um
parceiro novo): [link enviado junto com este pacote].

Resumo rápido dos arquivos:

- `app.py` — interface web (login, telas, painel do administrador).
- `processa_sswweb.py` — toda a lógica de classificação de rota, prazos e geração
  dos arquivos finais. Não depende de mais nenhum outro `.py`.
- `requirements.txt` — bibliotecas que o Streamlit Cloud instala sozinho no deploy.
- `base_rotas.xlsx` — base de rotas padrão (cidade/bairro/CEP → rota) usada hoje no RN.
- `.streamlit/config.toml` — tema visual (cores) e configuração do Streamlit.
- `.devcontainer/` — permite abrir e editar este projeto pelo GitHub Codespaces
  (navegador), sem precisar instalar nada no computador.

Este pacote **não inclui** `usuarios_autorizados.json` nem imagens de wallpaper —
isso é proposital, o guia explica como criar os seus próprios a partir do zero.
Sem as imagens de wallpaper, o app funciona 100% normal, só abre com fundo liso
(tema escuro padrão) em vez de ter uma imagem atrás — não é erro nenhum.
