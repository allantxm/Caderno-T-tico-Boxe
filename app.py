import io
import json
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Caderno Tático de Boxe",
    page_icon="🥊",
    layout="wide",
)


# ============================================================
# CONSTANTES
# ============================================================

COLUNAS = [
    "id",
    "tipo",
    "titulo",
    "categoria",
    "descricao",
    "fonte",
    "link_video",
    "dificuldade",
    "favorito",
    "data",
]

TIPOS = [
    "Dica",
    "Combo",
    "Defesa",
    "Erro no Sparring",
    "Correção do Professor",
    "Treino",
]

CATEGORIAS = [
    "Footwork / Movimentação",
    "Ataque & Combos",
    "Defesa & Esquivas",
    "Estratégia & Sparring",
    "Técnica",
    "Condicionamento",
    "Mentalidade",
    "Outro",
]

DIFICULDADES = [
    "Fácil",
    "Médio",
    "Difícil",
]

GITHUB_API = "https://api.github.com"


# ============================================================
# CONFIGURAÇÃO DO GITHUB
# ============================================================

def obter_config_github():
    """
    Lê as configurações do GitHub nos Secrets do Streamlit.
    """

    token = st.secrets.get("GITHUB_TOKEN", "")
    repo = st.secrets.get("GITHUB_REPO", "")
    arquivo = st.secrets.get(
        "GITHUB_FILE",
        "data/caderno_boxe.json"
    )

    return token, repo, arquivo


# ============================================================
# DATAFRAME
# ============================================================

def criar_dataframe_vazio():
    """
    Cria um DataFrame vazio com todas as colunas necessárias.
    """

    return pd.DataFrame(columns=COLUNAS)


def preparar_dataframe(df):
    """
    Garante que o DataFrame tenha todas as colunas esperadas
    e converte dados antigos para o novo padrão.
    """

    if df is None:
        return criar_dataframe_vazio()

    df = df.copy()

    # --------------------------------------------------------
    # Compatibilidade com modelo antigo
    # --------------------------------------------------------

    if len(df.columns) > 0:

        # Colunas antigas possíveis
        if "descricao" not in df.columns:
            df["descricao"] = ""

        if "fonte" not in df.columns:
            df["fonte"] = ""

    # --------------------------------------------------------
    # Criar colunas que não existem
    # --------------------------------------------------------

    valores_padrao = {
        "id": "",
        "tipo": "Dica",
        "titulo": "",
        "categoria": "Outro",
        "descricao": "",
        "fonte": "",
        "link_video": "",
        "dificuldade": "Médio",
        "favorito": False,
        "data": "",
    }

    for coluna, valor in valores_padrao.items():
        if coluna not in df.columns:
            df[coluna] = valor

    # --------------------------------------------------------
    # Selecionar somente as colunas oficiais
    # --------------------------------------------------------

    df = df[COLUNAS]

    # --------------------------------------------------------
    # Normalizar ID
    # --------------------------------------------------------

    if len(df) > 0:

        ids = []

        for valor in df["id"]:
            if pd.isna(valor) or str(valor).strip() == "":
                ids.append("")
            else:
                ids.append(str(valor))

        df["id"] = ids

    # --------------------------------------------------------
    # Gerar IDs ausentes
    # --------------------------------------------------------

    for indice in df.index:

        if not str(df.at[indice, "id"]).strip():
            df.at[indice, "id"] = (
                datetime.now().strftime("%Y%m%d%H%M%S")
                + "_"
                + str(indice)
            )

    # --------------------------------------------------------
    # Normalizar texto
    # --------------------------------------------------------

    for coluna in [
        "tipo",
        "titulo",
        "categoria",
        "descricao",
        "fonte",
        "link_video",
        "dificuldade",
        "data",
    ]:
        df[coluna] = df[coluna].fillna("").astype(str)

    # --------------------------------------------------------
    # Normalizar favorito
    # --------------------------------------------------------

    def converter_favorito(valor):

        if isinstance(valor, bool):
            return valor

        texto = str(valor).strip().lower()

        return texto in [
            "true",
            "1",
            "sim",
            "yes",
            "favorito",
        ]

    df["favorito"] = df["favorito"].apply(converter_favorito)

    return df


# ============================================================
# JSON <-> DATAFRAME
# ============================================================

def json_para_dataframe(dados):
    """
    Converte a estrutura JSON do GitHub para um DataFrame único.
    """

    registros = []

    if not isinstance(dados, dict):
        return criar_dataframe_vazio()

    # --------------------------------------------------------
    # Categorias existentes no JSON
    # --------------------------------------------------------

    categorias_json = [
        ("anotacoes", "Dica"),
        ("combos", "Combo"),
        ("defesas", "Defesa"),
        ("erros_sparring", "Erro no Sparring"),
        ("correcoes_professor", "Correção do Professor"),
        ("treinos", "Treino"),
    ]

    for chave, tipo_padrao in categorias_json:

        lista = dados.get(chave, [])

        if not isinstance(lista, list):
            continue

        for item in lista:

            if not isinstance(item, dict):
                continue

            registro = item.copy()

            if not registro.get("tipo"):
                registro["tipo"] = tipo_padrao

            registros.append(registro)

    if not registros:
        return criar_dataframe_vazio()

    return preparar_dataframe(
        pd.DataFrame(registros)
    )


def dataframe_para_json(df):
    """
    Converte o DataFrame para o formato organizado
    utilizado no arquivo JSON do GitHub.
    """

    df = preparar_dataframe(df)

    resultado = {
        "anotacoes": [],
        "combos": [],
        "defesas": [],
        "erros_sparring": [],
        "correcoes_professor": [],
        "treinos": [],
    }

    mapa_tipos = {
        "Dica": "anotacoes",
        "Combo": "combos",
        "Defesa": "defesas",
        "Erro no Sparring": "erros_sparring",
        "Correção do Professor": "correcoes_professor",
        "Treino": "treinos",
    }

    for _, linha in df.iterrows():

        registro = {}

        for coluna in COLUNAS:

            valor = linha[coluna]

            if pd.isna(valor):
                valor = ""

            if coluna == "favorito":
                valor = bool(valor)

            else:
                valor = str(valor)

            registro[coluna] = valor

        tipo = registro.get("tipo", "Dica")

        chave = mapa_tipos.get(
            tipo,
            "anotacoes"
        )

        resultado[chave].append(registro)

    return resultado


# ============================================================
# GITHUB
# ============================================================

def github_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def carregar_github():
    """
    Carrega o JSON diretamente do GitHub.
    """

    token, repo, arquivo = obter_config_github()

    if not token or not repo:
        return (
            criar_dataframe_vazio(),
            "GITHUB_TOKEN ou GITHUB_REPO não configurado."
        )

    url = (
        f"{GITHUB_API}/repos/"
        f"{repo}/contents/{arquivo}"
    )

    try:

        resposta = requests.get(
            url,
            headers=github_headers(token),
            timeout=20,
        )

        if resposta.status_code == 404:
            return (
                criar_dataframe_vazio(),
                "Arquivo JSON não encontrado no GitHub."
            )

        resposta.raise_for_status()

        dados = resposta.json()

        conteudo = dados.get("content", "")

        import base64

        texto = base64.b64decode(
            conteudo
        ).decode("utf-8")

        json_data = json.loads(texto)

        df = json_para_dataframe(json_data)

        return df, ""

    except Exception as erro:

        return (
            criar_dataframe_vazio(),
            f"Erro ao carregar GitHub: {erro}"
        )


def salvar_github(df):
    """
    Salva o DataFrame no JSON do GitHub.
    """

    token, repo, arquivo = obter_config_github()

    if not token or not repo:
        return False, "GITHUB_TOKEN ou GITHUB_REPO não configurado."

    url = (
        f"{GITHUB_API}/repos/"
        f"{repo}/contents/{arquivo}"
    )

    try:

        # ----------------------------------------------------
        # Primeiro precisamos obter o SHA atual
        # ----------------------------------------------------

        resposta_get = requests.get(
            url,
            headers=github_headers(token),
            timeout=20,
        )

        sha = None

        if resposta_get.status_code == 200:

            dados_arquivo = resposta_get.json()

            sha = dados_arquivo.get("sha")

        elif resposta_get.status_code != 404:

            resposta_get.raise_for_status()

        # ----------------------------------------------------
        # Converter DataFrame para JSON
        # ----------------------------------------------------

        json_data = dataframe_para_json(df)

        texto_json = json.dumps(
            json_data,
            ensure_ascii=False,
            indent=2,
        )

        import base64

        conteudo_base64 = base64.b64encode(
            texto_json.encode("utf-8")
        ).decode("utf-8")

        payload = {
            "message": "Atualiza Caderno Tático de Boxe",
            "content": conteudo_base64,
        }

        if sha:
            payload["sha"] = sha

        # ----------------------------------------------------
        # Enviar para GitHub
        # ----------------------------------------------------

        resposta_put = requests.put(
            url,
            headers=github_headers(token),
            json=payload,
            timeout=20,
        )

        resposta_put.raise_for_status()

        return True, "Dados salvos no GitHub com sucesso."

    except requests.exceptions.HTTPError as erro:

        try:
            detalhe = resposta_put.json()
        except Exception:
            detalhe = ""

        return (
            False,
            f"Erro HTTP ao salvar GitHub: {erro} {detalhe}"
        )

    except Exception as erro:

        return (
            False,
            f"Erro ao salvar GitHub: {erro}"
        )


# ============================================================
# EXCEL
# ============================================================

def gerar_excel(df):
    """
    Gera o Excel usando openpyxl diretamente.

    Isso evita o erro:
    'At least one sheet must be visible'
    """

    export_df = preparar_dataframe(df)

    # --------------------------------------------------------
    # Criar workbook diretamente
    # --------------------------------------------------------

    wb = Workbook()

    # O Workbook() já cria uma planilha visível.
    ws = wb.active

    ws.title = "Caderno Tático"

    # --------------------------------------------------------
    # Cabeçalho
    # --------------------------------------------------------

    ws.append(
        list(export_df.columns)
    )

    # --------------------------------------------------------
    # Dados
    # --------------------------------------------------------

    for linha in export_df.itertuples(
        index=False,
        name=None
    ):
        ws.append(
            list(linha)
        )

    # --------------------------------------------------------
    # Estilo do cabeçalho
    # --------------------------------------------------------

    for cell in ws[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # --------------------------------------------------------
    # Ajustar largura das colunas
    # --------------------------------------------------------

    larguras = {
        "A": 24,
        "B": 24,
        "C": 35,
        "D": 30,
        "E": 70,
        "F": 30,
        "G": 50,
        "H": 15,
        "I": 12,
        "J": 20,
    }

    for coluna, largura in larguras.items():

        ws.column_dimensions[
            coluna
        ].width = largura

    # --------------------------------------------------------
    # Congelar cabeçalho
    # --------------------------------------------------------

    ws.freeze_panes = "A2"

    # --------------------------------------------------------
    # Salvar em memória
    # --------------------------------------------------------

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def gerar_id():
    return (
        datetime.now().strftime(
            "%Y%m%d%H%M%S%f"
        )
    )


def formatar_data(data):
    if not data:
        return ""

    try:

        data_obj = pd.to_datetime(data)

        return data_obj.strftime(
            "%d/%m/%Y"
        )

    except Exception:

        return str(data)


def salvar_registro(
    tipo,
    titulo,
    categoria,
    descricao,
    fonte,
    link_video,
    dificuldade,
    favorito,
):
    novo = {
        "id": gerar_id(),
        "tipo": tipo,
        "titulo": titulo,
        "categoria": categoria,
        "descricao": descricao,
        "fonte": fonte,
        "link_video": link_video,
        "dificuldade": dificuldade,
        "favorito": favorito,
        "data": datetime.now().strftime(
            "%Y-%m-%d"
        ),
    }

    return novo


# ============================================================
# SESSION STATE
# ============================================================

if "df" not in st.session_state:

    df_inicial, erro = carregar_github()

    st.session_state.df = preparar_dataframe(
        df_inicial
    )

    st.session_state.erro_conexao = erro


if "pagina" not in st.session_state:

    st.session_state.pagina = "Visão Geral"


if "editando_id" not in st.session_state:

    st.session_state.editando_id = None


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.title("🥊 Caderno Tático")

    st.caption(
        "Seu diário pessoal de Boxe"
    )

    st.divider()

    paginas = [
        "🏠 Visão Geral",
        "📚 Caderno",
        "🥊 Combos",
        "🛡️ Defesas",
        "⚠️ Erros no Sparring",
        "🎯 Correções do Professor",
        "🏋️ Treinos",
        "⭐ Favoritos",
    ]

    for pagina in paginas:

        nome_pagina = pagina.split(
            " ",
            1
        )[1]

        if st.button(
            pagina,
            use_container_width=True,
        ):

            st.session_state.pagina = (
                nome_pagina
            )

            st.session_state.editando_id = None

            st.rerun()

    st.divider()

    st.subheader("💾 Dados")

    # --------------------------------------------------------
    # Salvar GitHub
    # --------------------------------------------------------

    if st.button(
        "💾 Salvar na nuvem",
        use_container_width=True,
    ):

        sucesso, mensagem = salvar_github(
            st.session_state.df
        )

        if sucesso:
            st.success(mensagem)

        else:
            st.error(mensagem)

    # --------------------------------------------------------
    # Recarregar GitHub
    # --------------------------------------------------------

    if st.button(
        "🔄 Recarregar da nuvem",
        use_container_width=True,
    ):

        df_novo, erro = carregar_github()

        if erro:

            st.error(erro)

        else:

            st.session_state.df = (
                preparar_dataframe(df_novo)
            )

            st.session_state.editando_id = None

            st.success(
                "Dados recarregados."
            )

            st.rerun()

    st.divider()

    # --------------------------------------------------------
    # Importar Excel
    # --------------------------------------------------------

    arquivo_excel = st.file_uploader(
        "📤 Importar Excel",
        type=["xlsx"],
    )

    if arquivo_excel is not None:

        try:

            df_importado = pd.read_excel(
                arquivo_excel
            )

            st.session_state.df = (
                preparar_dataframe(
                    df_importado
                )
            )

            st.success(
                "Excel importado."
            )

        except Exception as erro:

            st.error(
                f"Erro ao importar Excel: {erro}"
            )

    # --------------------------------------------------------
    # Exportar Excel
    # --------------------------------------------------------

    try:

        excel_bytes = gerar_excel(
            st.session_state.df
        )

        st.download_button(
            label="📥 Exportar Excel",
            data=excel_bytes,
            file_name="caderno_tatico_boxe.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    except Exception as erro:

        st.error(
            f"Erro ao gerar Excel: {erro}"
        )

    # --------------------------------------------------------
    # Erro de conexão
    # --------------------------------------------------------

    if st.session_state.get(
        "erro_conexao"
    ):

        st.warning(
            st.session_state.erro_conexao
        )


# ============================================================
# FUNÇÃO DE LISTAGEM
# ============================================================

def mostrar_lista(
    df,
    titulo=None,
    tipo_filtro=None,
):
    """
    Exibe filtros e registros.
    """

    if titulo:

        st.header(titulo)

    if tipo_filtro:

        df = df[
            df["tipo"] == tipo_filtro
        ].copy()

    if len(df) == 0:

        st.info(
            "Nenhum registro encontrado."
        )

        return

    # --------------------------------------------------------
    # Filtros
    # --------------------------------------------------------

    col1, col2, col3, col4 = st.columns(
        [2, 2, 2, 3]
    )

    with col1:

        categorias_disponiveis = [
            "Todas"
        ] + sorted(
            [
                x
                for x in df["categoria"].unique()
                if x
            ]
        )

        filtro_categoria = st.selectbox(
            "Categoria",
            categorias_disponiveis,
            key=f"categoria_{titulo}_{tipo_filtro}",
        )

    with col2:

        dificuldades_disponiveis = [
            "Todas"
        ] + DIFICULDADES

        filtro_dificuldade = st.selectbox(
            "Dificuldade",
            dificuldades_disponiveis,
            key=f"dificuldade_{titulo}_{tipo_filtro}",
        )

    with col3:

        filtro_favoritos = st.selectbox(
            "Favoritos",
            [
                "Todos",
                "Somente favoritos",
                "Não favoritos",
            ],
            key=f"favoritos_{titulo}_{tipo_filtro}",
        )

    with col4:

        busca = st.text_input(
            "🔎 Buscar",
            placeholder=(
                "Título, descrição, fonte..."
            ),
            key=f"busca_{titulo}_{tipo_filtro}",
        )

    # --------------------------------------------------------
    # Aplicar filtros
    # --------------------------------------------------------

    resultado = df.copy()

    if filtro_categoria != "Todas":

        resultado = resultado[
            resultado["categoria"]
            == filtro_categoria
        ]

    if filtro_dificuldade != "Todas":

        resultado = resultado[
            resultado["dificuldade"]
            == filtro_dificuldade
        ]

    if filtro_favoritos == "Somente favoritos":

        resultado = resultado[
            resultado["favorito"] == True
        ]

    elif filtro_favoritos == "Não favoritos":

        resultado = resultado[
            resultado["favorito"] == False
        ]

    if busca.strip():

        termo = busca.lower().strip()

        mascara = (
            resultado["titulo"].str.lower().str.contains(
                termo,
                na=False
            )
            |
            resultado["descricao"].str.lower().str.contains(
                termo,
                na=False
            )
            |
            resultado["categoria"].str.lower().str.contains(
                termo,
                na=False
            )
            |
            resultado["fonte"].str.lower().str.contains(
                termo,
                na=False
            )
        )

        resultado = resultado[
            mascara
        ]

    st.caption(
        f"{len(resultado)} registro(s)"
    )

    # --------------------------------------------------------
    # Exibir registros
    # --------------------------------------------------------

    for _, registro in resultado.iterrows():

        favorito = (
            "⭐"
            if registro["favorito"]
            else "☆"
        )

        with st.container(
            border=True
        ):

            col_a, col_b = st.columns(
                [6, 1]
            )

            with col_a:

                st.subheader(
                    f"{favorito} {registro['titulo']}"
                )

                info = (
                    f"**{registro['categoria']}**"
                )

                if registro["dificuldade"]:

                    info += (
                        f" • {registro['dificuldade']}"
                    )

                if registro["data"]:

                    info += (
                        f" • {formatar_data(registro['data'])}"
                    )

                st.markdown(info)

                if registro["descricao"]:

                    st.write(
                        registro["descricao"]
                    )

                if registro["fonte"]:

                    st.caption(
                        f"Fonte: {registro['fonte']}"
                    )

                if registro["link_video"]:

                    st.link_button(
                        "▶️ Abrir vídeo",
                        registro["link_video"],
                    )

            with col_b:

                if st.button(
                    "✏️",
                    key=f"editar_{registro['id']}",
                    help="Editar",
                ):

                    st.session_state.editando_id = (
                        registro["id"]
                    )

                    st.rerun()

                if st.button(
                    "🗑️",
                    key=f"excluir_{registro['id']}",
                    help="Excluir",
                ):

                    st.session_state.df = (
                        st.session_state.df[
                            st.session_state.df["id"]
                            != registro["id"]
                        ].reset_index(drop=True)
                    )

                    st.success(
                        "Registro excluído."
                    )

                    st.rerun()


# ============================================================
# FORMULÁRIO DE NOVO REGISTRO
# ============================================================

def formulario_novo(tipo_padrao=None):

    st.subheader(
        "➕ Novo registro"
    )

    with st.form(
        "form_novo_registro",
        clear_on_submit=True,
    ):

        col1, col2 = st.columns(2)

        with col1:

            tipo = st.selectbox(
                "Tipo",
                TIPOS,
                index=(
                    TIPOS.index(tipo_padrao)
                    if tipo_padrao in TIPOS
                    else 0
                ),
            )

            titulo = st.text_input(
                "Título *"
            )

            categoria = st.selectbox(
                "Categoria",
                CATEGORIAS,
            )

            dificuldade = st.selectbox(
                "Dificuldade",
                DIFICULDADES,
            )

        with col2:

            fonte = st.text_input(
                "Fonte"
            )

            link_video = st.text_input(
                "Link do vídeo"
            )

            favorito = st.checkbox(
                "⭐ Favorito"
            )

        descricao = st.text_area(
            "Descrição / Anotação *",
            height=180,
        )

        enviar = st.form_submit_button(
            "➕ Adicionar registro",
            use_container_width=True,
        )

    if enviar:

        if not titulo.strip():

            st.error(
                "Preencha o título."
            )

            return

        if not descricao.strip():

            st.error(
                "Preencha a descrição."
            )

            return

        novo = salvar_registro(
            tipo=tipo,
            titulo=titulo.strip(),
            categoria=categoria,
            descricao=descricao.strip(),
            fonte=fonte.strip(),
            link_video=link_video.strip(),
            dificuldade=dificuldade,
            favorito=favorito,
        )

        novo_df = pd.DataFrame(
            [novo]
        )

        st.session_state.df = pd.concat(
            [
                st.session_state.df,
                novo_df,
            ],
            ignore_index=True,
        )

        st.success(
            "Registro adicionado!"
        )

        st.rerun()


# ============================================================
# FORMULÁRIO DE EDIÇÃO
# ============================================================

def formulario_edicao(registro):

    st.subheader(
        "✏️ Editar registro"
    )

    with st.form(
        "form_edicao"
    ):

        col1, col2 = st.columns(2)

        with col1:

            tipo = st.selectbox(
                "Tipo",
                TIPOS,
                index=(
                    TIPOS.index(registro["tipo"])
                    if registro["tipo"] in TIPOS
                    else 0
                ),
            )

            titulo = st.text_input(
                "Título",
                value=registro["titulo"],
            )

            categoria = st.selectbox(
                "Categoria",
                CATEGORIAS,
                index=(
                    CATEGORIAS.index(
                        registro["categoria"]
                    )
                    if registro["categoria"]
                    in CATEGORIAS
                    else 0
                ),
            )

            dificuldade = st.selectbox(
                "Dificuldade",
                DIFICULDADES,
                index=(
                    DIFICULDADES.index(
                        registro["dificuldade"]
                    )
                    if registro["dificuldade"]
                    in DIFICULDADES
                    else 1
                ),
            )

        with col2:

            fonte = st.text_input(
                "Fonte",
                value=registro["fonte"],
            )

            link_video = st.text_input(
                "Link do vídeo",
                value=registro["link_video"],
            )

            favorito = st.checkbox(
                "⭐ Favorito",
                value=bool(
                    registro["favorito"]
                ),
            )

        descricao = st.text_area(
            "Descrição / Anotação",
            value=registro["descricao"],
            height=180,
        )

        col_a, col_b = st.columns(2)

        with col_a:

            salvar = st.form_submit_button(
                "💾 Salvar alterações",
                use_container_width=True,
            )

        with col_b:

            cancelar = st.form_submit_button(
                "❌ Cancelar",
                use_container_width=True,
            )

    if cancelar:

        st.session_state.editando_id = None

        st.rerun()

    if salvar:

        if not titulo.strip():

            st.error(
                "Preencha o título."
            )

            return

        if not descricao.strip():

            st.error(
                "Preencha a descrição."
            )

            return

        indice = st.session_state.df.index[
            st.session_state.df["id"]
            == registro["id"]
        ]

        if len(indice) > 0:

            i = indice[0]

            st.session_state.df.at[
                i,
                "tipo"
            ] = tipo

            st.session_state.df.at[
                i,
                "titulo"
            ] = titulo.strip()

            st.session_state.df.at[
                i,
                "categoria"
            ] = categoria

            st.session_state.df.at[
                i,
                "descricao"
            ] = descricao.strip()

            st.session_state.df.at[
                i,
                "fonte"
            ] = fonte.strip()

            st.session_state.df.at[
                i,
                "link_video"
            ] = link_video.strip()

            st.session_state.df.at[
                i,
                "dificuldade"
            ] = dificuldade

            st.session_state.df.at[
                i,
                "favorito"
            ] = favorito

            st.session_state.editando_id = None

            st.success(
                "Registro atualizado!"
            )

            st.rerun()


# ============================================================
# VISÃO GERAL
# ============================================================

def pagina_visao_geral():

    st.title(
        "🥊 Caderno Tático de Boxe"
    )

    st.markdown(
        "Seu diário pessoal de aprendizado, "
        "treinos e evolução no Boxe."
    )

    df = st.session_state.df

    # --------------------------------------------------------
    # Métricas
    # --------------------------------------------------------

    total = len(df)

    combos = len(
        df[df["tipo"] == "Combo"]
    )

    defesas = len(
        df[df["tipo"] == "Defesa"]
    )

    erros = len(
        df[df["tipo"] == "Erro no Sparring"]
    )

    correcoes = len(
        df[df["tipo"] == "Correção do Professor"]
    )

    treinos = len(
        df[df["tipo"] == "Treino"]
    )

    favoritos = len(
        df[df["favorito"] == True]
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "📚 Total",
        total
    )

    c2.metric(
        "🥊 Combos",
        combos
    )

    c3.metric(
        "🛡️ Defesas",
        defesas
    )

    c4.metric(
        "⭐ Favoritos",
        favoritos
    )

    c5, c6, c7, c8 = st.columns(4)

    c5.metric(
        "⚠️ Erros",
        erros
    )

    c6.metric(
        "🎯 Correções",
        correcoes
    )

    c7.metric(
        "🏋️ Treinos",
        treinos
    )

    c8.metric(
        "🥊 Registros",
        total
    )

    st.divider()

    # --------------------------------------------------------
    # Últimos registros
    # --------------------------------------------------------

    st.subheader(
        "🕐 Últimos registros"
    )

    if len(df) == 0:

        st.info(
            "Seu caderno ainda está vazio. "
            "Comece adicionando sua primeira anotação."
        )

    else:

        ultimos = df.tail(5).iloc[::-1]

        for _, registro in ultimos.iterrows():

            favorito = (
                "⭐"
                if registro["favorito"]
                else ""
            )

            st.markdown(
                f"**{favorito} {registro['titulo']}**  \n"
                f"{registro['tipo']} • "
                f"{registro['categoria']}"
            )

            if registro["descricao"]:

                st.caption(
                    registro["descricao"][:180]
                    + (
                        "..."
                        if len(registro["descricao"]) > 180
                        else ""
                    )
                )

            st.divider()


# ============================================================
# CADERNO
# ============================================================

def pagina_caderno():

    st.title(
        "📚 Caderno"
    )

    if st.session_state.editando_id:

        registros = st.session_state.df[
            st.session_state.df["id"]
            == st.session_state.editando_id
        ]

        if len(registros) > 0:

            formulario_edicao(
                registros.iloc[0]
            )

            st.divider()

    formulario_novo()

    st.divider()

    mostrar_lista(
        st.session_state.df,
        titulo="📚 Todos os registros",
    )


# ============================================================
# PÁGINAS ESPECÍFICAS
# ============================================================

def pagina_combos():

    st.title(
        "🥊 Combos"
    )

    formulario_novo(
        tipo_padrao="Combo"
    )

    st.divider()

    mostrar_lista(
        st.session_state.df,
        tipo_filtro="Combo",
    )


def pagina_defesas():

    st.title(
        "🛡️ Defesas"
    )

    formulario_novo(
        tipo_padrao="Defesa"
    )

    st.divider()

    mostrar_lista(
        st.session_state.df,
        tipo_filtro="Defesa",
    )


def pagina_erros():

    st.title(
        "⚠️ Erros no Sparring"
    )

    formulario_novo(
        tipo_padrao="Erro no Sparring"
    )

    st.divider()

    mostrar_lista(
        st.session_state.df,
        tipo_filtro="Erro no Sparring",
    )


def pagina_correcoes():

    st.title(
        "🎯 Correções do Professor"
    )

    formulario_novo(
        tipo_padrao="Correção do Professor"
    )

    st.divider()

    mostrar_lista(
        st.session_state.df,
        tipo_filtro="Correção do Professor",
    )


def pagina_treinos():

    st.title(
        "🏋️ Treinos"
    )

    formulario_novo(
        tipo_padrao="Treino"
    )

    st.divider()

    mostrar_lista(
        st.session_state.df,
        tipo_filtro="Treino",
    )


def pagina_favoritos():

    st.title(
        "⭐ Favoritos"
    )

    df = st.session_state.df[
        st.session_state.df["favorito"] == True
    ].copy()

    mostrar_lista(
        df,
        titulo="⭐ Meus favoritos",
    )


# ============================================================
# ROTEAMENTO
# ============================================================

pagina = st.session_state.pagina

if pagina == "Visão Geral":

    pagina_visao_geral()

elif pagina == "Caderno":

    pagina_caderno()

elif pagina == "Combos":

    pagina_combos()

elif pagina == "Defesas":

    pagina_defesas()

elif pagina == "Erros no Sparring":

    pagina_erros()

elif pagina == "Correções do Professor":

    pagina_correcoes()

elif pagina == "Treinos":

    pagina_treinos()

elif pagina == "Favoritos":

    pagina_favoritos()

else:

    pagina_visao_geral()
