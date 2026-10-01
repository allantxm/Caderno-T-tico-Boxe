import pandas as pd
import streamlit as st
from io import BytesIO


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Caderno Tático de Boxe",
    page_icon="🥊",
    layout="centered"
)

st.title("🥊 Caderno Tático de Boxe")
st.subheader("Dicas do Pantera & Anotações de Treino")


# ============================================================
# ESTRUTURA PADRÃO DA PLANILHA
# ============================================================

COLUNAS = [
    "titulo",
    "categoria",
    "descricao",
    "fonte"
]


# ============================================================
# INICIALIZAÇÃO DA MEMÓRIA DA SESSÃO
# ============================================================

if "df_dicas" not in st.session_state:
    st.session_state.df_dicas = pd.DataFrame(columns=COLUNAS)

if "arquivo_carregado" not in st.session_state:
    st.session_state.arquivo_carregado = False


# ============================================================
# FUNÇÕES
# ============================================================

def preparar_dataframe(df):
    """
    Garante que a planilha tenha as colunas necessárias.
    """

    # Remove espaços extras dos nomes das colunas
    df.columns = df.columns.astype(str).str.strip().str.lower()

    # Verifica se existem as colunas obrigatórias
    colunas_faltantes = [
        coluna for coluna in COLUNAS
        if coluna not in df.columns
    ]

    if colunas_faltantes:
        return None, colunas_faltantes

    # Mantém somente as colunas utilizadas pelo aplicativo
    df = df[COLUNAS].copy()

    # Substitui valores vazios
    df = df.fillna("")

    return df, []


def gerar_excel(df):
    """
    Converte o DataFrame para um arquivo Excel em memória.
    """

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Caderno Tático"
        )

    output.seek(0)

    return output


# ============================================================
# ÁREA DE IMPORTAÇÃO / EXPORTAÇÃO
# ============================================================

st.markdown("### 📁 Gerenciar Caderno")

col1, col2, col3 = st.columns(3)


# ------------------------------------------------------------
# IMPORTAR
# ------------------------------------------------------------

with col1:

    arquivo = st.file_uploader(
        "📂 Importar planilha",
        type=["xlsx"],
        help="Importe uma planilha Excel com as colunas: titulo, categoria, descricao e fonte."
    )

    if arquivo is not None:

        try:

            df_importado = pd.read_excel(
                arquivo,
                sheet_name=0
            )

            df_importado, erros = preparar_dataframe(
                df_importado
            )

            if erros:

                st.error(
                    "A planilha não possui as colunas necessárias: "
                    + ", ".join(erros)
                )

            else:

                st.session_state.df_dicas = df_importado
                st.session_state.arquivo_carregado = True

                st.success(
                    f"✅ {len(df_importado)} dicas importadas!"
                )

        except Exception as e:

            st.error(
                f"Erro ao importar a planilha: {e}"
            )


# ------------------------------------------------------------
# SALVAR NA MEMÓRIA
# ------------------------------------------------------------

with col2:

    if st.button(
        "💾 Salvar na memória",
        use_container_width=True
    ):

        # O DataFrame já está no session_state.
        # Este botão serve para confirmar que as alterações
        # atuais foram mantidas.

        st.session_state.df_dicas = (
            st.session_state.df_dicas.copy()
        )

        st.success(
            "✅ Dados salvos na memória da sessão!"
        )


# ------------------------------------------------------------
# EXPORTAR
# ------------------------------------------------------------

with col3:

    arquivo_excel = gerar_excel(
        st.session_state.df_dicas
    )

    st.download_button(
        label="📥 Exportar Excel",
        data=arquivo_excel,
        file_name="Caderno_Tatico_de_Boxe.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )


st.divider()


# ============================================================
# CADASTRAR NOVA DICA
# ============================================================

with st.expander(
    "➕ Cadastrar Nova Dica / Conceito",
    expanded=True
):

    with st.form(
        "form_dica",
        clear_on_submit=True
    ):

        titulo = st.text_input(
            "Título da Dica",
            placeholder="Ex: Ajuste no Cruzado de Esquerda"
        )

        categoria = st.selectbox(
            "Categoria",
            [
                "Footwork / Movimentação",
                "Ataque & Combos",
                "Defesa & Esquivas",
                "Estratégia & Sparring",
                "Condicionamento",
            ]
        )

        descricao = st.text_area(
            "Detalhes / Passo a Passo",
            placeholder=(
                "Ex: Manter peso na perna de trás..."
            ),
            height=150
        )

        fonte = st.text_input(
            "Fonte / Link",
            value="Guilherme Pantera"
        )

        submitted = st.form_submit_button(
            "💾 Salvar Dica",
            use_container_width=True
        )

        if submitted:

            if titulo and descricao:

                nova_dica = pd.DataFrame(
                    [{
                        "titulo": titulo,
                        "categoria": categoria,
                        "descricao": descricao,
                        "fonte": fonte
                    }]
                )

                st.session_state.df_dicas = pd.concat(
                    [
                        st.session_state.df_dicas,
                        nova_dica
                    ],
                    ignore_index=True
                )

                st.success(
                    "✅ Dica adicionada ao caderno!"
                )

                st.rerun()

            else:

                st.warning(
                    "Preencha o título e a descrição!"
                )


st.divider()


# ============================================================
# FILTROS
# ============================================================

df_dicas = st.session_state.df_dicas


if not df_dicas.empty:

    st.sidebar.header("🔍 Filtros")

    categorias_existentes = (
        df_dicas["categoria"]
        .dropna()
        .unique()
        .tolist()
    )

    filtro_cat = st.sidebar.selectbox(
        "Filtrar por Categoria",
        ["Todas"] + categorias_existentes
    )

    if filtro_cat != "Todas":

        dicas_filtradas = df_dicas[
            df_dicas["categoria"] == filtro_cat
        ]

    else:

        dicas_filtradas = df_dicas


    # ========================================================
    # CONTADOR
    # ========================================================

    st.write(
        f"### 📋 Dicas Registradas "
        f"({len(dicas_filtradas)})"
    )


    # ========================================================
    # EXIBIÇÃO DAS DICAS
    # ========================================================

    for idx, row in dicas_filtradas.iloc[::-1].iterrows():

        with st.container(border=True):

            st.markdown(
                f"#### 🥊 {row['titulo']}"
            )

            st.caption(
                f"📌 **Categoria:** {row['categoria']} "
                f"| 👤 **Fonte:** {row['fonte']}"
            )

            st.write(
                row["descricao"]
            )


else:

    st.info(
        "Nenhuma dica cadastrada ainda. "
        "Importe uma planilha ou cadastre sua primeira dica!"
    )
