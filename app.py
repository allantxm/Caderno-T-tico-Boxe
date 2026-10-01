import io
from datetime import datetime

import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection


# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Caderno Tático de Boxe",
    page_icon="🥊",
    layout="wide",
)

st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 0;
    }
    .subtitle {
        color: #888;
        margin-top: -8px;
        margin-bottom: 25px;
    }
    .metric-card {
        padding: 18px;
        border-radius: 12px;
        border: 1px solid rgba(128,128,128,.25);
        text-align: center;
    }
    .metric-number {
        font-size: 1.8rem;
        font-weight: 800;
    }
    .metric-label {
        color: #888;
        font-size: .9rem;
    }
    .tag {
        display: inline-block;
        padding: 3px 9px;
        border-radius: 20px;
        border: 1px solid rgba(128,128,128,.3);
        font-size: .8rem;
        margin-right: 5px;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# MODELO DE DADOS
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

COLUNAS_ANTIGAS = [
    "titulo",
    "categoria",
    "descricao",
    "fonte",
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


# ============================================================
# FUNÇÕES DE DADOS
# ============================================================

def novo_id(df):
    """Gera um ID numérico simples e único."""
    if df.empty or "id" not in df.columns:
        return 1

    ids = pd.to_numeric(df["id"], errors="coerce").dropna()

    if ids.empty:
        return 1

    return int(ids.max()) + 1


def preparar_dataframe(df):
    """
    Converte a planilha antiga ou nova para o formato atual.
    """

    if df is None:
        return pd.DataFrame(columns=COLUNAS)

    df = df.copy()

    # Padroniza nomes das colunas
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
    )

    # Compatibilidade com a planilha antiga
    if "titulo" not in df.columns:
        df["titulo"] = ""

    if "categoria" not in df.columns:
        df["categoria"] = ""

    if "descricao" not in df.columns:
        df["descricao"] = ""

    if "fonte" not in df.columns:
        df["fonte"] = ""

    # Adiciona colunas novas
    defaults = {
        "id": "",
        "tipo": "Dica",
        "link_video": "",
        "dificuldade": "Médio",
        "favorito": False,
        "data": "",
    }

    for coluna, valor in defaults.items():
        if coluna not in df.columns:
            df[coluna] = valor

    # Mantém somente o modelo atual
    df = df[COLUNAS].copy()

    # Limpeza
    df = df.fillna("")

    # IDs
    ids = []
    usado = set()

    for valor in df["id"]:
        try:
            numero = int(float(valor))
            if numero <= 0 or numero in usado:
                raise ValueError
        except Exception:
            numero = 1
            while numero in usado:
                numero += 1

        usado.add(numero)
        ids.append(numero)

    df["id"] = ids

    # Tipo
    df["tipo"] = df["tipo"].replace("", "Dica")
    df.loc[~df["tipo"].isin(TIPOS), "tipo"] = "Dica"

    # Dificuldade
    df["dificuldade"] = df["dificuldade"].replace("", "Médio")
    df.loc[
        ~df["dificuldade"].isin(DIFICULDADES),
        "dificuldade"
    ] = "Médio"

    # Favorito
    df["favorito"] = df["favorito"].apply(
        lambda x: (
            True if str(x).strip().lower()
            in ["true", "1", "sim", "yes", "⭐"]
            else False
        )
    )

    # Data
    df["data"] = df["data"].astype(str)

    return df


def gerar_excel(df):
    """Gera o Excel em memória para download."""

    output = io.BytesIO()

    export_df = preparar_dataframe(df)

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        export_df.to_excel(
            writer,
            index=False,
            sheet_name="Caderno Tático"
        )

    output.seek(0)
    return output.getvalue()


def carregar_google_sheets():
    """Carrega dados do Google Sheets."""

    try:
        conn = st.connection(
            "gsheets",
            type=GSheetsConnection
        )

        df = conn.read(ttl=0)
        return preparar_dataframe(df), None

    except Exception as e:
        return (
            pd.DataFrame(columns=COLUNAS),
            str(e)
        )


def salvar_google_sheets(df):
    """Salva o DataFrame inteiro no Google Sheets."""

    conn = st.connection(
        "gsheets",
        type=GSheetsConnection
    )

    conn.update(
        data=preparar_dataframe(df)
    )


def adicionar_registro(
    df,
    tipo,
    titulo,
    categoria,
    descricao,
    fonte,
    link_video,
    dificuldade,
    favorito,
):
    """Adiciona uma anotação."""

    registro = {
        "id": novo_id(df),
        "tipo": tipo,
        "titulo": titulo.strip(),
        "categoria": categoria,
        "descricao": descricao.strip(),
        "fonte": fonte.strip(),
        "link_video": link_video.strip(),
        "dificuldade": dificuldade,
        "favorito": favorito,
        "data": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

    novo = pd.DataFrame([registro])

    return pd.concat(
        [df, novo],
        ignore_index=True
    )


# ============================================================
# MEMÓRIA DA SESSÃO
# ============================================================

if "df" not in st.session_state:
    df_inicial, erro = carregar_google_sheets()
    st.session_state.df = df_inicial
    st.session_state.erro_conexao = erro

if "pagina" not in st.session_state:
    st.session_state.pagina = "🏠 Visão Geral"


# ============================================================
# CABEÇALHO
# ============================================================

st.markdown(
    '<div class="main-title">🥊 Caderno Tático de Boxe</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">Seu diário técnico de boxe, '
    'combos, defesas e evolução nos treinos.</div>',
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown("## 🥊 Navegação")

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

    pagina = st.radio(
        "Ir para",
        paginas,
        index=paginas.index(st.session_state.pagina),
    )

    st.session_state.pagina = pagina

    st.divider()

    st.markdown("### 📁 Arquivos")

    arquivo = st.file_uploader(
        "Importar Excel",
        type=["xlsx"],
        help=(
            "Aceita sua planilha antiga com "
            "titulo, categoria, descricao e fonte."
        ),
    )

    if arquivo is not None:

        if st.button(
            "📂 Carregar planilha",
            use_container_width=True
        ):

            try:
                df_importado = pd.read_excel(
                    arquivo,
                    sheet_name=0
                )

                st.session_state.df = preparar_dataframe(
                    df_importado
                )

                st.success(
                    f"{len(st.session_state.df)} "
                    "anotações importadas."
                )

                st.rerun()

            except Exception as e:
                st.error(f"Erro ao importar: {e}")

    st.download_button(
        "📥 Exportar Excel",
        data=gerar_excel(st.session_state.df),
        file_name="Caderno_Tatico_de_Boxe.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    if st.button(
        "💾 Salvar na nuvem",
        use_container_width=True
    ):

        try:
            salvar_google_sheets(st.session_state.df)

            st.success("Dados salvos com sucesso!")

        except Exception as e:
            st.error(
                "Não foi possível salvar no Google Sheets. "
                f"Detalhes: {e}"
            )

    st.divider()

    st.caption(
        "💡 O Excel funciona como backup e "
        "transferência de dados."
    )


df = preparar_dataframe(st.session_state.df)


# ============================================================
# FUNÇÃO PARA EDITAR / EXCLUIR
# ============================================================

def editar_registro(registro_id):

    registro = df[df["id"] == registro_id]

    if registro.empty:
        st.error("Anotação não encontrada.")
        return

    row = registro.iloc[0]

    st.markdown("### ✏️ Editar anotação")

    with st.form(f"editar_{registro_id}"):

        tipo = st.selectbox(
            "Tipo",
            TIPOS,
            index=(
                TIPOS.index(row["tipo"])
                if row["tipo"] in TIPOS
                else 0
            ),
        )

        titulo = st.text_input(
            "Título",
            value=str(row["titulo"])
        )

        categoria = st.selectbox(
            "Categoria",
            CATEGORIAS,
            index=(
                CATEGORIAS.index(row["categoria"])
                if row["categoria"] in CATEGORIAS
                else len(CATEGORIAS) - 1
            ),
        )

        dificuldade = st.selectbox(
            "Dificuldade",
            DIFICULDADES,
            index=(
                DIFICULDADES.index(row["dificuldade"])
                if row["dificuldade"] in DIFICULDADES
                else 1
            ),
        )

        descricao = st.text_area(
            "Descrição",
            value=str(row["descricao"]),
            height=180
        )

        fonte = st.text_input(
            "Fonte",
            value=str(row["fonte"])
        )

        link_video = st.text_input(
            "Link de vídeo",
            value=str(row["link_video"])
        )

        favorito = st.checkbox(
            "⭐ Favorito",
            value=bool(row["favorito"])
        )

        salvar = st.form_submit_button(
            "💾 Salvar alterações",
            use_container_width=True
        )

        if salvar:

            if not titulo.strip() or not descricao.strip():
                st.warning(
                    "Título e descrição são obrigatórios."
                )

            else:

                mask = df["id"] == registro_id

                df.loc[mask, "tipo"] = tipo
                df.loc[mask, "titulo"] = titulo.strip()
                df.loc[mask, "categoria"] = categoria
                df.loc[mask, "dificuldade"] = dificuldade
                df.loc[mask, "descricao"] = descricao.strip()
                df.loc[mask, "fonte"] = fonte.strip()
                df.loc[mask, "link_video"] = link_video.strip()
                df.loc[mask, "favorito"] = favorito

                st.session_state.df = df

                st.success("Anotação atualizada!")

                st.rerun()


# ============================================================
# NOVA ANOTAÇÃO
# ============================================================

with st.expander("➕ Nova anotação", expanded=False):

    with st.form("nova_anotacao"):

        col1, col2 = st.columns(2)

        with col1:

            tipo_novo = st.selectbox(
                "Tipo",
                TIPOS,
            )

            titulo_novo = st.text_input(
                "Título",
                placeholder=(
                    "Ex: Jab + direto + gancho"
                ),
            )

            categoria_nova = st.selectbox(
                "Categoria",
                CATEGORIAS,
            )

        with col2:

            dificuldade_nova = st.selectbox(
                "Dificuldade",
                DIFICULDADES,
                index=1,
            )

            fonte_nova = st.text_input(
                "Fonte",
                placeholder=(
                    "Ex: Professor, Pantera, treino..."
                ),
            )

            link_novo = st.text_input(
                "Link de vídeo",
                placeholder="https://...",
            )

        descricao_nova = st.text_area(
            "Descrição / Passo a passo",
            placeholder=(
                "Descreva a técnica, o erro, "
                "a correção ou o treino..."
            ),
            height=150,
        )

        favorito_novo = st.checkbox(
            "⭐ Adicionar aos favoritos"
        )

        criar = st.form_submit_button(
            "🥊 Adicionar ao caderno",
            use_container_width=True
        )

        if criar:

            if not titulo_novo.strip():
                st.warning("Informe um título.")

            elif not descricao_nova.strip():
                st.warning("Informe uma descrição.")

            else:

                st.session_state.df = adicionar_registro(
                    df,
                    tipo_novo,
                    titulo_novo,
                    categoria_nova,
                    descricao_nova,
                    fonte_nova,
                    link_novo,
                    dificuldade_nova,
                    favorito_novo,
                )

                st.success(
                    "Anotação adicionada ao caderno!"
                )

                st.rerun()


# ============================================================
# COMPONENTE DE LISTAGEM
# ============================================================

def mostrar_lista(df_lista, titulo_secao):

    st.markdown(f"## {titulo_secao}")

    if df_lista.empty:

        st.info(
            "Nenhuma anotação encontrada."
        )

        return

    busca = st.text_input(
        "🔎 Buscar",
        placeholder=(
            "Digite uma palavra, técnica, "
            "erro, professor..."
        ),
        key=f"busca_{titulo_secao}",
    )

    if busca.strip():

        termo = busca.strip().lower()

        mascara = (
            df_lista["titulo"].astype(str).str.lower().str.contains(
                termo, na=False
            )
            |
            df_lista["descricao"].astype(str).str.lower().str.contains(
                termo, na=False
            )
            |
            df_lista["categoria"].astype(str).str.lower().str.contains(
                termo, na=False
            )
            |
            df_lista["fonte"].astype(str).str.lower().str.contains(
                termo, na=False
            )
        )

        df_lista = df_lista[mascara]

    col1, col2, col3 = st.columns(3)

    with col1:
        filtro_categoria = st.selectbox(
            "Categoria",
            ["Todas"] + sorted(
                [
                    str(x)
                    for x in df_lista["categoria"].dropna().unique()
                    if str(x).strip()
                ]
            ),
            key=f"cat_{titulo_secao}",
        )

    with col2:
        filtro_dificuldade = st.selectbox(
            "Dificuldade",
            ["Todas"] + DIFICULDADES,
            key=f"dif_{titulo_secao}",
        )

    with col3:
        filtro_favorito = st.selectbox(
            "Favoritos",
            ["Todos", "Somente favoritos"],
            key=f"fav_{titulo_secao}",
        )

    if filtro_categoria != "Todas":
        df_lista = df_lista[
            df_lista["categoria"] == filtro_categoria
        ]

    if filtro_dificuldade != "Todas":
        df_lista = df_lista[
            df_lista["dificuldade"] == filtro_dificuldade
        ]

    if filtro_favorito == "Somente favoritos":
        df_lista = df_lista[
            df_lista["favorito"] == True
        ]

    st.caption(
        f"{len(df_lista)} anotação(ões) encontrada(s)"
    )

    for _, row in df_lista.iloc[::-1].iterrows():

        with st.container(border=True):

            estrela = "⭐ " if row["favorito"] else ""

            st.markdown(
                f"### {estrela}{row['titulo']}"
            )

            st.markdown(
                f'<span class="tag">{row["tipo"]}</span>'
                f'<span class="tag">{row["categoria"]}</span>'
                f'<span class="tag">{row["dificuldade"]}</span>',
                unsafe_allow_html=True
            )

            st.write("")

            st.write(row["descricao"])

            info = []

            if row["fonte"]:
                info.append(f"Fonte: {row['fonte']}")

            if row["data"]:
                info.append(f"Data: {row['data']}")

            if info:
                st.caption(" | ".join(info))

            if row["link_video"]:

                st.link_button(
                    "🎥 Abrir vídeo",
                    row["link_video"],
                )

            edit_col, delete_col = st.columns(2)

            with edit_col:

                if st.button(
                    "✏️ Editar",
                    key=f"editar_{row['id']}",
                    use_container_width=True,
                ):

                    st.session_state.editando = int(
                        row["id"]
                    )

                    st.rerun()

            with delete_col:

                if st.button(
                    "🗑️ Excluir",
                    key=f"excluir_{row['id']}",
                    use_container_width=True,
                ):

                    st.session_state.confirmar_exclusao = int(
                        row["id"]
                    )

                    st.rerun()


# ============================================================
# EXCLUSÃO
# ============================================================

if "confirmar_exclusao" in st.session_state:

    excluir_id = st.session_state.confirmar_exclusao

    registro = df[df["id"] == excluir_id]

    if not registro.empty:

        st.warning(
            f"Excluir permanentemente "
            f'"{registro.iloc[0]["titulo"]}"?'
        )

        sim, nao = st.columns(2)

        with sim:

            if st.button(
                "Sim, excluir",
                type="primary",
                use_container_width=True,
            ):

                st.session_state.df = df[
                    df["id"] != excluir_id
                ].reset_index(drop=True)

                del st.session_state.confirmar_exclusao

                st.success("Anotação excluída.")
                st.rerun()

        with nao:

            if st.button(
                "Cancelar",
                use_container_width=True,
            ):

                del st.session_state.confirmar_exclusao
                st.rerun()


# ============================================================
# EDIÇÃO
# ============================================================

if "editando" in st.session_state:

    editar_id = st.session_state.editando

    editar_registro(editar_id)

    if st.button("↩️ Voltar"):
        del st.session_state.editando
        st.rerun()

    st.stop()


# ============================================================
# PÁGINA: VISÃO GERAL
# ============================================================

if pagina == "🏠 Visão Geral":

    st.markdown("## 🏠 Visão Geral")

    total = len(df)
    favoritos = int(df["favorito"].sum())

    combos = int(
        (df["tipo"] == "Combo").sum()
    )

    defesas = int(
        (df["tipo"] == "Defesa").sum()
    )

    erros = int(
        (df["tipo"] == "Erro no Sparring").sum()
    )

    correcoes = int(
        (df["tipo"] == "Correção do Professor").sum()
    )

    treinos = int(
        (df["tipo"] == "Treino").sum()
    )

    cards = [
        ("📚", total, "Anotações"),
        ("⭐", favoritos, "Favoritos"),
        ("🥊", combos, "Combos"),
        ("🛡️", defesas, "Defesas"),
        ("⚠️", erros, "Erros no Sparring"),
        ("🎯", correcoes, "Correções"),
        ("🏋️", treinos, "Treinos"),
    ]

    cols = st.columns(4)

    for i, (icone, numero, label) in enumerate(cards):

        with cols[i % 4]:

            st.markdown(
                f"""
                <div class="metric-card">
                    <div>{icone}</div>
                    <div class="metric-number">
                        {numero}
                    </div>
                    <div class="metric-label">
                        {label}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.divider()

    st.markdown("### 🧠 Últimas anotações")

    if df.empty:

        st.info(
            "Seu caderno ainda está vazio. "
            "Comece adicionando uma anotação acima."
        )

    else:

        ultimas = df.iloc[::-1].head(5)

        for _, row in ultimas.iterrows():

            estrela = "⭐ " if row["favorito"] else ""

            st.markdown(
                f"**{estrela}{row['titulo']}**  \n"
                f"{row['tipo']} · {row['categoria']}"
            )

            st.caption(
                str(row["descricao"])[:180]
                + (
                    "..."
                    if len(str(row["descricao"])) > 180
                    else ""
                )
            )

            st.divider()


# ============================================================
# PÁGINA: CADERNO
# ============================================================

elif pagina == "📚 Caderno":

    mostrar_lista(
        df,
        "📚 Caderno"
    )


# ============================================================
# PÁGINAS POR TIPO
# ============================================================

elif pagina == "🥊 Combos":

    mostrar_lista(
        df[df["tipo"] == "Combo"],
        "🥊 Combos"
    )


elif pagina == "🛡️ Defesas":

    mostrar_lista(
        df[df["tipo"] == "Defesa"],
        "🛡️ Defesas"
    )


elif pagina == "⚠️ Erros no Sparring":

    mostrar_lista(
        df[df["tipo"] == "Erro no Sparring"],
        "⚠️ Erros no Sparring"
    )


elif pagina == "🎯 Correções do Professor":

    mostrar_lista(
        df[df["tipo"] == "Correção do Professor"],
        "🎯 Correções do Professor"
    )


elif pagina == "🏋️ Treinos":

    mostrar_lista(
        df[df["tipo"] == "Treino"],
        "🏋️ Treinos"
    )


elif pagina == "⭐ Favoritos":

    mostrar_lista(
        df[df["favorito"] == True],
        "⭐ Favoritos"
    )
