import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# 1. Configuração da página
st.set_page_config(
    page_title="Caderno Tático de Boxe", page_icon="🥊", layout="centered"
)

st.title("🥊 Caderno Tático de Boxe")
st.subheader("Dicas do Pantera & Anotações de Treino")

# 2. Conexão com o Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# 3. Carregar dados existentes (garante que 'df_dicas' existe antes de ser usada)
try:
  df_dicas = conn.read(ttl=0)
except Exception:
  # Se a planilha estiver vazia, cria um DataFrame estruturado
  df_dicas = pd.DataFrame(
      columns=["titulo", "categoria", "descricao", "fonte"]
  )

# 4. Formulário para cadastrar nova dica
with st.expander("➕ Cadastrar Nova Dica / Conceito"):
  with st.form("form_dica", clear_on_submit=True):
    titulo = st.text_input("Título da Dica (ex: Ajuste no Cruzado de Esquerda)")
    categoria = st.selectbox(
        "Categoria",
        [
            "Footwork / Movimentação",
            "Ataque & Combos",
            "Defesa & Esquivas",
            "Estratégia & Sparring",
            "Condicionamento",
        ],
    )
    descricao = st.text_area(
        "Detalhes / Passo a Passo (ex: Manter peso na perna de trás...)"
    )
    fonte = st.text_input(
        "Fonte / Link (ex: Guilherme Pantera)", value="Guilherme Pantera"
    )

    submitted = st.form_submit_button("Salvar Dica")

    if submitted:
      if titulo and descricao:
        nova_dica = pd.DataFrame([{
            "titulo": titulo,
            "categoria": categoria,
            "descricao": descricao,
            "fonte": fonte,
        }])

        # Concatena a nova dica ao DataFrame existente
        df_atualizado = pd.concat([df_dicas, nova_dica], ignore_index=True)
        conn.update(data=df_atualizado)

        st.success("Dica salva na nuvem com sucesso!")
        st.rerun()
      else:
        st.warning("Preencha o título e a descrição!")

st.divider()

# 5. Filtros e Exibição das Dicas
if not df_dicas.empty and "categoria" in df_dicas.columns:
  st.sidebar.header("🔍 Filtros")
  categorias_existentes = df_dicas["categoria"].dropna().unique().tolist()
  filtro_cat = st.sidebar.selectbox(
      "Filtrar por Categoria", ["Todas"] + categorias_existentes
  )

  if filtro_cat != "Todas":
    dicas_filtradas = df_dicas[df_dicas["categoria"] == filtro_cat]
  else:
    dicas_filtradas = df_dicas

  st.write(f"### 📋 Dicas Registradas ({len(dicas_filtradas)})")

  # Exibe da mais recente para a mais antiga
  for idx, row in dicas_filtradas.iloc[::-1].iterrows():
    with st.container(border=True):
      st.markdown(f"#### {row['titulo']}")
      st.caption(
          f"📌 **Categoria:** {row['categoria']} | 👤 **Fonte:**"
          f" {row['fonte']}"
      )
      st.write(row["descricao"])
else:
  st.info("Nenhuma dica cadastrada ainda. Use o formulário acima para começar!")
