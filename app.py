"""
Caderno Tático de Boxe
Diário pessoal de técnicas, combos, defesas, erros de sparring e treinos.

Dados: ficam em memória durante a sessão e são sincronizados com um
arquivo JSON num repositório do GitHub (Secrets: GITHUB_TOKEN, GITHUB_REPO,
e opcionalmente GITHUB_FILE e GITHUB_BRANCH).
"""

import base64
import inspect
import io
import json
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from html import escape
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from zoneinfo import ZoneInfo

    FUSO = ZoneInfo("America/Sao_Paulo")
except Exception:  # tzdata ausente
    FUSO = None


# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Caderno Tático de Boxe",
    page_icon="🥊",
    layout="wide",
    initial_sidebar_state="expanded",
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
    "status",
    "favorito",
    "data",
]

# Cada tipo tem uma cor própria (a "fita" do cartão) e a chave usada no JSON.
TIPOS_INFO = {
    "Dica": {"css": "t-dica", "cor": "#5BA4E6", "chave": "anotacoes"},
    "Combo": {"css": "t-combo", "cor": "#F28C45", "chave": "combos"},
    "Defesa": {"css": "t-defesa", "cor": "#3DB5A6", "chave": "defesas"},
    "Erro no Sparring": {"css": "t-erro", "cor": "#E8C547", "chave": "erros_sparring"},
    "Correção do Professor": {"css": "t-correcao", "cor": "#A98BEA", "chave": "correcoes_professor"},
    "Treino": {"css": "t-treino", "cor": "#7CC47C", "chave": "treinos"},
}
TIPOS = list(TIPOS_INFO)

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

DIFICULDADES = ["Fácil", "Médio", "Difícil"]
STATUS = ["Praticando", "Dominado"]
TIPOS_SEM_STATUS = {"Treino"}  # treino não "se domina"
TIPOS_REVISAO = ["Erro no Sparring", "Correção do Professor"]

# (rótulo no menu, tipo filtrado ou None)
PAGINAS = [
    ("Visão geral", None),
    ("Caderno", None),
    ("Dicas", "Dica"),
    ("Combos", "Combo"),
    ("Defesas", "Defesa"),
    ("Erros no sparring", "Erro no Sparring"),
    ("Correções do professor", "Correção do Professor"),
    ("Treinos", "Treino"),
    ("Favoritos", None),
]

VALORES_PADRAO = {
    "id": "",
    "tipo": "Dica",
    "titulo": "",
    "categoria": "Outro",
    "descricao": "",
    "fonte": "",
    "link_video": "",
    "dificuldade": "Médio",
    "status": "Praticando",
    "favorito": False,
    "data": "",
}

GITHUB_API = "https://api.github.com"
ITENS_POR_PAGINA = 10

MESES = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]
DIAS_SEMANA = [
    "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
    "sexta-feira", "sábado", "domingo",
]

# Largura total de botões: o parâmetro mudou nas versões novas do Streamlit.
try:
    if "width" in inspect.signature(st.button).parameters:
        LARGURA_BTN = {"width": "stretch"}
    else:
        LARGURA_BTN = {"use_container_width": True}
except Exception:
    LARGURA_BTN = {"use_container_width": True}


# ============================================================
# UTILITÁRIOS
# ============================================================

def agora():
    """Data/hora no fuso de São Paulo (o servidor costuma rodar em UTC)."""
    return datetime.now(FUSO) if FUSO else datetime.now()


def hoje():
    return agora().date()


def novo_id():
    return uuid.uuid4().hex[:16]


def normalizar_texto(texto):
    """Minúsculas e sem acentos, para busca tolerante."""
    texto = unicodedata.normalize("NFKD", str(texto))
    return texto.encode("ascii", "ignore").decode("ascii").lower()


def normalizar_data(valor):
    """Converte vários formatos para AAAA-MM-DD (ou '' se inválido)."""
    texto = str(valor if valor is not None else "").strip()

    if not texto or texto.lower() in ("nan", "nat", "none"):
        return ""

    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto[:10], formato).strftime("%Y-%m-%d")
        except ValueError:
            continue

    try:
        return pd.to_datetime(texto).strftime("%Y-%m-%d")
    except Exception:
        return ""


def formatar_data(iso):
    if not iso:
        return ""
    try:
        return datetime.strptime(iso[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(iso)


def para_date(iso):
    try:
        return datetime.strptime(iso[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return hoje()


def normalizar_link(link):
    link = str(link or "").strip()
    if link and not link.lower().startswith(("http://", "https://")):
        link = "https://" + link
    return link


def plural(n, singular, plural_):
    return f"{n} {singular if n == 1 else plural_}"


def trecho(texto, limite=170):
    texto = " ".join(str(texto).split())
    return texto if len(texto) <= limite else texto[: limite - 1].rstrip() + "…"


def indice_de(lista, valor, padrao=0):
    return lista.index(valor) if valor in lista else padrao


# ============================================================
# DATAFRAME
# ============================================================

def criar_dataframe_vazio():
    return pd.DataFrame(columns=COLUNAS)


def converter_favorito(valor):
    if isinstance(valor, bool):
        return valor
    return str(valor).strip().lower() in ("true", "1", "sim", "yes", "favorito", "x")


def preparar_dataframe(df):
    """
    Garante colunas, tipos e valores válidos. Aceita dados antigos
    (sem 'status', sem 'descricao', etc.) e planilhas importadas.
    """
    if df is None:
        return criar_dataframe_vazio()

    df = df.copy()

    for coluna, padrao in VALORES_PADRAO.items():
        if coluna not in df.columns:
            df[coluna] = padrao

    df = df[COLUNAS].reset_index(drop=True)

    # Textos limpos
    for coluna in COLUNAS:
        if coluna == "favorito":
            continue
        df[coluna] = df[coluna].apply(
            lambda v: "" if pd.isna(v) else str(v).strip()
        )

    # IDs únicos (gera para os ausentes e para os repetidos)
    vistos, ids = set(), []
    for valor in df["id"]:
        if not valor or valor in vistos:
            valor = novo_id()
        vistos.add(valor)
        ids.append(valor)
    df["id"] = ids

    # Valores de listas fechadas (ignora maiúsculas/minúsculas)
    def forcar(valor, opcoes, padrao):
        mapa = {o.lower(): o for o in opcoes}
        return mapa.get(valor.lower(), padrao)

    df["tipo"] = df["tipo"].apply(lambda v: forcar(v, TIPOS, "Dica"))
    df["dificuldade"] = df["dificuldade"].apply(lambda v: forcar(v, DIFICULDADES, "Médio"))
    df["status"] = df["status"].apply(lambda v: forcar(v, STATUS, "Praticando"))
    df["categoria"] = df["categoria"].apply(lambda v: v or "Outro")
    df["data"] = df["data"].apply(normalizar_data)
    df["favorito"] = df["favorito"].apply(converter_favorito).astype(bool)

    return df


# ============================================================
# JSON <-> DATAFRAME
# ============================================================

def json_para_dataframe(dados):
    """Converte a estrutura JSON (listas por tipo) para um DataFrame único."""
    if not isinstance(dados, dict):
        return criar_dataframe_vazio()

    registros = []

    for tipo, info in TIPOS_INFO.items():
        lista = dados.get(info["chave"], [])

        if not isinstance(lista, list):
            continue

        for item in lista:
            if not isinstance(item, dict):
                continue
            registro = item.copy()
            if not registro.get("tipo"):
                registro["tipo"] = tipo
            registros.append(registro)

    if not registros:
        return criar_dataframe_vazio()

    return preparar_dataframe(pd.DataFrame(registros))


def dataframe_para_json(df):
    """Converte o DataFrame para o formato do arquivo JSON."""
    df = preparar_dataframe(df)
    resultado = {info["chave"]: [] for info in TIPOS_INFO.values()}

    for _, linha in df.iterrows():
        registro = {
            coluna: (bool(linha[coluna]) if coluna == "favorito" else str(linha[coluna]))
            for coluna in COLUNAS
        }
        resultado[TIPOS_INFO[registro["tipo"]]["chave"]].append(registro)

    return resultado


# ============================================================
# GITHUB
# ============================================================

def _segredo(nome, padrao=""):
    try:
        return st.secrets.get(nome, padrao)
    except Exception:  # sem secrets.toml
        return padrao


def obter_config_github():
    return (
        _segredo("GITHUB_TOKEN"),
        _segredo("GITHUB_REPO"),
        _segredo("GITHUB_FILE", "data/caderno_boxe.json"),
        _segredo("GITHUB_BRANCH"),
    )


def github_configurado():
    token, repo, _, _ = obter_config_github()
    return bool(token and repo)


def github_headers(token, aceitar="application/vnd.github+json"):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": aceitar,
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _url_arquivo(repo, arquivo):
    return f"{GITHUB_API}/repos/{repo}/contents/{quote(arquivo, safe='/')}"


def _detalhe_erro(resposta):
    try:
        return resposta.json().get("message", "")
    except Exception:
        return ""


def carregar_github():
    """
    Carrega o JSON do GitHub.
    Retorna (DataFrame, mensagem_de_erro, estado) com estado em:
    'ok', 'novo' (arquivo ainda não existe), 'erro' ou 'sem_config'.
    """
    token, repo, arquivo, branch = obter_config_github()

    if not token or not repo:
        return criar_dataframe_vazio(), "Nuvem não configurada (GITHUB_TOKEN e GITHUB_REPO).", "sem_config"

    try:
        resposta = requests.get(
            _url_arquivo(repo, arquivo),
            headers=github_headers(token, "application/vnd.github.raw+json"),
            params={"ref": branch} if branch else None,
            timeout=20,
        )

        if resposta.status_code == 404:
            return criar_dataframe_vazio(), "", "novo"

        if resposta.status_code in (401, 403):
            return (
                criar_dataframe_vazio(),
                "O GitHub recusou o acesso. Confira se o token é válido e tem permissão no repositório.",
                "erro",
            )

        resposta.raise_for_status()
        dados = json.loads(resposta.content.decode("utf-8"))
        return json_para_dataframe(dados), "", "ok"

    except Exception as erro:
        return criar_dataframe_vazio(), f"Não foi possível carregar da nuvem: {erro}", "erro"


def salvar_github(df):
    """Salva o DataFrame no JSON do GitHub. Retorna (sucesso, mensagem)."""
    token, repo, arquivo, branch = obter_config_github()

    if not token or not repo:
        return False, "Nuvem não configurada (GITHUB_TOKEN e GITHUB_REPO)."

    url = _url_arquivo(repo, arquivo)
    parametros = {"ref": branch} if branch else None

    try:
        # SHA atual do arquivo (necessário para atualizar)
        resposta_get = requests.get(
            url, headers=github_headers(token), params=parametros, timeout=20
        )

        sha = None
        if resposta_get.status_code == 200:
            sha = resposta_get.json().get("sha")
        elif resposta_get.status_code != 404:
            resposta_get.raise_for_status()

        texto_json = json.dumps(dataframe_para_json(df), ensure_ascii=False, indent=2)

        payload = {
            "message": f"Atualiza caderno ({agora():%d/%m/%Y %H:%M})",
            "content": base64.b64encode(texto_json.encode("utf-8")).decode("ascii"),
        }
        if sha:
            payload["sha"] = sha
        if branch:
            payload["branch"] = branch

        resposta_put = requests.put(
            url, headers=github_headers(token), json=payload, timeout=20
        )

        if resposta_put.status_code in (409, 422):
            return False, (
                "O arquivo mudou na nuvem enquanto você editava. "
                "Baixe um backup e recarregue da nuvem."
            )

        if resposta_put.status_code in (401, 403):
            return False, "O GitHub recusou a gravação. Confira a permissão de escrita do token."

        resposta_put.raise_for_status()
        return True, "Caderno salvo na nuvem."

    except Exception as erro:
        return False, f"Não foi possível salvar na nuvem: {erro}"


# ============================================================
# EXPORTAÇÃO
# ============================================================

def gerar_excel(df):
    """Gera o .xlsx em memória (cabeçalho estilizado, filtros e colunas ajustadas)."""
    export_df = preparar_dataframe(df)

    wb = Workbook()
    ws = wb.active
    ws.title = "Caderno Tático"
    ws.append(COLUNAS)

    for linha in export_df.itertuples(index=False, name=None):
        valores = dict(zip(COLUNAS, linha))
        valores["favorito"] = "Sim" if valores["favorito"] else "Não"
        valores["data"] = formatar_data(valores["data"])
        ws.append([valores[c] for c in COLUNAS])

    cabecalho_fill = PatternFill("solid", fgColor="162434")
    borda = Border(bottom=Side(style="thin", color="2A4057"))

    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = cabecalho_fill
        celula.alignment = Alignment(horizontal="left", vertical="center")

    for linha in ws.iter_rows(min_row=2):
        for celula in linha:
            celula.alignment = Alignment(vertical="top", wrap_text=True)
            celula.border = borda

    larguras = {
        "id": 18, "tipo": 22, "titulo": 36, "categoria": 26, "descricao": 70,
        "fonte": 26, "link_video": 40, "dificuldade": 13, "status": 13,
        "favorito": 10, "data": 12,
    }
    for posicao, coluna in enumerate(COLUNAS, start=1):
        ws.column_dimensions[get_column_letter(posicao)].width = larguras[coluna]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


def gerar_backup_json(df):
    return json.dumps(dataframe_para_json(df), ensure_ascii=False, indent=2).encode("utf-8")


# ============================================================
# ESTADO E MUTAÇÕES
# ============================================================

def iniciar_estado():
    if "df" not in st.session_state:
        with st.spinner("Carregando seu caderno..."):
            df_inicial, erro, estado = carregar_github()

        st.session_state.df = preparar_dataframe(df_inicial)
        st.session_state.erro_conexao = erro
        st.session_state.estado_nuvem = estado
        st.session_state.alterado = False
        st.session_state.ultimo_salvamento = ""
        st.session_state.autosave = estado in ("ok", "novo")

    st.session_state.setdefault("pagina", "Visão geral")
    st.session_state.setdefault("editando_id", None)
    st.session_state.setdefault("excluindo_id", None)
    st.session_state.setdefault("avisos", [])
    st.session_state.setdefault("confirmar_recarga", False)
    st.session_state.setdefault("uploader_versao", 0)


def avisar(texto, icone="✅"):
    """Guarda uma mensagem para aparecer depois do próximo rerun."""
    st.session_state.avisos.append((icone, texto))


def pode_salvar_na_nuvem():
    return github_configurado() and st.session_state.estado_nuvem in ("ok", "novo")


def salvar_na_nuvem(silencioso=False):
    """Salva agora. Retorna True se deu certo. Com silencioso=True, só avisa em caso de erro."""
    if not github_configurado():
        avisar("Configure GITHUB_TOKEN e GITHUB_REPO nos Secrets para usar a nuvem.", "⚠️")
        return False

    if not pode_salvar_na_nuvem():
        avisar(
            "O caderno não foi carregado da nuvem. Recarregue antes de salvar "
            "para não sobrescrever seus dados.",
            "⚠️",
        )
        return False

    sucesso, mensagem = salvar_github(st.session_state.df)

    if sucesso:
        st.session_state.alterado = False
        st.session_state.estado_nuvem = "ok"
        st.session_state.ultimo_salvamento = agora().strftime("%H:%M")
        if not silencioso:
            avisar(mensagem)
    else:
        avisar(mensagem, "⚠️")

    return sucesso


def confirmar_mudanca(df_novo, mensagem, icone="✅"):
    """Aplica uma mudança no caderno e salva sozinho, se estiver ativado."""
    st.session_state.df = preparar_dataframe(df_novo)
    st.session_state.alterado = True
    avisar(mensagem, icone)

    if st.session_state.autosave and pode_salvar_na_nuvem():
        salvar_na_nuvem(silencioso=True)


def atualizar_registro(id_registro, campos, mensagem, icone="✅"):
    df = st.session_state.df.copy()
    posicoes = df.index[df["id"] == id_registro]

    if len(posicoes) == 0:
        return

    for coluna, valor in campos.items():
        df.at[posicoes[0], coluna] = valor

    confirmar_mudanca(df, mensagem, icone)


def excluir_registro(id_registro):
    df = st.session_state.df
    df = df[df["id"] != id_registro].reset_index(drop=True)
    confirmar_mudanca(df, "Registro excluído.", "🗑️")


def adicionar_registro(campos):
    novo = {"id": novo_id(), **campos}
    df = pd.concat(
        [st.session_state.df, pd.DataFrame([novo])], ignore_index=True
    )
    confirmar_mudanca(df, "Registro adicionado.")


def mostrar_avisos():
    avisos, st.session_state.avisos = st.session_state.avisos, []
    for icone, texto in avisos:
        st.toast(texto, icon=icone)


# ============================================================
# ESTILO
# ============================================================

def montar_css():
    # Regras por tipo (a fita colorida lateral e as pílulas)
    por_tipo = ""
    for info in TIPOS_INFO.values():
        classe, cor = info["css"], info["cor"]
        por_tipo += (
            f".{classe}{{--c:{cor};}}"
            f"div[data-testid='stVerticalBlockBorderWrapper']:has(.cm.{classe})"
            f"{{border-left:5px solid {cor} !important;}}"
        )

    base = """
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600&family=Barlow+Condensed:wght@500;600;700&display=swap');

:root{
  --bg:#0F1A26; --surface:#162434; --raised:#1D3046; --line:#2A4057;
  --text:#E9EEF3; --muted:#98A9BB; --accent:#DB3A3F; --gold:#E8C547; --ok:#7CC47C;
}

html, body, .stApp, p, li, label, input, textarea, button, a,
[data-baseweb="select"] *{
  font-family:'Barlow', system-ui, -apple-system, 'Segoe UI', sans-serif;
}
.stApp{ background:var(--bg); color:var(--text); }

.block-container{ max-width:1160px; padding:2.4rem 2rem 4rem; }
@media (max-width:640px){ .block-container{ padding:1.4rem 1rem 3rem; } }

header[data-testid="stHeader"]{ background:transparent; }
#MainMenu, footer, .stDeployButton{ visibility:hidden; }
[data-testid="stHeaderActionElements"]{ display:none; }

/* ---------- Tipografia ---------- */
h1, h2, h3, .page-title, .sec, .brand-name{
  font-family:'Barlow Condensed','Barlow',sans-serif !important;
}
.stApp .page-title{
  font-size:3.2rem; font-weight:700; line-height:1; letter-spacing:-.01em;
  margin:0; padding:0; color:var(--text);
}
.stApp .page-sub{ color:var(--muted); font-size:1.05rem; margin:.45rem 0 1.6rem; }
.stApp .sec{
  font-size:1.65rem; font-weight:600; line-height:1.1; margin:2rem 0 .3rem;
  padding:0; color:var(--text);
}
.stApp .sec-sub{ color:var(--muted); font-size:.98rem; margin:0 0 .9rem; }
@media (max-width:640px){ .stApp .page-title{ font-size:2.5rem; } }

/* ---------- Barra lateral ---------- */
section[data-testid="stSidebar"]{
  background:var(--surface); border-right:1px solid var(--line);
}
section[data-testid="stSidebar"] > div{ padding-top:.4rem; }
.brand{ padding:.9rem .75rem 1.1rem; }
.brand-bar{ display:block; width:38px; height:5px; border-radius:3px; background:var(--accent); margin-bottom:.7rem; }
.brand-name{ display:block; font-size:1.85rem; font-weight:700; line-height:1; color:var(--text); }
.brand-sub{ display:block; color:var(--muted); font-size:.95rem; margin-top:.3rem; }
.side-title{ color:var(--muted); font-size:.9rem; font-weight:600; padding:1rem .75rem .3rem; }

section[data-testid="stSidebar"] .stButton > button,
section[data-testid="stSidebar"] .stDownloadButton > button{
  justify-content:flex-start; text-align:left; background:transparent;
  border:0; border-left:3px solid transparent; border-radius:6px;
  color:var(--muted); font-weight:500; font-size:1rem;
  padding:.45rem .75rem; box-shadow:none; min-height:2.3rem;
}
section[data-testid="stSidebar"] .stButton > button:hover,
section[data-testid="stSidebar"] .stDownloadButton > button:hover{
  background:var(--raised); color:var(--text); border-left-color:transparent;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"],
section[data-testid="stSidebar"] .stButton > button[data-testid="stBaseButton-primary"]{
  background:rgba(219,58,63,.16); color:#fff; border-left:3px solid var(--accent);
}
section[data-testid="stSidebar"] .stButton > button p{ text-align:left; }

.sync{ display:flex; align-items:center; gap:.55rem; color:var(--muted); font-size:.92rem; padding:.2rem .75rem .5rem; }
.sync i{ width:9px; height:9px; border-radius:50%; background:var(--line); flex:none; }
.sync.ok i{ background:var(--ok); }
.sync.dirty i{ background:var(--gold); }
.sync.dirty{ color:var(--text); }
.sync-erro{ color:var(--muted); font-size:.85rem; padding:0 .75rem .6rem; line-height:1.35; }

/* ---------- Botões ---------- */
.stApp .stButton > button,
.stApp .stDownloadButton > button,
.stApp [data-testid="stFormSubmitButton"] > button,
.stApp a[data-testid="stBaseLinkButton-secondary"]{
  border-radius:6px; border:1px solid var(--line); background:transparent;
  color:var(--text); font-weight:500; min-height:2.4rem; box-shadow:none;
  transition:background .15s, border-color .15s;
}
.stApp .stButton > button:hover,
.stApp .stDownloadButton > button:hover,
.stApp [data-testid="stFormSubmitButton"] > button:hover,
.stApp a[data-testid="stBaseLinkButton-secondary"]:hover{
  background:var(--raised); border-color:var(--muted); color:var(--text);
}
.stApp .stButton > button[kind="primary"],
.stApp .stButton > button[data-testid="stBaseButton-primary"],
.stApp [data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"],
.stApp [data-testid="stFormSubmitButton"] > button[data-testid="stBaseButton-primaryFormSubmit"]{
  background:var(--accent); border-color:var(--accent); color:#fff; font-weight:600;
}
.stApp .stButton > button[kind="primary"]:hover,
.stApp .stButton > button[data-testid="stBaseButton-primary"]:hover,
.stApp [data-testid="stFormSubmitButton"] > button[kind="primaryFormSubmit"]:hover,
.stApp [data-testid="stFormSubmitButton"] > button[data-testid="stBaseButton-primaryFormSubmit"]:hover{
  background:#C42F34; border-color:#C42F34; color:#fff;
}
.stApp button:focus-visible, .stApp a:focus-visible{ outline:2px solid var(--gold); outline-offset:2px; }

/* ---------- Campos, expansores, formulários ---------- */
[data-testid="stExpander"] details{
  border:1px solid var(--line); border-radius:8px; background:var(--surface);
}
[data-testid="stForm"]{ border:1px solid var(--line); border-radius:8px; background:transparent; }
[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"]{ border-radius:6px; }
[data-testid="stCaptionContainer"]{ color:var(--muted); }

/* ---------- Cartões de registro ---------- */
div[data-testid="stVerticalBlockBorderWrapper"]:has(.cm){
  background:var(--surface); border:1px solid var(--line); border-radius:8px;
}
div[data-testid="stVerticalBlockBorderWrapper"]:has(.hm){
  background:var(--surface); border:1px solid var(--line);
  border-top:4px solid var(--accent); border-radius:8px; padding:.35rem .35rem .15rem;
}
.cm, .hm{ display:none; }
.card-top{ display:flex; flex-wrap:wrap; align-items:center; gap:.4rem .75rem; margin-bottom:.35rem; }
.card-meta{ color:var(--muted); font-size:.95rem; }
.diff{ display:inline-flex; align-items:center; gap:.4rem; color:var(--muted); font-size:.92rem; }
.pips{ display:inline-flex; gap:3px; }
.pips i{ width:6px; height:12px; border-radius:2px; background:var(--line); }
.pips i.on{ background:var(--text); }
.card-date{ color:var(--muted); font-size:.92rem; margin-left:auto; }
.card-title{
  font-family:'Barlow Condensed','Barlow',sans-serif; font-size:1.7rem;
  font-weight:600; line-height:1.15; color:var(--text); margin:.1rem 0 .1rem;
}
.card-title .fav{ color:var(--gold); margin-right:.3rem; }
.card-src{ color:var(--muted); font-size:.92rem; margin:.2rem 0 .5rem; }

.pill{
  display:inline-block; padding:.1rem .6rem; border-radius:999px;
  font-size:.85rem; font-weight:600; color:#0F1A26; background:var(--c, var(--muted));
}
.pill.ok{ background:transparent; color:var(--ok); border:1px solid var(--ok); }

.acoes-marker{ display:none; }
[data-testid="stElementContainer"]:has(.acoes-marker),
.element-container:has(.acoes-marker){ display:none; }
@media (max-width:640px){
  [data-testid="stHorizontalBlock"]:has(.acoes-marker){ flex-wrap:nowrap !important; gap:.35rem !important; }
  [data-testid="stHorizontalBlock"]:has(.acoes-marker) > [data-testid="stColumn"],
  [data-testid="stHorizontalBlock"]:has(.acoes-marker) > [data-testid="column"]{
    min-width:0 !important; flex:1 1 0 !important; width:auto !important;
  }
  [data-testid="stHorizontalBlock"]:has(.acoes-marker) > [data-testid="stColumn"]:last-child,
  [data-testid="stHorizontalBlock"]:has(.acoes-marker) > [data-testid="column"]:last-child{ display:none; }
  [data-testid="stHorizontalBlock"]:has(.acoes-marker) button p{ font-size:.85rem; }
}

/* ---------- Visão geral ---------- */
.pre-item{ padding:.7rem 0; border-top:1px solid var(--line); }
.pre-item.first{ border-top:0; }
.pre-title{ font-family:'Barlow Condensed','Barlow',sans-serif; font-size:1.35rem; font-weight:600; line-height:1.15; margin:.35rem 0 .1rem; }
.pre-desc{ color:var(--muted); font-size:.98rem; line-height:1.4; }

.bar-row{ display:grid; grid-template-columns:minmax(110px,170px) 1fr 34px; gap:.7rem; align-items:center; margin:.4rem 0; }
.bar-label{ color:var(--text); font-size:.98rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.bar-track{ height:10px; background:var(--raised); border-radius:5px; overflow:hidden; }
.bar-fill{ display:block; height:100%; border-radius:5px; background:var(--c, var(--muted)); }
.bar-n{ text-align:right; color:var(--muted); font-size:.95rem; font-variant-numeric:tabular-nums; }

.week{ display:flex; gap:8px; height:150px; margin-top:.4rem; }
.week .col{ flex:1; display:flex; flex-direction:column; align-items:center; gap:4px; min-width:0; }
.week .n{ color:var(--muted); font-size:.85rem; height:1rem; font-variant-numeric:tabular-nums; }
.week .wrap{ flex:1; width:100%; display:flex; align-items:flex-end; }
.week .wrap b{ display:block; width:100%; min-height:3px; border-radius:3px 3px 0 0; background:#4A6580; }
.week .col.atual .wrap b{ background:var(--accent); }
.week .d{ color:var(--muted); font-size:.8rem; white-space:nowrap; }

.recente{ display:flex; flex-wrap:wrap; align-items:center; gap:.3rem .75rem; padding:.55rem 0; border-top:1px solid var(--line); }
.recente:first-child{ border-top:0; }
.recente .t{ font-weight:500; flex:1 1 220px; min-width:0; }
.recente .d{ color:var(--muted); font-size:.92rem; }

.vazio{ border:1px dashed var(--line); border-radius:8px; padding:1.6rem 1.4rem; color:var(--muted); line-height:1.5; }
.vazio b{ color:var(--text); font-weight:600; }
"""
    return "<style>" + base + por_tipo + "</style>"


# ============================================================
# BLOCOS DE HTML
# ============================================================

def html_pilula(tipo):
    classe = TIPOS_INFO[tipo]["css"]
    return f'<span class="pill {classe}">{escape(tipo)}</span>'


def html_dificuldade(dificuldade):
    nivel = {"Fácil": 1, "Médio": 2, "Difícil": 3}.get(dificuldade, 2)
    barras = "".join(
        f'<i class="{"on" if k < nivel else ""}"></i>' for k in range(3)
    )
    return f'<span class="diff"><span class="pips">{barras}</span>{escape(dificuldade)}</span>'


def cabecalho_pagina(titulo, subtitulo=""):
    st.markdown(
        f'<h1 class="page-title">{escape(titulo)}</h1>'
        f'<div class="page-sub">{escape(subtitulo)}</div>',
        unsafe_allow_html=True,
    )


def secao(titulo, subtitulo=""):
    sub = f'<div class="sec-sub">{escape(subtitulo)}</div>' if subtitulo else ""
    st.markdown(f'<h2 class="sec">{escape(titulo)}</h2>{sub}', unsafe_allow_html=True)


def estado_vazio(titulo, texto):
    st.markdown(
        f'<div class="vazio"><b>{escape(titulo)}</b><br>{escape(texto)}</div>',
        unsafe_allow_html=True,
    )


# ============================================================
# BARRA LATERAL
# ============================================================

def contar_para_menu(nome, tipo):
    df = st.session_state.df
    if tipo:
        return int((df["tipo"] == tipo).sum())
    if nome == "Caderno":
        return len(df)
    if nome == "Favoritos":
        return int(df["favorito"].sum())
    return None


def barra_lateral():
    with st.sidebar:
        st.markdown(
            '<div class="brand"><span class="brand-bar"></span>'
            '<span class="brand-name">Caderno tático</span>'
            '<span class="brand-sub">Seu diário de boxe</span></div>',
            unsafe_allow_html=True,
        )

        for nome, tipo in PAGINAS:
            quantidade = contar_para_menu(nome, tipo)
            rotulo = nome if quantidade is None else f"{nome} ({quantidade})"

            if st.button(
                rotulo,
                key=f"nav_{nome}",
                type="primary" if st.session_state.pagina == nome else "secondary",
                **LARGURA_BTN,
            ):
                st.session_state.pagina = nome
                st.session_state.editando_id = None
                st.session_state.excluindo_id = None
                st.rerun()

        st.markdown('<div class="side-title">Dados</div>', unsafe_allow_html=True)

        # Situação da sincronização
        if not github_configurado():
            chip = '<div class="sync off"><i></i>Nuvem não configurada</div>'
        elif st.session_state.estado_nuvem == "erro":
            chip = '<div class="sync dirty"><i></i>Sem conexão com a nuvem</div>'
        elif st.session_state.alterado:
            chip = '<div class="sync dirty"><i></i>Alterações não salvas</div>'
        elif st.session_state.ultimo_salvamento:
            chip = f'<div class="sync ok"><i></i>Salvo às {st.session_state.ultimo_salvamento}</div>'
        else:
            chip = '<div class="sync ok"><i></i>Sincronizado com a nuvem</div>'

        st.markdown(chip, unsafe_allow_html=True)

        if st.session_state.erro_conexao and st.session_state.estado_nuvem != "sem_config":
            st.markdown(
                f'<div class="sync-erro">{escape(st.session_state.erro_conexao)}</div>',
                unsafe_allow_html=True,
            )

        if st.button("Salvar na nuvem", key="btn_salvar", **LARGURA_BTN):
            salvar_na_nuvem()
            st.rerun()

        # Recarregar pede confirmação se houver alterações pendentes
        if st.button("Recarregar da nuvem", key="btn_recarregar", **LARGURA_BTN):
            if st.session_state.alterado and not st.session_state.confirmar_recarga:
                st.session_state.confirmar_recarga = True
                st.rerun()
            else:
                recarregar_da_nuvem()

        if st.session_state.confirmar_recarga:
            st.warning("Há alterações não salvas. Recarregar vai descartá-las.")
            c1, c2 = st.columns(2)
            if c1.button("Descartar", key="recarga_sim", type="primary", **LARGURA_BTN):
                recarregar_da_nuvem()
            if c2.button("Voltar", key="recarga_nao", **LARGURA_BTN):
                st.session_state.confirmar_recarga = False
                st.rerun()

        if github_configurado():
            st.toggle(
                "Salvar sozinho a cada mudança",
                key="autosave",
                help="Cada alteração gera um commit no repositório do GitHub.",
            )

        with st.expander("Importar e exportar"):
            importar_e_exportar()


def recarregar_da_nuvem():
    df_novo, erro, estado = carregar_github()
    st.session_state.confirmar_recarga = False

    if estado in ("ok", "novo"):
        st.session_state.df = preparar_dataframe(df_novo)
        st.session_state.alterado = False
        st.session_state.estado_nuvem = estado
        st.session_state.erro_conexao = ""
        st.session_state.editando_id = None
        st.session_state.excluindo_id = None
        avisar("Caderno recarregado da nuvem.", "🔄")
    else:
        st.session_state.erro_conexao = erro
        avisar(erro or "Não foi possível recarregar.", "⚠️")

    st.rerun()


def importar_e_exportar():
    st.caption("Importe uma planilha (.xlsx) ou um backup (.json).")

    arquivo = st.file_uploader(
        "Arquivo",
        type=["xlsx", "json"],
        label_visibility="collapsed",
        key=f"uploader_{st.session_state.uploader_versao}",
    )

    modo = st.radio(
        "Ao importar",
        ["Adicionar aos existentes", "Substituir tudo"],
        key="modo_importacao",
    )

    if st.button("Importar", key="btn_importar", disabled=arquivo is None, **LARGURA_BTN):
        try:
            df_importado = ler_arquivo_importado(arquivo)
            quantidade = len(df_importado)

            if modo == "Substituir tudo":
                df_final = df_importado
            else:
                df_final = pd.concat(
                    [st.session_state.df, df_importado], ignore_index=True
                ).drop_duplicates(subset="id", keep="first")

            confirmar_mudanca(df_final, f"{plural(quantidade, 'registro importado', 'registros importados')}.", "📥")
            st.session_state.uploader_versao += 1
            st.rerun()

        except ValueError as erro:
            st.error(str(erro))
        except Exception as erro:
            st.error(f"Não foi possível ler o arquivo: {erro}")

    st.divider()

    try:
        st.download_button(
            "Baixar planilha (.xlsx)",
            data=gerar_excel(st.session_state.df),
            file_name=f"caderno_tatico_{hoje():%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="baixar_excel",
            **LARGURA_BTN,
        )
        st.download_button(
            "Baixar backup (.json)",
            data=gerar_backup_json(st.session_state.df),
            file_name=f"caderno_tatico_{hoje():%Y-%m-%d}.json",
            mime="application/json",
            key="baixar_json",
            **LARGURA_BTN,
        )
    except Exception as erro:
        st.error(f"Erro ao gerar os arquivos: {erro}")


def ler_arquivo_importado(arquivo):
    nome = arquivo.name.lower()

    if nome.endswith(".json"):
        dados = json.loads(arquivo.getvalue().decode("utf-8"))
        df = json_para_dataframe(dados)
    else:
        df = pd.read_excel(arquivo)
        df.columns = [str(c).strip().lower() for c in df.columns]

        if "titulo" not in df.columns:
            raise ValueError("A planilha precisa ter uma coluna chamada 'titulo'.")

        df = preparar_dataframe(df)

    df = df[df["titulo"] != ""]

    if len(df) == 0:
        raise ValueError("Nenhum registro com título foi encontrado no arquivo.")

    return df.reset_index(drop=True)


# ============================================================
# FORMULÁRIOS
# ============================================================

def campos_do_registro(prefixo, base):
    """Desenha os campos dentro de um st.form e devolve os valores digitados."""
    c1, c2 = st.columns(2)

    with c1:
        tipo = st.selectbox(
            "Tipo", TIPOS, index=indice_de(TIPOS, base["tipo"]), key=f"{prefixo}_tipo"
        )
        titulo = st.text_input(
            "Título *",
            value=base["titulo"],
            placeholder="Ex.: Jab duplo seguido de direto",
            key=f"{prefixo}_titulo",
        )
        categorias = CATEGORIAS + (
            [base["categoria"]] if base["categoria"] not in CATEGORIAS else []
        )
        categoria = st.selectbox(
            "Categoria",
            categorias,
            index=indice_de(categorias, base["categoria"], len(categorias) - 1),
            key=f"{prefixo}_categoria",
        )
        dificuldade = st.selectbox(
            "Dificuldade",
            DIFICULDADES,
            index=indice_de(DIFICULDADES, base["dificuldade"], 1),
            key=f"{prefixo}_dificuldade",
        )

    with c2:
        data_registro = st.date_input(
            "Data",
            value=para_date(base["data"]) if base["data"] else hoje(),
            format="DD/MM/YYYY",
            key=f"{prefixo}_data",
        )
        status = st.selectbox(
            "Status",
            STATUS,
            index=indice_de(STATUS, base["status"]),
            key=f"{prefixo}_status",
        )
        fonte = st.text_input(
            "Fonte",
            value=base["fonte"],
            placeholder="Professor, vídeo, sparring com...",
            key=f"{prefixo}_fonte",
        )
        link_video = st.text_input(
            "Link do vídeo",
            value=base["link_video"],
            placeholder="https://",
            key=f"{prefixo}_link",
        )

    favorito = st.checkbox(
        "Marcar como favorito", value=bool(base["favorito"]), key=f"{prefixo}_favorito"
    )
    descricao = st.text_area(
        "Descrição *",
        value=base["descricao"],
        height=170,
        placeholder="O passo a passo, o que deu errado e como corrigir. Aceita Markdown, como listas com '-'.",
        key=f"{prefixo}_descricao",
    )

    return {
        "tipo": tipo,
        "titulo": titulo.strip(),
        "categoria": categoria,
        "descricao": descricao.strip(),
        "fonte": fonte.strip(),
        "link_video": normalizar_link(link_video),
        "dificuldade": dificuldade,
        "status": status,
        "favorito": favorito,
        "data": data_registro.strftime("%Y-%m-%d"),
    }


def validar_campos(campos):
    if not campos["titulo"]:
        return "Dê um título ao registro."
    if not campos["descricao"]:
        return "Escreva a descrição do registro."
    return ""


def formulario_novo(chave, tipo_padrao=None, expandido=False):
    base = {**VALORES_PADRAO, "tipo": tipo_padrao or "Dica", "data": ""}

    with st.expander("Novo registro", expanded=expandido):
        with st.form(f"form_novo_{chave}", clear_on_submit=True):
            campos = campos_do_registro(f"novo_{chave}", base)
            enviar = st.form_submit_button("Adicionar ao caderno", type="primary", **LARGURA_BTN)

        if enviar:
            erro = validar_campos(campos)
            if erro:
                st.error(erro)
            else:
                adicionar_registro(campos)
                st.rerun()


def formulario_edicao(registro, chave):
    with st.form(f"form_edicao_{chave}_{registro['id']}"):
        campos = campos_do_registro(f"ed_{chave}_{registro['id']}", registro)

        c1, c2 = st.columns(2)
        salvar = c1.form_submit_button("Salvar alterações", type="primary", **LARGURA_BTN)
        cancelar = c2.form_submit_button("Cancelar", **LARGURA_BTN)

    if cancelar:
        st.session_state.editando_id = None
        st.rerun()

    if salvar:
        erro = validar_campos(campos)
        if erro:
            st.error(erro)
        else:
            st.session_state.editando_id = None
            atualizar_registro(registro["id"], campos, "Registro atualizado.")
            st.rerun()


# ============================================================
# CARTÃO DE REGISTRO
# ============================================================

def cartao_registro(registro, chave):
    id_registro = registro["id"]
    tipo = registro["tipo"]
    info = TIPOS_INFO[tipo]

    with st.container(border=True):
        if st.session_state.editando_id == id_registro:
            st.markdown(
                f'<span class="cm {info["css"]}"></span>'
                f'<div class="card-title">Editando: {escape(registro["titulo"])}</div>',
                unsafe_allow_html=True,
            )
            formulario_edicao(registro, chave)
            return

        # Topo: tipo, categoria, dificuldade, status, data
        status_html = (
            '<span class="pill ok">Dominado</span>'
            if registro["status"] == "Dominado" and tipo not in TIPOS_SEM_STATUS
            else ""
        )
        estrela = '<span class="fav">★</span>' if registro["favorito"] else ""
        data_html = (
            f'<span class="card-date">{formatar_data(registro["data"])}</span>'
            if registro["data"]
            else ""
        )

        st.markdown(
            f'<span class="cm {info["css"]}"></span>'
            f'<div class="card-top">{html_pilula(tipo)}{status_html}'
            f'<span class="card-meta">{escape(registro["categoria"])}</span>'
            f'{html_dificuldade(registro["dificuldade"])}{data_html}</div>'
            f'<div class="card-title">{estrela}{escape(registro["titulo"])}</div>',
            unsafe_allow_html=True,
        )

        if registro["descricao"]:
            st.markdown(registro["descricao"])

        if registro["fonte"]:
            st.markdown(
                f'<div class="card-src">Fonte: {escape(registro["fonte"])}</div>',
                unsafe_allow_html=True,
            )

        # Confirmação de exclusão
        if st.session_state.excluindo_id == id_registro:
            st.warning("Excluir este registro? Essa ação não pode ser desfeita.")
            c1, c2, _ = st.columns([1, 1, 2])
            if c1.button("Excluir", key=f"{chave}_conf_{id_registro}", type="primary", **LARGURA_BTN):
                st.session_state.excluindo_id = None
                excluir_registro(id_registro)
                st.rerun()
            if c2.button("Cancelar", key=f"{chave}_canc_{id_registro}", **LARGURA_BTN):
                st.session_state.excluindo_id = None
                st.rerun()
            return

        acoes_do_cartao(registro, chave)


def acoes_do_cartao(registro, chave):
    id_registro = registro["id"]
    tem_status = registro["tipo"] not in TIPOS_SEM_STATUS
    tem_video = bool(registro["link_video"])
    dominado = registro["status"] == "Dominado"

    quantidade = 3 + int(tem_status) + int(tem_video)
    colunas = st.columns([1] * quantidade + [max(1, 6 - quantidade)])
    posicao = 0

    # Marcador invisível: permite manter os botões numa linha só no celular
    with colunas[0]:
        st.markdown('<span class="acoes-marker"></span>', unsafe_allow_html=True)

    if tem_video:
        with colunas[posicao]:
            st.link_button("Vídeo", registro["link_video"], **LARGURA_BTN)
        posicao += 1

    with colunas[posicao]:
        if st.button(
            "★" if registro["favorito"] else "☆",
            key=f"{chave}_fav_{id_registro}",
            help="Tirar dos favoritos" if registro["favorito"] else "Adicionar aos favoritos",
            **LARGURA_BTN,
        ):
            atualizar_registro(
                id_registro,
                {"favorito": not registro["favorito"]},
                "Removido dos favoritos." if registro["favorito"] else "Adicionado aos favoritos.",
                "⭐",
            )
            st.rerun()
    posicao += 1

    if tem_status:
        with colunas[posicao]:
            if st.button(
                "Rever" if dominado else "Dominei",
                key=f"{chave}_status_{id_registro}",
                help="Voltar para 'Praticando'" if dominado else "Marcar como dominado",
                **LARGURA_BTN,
            ):
                atualizar_registro(
                    id_registro,
                    {"status": "Praticando" if dominado else "Dominado"},
                    "Voltou para a lista de prática." if dominado else "Marcado como dominado.",
                    "🥊",
                )
                st.rerun()
        posicao += 1

    with colunas[posicao]:
        if st.button("Editar", key=f"{chave}_edit_{id_registro}", **LARGURA_BTN):
            st.session_state.editando_id = id_registro
            st.session_state.excluindo_id = None
            st.rerun()
    posicao += 1

    with colunas[posicao]:
        if st.button("Excluir", key=f"{chave}_del_{id_registro}", **LARGURA_BTN):
            st.session_state.excluindo_id = id_registro
            st.session_state.editando_id = None
            st.rerun()


# ============================================================
# LISTAGEM COM FILTROS
# ============================================================

ORDENACOES = [
    "Mais recentes",
    "Mais antigos",
    "Título (A a Z)",
    "Mais difíceis primeiro",
    "Favoritos primeiro",
]


def aplicar_filtros(df, busca, categoria, dificuldade, status, so_favoritos, tipo, ordem):
    resultado = df.assign(_pos=df.index)

    if tipo != "Todos":
        resultado = resultado[resultado["tipo"] == tipo]
    if categoria != "Todas":
        resultado = resultado[resultado["categoria"] == categoria]
    if dificuldade != "Todas":
        resultado = resultado[resultado["dificuldade"] == dificuldade]
    if status != "Todos":
        resultado = resultado[resultado["status"] == status]
    if so_favoritos:
        resultado = resultado[resultado["favorito"]]

    termos = normalizar_texto(busca).split()
    if termos and len(resultado):
        palheiro = (
            resultado["titulo"] + " " + resultado["descricao"] + " "
            + resultado["categoria"] + " " + resultado["fonte"] + " "
            + resultado["tipo"]
        ).map(normalizar_texto)
        mascara = pd.Series(True, index=resultado.index)
        for termo in termos:
            mascara &= palheiro.str.contains(termo, regex=False)
        resultado = resultado[mascara]

    peso = {"Fácil": 1, "Médio": 2, "Difícil": 3}

    if ordem == "Mais antigos":
        resultado = resultado.sort_values(["data", "_pos"])
    elif ordem == "Título (A a Z)":
        resultado = resultado.assign(_t=resultado["titulo"].map(normalizar_texto)).sort_values("_t")
    elif ordem == "Mais difíceis primeiro":
        resultado = resultado.assign(_d=resultado["dificuldade"].map(peso)).sort_values(
            ["_d", "_pos"], ascending=[False, False]
        )
    elif ordem == "Favoritos primeiro":
        resultado = resultado.sort_values(["favorito", "_pos"], ascending=[False, False])
    else:
        resultado = resultado.sort_values(["data", "_pos"], ascending=[False, False])

    return resultado


def lista_de_registros(df_base, chave, filtrar_tipo=False, vazio=None):
    """Mostra filtros e cartões para o DataFrame recebido."""
    if len(df_base) == 0:
        titulo, texto = vazio or ("Nada por aqui ainda.", "Adicione um registro no botão acima.")
        estado_vazio(titulo, texto)
        return

    busca = st.text_input(
        "Buscar",
        placeholder="Buscar por título, descrição, fonte ou categoria",
        label_visibility="collapsed",
        key=f"busca_{chave}",
    )

    colunas = st.columns([2, 2, 2, 2, 2] if filtrar_tipo else [2, 2, 2, 2])
    posicao = iter(colunas)

    tipo = "Todos"
    if filtrar_tipo:
        with next(posicao):
            tipo = st.selectbox("Tipo", ["Todos"] + TIPOS, key=f"tipo_{chave}")

    with next(posicao):
        categorias = ["Todas"] + sorted(c for c in df_base["categoria"].unique() if c)
        categoria = st.selectbox("Categoria", categorias, key=f"cat_{chave}")
    with next(posicao):
        dificuldade = st.selectbox("Dificuldade", ["Todas"] + DIFICULDADES, key=f"dif_{chave}")
    with next(posicao):
        status = st.selectbox("Status", ["Todos"] + STATUS, key=f"sta_{chave}")
    with next(posicao):
        ordem = st.selectbox("Ordenar por", ORDENACOES, key=f"ord_{chave}")

    so_favoritos = st.toggle("Só favoritos", key=f"fav_{chave}")

    resultado = aplicar_filtros(
        df_base, busca, categoria, dificuldade, status, so_favoritos, tipo, ordem
    )

    st.caption(plural(len(resultado), "registro", "registros"))

    if len(resultado) == 0:
        estado_vazio(
            "Nenhum registro com esses filtros.",
            "Limpe a busca ou escolha outras opções.",
        )
        return

    chave_limite = f"limite_{chave}"
    limite = st.session_state.get(chave_limite, ITENS_POR_PAGINA)

    for _, registro in resultado.head(limite).iterrows():
        cartao_registro(registro, chave)

    if len(resultado) > limite:
        restantes = len(resultado) - limite
        if st.button(
            f"Mostrar mais ({restantes})", key=f"mais_{chave}", **LARGURA_BTN
        ):
            st.session_state[chave_limite] = limite + ITENS_POR_PAGINA
            st.rerun()


# ============================================================
# PÁGINAS
# ============================================================

def subtitulo_dados(df):
    return (
        f"{plural(len(df), 'registro', 'registros')}, "
        f"{plural(int(df['favorito'].sum()), 'favorito', 'favoritos')}"
    )


def pagina_visao_geral():
    df = st.session_state.df
    agora_ = agora()
    data_extenso = (
        f"{DIAS_SEMANA[agora_.weekday()].capitalize()}, "
        f"{agora_.day} de {MESES[agora_.month - 1]}"
    )

    if len(df) == 0:
        cabecalho_pagina("Caderno tático de boxe", data_extenso)
        estado_vazio(
            "Seu caderno está vazio.",
            "Registre a primeira dica, combo ou correção do professor. "
            "Depois do primeiro sparring, anote os erros: eles aparecem aqui "
            "para você rever antes do próximo.",
        )
        formulario_novo("visao", expandido=True)
        return

    dominados = int(
        ((df["status"] == "Dominado") & (~df["tipo"].isin(TIPOS_SEM_STATUS))).sum()
    )
    cabecalho_pagina(
        "Caderno tático de boxe",
        f"{data_extenso}. {plural(len(df), 'registro', 'registros')}, "
        f"{plural(int(df['favorito'].sum()), 'favorito', 'favoritos')} e "
        f"{plural(dominados, 'item dominado', 'itens dominados')}.",
    )

    formulario_novo("visao")

    esquerda, direita = st.columns([3, 2], gap="large")

    with esquerda:
        bloco_antes_do_sparring(df)
        bloco_recentes(df)

    with direita:
        bloco_por_tipo(df)
        bloco_por_semana(df)
        bloco_por_categoria(df)


def bloco_antes_do_sparring(df):
    pendentes = df[df["tipo"].isin(TIPOS_REVISAO) & (df["status"] != "Dominado")]
    pendentes = pendentes.assign(_pos=pendentes.index).sort_values(
        ["data", "_pos"], ascending=[False, False]
    )

    secao(
        "Antes do sparring",
        "Erros e correções que você ainda não domina, do mais recente ao mais antigo.",
    )

    with st.container(border=True):
        st.markdown('<span class="hm"></span>', unsafe_allow_html=True)

        if len(pendentes) == 0:
            st.markdown(
                '<div class="pre-desc" style="padding:.6rem .4rem">'
                "Nada pendente. Depois do próximo sparring, registre o que deu errado "
                "como <b>Erro no Sparring</b> e ele aparece aqui.</div>",
                unsafe_allow_html=True,
            )
            return

        for numero, (_, item) in enumerate(pendentes.head(4).iterrows()):
            c1, c2 = st.columns([5, 1], vertical_alignment="center") if _suporta_alinhamento() else st.columns([5, 1])

            with c1:
                st.markdown(
                    f'<div class="pre-item{" first" if numero == 0 else ""}">{html_pilula(item["tipo"])}'
                    f'<div class="pre-title">{escape(item["titulo"])}</div>'
                    f'<div class="pre-desc">{escape(trecho(item["descricao"]))}</div></div>',
                    unsafe_allow_html=True,
                )
            with c2:
                if st.button("Dominei", key=f"ov_dom_{item['id']}", **LARGURA_BTN):
                    atualizar_registro(item["id"], {"status": "Dominado"}, "Marcado como dominado.", "🥊")
                    st.rerun()

        if len(pendentes) > 4:
            st.caption(f"Mais {len(pendentes) - 4} pendentes nas páginas de erros e correções.")


def _suporta_alinhamento():
    try:
        return "vertical_alignment" in inspect.signature(st.columns).parameters
    except Exception:
        return False


def bloco_recentes(df):
    secao("Últimos registros")

    recentes = df.assign(_pos=df.index).sort_values(
        ["data", "_pos"], ascending=[False, False]
    ).head(6)

    linhas = ""
    for _, r in recentes.iterrows():
        estrela = "★ " if r["favorito"] else ""
        linhas += (
            f'<div class="recente">{html_pilula(r["tipo"])}'
            f'<span class="t">{estrela}{escape(r["titulo"])}</span>'
            f'<span class="d">{formatar_data(r["data"])}</span></div>'
        )

    st.markdown(linhas, unsafe_allow_html=True)


def bloco_por_tipo(df):
    secao("Por tipo")

    contagem = df["tipo"].value_counts()
    maximo = max(int(contagem.max()), 1) if len(contagem) else 1

    linhas = ""
    for tipo, info in TIPOS_INFO.items():
        n = int(contagem.get(tipo, 0))
        largura = round(100 * n / maximo)
        linhas += (
            f'<div class="bar-row"><span class="bar-label">{escape(tipo)}</span>'
            f'<span class="bar-track"><span class="bar-fill {info["css"]}" style="width:{largura}%"></span></span>'
            f'<span class="bar-n">{n}</span></div>'
        )

    st.markdown(linhas, unsafe_allow_html=True)


def bloco_por_semana(df):
    secao("Ritmo de registros", "Últimas 8 semanas")

    hoje_ = hoje()
    inicio_atual = hoje_ - timedelta(days=hoje_.weekday())
    contagem = [0] * 8

    for iso in df["data"]:
        if not iso:
            continue
        dia = para_date(iso)
        inicio = dia - timedelta(days=dia.weekday())
        distancia = (inicio_atual - inicio).days // 7
        if 0 <= distancia < 8:
            contagem[7 - distancia] += 1

    maximo = max(max(contagem), 1)
    colunas = ""

    for posicao, n in enumerate(contagem):
        inicio = inicio_atual - timedelta(weeks=7 - posicao)
        altura = round(100 * n / maximo)
        atual = " atual" if posicao == 7 else ""
        colunas += (
            f'<div class="col{atual}"><span class="n">{n or ""}</span>'
            f'<div class="wrap"><b style="height:{altura}%"></b></div>'
            f'<span class="d">{inicio:%d/%m}</span></div>'
        )

    st.markdown(f'<div class="week">{colunas}</div>', unsafe_allow_html=True)


def bloco_por_categoria(df):
    secao("Categorias mais anotadas")

    contagem = df["categoria"].value_counts().head(5)
    maximo = max(int(contagem.max()), 1)

    linhas = ""
    for categoria, n in contagem.items():
        largura = round(100 * n / maximo)
        linhas += (
            f'<div class="bar-row"><span class="bar-label">{escape(categoria)}</span>'
            f'<span class="bar-track"><span class="bar-fill" style="width:{largura}%;--c:#98A9BB"></span></span>'
            f'<span class="bar-n">{int(n)}</span></div>'
        )

    st.markdown(linhas, unsafe_allow_html=True)


def pagina_caderno():
    df = st.session_state.df
    cabecalho_pagina("Caderno", subtitulo_dados(df))
    formulario_novo("caderno", expandido=len(df) == 0)
    lista_de_registros(
        df,
        "caderno",
        filtrar_tipo=True,
        vazio=("Seu caderno está vazio.", "Use o botão 'Novo registro' para anotar a primeira dica."),
    )


def pagina_tipo(nome, tipo):
    df = st.session_state.df
    df_tipo = df[df["tipo"] == tipo]

    cabecalho_pagina(nome, subtitulo_dados(df_tipo))
    formulario_novo(f"tipo_{tipo}", tipo_padrao=tipo, expandido=len(df_tipo) == 0)

    mensagens_vazias = {
        "Combo": "Anote as sequências que você quer automatizar, como 'jab, direto, gancho de esquerda'.",
        "Defesa": "Registre esquivas, bloqueios e contra-ataques que funcionam para você.",
        "Erro no Sparring": "Depois de cada sparring, anote o que deu errado. Os erros pendentes aparecem na Visão geral.",
        "Correção do Professor": "Anote cada correção que o professor fizer, enquanto ainda estiver fresca na memória.",
        "Treino": "Registre seus treinos para acompanhar a evolução.",
        "Dica": "Guarde dicas de técnica, estratégia e mentalidade.",
    }

    lista_de_registros(
        df_tipo,
        f"tipo_{tipo}",
        vazio=(f"Nenhum registro em {nome.lower()} ainda.", mensagens_vazias.get(tipo, "")),
    )


def pagina_favoritos():
    df = st.session_state.df
    favoritos = df[df["favorito"]]

    cabecalho_pagina("Favoritos", plural(len(favoritos), "registro marcado", "registros marcados"))
    lista_de_registros(
        favoritos,
        "favoritos",
        filtrar_tipo=True,
        vazio=(
            "Nenhum favorito ainda.",
            "Toque na estrela (☆) de qualquer registro para guardá-lo aqui.",
        ),
    )


# ============================================================
# EXECUÇÃO
# ============================================================

iniciar_estado()
st.markdown(montar_css(), unsafe_allow_html=True)
barra_lateral()
mostrar_avisos()

_pagina = st.session_state.pagina
_tipo_da_pagina = dict(PAGINAS).get(_pagina)

if _pagina == "Caderno":
    pagina_caderno()
elif _pagina == "Favoritos":
    pagina_favoritos()
elif _tipo_da_pagina:
    pagina_tipo(_pagina, _tipo_da_pagina)
else:
    pagina_visao_geral()
