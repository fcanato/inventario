import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import traceback
from datetime import datetime, date

try:
    import psycopg2
    import psycopg2.extras
    PSYCOPG2_OK = True
except ImportError:
    PSYCOPG2_OK = False

st.set_page_config(page_title="Contagem de Estoque", page_icon="📦", layout="wide")

# ── Estilos globais ────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Tipografia base ─────────────────────────────────── */
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif; }

/* ── Cartões de métricas ─────────────────────────────── */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, #f0f7ff 0%, #e8f4fd 100%);
    border-radius: 14px;
    padding: 16px 20px !important;
    border: 1px solid #cce0f5;
    box-shadow: 0 2px 8px rgba(21,101,192,0.08);
    transition: box-shadow 0.2s;
}
[data-testid="stMetric"]:hover {
    box-shadow: 0 4px 16px rgba(21,101,192,0.15);
}
[data-testid="stMetricValue"] {
    font-size: 2rem !important;
    font-weight: 900 !important;
    color: #1a237e !important;
    line-height: 1.15 !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.75rem !important;
    font-weight: 700 !important;
    color: #546e7a !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* ── Botões primários ────────────────────────────────── */
[data-testid="baseButton-primary"] {
    border-radius: 10px !important;
    font-weight: 700 !important;
    background: linear-gradient(135deg, #1565c0 0%, #1976d2 100%) !important;
    border: none !important;
    box-shadow: 0 3px 10px rgba(21,101,192,0.28) !important;
    letter-spacing: 0.3px;
    transition: all 0.2s ease !important;
}
[data-testid="baseButton-primary"]:hover {
    box-shadow: 0 5px 18px rgba(21,101,192,0.42) !important;
    transform: translateY(-1px) !important;
}

/* ── Botões secundários ──────────────────────────────── */
[data-testid="baseButton-secondary"] {
    border-radius: 10px !important;
    font-weight: 600 !important;
    transition: all 0.2s ease !important;
}

/* ── Expanders ───────────────────────────────────────── */
[data-testid="stExpander"] > details {
    border-radius: 12px !important;
    border: 1px solid #dee6f0 !important;
    overflow: hidden;
}
[data-testid="stExpander"] > details > summary {
    border-radius: 12px !important;
    font-weight: 600;
}

/* ── Alertas (success / warning / error / info) ─────── */
[data-testid="stAlert"] {
    border-radius: 12px !important;
    border-width: 0 0 0 5px !important;
}

/* ── Inputs de texto e número ────────────────────────── */
[data-testid="stTextInput"] > div > div > input,
[data-testid="stNumberInput"] > div > div > input {
    border-radius: 8px !important;
}
[data-testid="stSelectbox"] > div > div {
    border-radius: 8px !important;
}

/* ── Tabs – visual de "pill" ─────────────────────────── */
[data-baseweb="tab-list"] {
    background: #f0f4fa !important;
    border-radius: 12px !important;
    padding: 5px !important;
    gap: 3px !important;
}
[data-baseweb="tab"] {
    border-radius: 9px !important;
    font-weight: 600 !important;
    transition: background 0.18s !important;
}
[aria-selected="true"][data-baseweb="tab"] {
    background: #ffffff !important;
    box-shadow: 0 1px 5px rgba(0,0,0,0.12) !important;
    color: #1565c0 !important;
}

/* ── Container padrão (st.container border) ─────────── */
[data-testid="stVerticalBlock"] [data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 14px !important;
    border-color: #d0e4f7 !important;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}

/* ── Sidebar – fundo escuro premium ─────────────────── */
[data-testid="stSidebar"] > div:first-child {
    background: linear-gradient(180deg, #0d1b4b 0%, #1a237e 50%, #1565c0 100%) !important;
}
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown span,
[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #e8eaf6 !important;
}
[data-testid="stSidebar"] [data-testid="stMetric"] {
    background: rgba(255,255,255,0.12) !important;
    border-color: rgba(255,255,255,0.2) !important;
}
[data-testid="stSidebar"] [data-testid="stMetricValue"] {
    color: #ffffff !important;
}
[data-testid="stSidebar"] [data-testid="stMetricLabel"] {
    color: #90caf9 !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.2) !important; }

/* ── Dataframe / tabela ──────────────────────────────── */
[data-testid="stDataFrame"] > div {
    border-radius: 12px !important;
    overflow: hidden;
    border: 1px solid #e0e8f4 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Paths ──────────────────────────────────────────────────────────────────────
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.path.abspath(os.getcwd())

ESTOQUE_FILE = os.path.join(BASE_DIR, "estoque.xlsx")
DB_FILE      = os.path.join(BASE_DIR, "contagem.db")

# ── Detecção de modo: Supabase (PostgreSQL) ou SQLite local ───────────────────
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "") if hasattr(st, "secrets") else ""
USE_PG = bool(SUPABASE_URL and PSYCOPG2_OK)

# ══════════════════════════════════════════════════════════════════════════════
# CAMADA DE BANCO — funções que abstraem SQLite ↔ PostgreSQL
# ══════════════════════════════════════════════════════════════════════════════

def get_conn():
    if USE_PG:
        conn = psycopg2.connect(SUPABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
        return conn
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def _execute(conn, sql, params=()):
    """Executa SQL adaptando placeholders %s (PG) ↔ ? (SQLite)."""
    if USE_PG:
        sql = sql.replace("?", "%s")
    cur = conn.cursor()
    cur.execute(sql, params)
    return cur

def _fetchone(cur):
    row = cur.fetchone()
    if row is None:
        return None
    return dict(row)

def _fetchall(cur):
    rows = cur.fetchall()
    return [dict(r) for r in rows]

def init_db():
    conn = get_conn()
    if USE_PG:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS inventarios (
                id          SERIAL PRIMARY KEY,
                nome        TEXT NOT NULL,
                descricao   TEXT,
                data_inicio TEXT NOT NULL,
                data_fim    TEXT,
                status      TEXT DEFAULT 'Aberto'
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS contagens (
                id            SERIAL PRIMARY KEY,
                inventario_id INTEGER NOT NULL,
                id_estoque    TEXT,
                desc_estoque  TEXT,
                cod_produto   TEXT,
                desc_produto  TEXT,
                unid_medida   TEXT,
                qtd_sistema   REAL,
                qtd_contada   REAL,
                diferenca     REAL,
                ativo         TEXT,
                lote          TEXT,
                observacao    TEXT,
                operador      TEXT,
                data_hora     TEXT,
                FOREIGN KEY (inventario_id) REFERENCES inventarios(id)
            )
        """)
        # Migração: adiciona coluna lote em bancos existentes (PG suporta IF NOT EXISTS)
        cur.execute(
            "ALTER TABLE contagens ADD COLUMN IF NOT EXISTS lote TEXT DEFAULT ''"
        )
        conn.commit()
        cur.close()
        conn.close()
        return

    # ── SQLite ────────────────────────────────────────────────────────────────
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS inventarios (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nome        TEXT NOT NULL,
            descricao   TEXT,
            data_inicio TEXT NOT NULL,
            data_fim    TEXT,
            status      TEXT DEFAULT 'Aberto'
        );
        CREATE TABLE IF NOT EXISTS contagens (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            inventario_id INTEGER NOT NULL,
            id_estoque    TEXT,
            desc_estoque  TEXT,
            cod_produto   TEXT,
            desc_produto  TEXT,
            unid_medida   TEXT,
            qtd_sistema   REAL,
            qtd_contada   REAL,
            diferenca     REAL,
            ativo         TEXT,
            lote          TEXT,
            observacao    TEXT,
            operador      TEXT,
            data_hora     TEXT,
            FOREIGN KEY (inventario_id) REFERENCES inventarios(id)
        );
    """)
    conn.commit()
    cols = [r[1] for r in conn.execute("PRAGMA table_info(contagens)").fetchall()]
    if "inventario_id" not in cols:
        conn.executescript("""
            DROP TABLE IF EXISTS contagens;
            DROP TABLE IF EXISTS inventarios;
            CREATE TABLE inventarios (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                nome        TEXT NOT NULL,
                descricao   TEXT,
                data_inicio TEXT NOT NULL,
                data_fim    TEXT,
                status      TEXT DEFAULT 'Aberto'
            );
            CREATE TABLE contagens (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                inventario_id INTEGER NOT NULL,
                id_estoque    TEXT,
                desc_estoque  TEXT,
                cod_produto   TEXT,
                desc_produto  TEXT,
                unid_medida   TEXT,
                qtd_sistema   REAL,
                qtd_contada   REAL,
                diferenca     REAL,
                ativo         TEXT,
                lote          TEXT,
                observacao    TEXT,
                operador      TEXT,
                data_hora     TEXT,
                FOREIGN KEY (inventario_id) REFERENCES inventarios(id)
            );
        """)
        conn.commit()
    # Migração: adiciona coluna lote em bancos SQLite existentes
    cols_now = [r[1] for r in conn.execute("PRAGMA table_info(contagens)").fetchall()]
    if "lote" not in cols_now:
        conn.execute("ALTER TABLE contagens ADD COLUMN lote TEXT DEFAULT ''")
        conn.commit()
    conn.close()

# Inicializa banco — erros aparecem no frontend
try:
    init_db()
except Exception as e:
    st.error(f"❌ Erro ao inicializar banco de dados:\n\n{e}")
    st.code(traceback.format_exc())
    st.stop()

# ── Funções de Inventário ──────────────────────────────────────────────────────
def criar_inventario(nome, descricao):
    conn = get_conn()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if USE_PG:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO inventarios (nome, descricao, data_inicio, status) VALUES (%s,%s,%s,%s) RETURNING id",
            (nome, descricao, agora, "Aberto")
        )
        novo_id = cur.fetchone()["id"]
        conn.commit(); cur.close(); conn.close()
    else:
        cur = _execute(conn,
            "INSERT INTO inventarios (nome, descricao, data_inicio, status) VALUES (?,?,?,?)",
            (nome, descricao, agora, "Aberto")
        )
        conn.commit(); novo_id = cur.lastrowid; conn.close()
    return novo_id

def fechar_inventario(inv_id):
    conn = get_conn()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur = _execute(conn, "UPDATE inventarios SET status='Fechado', data_fim=? WHERE id=?", (agora, inv_id))
    conn.commit(); conn.close()

def reabrir_inventario(inv_id):
    conn = get_conn()
    _execute(conn, "UPDATE inventarios SET status='Aberto', data_fim=NULL WHERE id=?", (inv_id,))
    conn.commit(); conn.close()

def deletar_inventario(inv_id):
    conn = get_conn()
    _execute(conn, "DELETE FROM contagens WHERE inventario_id=?", (inv_id,))
    _execute(conn, "DELETE FROM inventarios WHERE id=?", (inv_id,))
    conn.commit(); conn.close()

def listar_inventarios():
    conn = get_conn()
    if USE_PG:
        cur = conn.cursor()
        cur.execute("SELECT * FROM inventarios ORDER BY id DESC")
        rows = _fetchall(cur)
        cur.close(); conn.close()
        return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["id","nome","descricao","data_inicio","data_fim","status"])
    df = pd.read_sql("SELECT * FROM inventarios ORDER BY id DESC", conn)
    conn.close()
    return df

# ── Funções de Contagem ────────────────────────────────────────────────────────
def salvar_contagem(inv_id, dados, operador):
    conn = get_conn()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lote = dados.get("lote", "") or ""
    cur = _execute(conn,
        "SELECT id FROM contagens WHERE inventario_id=? AND cod_produto=? AND id_estoque=? AND ativo=? AND lote=?",
        (inv_id, dados["cod_produto"], dados["id_estoque"], dados["ativo"], lote)
    )
    existe = _fetchone(cur)
    if existe:
        _execute(conn,
            "UPDATE contagens SET qtd_contada=?, diferenca=?, observacao=?, operador=?, data_hora=? WHERE id=?",
            (dados["qtd_contada"], dados["diferenca"], dados["observacao"], operador, agora, existe["id"])
        )
    else:
        _execute(conn,
            """INSERT INTO contagens
               (inventario_id,id_estoque,desc_estoque,cod_produto,desc_produto,
                unid_medida,qtd_sistema,qtd_contada,diferenca,ativo,lote,observacao,operador,data_hora)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (inv_id, dados["id_estoque"], dados["desc_estoque"], dados["cod_produto"],
             dados["desc_produto"], dados["unid_medida"], dados["qtd_sistema"],
             dados["qtd_contada"], dados["diferenca"], dados["ativo"],
             lote, dados["observacao"], operador, agora)
        )
    conn.commit(); conn.close()

def listar_contagens(inv_id):
    conn = get_conn()
    if USE_PG:
        cur = conn.cursor()
        cur.execute("SELECT * FROM contagens WHERE inventario_id=%s ORDER BY data_hora DESC", (inv_id,))
        rows = _fetchall(cur); cur.close(); conn.close()
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    df = pd.read_sql("SELECT * FROM contagens WHERE inventario_id=? ORDER BY data_hora DESC", conn, params=(inv_id,))
    conn.close()
    return df

def buscar_contagem_existente(inv_id, cod_produto, id_estoque, ativo, lote=""):
    conn = get_conn()
    cur = _execute(conn,
        "SELECT * FROM contagens WHERE inventario_id=? AND cod_produto=? AND id_estoque=? AND ativo=? AND lote=?",
        (inv_id, cod_produto, id_estoque, ativo, lote or "")
    )
    row = _fetchone(cur); conn.close()
    return row

def listar_contagens_historico(inv_ids):
    if not inv_ids:
        return pd.DataFrame()
    conn = get_conn()
    placeholders = ",".join(("%" + "s") * len(inv_ids) if USE_PG else ["?"] * len(inv_ids))
    sql = f"""SELECT i.nome AS Inventario, i.data_inicio AS DataInicio, i.status AS Status, c.*
              FROM contagens c JOIN inventarios i ON c.inventario_id = i.id
              WHERE c.inventario_id IN ({placeholders}) ORDER BY c.data_hora DESC"""
    if USE_PG:
        cur = conn.cursor(); cur.execute(sql, inv_ids)
        rows = _fetchall(cur); cur.close(); conn.close()
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    df = pd.read_sql(sql, conn, params=inv_ids); conn.close()
    return df

# ── Estoque Excel ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando base de estoque...")
def carregar_estoque():
    if not os.path.exists(ESTOQUE_FILE):
        return pd.DataFrame()
    df = pd.read_excel(ESTOQUE_FILE, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].fillna("").str.strip()
    return df

def buscar_produto(df, codigo, id_estoque=None):
    codigo = str(codigo).strip().upper()
    resultado = df[df["Cód. Produto"].str.upper() == codigo]
    if id_estoque and id_estoque != "Todos":
        filtrado = resultado[resultado["Id. Estoq. Físico"] == str(id_estoque).strip()]
        if not filtrado.empty:
            return filtrado
    return resultado

def interpretar_ativo(valor):
    """Retorna True se o produto está ativo. Vazio/NaN = Ativo por padrão."""
    v = str(valor).strip().upper()
    if v in ["", "NAN", "NONE"]:
        return True   # sem informação = considera ativo
    if v in ["N", "NAO", "NÃO", "INATIVO", "0", "FALSE", "F"]:
        return False
    return True       # S, SIM, ATIVO, 1, TRUE ou qualquer outro valor = ativo

# ── Carrega estoque — sem o arquivo, exibe aviso de upload ───────────────────
try:
    df_estoque = carregar_estoque()
    estoques_disponiveis = sorted(df_estoque["Id. Estoq. Físico"].unique().tolist()) if not df_estoque.empty else []
except Exception as e:
    st.error(f"❌ Erro ao carregar `estoque.xlsx`:\n\n{e}")
    st.code(traceback.format_exc())
    df_estoque = pd.DataFrame()
    estoques_disponiveis = []

# ── Session State ──────────────────────────────────────────────────────────────
for k, v in {
    "inv_id": None, "inv_nome": "", "operador": "", "produto": None,
    "ultimo_cod": "", "input_key": 0, "permitir_recontagem": False,
    "linha_selecionada": None,  # índice da linha escolhida quando há múltiplos patrimônios
    "pag_base4": 1,             # página atual da Aba 4
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("📦 Contagem de Estoque")
    st.divider()

    st.header("🗂️ Inventário Ativo")
    df_inv = listar_inventarios()
    abertos = df_inv[df_inv["status"] == "Aberto"] if not df_inv.empty else pd.DataFrame()

    if not abertos.empty:
        opcoes = {f"#{r['id']} – {r['nome']} ({r['data_inicio'][:10]})": r["id"]
                  for _, r in abertos.iterrows()}
        escolha = st.selectbox("Selecione o inventário", list(opcoes.keys()))
        st.session_state.inv_id   = opcoes[escolha]
        st.session_state.inv_nome = escolha
    else:
        st.info("Nenhum inventário aberto.\nCrie um abaixo ↓")
        st.session_state.inv_id = None

    with st.expander("➕ Novo Inventário"):
        nome_inv = st.text_input("Nome", placeholder="Ex: Inventário Mai/2026")
        desc_inv = st.text_input("Descrição (opcional)")
        if st.button("Criar", type="primary", use_container_width=True):
            if nome_inv.strip():
                nid = criar_inventario(nome_inv.strip(), desc_inv.strip())
                st.session_state.inv_id   = nid
                st.session_state.inv_nome = nome_inv.strip()
                st.success(f"Inventário #{nid} criado!")
                st.rerun()
            else:
                st.warning("Informe um nome.")

    if st.session_state.inv_id:
        if st.button("🔒 Fechar inventário atual", use_container_width=True):
            fechar_inventario(st.session_state.inv_id)
            st.session_state.inv_id = None
            st.rerun()

    st.divider()
    st.session_state.operador = st.text_input(
        "👤 Operador", value=st.session_state.operador, placeholder="Seu nome"
    )

    if st.session_state.inv_id:
        st.divider()
        df_cnt_side = listar_contagens(st.session_state.inv_id)
        _total_base = max(1, len(df_estoque))
        _contados   = len(df_cnt_side)
        _pendentes  = max(0, len(df_estoque) - _contados)
        _pct        = _contados / _total_base

        st.metric("📋 Itens na base",  len(df_estoque))
        st.metric("✅ Contados",        _contados)
        st.metric("⏳ Pendentes",       _pendentes)

        # Barra de progresso visual
        st.markdown(
            f"<div style='margin-top:10px'>"
            f"<div style='font-size:11px;color:#90caf9;font-weight:700;"
            f"text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px'>"
            f"Progresso da Contagem</div>"
            f"<div style='background:rgba(255,255,255,0.15);border-radius:8px;"
            f"height:10px;overflow:hidden'>"
            f"<div style='background:linear-gradient(90deg,#43a047,#66bb6a);"
            f"width:{_pct*100:.1f}%;height:100%;border-radius:8px;"
            f"transition:width 0.4s ease'></div></div>"
            f"<div style='font-size:13px;color:#e8eaf6;font-weight:700;margin-top:5px'>"
            f"{_pct*100:.1f}% concluído</div>"
            f"</div>",
            unsafe_allow_html=True,
        )

    st.divider()
    with st.expander("📂 Carregar / Atualizar Estoque"):
        uploaded = st.file_uploader(
            "Selecione o arquivo estoque.xlsx",
            type=["xlsx"],
            label_visibility="collapsed"
        )
        if uploaded is not None:
            with open(ESTOQUE_FILE, "wb") as f:
                f.write(uploaded.read())
            st.cache_data.clear()
            st.success("✅ Estoque atualizado com sucesso!")
            st.rerun()
        st.caption(f"Arquivo atual: `{os.path.basename(ESTOQUE_FILE)}`")

    if st.button("🔄 Recarregar estoque", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    if USE_PG:
        st.success("☁️ Banco: **Supabase** (nuvem)")
    else:
        st.info("💾 Banco: **SQLite** (local)")

# ══════════════════════════════════════════════════════════════════════════════
# TÍTULO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    """
    <div style="background:linear-gradient(135deg,#0d1b4b 0%,#1565c0 65%,#1976d2 100%);
                border-radius:18px;padding:22px 32px;margin-bottom:6px;
                box-shadow:0 6px 28px rgba(13,27,75,0.35)">
      <div style="display:flex;align-items:center;gap:18px">
        <span style="font-size:52px;filter:drop-shadow(0 2px 4px rgba(0,0,0,0.3))">📦</span>
        <div>
          <div style="font-size:24px;font-weight:900;color:#ffffff;
                      letter-spacing:-0.5px;line-height:1.2;
                      text-shadow:0 1px 4px rgba(0,0,0,0.3)">
            Contagem de Estoque Físico
          </div>
          <div style="font-size:13px;color:#90caf9;margin-top:5px;font-weight:500;
                      letter-spacing:0.3px">
            Tel Ribeirão Preto &nbsp;·&nbsp; Sistema de Inventário
          </div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state.inv_id:
    st.markdown(
        f"<div style='background:#e8f5e9;border-left:5px solid #43a047;"
        f"border-radius:10px;padding:10px 16px;margin-top:6px;font-size:15px'>"
        f"🟢 &nbsp;<b>Inventário ativo:</b> {st.session_state.inv_nome}</div>",
        unsafe_allow_html=True,
    )
else:
    st.warning("⚠️ Nenhum inventário ativo. Crie ou selecione um na barra lateral.")

# ══════════════════════════════════════════════════════════════════════════════
# ABAS
# ══════════════════════════════════════════════════════════════════════════════
aba1, aba2, aba3, aba4 = st.tabs([
    "🔍 Contar Item",
    "📊 Contagem Atual",
    "📁 Histórico de Inventários",
    "📋 Base de Estoque"
])

# ─── ABA 1: CONTAR ITEM ───────────────────────────────────────────────────────
with aba1:
    if not st.session_state.inv_id:
        st.info("Selecione ou crie um inventário na barra lateral para começar.")
    else:
        if df_estoque.empty:
            st.warning("⚠️ Nenhuma base de estoque carregada. Faça o upload do arquivo `estoque.xlsx` na barra lateral (📂 Carregar / Atualizar Estoque).")
        # ── Campo de busca ────────────────────────────────────────────────────
        col_cod, col_est = st.columns([3, 1])
        with col_cod:
            codigo = st.text_input(
                "📷 Código do Produto (etiqueta ou manual)",
                placeholder="Ex: TEFE2093Z",
                key=f"cod_input_{st.session_state.input_key}"
            )
        with col_est:
            est_filtro = st.selectbox("📍 Estoque Físico", ["Todos"] + estoques_disponiveis)

        buscar = st.button("🔎 Buscar", type="primary", use_container_width=True)
        cod = codigo.strip()

        # ── Leitor de câmera ──────────────────────────────────────────────────
        with st.expander("📷 Ler QR Code / Código de Barras"):
            _tab_qr, _tab_foto = st.tabs(["📷 QR Code ao vivo", "📸 Foto (todos os formatos)"])

            # ── Tab 1: scanner ao vivo via câmera ──────────────────────────────
            with _tab_qr:
                st.caption(
                    "Aponte a câmera para o QR Code — leitura automática, sem tirar foto.\n"
                    "⚠️ Requer conexão HTTPS (funciona no Streamlit Cloud e localhost)."
                )
                try:
                    from streamlit_qrcode_scanner import qrcode_scanner as _qr_scanner
                    _qr_result = _qr_scanner(key=f"qr_{st.session_state.input_key}")
                    if _qr_result:
                        _cod_cam = str(_qr_result).strip()
                        st.success(f"✅ Código lido: **{_cod_cam}**")
                        st.session_state.ultimo_cod          = _cod_cam
                        st.session_state.linha_selecionada   = None
                        st.session_state.permitir_recontagem = False
                        _res_cam = buscar_produto(df_estoque, _cod_cam, est_filtro)
                        st.session_state.produto = _res_cam if not _res_cam.empty else None
                        if _res_cam.empty:
                            st.error(f"❌ Produto **{_cod_cam}** não encontrado na base.")
                        st.rerun()
                except ImportError:
                    st.warning(
                        "Pacote não instalado. Execute:\n```\npip install streamlit-qrcode-scanner\n```"
                    )

            # ── Tab 2: foto → pyzbar (Code128, EAN, QR e outros) ───────────────
            with _tab_foto:
                st.caption(
                    "📱 **Celular:** Procurar arquivos → **Câmera** → fotografar a etiqueta\n"
                    "🖥️ **PC:** selecionar imagem salva — suporta QR Code e código de barras."
                )
                _img_arq = st.file_uploader(
                    "foto",
                    type=["jpg", "jpeg", "png", "webp", "bmp"],
                    key=f"cam_{st.session_state.input_key}",
                    label_visibility="collapsed",
                )
                if _img_arq is not None:
                    try:
                        from PIL import Image
                        from pyzbar import pyzbar as _pyzbar
                        _decoded = _pyzbar.decode(Image.open(_img_arq).convert("RGB"))
                        if _decoded:
                            _cod_cam = _decoded[0].data.decode("utf-8").strip()
                            st.success(f"✅ Código identificado: **{_cod_cam}**")
                            st.session_state.ultimo_cod          = _cod_cam
                            st.session_state.linha_selecionada   = None
                            st.session_state.permitir_recontagem = False
                            _res_cam = buscar_produto(df_estoque, _cod_cam, est_filtro)
                            st.session_state.produto = _res_cam if not _res_cam.empty else None
                            if _res_cam.empty:
                                st.error(f"❌ Produto **{_cod_cam}** não encontrado na base.")
                            st.rerun()
                        else:
                            st.warning(
                                "⚠️ Código não identificado. Dicas:\n"
                                "- Aproxime mais a câmera da etiqueta\n"
                                "- Melhore a iluminação\n"
                                "- Evite reflexos ou sombras"
                            )
                    except ImportError:
                        st.error("❌ Execute: `pip install pyzbar Pillow`")
                    except Exception as _e:
                        st.error(f"❌ Erro ao processar imagem: {_e}")

        # Nova busca → zera seleção anterior
        if cod and (cod != st.session_state.ultimo_cod or buscar):
            st.session_state.ultimo_cod       = cod
            st.session_state.linha_selecionada = None
            st.session_state.permitir_recontagem = False
            resultado = buscar_produto(df_estoque, cod, est_filtro)
            st.session_state.produto = resultado if not resultado.empty else None

        # ── Produto não encontrado ────────────────────────────────────────────
        if cod and st.session_state.ultimo_cod == cod and st.session_state.produto is None:
            st.error(f"❌ Produto **{cod}** não encontrado na base de estoque.")

        # ── Produto(s) encontrado(s) ──────────────────────────────────────────
        elif st.session_state.produto is not None:
            df_p  = st.session_state.produto
            total = len(df_p)

            # ════════════════════════════════════════════════════════════════
            # ETAPA A — múltiplos patrimônios: contagem em massa numa tela só
            # ════════════════════════════════════════════════════════════════
            if total > 1 and st.session_state.linha_selecionada is None:

                # Calcula totais para o banner
                qtd_total_sistema = 0.0
                for _, _r in df_p.iterrows():
                    try:
                        qtd_total_sistema += float(str(_r.get("Qtd Estoque", "0")).replace(",", "."))
                    except Exception:
                        pass
                contados_ate_agora = 0
                for _, _r in df_p.iterrows():
                    _pat = str(_r.get("Ativo", "")).strip()
                    _as  = _pat if _pat.upper() not in ["", "NAN", "NONE"] else ("Ativo" if interpretar_ativo(_pat) else "Inativo")
                    _lt  = str(_r.get("Lote", "")).strip()
                    _lt  = "" if _lt.upper() in ["NAN", "NONE"] else _lt
                    if buscar_contagem_existente(st.session_state.inv_id, _r["Cód. Produto"], _r["Id. Estoq. Físico"], _as, _lt):
                        contados_ate_agora += 1

                st.markdown(
                    f"<div style='background:#fff8e1;border-left:5px solid #f9a825;"
                    f"padding:14px 18px;border-radius:10px;margin-bottom:16px'>"
                    f"<b style='font-size:18px;color:#5d4037'>📦 {df_p.iloc[0]['Cód. Produto']}</b><br>"
                    f"<span style='color:#555;font-size:13px'>{df_p.iloc[0]['Desc. Produto']}</span><br>"
                    f"<div style='margin-top:10px;display:flex;gap:20px'>"
                    f"<span style='background:#fff3e0;border:1px solid #ffb74d;border-radius:8px;"
                    f"padding:6px 14px;font-size:13px;font-weight:700;color:#e65100'>"
                    f"🏷️ {total} patrimônios</span>"
                    f"<span style='background:#e8f5e9;border:1px solid #66bb6a;border-radius:8px;"
                    f"padding:6px 14px;font-size:13px;font-weight:700;color:#2e7d32'>"
                    f"✅ {contados_ate_agora}/{total} contados</span>"
                    f"<span style='background:#e3f2fd;border:1px solid #64b5f6;border-radius:8px;"
                    f"padding:6px 14px;font-size:13px;font-weight:700;color:#1565c0'>"
                    f"📊 Qtd Sistema Total: {qtd_total_sistema:.0f}</span>"
                    f"</div></div>",
                    unsafe_allow_html=True
                )

                # Form único com campo de quantidade por patrimônio
                with st.form(f"form_multi_{st.session_state.input_key}", clear_on_submit=True):
                    st.markdown("### Informe a quantidade para cada patrimônio:")
                    entradas = []  # lista de (idx, row, qtd_key, obs_key, ja_contado)

                    for idx in range(total):
                        row      = df_p.iloc[idx]
                        pat      = str(row.get("Ativo", "")).strip()
                        tem_pat  = pat.upper() not in ["", "NAN", "NONE"]
                        ativo_ok_r = interpretar_ativo(pat)
                        asalvo_r = pat if tem_pat else ("Ativo" if ativo_ok_r else "Inativo")
                        lote_r   = str(row.get("Lote", "")).strip()
                        lote_r   = "" if lote_r.upper() in ["NAN", "NONE"] else lote_r
                        ja_r = buscar_contagem_existente(
                            st.session_state.inv_id,
                            row["Cód. Produto"], row["Id. Estoq. Físico"], asalvo_r, lote_r
                        )
                        try:
                            qtd_sist_r = float(str(row.get("Qtd Estoque", "0")).replace(",", "."))
                        except Exception:
                            qtd_sist_r = 0.0

                        borda = "#2e7d32" if ja_r else "#1565c0"
                        bg    = "#f1f8e9" if ja_r else "#e3f2fd"
                        badge_txt = "✅ Já contado" if ja_r else "⏳ Pendente"
                        badge_bg  = "#c8e6c9" if ja_r else "#bbdefb"
                        badge_cor = "#2e7d32" if ja_r else "#1565c0"

                        # HTML do identificador (patrimônio e/ou lote)
                        _ident_html = ""
                        if tem_pat:
                            _ident_html += (
                                f"<div style='margin-top:6px;background:#fff9c4;border:1px solid #f9a825;"
                                f"border-radius:8px;padding:6px 12px;display:inline-block;margin-right:8px'>"
                                f"<span style='font-size:11px;color:#795548'>🏷️ PATRIMÔNIO: </span>"
                                f"<span style='font-size:20px;font-weight:900;color:#e65100'>{pat}</span></div>"
                            )
                        if lote_r:
                            _ident_html += (
                                f"<div style='margin-top:6px;background:#e8f5e9;border:1px solid #66bb6a;"
                                f"border-radius:8px;padding:6px 12px;display:inline-block'>"
                                f"<span style='font-size:11px;color:#2e7d32'>📦 LOTE: </span>"
                                f"<span style='font-size:20px;font-weight:900;color:#1b5e20'>{lote_r}</span></div>"
                            )
                        if not tem_pat and not lote_r:
                            _ident_html = (
                                "<div style='margin-top:6px;background:#f5f5f5;border-radius:8px;"
                                "padding:6px 12px;display:inline-block;color:#9e9e9e;font-size:14px'>"
                                "Sem patrimônio ou lote</div>"
                            )
                        st.markdown(
                            f"<div style='border:2px solid {borda};background:{bg};"
                            f"border-radius:12px;padding:12px 16px;margin-bottom:4px'>"
                            f"<div style='display:flex;justify-content:space-between;align-items:center'>"
                            f"<span style='font-weight:700;font-size:15px'>Item {idx+1}</span>"
                            f"<span style='background:{badge_bg};color:{badge_cor};padding:2px 10px;"
                            f"border-radius:12px;font-size:11px;font-weight:700'>{badge_txt}</span></div>"
                            + _ident_html
                            + f"<div style='font-size:12px;color:#555;margin-top:6px'>"
                            f"📍 {row['Id. Estoq. Físico']} | "
                            f"Qtd Sistema: <b>{qtd_sist_r:.0f}</b> {row.get('Unid. Medida','')}"
                            + (f" | ✅ Já contado: <b>{ja_r['qtd_contada']:.0f}</b> em {ja_r['data_hora'][:16]}" if ja_r else "")
                            + f"</div></div>",
                            unsafe_allow_html=True
                        )

                        qtd_default = float(ja_r["qtd_contada"]) if ja_r else 0.0
                        obs_default = ja_r["observacao"] if ja_r else ""

                        col_q, col_o = st.columns([1, 2])
                        with col_q:
                            qtd_val = st.number_input(
                                f"Qtd contada — Pat. {pat if tem_pat else idx+1}",
                                min_value=0.0, value=qtd_default, step=1.0,
                                key=f"qtd_multi_{idx}_{st.session_state.input_key}"
                            )
                        with col_o:
                            obs_val = st.text_input(
                                "Observação (opcional)",
                                value=obs_default,
                                key=f"obs_multi_{idx}_{st.session_state.input_key}"
                            )
                        entradas.append((idx, row, asalvo_r, lote_r, qtd_sist_r, qtd_val, obs_val))
                        st.divider()

                    submitted = st.form_submit_button(
                        f"✅ Confirmar contagem de todos os {total} patrimônios",
                        type="primary", use_container_width=True
                    )

                    if submitted:
                        salvos = 0
                        for idx, row, asalvo_r, lote_r, qtd_sist_r, qtd_val, obs_val in entradas:
                            diferenca = qtd_val - qtd_sist_r
                            dados = {
                                "id_estoque":   row["Id. Estoq. Físico"],
                                "desc_estoque": row["Desc. Estoque Físico"],
                                "cod_produto":  row["Cód. Produto"],
                                "desc_produto": row["Desc. Produto"],
                                "unid_medida":  row["Unid. Medida"],
                                "qtd_sistema":  qtd_sist_r,
                                "qtd_contada":  qtd_val,
                                "diferenca":    diferenca,
                                "ativo":        asalvo_r,
                                "lote":         lote_r,
                                "observacao":   obs_val,
                            }
                            salvar_contagem(st.session_state.inv_id, dados, st.session_state.operador)
                            salvos += 1

                        st.session_state.produto             = None
                        st.session_state.ultimo_cod          = ""
                        st.session_state.linha_selecionada   = None
                        st.session_state.permitir_recontagem = False
                        st.session_state.input_key          += 1
                        st.toast(f"✅ {salvos} patrimônio(s) registrado(s)!", icon="✅")
                        st.rerun()

                # Botão cancelar fora do form
                def _cancelar():
                    st.session_state.produto            = None
                    st.session_state.ultimo_cod         = ""
                    st.session_state.linha_selecionada  = None
                    st.session_state.permitir_recontagem = False
                    st.session_state.input_key         += 1

                st.button("🔄 Cancelar e buscar outro código",
                          on_click=_cancelar, use_container_width=True)

            # ════════════════════════════════════════════════════════════════
            # ETAPA B — linha definida (único resultado ou já escolhido)
            # ════════════════════════════════════════════════════════════════
            else:
                idx_linha = st.session_state.linha_selecionada if st.session_state.linha_selecionada is not None else 0
                linha = df_p.iloc[idx_linha]

                ativo_num        = str(linha.get("Ativo", "")).strip()
                ativo_tem_numero = ativo_num.upper() not in ["", "NAN", "NONE"]
                ativo_ok         = interpretar_ativo(ativo_num)
                badge            = "🟢 Ativo" if ativo_ok else "🔴 Inativo"
                ativo_salvo      = ativo_num if ativo_tem_numero else ("Ativo" if ativo_ok else "Inativo")
                lote_val         = str(linha.get("Lote", "")).strip()
                lote_val         = "" if lote_val.upper() in ["NAN", "NONE"] else lote_val

                # ── Card do produto ──────────────────────────────────────────
                with st.container(border=True):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Cód. Produto",   linha["Cód. Produto"])
                    c2.metric("Estoque Físico",  linha["Id. Estoq. Físico"])
                    c3.metric("Unid. Medida",    linha["Unid. Medida"])
                    c4.metric("Status",          badge)
                    st.markdown(f"**Descrição:** {linha['Desc. Produto']}")
                    st.markdown(f"**Local:** {linha['Desc. Estoque Físico']}")

                    if ativo_tem_numero:
                        st.markdown(
                            f"<div style='background:#fff9c4;border:1px solid #f9a825;"
                            f"border-radius:8px;padding:10px 16px;margin-top:8px'>"
                            f"<span style='font-size:12px;color:#795548'>🏷️ Nº PATRIMÔNIO / ATIVO</span><br>"
                            f"<span style='font-size:28px;font-weight:900;color:#e65100'>"
                            f"{ativo_num}</span></div>",
                            unsafe_allow_html=True
                        )
                    if lote_val:
                        st.markdown(
                            f"<div style='background:#e8f5e9;border:1px solid #66bb6a;"
                            f"border-radius:8px;padding:10px 16px;margin-top:8px'>"
                            f"<span style='font-size:12px;color:#2e7d32'>📦 LOTE</span><br>"
                            f"<span style='font-size:28px;font-weight:900;color:#1b5e20'>"
                            f"{lote_val}</span></div>",
                            unsafe_allow_html=True
                        )

                    qtd_sist_raw = linha.get("Qtd Estoque", "0")
                    try:
                        qtd_sist = float(str(qtd_sist_raw).replace(",", ".")) if qtd_sist_raw not in ["", "nan"] else 0.0
                    except Exception:
                        qtd_sist = 0.0
                    st.metric("Qtd Sistema", f"{qtd_sist:.0f}")

                # ── Botão voltar (apenas informativo, total=1 sempre aqui) ────

                # ── Trava: item já contado? ──────────────────────────────────
                ja_contado = buscar_contagem_existente(
                    st.session_state.inv_id,
                    linha["Cód. Produto"], linha["Id. Estoq. Físico"], ativo_salvo, lote_val
                )

                if ja_contado and not st.session_state.permitir_recontagem:
                    st.warning(
                        f"⚠️ **Item já contado neste inventário!**\n\n"
                        f"- 📦 Qtd contada: **{ja_contado['qtd_contada']:.0f}**\n"
                        f"- 📋 Qtd sistema: **{ja_contado['qtd_sistema']:.0f}**\n"
                        f"- ↕️ Diferença: **{ja_contado['diferenca']:+.0f}**\n"
                        f"- 👤 Operador: **{ja_contado['operador'] or '–'}**\n"
                        f"- 🕐 Data/hora: **{ja_contado['data_hora']}**"
                        + (f"\n- 🏷️ Patrimônio: **{ativo_num}**" if ativo_tem_numero else "")
                    )
                    ca1, ca2 = st.columns(2)
                    with ca1:
                        if st.button("✏️ Corrigir / recontar", type="primary", use_container_width=True):
                            st.session_state.permitir_recontagem = True
                            st.rerun()
                    with ca2:
                        if st.button("🔄 Buscar outro item", use_container_width=True):
                            st.session_state.produto            = None
                            st.session_state.ultimo_cod         = ""
                            st.session_state.linha_selecionada  = None
                            st.session_state.permitir_recontagem = False
                            st.session_state.input_key         += 1
                            st.rerun()

                else:
                    # ── Formulário de contagem ───────────────────────────────
                    if ja_contado:
                        st.info(
                            f"✏️ **Modo correção** — contagem anterior: "
                            f"**{ja_contado['qtd_contada']:.0f}** em {ja_contado['data_hora']}"
                        )

                    with st.form("form_contagem", clear_on_submit=True):
                        qtd_default = float(ja_contado["qtd_contada"]) if ja_contado else 0.0
                        qtd_contada = st.number_input(
                            "📦 Quantidade contada fisicamente",
                            min_value=0.0, value=qtd_default, step=1.0
                        )
                        obs_default = ja_contado["observacao"] if ja_contado else ""
                        obs = st.text_input("📝 Observação (opcional)", value=obs_default)

                        if st.form_submit_button("✅ Confirmar Contagem", type="primary",
                                                 use_container_width=True):
                            diferenca = qtd_contada - qtd_sist
                            dados = {
                                "id_estoque":   linha["Id. Estoq. Físico"],
                                "desc_estoque": linha["Desc. Estoque Físico"],
                                "cod_produto":  linha["Cód. Produto"],
                                "desc_produto": linha["Desc. Produto"],
                                "unid_medida":  linha["Unid. Medida"],
                                "qtd_sistema":  qtd_sist,
                                "qtd_contada":  qtd_contada,
                                "diferenca":    diferenca,
                                "ativo":        ativo_salvo,
                                "lote":         lote_val,
                                "observacao":   obs,
                            }
                            salvar_contagem(st.session_state.inv_id, dados, st.session_state.operador)

                            sinal   = "+" if diferenca >= 0 else ""
                            pat_reg = f" | Pat: {ativo_num}" if ativo_tem_numero else ""
                            if lote_val:
                                pat_reg += f" | Lote: {lote_val}"
                            acao    = "Corrigido" if ja_contado else "Registrado"

                            # Zera tudo para próxima leitura
                            st.session_state.produto             = None
                            st.session_state.ultimo_cod          = ""
                            st.session_state.linha_selecionada   = None
                            st.session_state.permitir_recontagem = False
                            st.session_state.input_key          += 1

                            _msg_toast = (
                                f"✅ {acao}: {linha['Cód. Produto']}{pat_reg} "
                                f"| Dif: {sinal}{diferenca:.0f}"
                            )
                            st.toast(_msg_toast, icon="✅")
                            st.rerun()
            # ── bloco legado removido ──

            if False:  # bloco legado desativado — lógica migrada para ETAPA A/B acima
                cod_exibido = linhas_lista[0]["Cód. Produto"]
                desc_exibida = linhas_lista[0]["Desc. Produto"]

                st.markdown("---")
                st.markdown(
                    f"<div style='background:#fff8e1;border-left:5px solid #f9a825;"
                    f"padding:12px 18px;border-radius:8px;margin-bottom:12px'>"
                    f"<span style='font-size:18px;font-weight:700;color:#5d4037'>"
                    f"📦 {cod_exibido}</span><br>"
                    f"<span style='color:#555;font-size:14px'>{desc_exibida}</span><br>"
                    f"<span style='color:#e65100;font-size:13px;font-weight:600'>"
                    f"⚠️ {len(linhas_lista)} patrimônios encontrados — selecione o correto abaixo</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )

                # Cards em grade: 2 por linha
                pares = [linhas_lista[i:i+2] for i in range(0, len(linhas_lista), 2)]
                for par in pares:
                    cols_cards = st.columns(len(par))
                    for col_c, (j, row) in zip(cols_cards, [(linhas_lista.index(r), r) for r in par]):
                        pat        = str(row.get("Ativo", "")).strip()
                        tem_pat    = pat not in ["", "nan", "NaN", "NAN", "None", "NONE"]
                        ativo_ok_r = interpretar_ativo(pat)

                        # Verifica se este patrimônio já foi contado
                        ativo_salvo_r = pat if tem_pat else ("Ativo" if ativo_ok_r else "Inativo")
                        ja_cnt_r = buscar_contagem_existente(
                            st.session_state.inv_id,
                            row["Cód. Produto"],
                            row["Id. Estoq. Físico"],
                            ativo_salvo_r
                        ) if st.session_state.inv_id else None

                        # Cor da borda: verde=contado, azul=pendente
                        borda_cor = "#2e7d32" if ja_cnt_r else "#1565c0"
                        bg_cor    = "#f1f8e9" if ja_cnt_r else "#e3f2fd"
                        status_html = (
                            "<span style='background:#c8e6c9;color:#1b5e20;padding:2px 10px;"
                            "border-radius:12px;font-size:12px;font-weight:700'>✅ Já contado</span>"
                            if ja_cnt_r else
                            "<span style='background:#bbdefb;color:#0d47a1;padding:2px 10px;"
                            "border-radius:12px;font-size:12px;font-weight:700'>⏳ Pendente</span>"
                        )
                        pat_html = (
                            f"<div style='margin:8px 0;background:#fff9c4;border:1px solid #f9a825;"
                            f"border-radius:8px;padding:6px 12px;text-align:center'>"
                            f"<span style='font-size:11px;color:#795548'>Nº PATRIMÔNIO</span><br>"
                            f"<span style='font-size:22px;font-weight:900;color:#e65100'>{pat}</span>"
                            f"</div>"
                            if tem_pat else
                            "<div style='margin:8px 0;background:#f5f5f5;border-radius:8px;"
                            "padding:6px 12px;text-align:center;color:#9e9e9e;font-size:13px'>"
                            "🏷️ Sem nº de patrimônio</div>"
                        )
                        qtd_sist_r = row.get("Qtd Estoque", "0")
                        qtd_cnt_r  = (
                            f"<br><span style='font-size:12px;color:#2e7d32'>"
                            f"✅ Contado: <b>{ja_cnt_r['qtd_contada']:.0f}</b> "
                            f"({ja_cnt_r['data_hora'][:16]})</span>"
                            if ja_cnt_r else ""
                        )

                        with col_c:
                            st.markdown(
                                f"<div style='border:2px solid {borda_cor};background:{bg_cor};"
                                f"border-radius:12px;padding:14px 16px;margin-bottom:8px'>"
                                f"<div style='display:flex;justify-content:space-between;"
                                f"align-items:center;margin-bottom:6px'>"
                                f"<span style='font-weight:700;font-size:15px;color:#333'>"
                                f"Opção {j+1}</span>{status_html}</div>"
                                f"<div style='font-size:13px;color:#444;margin-bottom:4px'>"
                                f"📍 <b>{row['Id. Estoq. Físico']}</b> — {row['Desc. Estoque Físico']}</div>"
                                f"{pat_html}"
                                f"<div style='font-size:12px;color:#666;margin-top:4px'>"
                                f"Qtd Sistema: <b>{qtd_sist_r}</b> | "
                                f"Unid: <b>{row.get('Unid. Medida','–')}</b> | "
                                f"{'🟢 Ativo' if ativo_ok_r else '🔴 Inativo'}"
                                f"{qtd_cnt_r}</div>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                            btn_label = "✏️ Selecionar e corrigir" if ja_cnt_r else "✅ Selecionar este"
                            btn_type  = "secondary" if ja_cnt_r else "primary"
                            if st.button(
                                btn_label,
                                key=f"sel_{j}_{st.session_state.input_key}",
                                use_container_width=True,
                                type=btn_type
                            ):
                                st.session_state.produto = df_p.iloc[[j]]
                                st.session_state.permitir_recontagem = bool(ja_cnt_r)
                                st.rerun()

                st.button(
                    "🔄 Cancelar e buscar outro código",
                    on_click=lambda: st.session_state.update({
                        "produto": None, "ultimo_cod": "",
                        "permitir_recontagem": False,
                        "input_key": st.session_state.input_key + 1
                    }),
                    use_container_width=True
                )
                st.stop()  # Aguarda seleção — não renderiza formulário ainda

# ─── ABA 2: CONTAGEM ATUAL ────────────────────────────────────────────────────
with aba2:
    if not st.session_state.inv_id:
        st.info("Selecione um inventário na barra lateral.")
    else:
        st.subheader(f"Inventário: {st.session_state.inv_nome}")
        df_cur = listar_contagens(st.session_state.inv_id)

        if df_cur.empty:
            st.info("Nenhum item contado ainda neste inventário.")
        else:
            _div_count = int((df_cur["diferenca"].astype(float) != 0).sum())
            _qtd_cnt   = df_cur["qtd_contada"].astype(float).sum()
            _qtd_sis   = df_cur["qtd_sistema"].astype(float).sum()
            _diff_tot  = _qtd_cnt - _qtd_sis

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📦 Itens contados",    len(df_cur))
            m2.metric("⚠️ Com divergência",    _div_count)
            m3.metric("🔢 Qtd total contada",  f"{_qtd_cnt:.0f}")
            m4.metric("📋 Qtd total sistema",  f"{_qtd_sis:.0f}",
                      delta=f"{_diff_tot:+.0f}", delta_color="inverse")

            if _div_count > 0:
                _pct_div = _div_count / max(1, len(df_cur)) * 100
                st.markdown(
                    f"<div style='background:#fff3e0;border-left:5px solid #fb8c00;"
                    f"border-radius:12px;padding:12px 18px;margin:10px 0;"
                    f"display:flex;align-items:center;gap:16px'>"
                    f"<span style='font-size:28px'>⚠️</span>"
                    f"<div>"
                    f"<div style='font-weight:700;font-size:15px;color:#e65100'>"
                    f"{_div_count} ite{'m' if _div_count==1 else 'ns'} com divergência "
                    f"({_pct_div:.0f}% do total)</div>"
                    f"<div style='font-size:13px;color:#bf360c;margin-top:2px'>"
                    f"Diferença acumulada: <b>{_diff_tot:+.0f}</b> unidades</div>"
                    f"</div></div>",
                    unsafe_allow_html=True,
                )

            flt = st.selectbox("Filtrar estoque", ["Todos"] + sorted(df_cur["id_estoque"].unique().tolist()))
            df_v = df_cur if flt == "Todos" else df_cur[df_cur["id_estoque"] == flt]

            st.dataframe(df_v, use_container_width=True, hide_index=True)

            c1, c2 = st.columns(2)
            with c1:
                csv = df_v.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
                st.download_button("⬇️ Exportar CSV", csv, "contagem_atual.csv",
                                   "text/csv", use_container_width=True)
            with c2:
                buf = io.BytesIO()
                df_v.to_excel(buf, index=False, engine="openpyxl")
                st.download_button("⬇️ Exportar Excel", buf.getvalue(), "contagem_atual.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)

# ─── ABA 3: HISTÓRICO DE INVENTÁRIOS ─────────────────────────────────────────
with aba3:
    st.subheader("📁 Histórico de Inventários")
    df_todos = listar_inventarios()

    if df_todos.empty:
        st.info("Nenhum inventário criado ainda.")
    else:
        for _, inv in df_todos.iterrows():
            conn_h = get_conn()
            cur_h = _execute(conn_h, "SELECT COUNT(*) AS cnt FROM contagens WHERE inventario_id=?", (inv["id"],))
            qtd = _fetchone(cur_h)["cnt"]
            conn_h.close()

            icone = "🟢" if inv["status"] == "Aberto" else "🔴"
            with st.expander(f"{icone} #{inv['id']} – {inv['nome']} | {inv['data_inicio'][:10]} | {qtd} itens"):

                col1, col2, col3 = st.columns(3)
                col1.write(f"**Status:** {inv['status']}")
                col2.write(f"**Início:** {inv['data_inicio'][:16]}")
                col3.write(f"**Fim:** {inv['data_fim'][:16] if pd.notna(inv['data_fim']) else '–'}")
                if inv["descricao"]:
                    st.caption(inv["descricao"])

                df_hist = listar_contagens(inv["id"])
                if not df_hist.empty:
                    st.dataframe(df_hist, use_container_width=True, hide_index=True)

                    b1, b2, b3, b4 = st.columns(4)
                    with b1:
                        buf_h = io.BytesIO()
                        df_hist.to_excel(buf_h, index=False, engine="openpyxl")
                        st.download_button(
                            "⬇️ Exportar Excel", buf_h.getvalue(),
                            f"inventario_{inv['id']}_{inv['nome'].replace(' ','_')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"exp_{inv['id']}", use_container_width=True
                        )
                    with b2:
                        csv_h = df_hist.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
                        st.download_button(
                            "⬇️ Exportar CSV", csv_h,
                            f"inventario_{inv['id']}.csv", "text/csv",
                            key=f"csv_{inv['id']}", use_container_width=True
                        )
                    with b3:
                        if inv["status"] == "Aberto":
                            if st.button("🔒 Fechar", key=f"fch_{inv['id']}", use_container_width=True):
                                fechar_inventario(inv["id"])
                                st.rerun()
                        else:
                            if st.button("🔓 Reabrir", key=f"rab_{inv['id']}", use_container_width=True):
                                reabrir_inventario(inv["id"])
                                st.rerun()
                    with b4:
                        if st.button("🗑️ Excluir", key=f"del_{inv['id']}", use_container_width=True):
                            deletar_inventario(inv["id"])
                            if st.session_state.inv_id == inv["id"]:
                                st.session_state.inv_id = None
                            st.rerun()
                else:
                    st.info("Nenhum item contado neste inventário.")
                    if st.button("🗑️ Excluir inventário vazio", key=f"del2_{inv['id']}"):
                        deletar_inventario(inv["id"])
                        st.rerun()

# ─── ABA 4: BASE DE ESTOQUE ───────────────────────────────────────────────────
with aba4:
    st.subheader("📋 Base de Estoque")
    f_est = st.selectbox("Filtrar por Estoque Físico", ["Todos"] + estoques_disponiveis, key="f_est4")
    pesq  = st.text_input("🔎 Pesquisar (código ou descrição)", key="pesq4")

    df_v4 = df_estoque.copy()
    if f_est != "Todos":
        df_v4 = df_v4[df_v4["Id. Estoq. Físico"] == f_est]
    if pesq:
        mask = (
            df_v4["Cód. Produto"].str.contains(pesq, case=False, na=False) |
            df_v4["Desc. Produto"].str.contains(pesq, case=False, na=False)
        )
        df_v4 = df_v4[mask]

    # ── Coluna de status vinculada à linha (segura para ordenação) ───────────
    if st.session_state.inv_id and not df_v4.empty:
        df_cnt4 = listar_contagens(st.session_state.inv_id)
        if not df_cnt4.empty:
            _lotes4 = df_cnt4["lote"].astype(str).str.strip() if "lote" in df_cnt4.columns else [""] * len(df_cnt4)
            contados_set = set(
                zip(
                    df_cnt4["cod_produto"].astype(str).str.strip(),
                    df_cnt4["id_estoque"].astype(str).str.strip(),
                    df_cnt4["ativo"].astype(str).str.strip(),
                    _lotes4,
                )
            )
        else:
            contados_set = set()

        def _status_linha(row):
            pat = str(row.get("Ativo", "")).strip()
            tem_pat = pat.upper() not in ["", "NAN", "NONE"]
            ativo_ok_r = interpretar_ativo(pat)
            ativo_salvo_r = pat if tem_pat else ("Ativo" if ativo_ok_r else "Inativo")
            lote_r = str(row.get("Lote", "")).strip()
            lote_r = "" if lote_r.upper() in ["NAN", "NONE"] else lote_r
            chave = (
                str(row["Cód. Produto"]).strip(),
                str(row["Id. Estoq. Físico"]).strip(),
                ativo_salvo_r,
                lote_r,
            )
            return "✅ Contado" if chave in contados_set else "⏳ Pendente"

        df_exib = df_v4.copy().reset_index(drop=True)
        df_exib.insert(0, "Status", df_exib.apply(_status_linha, axis=1))
        df_final = df_exib
    else:
        df_final = df_v4.copy().reset_index(drop=True)

    # ── Paginação ─────────────────────────────────────────────────────────────
    total_itens = len(df_final)
    pc1, pc2 = st.columns([2, 1])
    with pc1:
        st.caption(f"**{total_itens}** itens encontrados | Arquivo: {os.path.basename(ESTOQUE_FILE)}")
    with pc2:
        por_pagina = st.select_slider(
            "Itens por página", options=[25, 50, 100, 200], value=50, key="por_pag4"
        )

    total_pags = max(1, (total_itens + por_pagina - 1) // por_pagina)

    # Reseta para página 1 quando filtro ou pesquisa mudam
    _sig4 = f"{f_est}|{pesq}"
    if st.session_state.get("_sig_aba4") != _sig4:
        st.session_state["_sig_aba4"] = _sig4
        st.session_state.pag_base4 = 1
    # Garante que a página não ultrapasse o total (ex: após filtro mais restrito)
    st.session_state.pag_base4 = min(st.session_state.pag_base4, total_pags)

    pn1, pn2, pn3 = st.columns([1, 4, 1])
    with pn1:
        if st.button("◀ Anterior", disabled=st.session_state.pag_base4 <= 1,
                     key="btn_p4_prev", use_container_width=True):
            st.session_state.pag_base4 -= 1
            st.rerun()
    with pn2:
        st.markdown(
            f"<div style='text-align:center;padding:6px 0'>"
            f"Página <b>{st.session_state.pag_base4}</b> de <b>{total_pags}</b></div>",
            unsafe_allow_html=True
        )
    with pn3:
        if st.button("Próxima ▶", disabled=st.session_state.pag_base4 >= total_pags,
                     key="btn_p4_next", use_container_width=True):
            st.session_state.pag_base4 += 1
            st.rerun()

    _ini = (st.session_state.pag_base4 - 1) * por_pagina
    st.dataframe(df_final.iloc[_ini:_ini + por_pagina], use_container_width=True, hide_index=True)
