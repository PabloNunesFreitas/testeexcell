#!/usr/bin/env python3
"""Interface web — Sistema de Admissão FIBRA. Execute: streamlit run app.py"""

import os
import sys
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))

from solicitacao_emprego import (
    GeminiKeyManager, DOCUMENT_TYPES, REQUIRED_DOCUMENTS,
    REQUIRED_FIELDS, OPTIONAL_FIELDS, DAILY_LIMIT, STOP_AT_REMAINING,
    extract_document_info, merge_extracted_data, fill_excel,
)

TEMPLATE_PATH = Path(__file__).parent / "template_solicitacao.xlsx"

DOCUMENT_GROUPS = [
    ("🪪", "Identidade", [
        ("rg_frente",   "RG — Frente",   True),
        ("rg_verso",    "RG — Verso",    False),
        ("cpf_frente",  "CPF — Frente",  True),
        ("cpf_verso",   "CPF — Verso",   False),
        ("cnh",         "CNH",           False),
    ]),
    ("💼", "Trabalho", [
        ("carteira_trabalho",         "Carteira de Trabalho (CTPS)", True),
        ("carteira_trabalho_digital", "Carteira de Trabalho Digital", False),
    ]),
    ("🏠", "Residência", [
        ("comprovante_residencia", "Comprovante de Residência", True),
    ]),
    ("🗳️", "Eleitorais & Militares", [
        ("titulo_eleitor",         "Título de Eleitor",          True),
        ("certificado_reservista", "Certificado de Reservista",  False),
    ]),
    ("📚", "Formação & Currículo", [
        ("curriculo",             "Currículo",               True),
        ("historico_escolar",     "Histórico Escolar",       False),
        ("certificado_conclusao", "Certificado de Conclusão", False),
    ]),
    ("📸", "Outros", [
        ("foto_3x4",             "Foto 3x4",                 True),
        ("pis_doc",              "Documento PIS/NIT",         False),
        ("cartao_vacina_frente", "Cartão de Vacina — Frente", False),
        ("cartao_vacina_verso",  "Cartão de Vacina — Verso",  False),
    ]),
]


# ════════════════════════════════════════════════════════════
# CSS
# ════════════════════════════════════════════════════════════

def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

    .stApp { background: #eef2f7; }
    .main .block-container { padding: 1.5rem 2rem 3rem; max-width: 1200px; }
    #MainMenu, footer { visibility: hidden; }
    [data-testid="stDeployButton"] { display:none; }

    /* SIDEBAR */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg,#0b1929 0%,#112240 55%,#0d3060 100%) !important;
        border-right: 1px solid #1e3a5f;
    }
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] label { color: #ccd6f6 !important; }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: #e6f1ff !important; font-weight:700 !important; }
    [data-testid="stSidebar"] .stTextInput input {
        background: rgba(255,255,255,.08) !important;
        border: 1px solid rgba(100,163,255,.3) !important;
        color: #e6f1ff !important; border-radius: 8px !important;
    }
    [data-testid="stSidebar"] .stTextInput input::placeholder { color: #8892b0 !important; }
    [data-testid="stSidebar"] .stButton button {
        background: linear-gradient(135deg,#1d6fde,#4ca3ff) !important;
        color: white !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
    }
    [data-testid="stSidebar"] hr { border-color: rgba(100,163,255,.2) !important; }
    [data-testid="stSidebar"] .stProgress > div > div > div > div {
        background: linear-gradient(90deg,#1d6fde,#4ca3ff) !important;
    }
    [data-testid="stSidebar"] .stProgress > div > div {
        background: rgba(255,255,255,.1) !important; border-radius: 99px !important;
    }

    /* EXPANDER ADMIN — tema escuro dentro da sidebar */
    [data-testid="stSidebar"] [data-testid="stExpander"] {
        background: rgba(255,255,255,.04) !important;
        border: 1px solid rgba(100,163,255,.2) !important;
        box-shadow: none !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary {
        background: rgba(255,255,255,.07) !important;
        border-bottom: 1px solid rgba(100,163,255,.15) !important;
        color: #e6f1ff !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] summary:hover {
        background: rgba(255,255,255,.11) !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] label {
        color: #ccd6f6 !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] p,
    [data-testid="stSidebar"] [data-testid="stExpander"] span {
        color: #8892b0 !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] .stTextInput input {
        background: rgba(10,25,47,.85) !important;
        color: #e6f1ff !important;
        border: 1px solid rgba(100,163,255,.4) !important;
    }
    [data-testid="stSidebar"] [data-testid="stExpander"] .stTextInput input::placeholder {
        color: #4a6080 !important;
    }

    /* ABAS */
    .stTabs [data-testid="stTab"] {
        font-size: .98rem !important; font-weight: 500 !important;
        padding: .55rem 1.3rem !important;
    }
    .stTabs [aria-selected="true"] {
        font-weight: 700 !important; color: #1d6fde !important;
        border-bottom: 3px solid #1d6fde !important;
    }

    /* EXPANDERS */
    [data-testid="stExpander"] {
        background: white !important;
        border: 1px solid #dce6f5 !important; border-radius: 14px !important;
        box-shadow: 0 2px 10px rgba(13,48,96,.07) !important;
        margin-bottom: 1rem !important; overflow: hidden !important;
    }
    [data-testid="stExpander"] summary {
        padding: .9rem 1.4rem !important;
        font-size: 1rem !important; font-weight: 600 !important;
        color: #0b1929 !important; background: #f7faff !important;
        border-bottom: 1px solid #dce6f5 !important;
    }
    [data-testid="stExpander"] summary:hover { background: #eef4ff !important; }
    [data-testid="stExpander"] > div > div { padding: 1.2rem 1.4rem !important; }

    /* FILE UPLOADER */
    [data-testid="stFileUploaderDropzone"] {
        background: #f7faff !important;
        border: 2px dashed #b8d0f5 !important; border-radius: 10px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #1d6fde !important; background: #eef4ff !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span { color: #4a6fa5 !important; }

    /* BOTÃO PRIMÁRIO */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg,#0d3060,#1d6fde) !important;
        color: white !important; border: none !important;
        border-radius: 10px !important; font-size: 1.05rem !important;
        font-weight: 700 !important; height: 3.8rem !important;
        box-shadow: 0 4px 16px rgba(29,111,222,.4) !important;
        transition: all .2s !important;
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 20px rgba(29,111,222,.5) !important;
    }
    .stButton > button[kind="primary"]:disabled {
        background: #b0bec5 !important; box-shadow: none !important;
    }

    /* BOTÃO DOWNLOAD */
    .stDownloadButton > button {
        background: linear-gradient(135deg,#00695c,#00897b) !important;
        color: white !important; border: none !important;
        border-radius: 12px !important; font-size: 1.15rem !important;
        font-weight: 800 !important; height: 4.5rem !important;
        box-shadow: 0 6px 20px rgba(0,137,123,.4) !important;
        transition: all .2s !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 24px rgba(0,137,123,.5) !important;
    }

    /* MÉTRICAS */
    [data-testid="stMetric"] {
        background: white !important; padding: 1.25rem 1.5rem !important;
        border-radius: 12px !important; border: 1px solid #dce6f5 !important;
        box-shadow: 0 2px 8px rgba(13,48,96,.06) !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: .82rem !important; color: #607d9b !important; font-weight: 500 !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 2rem !important; font-weight: 800 !important; color: #0b1929 !important;
    }

    /* ALERTAS */
    [data-testid="stAlert"] {
        border-radius: 10px !important;
        border-left-width: 4px !important;
    }

    /* PROGRESS */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg,#1d6fde,#4ca3ff) !important;
        border-radius: 99px !important;
    }
    .stProgress > div > div { border-radius: 99px !important; }

    /* TABELA */
    [data-testid="stTable"] table {
        border-radius: 10px !important; overflow: hidden !important;
        box-shadow: 0 2px 8px rgba(13,48,96,.06) !important;
        border: 1px solid #dce6f5 !important;
    }
    [data-testid="stTable"] th {
        background: #0d3060 !important; color: white !important;
        font-weight: 600 !important; padding: .75rem 1rem !important;
    }
    [data-testid="stTable"] tr:nth-child(even) { background: #f7faff !important; }
    [data-testid="stTable"] td {
        padding: .6rem 1rem !important; color: #1e293b !important;
    }

    hr { border-color: #dce6f5 !important; margin: 1.5rem 0 !important; }
    </style>
    """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════

@st.cache_resource
def get_key_manager() -> GeminiKeyManager:
    """Singleton compartilhado entre todas as sessões."""
    secret_keys = []
    try:
        for k in ["GEMINI_API_KEY", "GEMINI_API_KEY_2", "GEMINI_API_KEY_3"]:
            v = st.secrets.get(k)
            if v:
                secret_keys.append(v)
    except Exception:
        pass
    return GeminiKeyManager(env_keys=secret_keys)


def _check_admin_credentials(username: str, password: str) -> bool:
    try:
        ok_user = st.secrets.get("ADMIN_USERNAME", "admin")
        ok_pass = st.secrets.get("ADMIN_PASSWORD", "fibra2024")
    except Exception:
        ok_user, ok_pass = "admin", "fibra2024"
    return username == ok_user and password == ok_pass


def render_sidebar() -> GeminiKeyManager:
    km = get_key_manager()

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    with st.sidebar:
        # ── Logo ──────────────────────────────────────────
        st.html("""
        <div style="padding:.5rem 0 1.5rem;text-align:center;">
            <div style="font-size:2.8rem;margin-bottom:.4rem;">🏢</div>
            <div style="font-size:1.35rem;font-weight:800;color:#e6f1ff;letter-spacing:.5px;">
                FIBRA
            </div>
            <div style="font-size:.72rem;color:#8892b0;letter-spacing:1.5px;text-transform:uppercase;
                        margin-top:.2rem;">
                Sistema de Admissão
            </div>
        </div>
        """)

        # ── Status visível para todos ──────────────────────
        entry     = km._current_entry()
        active    = km.get_active_key()
        used      = entry.get("requests_today", 0) if entry else 0
        remaining = (DAILY_LIMIT - used) if entry else 0
        dot_color = "#22c55e" if active else "#ef4444"
        status    = "Sistema pronto" if active else "Limite atingido hoje"

        st.html(f"""
        <div style="background:rgba(255,255,255,.06);border-radius:12px;
                    padding:.9rem 1.1rem;margin-bottom:.5rem;">
            <div style="display:flex;align-items:center;gap:.6rem;margin-bottom:.5rem;">
                <div style="width:9px;height:9px;border-radius:50%;
                            background:{dot_color};flex-shrink:0;"></div>
                <span style="color:#e6f1ff;font-size:.88rem;font-weight:700;">{status}</span>
            </div>
            <div style="font-size:.76rem;color:#8892b0;line-height:1.9;">
                🤖 Google Gemini AI<br>
                📊 {used:,} / {DAILY_LIMIT:,} usos hoje<br>
                ✅ {remaining:,} restantes
            </div>
        </div>
        """)
        if km.has_keys():
            st.progress(min(used / DAILY_LIMIT, 1.0))

        # ── Aviso se sem chaves ────────────────────────────
        if not km.has_keys() and not st.session_state.admin_logged_in:
            st.html("""
            <div style="margin:.8rem 0;padding:.8rem 1rem;
                        background:rgba(255,193,7,.12);border:1px solid rgba(255,193,7,.4);
                        border-radius:10px;">
                <p style="color:#ffc107;font-size:.82rem;margin:0;font-weight:600;">
                    ⚠️ Nenhuma chave configurada.<br>Faça login como admin.
                </p>
            </div>
            """)

        # ── Botão Admin (discreto, no rodapé da sidebar) ───
        st.divider()

        if not st.session_state.admin_logged_in:
            with st.expander("⚙️ Admin", expanded=False):
                st.html("""
                <p style="font-size:.8rem;color:#8892b0;margin-bottom:.5rem;">
                    Área restrita — acesso autorizado apenas.
                </p>
                """)
                u = st.text_input("Usuário", key="admin_user", placeholder="usuário")
                p = st.text_input("Senha", key="admin_pass", type="password", placeholder="••••••")
                if st.button("Entrar", use_container_width=True, type="primary", key="btn_login"):
                    if _check_admin_credentials(u, p):
                        st.session_state.admin_logged_in = True
                        st.rerun()
                    else:
                        st.error("Usuário ou senha incorretos.")
        else:
            # ── Painel Admin ───────────────────────────────
            st.html("""
            <p style="font-size:.9rem;font-weight:700;color:#4ca3ff;margin:.2rem 0 .8rem;">
                ⚙️ Painel Admin
            </p>
            """)

            # Lista de chaves com botão remover
            if km.data["keys"]:
                current_idx = km.data.get("current_index", 0)
                for i, entry_k in enumerate(km.data["keys"]):
                    rem  = DAILY_LIMIT - entry_k.get("requests_today", 0)
                    used_k = entry_k.get("requests_today", 0)
                    dot  = "🟢" if (i == current_idx and rem > STOP_AT_REMAINING) else ("🔴" if rem <= STOP_AT_REMAINING else "⚪")
                    col_a, col_b = st.columns([4, 1])
                    with col_a:
                        st.html(f"""
                        <p style="font-size:.8rem;margin:.3rem 0 0;color:#ccd6f6;font-weight:600;">
                            {dot} Conta {i+1}
                            <code style="font-size:.7rem;background:rgba(255,255,255,.1);
                                         padding:1px 5px;border-radius:4px;color:#a8d8ff;">
                                ...{entry_k['key'][-8:]}
                            </code>
                        </p>
                        <p style="font-size:.72rem;color:#8892b0;margin:.1rem 0 .4rem;">
                            {used_k:,}/{DAILY_LIMIT:,} · {rem:,} restantes
                        </p>
                        """)
                    with col_b:
                        if st.button("🗑️", key=f"del_{i}", help="Remover"):
                            km.remove_key(i)
                            st.rerun()
            else:
                st.info("Nenhuma chave cadastrada.")

            # Adicionar nova chave
            st.html("""
            <p style="font-size:.83rem;font-weight:700;color:#e6f1ff;
                      margin:.8rem 0 .3rem;">➕ Adicionar chave</p>
            <p style="font-size:.75rem;color:#8892b0;margin:0 0 .4rem;">
                <a href="https://aistudio.google.com/app/apikey" target="_blank"
                   style="color:#4ca3ff;text-decoration:none;">
                    Obter chave gratuita ↗
                </a>
            </p>
            """)
            new_key = st.text_input("Nova chave", type="password",
                                    placeholder="AIzaSy...", label_visibility="collapsed",
                                    key="new_key_input")
            if st.button("Adicionar", use_container_width=True, type="primary", key="btn_add_key"):
                if new_key.strip():
                    if km.add_key(new_key.strip()):
                        st.success(f"✅ Chave ...{new_key.strip()[-8:]} adicionada!")
                        st.rerun()
                    else:
                        st.info("Chave já cadastrada.")
                else:
                    st.error("Cole uma chave válida.")

            st.divider()
            if st.button("🚪 Sair do admin", use_container_width=True, key="btn_logout"):
                st.session_state.admin_logged_in = False
                st.rerun()

        if not km.has_keys():
            st.stop()

    return km


# ════════════════════════════════════════════════════════════
# CABEÇALHO
# ════════════════════════════════════════════════════════════

def render_header():
    now = datetime.now()
    st.html(f"""
    <div style="background:linear-gradient(135deg,#0b1929 0%,#0d3060 50%,#1d6fde 100%);
                padding:2.2rem 2.5rem;border-radius:18px;margin-bottom:1.8rem;
                box-shadow:0 8px 32px rgba(13,48,96,.25);
                display:flex;align-items:center;gap:1.5rem;">

        <div style="background:rgba(255,255,255,.12);border-radius:16px;
                    width:64px;height:64px;display:flex;align-items:center;
                    justify-content:center;font-size:2rem;flex-shrink:0;">
            📋
        </div>

        <div style="flex:1;">
            <h1 style="margin:0;color:white;font-size:1.75rem;font-weight:800;
                       letter-spacing:-.3px;font-family:'Inter',sans-serif;">
                Solicitação de Emprego
            </h1>
            <p style="margin:.35rem 0 0;color:rgba(255,255,255,.65);
                      font-size:.95rem;font-family:'Inter',sans-serif;">
                Envie os documentos do candidato — o Excel é preenchido automaticamente com
                <strong style="color:#64b5f6;">Gemini AI</strong>
            </p>
        </div>

        <div style="text-align:right;color:rgba(255,255,255,.5);
                    font-size:.78rem;line-height:2;flex-shrink:0;
                    font-family:'Inter',sans-serif;">
            <div>📅 {now.strftime('%d/%m/%Y')}</div>
            <div>🤖 Google Gemini</div>
            <div style="color:rgba(255,255,255,.35);">gratuito · 1.500 req/dia</div>
        </div>
    </div>
    """)


# ════════════════════════════════════════════════════════════
# ABA UPLOAD
# ════════════════════════════════════════════════════════════

def render_upload_tab() -> dict:
    uploaded: dict = {}

    for icon, group_name, docs in DOCUMENT_GROUPS:
        recv  = sum(1 for dt, _, _ in docs if st.session_state.get(f"up_{dt}"))
        total = len(docs)

        with st.expander(f"{icon}  **{group_name}**  —  {recv}/{total} arquivos", expanded=True):
            col_a, col_b = st.columns(2)
            for idx, (doc_type, label, required) in enumerate(docs):
                with (col_a if idx % 2 == 0 else col_b):
                    badge = (
                        '<span style="background:#fef3c7;color:#92400e;font-size:.68rem;'
                        'font-weight:700;padding:2px 7px;border-radius:99px;margin-right:5px;">'
                        'OBRIGATÓRIO</span>'
                    ) if required else ""
                    st.html(f"""
                    <p style="font-size:.87rem;font-weight:600;color:#1e3a5f;
                               margin-bottom:.25rem;font-family:'Inter',sans-serif;">
                        {badge}{label}
                    </p>
                    """)
                    uf = st.file_uploader(
                        label,
                        type=["jpg", "jpeg", "png", "pdf", "webp"],
                        key=f"up_{doc_type}",
                        label_visibility="collapsed",
                    )
                    if uf:
                        uploaded[doc_type] = uf
                        st.html(f"""
                        <p style="font-size:.8rem;color:#059669;font-weight:600;
                                   margin-top:.15rem;font-family:'Inter',sans-serif;">
                            ✅ {uf.name}
                        </p>
                        """)
    return uploaded


# ════════════════════════════════════════════════════════════
# PAINEL DE RESUMO
# ════════════════════════════════════════════════════════════

def render_summary(uploaded: dict):
    total_req = len(REQUIRED_DOCUMENTS)
    recv_req  = sum(1 for d in REQUIRED_DOCUMENTS if d in uploaded)
    recv_opt  = len(uploaded) - recv_req

    st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("📎 Enviados", len(uploaded))
    c2.metric("⭐ Obrigatórios", f"{recv_req}/{total_req}",
              delta="completo!" if recv_req == total_req else f"faltam {total_req - recv_req}",
              delta_color="normal" if recv_req == total_req else "inverse")
    c3.metric("📄 Opcionais", recv_opt)

    missing = [DOCUMENT_TYPES[d] for d in REQUIRED_DOCUMENTS if d not in uploaded]
    if missing:
        items_html = "".join(f"<li>{d}</li>" for d in missing)
        st.html(f"""
        <div style="background:#fffbeb;border:1px solid #fbbf24;
                    border-left:4px solid #f59e0b;border-radius:10px;
                    padding:1rem 1.25rem;margin-top:.75rem;">
            <p style="margin:0 0 .5rem;font-weight:700;color:#92400e;
                      font-family:'Inter',sans-serif;font-size:.92rem;">
                ⚠️ Documentos obrigatórios faltando:
            </p>
            <ul style="margin:0;padding-left:1.2rem;color:#78350f;
                       font-family:'Inter',sans-serif;font-size:.88rem;
                       line-height:1.8;">
                {items_html}
            </ul>
        </div>
        """)


# ════════════════════════════════════════════════════════════
# PROCESSAMENTO
# ════════════════════════════════════════════════════════════

def process_documents(uploaded: dict, km: GeminiKeyManager):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_docs: dict = {}
        for doc_type, uf in uploaded.items():
            suffix = Path(uf.name).suffix or ".jpg"
            p = os.path.join(tmpdir, f"{doc_type}{suffix}")
            with open(p, "wb") as f:
                f.write(uf.getvalue())
            tmp_docs[doc_type] = p

        total       = len(tmp_docs)
        prog        = st.progress(0)
        status_slot = st.empty()
        all_extractions: dict = {}

        for i, (doc_type, file_path) in enumerate(tmp_docs.items()):
            name = DOCUMENT_TYPES.get(doc_type, doc_type)
            prog.progress(i / total)
            status_slot.html(f"""
            <div style="background:white;border-radius:10px;padding:.85rem 1.2rem;
                        border:1px solid #dce6f5;font-size:.93rem;color:#1e3a5f;
                        font-family:'Inter',sans-serif;">
                🔍 <b>Lendo</b> {name}…
                <span style="color:#8892b0;">({i+1}/{total})</span>
            </div>
            """)

            if km.get_active_key() is None:
                status_slot.error("⚠️ Todas as chaves atingiram o limite. Adicione nova conta.")
                prog.empty()
                return

            try:
                extraction = extract_document_info(km, doc_type, file_path)
            except Exception as exc:
                status_slot.error(f"❌ Erro ao processar {name}: {exc}")
                prog.empty()
                return
            if extraction:
                all_extractions[doc_type] = extraction

        prog.progress(1.0)
        status_slot.html("""
        <div style="background:#f0fdf4;border-radius:10px;padding:.85rem 1.2rem;
                    border:1px solid #86efac;font-size:.93rem;color:#14532d;
                    font-family:'Inter',sans-serif;">
            📊 <b>Preenchendo o Excel…</b>
        </div>
        """)

        merged     = merge_extracted_data(all_extractions)
        nome       = str(merged.get("nome_completo", "CANDIDATO")).upper()
        nome_safe  = nome.replace(" ", "_").replace("/", "-")[:50]
        excel_name = f"SOLICITACAO_{nome_safe}.xlsx"
        excel_path = os.path.join(tmpdir, excel_name)

        missing_fields = fill_excel(str(TEMPLATE_PATH), excel_path, merged)

        with open(excel_path, "rb") as f:
            excel_bytes = f.read()

        prog.empty()
        status_slot.empty()

        st.session_state["results"] = {
            "nome":           nome,
            "merged":         merged,
            "missing_fields": missing_fields,
            "found_docs":     list(uploaded.keys()),
            "excel_bytes":    excel_bytes,
            "excel_name":     excel_name,
        }


# ════════════════════════════════════════════════════════════
# ABA RESULTADO
# ════════════════════════════════════════════════════════════

def render_result_tab():
    r = st.session_state.get("results")

    if r is None:
        st.html("""
        <div style="text-align:center;padding:4rem 2rem;background:white;
                    border-radius:16px;border:2px dashed #dce6f5;margin-top:1rem;">
            <div style="font-size:3.5rem;margin-bottom:1rem;">📊</div>
            <h3 style="color:#1e3a5f;margin:0 0 .5rem;font-family:'Inter',sans-serif;">
                Nenhum resultado ainda
            </h3>
            <p style="color:#607d9b;margin:0;font-size:.95rem;font-family:'Inter',sans-serif;">
                Vá para a aba <b>📤 Enviar Documentos</b>,
                envie os arquivos e clique em <b>Processar</b>.
            </p>
        </div>
        """)
        return

    # Card do candidato
    total_req = len(REQUIRED_FIELDS)
    filled    = total_req - len(r["missing_fields"])
    pct       = filled / total_req
    bar_color = "#22c55e" if pct == 1 else ("#f59e0b" if pct >= 0.7 else "#ef4444")
    bar_w     = f"{pct * 100:.0f}%"

    st.html(f"""
    <div style="background:white;border-radius:16px;padding:1.75rem 2rem;
                border:1px solid #dce6f5;box-shadow:0 2px 12px rgba(13,48,96,.08);
                margin-bottom:1.25rem;display:flex;align-items:center;gap:1.2rem;">
        <div style="width:56px;height:56px;border-radius:14px;
                    background:linear-gradient(135deg,#0d3060,#1d6fde);
                    display:flex;align-items:center;justify-content:center;
                    font-size:1.6rem;flex-shrink:0;">👤</div>
        <div style="flex:1;">
            <div style="font-size:.75rem;color:#8892b0;font-weight:500;
                        text-transform:uppercase;letter-spacing:.6px;
                        font-family:'Inter',sans-serif;">Candidato</div>
            <div style="font-size:1.4rem;font-weight:800;color:#0b1929;
                        margin:.15rem 0 .6rem;font-family:'Inter',sans-serif;">
                {r['nome']}
            </div>
            <div style="background:#eef2f7;border-radius:99px;height:7px;overflow:hidden;">
                <div style="background:{bar_color};width:{bar_w};height:100%;
                            border-radius:99px;"></div>
            </div>
            <div style="font-size:.77rem;color:#607d9b;margin-top:.3rem;
                        font-family:'Inter',sans-serif;">
                {filled}/{total_req} campos obrigatórios preenchidos
            </div>
        </div>
    </div>
    """)

    # Métricas
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Preenchidos",  f"{filled}/{total_req}")
    c2.metric("⚠️ Faltando",    len(r["missing_fields"]),
              delta=None if r["missing_fields"] else "tudo ok!",
              delta_color="inverse")
    c3.metric("📄 Documentos",  len(r["found_docs"]))

    st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)

    # Botão download
    st.download_button(
        label="⬇️   Baixar Excel Preenchido",
        data=r["excel_bytes"],
        file_name=r["excel_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)

    # Campos faltando
    if r["missing_fields"]:
        items_html = "".join(
            f"<li style='margin:.35rem 0;'>{label}</li>"
            for _, label in r["missing_fields"]
        )
        st.html(f"""
        <div style="background:#fffbeb;border:1px solid #fbbf24;
                    border-left:4px solid #f59e0b;border-radius:12px;
                    padding:1.2rem 1.5rem;margin-bottom:1rem;">
            <p style="margin:0 0 .6rem;font-weight:700;font-size:.93rem;color:#78350f;
                      font-family:'Inter',sans-serif;">
                ⚠️ {len(r['missing_fields'])} campo(s) para preencher manualmente no Excel:
            </p>
            <ul style="margin:0;padding-left:1.2rem;font-size:.88rem;color:#78350f;
                       font-family:'Inter',sans-serif;line-height:1.8;">
                {items_html}
            </ul>
        </div>
        """)
    else:
        st.success("🎉 Todos os campos obrigatórios foram preenchidos automaticamente!")

    st.divider()

    with st.expander("📋 Ver todos os dados extraídos", expanded=False):
        rows = [
            {"Campo": label, "Valor": str(r["merged"].get(field, "—"))}
            for field, label in {**REQUIRED_FIELDS, **OPTIONAL_FIELDS}.items()
        ]
        st.table(rows)

    with st.expander("📄 Status dos documentos", expanded=False):
        col_a, col_b = st.columns(2)
        items = list(DOCUMENT_TYPES.items())
        mid   = (len(items) + 1) // 2
        for col, chunk in [(col_a, items[:mid]), (col_b, items[mid:])]:
            with col:
                for doc_type, dname in chunk:
                    ok  = doc_type in r["found_docs"]
                    req = " *(obrigatório)*" if doc_type in REQUIRED_DOCUMENTS else ""
                    st.markdown(f"{'✅' if ok else '❌'} **{dname}**{req}")


# ════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Admissão — FIBRA",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()

    if not TEMPLATE_PATH.exists():
        st.error(f"Template não encontrado: `{TEMPLATE_PATH}`")
        st.stop()

    km = render_sidebar()
    render_header()

    tab_upload, tab_result = st.tabs(["📤  Enviar Documentos", "📊  Resultado"])

    with tab_upload:
        st.html("""
        <p style="color:#607d9b;margin-bottom:1rem;font-size:.93rem;
                   font-family:'Inter',sans-serif;">
            Envie os arquivos abaixo.
            <span style="background:#fef3c7;color:#92400e;font-size:.72rem;font-weight:700;
                         padding:2px 7px;border-radius:99px;">OBRIGATÓRIO</span>
            indica documentos essenciais.
            Aceita <b>JPG · PNG · PDF · WEBP</b>.
        </p>
        """)
        uploaded = render_upload_tab()
        render_summary(uploaded)
        st.markdown("<div style='height:.75rem'></div>", unsafe_allow_html=True)

        if st.button(
            "🚀   Processar Documentos e Gerar Excel",
            type="primary",
            use_container_width=True,
            disabled=not uploaded,
        ):
            if km.get_active_key() is None:
                st.error("Nenhuma chave disponível. Adicione uma conta na barra lateral.")
            else:
                with st.spinner("Processando com Gemini AI…"):
                    process_documents(uploaded, km)
                if st.session_state.get("results"):
                    st.success("✅ Pronto! Clique na aba **📊 Resultado** para baixar o Excel.")

    with tab_result:
        render_result_tab()


if __name__ == "__main__":
    main()
