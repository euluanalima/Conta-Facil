# -*- coding: utf-8 -*-
"""Interface web do Conta Fácil.

A camada web chama as mesmas funções de extração, consolidação, cálculos,
Excel, gráficos e Word presentes no núcleo da versão desktop.
"""
from __future__ import annotations

import re
import tempfile
from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st

from conta_facil_core import (
    brl,
    pct,
    month_key,
    month_sort_key,
    extract_pdf,
    normalize_history_excel,
    build_control,
    export_excel,
    make_chart_images,
    generate_word_report,
    history_snapshot_from_control,
    load_local_history,
    save_local_history,
)

APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "logo_conta_facil.svg"
MONTH_RE = re.compile(r"(?:Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez)/20\d{2}")


# Apresentação apenas: reaproveita a arte existente como ícone compacto.
LOGO_SVG = ""
if LOGO_PATH.exists():
    LOGO_SVG = LOGO_PATH.read_text(encoding="utf-8")
    LOGO_SVG = re.sub(r'width="[^"]+"', '', LOGO_SVG, count=1)
    LOGO_SVG = re.sub(r'height="[^"]+"', '', LOGO_SVG, count=1)
    LOGO_SVG = re.sub(r'viewBox="[^"]+"', 'viewBox="250 45 430 400"', LOGO_SVG, count=1)
    LOGO_SVG = LOGO_SVG.replace("<svg ", '<svg class="cf-logo-svg" preserveAspectRatio="xMidYMid meet" ', 1)

st.set_page_config(
    page_title="Conta Fácil",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      :root {
        --cf-navy: #123b67;
        --cf-blue: #176da5;
        --cf-green: #08744f;
        --cf-border: #dbe3ec;
        --cf-muted: #65758b;
        --cf-bg: #f5f7fb;
      }

      html, body, [class*="css"] { font-family: "Segoe UI", Arial, sans-serif; }
      .stApp { background: var(--cf-bg); }
      .block-container {
        width: min(100%, 1500px);
        max-width: 1500px;
        padding: 1.15rem 1.35rem 2rem;
      }
      [data-testid="stHeader"] { background: rgba(245,247,251,.92); }
      [data-testid="stToolbar"] { right: .75rem; }

      .cf-header {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        align-items: center;
        gap: 1.25rem;
        width: 100%;
        padding: 1rem 1.2rem;
        border: 1px solid var(--cf-border);
        border-radius: 14px;
        background: #fff;
        box-shadow: 0 2px 12px rgba(15,39,71,.05);
      }
      .cf-brand {
        display: flex;
        align-items: center;
        gap: .9rem;
        min-width: 0;
      }
      .cf-logo {
        width: 82px;
        height: 82px;
        flex: 0 0 82px;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow: hidden;
      }
      .cf-logo-svg {
        width: 82px;
        height: 82px;
        display: block;
      }
      .cf-brand-copy { min-width: 0; }
      .cf-title {
        font-size: clamp(1.45rem, 2vw, 1.85rem);
        line-height: 1.08;
        font-weight: 780;
        color: #0f2f55;
        margin: 0;
        letter-spacing: -.02em;
      }
      .cf-sub {
        color: var(--cf-muted);
        font-size: clamp(.78rem, 1vw, .92rem);
        line-height: 1.35;
        margin-top: .28rem;
      }
      .cf-org {
        color: var(--cf-navy);
        background: #f3f7fb;
        border: 1px solid #d7e2ee;
        border-radius: 10px;
        padding: .65rem .9rem;
        font-weight: 750;
        font-size: .86rem;
        line-height: 1.2;
        text-align: center;
        white-space: nowrap;
      }

      .cf-card {
        background:#fff;
        border:1px solid var(--cf-border);
        border-radius:14px;
        padding:18px 20px;
        box-shadow:0 2px 10px rgba(15,39,71,.035);
      }
      .cf-kicker {
        color:#668096;
        font-size:.75rem;
        font-weight:700;
        letter-spacing:.08em;
        text-transform:uppercase;
      }
      .cf-value { color:var(--cf-green); font-size:2rem; font-weight:800; margin-top:2px; }
      .cf-muted { color:#6d7d90; font-size:.88rem; }

      div[data-testid="stMetric"] {
        background:#fff;
        border:1px solid var(--cf-border);
        border-radius:12px;
        padding:10px 14px;
        min-width:0;
      }
      [data-testid="stFileUploader"] { width:100%; min-width:0; }
      div[data-testid="stColumn"] { min-width:0; }
      div[data-testid="stHorizontalBlock"] { width:100%; align-items:stretch; }
      [data-testid="stFileUploaderDropzone"] {
        border-radius:12px;
        border-color:#cad7e4;
        background:#fff;
      }
      .stTabs [data-baseweb="tab-list"] {
        gap:8px;
        border-bottom:1px solid #d9e1ea;
        overflow-x:auto;
        scrollbar-width:thin;
      }
      .stTabs [data-baseweb="tab"] {
        height:46px;
        padding:0 18px;
        border-radius:9px 9px 0 0;
        font-weight:650;
        white-space:nowrap;
        flex:0 0 auto;
      }
      .stButton>button, .stDownloadButton>button {
        border-radius:9px;
        min-height:40px;
        font-weight:650;
      }
      [data-testid="stDataFrame"] {
        width:100%;
        border:1px solid var(--cf-border);
        border-radius:10px;
        overflow:auto;
      }
      [data-testid="stImage"] img {
        max-width:100%;
        height:auto;
      }

      @media (max-width: 980px) {
        .block-container {
          width:100%;
          max-width:100%;
          padding:.85rem .8rem 1.5rem;
        }
        .cf-header {
          grid-template-columns:1fr;
          gap:.75rem;
          padding:.9rem 1rem;
        }
        .cf-org {
          width:100%;
          white-space:normal;
        }
        div[data-testid="stHorizontalBlock"] {
          display:flex !important;
          flex-direction:column !important;
          align-items:stretch !important;
          gap:.75rem !important;
          width:100% !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
          width:100% !important;
          min-width:0 !important;
          max-width:100% !important;
          flex:1 1 100% !important;
        }
        div[data-testid="stMetric"] {
          width:100% !important;
        }
      }

      @media (max-width: 620px) {
        .block-container { padding:.6rem .6rem 1.25rem; }
        .cf-header {
          padding:.8rem;
          border-radius:12px;
        }
        .cf-brand {
          width:100%;
          gap:.7rem;
          align-items:center;
        }
        .cf-brand-copy {
          width:calc(100% - 76px);
          min-width:0;
        }
        .cf-logo, .cf-logo-svg {
          width:68px;
          height:68px;
          flex-basis:68px;
        }
        .cf-title { font-size:1.4rem; }
        .cf-sub { font-size:.78rem; }
        .cf-org {
          font-size:.8rem;
          padding:.55rem .7rem;
        }

        div[data-testid="stHorizontalBlock"] {
          display:flex !important;
          flex-direction:column !important;
          align-items:stretch !important;
          gap:.6rem !important;
          width:100% !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
          width:100% !important;
          min-width:0 !important;
          max-width:100% !important;
          flex:1 1 100% !important;
        }
        .stButton, .stDownloadButton { width:100%; }
        .stButton>button, .stDownloadButton>button { width:100%; }
        [data-testid="stFileUploaderDropzone"] {
          padding:.7rem !important;
          min-height:84px;
        }
        [data-testid="stMetric"] { width:100%; }
        .stTabs [data-baseweb="tab"] {
          height:42px;
          padding:0 13px;
          font-size:.86rem;
        }
        [data-testid="stDataFrame"] {
          max-width:calc(100vw - 1.2rem);
        }
        h2, h3 { overflow-wrap:anywhere; }
      }

      @media (max-width: 390px) {
        .cf-logo, .cf-logo-svg {
          width:58px;
          height:58px;
          flex-basis:58px;
        }
        .cf-title { font-size:1.25rem; }
        .cf-sub { font-size:.72rem; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)

if "history_df" not in st.session_state:
    local_history, _ = load_local_history()
    st.session_state.history_df = local_history
if "faturas" not in st.session_state:
    st.session_state.faturas = None
if "control" not in st.session_state:
    st.session_state.control = None
if "resumo" not in st.session_state:
    st.session_state.resumo = None
if "mes_ano" not in st.session_state:
    st.session_state.mes_ano = ""

st.markdown(
    f"""
    <div class="cf-header">
      <div class="cf-brand">
        <div class="cf-logo">{LOGO_SVG}</div>
        <div class="cf-brand-copy">
          <div class="cf-title">Conta Fácil</div>
          <div class="cf-sub">Gestão e conferência de faturamento de água • EMBASA</div>
        </div>
      </div>
      <div class="cf-org">UFRB - NUMAM</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.write("")

left, right = st.columns([4.7, 2], gap="large")
with left:
    st.subheader("Arquivo de faturamento")
    st.caption("Envie o PDF consolidado da EMBASA para processar as contas do período.")
    pdf_file = st.file_uploader("PDF consolidado", type=["pdf"], label_visibility="collapsed")
with right:
    total = 0
    if st.session_state.faturas is not None and not st.session_state.faturas.empty:
        total = int(st.session_state.faturas["Valor_Centavos"].sum())
    st.metric("Valor total geral", brl(total))
    if st.session_state.mes_ano:
        st.caption(f"Referência: {st.session_state.mes_ano}")

history_file = st.file_uploader(
    "Carregar histórico Excel (opcional)",
    type=["xlsx", "xls"],
    help="Aceita o Excel exportado pelo próprio Conta Fácil.",
)

c1, c2 = st.columns([1, 5])
with c1:
    process = st.button("Processar PDF", type="primary", use_container_width=True, disabled=pdf_file is None)

if history_file is not None:
    try:
        suffix = Path(history_file.name).suffix or ".xlsx"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(history_file.getbuffer())
            hist_path = tmp.name
        st.session_state.history_df = normalize_history_excel(hist_path)
        st.success("Histórico carregado com sucesso.")
    except Exception as exc:
        st.error(f"Não foi possível carregar o histórico: {exc}")

if process and pdf_file is not None:
    try:
        with st.spinner("Processando faturas..."):
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(pdf_file.getbuffer())
                pdf_path = tmp.name
            faturas = extract_pdf(pdf_path)
            control, resumo, mes_ano = build_control(faturas, st.session_state.history_df)

            history_df = history_snapshot_from_control(control)
            st.session_state.history_df = history_df
            try:
                save_local_history(history_df)
            except Exception:
                pass

            st.session_state.faturas = faturas
            st.session_state.control = control
            st.session_state.resumo = resumo
            st.session_state.mes_ano = mes_ano
        st.success(f"{len(faturas)} faturas processadas • Referência {mes_ano} • Total {brl(int(faturas['Valor_Centavos'].sum()))}")
    except Exception as exc:
        st.error(f"Erro ao processar o PDF: {exc}")

control = st.session_state.control
faturas = st.session_state.faturas
resumo = st.session_state.resumo
mes_ano = st.session_state.mes_ano

if control is None or faturas is None or resumo is None or not mes_ano:
    st.info("Envie e processe um PDF para habilitar a planilha, os gráficos e o relatório.")
    st.stop()

def display_control(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if MONTH_RE.fullmatch(str(c)):
            out[c] = out[c].apply(lambda x: "" if pd.isna(x) else brl(int(x)))
        elif c in {"Aum / Red", "% Centro", "% Total", "Variação %"}:
            out[c] = out[c].apply(pct)
    return out

tab1, tab2, tab3 = st.tabs(["Planilha de Controle", "Gráficos Analíticos", "Relatório de Ateste"])

with tab1:
    a, b = st.columns([3.8, 1.2], gap="large")
    with a:
        query = st.text_input("Pesquisar local", placeholder="Digite parte do local, centro ou matrícula...")
    filtered = control
    if query.strip():
        q = query.casefold().strip()
        mask = pd.Series(False, index=control.index)
        for col in ("LOCAL", "CEN", "MAT"):
            if col in control.columns:
                mask = mask | control[col].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
        filtered = control.loc[mask]
    with b:
        st.metric("Contas exibidas", f"{len(filtered)} / {len(control)}")

    st.dataframe(display_control(filtered), use_container_width=True, hide_index=True, height=520)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        excel_path = tmp.name
    export_excel(excel_path, control, faturas, resumo, mes_ano)
    excel_bytes = Path(excel_path).read_bytes()
    st.download_button(
        "Baixar Excel",
        data=excel_bytes,
        file_name=f"Controle_EMBASA_{mes_ano.replace('/', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab2:
    with tempfile.TemporaryDirectory() as td:
        center_img, trend_img = make_chart_images(control, resumo, mes_ano, td)
        st.subheader("Distribuição do faturamento por centro")
        st.image(center_img, use_container_width=True)
        st.subheader("Evolução mensal do faturamento por centro")
        st.image(trend_img, use_container_width=True)

with tab3:
    st.subheader("Dados do ateste")
    col1, col2 = st.columns(2, gap="large")
    with col1:
        contrato = st.text_input("Contrato", value="Nº 06/2015")
        gestor = st.text_input("Nome do gestor")
    with col2:
        siape = st.text_input("SIAPE")
        data_ateste = st.date_input("Data do ateste", value=date.today(), format="DD/MM/YYYY")
    ocorrencias = st.text_area(
        "Ocorrências / justificativas do mês (opcional)",
        height=130,
        placeholder="Ex.: vazamento identificado, intervenção executada, alteração de uso do imóvel...",
    )

    if st.button("Gerar relatório Word", type="primary"):
        if not gestor.strip() or not siape.strip() or not contrato.strip():
            st.warning("Preencha contrato, nome do gestor e SIAPE antes de gerar o relatório.")
        else:
            try:
                with st.spinner("Gerando relatório..."):
                    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                        docx_path = tmp.name
                    generate_word_report(
                        docx_path,
                        control,
                        faturas,
                        resumo,
                        mes_ano,
                        contrato.strip(),
                        gestor.strip(),
                        siape.strip(),
                        data_ateste.strftime("%d/%m/%Y"),
                        ocorrencias.strip(),
                    )
                    st.session_state.report_bytes = Path(docx_path).read_bytes()
                st.success("Relatório gerado com sucesso.")
            except Exception as exc:
                st.error(f"Erro ao gerar o relatório: {exc}")

    if st.session_state.get("report_bytes"):
        st.download_button(
            "Baixar relatório Word",
            data=st.session_state.report_bytes,
            file_name=f"Relatorio_Mensal_Ateste_EMBASA_{mes_ano.replace('/', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

st.caption("Conta Fácil • UFRB - NUMAM • EMBASA")
