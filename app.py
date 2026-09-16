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

st.set_page_config(
    page_title="Conta Fácil",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .stApp { background: #f5f7fb; }
      .block-container { max-width: 1500px; padding-top: 1.25rem; padding-bottom: 2rem; }
      [data-testid="stHeader"] { background: rgba(0,0,0,0); }
      .cf-header { display:flex; align-items:center; justify-content:space-between; gap:24px;
                   padding:18px 22px; border:1px solid #dbe3ec; border-radius:14px;
                   background:#ffffff; box-shadow:0 2px 12px rgba(15,39,71,.05); }
      .cf-title { font-size:1.65rem; font-weight:750; color:#0f2f55; margin:0; }
      .cf-sub { color:#65758b; font-size:.92rem; margin-top:2px; }
      .cf-org { color:#174f7a; font-weight:700; font-size:.83rem; text-align:right; }
      .cf-card { background:#fff; border:1px solid #dbe3ec; border-radius:14px; padding:18px 20px;
                 box-shadow:0 2px 10px rgba(15,39,71,.035); }
      .cf-kicker { color:#668096; font-size:.75rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; }
      .cf-value { color:#08744f; font-size:2rem; font-weight:800; margin-top:2px; }
      .cf-muted { color:#6d7d90; font-size:.88rem; }
      div[data-testid="stMetric"] { background:#fff; border:1px solid #dbe3ec; border-radius:12px; padding:10px 14px; }
      .stTabs [data-baseweb="tab-list"] { gap:8px; border-bottom:1px solid #d9e1ea; }
      .stTabs [data-baseweb="tab"] { height:46px; padding:0 18px; border-radius:9px 9px 0 0; font-weight:650; }
      .stButton>button, .stDownloadButton>button { border-radius:9px; min-height:40px; font-weight:650; }
      [data-testid="stDataFrame"] { border:1px solid #dbe3ec; border-radius:10px; overflow:hidden; }
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

h1, h2, h3 = st.columns([1.25, 4.6, 2.2], vertical_alignment="center")
with h1:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=155)
with h2:
    st.markdown('<div class="cf-title">Conta Fácil</div><div class="cf-sub">Gestão e conferência de faturamento de água • EMBASA</div>', unsafe_allow_html=True)
with h3:
    st.markdown('<div class="cf-org">UFRB<br>PRÓ-REITORIA DE ADMINISTRAÇÃO</div>', unsafe_allow_html=True)
st.divider()

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

st.caption("Conta Fácil • UFRB • EMBASA")
