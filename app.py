import platform
import time
from io import BytesIO

import pandas as pd
import streamlit as st

# -------- SAP automation imports (Windows only) --------
if platform.system().lower().startswith("win"):
    import win32com.client
    import win32clipboard


# ======================================================
# CONFIG PADRÃO (IDs do SAP GUI Scripting - IW32)
# ======================================================
SAP_IDS = {
    "OKCD_ID": "wnd[0]/tbar[0]/okcd",
    "OS_FIELD_ID": "wnd[0]/usr/ctxtCAUFVD-AUFNR",
    "TAB_OPERACOES_ID": "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:110 7/tabsTS_1100/tabpVGUE",
    "TBL_PATH": (
        "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:110 7/"
        "tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010"
    ),
    "BTN_TEXTO_LONGO_FMT": (
        "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:110 7/"
        "tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/btnLT ICON-LTOPR[8,{row}]"
    ),
    "LONGTEXT_SHELL_ID": "wnd[0]/usr/cntlSCMSW_CONTAINER_2102/shellcont/shell",
    "BTN_BACK_ID": "wnd[0]/tbar[0]/btn[3]",
    "BTN_SAVE_ID": "wnd[0]/tbar[0]/btn[11]",
}

VISIBLE_ROWS_DEFAULT = 15


# ======================================================
# UTILIDADES - EXCEL
# ======================================================

def _normalize_cols(cols):
    out = []
    for c in cols:
        if isinstance(c, str):
            out.append(c.strip())
        else:
            out.append(c)
    return out


def autodetect_header_row(excel_bytes: bytes, sheet_name: str, max_scan_rows: int = 30):
    """Tenta localizar a linha de cabeçalho procurando por 'OS' e 'Máscara'."""
    preview = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet_name, engine="openpyxl", header=None, nrows=max_scan_rows)
    for idx in range(len(preview)):
        row = preview.iloc[idx].astype(str).str.strip().str.lower().tolist()
        if ("os" in row) and ("máscara" in row or "mascara" in row):
            return idx
    return None


def load_dataframe(excel_bytes: bytes, sheet_name: str, header_row: int | None):
    if header_row is None:
        # fallback: tenta padrão do seu arquivo
        header_row = 3

    df = pd.read_excel(BytesIO(excel_bytes), sheet_name=sheet_name, engine="openpyxl", header=header_row)
    df.columns = _normalize_cols(df.columns)

    # Normaliza nomes principais se vierem com espaços
    rename_map = {}
    for c in df.columns:
        if isinstance(c, str):
            if c.strip().lower() == "status":
                rename_map[c] = "Status"
            if c.strip().lower() == "máscara" or c.strip().lower() == "mascara":
                rename_map[c] = "Máscara"
    if rename_map:
        df = df.rename(columns=rename_map)

    return df


def coerce_os_to_str(series: pd.Series) -> pd.Series:
    def _conv(x):
        if pd.isna(x):
            return ""
        # Muitos arquivos vêm com OS como float (ex.: 6000794541.0)
        try:
            if isinstance(x, (int,)):
                return str(x)
            if isinstance(x, float):
                return str(int(x))
            s = str(x).strip()
            # remove .0
            if s.endswith(".0"):
                s = s[:-2]
            return s
        except Exception:
            return str(x)

    return series.apply(_conv)


# ======================================================
# UTILIDADES - SAP GUI Scripting
# ======================================================

def wait_not_busy(session, timeout=60):
    t0 = time.time()
    while session.Busy:
        if time.time() - t0 > timeout:
            raise TimeoutError("SAP ficou ocupado tempo demais (Busy).")
        time.sleep(0.1)


def connect_sap_session(connection_index: int = 0, session_index: int = 0):
    sap = win32com.client.GetObject("SAPGUI")
    app = sap.GetScriptingEngine
    conn = app.Children(connection_index)
    sess = conn.Children(session_index)
    return sess


def set_clipboard(texto: str):
    # preserva quebras de linha
    texto = (texto or "").replace("\n", "\r\n")
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(texto)
    finally:
        win32clipboard.CloseClipboard()


def ensure_visible(tbl, abs_row: int, visible_rows: int):
    if abs_row < visible_rows:
        scroll_pos = 0
    else:
        scroll_pos = abs_row - (visible_rows - 1)
    tbl.VerticalScrollbar.Position = scroll_pos
    return abs_row - scroll_pos


def open_iw32_and_load_os(session, os_str: str):
    session.findById(SAP_IDS["OKCD_ID"]).text = "/nIW32"
    session.findById("wnd[0]").sendVKey(0)
    wait_not_busy(session)

    session.findById(SAP_IDS["OS_FIELD_ID"]).text = os_str
    session.findById(SAP_IDS["OS_FIELD_ID"]).caretPosition = len(os_str)
    session.findById("wnd[0]").sendVKey(0)
    wait_not_busy(session)


def push_to_sap(os_str: str, long_texts: list[str], visible_rows: int, save_after: bool,
                connection_index: int, session_index: int, progress_cb=None, log_cb=None):
    if not platform.system().lower().startswith("win"):
        raise RuntimeError("Este envio ao SAP requer Windows (SAP GUI + COM).")

    session = connect_sap_session(connection_index=connection_index, session_index=session_index)
    session.findById("wnd[0]").maximize

    if log_cb:
        log_cb(f"Abrindo IW32 e carregando OS {os_str}...")

    open_iw32_and_load_os(session, os_str)

    # Aba Operações
    session.findById(SAP_IDS["TAB_OPERACOES_ID"]).select
    wait_not_busy(session)

    tbl = session.findById(SAP_IDS["TBL_PATH"])  # GuiTableControl

    total = len(long_texts)
    for i, texto in enumerate(long_texts):
        if progress_cb:
            progress_cb((i + 1) / max(total, 1))

        vis_row = ensure_visible(tbl, i, visible_rows)

        # abre texto longo da linha
        session.findById(SAP_IDS["BTN_TEXTO_LONGO_FMT"].format(row=vis_row)).press
        wait_not_busy(session)

        # cola do Excel via clipboard e aplica setDocum
        set_clipboard(texto)
        session.findById(SAP_IDS["LONGTEXT_SHELL_ID"]).setDocum
        wait_not_busy(session)

        # voltar para a tabela
        session.findById(SAP_IDS["BTN_BACK_ID"]).press
        wait_not_busy(session)

        if log_cb:
            log_cb(f"Linha {i+1}/{total}: texto longo aplicado.")

    if save_after:
        if log_cb:
            log_cb("Salvando a ordem...")
        session.findById(SAP_IDS["BTN_SAVE_ID"]).press
        wait_not_busy(session)

    if progress_cb:
        progress_cb(1.0)

    return True


# ======================================================
# STREAMLIT APP
# ======================================================

st.set_page_config(page_title="IW32 - Preencher Texto Longo (ZSUB)", layout="wide")

st.title("Automação IW32 – Preencher Texto Longo por OS (via SAP GUI Scripting)")

with st.expander("⚠️ Pré-requisitos", expanded=True):
    st.markdown(
        """
**Para funcionar:**
- Este app deve rodar **no mesmo PC** onde o **SAP GUI** está instalado.
- O SAP deve estar **aberto**, com você **logado** no ambiente correto.
- O SAP GUI Scripting precisa estar **habilitado** no cliente.

> Observação: o envio ao SAP usa integração COM (Windows). Em Linux/Mac a interface abre, mas o envio ao SAP não funciona.
"""
    )

uploaded = st.file_uploader("1) Envie a planilha (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.stop()

excel_bytes = uploaded.read()

# Seleção de aba
try:
    xls = pd.ExcelFile(BytesIO(excel_bytes), engine="openpyxl")
    sheet = st.selectbox("2) Selecione a aba (sheet)", xls.sheet_names)
except Exception as e:
    st.error(f"Não consegui ler o arquivo Excel: {e}")
    st.stop()

# Detecta header
auto_header = autodetect_header_row(excel_bytes, sheet)
col1, col2 = st.columns([1, 1])
with col1:
    use_auto = st.checkbox("Detectar cabeçalho automaticamente", value=True)
with col2:
    header_row = st.number_input("Linha do cabeçalho (0 = primeira)", min_value=0, max_value=200, value=int(auto_header if auto_header is not None else 3))

try:
    df = load_dataframe(excel_bytes, sheet, header_row if not use_auto else (auto_header if auto_header is not None else header_row))
except Exception as e:
    st.error(f"Erro ao carregar dados: {e}")
    st.stop()

# Mapeamento de colunas (caso o usuário use outra variação)
cols = list(df.columns)
cols_lower = [str(c).strip().lower() for c in cols]

def _suggest(name: str):
    name = name.lower()
    for c in cols:
        if str(c).strip().lower() == name:
            return c
    # fallback aproximado
    for c in cols:
        if name in str(c).strip().lower():
            return c
    return cols[0]

st.subheader("3) Mapeamento de colunas")
map_col_os = st.selectbox("Coluna da OS (ordem)", cols, index=cols.index(_suggest("os")) if _suggest("os") in cols else 0)
map_col_texto = st.selectbox("Coluna do Texto Longo", cols, index=cols.index(_suggest("máscara")) if _suggest("máscara") in cols else (cols.index(_suggest("mascara")) if _suggest("mascara") in cols else 0))

# Prepara OS
if map_col_os not in df.columns:
    st.error("Coluna de OS não encontrada.")
    st.stop()

work = df.copy()
work["OS_str"] = coerce_os_to_str(work[map_col_os])
work = work[work["OS_str"] != ""]

if work.empty:
    st.warning("Não encontrei nenhuma OS válida na planilha.")
    st.stop()

os_list = work["OS_str"].dropna().unique().tolist()

st.subheader("4) Escolha a OS e confira os dados")
selected_os = st.selectbox("OS", os_list)

rows_os = work[work["OS_str"] == selected_os].copy()

# Ordenação: mantém ordem original do arquivo (índice)
rows_os = rows_os.reset_index(drop=False).rename(columns={"index": "_linha_excel"})

# Colunas para exibição
show_mask = st.checkbox("Mostrar coluna do texto longo na tabela (pode ficar pesado)", value=False)

default_show = [c for c in ["_linha_excel", map_col_os, "Operação", "Material", "Texto breve material", "Quantidade", "Centro"] if c in rows_os.columns]

if show_mask:
    default_show.append(map_col_texto)

st.dataframe(rows_os[default_show], use_container_width=True, height=350)

# Visualização do texto longo por linha
st.markdown("**Visualizar texto longo de uma linha específica**")
line_choice = st.number_input("Linha (0 = primeira dentro da OS)", min_value=0, max_value=max(len(rows_os)-1, 0), value=0)
preview_text = ""
if map_col_texto in rows_os.columns and len(rows_os) > 0:
    preview_text = str(rows_os.loc[int(line_choice), map_col_texto])
st.text_area("Texto longo", preview_text, height=220)

# Botão para exportar prévia
st.download_button(
    "Baixar prévia (CSV)",
    data=rows_os.to_csv(index=False).encode("utf-8"),
    file_name=f"preview_OS_{selected_os}.csv",
    mime="text/csv",
)

st.divider()

st.subheader("5) Enviar para o SAP (IW32)")

if not platform.system().lower().startswith("win"):
    st.error("Este recurso só funciona no Windows (SAP GUI Scripting via COM).")
    st.stop()

ack = st.checkbox("✅ SAP GUI está aberto e eu estou logado. Quero executar a automação.", value=False)

c1, c2, c3 = st.columns([1, 1, 1])
with c1:
    visible_rows = st.number_input("Linhas visíveis na tabela", min_value=5, max_value=30, value=VISIBLE_ROWS_DEFAULT)
with c2:
    conn_idx = st.number_input("Connection index", min_value=0, max_value=9, value=0)
with c3:
    sess_idx = st.number_input("Session index", min_value=0, max_value=9, value=0)

save_after = st.checkbox("Salvar a OS ao final", value=True)

st.info("Dica: durante a execução, **não mexa no SAP** (mouse/teclado) para evitar perder foco.")

run_btn = st.button("🚀 Enviar texto longo para o SAP", type="primary", disabled=not ack)

if run_btn:
    if map_col_texto not in rows_os.columns:
        st.error("A coluna de texto longo selecionada não existe no dataframe.")
        st.stop()

    long_texts = rows_os[map_col_texto].fillna("").astype(str).tolist()

    prog = st.progress(0.0)
    log_area = st.empty()
    logs = []

    def log_cb(msg):
        logs.append(msg)
        # mostra os últimos 15 logs
        log_area.code("\n".join(logs[-15:]))

    try:
        push_to_sap(
            os_str=selected_os,
            long_texts=long_texts,
            visible_rows=int(visible_rows),
            save_after=save_after,
            connection_index=int(conn_idx),
            session_index=int(sess_idx),
            progress_cb=lambda v: prog.progress(min(max(float(v), 0.0), 1.0)),
            log_cb=log_cb,
        )
        st.success("✅ Envio concluído com sucesso!")
    except Exception as e:
        st.error(f"Falha ao executar: {e}")
        st.warning("Se o SAP tiver mais de uma sessão aberta, ajuste Connection/Session index.")
