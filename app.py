"""
app.py — VIP / EFL Pick Generation System (Streamlit)
=====================================================
Requament -> SKU_MASTER + Inventory check -> VIP PICK + INDIA SO Pick Excel generate.

Run:    streamlit run app.py
Deploy: GitHub -> Streamlit Cloud (README බලන්න)
"""

from __future__ import annotations
from datetime import datetime

import pandas as pd
import streamlit as st

import pick_engine as E

st.set_page_config(page_title="VIP Pick Generator", page_icon="📦", layout="wide")

# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
st.title("📦 VIP / EFL Pick Generation System")
st.caption(
    "Requirement එක → SKU_MASTER + Inventory check → **VIP PICK** සහ "
    "**INDIA SO Pick** Excel files generate කරගන්න."
)

# --------------------------------------------------------------------------- #
# Sidebar — data source + settings
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("⚙️ Settings")
    source = st.radio("Data source", ["Excel Upload", "Google Sheets"], index=0)

    st.divider()
    st.subheader("Pick options")
    rounding = st.selectbox(
        "Carton rounding", ["floor", "round", "ceil"], index=0,
        help="Req Qty / divisor → carton ගණන round කරන විදිය. "
             "Reference file එකට ගැළපෙන්නේ floor (98.2%).",
    )
    per_min_cbm = st.number_input(
        "Per Minute CBM", min_value=0.0001, value=0.333333, step=0.01, format="%.6f",
        help="Productivity assumption — pick time එක estimate කරන්න.",
    )
    pick_date = st.date_input("Pick Date", value=datetime.now())

    st.divider()
    with st.expander("WMS constants (INDIA SO)"):
        wh_id = st.text_input("WH_ID", value="LPGL")
        client_code = st.text_input("CLIENT_CODE", value="INM0VIP")
        order_type = st.text_input("ORDER_TYPE", value="Sales Orders")

    with st.expander("LOAD ID QR — fixed header codes"):
        header_codes_raw = st.text_input(
            "Header QR codes (comma-separated)",
            value="INM0VIP, PKINM0, IMSA05",
            help="VIP PICK 'LOAD ID QR' sheet එකේ උඩින් පෙන්නන fixed QR codes.",
        )

header_qr_codes = [c.strip() for c in header_codes_raw.split(",") if c.strip()]

cfg = E.EngineConfig(
    rounding=rounding, per_minute_cbm=per_min_cbm,
    wh_id=wh_id, client_code=client_code, order_type=order_type,
    pick_date=datetime.combine(pick_date, datetime.min.time()),
    header_qr_codes=header_qr_codes,
)

# --------------------------------------------------------------------------- #
# Inputs
# --------------------------------------------------------------------------- #
req_df = sku_df = inv_df = None

if source == "Excel Upload":
    c1, c2, c3 = st.columns(3)
    with c1:
        f_req = st.file_uploader("1️⃣ Requirement (Requament)", type=["xlsx", "xls"])
    with c2:
        f_sku = st.file_uploader("2️⃣ SKU_MASTER", type=["xlsx", "xls"])
    with c3:
        f_inv = st.file_uploader("3️⃣ Inventory_Report", type=["xlsx", "xls"])

    if f_req:
        req_df = pd.read_excel(f_req)
    if f_sku:
        sku_df = pd.read_excel(f_sku)
    if f_inv:
        inv_df = pd.read_excel(f_inv)

else:  # Google Sheets
    try:
        import gsheet
        sa_info = st.secrets.get("gcp_service_account", None)
    except Exception:
        sa_info = None

    if not sa_info:
        st.warning(
            "Google Sheets මට වැඩ කරන්න `st.secrets['gcp_service_account']` ඕනේ. "
            "`.streamlit/secrets.toml.example` සහ README බලන්න. "
            "දැනට Excel Upload mode එක පාවිච්චි කරන්න පුළුවන්."
        )
    else:
        st.info("Google Sheet URL/Key එක සහ worksheet නම් දෙන්න.")
        gc1, gc2 = st.columns(2)
        with gc1:
            req_key = st.text_input("Requirement Sheet URL/Key")
            req_ws = st.text_input("Requirement worksheet", value="Sheet1")
            sku_key = st.text_input("SKU_MASTER Sheet URL/Key")
            sku_ws = st.text_input("SKU_MASTER worksheet", value="Sheet1")
        with gc2:
            inv_key = st.text_input("Inventory Sheet URL/Key")
            inv_ws = st.text_input("Inventory worksheet", value="Sheet1")
            write_back = st.checkbox("Output Google Sheet එකට ආපහු ලියන්න", value=False)
            out_key = st.text_input("Output Sheet URL/Key (write-back)") if write_back else ""

        if st.button("📥 Google Sheets වලින් load කරන්න"):
            try:
                req_df = gsheet.read_sheet(sa_info, req_key, req_ws)
                sku_df = gsheet.read_sheet(sa_info, sku_key, sku_ws)
                inv_df = gsheet.read_sheet(sa_info, inv_key, inv_ws)
                st.session_state["gs_loaded"] = (req_df, sku_df, inv_df, out_key if write_back else "")
                st.success("Load වුණා ✅")
            except Exception as ex:
                st.error(f"Google Sheets error: {ex}")

        if "gs_loaded" in st.session_state and req_df is None:
            req_df, sku_df, inv_df, _ = st.session_state["gs_loaded"]

# --------------------------------------------------------------------------- #
# Generate
# --------------------------------------------------------------------------- #
ready = req_df is not None and sku_df is not None and inv_df is not None

if not ready:
    st.info("ⓘ files 3ම (Requirement, SKU_MASTER, Inventory) දුන්නම generate කරන්න පුළුවන්.")
    st.stop()

if st.button("🚀 Generate Pick Files", type="primary", use_container_width=True):
    try:
        with st.spinner("Processing..."):
            res = E.run_pipeline(req_df, sku_df, inv_df, cfg)
        st.session_state["result"] = res
        st.success("Generate වුණා ✅")
    except Exception as ex:
        st.error(f"Error: {ex}")
        st.stop()

res = st.session_state.get("result")
if not res:
    st.stop()

# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
pick = res["pick"]
summary = res["summary"]
n_short = int(pick["REMARKS"].str.startswith("Shortage").sum())
n_flag = int((pick["REMARKS"] != "").sum())

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Pick lines", len(pick))
m2.metric("Deliveries", pick["OBD"].nunique())
m3.metric("Total Boxes", int(pick["HJ Box Qty"].sum()))
m4.metric("Total CBM", summary["Total CBM Of Pick"])
m5.metric("Shortage rows", n_short, delta=None)

st.subheader("⏱️ CBM / Pick-Time Summary")
st.dataframe(pd.DataFrame([summary]), hide_index=True, use_container_width=True)

if n_flag:
    with st.expander(f"⚠️ Flagged rows ({n_flag}) — Shortage / SKU/Inventory නැති ඒවා", expanded=n_short > 0):
        st.dataframe(
            pick[pick["REMARKS"] != ""][
                ["Material", "Qty", "OBD", "HJ Box Qty", "HJ Pcs Qty", "Pcs/Box", "_avail", "REMARKS"]
            ].rename(columns={"_avail": "Avail Inv"}),
            hide_index=True, use_container_width=True,
        )

tab1, tab2, tab3, tab4 = st.tabs(
    ["📋 VIP PICK", "🧾 INDIA SO — MASTER", "🧾 INDIA SO — Detail", "🏷️ LOAD ID QR"]
)
with tab1:
    st.dataframe(res["vip_table"], hide_index=True, use_container_width=True, height=420)
with tab2:
    st.dataframe(res["india_master"], hide_index=True, use_container_width=True, height=420)
with tab3:
    st.dataframe(res["india_detail"], hide_index=True, use_container_width=True, height=420)
with tab4:
    st.caption(
        f"{len(res['load_ids'])} unique LOAD IDs — VIP PICK Excel එකේ 'LOAD ID QR' "
        "sheet එකේ scannable QR code එකක් සමඟ generate වෙනවා. "
        f"Fixed header QR codes: {', '.join(cfg.header_qr_codes)}."
    )
    st.dataframe(
        pd.DataFrame({"LOAD ID": res["load_ids"]}),
        hide_index=True, use_container_width=True, height=420,
    )

st.divider()
st.subheader("⬇️ Downloads")
d1, d2 = st.columns(2)
stamp = cfg.pick_date.strftime("%Y%m%d")
with d1:
    st.download_button(
        "📦 VIP PICK download", data=res["vip_pick_bytes"],
        file_name=f"VIP_PICK_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with d2:
    st.download_button(
        "🧾 INDIA SO Pick download", data=res["india_so_bytes"],
        file_name=f"INDIA_SO_Pick_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# Optional Google Sheets write-back
if source == "Google Sheets" and "gs_loaded" in st.session_state:
    out_key = st.session_state["gs_loaded"][3]
    if out_key and st.button("📝 Output Google Sheet එකට ලියන්න"):
        try:
            import gsheet
            sa_info = st.secrets["gcp_service_account"]
            gsheet.write_sheet(sa_info, out_key, "VIP PICK", res["vip_table"])
            gsheet.write_sheet(sa_info, out_key, "OutBound MASTER", res["india_master"])
            gsheet.write_sheet(sa_info, out_key, "OutBound Detail", res["india_detail"])
            st.success("Google Sheet එකට ලිව්වා ✅")
        except Exception as ex:
            st.error(f"Write error: {ex}")

st.caption(
    "Logic: boxsize = inventory carton mode · divisor = boxsize(LOOSE)/SAP(SET) · "
    "HJ Box = floor(Qty/divisor) · HJ Pcs = Box×boxsize · "
    "INDIA SO Detail QTY = HJ Pcs Qty."
)
