"""
app.py — VIP / EFL Pick Generation System (Streamlit)
=====================================================
Requirement -> SKU_MASTER + Inventory check -> VIP PICK + INDIA SO Pick.

Features
  - Pick formula: target = Req Qty * HJ / SAP, picked only in whole cartons.
  - Cannot-Pick report: carton-split + insufficient-stock lines.
  - SKU_MASTER stored in Google Sheet with add / edit / delete (CRUD).
  - LOAD ID global uniqueness via a Google Sheet registry (-A/-B suffix).
  - Auto-save every output + monthly pick history to Google Sheet.

Run:    streamlit run app.py
Deploy: GitHub -> Streamlit Cloud (README බලන්න)
"""
from __future__ import annotations
from datetime import datetime

import pandas as pd
import streamlit as st

import pick_engine as E

st.set_page_config(page_title="VIP Pick Generator", page_icon="📦", layout="wide")


def get_sa():
    try:
        return st.secrets.get("gcp_service_account", None)
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# Sidebar — settings
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("⚙️ Settings")
    source = "Excel Upload"

    st.divider()
    st.subheader("Pick options")
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

    st.divider()
    st.subheader("💾 Google Sheet (save + registry + history)")
    save_key = st.text_input(
        "Data Sheet URL / Key",
        help="හැම output එකක්ම, LOAD_ID registry, monthly history මෙතන save වෙනවා. "
             "Service-account email එකට Editor විදිහට share කරන්න.",
    )
    autosave = st.checkbox("Generate කළාම auto-save", value=bool(save_key))
    use_registry = st.checkbox(
        "LOAD_ID duplicate check (registry)", value=bool(save_key),
        help="මේ Sheet එකේ registry එකට බලලා duplicate LOAD ID එකකට -A/-B වගේ "
             "suffix එකක් දානවා.",
    )
    sku_ws = st.text_input("SKU_MASTER worksheet name", value="SKU_MASTER")

header_qr_codes = [c.strip() for c in header_codes_raw.split(",") if c.strip()]
cfg = E.EngineConfig(
    per_minute_cbm=per_min_cbm,
    wh_id=wh_id, client_code=client_code, order_type=order_type,
    pick_date=datetime.combine(pick_date, datetime.min.time()),
    header_qr_codes=header_qr_codes,
)
sa_info = get_sa()

st.title("📦 VIP / EFL Pick Generation System")

tab_gen, tab_sku, tab_hist = st.tabs(
    ["🚀 Generate Pick", "🗂️ SKU_MASTER", "📅 History"]
)

# =========================================================================== #
# TAB: SKU_MASTER (CRUD)
# =========================================================================== #
with tab_sku:
    st.subheader("🗂️ SKU_MASTER — upload / add / edit / delete")
    st.caption(
        "SKU_MASTER Excel එකක් upload කරලා හරි, Google Sheet එකෙන් load කරලා හරි — "
        "rows add/edit/delete කරලා Google Sheet එකට save කරන්න. "
        "LOOSE → SAP = HJ · SET → SAP, HJ වෙනස්."
    )

    # ---- 1) Upload SKU_MASTER from Excel (separate, always available) ----
    with st.container(border=True):
        st.markdown("**1️⃣ SKU_MASTER Excel upload කරන්න**")
        up = st.file_uploader("SKU_MASTER (.xlsx / .xls / .csv)",
                              type=["xlsx", "xls", "csv"], key="sku_upload")
        uc1, uc2 = st.columns([1, 2])
        if up is not None and uc1.button("⬆️ Editor එකට load කරන්න"):
            try:
                raw = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                cols = {c.strip().lower(): c for c in raw.columns}
                want = ["Material code", "Material Desc", "Catergory", "HJ", "SAP"]
                alias = {
                    "Material code": ["material code", "material", "code", "sku"],
                    "Material Desc": ["material desc", "description", "material description", "desc"],
                    "Catergory": ["catergory", "category", "type"],
                    "HJ": ["hj"], "SAP": ["sap"],
                }
                out = pd.DataFrame()
                for w in want:
                    found = next((cols[a] for a in alias[w] if a in cols), None)
                    out[w] = raw[found] if found else ""
                st.session_state["sku_edit"] = out
                st.success(f"Editor එකට load වුණා ✅  ({len(out)} rows). පහළ edit කරලා save කරන්න.")
            except Exception as ex:
                st.error(f"Upload error: {ex}")
        uc2.caption("Columns auto-map වෙනවා: Material code · Material Desc · Catergory · HJ · SAP")

    # ---- 2) Google Sheet load / save ----
    st.markdown("**2️⃣ Google Sheet (load / save)**")
    if not sa_info:
        st.warning("Google Sheets credentials නෑ — upload කරපු data edit කරන්න පුළුවන්, "
                   "save කරන්න `st.secrets['gcp_service_account']` ඕන (README බලන්න).")
    elif not save_key:
        st.info("Sidebar එකේ **Data Sheet URL/Key** එක දාන්න (save/load කරන්න).")
    else:
        import gsheet
        cinit, cload, csave = st.columns([1.2, 1, 1])
        if cinit.button("🆕 Initialize sheet (worksheets + headers)"):
            try:
                r = gsheet.init_sheet(sa_info, save_key)
                created = [k for k, v in r.items() if v is True]
                st.success("Initialize වුණා ✅ — " +
                           (f"created: {', '.join(created)}" if created else "ඔක්කොම දැනටමත් තිබුණා"))
                st.markdown(f"🔗 [Google Sheet එක open කරන්න]({r.get('url','')})")
            except Exception as ex:
                st.error(f"Init error: {ex}")
        if cload.button("📥 SKU_MASTER load කරන්න"):
            try:
                st.session_state["sku_edit"] = gsheet.read_sku_master(sa_info, save_key, sku_ws)
                st.success("Load වුණා ✅")
            except Exception as ex:
                st.error(f"Load error: {ex}")

    # ---- 3) Editor + save ----
    if "sku_edit" in st.session_state:
        st.markdown("**3️⃣ Edit (add / change / delete rows)**")
        edited = st.data_editor(
            st.session_state["sku_edit"],
            num_rows="dynamic", use_container_width=True, height=460,
            key="sku_editor",
            column_config={
                "Catergory": st.column_config.SelectboxColumn(
                    "Catergory", options=["LOOSE", "SET"]),
                "HJ": st.column_config.NumberColumn("HJ", min_value=0, step=1),
                "SAP": st.column_config.NumberColumn("SAP", min_value=0, step=1),
            },
        )
        st.caption(f"{len(edited)} rows · row delete කරන්න row එක select කරලා 🗑️ icon එක.")
        if sa_info and save_key:
            if st.button("💾 SKU_MASTER Google Sheet එකට save කරන්න", type="primary"):
                try:
                    import gsheet
                    gsheet.save_sku_master(sa_info, save_key, edited, sku_ws)
                    st.session_state["sku_edit"] = edited
                    st.success(f"Save වුණා ✅  ({len(edited)} rows)")
                except Exception as ex:
                    st.error(f"Save error: {ex}")
        else:
            st.info("Google Sheet එකට save කරන්න credentials + Data Sheet key එක ඕන.")

# =========================================================================== #
# TAB: History
# =========================================================================== #
with tab_hist:
    st.subheader("📅 Monthly Pick History")
    st.caption("හැම pick run එකක්ම මාසෙට අදාල tab එකකට save වෙනවා — මෙතන බලන්න පුළුවන්.")
    if not sa_info:
        st.warning("Google Sheets credentials නෑ.")
    elif not save_key:
        st.info("Sidebar එකේ **Data Sheet URL/Key** එක දාන්න.")
    else:
        import gsheet
        if st.button("🔄 History months refresh"):
            try:
                st.session_state["hist_months"] = gsheet.list_history_months(sa_info, save_key)
            except Exception as ex:
                st.error(f"Error: {ex}")
        months = st.session_state.get("hist_months", [])
        if months:
            sel = st.selectbox("Month", months)
            if st.button("📄 මේ මාසේ data පෙන්නන්න"):
                try:
                    hdf = gsheet.read_history(sa_info, save_key, sel)
                    st.dataframe(hdf, hide_index=True, use_container_width=True, height=460)
                    st.metric("Rows", len(hdf))
                except Exception as ex:
                    st.error(f"Error: {ex}")
        else:
            st.info("History months බලන්න 'refresh' click කරන්න (අඩුම තරමේ එක pick එකක් save කරලා තියෙන්න ඕන).")

# =========================================================================== #
# TAB: Generate
# =========================================================================== #
with tab_gen:
    st.caption(
        "**Requirement + Inventory_Report** upload කරන්න — SKU_MASTER Google Sheet "
        "එකෙන් ගන්නවා → **VIP PICK** + **INDIA SO Pick**. "
        "Pick = Req Qty × HJ ÷ SAP, carton (Actual Qty) නොබෙදා."
    )
    req_df = inv_df = sku_df = None

    c1, c2 = st.columns(2)
    with c1:
        f_req = st.file_uploader("1️⃣ Requirement (Requament)", type=["xlsx", "xls"])
    with c2:
        f_inv = st.file_uploader("2️⃣ Inventory_Report", type=["xlsx", "xls"])
    if f_req:
        req_df = pd.read_excel(f_req)
    if f_inv:
        inv_df = pd.read_excel(f_inv)

    # SKU_MASTER always comes from the Google Sheet (managed in the 🗂️ SKU_MASTER tab)
    sku_ready = False
    if sa_info and save_key:
        try:
            import gsheet
            sku_df = gsheet.read_sku_master(sa_info, save_key, sku_ws)
            if sku_df is not None and len(sku_df):
                sku_ready = True
                st.caption(f"ℹ️ SKU_MASTER — Google Sheet එකෙන් ({len(sku_df)} rows).")
            else:
                st.warning("SKU_MASTER හිස් — **🗂️ SKU_MASTER** tab එකෙන් upload කරලා save කරන්න.")
        except Exception as ex:
            st.warning(f"SKU_MASTER load කරගන්න බැරි වුණා: {ex}")
    else:
        st.warning("Sidebar එකේ **Data Sheet key** + credentials දාලා, "
                   "**🗂️ SKU_MASTER** tab එකෙන් SKU_MASTER save කරන්න.")

    # ---- missing-material pre-check (notify before creating pick) ----
    blocking = False
    if req_df is not None and sku_ready:
        try:
            miss = E.missing_materials(req_df, sku_df)
        except Exception:
            miss = []
        if miss:
            blocking = True
            st.error(
                f"⛔ Requirement එකේ Material **{len(miss)}**ක් SKU_MASTER එකේ නෑ. "
                "Pick create කරන්න කලින් මේවා **🗂️ SKU_MASTER** tab එකෙන් add කරන්න:"
            )
            st.dataframe(pd.DataFrame({"Missing Material (SKU_MASTER එකට add කරන්න)": miss}),
                         hide_index=True, use_container_width=True,
                         height=min(340, 60 + 30 * len(miss)))

    ready = (req_df is not None and inv_df is not None and sku_ready and not blocking)
    if not ready:
        if not blocking:
            st.info("ⓘ Requirement + Inventory upload කරන්න · SKU_MASTER, Google Sheet එකේ තියෙන්න ඕන.")
    else:
        if st.button("🚀 Generate Pick Files", type="primary", use_container_width=True):
            try:
                existing = None
                if use_registry and sa_info and save_key:
                    import gsheet
                    if hasattr(gsheet, "read_load_id_registry"):
                        try:
                            existing = gsheet.read_load_id_registry(sa_info, save_key)
                        except Exception as ex:
                            st.warning(f"LOAD_ID registry කියවන්න බැරි වුණා ({ex}) — duplicate check skip කළා.")
                    else:
                        st.warning("gsheet.py පරණ version එකක් — latest zip එක re-download කරන්න (registry check skip කළා).")
                with st.spinner("Processing..."):
                    res = E.run_pipeline(req_df, sku_df, inv_df, cfg, existing_load_ids=existing)
                st.session_state["result"] = res
                st.success("Generate වුණා ✅")
            except Exception as ex:
                st.error(f"Error: {ex}")
                res = None

            if st.session_state.get("result") and autosave and sa_info and save_key:
                try:
                    import gsheet
                    with st.spinner("Google Sheet එකට save කරනවා..."):
                        url = gsheet.save_all(sa_info, save_key,
                                              st.session_state["result"], cfg,
                                              source_label=source)
                    st.success("Auto-save + history + registry update වුණා ✅")
                    st.markdown(f"🔗 [Google Sheet එක open කරන්න]({url})")
                except Exception as ex:
                    st.error(f"Auto-save error: {ex}")
            elif autosave and not (sa_info and save_key):
                st.info("Auto-save on — හැබැයි credentials/Sheet key එක නෑ.")

        res = st.session_state.get("result")
        if res:
            pick = res["pick"]
            summary = res["summary"]
            exc = res["exceptions"]

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Pickable lines", len(pick))
            m2.metric("Deliveries", pick["OBD"].nunique() if len(pick) else 0)
            m3.metric("Total Boxes", int(pick["HJ Box Qty"].sum()) if len(pick) else 0)
            m4.metric("Total CBM", summary["Total CBM Of Pick"])
            m5.metric("Cannot Pick", len(exc), delta=None,
                      delta_color="inverse" if len(exc) else "off")

            st.subheader("⏱️ CBM / Pick-Time Summary")
            st.dataframe(pd.DataFrame([summary]), hide_index=True, use_container_width=True)

            if len(exc):
                short = exc[exc["Issue Type"].astype(str).str.contains("Inventory")]
                split = exc[exc["Issue Type"].astype(str).str.contains("Carton")]
                tot_var = int(exc["Variance"].astype(int).sum()) if "Variance" in exc.columns else 0
                with st.expander(f"⚠️ Cannot Pick report ({len(exc)}) — "
                                 f"carton-split {len(split)} · stock-short {len(short)} · "
                                 f"total variance {tot_var} pcs", expanded=True):
                    st.caption("Picked Pcs = whole cartons පිකියද · Variance = pick කරන්න බැරි ඉතුරු pcs.")
                    st.dataframe(exc, hide_index=True, use_container_width=True)

            t1, t2, t3, t4 = st.tabs(
                ["📋 VIP PICK", "🧾 INDIA SO — MASTER", "🧾 INDIA SO — Detail", "🏷️ LOAD ID QR"]
            )
            with t1:
                st.dataframe(res["vip_table"], hide_index=True, use_container_width=True, height=420)
            with t2:
                st.dataframe(res["india_master"], hide_index=True, use_container_width=True, height=420)
            with t3:
                st.dataframe(res["india_detail"], hide_index=True, use_container_width=True, height=420)
            with t4:
                suffixed = [str(x) for x in res["load_ids"] if "-" in str(x)]
                st.caption(
                    f"{len(res['load_ids'])} unique LOAD IDs · suffixed (duplicate-fixed): "
                    f"{len(suffixed)} · header QR: {', '.join(cfg.header_qr_codes)}."
                )
                st.dataframe(pd.DataFrame({"LOAD ID": res["load_ids"]}),
                             hide_index=True, use_container_width=True, height=420)

            st.divider()
            st.subheader("⬇️ Downloads")
            stamp = cfg.pick_date.strftime("%Y%m%d")
            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "📦 VIP PICK download", data=res["vip_pick_bytes"],
                    file_name=f"VIP_PICK_{stamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with d2:
                st.download_button(
                    "🧾 INDIA SO Pick download", data=res["india_so_bytes"],
                    file_name=f"INDIA_SO_Pick_{stamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

            if sa_info and save_key and not autosave:
                if st.button("📝 හැම data එකම Google Sheet එකට save කරන්න (+history +registry)"):
                    try:
                        import gsheet
                        url = gsheet.save_all(sa_info, save_key, res, cfg, source_label=source)
                        st.success("Save වුණා ✅")
                        st.markdown(f"🔗 [Google Sheet එක open කරන්න]({url})")
                    except Exception as ex:
                        st.error(f"Save error: {ex}")

st.caption(
    "Logic: pick = Req Qty × HJ ÷ SAP · carton (Actual Qty) නොබෙදා · "
    "බෙදෙන්නේ නැති / stock මදි lines → Cannot Pick report · "
    "LOAD ID duplicate → -A/-B suffix."
)
