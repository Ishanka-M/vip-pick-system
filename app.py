"""
app.py — Donaldson OutBound Pick Generator (EFL / Körber One)
============================================================
Invoice / Delivery Challan PDF + Inventory Report
        -> base-ID item match -> plant confirm -> pallet-level pick
        -> Google Sheet ledger  +  "OutBound MASTER / OutBound Detail" Excel

Run:    streamlit run app.py
"""
from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

import doc_parser as P
import pick_engine as E
import pick_pdf as PP

st.set_page_config(page_title="Donaldson OutBound Pick", page_icon="📦", layout="wide")

# --------------------------------------------------------------------------- #
# Theme
# --------------------------------------------------------------------------- #
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');
:root{
  --bg:#070F1A; --bg2:#0B1626; --panel:#0F1F33; --panel2:#13253B;
  --line:#22374F; --text:#EAF1F9; --muted:#9DB0C9; --faint:#7C90AB;
  --accent:#FF365B; --accent-soft:rgba(255,54,91,.14);
  --ok:#34D399; --warn:#F2B33D;
}
html, body, [class*="css"]{ font-family:'Inter',system-ui,sans-serif; }
.stApp{
  background:
   radial-gradient(1100px 480px at 100% -8%, #11223a 0%, rgba(17,34,58,0) 62%),
   radial-gradient(900px 420px at -5% 0%, #0c1a2c 0%, rgba(12,26,44,0) 55%),
   var(--bg);
  color:var(--text);
}
.block-container{ padding-top:1.1rem; max-width:1400px; }
.stApp, .stMarkdown, p, span, label, li, .stCaption,
[data-testid="stWidgetLabel"] label, [data-testid="stWidgetLabel"] p{ color:var(--text); }
small, .stCaption, [data-testid="stCaptionContainer"]{ color:var(--muted) !important; }
h1,h2,h3,h4{ font-family:'Space Grotesk','Inter',sans-serif; color:var(--text);
  letter-spacing:-.01em; font-weight:700; }
h2{ font-size:1.25rem; } h3{ font-size:1.05rem; }
.appbar{ display:flex; align-items:center; justify-content:space-between;
  gap:1rem; padding:18px 22px; border-radius:16px; margin-bottom:10px;
  background:linear-gradient(120deg,#0A1730 0%,#10243E 55%,#15304F 100%);
  border:1px solid #1C3450; box-shadow:0 16px 40px -18px rgba(0,0,0,.7); color:#fff; }
.appbar .brand{ display:flex; align-items:center; gap:14px; }
.appbar .mark{ width:42px; height:42px; border-radius:11px; display:grid;
  place-items:center; background:var(--accent); color:#fff; font-size:20px;
  box-shadow:0 8px 20px -7px rgba(255,54,91,.65); }
.appbar .title{ font-family:'Space Grotesk',sans-serif; font-weight:700;
  font-size:1.32rem; line-height:1.1; color:#fff; }
.appbar .sub{ font-size:.78rem; color:#9FB4D0; margin-top:3px;
  font-family:'IBM Plex Mono',monospace; }
.appbar .pill{ font-family:'IBM Plex Mono',monospace; font-size:.72rem;
  font-weight:600; padding:7px 13px; border-radius:999px;
  border:1px solid rgba(255,255,255,.14); background:rgba(255,255,255,.06);
  color:#dbe6f3; display:inline-flex; align-items:center; gap:7px; white-space:nowrap;}
.appbar .dot{ width:8px; height:8px; border-radius:50%; }
.dot-on{ background:#43D39E; box-shadow:0 0 0 3px rgba(67,211,158,.22);}
.dot-off{ background:#F2B33D; box-shadow:0 0 0 3px rgba(242,179,61,.20);}
[data-testid="stTabs"] [role="tablist"]{ gap:6px; border-bottom:1px solid var(--line);}
[data-testid="stTabs"] [role="tab"]{ font-family:'Space Grotesk',sans-serif;
  font-weight:600; font-size:.92rem; color:var(--muted); padding:8px 16px;
  border-radius:9px 9px 0 0; }
[data-testid="stTabs"] [role="tab"] p{ color:var(--muted); font-weight:600; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"],
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p{ color:#fff; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{ background:var(--accent-soft); }
[data-testid="stMetric"]{ background:var(--panel); border:1px solid var(--line);
  border-radius:14px; padding:14px 16px; }
[data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p{ color:var(--muted);
  font-weight:600; font-size:.74rem; text-transform:uppercase; letter-spacing:.06em; }
[data-testid="stMetricValue"]{ font-family:'IBM Plex Mono',monospace; color:var(--text); }
.stButton>button, .stDownloadButton>button{ font-family:'Space Grotesk',sans-serif;
  font-weight:600; border-radius:10px; border:1px solid var(--line);
  background:var(--panel2); color:var(--text); padding:.5rem 1rem; }
.stButton>button:hover, .stDownloadButton>button:hover{ border-color:#33506f; }
.stButton>button[kind="primary"], .stDownloadButton>button{
  background:var(--accent); border-color:var(--accent); color:#fff; }
[data-testid="stSidebar"]{ background:var(--bg2); border-right:1px solid var(--line); }
[data-testid="stSidebar"] *{ color:var(--text); }
.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea{
  background:var(--panel); color:var(--text); border:1px solid var(--line); border-radius:9px;}
[data-baseweb="select"]>div{ background:var(--panel); border-color:var(--line); color:var(--text);}
[data-testid="stFileUploaderDropzone"]{ background:var(--panel);
  border:1px dashed #2C466A; border-radius:12px; }
[data-testid="stFileUploaderDropzone"] *{ color:var(--muted); }
[data-testid="stExpander"]{ border:1px solid var(--line); border-radius:12px;
  background:var(--panel); }
[data-testid="stDataFrame"]{ border:1px solid var(--line); border-radius:12px; }
.docok{ border-left:3px solid var(--ok); padding:8px 12px; border-radius:8px;
  background:rgba(52,211,153,.08); margin-bottom:6px; font-size:.9rem;}
.docbad{ border-left:3px solid var(--accent); padding:8px 12px; border-radius:8px;
  background:var(--accent-soft); margin-bottom:6px; font-size:.9rem;}
.footnote{ color:var(--faint); font-size:.78rem; font-family:'IBM Plex Mono',monospace;
  border-top:1px solid var(--line); padding-top:12px; margin-top:18px; }
</style>
"""
st.markdown(_CSS, unsafe_allow_html=True)


def appbar(connected: bool, detail: str = "") -> None:
    dot = "dot-on" if connected else "dot-off"
    status = "Sheet connected" if connected else "Sheet not set"
    st.markdown(f"""
<div class="appbar">
  <div class="brand">
    <div class="mark">📦</div>
    <div>
      <div class="title">Donaldson &middot; OutBound Pick Generator</div>
      <div class="sub">Invoice / Delivery Challan → Pallet Pick → Körber One upload · {detail}</div>
    </div>
  </div>
  <div class="pill"><span class="dot {dot}"></span>{status}</div>
</div>""", unsafe_allow_html=True)


def get_sa():
    try:
        return st.secrets.get("gcp_service_account", None)
    except Exception:
        return None


def gs_conf() -> dict:
    try:
        return dict(st.secrets.get("google_sheet", {}))
    except Exception:
        return {}


def _reset_pw() -> str:
    """Password එක secrets එකේ තිබ්බොත් ඒක — නැත්නම් default."""
    try:
        return str(st.secrets.get("app", {}).get("reset_password", "") or "Isha@1996")
    except Exception:
        return "Isha@1996"


RESET_PASSWORD = _reset_pw()


# --------------------------------------------------------------------------- #
# Sidebar
# --------------------------------------------------------------------------- #
sa_info = get_sa()
conf = gs_conf()
sheet_key = str(conf.get("data_sheet", "")).strip()
autosave = bool(conf.get("auto_save", True))
gs_ready = bool(sa_info and sheet_key)

with st.sidebar:
    st.header("⚙️ Settings")

    with st.expander("WMS constants", expanded=True):
        wh_id = st.text_input("WH_ID", value=str(conf.get("wh_id", "INMM01")))
        client_code = st.text_input("CLIENT_CODE", value=str(conf.get("client_code", "INM0DONA")))
        order_type = st.text_input("ORDER_TYPE", value="Sales Orders")

    with st.expander("Pick options", expanded=True):
        strategy = st.selectbox(
            "Pick strategy",
            options=["FIFO", "SINGLE_PALLET_FIRST", "LEAST_PALLETS"],
            format_func=lambda s: {
                "FIFO": "FIFO — පරණ stock මුලින්",
                "SINGLE_PALLET_FIRST": "Single pallet first — හැකි නම් එක pallet එකෙන්",
                "LEAST_PALLETS": "Least pallets — ලොකු pallet මුලින්",
            }[s],
        )
        statuses = st.multiselect("Inventory Status", ["Available", "Hold", "Damage", "QC"],
                                  default=["Available"])
        exact_first = st.checkbox("Exact item number එකට මුල් තැන", value=True,
                                  help="Base ID එක ගැලපුනත්, document එකේ තියෙන "
                                       "full item number එකට මුලින් priority.")
        use_ledger = st.checkbox("Pallet ledger balance logic", value=True,
                                 help="Pallet එකේ Actual Qty == ledger QTY_BEFORE නම් → "
                                      "QTY_BALANCE එකෙන් pick. වෙනස් නම් → Inventory "
                                      "Actual Qty එක අලුත් QTY_BEFORE එක විදිහට අරගෙන pick.")
        blank_fill = st.text_input("හිස් attribute වලට දාන value", value="TBC")
        fill_item_col = st.checkbox("ITEM_NUMBER column එකත් පුරවන්න", value=False)
        merge_lines = st.checkbox("එකම item එකේ lines merge කරන්න", value=False)
        override = st.checkbox("⚠️ Document check bypass (manual verify කළා)", value=False,
                               help="Total Quantity / Grand Total / S.No check එක fail වුණත් "
                                    "pick කරනවා. Stock check එක bypass වෙන්නේ නෑ. "
                                    "Registry එකේ 'MANUAL OVERRIDE' කියලා log වෙනවා.")
        if override:
            st.warning("Document completeness gate off — parse එක ඇස්සෙන් check කරන්න.")
        pick_date = st.date_input("Pick Date", value=datetime.now())

    with st.expander("📧 Email settings", expanded=False):
        book = st.session_state.get("mail_book")
        if book is None:
            book = str(conf.get("mail_to", ""))
            if gs_ready:
                try:
                    import gsheet
                    book = gsheet.read_setting(sa_info, sheet_key, "MAIL_TO", book) or book
                except Exception:
                    pass
            st.session_state["mail_book"] = book

        saved = PP._addr_list(st.session_state.get("mail_book", ""))
        mail_to = st.multiselect("To (save කරපු addresses)", options=saved, default=saved)
        extra_to = st.text_input("තව address (comma වලින්)", key="mail_extra")
        mail_to = list(mail_to) + PP._addr_list(extra_to)
        mail_cc = st.text_input("Cc", value=str(conf.get("mail_cc", "")))
        mail_from = st.text_input("From (ඔයාගේ mail)", value=str(conf.get("mail_from", "")))
        mail_sign = st.text_area("Signature", value=str(conf.get("mail_sign",
                                                                "Thanks & regards,")),
                                 height=70)

        new_addr = st.text_input("➕ Address book එකට add කරන්න")
        ab1, ab2 = st.columns(2)
        if ab1.button("Add", use_container_width=True) and PP._addr_list(new_addr):
            merged = PP._addr_list(st.session_state.get("mail_book", "")) \
                + PP._addr_list(new_addr)
            st.session_state["mail_book"] = ", ".join(dict.fromkeys(merged))
            if gs_ready:
                try:
                    import gsheet
                    gsheet.save_setting(sa_info, sheet_key, "MAIL_TO",
                                        st.session_state["mail_book"])
                except Exception as ex:
                    st.warning(f"Sheet save error: {ex}")
            st.rerun()
        if ab2.button("Clear book", use_container_width=True):
            st.session_state["mail_book"] = ""
            if gs_ready:
                try:
                    import gsheet
                    gsheet.save_setting(sa_info, sheet_key, "MAIL_TO", "")
                except Exception:
                    pass
            st.rerun()
        if saved:
            st.caption("Book: " + ", ".join(saved))

    st.divider()
    st.subheader("💾 Google Sheet")
    if gs_ready:
        st.success("✅ secrets වලින් connected")
        st.caption(f"auto-save: {'on' if autosave else 'off'}")
        if st.button("🆕 Initialize worksheets"):
            try:
                import gsheet
                r = gsheet.init_sheet(sa_info, sheet_key)
                new = [k for k, v in r.items() if v is True]
                st.success("Init වුණා ✅ " + (f"created: {', '.join(new)}" if new
                                             else "ඔක්කොම දැනටමත් තිබුණා"))
                st.markdown(f"🔗 [Sheet එක open කරන්න]({r.get('url','')})")
            except Exception as ex:
                st.error(f"Init error: {ex}")
    elif not sa_info:
        st.warning("`[gcp_service_account]` secret නෑ — download විතරක් වැඩ කරයි.")
    else:
        st.warning("`[google_sheet] data_sheet` secret නෑ.")

    with st.expander("🧹 Reset / Undo"):
        undo = st.text_input("RUN_ID එකක් undo කරන්න")
        if st.button("↩️ Run එක delete කරන්න") and gs_ready and undo.strip():
            try:
                import gsheet
                r = gsheet.delete_run(sa_info, sheet_key, undo.strip())
                st.success(f"Delete වුණා ✅ {r}")
            except Exception as ex:
                st.error(f"Undo error: {ex}")

        st.markdown("---")
        st.markdown("**🔐 DB Reset**")
        if not st.session_state.get("reset_ok"):
            pw = st.text_input("Password", type="password", key="reset_pw")
            if st.button("🔓 Unlock", use_container_width=True):
                if pw == RESET_PASSWORD:
                    st.session_state["reset_ok"] = True
                    st.rerun()
                else:
                    st.error("Password වැරදියි.")
        else:
            st.success("🔓 Unlocked")
            scope = st.multiselect(
                "Clear කරන්නේ",
                ["outputs", "ledger", "registry", "rejected", "runlog", "settings"],
                default=["outputs", "ledger", "registry", "rejected", "runlog"],
                format_func=lambda s: {
                    "outputs": "OUTBOUND_MASTER + DETAIL",
                    "ledger": "PALLET_LEDGER (pallet balance!)",
                    "registry": "DOC_REGISTRY (duplicate gate!)",
                    "rejected": "REJECTED_LOG",
                    "runlog": "RUN_LOG",
                    "settings": "APP_SETTINGS (email book)",
                }[s],
            )
            sure = st.checkbox("මට විශ්වාසයි — back ගන්න බෑ", key="reset_sure")
            r1, r2 = st.columns(2)
            if r1.button("🗑️ Reset", use_container_width=True):
                if not gs_ready:
                    st.error("Google Sheet connect වෙලා නෑ.")
                elif not scope or not sure:
                    st.warning("Scope select කරලා confirm කරන්න.")
                else:
                    try:
                        import gsheet
                        done = gsheet.reset_data(sa_info, sheet_key, scope)
                        for k in ("result", "bundles", "bundle_key", "zipfile", "hist"):
                            st.session_state.pop(k, None)
                        st.success(f"Reset වුණා ✅ {done}")
                    except Exception as ex:
                        st.error(f"Reset error: {ex}")
            if r2.button("💣 FULL DB RESET", use_container_width=True, type="primary"):
                if not gs_ready:
                    st.error("Google Sheet connect වෙලා නෑ.")
                elif not sure:
                    st.warning("Confirm checkbox එක tick කරන්න.")
                else:
                    try:
                        import gsheet
                        r = gsheet.reset_all(sa_info, sheet_key, keep_settings=True)
                        for k in ("result", "bundles", "bundle_key", "zipfile", "hist"):
                            st.session_state.pop(k, None)
                        st.success(f"FULL RESET ✅ {r['count']} worksheets — "
                                   f"{', '.join(r['cleared'])}")
                        st.caption("Email book (APP_SETTINGS) එක ඉතුරු කළා.")
                    except Exception as ex:
                        st.error(f"Reset error: {ex}")
            if st.button("🔒 Lock ආපහු"):
                st.session_state.pop("reset_ok", None)
                st.rerun()

    if st.button("🔄 Session clear"):
        for k in list(st.session_state.keys()):
            st.session_state.pop(k, None)
        st.rerun()

appbar(gs_ready, "Körber One · INM0DONA" if gs_ready else "download-only mode")

tab_gen, tab_search, tab_bal, tab_hist, tab_help = st.tabs(
    ["🚀 Pick Generate", "🔎 Search", "📦 Pallet Balance", "📜 History", "📘 Guide"]
)

# =========================================================================== #
# TAB 1 — Generate
# =========================================================================== #
with tab_gen:
    st.caption(
        "**1** Invoice / Delivery Challan PDF + Inventory Report upload → "
        "**2** parse කරපු lines confirm → **3** Plant confirm → **4** pick + Excel."
    )

    c1, c2 = st.columns(2)
    with c1:
        f_docs = st.file_uploader("1️⃣ Invoice / Delivery Challan (PDF · multiple)",
                                  type=["pdf"], accept_multiple_files=True)
    with c2:
        f_inv = st.file_uploader("2️⃣ Inventory Report (Excel)", type=["xlsx", "xls"])

    # ---------------- parse documents ---------------- #
    docs: list[P.ParsedDoc] = []
    if f_docs:
        sig = tuple(sorted((f.name, f.size) for f in f_docs))
        if st.session_state.get("doc_sig") != sig:
            parsed = []
            raw: dict[str, bytes] = {}
            with st.spinner("PDF read කරනවා..."):
                for f in f_docs:
                    try:
                        data = f.getvalue()
                        raw[f.name] = data
                        parsed.append(P.parse_pdf(data, f.name))
                    except Exception as ex:
                        st.error(f"{f.name} — parse error: {ex}")
            st.session_state["doc_bytes"] = raw
            st.session_state["doc_sig"] = sig
            st.session_state["docs"] = parsed
            st.session_state["doc_frame"] = P.docs_to_frame(parsed)
            st.session_state.pop("result", None)
        docs = st.session_state.get("docs", [])

    if docs:
        st.markdown("### 📄 Parse කරපු documents")
        seen: set[str] = set()
        for d in docs:
            ok, probs = d.completeness()
            dup = d.doc_number in seen
            seen.add(d.doc_number)
            qty = sum(l.qty for l in d.lines)
            head = (f"<b>{d.doc_type}</b> · <code>{d.doc_number or '???'}</code> · "
                    f"{d.doc_date} · {len(d.lines)} lines · Qty {qty:g} "
                    f"<span style='color:#7C90AB'>({d.source_file})</span>")
            if dup:
                st.markdown(f"<div class='docbad'>{head}<br>⚠️ Duplicate number — "
                            f"පළවෙනි එක විතරක් process වෙනවා</div>", unsafe_allow_html=True)
            elif ok:
                st.markdown(f"<div class='docok'>✅ {head}</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='docbad'>⛔ {head}<br>" +
                            "<br>".join("• " + p for p in probs) + "</div>",
                            unsafe_allow_html=True)

        with st.expander("✏️ Lines review / edit (parse වැරදුනොත් මෙතන හදන්න)"):
            st.caption("Base ID = match කරන ID එක · row එකක් අයින් කරන්න `Use` uncheck කරන්න.")
            edited = st.data_editor(
                st.session_state.get("doc_frame", pd.DataFrame()),
                num_rows="dynamic", use_container_width=True, height=320,
                key="doc_editor",
                column_config={
                    "Use": st.column_config.CheckboxColumn("Use", default=True),
                    "Qty": st.column_config.NumberColumn("Qty", min_value=0, step=1),
                    "Base ID": st.column_config.TextColumn("Base ID", disabled=True),
                },
            )
            if st.button("💡 Edit කරපු lines apply කරන්න"):
                st.session_state["docs"] = P.frame_to_docs(edited, docs)
                st.session_state["doc_frame"] = P.docs_to_frame(st.session_state["docs"])
                st.session_state.pop("result", None)
                st.success("Apply වුණා ✅")
                st.rerun()

    # ---------------- inventory + plant confirm ---------------- #
    inv_raw = None
    if f_inv is not None:
        if st.session_state.get("inv_sig") != (f_inv.name, f_inv.size):
            with st.spinner("Inventory read කරනවා..."):
                st.session_state["inv_raw"] = pd.read_excel(f_inv, dtype=str)
            st.session_state["inv_sig"] = (f_inv.name, f_inv.size)
            st.session_state.pop("plants_ok", None)
            st.session_state.pop("result", None)
        inv_raw = st.session_state.get("inv_raw")

    if inv_raw is not None:
        inv_norm = E.normalize_inventory(inv_raw)
        psum = E.plant_summary(inv_norm)
        st.markdown("### 🏭 Plant confirmation")
        st.caption("Inventory එකේ තියෙන plant ටික මේවා — **මොන plant එකෙන්ද pick කරන්නේ?**")
        pc1, pc2 = st.columns([1.2, 1])
        with pc1:
            st.dataframe(psum, hide_index=True, use_container_width=True)
        with pc2:
            choice = st.multiselect(
                "Plant(s)", options=psum["Plant"].tolist(),
                default=st.session_state.get("plants_ok", []),
            )
            if st.button("✅ මේ plant එකෙන් pick කරන්න — confirm", type="primary",
                         use_container_width=True):
                if not choice:
                    st.warning("Plant එකක් හරි select කරන්න.")
                else:
                    st.session_state["plants_ok"] = choice
                    st.session_state.pop("result", None)
                    st.success(f"Confirm වුණා ✅ — {', '.join(choice)}")
        if st.session_state.get("plants_ok"):
            st.info(f"🏭 Confirmed plant: **{', '.join(st.session_state['plants_ok'])}**")

    # ---------------- generate ---------------- #
    ready = bool(st.session_state.get("docs")) and inv_raw is not None \
        and bool(st.session_state.get("plants_ok"))

    if not ready:
        st.info("ⓘ Document PDF + Inventory upload කරලා, Plant එක confirm කරන්න.")
    else:
        cfg = E.EngineConfig(
            wh_id=wh_id, client_code=client_code, order_type=order_type,
            plants=st.session_state["plants_ok"], statuses=statuses, strategy=strategy,
            exact_item_first=exact_first, use_ledger=use_ledger, blank_fill=blank_fill,
            fill_item_number_col=fill_item_col, merge_same_item_lines=merge_lines,
            override_doc_check=override,
            pick_date=datetime.combine(pick_date, datetime.now().time()),
        )
        note = st.text_input("Run note (optional)", placeholder="උදා: 13-Aug morning batch")

        if st.button("🚀 Pick generate කරන්න", type="primary", use_container_width=True):
            ledger = None
            done_docs: set[str] = set()
            if gs_ready:
                try:
                    import gsheet
                    ledger = gsheet.read_ledger(sa_info, sheet_key)
                    done_docs = gsheet.read_processed_docs(sa_info, sheet_key)
                except Exception as ex:
                    st.warning(f"Google Sheet read කරන්න බැරි වුණා ({ex}) — "
                               "ledger/duplicate check skip කළා.")
            try:
                with st.spinner("Pick calculate කරනවා..."):
                    res = E.run_pick(st.session_state["docs"], inv_raw, cfg,
                                     ledger=ledger, processed_docs=done_docs)
                res["note"] = note
                st.session_state["result"] = res
                st.success(f"Generate වුණා ✅  RUN_ID `{res['run_id']}`")
            except Exception as ex:
                st.error(f"Engine error: {ex}")
                st.session_state.pop("result", None)

            res = st.session_state.get("result")
            if res is not None and gs_ready and autosave and len(res["master"]):
                try:
                    import gsheet
                    with st.spinner("Google Sheet එකට save කරනවා..."):
                        r = gsheet.save_run(sa_info, sheet_key, res, cfg, note=note)
                    st.success(f"Sheet save වුණා ✅ master {r['master']} · detail {r['detail']} "
                               f"· ledger {r['ledger']} rows")
                    st.markdown(f"🔗 [Sheet එක open කරන්න]({r['url']})")
                    st.session_state["saved"] = res["run_id"]
                except Exception as ex:
                    st.error(f"Save error: {ex}")

    # ---------------- results ---------------- #
    res = st.session_state.get("result")
    if res:
        st.divider()
        acc, rej = res["accepted"], res["rejected"]
        alloc = res["allocations"]

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Docs picked", len(acc))
        m2.metric("Rejected", len(rej), delta_color="inverse")
        m3.metric("Detail lines", len(res["detail"]))
        m4.metric("Pallets", int(alloc["PALLET"].nunique()) if len(alloc) else 0)
        m5.metric("Total Qty",
                  f"{pd.to_numeric(alloc['QTY_PICKED'], errors='coerce').sum():g}"
                  if len(alloc) else "0")

        vdf = res.get("verify", pd.DataFrame())
        bad_v = vdf[vdf["STATUS"].astype(str).str.contains("MISMATCH")] if len(vdf) else vdf
        if len(vdf) and not len(bad_v):
            st.success("🔢 Quantity verify — Invoice / DC qty එකට **හරියටම** ගැලපෙනවා "
                       "(line · document total · WMS file total).")
        elif len(bad_v):
            st.error("🔢 Quantity mismatch! මේ documents pick කළේ නෑ:")
            st.dataframe(bad_v, hide_index=True, use_container_width=True)

        if len(rej):
            with st.expander(f"⛔ Pick කරන්න බැරි වුණ documents ({len(rej)})", expanded=True):
                st.dataframe(rej, hide_index=True, use_container_width=True)
                if len(res["shortage"]):
                    st.caption("Stock short lines:")
                    st.dataframe(res["shortage"], hide_index=True, use_container_width=True)

        t1, t2, t3, t4, t7, t5, t6 = st.tabs(
            ["🧾 OutBound MASTER", "📋 OutBound Detail", "🎯 Pallet Allocation",
             "📦 Pallet Balance", "📊 Stock Basis", "🔢 Qty Verify", "✅ Doc Summary"])
        with t1:
            st.dataframe(res["master"], hide_index=True, use_container_width=True, height=280)
        with t2:
            st.dataframe(res["detail"], hide_index=True, use_container_width=True, height=380)
        with t3:
            st.caption("කොයි pallet එකෙන් කීයද ගත්තේ — balance එකත් එක්කම.")
            st.dataframe(alloc, hide_index=True, use_container_width=True, height=380)
        with t4:
            st.caption("QTY_BEFORE → QTY_PICKED → QTY_BALANCE. **MODE** එකෙන් පේනවා "
                       "ledger balance එකෙන්ද, අලුත් inventory qty එකෙන්ද pick කරේ කියලා.")
            st.dataframe(res["balance"], hide_index=True, use_container_width=True, height=380)
        with t7:
            bs = res.get("basis", pd.DataFrame())
            st.caption("Pallet එකකට pick කරන්න පුළුවන් උපරිමය කොහොමද තීරණය වුණේ කියලා.")
            if len(bs):
                mc = bs["MODE"].value_counts()
                b1, b2, b3 = st.columns(3)
                b1.metric("NEW (ledger නෑ)", int(mc.get("NEW", 0)))
                b2.metric("LEDGER BALANCE", int(mc.get("LEDGER BALANCE", 0)),
                          help="Inventory එක refresh වෙලා නෑ — ledger balance එකෙන් pick")
                b3.metric("NEW BASELINE", int(mc.get("NEW BASELINE", 0)),
                          help="Inventory Actual Qty වෙනස් වෙලා — ඒක අලුත් QTY_BEFORE")
                only = st.checkbox("Ledger එකේ තියෙන pallet විතරක්", value=True,
                                   key="basis_filter")
                view = bs[bs["MODE"] != "NEW"] if only else bs
                st.dataframe(view, hide_index=True, use_container_width=True, height=380)
            else:
                st.info("Basis data නෑ.")
        with t5:
            st.caption("Invoice / DC එකේ Quantity එකට **හරියටම** ගැලපෙනවද — line by line, "
                       "document total, WMS file total.")
            st.dataframe(res["verify"], hide_index=True, use_container_width=True, height=380)
        with t6:
            st.dataframe(acc, hide_index=True, use_container_width=True, height=280)

        # ------------------------------------------------------------------ #
        # Downloads — LOAD_ID එකෙන් save
        # ------------------------------------------------------------------ #
        st.divider()
        st.subheader("⬇️ Downloads — LOAD_ID එකෙන්")
        ids = E.load_ids(res)
        src_map = st.session_state.get("doc_bytes", {})

        if not ids:
            st.warning("Pick වුණ document නෑ — download කරන්න දෙයක් නෑ.")
        else:
            attach_src = st.checkbox("PDF එකට upload කරපු Invoice / DC pages එකතු කරන්න",
                                     value=True)

            cache_key = (res["run_id"], bool(attach_src))
            if st.session_state.get("bundle_key") != cache_key:
                with st.spinner("PDF + Excel හදනවා..."):
                    made = {}
                    zfiles: list[tuple[str, bytes]] = []
                    for lid in ids:
                        b = E.doc_bundle(res, lid)
                        b["xlsx"] = E.build_wms_excel(b["master"], b["detail"])
                        b["pdf"] = PP.build_doc_pdf(
                            b["info"], b["allocations"], b["verify"],
                            src_map.get(b["info"].get("SOURCE_FILE", "")),
                            attach_source=attach_src)
                        made[lid] = b
                        zfiles += [(f"{b['safe']}.xlsx", b["xlsx"]),
                                   (f"{b['safe']}.pdf", b["pdf"])]
                    st.session_state["bundles"] = made
                    st.session_state["zipfile"] = E.build_zip(zfiles)
                    st.session_state["bundle_key"] = cache_key
            bundles = st.session_state["bundles"]

            for lid in ids:
                b = bundles[lid]
                with st.container(border=True):
                    h1, h2, h3 = st.columns([2, 1, 1])
                    i = b["info"]
                    h1.markdown(f"**LOAD ID `{lid}`** · {i.get('DOC_TYPE','')} · "
                                f"{i.get('LINES','')} lines · Qty "
                                f"{E._qty_str(float(i.get('TOTAL_QTY') or 0))} · "
                                f"{i.get('PALLETS','')} pallets · {i.get('VERIFY','')}")
                    if b["safe"] != lid:
                        h1.caption(f"File name: `{b['safe']}` (`/` filename වලට දාන්න බෑ)")
                    h2.download_button("📥 Excel", data=b["xlsx"],
                                       file_name=f"{b['safe']}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument."
                                            "spreadsheetml.sheet",
                                       use_container_width=True, key=f"x_{b['safe']}")
                    h3.download_button("🏷️ PDF + QR", data=b["pdf"],
                                       file_name=f"{b['safe']}.pdf",
                                       mime="application/pdf", use_container_width=True,
                                       key=f"p_{b['safe']}")

            st.markdown("**සියල්ලම එකට**")
            z1, z2, z3 = st.columns(3)
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            z1.download_button("🗜️ ZIP (හැම LOAD_ID එකකටම Excel + PDF)",
                               data=st.session_state["zipfile"],
                               file_name=f"OutBound_{stamp}.zip",
                               mime="application/zip", use_container_width=True)
            z2.download_button("📚 එකම Excel එකක (ඔක්කොම docs)",
                               data=E.build_wms_excel(res["master"], res["detail"]),
                               file_name=f"OutBound_Upload_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet", use_container_width=True)
            z3.download_button("📊 Pick Report", data=E.build_report_excel(res),
                               file_name=f"Pick_Report_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet", use_container_width=True)

            # -------------------------------------------------------------- #
            # Email
            # -------------------------------------------------------------- #
            st.divider()
            st.subheader("📧 Email — pick details")
            pick_for_mail = st.multiselect("මොන LOAD_ID ද mail එකට", ids, default=ids)
            infos = [bundles[l]["info"] for l in pick_for_mail]
            m_alloc = alloc[alloc["DOC_NUMBER"].astype(str).isin(pick_for_mail)] \
                if len(alloc) else alloc

            if infos:
                subj0, body0, html0 = PP.pick_email_text(infos, m_alloc, mail_sign)
                e1, e2 = st.columns([2, 1])
                subject = e1.text_input("Subject", value=subj0, key="mail_subject")
                add_att = e2.checkbox("Excel + PDF attach කරන්න", value=True)
                body = st.text_area("Body", value=body0, height=240, key="mail_body")

                to_list = PP._addr_list(mail_to)
                cc_list = PP._addr_list(mail_cc)
                if not to_list:
                    st.warning("Sidebar → **📧 Email settings** එකේ 'To' address එකක් දාන්න.")
                st.caption(f"To: {', '.join(to_list) or '—'}"
                           + (f"  ·  Cc: {', '.join(cc_list)}" if cc_list else ""))

                atts: list[tuple[str, bytes, str]] = []
                if add_att:
                    for l in pick_for_mail:
                        b = bundles[l]
                        atts.append((f"{b['safe']}.xlsx", b["xlsx"],
                                     "application/vnd.openxmlformats-officedocument."
                                     "spreadsheetml.sheet"))
                        atts.append((f"{b['safe']}.pdf", b["pdf"], "application/pdf"))

                mc1, mc2 = st.columns(2)
                with mc1:
                    st.link_button("✉️ Default mail app එකෙන් open කරන්න",
                                   PP.mailto_link(to_list, subject, body, cc_list),
                                   use_container_width=True,
                                   disabled=not to_list)
                    st.caption("mailto: — attachment යන්නේ නෑ, body විතරයි.")
                with mc2:
                    st.download_button(
                        "📎 Draft (.eml) download — attachment එක්ක",
                        data=PP.build_eml(to_list, subject, body, html0, cc_list,
                                          sender=mail_from, attachments=atts),
                        file_name=f"PICK_{E.safe_name(pick_for_mail[0])}"
                                  f"{'_+' + str(len(pick_for_mail) - 1) if len(pick_for_mail) > 1 else ''}.eml",
                        mime="message/rfc822", use_container_width=True)
                    st.caption("Double-click කරාම Outlook/Mail එකේ draft එකක් විදිහට "
                               "attachment එක්කම open වෙනවා.")

        if gs_ready and not autosave and st.session_state.get("saved") != res["run_id"]:
            if st.button("📝 Google Sheet එකට save කරන්න"):
                try:
                    import gsheet
                    r = gsheet.save_run(sa_info, sheet_key, res, res["cfg"],
                                        note=res.get("note", ""))
                    st.session_state["saved"] = res["run_id"]
                    st.success("Save වුණා ✅")
                    st.markdown(f"🔗 [Sheet එක open කරන්න]({r['url']})")
                except Exception as ex:
                    st.error(f"Save error: {ex}")


# =========================================================================== #
# TAB — Global search
# =========================================================================== #
with tab_search:
    st.subheader("🔎 Search — ඕනෑම data එකක්")
    st.caption("Item code · LOAD ID · pallet · location · GRN · plant · lot — "
               "ඕන දෙයක් type කරන්න. Word කීපයක් දුන්නොත් ඔක්කොම තියෙන rows විතරයි.")

    q = st.text_input("🔍", placeholder="උදා:  P550945   ·   333262712337   ·   "
                                       "IMDS01 P502639   ·   DONAL130826",
                      label_visibility="collapsed", key="global_q")

    frames: dict[str, pd.DataFrame] = {}
    res_s = st.session_state.get("result")
    if st.session_state.get("doc_frame") is not None:
        frames["📄 Document lines"] = st.session_state["doc_frame"]
    if res_s:
        frames["🎯 Pallet Allocation"] = res_s["allocations"]
        frames["📋 OutBound Detail"] = res_s["detail"]
        frames["🧾 OutBound MASTER"] = res_s["master"]
        frames["🔢 Qty Verify"] = res_s.get("verify", pd.DataFrame())
        frames["⛔ Rejected"] = res_s["rejected"]

    src = st.multiselect(
        "කොහෙන්ද හොයන්නේ",
        options=["Current run", "Inventory (stock)", "Google Sheet — ledger",
                 "Google Sheet — registry", "Google Sheet — detail"],
        default=["Current run", "Inventory (stock)"],
    )

    if "Inventory (stock)" in src and st.session_state.get("inv_raw") is not None:
        led = None
        if gs_ready:
            try:
                import gsheet
                led = gsheet.read_ledger(sa_info, sheet_key)
            except Exception:
                led = None
        frames["📦 Inventory / balance"] = E.stock_view(st.session_state["inv_raw"], led,
                                                       use_ledger=use_ledger)
    if "Current run" not in src:
        for k in ["🎯 Pallet Allocation", "📋 OutBound Detail", "🧾 OutBound MASTER",
                  "🔢 Qty Verify", "📊 Stock Basis", "⛔ Rejected", "📄 Document lines"]:
            frames.pop(k, None)
    if gs_ready:
        try:
            import gsheet
            if "Google Sheet — ledger" in src:
                frames["📜 Sheet · PALLET_LEDGER"] = gsheet.read_ws(sa_info, sheet_key,
                                                                   gsheet.WS_LEDGER)
            if "Google Sheet — registry" in src:
                frames["📜 Sheet · DOC_REGISTRY"] = gsheet.read_ws(sa_info, sheet_key,
                                                                  gsheet.WS_REGISTRY)
            if "Google Sheet — detail" in src:
                frames["📜 Sheet · OUTBOUND_DETAIL"] = gsheet.read_ws(sa_info, sheet_key,
                                                                     gsheet.WS_DETAIL)
        except Exception as ex:
            st.warning(f"Google Sheet read error: {ex}")

    if not q.strip():
        st.info("ⓘ හොයන්න ඕන දේ type කරන්න.")
    elif not frames:
        st.warning("Search කරන්න data නෑ — Generate tab එකෙන් upload කරන්න.")
    else:
        hits = E.search_frames(q, frames)
        total = sum(len(v) for v in hits.values())
        if not total:
            st.warning(f"'{q}' — කිසිම තැනක නෑ.")
        else:
            st.success(f"**{total}** rows · {len(hits)} තැනක හම්බුණා")
            for name, df in hits.items():
                with st.expander(f"{name} — {len(df)} rows", expanded=len(hits) <= 2):
                    st.dataframe(df, hide_index=True, use_container_width=True,
                                 height=min(420, 60 + 32 * len(df)))
                    st.download_button("⬇️ CSV", data=df.to_csv(index=False).encode(),
                                       file_name=f"search_{E.safe_name(name)}.csv",
                                       mime="text/csv", key=f"dl_{name}")

# =========================================================================== #
# TAB — Pallet balance
# =========================================================================== #
with tab_bal:
    st.subheader("📦 Pallet-level balance")
    st.caption("Inventory එකේ තියෙන qty එකෙන් කලින් pick කරපු ප්‍රමාණය අඩු කරලා "
               "ඉතුරු balance එක.")
    inv_raw = st.session_state.get("inv_raw")
    if inv_raw is None:
        st.info("Generate tab එකෙන් Inventory Report එක upload කරන්න.")
    else:
        ledger = None
        if gs_ready:
            try:
                import gsheet
                ledger = gsheet.read_ledger(sa_info, sheet_key)
            except Exception as ex:
                st.warning(f"Ledger read error: {ex}")
        view = E.stock_view(inv_raw, ledger, use_ledger=use_ledger)
        f1, f2, f3, f4 = st.columns([1.4, 1, 1, 1])
        q = f1.text_input("Item / Base ID / Pallet search", placeholder="P550945")
        plants = f2.multiselect("Plant", sorted(view["PLANT"].dropna().unique().tolist()))
        modes = f3.multiselect("Mode", ["NEW", "LEDGER BALANCE", "NEW BASELINE"])
        only_bal = f4.checkbox("Balance > 0 විතරක්", value=True)

        v = view
        if q.strip():
            k = q.strip().upper()
            v = v[v["ITEM_NUMBER"].str.upper().str.contains(k, na=False)
                  | v["BASE_ID"].str.upper().str.contains(k, na=False)
                  | v["PALLET"].str.upper().str.contains(k, na=False)]
        if plants:
            v = v[v["PLANT"].isin(plants)]
        if modes:
            v = v[v["MODE"].isin(modes)]
        if only_bal:
            v = v[v["BALANCE"] > 0]

        a, b, c, dcol = st.columns(4)
        a.metric("Rows", len(v))
        b.metric("Pallets", int(v["PALLET"].nunique()) if len(v) else 0)
        c.metric("Actual Qty", f"{v['ACTUAL_QTY'].sum():g}" if len(v) else "0")
        dcol.metric("Pickable Balance", f"{v['BALANCE'].sum():g}" if len(v) else "0")
        st.dataframe(v, hide_index=True, use_container_width=True, height=520)

# =========================================================================== #
# TAB 3 — History
# =========================================================================== #
with tab_hist:
    st.subheader("📜 History")
    if not gs_ready:
        st.info("Google Sheet secrets දාන්න.")
    else:
        import gsheet
        which = st.selectbox("Worksheet", [gsheet.WS_RUNLOG, gsheet.WS_REGISTRY,
                                           gsheet.WS_LEDGER, gsheet.WS_REJECT,
                                           gsheet.WS_MASTER, gsheet.WS_DETAIL])
        if st.button("🔄 Load"):
            try:
                st.session_state["hist"] = (which, gsheet.read_ws(sa_info, sheet_key, which))
            except Exception as ex:
                st.error(f"Read error: {ex}")
        h = st.session_state.get("hist")
        if h and h[0] == which:
            df = h[1]
            st.metric("Rows", len(df))
            st.dataframe(df, hide_index=True, use_container_width=True, height=520)

# =========================================================================== #
# TAB 4 — Guide
# =========================================================================== #
with tab_help:
    st.subheader("📘 කොහොමද වැඩ කරන්නේ")
    st.markdown("""
**Flow**
1. Invoice / Delivery Challan **PDF** ටික + **Inventory Report** Excel එක upload කරන්න.
2. App එක PDF එකෙන් line items කියවලා, **completeness check** එකක් කරනවා
   (S.No sequence · Total Quantity · Grand Total amount). හරියන්නේ නැත්නම් ⛔ — pick කරන්නේ නෑ.
3. Inventory එකේ තියෙන **Plant** ටික පෙන්නලා confirm ගන්නවා.
4. Item match වෙන්නේ **base ID** එකෙන් විතරයි:
   `P162400-000-140` → **`P162400`**.
   Inventory එකේ `P162400-016-140` තිබ්බත් ගැලපෙනවා.
5. හැම line එකකටම stock තිබ්බොත් විතරයි pick කරන්නේ — **එක line එකක් මදි වුණත් මුළු
   document එකම reject** වෙනවා (`STOCK SHORT`).
6. එකම Invoice / DC number එකක් **දෙපාරක් process වෙන්නේ නෑ** (DOC_REGISTRY gate).
7. Pallet level එකෙන් `QTY_BEFORE → QTY_PICKED → QTY_BALANCE` ledger එකට save වෙනවා.

**Output — `OutBound_Upload_*.xlsx`** (හැම cell එකක්ම **text**)

*OutBound MASTER* — `DISPLAY_ORDER_NUMBER`, `STORE_ORDER_NUMBER`, `CUSTOMER_PO_NUMBER`,
`LOAD_ID` හතරටම Invoice No / DC No. අනිත් ඔක්කොම template එකේ විදිහටම.

*OutBound Detail*

| Column | එන්නේ කොහෙන්ද |
|---|---|
| DISPLAY_ORDER_NUMBER | Invoice No / DC No |
| LINE_NUMBER | 1, 2, 3, … |
| DISPLAY_ITEM_NUMBER | Inventory **Item Number** |
| LOT_NUMBER | Inventory Lot Number |
| QTY | Invoice / DC Quantity |
| ORDER_UOM | Inventory Uom |
| GEN_ATTRIBUTE_VALUE1–11 | Color · Size · Style · Supplier · **Plant** · Client So · Client So Line · Po Cust Dec · Customer Ref Number · Item Id · Invoice Number 1 |

**secrets.toml**
```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "xxx@yyy.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"

[google_sheet]
data_sheet = "https://docs.google.com/spreadsheets/d/<KEY>/edit"
auto_save  = true
wh_id       = "INMM01"
client_code = "INM0DONA"
```
Service-account email එකට Sheet එක **Editor** විදිහට share කරන්න ඕන.
""")

st.markdown(
    "<div class='footnote'>base-ID match · plant confirm · all-or-nothing per document · "
    "duplicate doc gate · pallet-level ledger · WMS cells = text</div>",
    unsafe_allow_html=True,
)
