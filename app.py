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
import sku_master as SKU
import ui

st.set_page_config(
    page_title="OutBound Pick · EFL",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"about": "OutBound Pick Generator — EFL / Donaldson · Körber One"},
)

ui.inject()


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

if "user_id" not in st.session_state:
    import uuid as _uuid
    st.session_state["user_id"] = f"user-{_uuid.uuid4().hex[:6]}"
USER = st.session_state["user_id"]


# --------------------------------------------------------------------------- #
# Sidebar
# --------------------------------------------------------------------------- #
sa_info = get_sa()
conf = gs_conf()
sheet_key = str(conf.get("data_sheet", "")).strip()
autosave = bool(conf.get("auto_save", True))
gs_ready = bool(sa_info and sheet_key)

with st.sidebar:
    st.markdown("<div class='eyebrow'>Setup</div>", unsafe_allow_html=True)

    with st.expander("Warehouse & client", expanded=False):
        wh_id = st.text_input("WH_ID", value=str(conf.get("wh_id", "INMM01")))
        client_code = st.text_input("CLIENT_CODE", value=str(conf.get("client_code", "INM0DONA")))
        order_type = st.text_input("ORDER_TYPE", value="Sales Orders")

    with st.expander("Pick options", expanded=False):
        strategy = st.selectbox(
            "Strategy",
            options=["FIFO", "SINGLE_PALLET_FIRST", "LEAST_PALLETS"],
            format_func=lambda s: {
                "FIFO": "FIFO — පරණ stock මුලින්",
                "SINGLE_PALLET_FIRST": "Single pallet first — හැකි නම් එක pallet එකෙන්",
                "LEAST_PALLETS": "Least pallets — ලොකු pallet මුලින්",
            }[s],
        )
        statuses = st.multiselect("Inventory status", ["Available", "Hold", "Damage", "QC"],
                                  default=["Available"])
        exact_first = st.checkbox("Exact item number first", value=True,
                                  help="Base ID එක ගැලපුනත්, document එකේ තියෙන "
                                       "full item number එකට මුලින් priority.")
        use_ledger = st.checkbox("Ledger balance logic", value=True,
                                 help="Pallet එකේ Actual Qty == ledger QTY_BEFORE නම් → "
                                      "QTY_BALANCE එකෙන් pick. වෙනස් නම් → Inventory "
                                      "Actual Qty එක අලුත් QTY_BEFORE එක විදිහට අරගෙන pick.")
        blank_fill = st.text_input("Blank attribute fill", value="TBC")
        fill_item_col = st.checkbox("Fill ITEM_NUMBER column", value=False)
        merge_lines = st.checkbox("Merge same-item lines", value=False)
        override = st.checkbox("Bypass document check", value=False,
                               help="Total Quantity / Grand Total / S.No check එක fail වුණත් "
                                    "pick කරනවා. Stock check එක bypass වෙන්නේ නෑ. "
                                    "Registry එකේ 'MANUAL OVERRIDE' කියලා log වෙනවා.")
        if override:
            st.warning("Completeness gate off — parse එක ඇස්සෙන් check කරන්න.")
        pick_date = st.date_input("Pick date", value=datetime.now())

    with st.expander("Email", expanded=False):
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
        mail_to = st.multiselect("To", options=saved, default=saved)
        extra_to = st.text_input("Add more (comma separated)", key="mail_extra")
        mail_to = list(mail_to) + PP._addr_list(extra_to)
        mail_cc = st.text_input("Cc", value=str(conf.get("mail_cc", "")))
        mail_from = st.text_input("From", value=str(conf.get("mail_from", "")))
        mail_sign = st.text_area("Signature", value=str(conf.get("mail_sign",
                                                                "Thanks & regards,")),
                                 height=70)

        new_addr = st.text_input("Save an address")
        ab1, ab2 = st.columns(2)
        if ab1.button("Add", width="stretch") and PP._addr_list(new_addr):
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
        if ab2.button("Clear", width="stretch"):
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
    st.markdown("<div class='eyebrow'>Database</div>", unsafe_allow_html=True)
    USER = st.text_input("Your name", value=USER, key="user_name",
                         help="Lock owner සහ SKU UPDATED_BY එකට යනවා.") or USER
    st.session_state["user_id"] = USER
    if gs_ready:
        st.caption(f"Connected · auto-save {'on' if autosave else 'off'}")
        if st.button("Set up worksheets", width="stretch"):
            try:
                import gsheet
                r = gsheet.init_sheet(sa_info, sheet_key)
                new = [k for k, v in r.items() if v is True]
                st.success("Ready · " + (f"created {', '.join(new)}" if new
                                         else "ඔක්කොම දැනටමත් තිබුණා"))
                st.markdown(f"[Open the sheet]({r.get('url','')})")
            except Exception as ex:
                st.error(f"Init error: {ex}")
    elif not sa_info:
        st.warning("`[gcp_service_account]` secret නෑ. Download වැඩ කරනවා, "
                   "ledger සහ duplicate check වැඩ කරන්නේ නෑ.")
    else:
        st.warning("`[google_sheet] data_sheet` secret නෑ.")

    with st.expander("API & multi-user"):
        ttl = st.slider("Read cache (seconds)", 0, 180, 45, 5,
                        help="User කීපදෙනෙක් වැඩ කරද්දී Google API quota ඉතුරු වෙනවා. "
                             "0 = cache නෑ (හැම වෙලේම fresh read).")
        try:
            import gsheet
            gsheet.set_cache_ttl(ttl)
            stt = gsheet.STATS
            a1, a2 = st.columns(2)
            a1.metric("API calls", stt["calls"])
            a2.metric("Cache hits", stt["cache_hits"])
            b1, b2 = st.columns(2)
            b1.metric("Retries", stt["retries"])
            b2.metric("Errors", stt["errors"], delta_color="inverse")
            if stt["last_error"]:
                st.caption(f"⚠️ {stt['last_error'][:120]}")
            if gs_ready:
                c1, c2 = st.columns(2)
                if c1.button("Health check", width="stretch"):
                    try:
                        h = gsheet.health(sa_info, sheet_key)
                        st.success(f"OK · {h['ms']} ms · {len(h['worksheets'])} worksheets")
                        if h["missing"]:
                            st.warning(f"නැති ඒවා: {', '.join(h['missing'])} — "
                                       "Initialize කරන්න.")
                    except Exception as ex:
                        st.error(f"{ex}")
                if c2.button("Clear cache", width="stretch"):
                    gsheet.cache_clear(sheet_key)
                    st.success("Cache clear ✅")
                locks = gsheet.active_locks(sa_info, sheet_key)
                if len(locks):
                    st.warning(f"{len(locks)} active lock — තව කෙනෙක් save කරමින්")
                    st.dataframe(locks, hide_index=True, width="stretch")
                    if st.button("Clear stale locks"):
                        st.success(f"{gsheet.clear_locks(sa_info, sheet_key)} clear ✅")
        except Exception as ex:
            st.caption(f"API panel: {ex}")

    with st.expander("Reset & undo"):
        undo = st.text_input("Undo a RUN_ID")
        if st.button("Delete run") and gs_ready and undo.strip():
            try:
                import gsheet
                r = gsheet.delete_run(sa_info, sheet_key, undo.strip())
                st.success(f"Deleted · {r}")
            except Exception as ex:
                st.error(f"Undo error: {ex}")

        st.markdown("---")
        st.markdown("<div class='eyebrow'>Database reset</div>", unsafe_allow_html=True)
        if not st.session_state.get("reset_ok"):
            pw = st.text_input("Password", type="password", key="reset_pw")
            if st.button("Unlock", width="stretch"):
                if pw == RESET_PASSWORD:
                    st.session_state["reset_ok"] = True
                    st.rerun()
                else:
                    st.error("Password වැරදියි.")
        else:
            st.caption("Unlocked")
            scope = st.multiselect(
                "Clear කරන්නේ",
                ["outputs", "ledger", "registry", "rejected", "runlog", "sku", "settings"],
                default=["outputs", "ledger", "registry", "rejected", "runlog"],
                format_func=lambda s: {
                    "outputs": "OUTBOUND_MASTER + DETAIL",
                    "ledger": "PALLET_LEDGER (pallet balance!)",
                    "registry": "DOC_REGISTRY (duplicate gate!)",
                    "rejected": "REJECTED_LOG",
                    "runlog": "RUN_LOG",
                    "sku": "SKU_MASTER",
                    "settings": "APP_SETTINGS (email book)",
                }[s],
            )
            sure = st.checkbox("මට විශ්වාසයි — back ගන්න බෑ", key="reset_sure")
            r1, r2 = st.columns(2)
            if r1.button("Reset selected", width="stretch"):
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
            if r2.button("Reset everything", width="stretch",
                         type="primary"):
                if not gs_ready:
                    st.error("Google Sheet connect වෙලා නෑ.")
                elif not sure:
                    st.warning("Confirm checkbox එක tick කරන්න.")
                else:
                    try:
                        import gsheet
                        r = gsheet.reset_all(sa_info, sheet_key, keep_settings=True,
                                             keep_sku=True)
                        for k in ("result", "bundles", "bundle_key", "zipfile", "hist"):
                            st.session_state.pop(k, None)
                        st.success(f"FULL RESET ✅ {r['count']} worksheets — "
                                   f"{', '.join(r['cleared'])}")
                        st.caption("SKU master + email book ඉතුරු කළා.")
                    except Exception as ex:
                        st.error(f"Reset error: {ex}")
            if st.button("Lock again"):
                st.session_state.pop("reset_ok", None)
                st.rerun()

    st.divider()
    if st.button("Clear this session", width="stretch"):
        for k in list(st.session_state.keys()):
            st.session_state.pop(k, None)
        st.rerun()

# Top bar + step rail are drawn into placeholders and filled at the end of the
# run — otherwise a file uploaded during THIS run would still read as missing.
SLOT_TOPBAR = st.empty()


def _draw_topbar():
    docs_n = len(st.session_state.get("docs", []) or [])
    plants = st.session_state.get("plants_ok", [])
    ui.topbar(
        "OutBound Pick Generator",
        f"Körber One · {client_code} · WH {wh_id} · "
        f"Invoice / Delivery Challan → pallet pick",
        chips=[
            {"label": "DB", "value": "connected" if gs_ready else "local only",
             "tone": "ok" if gs_ready else "warn"},
            {"label": "PLANT", "value": ", ".join(plants) if plants else "not set",
             "tone": "ok" if plants else ""},
            {"label": "DOCS", "value": str(docs_n) if docs_n else "none",
             "tone": "ok" if docs_n else ""},
            {"label": "USER", "value": USER, "tone": ""},
        ],
    )


def _draw_rail():
    docs_n = len(st.session_state.get("docs", []) or [])
    inv_ok = st.session_state.get("inv_raw") is not None
    plants = st.session_state.get("plants_ok", [])
    res_now = st.session_state.get("result")
    ui.rail([
        {"label": "Documents", "state": "done" if docs_n else "now",
         "value": f"{docs_n} uploaded" if docs_n else "Invoice / DC PDF"},
        {"label": "Inventory", "state": "done" if inv_ok else
         ("now" if docs_n else "todo"),
         "value": (f"{len(st.session_state['inv_raw']):,} rows" if inv_ok
                   else "Inventory report")},
        {"label": "Plant", "state": "done" if plants else ("now" if inv_ok else "todo"),
         "value": ", ".join(plants) if plants else "confirm කරන්න"},
        {"label": "Pick", "state": "done" if res_now else
         ("now" if (docs_n and inv_ok and plants) else "todo"),
         "value": (f"{len(res_now['accepted'])} picked" if res_now else "not run")},
    ])

(tab_gen, tab_loads, tab_sku, tab_search, tab_bal, tab_hist, tab_help) = st.tabs(
    ["Pick", "Loads", "SKU master", "Search", "Stock", "History", "Guide"]
)

# =========================================================================== #
# TAB 1 — Generate
# =========================================================================== #
with tab_gen:
    SLOT_RAIL = st.empty()
    ui.section("Documents & inventory", "01",
               "PDF කීයක් හරි එකවර දාන්න — invoice සහ DC මිශ්‍ර වුණත් කමක් නෑ")
    c1, c2 = st.columns(2)
    with c1:
        f_docs = st.file_uploader("Invoice / Delivery Challan (PDF)",
                                  type=["pdf"], accept_multiple_files=True,
                                  help="Donaldson tax invoice හෝ delivery challan. "
                                       "Multi-copy DC එකක් වුණත් copy එකක් විතරයි ගන්නේ.")
    with c2:
        f_inv = st.file_uploader("Inventory report (Excel)", type=["xlsx", "xls"],
                                 help="Körber One pallet-level inventory export.")

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
        n_ok = sum(1 for d in docs if d.completeness()[0])
        ui.section("Documents read", "02",
                   f"{n_ok} of {len(docs)} ready to pick")
        seen: set[str] = set()
        for d in docs:
            ok, probs = d.completeness()
            dup = d.doc_number in seen
            seen.add(d.doc_number)
            meta = (f"{d.doc_type.title()} · {d.doc_date or 'no date'} · "
                    f"{len(d.lines)} lines · qty {sum(l.qty for l in d.lines):g} · "
                    f"{d.source_file}")
            if dup:
                ui.doc_card(d.doc_number or "unreadable", meta, "warn", "duplicate",
                            ["එකම number එක දෙපාරක් — පළවෙනි එක විතරයි process වෙන්නේ"])
            elif ok:
                ui.doc_card(d.doc_number or "unreadable", meta, "ok", "ready")
            else:
                ui.doc_card(d.doc_number or "unreadable", meta, "stop", "blocked", probs)

        with st.expander("Lines හදන්න — parse එක වැරදුනොත් විතරක්"):
            st.caption("Base ID = match වෙන ID එක · row එකක් අයින් කරන්න `Use` "
                       "uncheck කරන්න.")
            edited = st.data_editor(
                st.session_state.get("doc_frame", pd.DataFrame()),
                num_rows="dynamic", width="stretch", height=320,
                key="doc_editor",
                column_config={
                    "Use": st.column_config.CheckboxColumn("Use", default=True),
                    "Qty": st.column_config.NumberColumn("Qty", min_value=0, step=1),
                    "Base ID": st.column_config.TextColumn("Base ID", disabled=True),
                },
            )
            if st.button("Apply changes"):
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
        ui.section("Plant", "03", "මොන plant එකෙන්ද pick කරන්නේ?")
        pc1, pc2 = st.columns([1.2, 1])
        with pc1:
            st.dataframe(psum, hide_index=True, width="stretch")
        with pc2:
            choice = st.multiselect(
                "Plant(s)", options=psum["Plant"].tolist(),
                default=st.session_state.get("plants_ok", []),
            )
            if st.button("Confirm plant", type="primary", width="stretch"):
                if not choice:
                    st.warning("Plant එකක් තෝරන්න.")
                else:
                    st.session_state["plants_ok"] = choice
                    st.session_state.pop("result", None)
                    st.toast(f"Plant confirmed — {', '.join(choice)}", icon="🏭")
                    st.rerun()
            if st.session_state.get("plants_ok"):
                st.markdown(ui.stamp("confirmed", "ok") +
                            f" &nbsp;<code>{', '.join(st.session_state['plants_ok'])}</code>",
                            unsafe_allow_html=True)

    # ---------------- generate ---------------- #
    ready = bool(st.session_state.get("docs")) and inv_raw is not None \
        and bool(st.session_state.get("plants_ok"))

    if not ready:
        missing = []
        if not st.session_state.get("docs"):
            missing.append("Invoice / DC PDF")
        if inv_raw is None:
            missing.append("Inventory report")
        if not st.session_state.get("plants_ok"):
            missing.append("Plant confirm")
        ui.empty("තව ටිකක් ඉතුරුයි", " · ".join(missing), "📥")
    else:
        cfg = E.EngineConfig(
            wh_id=wh_id, client_code=client_code, order_type=order_type,
            plants=st.session_state["plants_ok"], statuses=statuses, strategy=strategy,
            exact_item_first=exact_first, use_ledger=use_ledger, blank_fill=blank_fill,
            fill_item_number_col=fill_item_col, merge_same_item_lines=merge_lines,
            override_doc_check=override,
            pick_date=datetime.combine(pick_date, datetime.now().time()),
        )
        ui.section("Generate", "04",
                   f"{len(st.session_state['docs'])} documents · "
                   f"{', '.join(st.session_state['plants_ok'])} · {strategy}")
        gc1, gc2 = st.columns([2, 1])
        note = gc1.text_input("Run note", placeholder="උදා: 13-Aug morning batch",
                              label_visibility="collapsed")
        go = gc2.button("Generate pick", type="primary", width="stretch")

        if go:
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
                with st.status("Pick calculate කරනවා…", expanded=False) as sbox:
                    res = E.run_pick(st.session_state["docs"], inv_raw, cfg,
                                     ledger=ledger, processed_docs=done_docs,
                                     sku_desc=SKU.lookup(st.session_state.get("sku_df")))
                    sbox.update(label=f"Picked {len(res['accepted'])} of "
                                      f"{len(st.session_state['docs'])} documents",
                                state="complete")
                res["note"] = note
                st.session_state["result"] = res
                st.toast(f"Pick generated · RUN {res['run_id']}", icon="✅")
            except Exception as ex:
                st.error(f"Pick generate කරන්න බැරි වුණා — {ex}")
                st.session_state.pop("result", None)

            res = st.session_state.get("result")
            if res is not None and gs_ready and autosave and len(res["master"]):
                try:
                    import gsheet
                    with st.spinner("Google Sheet එකට save කරනවා..."):
                        r = gsheet.save_run(sa_info, sheet_key, res, cfg, note=note,
                                            owner=USER)
                    st.success(f"Sheet save වුණා ✅ master {r['master']} · detail {r['detail']} "
                               f"· ledger {r['ledger']} rows")
                    if r.get("skipped"):
                        st.warning("තව user කෙනෙක් මේවා දැනටමත් දාලා තිබුණා — skip කළා: "
                                   + ", ".join(r["skipped"]))
                    st.markdown(f"🔗 [Sheet එක open කරන්න]({r['url']})")
                    st.session_state["saved"] = res["run_id"]
                except gsheet.LockBusy as ex:
                    st.warning(f"🔒 {ex}")
                except Exception as ex:
                    st.error(f"Save error: {ex}")

    # ---------------- results ---------------- #
    res = st.session_state.get("result")
    if res:
        st.divider()
        acc, rej = res["accepted"], res["rejected"]
        alloc = res["allocations"]

        ui.section("Result", "05", f"RUN {res['run_id']} · {res['pick_date']}")
        if len(acc) and not len(rej):
            st.markdown(ui.stamp("all picked", "ok") + " &nbsp;" +
                        ui.muted("හැම document එකක්ම pick වුණා"),
                        unsafe_allow_html=True)
        elif len(acc):
            st.markdown(ui.stamp(f"{len(acc)} picked", "ok") + " &nbsp;" +
                        ui.stamp(f"{len(rej)} blocked", "stop"),
                        unsafe_allow_html=True)
        else:
            st.markdown(ui.stamp("nothing picked", "stop") + " &nbsp;" +
                        ui.muted("පහළ හේතුව බලන්න"), unsafe_allow_html=True)

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Documents", len(acc))
        m2.metric("Blocked", len(rej), delta_color="inverse")
        m3.metric("Order lines", len(res["detail"]))
        m4.metric("Pallets", int(alloc["PALLET"].nunique()) if len(alloc) else 0)
        m5.metric("Total Qty",
                  f"{pd.to_numeric(alloc['QTY_PICKED'], errors='coerce').sum():g}"
                  if len(alloc) else "0")

        vdf = res.get("verify", pd.DataFrame())
        bad_v = vdf[vdf["STATUS"].astype(str).str.contains("MISMATCH")] if len(vdf) else vdf
        if len(vdf) and not len(bad_v):
            st.success("Quantity verified — line · document total · WMS file total "
                       "තුනම Invoice / DC qty එකට හරියටම ගැලපෙනවා.")
        elif len(bad_v):
            st.error("Quantity mismatch — මේ documents pick කළේ නෑ:")
            st.dataframe(bad_v, hide_index=True, width="stretch")

        if len(rej):
            with st.expander(f"Blocked documents ({len(rej)})", expanded=True):
                st.dataframe(rej, hide_index=True, width="stretch")
                if len(res["shortage"]):
                    st.caption("Stock short lines:")
                    st.dataframe(res["shortage"], hide_index=True, width="stretch")

        # ------------------------------------------------------------------ #
        # Shortage — PDF (invoice එකත් එක්කම) + email
        # ------------------------------------------------------------------ #
        sh = res.get("shortage", pd.DataFrame())
        if len(sh):
            st.divider()
            ui.section("Shortage", "06", "stock මදි වුණ documents")
            src_map0 = st.session_state.get("doc_bytes", {})
            doc_map = {d.doc_number: d for d in st.session_state.get("docs", [])}
            rej_map = {str(r["DOC_NUMBER"]): str(r.get("REASON", ""))
                       for _, r in rej.iterrows()} if len(rej) else {}

            sh_chart = PP.shortage_chart_png(sh)          # email + PDF (white)
            sh_nums = list(dict.fromkeys(sh["DOC_NUMBER"].astype(str)))
            sh_att = st.checkbox("Shortage PDF එකට Invoice / DC pages එකතු කරන්න",
                                 value=True, key="sh_attach")

            sh_files: list[tuple[str, bytes]] = []
            sh_infos: list[dict] = []
            for num in sh_nums:
                one = sh[sh["DOC_NUMBER"].astype(str) == num]
                doc = doc_map.get(num)
                info = {
                    "DOC_NUMBER": num,
                    "DOC_TYPE": doc.doc_type if doc else str(one.iloc[0].get("DOC_TYPE", "")),
                    "DOC_DATE": doc.doc_date if doc else "",
                    "PLANT": ", ".join(st.session_state.get("plants_ok", [])),
                    "TOTAL_QTY": sum(l.qty for l in doc.lines) if doc else "",
                    "REASON": rej_map.get(num, "Stock short"),
                    "RUN_ID": res["run_id"], "WH_ID": wh_id, "CLIENT": client_code,
                    "SOURCE_FILE": doc.source_file if doc else "",
                }
                sh_infos.append(info)
                lines_df = P.docs_to_frame([doc]) if doc else None
                pdf_b = PP.build_shortage_pdf(
                    info, one, lines_df, src_map0.get(info["SOURCE_FILE"]),
                    attach_source=sh_att,
                    chart=PP.shortage_chart_png(one, f"Shortage · {num}"))
                fn = f"SHORT_{E.safe_name(num)}.pdf"
                sh_files.append((fn, pdf_b))
                c1, c2 = st.columns([3, 1])
                c1.markdown(f"**`{num}`** · {len(one)} short lines · "
                            f"short qty **{E._qty_str(float(pd.to_numeric(one['SHORT'], errors='coerce').sum()))}**")
                c2.download_button("Shortage PDF", data=pdf_b, file_name=fn,
                                   mime="application/pdf", width="stretch",
                                   key=f"sh_{E.safe_name(num)}")

            if sh_chart:
                st.image(PP.shortage_chart_png(sh, dark=(ui.theme_type() == "dark")),
                         caption="Shortage by item — මේකම email එකටත් යනවා",
                         width="content")

            s_subj, s_body, s_html = PP.shortage_email_text(sh_infos, sh, mail_sign)
            s_subject = st.text_input("Shortage subject", value=s_subj, key="sh_subject")
            s_bodyx = st.text_area("Shortage body", value=s_body, height=200,
                                   key="sh_body")
            to_l = PP._addr_list(mail_to)
            cc_l = PP._addr_list(mail_cc)
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                st.link_button("Open in mail app",
                               PP.mailto_link(to_l, s_subject, s_bodyx, cc_l),
                               width="stretch", disabled=not to_l)
            with sc2:
                st.download_button(
                    "Download draft (.eml)",
                    data=PP.build_eml(to_l, s_subject, s_bodyx, s_html, cc_l,
                                      sender=mail_from,
                                      attachments=[(n, b, "application/pdf")
                                                   for n, b in sh_files],
                                      inline_png=sh_chart),
                    file_name=f"SHORTAGE_{E.safe_name(sh_nums[0])}.eml",
                    mime="message/rfc822", width="stretch")
            with sc3:
                st.download_button("ZIP — all shortage PDFs",
                                   data=E.build_zip(sh_files),
                                   file_name=f"Shortage_{res['run_id']}.zip",
                                   mime="application/zip", width="stretch",
                                   disabled=len(sh_files) < 2)

        t1, t2, t3, t4, t7, t5, t6 = st.tabs(
            ["OutBound MASTER", "OutBound Detail", "Pick allocation",
             "Pallet balance", "Stock basis", "Qty verify", "Summary"])
        with t1:
            st.dataframe(res["master"], hide_index=True, width="stretch", height=280)
        with t2:
            st.dataframe(res["detail"], hide_index=True, width="stretch", height=380)
        with t3:
            st.caption("කොයි pallet එකෙන් කීයද ගත්තේ — balance එකත් එක්කම.")
            st.dataframe(alloc, hide_index=True, width="stretch", height=380)
        with t4:
            st.caption("QTY_BEFORE → QTY_PICKED → QTY_BALANCE. **MODE** එකෙන් පේනවා "
                       "ledger balance එකෙන්ද, අලුත් inventory qty එකෙන්ද pick කරේ කියලා.")
            st.dataframe(res["balance"], hide_index=True, width="stretch", height=380)
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
                st.dataframe(view, hide_index=True, width="stretch", height=380)
            else:
                st.info("Basis data නෑ.")
        with t5:
            st.caption("Invoice / DC එකේ Quantity එකට **හරියටම** ගැලපෙනවද — line by line, "
                       "document total, WMS file total.")
            st.dataframe(res["verify"], hide_index=True, width="stretch", height=380)
        with t6:
            st.dataframe(acc, hide_index=True, width="stretch", height=280)

        # ------------------------------------------------------------------ #
        # Downloads — LOAD_ID එකෙන් save
        # ------------------------------------------------------------------ #
        st.divider()
        ui.section("Files", "07", "හැම LOAD_ID එකකටම වෙන වෙනම")
        ids = E.load_ids(res)
        src_map = st.session_state.get("doc_bytes", {})

        if not ids:
            ui.empty("Download කරන්න දෙයක් නෑ", "Pick වුණ document එකක් නෑ.", "🗂️")
        else:
            attach_src = st.checkbox("Pick sheet PDF එකට Invoice / DC pages එකතු කරන්න",
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
                    h2.download_button("WMS Excel", data=b["xlsx"],
                                       file_name=f"{b['safe']}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument."
                                            "spreadsheetml.sheet",
                                       width="stretch", key=f"x_{b['safe']}")
                    h3.download_button("Pick sheet", data=b["pdf"],
                                       file_name=f"{b['safe']}.pdf",
                                       mime="application/pdf", width="stretch",
                                       key=f"p_{b['safe']}")

            ui.eyebrow("සියල්ලම එකට")
            z1, z2, z3 = st.columns(3)
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            z1.download_button("ZIP — every load",
                               data=st.session_state["zipfile"],
                               file_name=f"OutBound_{stamp}.zip",
                               mime="application/zip", width="stretch")
            z2.download_button("One Excel — all docs",
                               data=E.build_wms_excel(res["master"], res["detail"]),
                               file_name=f"OutBound_Upload_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet", width="stretch")
            z3.download_button("Pick report", data=E.build_report_excel(res),
                               file_name=f"Pick_Report_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet", width="stretch")

            # -------------------------------------------------------------- #
            # Email
            # -------------------------------------------------------------- #
            st.divider()
            ui.section("Email", "08", "pick details + item chart")
            pick_for_mail = st.multiselect("Loads to include", ids, default=ids)
            infos = [bundles[l]["info"] for l in pick_for_mail]
            m_alloc = alloc[alloc["DOC_NUMBER"].astype(str).isin(pick_for_mail)] \
                if len(alloc) else alloc

            if infos:
                subj0, body0, html0 = PP.pick_email_text(infos, m_alloc, mail_sign)
                pick_chart = PP.pick_chart_png(m_alloc, "Picked qty by item")
                e1, e2 = st.columns([2, 1])
                subject = e1.text_input("Subject", value=subj0, key="mail_subject")
                add_att = e2.checkbox("Attach Excel + PDF", value=True)
                body = st.text_area("Body", value=body0, height=240, key="mail_body")

                to_list = PP._addr_list(mail_to)
                cc_list = PP._addr_list(mail_cc)
                if not to_list:
                    st.warning("To address එකක් නෑ — sidebar → Email එකේ දාන්න.")
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
                    st.link_button("Open in mail app",
                                   PP.mailto_link(to_list, subject, body, cc_list),
                                   width="stretch",
                                   disabled=not to_list)
                    st.caption("Body විතරයි යන්නේ — attachment යන්නේ නෑ.")
                with mc2:
                    st.download_button(
                        "Download draft (.eml)",
                        data=PP.build_eml(to_list, subject, body, html0, cc_list,
                                          sender=mail_from, attachments=atts,
                                          inline_png=pick_chart),
                        file_name=f"PICK_{E.safe_name(pick_for_mail[0])}"
                                  f"{'_+' + str(len(pick_for_mail) - 1) if len(pick_for_mail) > 1 else ''}.eml",
                        mime="message/rfc822", width="stretch")
                    st.caption("Double-click කරාම Outlook/Mail එකේ draft එකක් විදිහට "
                               "attachment + chart එක්කම open වෙනවා.")
                if pick_chart:
                    with st.expander("Email එකට යන item chart එක", expanded=False):
                        st.image(PP.pick_chart_png(m_alloc, "Picked qty by item",
                                                   dark=(ui.theme_type() == "dark")),
                                 width="content")

        if gs_ready and not autosave and st.session_state.get("saved") != res["run_id"]:
            if st.button("Save to database"):
                try:
                    import gsheet
                    r = gsheet.save_run(sa_info, sheet_key, res, res["cfg"],
                                        note=res.get("note", ""), owner=USER)
                    st.session_state["saved"] = res["run_id"]
                    st.success("Save වුණා ✅")
                    st.markdown(f"🔗 [Sheet එක open කරන්න]({r['url']})")
                except Exception as ex:
                    st.error(f"Save error: {ex}")



# =========================================================================== #
# TAB — Loads (LOAD_ID එකෙන් download / delete)
# =========================================================================== #
with tab_loads:
    ui.section("Load manager", hint="LOAD_ID එකෙන් බලන්න · download · delete")

    if not gs_ready:
        ui.empty("Database එක connect වෙලා නෑ",
                 "secrets.toml එකේ Google Sheet එක දාපුවම save කරපු loads මෙතන පේනවා.",
                 "🔌")
    else:
        import gsheet

        lc1, lc2, lc3 = st.columns([2, 2, 1])
        typed = lc1.text_input("LOAD_ID", placeholder="333262712337  ·  333/26-27/62",
                               key="load_q")
        try:
            reg = gsheet.list_loads(sa_info, sheet_key)
        except Exception as ex:
            reg = pd.DataFrame()
            st.error(f"Registry read error: {ex}")
        opts = [""] + (reg["DOC_NUMBER"].astype(str).tolist() if len(reg) else [])
        picked = lc2.selectbox("නැත්නම් list එකෙන් තෝරන්න", opts, key="load_sel")
        lc3.write("")
        if lc3.button("Refresh", width="stretch"):
            gsheet.cache_clear(sheet_key)
            st.rerun()

        load_id = (typed.strip() or picked.strip())

        with st.expander(f"Saved loads ({len(reg)})", expanded=not load_id):
            if len(reg):
                st.dataframe(reg, hide_index=True, width="stretch", height=300)
            else:
                ui.empty("තාම load එකක් නෑ",
                         "Pick tab එකෙන් document එකක් generate කරලා save කරන්න.", "🗂️")

        if load_id:
            try:
                data = gsheet.read_load(sa_info, sheet_key, load_id, fresh=True)
            except Exception as ex:
                data = {}
                st.error(f"Read error: {ex}")

            m = data.get(gsheet.WS_MASTER, pd.DataFrame())
            d = data.get(gsheet.WS_DETAIL, pd.DataFrame())
            led = data.get(gsheet.WS_LEDGER, pd.DataFrame())
            rg = data.get(gsheet.WS_REGISTRY, pd.DataFrame())

            if not len(m) and not len(d) and not len(led):
                ui.empty(f"{load_id} — DB එකේ නෑ",
                         "LOAD_ID එක check කරන්න, නැත්නම් list එකෙන් තෝරන්න.", "🔍")
            else:
                info_row = rg.iloc[-1].to_dict() if len(rg) else {}
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("Master rows", len(m))
                k2.metric("Detail lines", len(d))
                k3.metric("Ledger rows", len(led))
                k4.metric("Doc Qty", info_row.get("DOC_QTY", "—"))
                k5.metric("Pallets", info_row.get("PALLETS", "—"))
                if info_row:
                    st.caption(f"{info_row.get('DOC_TYPE','')} · {info_row.get('DOC_DATE','')}"
                               f" · plant {info_row.get('PLANTS','')} · "
                               f"RUN `{info_row.get('RUN_ID','')}` · "
                               f"{info_row.get('PROCESSED_AT','')} · "
                               f"qty check {info_row.get('VERIFY','')}")

                lt1, lt2, lt3 = st.tabs(["Pick details", "OutBound Detail",
                                         "OutBound MASTER"])
                lt1.dataframe(led, hide_index=True, width="stretch", height=320)
                lt2.dataframe(d, hide_index=True, width="stretch", height=320)
                lt3.dataframe(m, hide_index=True, width="stretch", height=320)

                safe = E.safe_name(load_id)
                drop = ["RUN_ID", "PROCESSED_AT"]
                m_out = m.drop(columns=[c for c in drop if c in m.columns], errors="ignore")
                d_out = d.drop(columns=[c for c in drop if c in d.columns], errors="ignore")

                info_pdf = {
                    "LOAD_ID": load_id, "DOC_NUMBER": load_id,
                    "DOC_TYPE": info_row.get("DOC_TYPE", ""),
                    "DOC_DATE": info_row.get("DOC_DATE", ""),
                    "REF_NUMBER": info_row.get("REF_NUMBER", ""),
                    "PLANT": info_row.get("PLANTS", ""),
                    "LINES": info_row.get("LINES", len(d)),
                    "TOTAL_QTY": info_row.get("DOC_QTY", ""),
                    "PALLETS": info_row.get("PALLETS", ""),
                    "VERIFY": info_row.get("VERIFY", ""),
                    "RUN_ID": info_row.get("RUN_ID", ""),
                    "PICK_DATE": info_row.get("PROCESSED_AT", ""),
                    "STRATEGY": "", "WH_ID": wh_id, "CLIENT": client_code,
                }
                ui.eyebrow("Download")
                g1, g2, g3 = st.columns(3)
                g1.download_button(
                    "WMS Excel", data=E.build_wms_excel(m_out, d_out),
                    file_name=f"{safe}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument."
                         "spreadsheetml.sheet", width="stretch",
                    disabled=not len(m_out))
                try:
                    pdf_b = PP.build_doc_pdf(info_pdf, led, None, None,
                                             attach_source=False)
                except Exception as ex:
                    pdf_b = b""
                    st.warning(f"PDF error: {ex}")
                g2.download_button("Pick sheet PDF", data=pdf_b,
                                   file_name=f"{safe}.pdf", mime="application/pdf",
                                   width="stretch", disabled=not pdf_b)
                g3.download_button("Pick details CSV",
                                   data=led.to_csv(index=False).encode(),
                                   file_name=f"{safe}_pick_details.csv", mime="text/csv",
                                   width="stretch", disabled=not len(led))
                st.caption("PDF එකේ තියෙන්නේ pick sheet එක විතරයි — original Invoice / DC "
                           "PDF එක DB එකේ save වෙන්නේ නෑ.")

                ui.eyebrow("Delete")
                st.caption("Delete කරාම ledger + registry එකෙන් අයින් වෙනවා — ඒ නිසා "
                           "**pallet balance ආපහු එනවා**, ආපහු pick කරන්නත් පුළුවන්.")
                if not st.session_state.get("reset_ok"):
                    st.info("Delete කරන්න sidebar → Reset & undo → password එකෙන් "
                            "unlock කරන්න.")
                else:
                    dl1, dl2 = st.columns([2, 1])
                    typed_id = dl1.text_input(f"Confirm — `{load_id}` ආපහු type කරන්න",
                                              key="del_confirm")
                    dl2.write("")
                    if dl2.button("Delete load", width="stretch",
                                  type="primary"):
                        if typed_id.strip() != load_id:
                            st.error("LOAD_ID එක හරියටම match වෙන්නේ නෑ.")
                        else:
                            try:
                                with st.spinner("Delete කරනවා..."):
                                    r = gsheet.delete_load(sa_info, sheet_key, load_id,
                                                           owner=USER)
                                st.toast(f"Deleted {load_id}", icon="🗑️")
                                st.rerun()
                            except gsheet.LockBusy as ex:
                                st.warning(str(ex))
                            except Exception as ex:
                                st.error(f"Delete error: {ex}")


# =========================================================================== #
# TAB — SKU Master
# =========================================================================== #
with tab_sku:
    ui.section("SKU master",
               hint="එකම item එකක් ආපහු දැම්මොත් update වෙනවා — අලුත් row එකක් නෑ")

    if "sku_df" not in st.session_state:
        base = pd.DataFrame(columns=SKU.CORE)
        if gs_ready:
            try:
                import gsheet
                base = gsheet.read_sku(sa_info, sheet_key)
            except Exception as ex:
                st.warning(f"SKU read error: {ex}")
        st.session_state["sku_df"] = base if base is not None else pd.DataFrame()

    master = st.session_state["sku_df"]
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("SKU count", len(master))
    s2.metric("Base IDs", int(master["BASE_ID"].nunique()) if len(master) and
              "BASE_ID" in master.columns else 0)
    s3.metric("Columns", len(master.columns) if len(master) else 0)
    if s4.button("Reload from sheet", width="stretch",
                 disabled=not gs_ready):
        try:
            import gsheet
            st.session_state["sku_df"] = gsheet.read_sku(sa_info, sheet_key, fresh=True)
            st.success("Reload වුණා ✅")
            st.rerun()
        except Exception as ex:
            st.error(f"Reload error: {ex}")

    sku_up, sku_find, sku_edit = st.tabs(["Upload", "Search", "Edit"])

    # ---------------- upload ----------------
    with sku_up:
        u1, u2 = st.columns([2, 1])
        f_sku = u1.file_uploader("SKU file (Excel / CSV)", type=["xlsx", "xls", "csv"],
                                 key="sku_file")
        u2.write("")
        u2.download_button("Template", data=SKU.template_excel(),
                           file_name="SKU_Master_Template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet", width="stretch")
        if f_sku is not None:
            try:
                raw_sku = (pd.read_csv(f_sku, dtype=str) if f_sku.name.lower().endswith("csv")
                           else pd.read_excel(f_sku, dtype=str))
                inc = SKU.normalize(raw_sku, source=f_sku.name, user=USER)
                st.caption(f"{len(inc)} rows read · {', '.join(inc.columns[:8])}")
                prev = SKU.upsert(master, inc)
                p1, p2, p3, p4 = st.columns(4)
                p1.metric("New", prev["new"])
                p2.metric("Updated", prev["updated"])
                p3.metric("Unchanged", prev["unchanged"])
                p4.metric("Total after save", len(prev["data"]))
                if len(prev["changes"]):
                    with st.expander("වෙනස් වෙන rows මොනවද", expanded=True):
                        cols = [c for c in ["ITEM_NUMBER", "ITEM_DESCRIPTION", "_STATUS",
                                            "_CHANGED"] if c in prev["changes"].columns]
                        st.dataframe(prev["changes"][cols], hide_index=True,
                                     width="stretch", height=260)
                if st.button("Save SKU master", type="primary",
                             width="stretch"):
                    st.session_state["sku_df"] = prev["data"]
                    if gs_ready:
                        try:
                            import gsheet
                            with st.spinner("Sheet එකට save කරනවා..."):
                                r = gsheet.save_sku(sa_info, sheet_key, prev["data"],
                                                    owner=USER)
                            st.toast(f"Saved · {r['rows']} SKUs "
                                     f"({prev['new']} new, {prev['updated']} updated)",
                                     icon="🏷️")
                        except gsheet.LockBusy as ex:
                            st.warning(str(ex))
                        except Exception as ex:
                            st.error(f"Save error: {ex}")
                    else:
                        st.info("Google Sheet නෑ — session එකේ විතරක් තියෙනවා.")
                    st.rerun()
            except Exception as ex:
                st.error(f"File error: {ex}")

    # ---------------- search ----------------
    with sku_find:
        q = st.text_input("🔍 Item number / base ID / description",
                          placeholder="07011636   ·   07011636-000-440   ·   gasket epdm",
                          key="sku_q")
        base_on = st.checkbox("Base ID match — suffix නැතුව හොයන්න", value=True)
        if not len(master):
            ui.empty("SKU master හිස්", "Upload tab එකෙන් file එකක් දාන්න.", "🏷️")
        elif q.strip():
            hit = SKU.search(master, q, base_match=base_on)
            if not len(hit):
                st.warning(f"'{q}' — හම්බුණේ නෑ. Base ID එකෙන් try කරන්න.")
            else:
                st.caption(f"{len(hit)} rows")
                st.dataframe(hit, hide_index=True, width="stretch", height=420)
                st.download_button("Download CSV",
                                   data=hit.to_csv(index=False).encode(),
                                   file_name="sku_search.csv", mime="text/csv")
        else:
            st.dataframe(master.head(200), hide_index=True, width="stretch",
                         height=420)
            st.caption(f"මුල් 200 rows පෙන්නනවා ({len(master)} total).")

    # ---------------- edit ----------------
    with sku_edit:
        st.caption("කෙලින්ම edit කරන්න / අලුත් row එකක් add කරන්න. "
                   "Save කරද්දී duplicate check එක ආපහු වෙනවා.")
        show = master if len(master) else pd.DataFrame(columns=SKU.CORE)
        ed = st.data_editor(show, num_rows="dynamic", width="stretch",
                            height=420, key="sku_editor",
                            column_config={
                                "BASE_ID": st.column_config.TextColumn("BASE_ID",
                                                                       disabled=True),
                                "MATCH_KEY": st.column_config.TextColumn("MATCH_KEY",
                                                                          disabled=True)})
        e1, e2 = st.columns(2)
        if e1.button("Save changes", type="primary", width="stretch"):
            try:
                clean = SKU.normalize(ed.rename(columns={"ITEM_NUMBER": "Item Number",
                                                         "ITEM_DESCRIPTION":
                                                             "Item Description"}),
                                      source="manual edit", user=USER)
                st.session_state["sku_df"] = clean
                if gs_ready:
                    import gsheet
                    r = gsheet.save_sku(sa_info, sheet_key, clean, owner=USER)
                    st.success(f"Save වුණා ✅ {r['rows']} rows")
                else:
                    st.success(f"Save වුණා ✅ {len(clean)} rows (session)")
                st.rerun()
            except Exception as ex:
                st.error(f"Save error: {ex}")
        e2.download_button("Download all", data=SKU.to_excel(master),
                           file_name="SKU_Master.xlsx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet", width="stretch",
                           disabled=not len(master))

# =========================================================================== #
# TAB — Global search
# =========================================================================== #
with tab_search:
    ui.section("Search everything",
               hint="word කීපයක් දුන්නොත් ඔක්කොම තියෙන rows විතරයි")

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
        ui.empty("හොයන්න ඕන දේ type කරන්න",
                 "Item code · LOAD ID · pallet · location · GRN · lot · plant", "🔎")
    elif not frames:
        ui.empty("Search කරන්න data නෑ", "Pick tab එකෙන් file දාන්න.", "📭")
    else:
        hits = E.search_frames(q, frames)
        total = sum(len(v) for v in hits.values())
        if not total:
            st.warning(f"'{q}' — කිසිම තැනක නෑ.")
        else:
            st.caption(f"{total} rows · {len(hits)} තැනක")
            for name, df in hits.items():
                with st.expander(f"{name} — {len(df)} rows", expanded=len(hits) <= 2):
                    st.dataframe(df, hide_index=True, width="stretch",
                                 height=min(420, 60 + 32 * len(df)))
                    st.download_button("Download CSV",
                                       data=df.to_csv(index=False).encode(),
                                       file_name=f"search_{E.safe_name(name)}.csv",
                                       mime="text/csv", key=f"dl_{name}")

# =========================================================================== #
# TAB — Pallet balance
# =========================================================================== #
with tab_bal:
    ui.section("Pallet stock",
               hint="MODE එකෙන් පේනවා ledger balance ද අලුත් inventory qty ද කියලා")
    inv_raw = st.session_state.get("inv_raw")
    if inv_raw is None:
        ui.empty("Inventory report එකක් නෑ", "Pick tab එකෙන් upload කරන්න.", "📦")
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
        f4.write("")
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
        st.dataframe(v, hide_index=True, width="stretch", height=520)

# =========================================================================== #
# TAB 3 — History
# =========================================================================== #
with tab_hist:
    ui.section("History", hint="DB එකේ තියෙන raw worksheets")
    if not gs_ready:
        ui.empty("Database එක connect වෙලා නෑ",
                 "secrets.toml එකේ Google Sheet එක දාන්න.", "🔌")
    else:
        import gsheet
        which = st.selectbox("Worksheet", [gsheet.WS_RUNLOG, gsheet.WS_REGISTRY,
                                           gsheet.WS_LEDGER, gsheet.WS_REJECT,
                                           gsheet.WS_MASTER, gsheet.WS_DETAIL])
        if st.button("Load worksheet"):
            try:
                st.session_state["hist"] = (which, gsheet.read_ws(sa_info, sheet_key, which))
            except Exception as ex:
                st.error(f"Read error: {ex}")
        h = st.session_state.get("hist")
        if h and h[0] == which:
            df = h[1]
            st.metric("Rows", len(df))
            st.dataframe(df, hide_index=True, width="stretch", height=520)

# =========================================================================== #
# TAB 4 — Guide
# =========================================================================== #
with tab_help:
    ui.section("How it works", hint="quick reference")
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

with SLOT_TOPBAR.container():
    _draw_topbar()
with SLOT_RAIL.container():
    _draw_rail()

ui.footnote("base-ID match · plant confirm · all-or-nothing per document · "
            "duplicate gate · pallet-level ledger · qty verified · WMS cells = text")
