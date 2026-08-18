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
import invoice_register as R
import sku_master as SKU
import ui

st.set_page_config(
    page_title="OutBound Pick · EFL",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"about": "OutBound Pick Generator — EFL / Donaldson · Körber One"},
)

# --------------------------------------------------------------------------- #
# Module version gate
# --------------------------------------------------------------------------- #
# Streamlit Cloud redacts exception text, so a half-updated deploy used to die
# with an unreadable TypeError. Every module carries an API number; refuse to
# run against a stale one and say exactly which file to replace.
_NEEDS = {"pick_engine.py": (E, 4), "doc_parser.py": (P, 3), "pick_pdf.py": (PP, 4),
          "sku_master.py": (SKU, 2), "ui.py": (ui, 2),
          "invoice_register.py": (R, 3)}
_STALE = [(f, getattr(m, "API", 0), n) for f, (m, n) in _NEEDS.items()
          if getattr(m, "API", 0) < n]
if _STALE:
    st.error(
        "**These files are out of date — replace them and redeploy.**\n\n"
        + "\n".join(f"- `{f}` — found API {have or 'none'}, needs {need}"
                     for f, have, need in _STALE)
        + "\n\nEvery module in this app has to come from the same release. "
          "Updating `app.py` on its own leaves the engine without the fields it "
          "is being handed."
    )
    st.stop()

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


def _mrp_contacts():
    """Who counts as MRP — stored in APP_SETTINGS, editable on the Register tab."""
    if "mrp_contacts" not in st.session_state:
        raw = ""
        if _gs_ready():
            try:
                import gsheet
                raw = gsheet.read_setting(get_sa(), str(gs_conf().get("data_sheet", "")),
                                          "MRP_CONTACTS", "")
            except Exception:
                raw = ""
        st.session_state["mrp_contacts"] = R.load_contacts(raw)
    return st.session_state["mrp_contacts"]


def _gs_ready() -> bool:
    return bool(get_sa() and str(gs_conf().get("data_sheet", "")).strip())

if "user_id" not in st.session_state:
    import uuid as _uuid
    st.session_state["user_id"] = f"user-{_uuid.uuid4().hex[:6]}"
USER = st.session_state["user_id"]
MRP_CONTACTS = _mrp_contacts()


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
                "FIFO": "FIFO - oldest stock first",
                "SINGLE_PALLET_FIRST": "Single pallet first - fewest pallet touches",
                "LEAST_PALLETS": "Least pallets - largest pallet first",
            }[s],
        )
        statuses = st.multiselect("Inventory status", ["Available", "Hold", "Damage", "QC"],
                                  default=["Available"])
        pick_id_gate = st.checkbox("Pick Id = 0 only", value=True,
                                   help="A pallet with a non-zero Pick Id is already "
                                        "allocated to a pick task in the WMS — its Status "
                                        "still reads 'Available', so it has to be excluded "
                                        "here or the same stock gets picked twice.")
        exact_first = st.checkbox("Exact item number first", value=True,
                                  help="When several item numbers share a base ID, the "
                                       "one printed on the document is picked first.")
        use_ledger = st.checkbox("Ledger balance logic", value=True,
                                 help="If a pallet's Actual Qty equals the ledger "
                                      "QTY_BEFORE, the report is stale and the pick comes "
                                      "out of QTY_BALANCE. If it differs, the WMS has "
                                      "moved on and Actual Qty becomes the new baseline.")
        blank_fill = st.text_input("Blank attribute fill", value="TBC")
        fill_item_col = st.checkbox("Fill ITEM_NUMBER column", value=False)
        merge_lines = st.checkbox("Merge same-item lines", value=False)
        override = st.checkbox("Bypass document check", value=False,
                               help="Picks even when the Total Quantity / Grand Total / "
                                    "S.No check fails. The stock check is never bypassed. "
                                    "Logged in the registry as 'MANUAL OVERRIDE'.")
        if override:
            st.warning("Completeness gate is off - check the parsed lines yourself.")
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
        chart_style = st.radio(
            "Chart in the email", ["Line by document line", "Bar by item"],
            help="Line: one point per document line, in order, with the pallet "
                 "balance on a second axis. Bar: quantity grouped by item. "
                 "A single line falls back to bars either way.")
        CHART_LINE = chart_style.startswith("Line")

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
                         help="Used as the lock owner and as SKU UPDATED_BY.") or USER
    st.session_state["user_id"] = USER
    if gs_ready:
        st.caption(f"Connected · auto-save {'on' if autosave else 'off'}")
        if st.button("Set up worksheets", width="stretch"):
            try:
                import gsheet
                r = gsheet.init_sheet(sa_info, sheet_key)
                new = [k for k, v in r.items() if v is True]
                st.success("Ready · " + (f"created {', '.join(new)}" if new
                                         else "everything already existed"))
                st.markdown(f"[Open the sheet]({r.get('url','')})")
            except Exception as ex:
                st.error(f"Init error: {ex}")
    elif not sa_info:
        st.warning("No `[gcp_service_account]` secret. Downloads still work, "
                   "but the ledger and the duplicate check do not.")
    else:
        st.warning("No `[google_sheet] data_sheet` secret.")

    with st.expander("API & multi-user"):
        ttl = st.slider("Read cache (seconds)", 0, 180, 45, 5,
                        help="Saves Google API quota when several people work at "
                             "once. 0 = no cache, every read hits the API.")
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
                            st.warning(f"Missing: {', '.join(h['missing'])} - run "
                                       "'Set up worksheets'.")
                    except Exception as ex:
                        st.error(f"{ex}")
                if c2.button("Clear cache", width="stretch"):
                    gsheet.cache_clear(sheet_key)
                    st.success("Cache clear ✅")
                locks = gsheet.active_locks(sa_info, sheet_key)
                if len(locks):
                    st.warning(f"{len(locks)} active lock - someone else is saving")
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
                    st.error("Wrong password.")
        else:
            st.caption("Unlocked")
            scope = st.multiselect(
                "What to clear",
                ["outputs", "ledger", "registry", "rejected", "runlog", "sku",
                 "register", "settings"],
                default=["outputs", "ledger", "registry", "rejected", "runlog"],
                format_func=lambda s: {
                    "outputs": "OUTBOUND_MASTER + DETAIL",
                    "ledger": "PALLET_LEDGER (pallet balance!)",
                    "registry": "DOC_REGISTRY (duplicate gate!)",
                    "rejected": "REJECTED_LOG",
                    "runlog": "RUN_LOG",
                    "sku": "SKU_MASTER",
                    "register": "INVOICE_SUMMARY + INVOICE_DETAIL",
                    "settings": "APP_SETTINGS (email book)",
                }[s],
            )
            sure = st.checkbox("I am sure - this cannot be undone", key="reset_sure")
            r1, r2 = st.columns(2)
            if r1.button("Reset selected", width="stretch"):
                if not gs_ready:
                    st.error("Google Sheet is not connected.")
                elif not scope or not sure:
                    st.warning("Pick what to clear and tick the confirm box.")
                else:
                    try:
                        import gsheet
                        done = gsheet.reset_data(sa_info, sheet_key, scope)
                        for k in ("result", "bundles", "bundle_key", "zipfile", "hist"):
                            st.session_state.pop(k, None)
                        st.success(f"Reset done · {done}")
                    except Exception as ex:
                        st.error(f"Reset error: {ex}")
            if r2.button("Reset everything", width="stretch",
                         type="primary"):
                if not gs_ready:
                    st.error("Google Sheet is not connected.")
                elif not sure:
                    st.warning("Tick the confirm box first.")
                else:
                    try:
                        import gsheet
                        r = gsheet.reset_all(sa_info, sheet_key, keep_settings=True,
                                             keep_sku=True, keep_register=True)
                        for k in ("result", "bundles", "bundle_key", "zipfile", "hist"):
                            st.session_state.pop(k, None)
                        st.success(f"FULL RESET ✅ {r['count']} worksheets — "
                                   f"{', '.join(r['cleared'])}")
                        st.caption("SKU master, invoice register and the email "
                                   "book were kept.")
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
         "value": ", ".join(plants) if plants else "confirm a plant"},
        {"label": "Pick", "state": "done" if res_now else
         ("now" if (docs_n and inv_ok and plants) else "todo"),
         "value": (f"{len(res_now['accepted'])} picked" if res_now else "not run")},
    ])

(tab_gen, tab_dash, tab_reg, tab_loads, tab_sku, tab_search, tab_bal, tab_hist,
 tab_help) = st.tabs(["Pick", "Dashboard", "Register", "Loads", "SKU master",
                      "Search", "Stock", "History", "Guide"])

# =========================================================================== #
# TAB 1 — Generate
# =========================================================================== #
with tab_gen:
    SLOT_RAIL = st.empty()
    ui.section("Documents & inventory", "01",
               "drop in as many PDFs as you like - invoices and DCs can be mixed")
    c1, c2 = st.columns(2)
    with c1:
        f_docs = st.file_uploader("Invoice / Delivery Challan (PDF)",
                                  type=["pdf"], accept_multiple_files=True,
                                  help="Donaldson tax invoice or delivery challan. A DC "
                                       "printed in triplicate is read only once.")
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
            with st.spinner("Reading the PDFs..."):
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
                            ["Same number twice - only the first one is processed"])
            elif ok:
                ui.doc_card(d.doc_number or "unreadable", meta, "ok", "ready")
            else:
                ui.doc_card(d.doc_number or "unreadable", meta, "stop", "blocked", probs)

        with st.expander("Fix the lines - only if the parse went wrong"):
            st.caption("Base ID is what the match runs on · untick `Use` to "
                       "drop a row.")
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
                st.success("Changes applied")
                st.rerun()

    # ---------------- inventory + plant confirm ---------------- #
    # the parsed frame lives in session state, so a rerun that loses the
    # uploader value still has the inventory to work with
    if f_inv is not None and st.session_state.get("inv_sig") != (f_inv.name, f_inv.size):
        with st.spinner("Reading the inventory..."):
            st.session_state["inv_raw"] = pd.read_excel(f_inv, dtype=str)
        st.session_state["inv_sig"] = (f_inv.name, f_inv.size)
        for _k in ("plants_ok", "result", "release_locked"):
            st.session_state.pop(_k, None)
    inv_raw = st.session_state.get("inv_raw")

    if inv_raw is not None:
        inv_norm = E.normalize_inventory(inv_raw)
        psum = E.plant_summary(inv_norm)
        ui.section("Plant", "03", "which plant should this be picked from?")
        pc1, pc2 = st.columns([1.2, 1])
        with pc1:
            st.dataframe(psum, hide_index=True, width="stretch")
            if "On pick task" in psum.columns and psum["On pick task"].sum():
                st.caption(f"**Qty** counts only pallets with `Pick Id = 0`. "
                           f"{int(psum['On pick task'].sum()):,} units sit on pallets "
                           f"already allocated to a pick task and are left alone."
                           if pick_id_gate else
                           f"`Pick Id = 0 only` is off — "
                           f"{int(psum['On pick task'].sum()):,} units allocated to a "
                           f"pick task will also be picked.")
        with pc2:
            choice = st.multiselect(
                "Plant(s)", options=psum["Plant"].tolist(),
                default=st.session_state.get("plants_ok", []),
            )
            if st.button("Confirm plant", type="primary", width="stretch"):
                if not choice:
                    st.warning("Select a plant first.")
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
        ui.empty("Almost there", "Still needed: " + " · ".join(missing), "📥")
    else:
        _cfg_kw = dict(
            wh_id=wh_id, client_code=client_code, order_type=order_type,
            plants=st.session_state["plants_ok"], statuses=statuses, strategy=strategy,
            exact_item_first=exact_first, use_ledger=use_ledger,
            pick_id_zero_only=pick_id_gate,
            release_locked=st.session_state.get("release_locked", {}),
            blank_fill=blank_fill,
            fill_item_number_col=fill_item_col, merge_same_item_lines=merge_lines,
            override_doc_check=override,
            pick_date=datetime.combine(pick_date, datetime.now().time()),
        )
        try:
            cfg = E.EngineConfig(**_cfg_kw)
        except TypeError as ex:
            st.error(f"`pick_engine.py` does not match this `app.py` — {ex}. "
                     "Replace `pick_engine.py` with the copy from the same release "
                     "and redeploy.")
            st.stop()
        ui.section("Generate", "04",
                   f"{len(st.session_state['docs'])} documents · "
                   f"{', '.join(st.session_state['plants_ok'])} · {strategy}")
        gc1, gc2 = st.columns([2, 1])
        note = gc1.text_input("Run note", placeholder="e.g. 13-Aug morning batch",
                              label_visibility="collapsed")
        go = gc2.button("Generate pick", type="primary", width="stretch")

        _auto = st.session_state.pop("rerun_pick", False)
        if go or _auto:
            ledger = None
            done_docs: set[str] = set()
            if gs_ready:
                try:
                    import gsheet
                    ledger = gsheet.read_ledger(sa_info, sheet_key)
                    done_docs = gsheet.read_processed_docs(sa_info, sheet_key)
                except Exception as ex:
                    st.warning(f"Could not read the Google Sheet ({ex}) - "
                               "ledger and duplicate check were skipped.")
            try:
                with st.status("Calculating the pick...", expanded=False) as sbox:
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
                st.error(f"Could not generate the pick - {ex}")
                st.session_state.pop("result", None)

            res = st.session_state.get("result")

            # every uploaded document is registered, picked or not
            if res is not None:
                _sum, _det = R.build(st.session_state["docs"], res, MRP_CONTACTS,
                                     user=USER,
                                     plant=", ".join(st.session_state["plants_ok"]))
                st.session_state["reg_run"] = (_sum, _det)
                if gs_ready:
                    try:
                        import gsheet
                        rr = gsheet.save_register(sa_info, sheet_key, _sum, _det,
                                                  owner=USER)
                        st.session_state.pop("reg_cache", None)
                        st.caption(f"Register · {rr['new']} new, {rr['updated']} updated "
                                   f"— {rr['summary_rows']} invoices on file")
                    except Exception as ex:
                        st.warning(f"Register not saved: {ex}")

            if res is not None and gs_ready and autosave and len(res["master"]):
                try:
                    import gsheet
                    with st.spinner("Saving to the Google Sheet..."):
                        r = gsheet.save_run(sa_info, sheet_key, res, cfg, note=note,
                                            owner=USER)
                    st.success(f"Saved · master {r['master']} · detail {r['detail']} "
                               f"· ledger {r['ledger']} rows")
                    if r.get("skipped"):
                        st.warning("Another user had already saved these - skipped: "
                                   + ", ".join(r["skipped"]))
                    st.markdown(f"[Open the sheet]({r['url']})")
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
                        ui.muted("every document was picked"),
                        unsafe_allow_html=True)
        elif len(acc):
            st.markdown(ui.stamp(f"{len(acc)} picked", "ok") + " &nbsp;" +
                        ui.stamp(f"{len(rej)} blocked", "stop"),
                        unsafe_allow_html=True)
        else:
            st.markdown(ui.stamp("nothing picked", "stop") + " &nbsp;" +
                        ui.muted("see the reason below"), unsafe_allow_html=True)

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
                       "all three match the Invoice / DC quantity exactly.")
        elif len(bad_v):
            st.error("Quantity mismatch - these documents were not picked:")
            st.dataframe(bad_v, hide_index=True, width="stretch")

        if len(rej):
            with st.expander(f"Blocked documents ({len(rej)})", expanded=True):
                st.dataframe(rej, hide_index=True, width="stretch")
                if len(res["shortage"]):
                    st.caption("Stock short lines:")
                    st.dataframe(res["shortage"], hide_index=True, width="stretch")

        # ------------------------------------------------------------------ #
        # Shortage — PDF (merged with the invoice) + email
        # ------------------------------------------------------------------ #
        rel = E.releasable(res)
        if len(rel):
            st.divider()
            ui.section("Release stock",
                       hint="blocked only because the stock sits on another pick task")
            st.markdown(
                ui.stamp("confirmation needed", "warn") + " &nbsp;" +
                ui.muted("These documents have enough stock — it is just committed "
                         "to another pick task in the WMS. Releasing takes it anyway."),
                unsafe_allow_html=True)
            st.dataframe(rel, hide_index=True, width="stretch")

            lk_all = res.get("locked", pd.DataFrame())
            if len(lk_all):
                with st.expander("The pallets that would be taken"):
                    ids_all = {i.strip() for v in rel["PICK_IDS"] for i in str(v).split(",")
                               if i.strip()}
                    st.dataframe(lk_all[lk_all["PICK_ID"].astype(str).isin(ids_all)],
                                 hide_index=True, width="stretch")

            rc1, rc2 = st.columns([2, 1])
            take = rc1.multiselect("Documents to release", rel["DOC_NUMBER"].tolist(),
                                   default=rel["DOC_NUMBER"].tolist(),
                                   key="release_pick")
            sure_rel = rc1.checkbox(
                "I have checked the other pick task and this stock is free to take",
                key="release_confirm")
            rc2.write("")
            if rc2.button("Release and pick", type="primary", width="stretch",
                          disabled=not (take and sure_rel)):
                book = dict(st.session_state.get("release_locked", {}))
                for _, r in rel[rel["DOC_NUMBER"].isin(take)].iterrows():
                    book[str(r["DOC_NUMBER"])] = [i.strip() for i in
                                                  str(r["PICK_IDS"]).split(",") if i.strip()]
                st.session_state["release_locked"] = book
                st.session_state["rerun_pick"] = True
                st.toast(f"Released {len(take)} document(s) — picking again", icon="🔓")
                st.rerun()
            if not sure_rel:
                rc1.caption("The tick box is the confirmation — nothing is released "
                            "until it is on.")

        if st.session_state.get("release_locked"):
            done_rel = ", ".join(st.session_state["release_locked"])
            st.caption(f"Released from a pick task in this run: {done_rel} — "
                       "logged against the document in DOC_REGISTRY.")
            if st.button("Undo the release"):
                st.session_state.pop("release_locked", None)
                st.session_state["rerun_pick"] = True
                st.rerun()

        lk = res.get("locked", pd.DataFrame())
        if len(lk):
            with st.expander(f"On a pick task — excluded ({len(lk)} pallets · "
                             f"{E._qty_str(float(pd.to_numeric(lk['QTY'], errors='coerce').sum()))} units)"):
                st.caption("These pallets carry a non-zero `Pick Id`, so the WMS has "
                           "already committed them to another pick. Nothing was taken "
                           "off them.")
                st.dataframe(lk, hide_index=True, width="stretch")

        sh = res.get("shortage", pd.DataFrame())
        if len(sh):
            st.divider()
            ui.section("Shortage", "06", "documents blocked by short stock")
            src_map0 = st.session_state.get("doc_bytes", {})
            doc_map = {d.doc_number: d for d in st.session_state.get("docs", [])}
            rej_map = {str(r["DOC_NUMBER"]): str(r.get("REASON", ""))
                       for _, r in rej.iterrows()} if len(rej) else {}

            # email + PDF always on white paper
            sh_chart = (PP.shortage_line_chart_png(sh) if CHART_LINE
                        else PP.shortage_chart_png(sh))
            sh_nums = list(dict.fromkeys(sh["DOC_NUMBER"].astype(str)))
            sh_att = st.checkbox("Add the Invoice / DC pages to the shortage PDF",
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
                    chart=(PP.shortage_line_chart_png(one, f"Shortage · {num}")
                           if CHART_LINE
                           else PP.shortage_chart_png(one, f"Shortage · {num}")))
                fn = f"SHORT_{E.safe_name(num)}.pdf"
                sh_files.append((fn, pdf_b))
                c1, c2 = st.columns([3, 1])
                c1.markdown(f"**`{num}`** · {len(one)} short lines · "
                            f"short qty **{E._qty_str(float(pd.to_numeric(one['SHORT'], errors='coerce').sum()))}**")
                c2.download_button("Shortage PDF", data=pdf_b, file_name=fn,
                                   mime="application/pdf", width="stretch",
                                   key=f"sh_{E.safe_name(num)}")

            if sh_chart:
                _dk = ui.theme_type() == "dark"
                st.image(PP.shortage_line_chart_png(sh, dark=_dk) if CHART_LINE
                         else PP.shortage_chart_png(sh, dark=_dk),
                         caption="The same chart goes into the shortage email",
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
            st.caption("Which pallet each unit came off, with the balance left behind.")
            st.dataframe(alloc, hide_index=True, width="stretch", height=380)
        with t4:
            st.caption("QTY_BEFORE → QTY_PICKED → QTY_BALANCE. **MODE** shows whether "
                       "the pick came out of the ledger balance or a fresh "
                       "inventory quantity.")
            st.dataframe(res["balance"], hide_index=True, width="stretch", height=380)
        with t7:
            bs = res.get("basis", pd.DataFrame())
            st.caption("How the maximum pickable quantity was decided for each pallet.")
            if len(bs):
                mc = bs["MODE"].value_counts()
                b1, b2, b3 = st.columns(3)
                b1.metric("NEW (no ledger)", int(mc.get("NEW", 0)))
                b2.metric("LEDGER BALANCE", int(mc.get("LEDGER BALANCE", 0)),
                          help="Inventory not refreshed yet - picked from the ledger balance")
                b3.metric("NEW BASELINE", int(mc.get("NEW BASELINE", 0)),
                          help="Actual Qty changed - it becomes the new QTY_BEFORE")
                only = st.checkbox("Only pallets with ledger history", value=True,
                                   key="basis_filter")
                view = bs[bs["MODE"] != "NEW"] if only else bs
                st.dataframe(view, hide_index=True, width="stretch", height=380)
            else:
                st.info("No basis data.")
        with t5:
            st.caption("Does it match the Invoice / DC quantity **exactly** - line by "
                       "document total, WMS file total.")
            st.dataframe(res["verify"], hide_index=True, width="stretch", height=380)
        with t6:
            st.dataframe(acc, hide_index=True, width="stretch", height=280)

        # ------------------------------------------------------------------ #
        # Downloads — one file set per LOAD_ID
        # ------------------------------------------------------------------ #
        st.divider()
        ui.section("Files", "07", "a separate set for every LOAD_ID")
        ids = E.load_ids(res)
        src_map = st.session_state.get("doc_bytes", {})

        if not ids:
            ui.empty("Nothing to download", "No document was picked.", "🗂️")
        else:
            attach_src = st.checkbox("Add the Invoice / DC pages to the pick sheet PDF",
                                     value=True)

            cache_key = (res["run_id"], bool(attach_src))
            if st.session_state.get("bundle_key") != cache_key:
                with st.spinner("Building the PDFs and Excel files..."):
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
                        h1.caption(f"Saved as `{b['safe']}` - a filename cannot contain `/`")
                    h2.download_button("WMS Excel", data=b["xlsx"],
                                       file_name=f"{b['safe']}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument."
                                            "spreadsheetml.sheet",
                                       width="stretch", key=f"x_{b['safe']}")
                    h3.download_button("Pick sheet", data=b["pdf"],
                                       file_name=f"{b['safe']}.pdf",
                                       mime="application/pdf", width="stretch",
                                       key=f"p_{b['safe']}")

            ui.eyebrow("Everything together")
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
                pick_chart = (PP.pick_line_chart_png(m_alloc) if CHART_LINE
                              else PP.pick_chart_png(m_alloc, "Picked qty by item"))
                e1, e2 = st.columns([2, 1])
                subject = e1.text_input("Subject", value=subj0, key="mail_subject")
                add_att = e2.checkbox("Attach Excel + PDF", value=True)
                body = st.text_area("Body", value=body0, height=240, key="mail_body")

                to_list = PP._addr_list(mail_to)
                cc_list = PP._addr_list(mail_cc)
                if not to_list:
                    st.warning("No To address - add one under sidebar → Email.")
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
                    st.caption("Carries the body only - no attachments.")
                with mc2:
                    st.download_button(
                        "Download draft (.eml)",
                        data=PP.build_eml(to_list, subject, body, html0, cc_list,
                                          sender=mail_from, attachments=atts,
                                          inline_png=pick_chart),
                        file_name=f"PICK_{E.safe_name(pick_for_mail[0])}"
                                  f"{'_+' + str(len(pick_for_mail) - 1) if len(pick_for_mail) > 1 else ''}.eml",
                        mime="message/rfc822", width="stretch")
                    st.caption("Double-click it and Outlook / Mail opens a draft with "
                               "the attachments and the chart already in place.")
                if pick_chart:
                    with st.expander("The chart that goes into the email",
                                     expanded=False):
                        _dk = ui.theme_type() == "dark"
                        st.image(PP.pick_line_chart_png(m_alloc, dark=_dk) if CHART_LINE
                                 else PP.pick_chart_png(m_alloc, "Picked qty by item",
                                                        dark=_dk),
                                 width="content")

        if gs_ready and not autosave and st.session_state.get("saved") != res["run_id"]:
            if st.button("Save to database"):
                try:
                    import gsheet
                    r = gsheet.save_run(sa_info, sheet_key, res, res["cfg"],
                                        note=res.get("note", ""), owner=USER)
                    st.session_state["saved"] = res["run_id"]
                    st.success("Saved")
                    st.markdown(f"[Open the sheet]({r['url']})")
                except Exception as ex:
                    st.error(f"Save error: {ex}")




def _register_frames():
    """Register data for the Dashboard and the Register tab — read once."""
    empty = (pd.DataFrame(columns=R.SUMMARY_COLS), pd.DataFrame(columns=R.DETAIL_COLS))
    if gs_ready:
        if "reg_cache" not in st.session_state:
            try:
                import gsheet
                st.session_state["reg_cache"] = gsheet.read_register(sa_info, sheet_key)
            except Exception as ex:
                st.session_state["reg_cache"] = empty
                st.session_state["reg_error"] = str(ex)
        a, b = st.session_state["reg_cache"]
    elif st.session_state.get("reg_run"):
        a, b = st.session_state["reg_run"]
    else:
        a, b = empty
    a = a if a is not None and len(a) else empty[0]
    b = b if b is not None and len(b) else empty[1]
    return a, b


# =========================================================================== #
# TAB — Dashboard (pending vs picked)
# =========================================================================== #
with tab_dash:
    ui.section("Pending vs picked", hint="how much is still to go")
    d_sum, d_det = _register_frames()

    if not len(d_sum):
        ui.empty("Nothing to show yet",
                 "Generate a pick and this fills up from the invoice register.", "📊")
    else:
        with st.expander("Filter", expanded=False):
            fa, fb, fc = st.columns(3)
            dates = d_sum["TAX_INVOICE_DATE"].map(R.parse_date).dropna()
            lo = dates.min().date() if len(dates) else datetime.now().date()
            hi = dates.max().date() if len(dates) else datetime.now().date()
            d_from = fa.date_input("Invoice date from", value=lo, key="dash_from")
            d_to = fb.date_input("to", value=hi, key="dash_to")
            d_types = fc.multiselect("Document type",
                                     sorted(d_sum["DOC_TYPE"].astype(str).unique()),
                                     key="dash_types")
        dash = R.dashboard(d_sum, d_from, d_to, d_types or None)
        k = dash["kpi"]

        # ---- the answer, first ----
        pct = k["pct"] / 100.0
        st.progress(min(1.0, max(0.0, pct)),
                    text=f"{k['picked']} of {k['total']} invoices picked · "
                         f"{k['pct']:.0f}%")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Pending invoices", k["pending"], delta_color="inverse")
        m2.metric("Pending qty", E._qty_str(k["qty_pending"]), delta_color="inverse")
        m3.metric("Picked qty", E._qty_str(k["qty_picked"]))
        m4.metric("Oldest pending", f"{k['oldest']} d",
                  help="Days since the invoice date of the oldest one still waiting.")

        if not k["pending"]:
            st.success("Nothing pending — every invoice in this range is picked.")
        else:
            c1, c2 = st.columns([1.1, 1])
            with c1:
                ui.eyebrow("Why they are waiting")
                br = dash["by_reason"]
                if len(br):
                    st.bar_chart(br.set_index("REASON")["QTY"], horizontal=True,
                                 color="#F0A81C", height=max(180, 46 * len(br)))
                    st.dataframe(br, hide_index=True, width="stretch")
            with c2:
                ui.eyebrow("Who is waiting")
                bc = dash["by_customer"].head(8)
                if len(bc):
                    st.bar_chart(bc.set_index("CUSTOMER_NAME")["QTY"], horizontal=True,
                                 color="#3E8FD0", height=max(180, 46 * len(bc)))
                    st.dataframe(bc, hide_index=True, width="stretch")

            ui.eyebrow("Oldest first — chase these")
            pend = dash["pending"]
            st.dataframe(pend, hide_index=True, width="stretch", height=340,
                         column_config={
                             "DAYS": st.column_config.NumberColumn(
                                 "Days", help="Since the invoice date", format="%d d"),
                             "QTY": st.column_config.NumberColumn("Qty"),
                             "SHORT_QTY": st.column_config.NumberColumn("Still to pick")})

            st.caption("A pending invoice only clears when it is picked — upload it "
                       "again on the Pick tab and the register updates itself.")

        # ---- every report, from here, under the same filter ----
        st.divider()
        ui.eyebrow("Download")
        d_stamp = datetime.now().strftime("%Y%m%d_%H%M")
        XL = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        dash_sum = dash["summary"]
        dash_det = R.details_for(d_det, dash["invoices"])
        pend = dash["pending"]

        q1, q2, q3, q4 = st.columns(4)
        q1.download_button("Summary report (Excel)", data=R.summary_excel(dash_sum),
                           file_name=f"Invoice_Summary_{d_stamp}.xlsx", mime=XL,
                           width="stretch", key="dash_sum_xl")
        q2.download_button("Details report (Excel)", data=R.details_excel(dash_det),
                           file_name=f"Invoice_Details_{d_stamp}.xlsx", mime=XL,
                           width="stretch", key="dash_det_xl")
        q3.download_button("Pending report (Excel)", data=R.pending_excel(dash),
                           file_name=f"Pending_{d_stamp}.xlsx", mime=XL,
                           width="stretch", key="dash_pend_xl",
                           disabled=not k["pending"])
        q4.download_button("Pending list (CSV)", data=pend.to_csv(index=False).encode(),
                           file_name=f"Pending_{d_stamp}.csv", mime="text/csv",
                           width="stretch", key="dash_pend_csv",
                           disabled=not k["pending"])
        st.caption(f"The filter above applies to all four — "
                   f"{len(dash_sum)} invoices · {len(dash_det)} lines · "
                   f"{k['pending']} pending.")


# =========================================================================== #
# TAB — Invoice register
# =========================================================================== #
with tab_reg:
    ui.section("Invoice register",
               hint="every document that has been uploaded, picked or not")

    reg_s, reg_d = _register_frames()
    if st.session_state.get("reg_error"):
        st.warning(f"Register read error: {st.session_state['reg_error']}")
    elif not gs_ready and st.session_state.get("reg_run"):
        st.info("No database connected — showing this session's run only.")

    if not len(reg_s):
        ui.empty("Nothing registered yet",
                 "Generate a pick and every uploaded document lands here.", "🗒️")
    else:
        k1, k2, k3, k4 = st.columns(4)
        yes = (reg_s["KORBER_PICK"].astype(str) == "Yes").sum()
        k1.metric("Invoices", len(reg_s))
        k2.metric("Körber picked", int(yes))
        k3.metric("Not picked", int(len(reg_s) - yes), delta_color="inverse")
        k4.metric("MRP", int((reg_s["MRP"].astype(str) == "Yes").sum()))

        f1, f2, f3 = st.columns([2, 1, 1])
        rq = f1.text_input("Search", placeholder="invoice no · customer · item · remark",
                           key="reg_q")
        fk = f2.multiselect("Körber Pick", ["Yes", "No"], key="reg_k")
        fm = f3.multiselect("MRP", ["Yes", "No"], key="reg_m")

        vs, vd = reg_s.copy(), reg_d.copy()
        if fk:
            vs = vs[vs["KORBER_PICK"].astype(str).isin(fk)]
        if fm:
            vs = vs[vs["MRP"].astype(str).isin(fm)]
        if rq.strip():
            vs = R.search(vs, rq)
        vd = vd[vd["TAX_INVOICE_NO"].astype(str).isin(set(vs["TAX_INVOICE_NO"].astype(str)))]

        t_sum, t_det = st.tabs(["Summary", "Details"])
        with t_sum:
            st.dataframe(vs[R.SUMMARY_COLS[:8]], hide_index=True, width="stretch",
                         height=430)
            with st.expander("All columns"):
                st.dataframe(vs, hide_index=True, width="stretch", height=380)
        with t_det:
            st.caption("One row per document line, with the pallets it came off.")
            st.dataframe(vd, hide_index=True, width="stretch", height=430)

        ui.eyebrow("Download")
        stamp2 = datetime.now().strftime("%Y%m%d_%H%M")
        dl1, dl2 = st.columns(2)
        dl1.download_button("Summary report (Excel)", data=R.summary_excel(vs),
                            file_name=f"Invoice_Summary_{stamp2}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument."
                                 "spreadsheetml.sheet", width="stretch")
        dl2.download_button("Details report (Excel)", data=R.details_excel(vd),
                            file_name=f"Invoice_Details_{stamp2}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument."
                                 "spreadsheetml.sheet", width="stretch")
        st.caption(f"Downloading {len(vs)} invoices and {len(vd)} lines — the filters "
                   "above apply to both files. The **Dashboard** tab has the same two "
                   "reports filtered by invoice date instead.")

        # ---- remarks are operational notes, so let them be edited ----
        with st.expander("Edit remarks"):
            st.caption("A remark on a picked invoice is kept; picking a `No` again "
                       "clears it automatically.")
            ed = st.data_editor(
                vs[["TAX_INVOICE_NO", "CUSTOMER_NAME", "KORBER_PICK", "REMARK"]],
                hide_index=True, width="stretch", height=280, key="reg_remark",
                column_config={
                    "TAX_INVOICE_NO": st.column_config.TextColumn(disabled=True),
                    "CUSTOMER_NAME": st.column_config.TextColumn(disabled=True),
                    "KORBER_PICK": st.column_config.TextColumn(disabled=True)})
            if st.button("Save remarks", disabled=not gs_ready):
                try:
                    import gsheet
                    upd = reg_s.copy()
                    idx = {str(v): i for i, v in enumerate(upd["TAX_INVOICE_NO"])}
                    for _, r in ed.iterrows():
                        i = idx.get(str(r["TAX_INVOICE_NO"]))
                        if i is not None:
                            upd.iat[i, upd.columns.get_loc("REMARK")] = str(r["REMARK"])
                    gsheet.save_register(sa_info, sheet_key, upd.head(0), reg_d.head(0),
                                         owner=USER)
                    book = gsheet.open_book(sa_info, sheet_key)
                    gsheet._replace_ws(book, gsheet.WS_INV_SUM, R.SUMMARY_COLS, upd)
                    gsheet.cache_clear(sheet_key)
                    st.session_state.pop("reg_cache", None)
                    st.toast("Remarks saved", icon="📝")
                    st.rerun()
                except Exception as ex:
                    st.error(f"Save error: {ex}")

    # ---- the MRP rule lives here ----
    with st.expander("MRP contacts — the rule behind the Yes / No"):
        st.caption("A document is **MRP = Yes** when its `Contact Person` / `Email` "
                   "matches one of these. Email wins when the document carries one.")
        mrp_ed = st.data_editor(pd.DataFrame(MRP_CONTACTS, columns=["NAME", "EMAIL"]),
                                num_rows="dynamic", width="stretch", hide_index=True,
                                key="mrp_editor")
        m1, m2 = st.columns(2)
        if m1.button("Save MRP contacts", type="primary", width="stretch"):
            st.session_state["mrp_contacts"] = R.load_contacts(
                mrp_ed.to_dict("records"))
            if gs_ready:
                try:
                    import gsheet
                    gsheet.save_setting(sa_info, sheet_key, "MRP_CONTACTS",
                                        R.dump_contacts(mrp_ed))
                except Exception as ex:
                    st.error(f"Save error: {ex}")
            st.toast("MRP contacts saved — re-run a pick to apply", icon="✅")
            st.rerun()
        if m2.button("Reset to default", width="stretch"):
            st.session_state["mrp_contacts"] = R.load_contacts(None)
            if gs_ready:
                try:
                    import gsheet
                    gsheet.save_setting(sa_info, sheet_key, "MRP_CONTACTS", "")
                except Exception:
                    pass
            st.rerun()
        st.caption("Changing this does not rewrite invoices already in the register — "
                   "it applies from the next pick run.")

# =========================================================================== #
# TAB — Loads (download / delete by LOAD_ID)
# =========================================================================== #
with tab_loads:
    ui.section("Load manager", hint="look up a LOAD_ID · download · delete")

    if not gs_ready:
        ui.empty("Database not connected",
                 "Add the Google Sheet to secrets.toml and saved loads appear here.",
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
        picked = lc2.selectbox("or pick from the list", opts, key="load_sel")
        lc3.write("")
        if lc3.button("Refresh", width="stretch"):
            gsheet.cache_clear(sheet_key)
            st.rerun()

        load_id = (typed.strip() or picked.strip())

        with st.expander(f"Saved loads ({len(reg)})", expanded=not load_id):
            if len(reg):
                st.dataframe(reg, hide_index=True, width="stretch", height=300)
            else:
                ui.empty("No loads yet",
                         "Generate a document on the Pick tab and save it.", "🗂️")

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
                ui.empty(f"{load_id} is not in the database",
                         "Check the LOAD_ID, or choose one from the list.", "🔍")
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
                st.caption("This PDF holds the pick sheet only - the original Invoice / "
                           "DC PDF is not stored in the database.")

                ui.eyebrow("Delete")
                st.caption("Deleting removes it from the ledger and the registry, so the "
                           "**pallet balance comes back** and it can be picked again.")
                if not st.session_state.get("reset_ok"):
                    st.info("To delete, unlock with the password under "
                            "sidebar → Reset & undo.")
                else:
                    dl1, dl2 = st.columns([2, 1])
                    typed_id = dl1.text_input(f"Type `{load_id}` again to confirm",
                                              key="del_confirm")
                    dl2.write("")
                    if dl2.button("Delete load", width="stretch",
                                  type="primary"):
                        if typed_id.strip() != load_id:
                            st.error("That does not match the LOAD_ID exactly.")
                        else:
                            try:
                                with st.spinner("Deleting..."):
                                    r = gsheet.delete_load(sa_info, sheet_key, load_id,
                                                           owner=USER)
                                st.session_state.pop("reg_cache", None)
                                st.toast(f"Deleted {load_id} — register set to No",
                                         icon="🗑️")
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
               hint="upload the same item again and it updates - no duplicate row")

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
            st.success("Reloaded")
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
                    with st.expander("What will change", expanded=True):
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
                            with st.spinner("Saving to the sheet..."):
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
                        st.info("No Google Sheet connected - kept in this session only.")
                    st.rerun()
            except Exception as ex:
                st.error(f"File error: {ex}")

    # ---------------- search ----------------
    with sku_find:
        q = st.text_input("🔍 Item number / base ID / description",
                          placeholder="07011636   ·   07011636-000-440   ·   gasket epdm",
                          key="sku_q")
        base_on = st.checkbox("Base ID match - search without the suffix", value=True)
        if not len(master):
            ui.empty("SKU master is empty", "Add a file on the Upload tab.", "🏷️")
        elif q.strip():
            hit = SKU.search(master, q, base_match=base_on)
            if not len(hit):
                st.warning(f"'{q}' not found. Try the base ID on its own.")
            else:
                st.caption(f"{len(hit)} rows")
                st.dataframe(hit, hide_index=True, width="stretch", height=420)
                st.download_button("Download CSV",
                                   data=hit.to_csv(index=False).encode(),
                                   file_name="sku_search.csv", mime="text/csv")
        else:
            st.dataframe(master.head(200), hide_index=True, width="stretch",
                         height=420)
            st.caption(f"Showing the first 200 rows of {len(master)}.")

    # ---------------- edit ----------------
    with sku_edit:
        st.caption("Edit in place or add a row. The duplicate check runs "
                   "again on save.")
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
                    st.success(f"Saved · {r['rows']} rows")
                else:
                    st.success(f"Saved · {len(clean)} rows (session only)")
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
               hint="give several words and only rows holding all of them come back")

    q = st.text_input("🔍", placeholder="e.g.  P550945   ·   333262712337   ·   "
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
        frames["📊 Stock Basis"] = res_s.get("basis", pd.DataFrame())
        frames["🔒 On Pick Task"] = res_s.get("locked", pd.DataFrame())
        frames["⛔ Rejected"] = res_s["rejected"]

    src = st.multiselect(
        "Where to search",
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
                  "🔢 Qty Verify", "📊 Stock Basis", "🔒 On Pick Task", "⛔ Rejected",
                  "📄 Document lines"]:
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
        ui.empty("Type what you are looking for",
                 "Item code · LOAD ID · pallet · location · GRN · lot · plant", "🔎")
    elif not frames:
        ui.empty("Nothing to search yet", "Upload a file on the Pick tab.", "📭")
    else:
        hits = E.search_frames(q, frames)
        total = sum(len(v) for v in hits.values())
        if not total:
            st.warning(f"'{q}' is not anywhere.")
        else:
            st.caption(f"{total} rows across {len(hits)} places")
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
               hint="MODE shows whether the balance came from the ledger or a fresh count")
    inv_raw = st.session_state.get("inv_raw")
    if inv_raw is None:
        ui.empty("No inventory report", "Upload one on the Pick tab.", "📦")
    else:
        ledger = None
        if gs_ready:
            try:
                import gsheet
                ledger = gsheet.read_ledger(sa_info, sheet_key)
            except Exception as ex:
                st.warning(f"Ledger read error: {ex}")
        view = E.stock_view(inv_raw, ledger, use_ledger=use_ledger)
        f1, f2, f3, f4, f5 = st.columns([1.5, 1, 1, 1, .9])
        q = f1.text_input("Item / Base ID / Pallet search", placeholder="P550945")
        plants = f2.multiselect("Plant", sorted(view["PLANT"].dropna().unique().tolist()))
        modes = f3.multiselect("Mode", ["NEW", "LEDGER BALANCE", "NEW BASELINE"])
        pstat = f4.multiselect("Pick Id", ["FREE", "ON PICK TASK"])
        f5.write("")
        only_bal = f5.checkbox("Balance > 0", value=True)

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
        if pstat:
            v = v[v["PICK_STATUS"].isin(pstat)]
        if only_bal:
            v = v[v["BALANCE"] > 0]

        a, b, c, dcol, ecol = st.columns(5)
        a.metric("Rows", len(v))
        b.metric("Pallets", int(v["PALLET"].nunique()) if len(v) else 0)
        c.metric("Actual Qty", f"{v['ACTUAL_QTY'].sum():g}" if len(v) else "0")
        dcol.metric("Pickable Balance", f"{v['BALANCE'].sum():g}" if len(v) else "0")
        ecol.metric("On Pick Task",
                    f"{v.loc[v['PICK_STATUS'] != 'FREE', 'ACTUAL_QTY'].sum():g}"
                    if len(v) else "0",
                    help="Non-zero Pick Id — balance is forced to 0 for these.")
        st.dataframe(v, hide_index=True, width="stretch", height=520)

# =========================================================================== #
# TAB 3 — History
# =========================================================================== #
with tab_hist:
    ui.section("History", hint="the raw worksheets behind the app")
    if not gs_ready:
        ui.empty("Database not connected",
                 "Add the Google Sheet to secrets.toml.", "🔌")
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
    g1, g2, g4, g5, g3 = st.tabs(["The pick", "Output files", "Register",
                                  "Dashboard", "Setup"])

    with g1:
        st.markdown("""
**The flow**

1. Upload the **Invoice / Delivery Challan PDFs** and the **Inventory Report**.
2. Each document is parsed and put through a **completeness check** — S.No
   sequence, `Total Quantity`, `Grand Total` amount. Anything that fails is
   marked `blocked` and is not picked.
3. The plants found in the inventory are listed. **Confirm one** before the
   pick can run.
4. Items match on the **base ID** only:
   `P162400-000-140` → **`P162400`**, so inventory holding `P162400-016-140`
   still matches. A suffix only counts as a suffix when every part after the
   first is three digits — `05-47174` stays whole.
5. Only pallets with **`Pick Id = 0`** are pickable. A non-zero Pick Id means the
   WMS has already committed that pallet to another pick task — its `Status` still
   reads `Available`, so without this gate the same stock would be picked twice.
   Whatever is skipped shows up under **On a pick task** with its Pick Id.
   If a document fails *only* because of this, a **Release stock** panel appears:
   tick the confirmation, press **Release and pick**, and those pallets are used.
   The release is recorded against the document (`DOC_CHECK`), printed on the pick
   sheet and written into the email — so it is never a silent override.
6. A document is picked **only if every line can be filled**. One short line
   blocks the whole document (`STOCK SHORT`) — there is no partial picking.
   When the missing quantity does exist but is locked, the reason says so:
   *"On another pick task — 128 locked (Pick Id 1282815, …)"*.
7. The same Invoice / DC number is **never processed twice** (`DOC_REGISTRY`).
8. Every pallet touched is written to the ledger as
   `QTY_BEFORE → QTY_PICKED → QTY_BALANCE`, and a pallet can never go negative.

**Where the balance comes from**

| Situation | Meaning | Picked from |
|---|---|---|
| Pick Id ≠ 0 | on another pick task | nothing — excluded |
| Pallet not in the ledger | new pallet | `NEW` → Actual Qty |
| Actual Qty **=** ledger QTY_BEFORE | inventory report is stale | `LEDGER BALANCE` → QTY_BALANCE |
| Actual Qty **≠** ledger QTY_BEFORE | WMS has been updated | `NEW BASELINE` → Actual Qty |

**Quantity verification**

Line by line, document total, and the total inside the generated file — all
three must equal the Invoice / DC quantity. Any mismatch blocks the document.
""")

    with g2:
        st.markdown("""
**`<LOAD_ID>.xlsx`** — the WMS upload file. Every cell is **text**, empty cells
stay truly empty.

*OutBound MASTER* — `DISPLAY_ORDER_NUMBER`, `STORE_ORDER_NUMBER`,
`CUSTOMER_PO_NUMBER` and `LOAD_ID` all carry the Invoice No / DC No. Everything
else follows the template.

*OutBound Detail*

| Column | Source |
|---|---|
| DISPLAY_ORDER_NUMBER | Invoice No / DC No |
| LINE_NUMBER | 1, 2, 3, … |
| DISPLAY_ITEM_NUMBER | Inventory **Item Number**, exactly as stored |
| LOT_NUMBER | Inventory Lot Number |
| QTY | Invoice / DC quantity |
| ORDER_UOM | Inventory Uom |
| GEN_ATTRIBUTE_VALUE1–11 | Color · Size · Style · Supplier · **Plant** · Client So · Client So Line · Po Cust Dec · Customer Ref Number · Item Id · Invoice Number 1 |

**`<LOAD_ID>.pdf`** — pick sheet with a LOAD_ID QR, the pallet-by-pallet pick
list, the quantity verification and a sign-off block, followed by the pages of
the original Invoice / DC.

**`SHORT_<LOAD_ID>.pdf`** — raised when stock is short: the short lines,
required vs available vs short, a chart, and the full document lines.

**`Pick_Report_*.xlsx`** — Doc Summary · Qty Verification · Pallet Allocation ·
Pallet Balance · Stock Basis · Shortage · **On Pick Task** · Rejected Docs.

A DC number such as `333/26-27/62` is saved as `333-26-27-62` because a filename
cannot hold a slash. The LOAD_ID inside the file and the QR is the real one.
""")

    with g4:
        st.markdown("""
Every document that is uploaded is written to the **Register**, whether it was
picked or not, one row per invoice — for good.

**Summary**

| Column | Where it comes from |
|---|---|
| Tax Invoice Date | `Invoice Date` / `Delivery Challan Date` |
| Tax Invoice No. | `Invoice No.` / `Delivery Challan no.` |
| AR Invoice No. | `AR Invoice No.` (invoices only) |
| Name of Customer | `Ship To / Consignee` · on a challan, `Name of Consignee(Shipped To)` |
| Qty | total document quantity |
| Körber Pick | **Yes** when the pick generated · **No** with the reason as the remark |
| MRP | **Yes** when the document's Contact Person / Email is one of the MRP contacts |

**Details** — one row per document line: item, quantity, and the pallets,
locations and lots it came off.

**How Körber Pick moves**

* first upload, stock short → `No` + *"STOCK SHORT — L2 X006252 (need 999, have 6)"*
* picked later → `Yes`, remark cleared automatically, same row
* load deleted on the **Loads** tab → back to `No` + *"Load deleted · <time>"*

A `Yes` is never quietly pushed back to `No` by a later run that skipped the
invoice as a duplicate — only deleting the load does that.

**MRP contacts** live under *MRP contacts — the rule behind the Yes / No* on the
Register tab. The default is `Sharma, Rahul` / `rahul.sharma1@donaldson.com`;
add or change rows there and it applies from the next run.

Both reports download as **separate Excel files**, and the filters on screen
apply to what you get.
""")

    with g5:
        st.markdown("""
**Pending vs picked — how much is still left.**

It reads the invoice register, so it covers every document ever uploaded, not
just the last run.

| | |
|---|---|
| Progress bar | picked invoices out of the total, as a percentage |
| Pending invoices / Pending qty | what is still to go — counted both ways, because one pending invoice for 1 000 units is not the same problem as ten for 2 units |
| Oldest pending | days since the invoice date of the oldest one still waiting |
| Why they are waiting | pending quantity grouped by reason — stock short · on another pick task · duplicate · document incomplete · not picked yet |
| Who is waiting | pending quantity by customer, biggest first |
| Oldest first | the chase list — days, invoice, customer, qty, still to pick, reason |

**Filter** by invoice date range and document type; every number and both
downloads follow the filter.

**Downloads — all four sit here**

| File | What is in it |
|---|---|
| Summary report (Excel) | one row per invoice, the whole register |
| Details report (Excel) | one row per document line |
| Pending report (Excel) | five sheets — Status · Pending · By reason · By customer · Picked |
| Pending list (CSV) | just the chase list |

The date and document-type filter applies to all four. The **Register** tab has
the same Summary and Details reports with its own filters (search · Körber Pick ·
MRP) — same data, pick whichever filter suits.

A pending invoice clears itself: upload it again on the **Pick** tab and once it
picks, the register flips it to `Yes` and it drops off this screen.
""")

    with g3:
        st.markdown("""
**`.streamlit/secrets.toml`**

```toml
[gcp_service_account]
type           = "service_account"
project_id     = "..."
private_key_id = "..."
private_key    = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email   = "xxx@yyy.iam.gserviceaccount.com"
client_id      = "..."
token_uri      = "https://oauth2.googleapis.com/token"

[google_sheet]
data_sheet  = "https://docs.google.com/spreadsheets/d/<KEY>/edit"
auto_save   = true
wh_id       = "INMM01"
client_code = "INM0DONA"
mail_to     = "stores@example.com"
mail_from   = "you@example.com"

[app]
reset_password = "..."      # optional, overrides the built-in one
```

Share the sheet with the **service-account email as an Editor**, then press
**Set up worksheets** in the sidebar once.

**Worksheets**

`OUTBOUND_MASTER` · `OUTBOUND_DETAIL` · `PALLET_LEDGER` · `DOC_REGISTRY` ·
`REJECTED_LOG` · `RUN_LOG` · `SKU_MASTER` · `INVOICE_SUMMARY` ·
`INVOICE_DETAIL` · `APP_SETTINGS` · `_LOCKS`

**Several people at once**

Reads are cached, writes take a lock, and the duplicate check runs again inside
that lock — so two people saving the same invoice at the same moment still
produce one order. Failed API calls back off and retry. Watch it under
sidebar → **API & multi-user**.
""")

with SLOT_TOPBAR.container():
    _draw_topbar()
with SLOT_RAIL.container():
    _draw_rail()

ui.footnote("base-ID match · plant confirm · all-or-nothing per document · "
            "duplicate gate · pallet-level ledger · qty verified · WMS cells = text")
