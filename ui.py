"""
ui.py — Design system for the OutBound Pick Generator
=====================================================
Direction: a warehouse operations console, not a dashboard.

  Ink + hi-vis      navy ink on light paper; safety amber for the one action
                    that matters on each screen; red means STOP, never decoration.
  Signage type      Barlow Condensed (labels, in the voice of rack signage)
                    Barlow (reading) · IBM Plex Mono (every code, always)
  Codes are mono    LOAD IDs, pallets, locations are read character by character.
  Signature         hazard rule — the striped tape from the warehouse floor,
                    used only under the top bar and on blocked documents.
"""
from __future__ import annotations

from typing import Any, Iterable

import streamlit as st

# --------------------------------------------------------------------------- #
# tokens
# --------------------------------------------------------------------------- #
INK = "#0B1A2B"
INK_SOFT = "#3D4E62"
MUTED = "#6B7B8C"
PAPER = "#F4F6F8"
SURFACE = "#FFFFFF"
SURFACE_2 = "#EDF1F5"
RULE = "#DCE3EA"
ACCENT = "#F0A81C"
OK = "#0F8A5F"
WARN = "#B4740A"
STOP = "#C62828"
INFO = "#1B6CA8"

_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Barlow:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root{{
  --ink:{INK}; --ink-soft:{INK_SOFT}; --muted:{MUTED};
  --paper:{PAPER}; --surface:{SURFACE}; --surface-2:{SURFACE_2}; --rule:{RULE};
  --accent:{ACCENT}; --ok:{OK}; --warn:{WARN}; --stop:{STOP}; --info:{INFO};
  --shadow:0 1px 2px rgba(11,26,43,.06), 0 8px 24px -18px rgba(11,26,43,.35);
  --r:10px;
  --display:'Barlow Condensed','Barlow',system-ui,sans-serif;
  --body:'Barlow',system-ui,-apple-system,'Segoe UI',sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,'SF Mono',monospace;
}}

/* ---------- canvas ---------- */
.stApp{{ background:var(--paper); color:var(--ink); }}
html, body, [class*="css"], .stMarkdown, p, li, label,
[data-testid="stWidgetLabel"] p{{ font-family:var(--body); color:var(--ink); }}
.block-container{{ padding-top:.9rem; padding-bottom:3rem; max-width:1440px; }}
h1,h2,h3,h4,h5{{ font-family:var(--display); color:var(--ink);
  letter-spacing:.01em; font-weight:700; }}
h2{{ font-size:1.28rem; }} h3{{ font-size:1.08rem; }}
code, kbd{{ font-family:var(--mono); font-size:.86em;
  background:var(--surface-2); color:var(--ink);
  padding:1px 5px; border-radius:5px; border:1px solid var(--rule); }}
a{{ color:var(--info); text-decoration:none; }}
a:hover{{ text-decoration:underline; }}
[data-testid="stCaptionContainer"], .stCaption, small{{ color:var(--muted) !important;
  font-size:.815rem; }}

/* ---------- top bar ---------- */
.topbar{{ background:var(--surface); border:1px solid var(--rule);
  border-radius:var(--r) var(--r) 0 0; padding:14px 18px 13px;
  display:flex; align-items:center; justify-content:space-between; gap:18px;
  flex-wrap:wrap; box-shadow:var(--shadow); }}
.tb-brand{{ display:flex; align-items:center; gap:13px; min-width:0; }}
.tb-mark{{ width:40px; height:40px; border-radius:9px; flex:0 0 auto;
  background:var(--ink); color:var(--accent); display:grid; place-items:center;
  font-family:var(--display); font-weight:700; font-size:1.02rem;
  letter-spacing:.02em; }}
.tb-title{{ font-family:var(--display); font-weight:700; font-size:1.36rem;
  line-height:1.05; letter-spacing:.005em; color:var(--ink); }}
.tb-sub{{ font-family:var(--mono); font-size:.715rem; color:var(--muted);
  margin-top:3px; letter-spacing:.01em; }}
.tb-chips{{ display:flex; gap:7px; flex-wrap:wrap; align-items:center; }}
.chip{{ display:inline-flex; align-items:center; gap:7px; padding:5px 11px;
  border:1px solid var(--rule); border-radius:999px; background:var(--surface);
  font-family:var(--mono); font-size:.7rem; color:var(--ink-soft);
  white-space:nowrap; }}
.chip b{{ color:var(--ink); font-weight:600; }}
.chip .dot{{ width:7px; height:7px; border-radius:50%; background:var(--muted); }}
.chip.ok .dot{{ background:var(--ok); box-shadow:0 0 0 3px rgba(15,138,95,.16); }}
.chip.warn .dot{{ background:var(--accent); box-shadow:0 0 0 3px rgba(240,168,28,.2); }}
.chip.stop .dot{{ background:var(--stop); box-shadow:0 0 0 3px rgba(198,40,40,.16); }}

/* signature: hazard rule */
.hazard{{ height:5px; border-radius:0 0 var(--r) var(--r); margin-bottom:14px;
  background:repeating-linear-gradient(135deg,
    var(--accent) 0 11px, var(--ink) 11px 22px); opacity:.92; }}

/* ---------- section head ---------- */
.sec{{ display:flex; align-items:baseline; gap:11px; margin:20px 0 8px; }}
.sec .num{{ font-family:var(--mono); font-size:.74rem; font-weight:600;
  color:var(--muted); border:1px solid var(--rule); border-radius:6px;
  padding:2px 7px; background:var(--surface); }}
.sec .t{{ font-family:var(--display); font-weight:700; font-size:1.02rem;
  text-transform:uppercase; letter-spacing:.055em; color:var(--ink); }}
.sec .h{{ font-size:.83rem; color:var(--muted); }}
.eyebrow{{ font-family:var(--display); font-size:.72rem; font-weight:600;
  text-transform:uppercase; letter-spacing:.15em; color:var(--muted);
  margin-bottom:5px; }}

/* ---------- step rail ---------- */
.rail{{ display:flex; gap:0; background:var(--surface); border:1px solid var(--rule);
  border-radius:var(--r); overflow:hidden; margin-bottom:6px; }}
.rail .st{{ flex:1; padding:11px 14px; border-right:1px solid var(--rule);
  min-width:0; position:relative; }}
.rail .st:last-child{{ border-right:0; }}
.rail .n{{ font-family:var(--mono); font-size:.68rem; font-weight:600;
  color:var(--muted); }}
.rail .l{{ font-family:var(--display); font-size:.95rem; font-weight:600;
  text-transform:uppercase; letter-spacing:.05em; color:var(--ink-soft);
  line-height:1.15; margin-top:2px; }}
.rail .v{{ font-family:var(--mono); font-size:.71rem; color:var(--muted);
  margin-top:3px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.rail .st.done{{ background:rgba(15,138,95,.055); }}
.rail .st.done .l{{ color:var(--ink); }}
.rail .st.done .n::after{{ content:" ✓"; color:var(--ok); font-weight:700; }}
.rail .st.now{{ background:rgba(240,168,28,.1);
  box-shadow:inset 0 -3px 0 var(--accent); }}
.rail .st.now .l{{ color:var(--ink); }}
.rail .st.now .n{{ color:var(--ink); }}

/* ---------- doc / status cards ---------- */
.dcard{{ background:var(--surface); border:1px solid var(--rule);
  border-left:4px solid var(--rule); border-radius:8px; padding:10px 14px;
  margin-bottom:7px; }}
.dcard.ok{{ border-left-color:var(--ok); }}
.dcard.stop{{ border-left:4px solid transparent;
  border-image:repeating-linear-gradient(135deg, {STOP} 0 6px, {INK} 6px 12px) 1;
  background:rgba(198,40,40,.035); }}
.dcard.warn{{ border-left-color:var(--accent); background:rgba(240,168,28,.05); }}
.dcard .hd{{ display:flex; align-items:center; gap:9px; flex-wrap:wrap; }}
.dcard .id{{ font-family:var(--mono); font-weight:600; font-size:.94rem;
  color:var(--ink); }}
.dcard .meta{{ font-size:.83rem; color:var(--muted); }}
.dcard ul{{ margin:6px 0 0 2px; padding-left:16px; }}
.dcard li{{ font-size:.83rem; color:var(--ink-soft); margin:1px 0; }}

/* stamp badges */
.stamp{{ display:inline-block; font-family:var(--display); font-weight:700;
  font-size:.7rem; text-transform:uppercase; letter-spacing:.11em;
  padding:2px 8px; border-radius:4px; border:1.5px solid currentColor;
  line-height:1.5; }}
.stamp.ok{{ color:var(--ok); }}
.stamp.stop{{ color:var(--stop); }}
.stamp.warn{{ color:var(--warn); }}
.stamp.info{{ color:var(--info); }}
.stamp.mute{{ color:var(--muted); }}

/* ---------- empty state ---------- */
.empty{{ background:var(--surface); border:1px dashed #C6D0DA; border-radius:var(--r);
  padding:26px 22px; text-align:center; }}
.empty .i{{ font-size:1.5rem; }}
.empty .t{{ font-family:var(--display); font-weight:700; font-size:1.02rem;
  text-transform:uppercase; letter-spacing:.05em; margin-top:6px; color:var(--ink); }}
.empty .b{{ font-size:.87rem; color:var(--muted); margin-top:4px; }}

/* ---------- streamlit widgets ---------- */
[data-testid="stTabs"] [role="tablist"]{{ gap:2px; border-bottom:1px solid var(--rule);
  background:transparent; }}
[data-testid="stTabs"] [role="tab"]{{ padding:8px 15px; border-radius:8px 8px 0 0; }}
[data-testid="stTabs"] [role="tab"] p{{ font-family:var(--display); font-weight:600;
  font-size:.95rem; letter-spacing:.035em; text-transform:uppercase;
  color:var(--muted); }}
[data-testid="stTabs"] [role="tab"]:hover p{{ color:var(--ink-soft); }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{ background:var(--surface);
  border:1px solid var(--rule); border-bottom-color:var(--surface);
  margin-bottom:-1px; }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p{{ color:var(--ink); }}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{{ display:none; }}

[data-testid="stMetric"]{{ background:var(--surface); border:1px solid var(--rule);
  border-radius:9px; padding:11px 13px 9px; }}
[data-testid="stMetricLabel"] p{{ font-family:var(--display); font-weight:600;
  font-size:.74rem; text-transform:uppercase; letter-spacing:.09em;
  color:var(--muted); }}
[data-testid="stMetricValue"]{{ font-family:var(--mono); font-weight:600;
  font-size:1.32rem; color:var(--ink); }}

.stButton>button, .stDownloadButton>button, .stLinkButton>a{{
  font-family:var(--display); font-weight:600; font-size:.95rem;
  letter-spacing:.035em; text-transform:uppercase; border-radius:8px;
  border:1px solid var(--rule); background:var(--surface); color:var(--ink);
  padding:.42rem 1rem; transition:border-color .12s, background .12s; }}
.stButton>button:hover, .stDownloadButton>button:hover, .stLinkButton>a:hover{{
  border-color:var(--ink); background:var(--surface-2); color:var(--ink); }}
.stButton>button[kind="primary"]{{ background:var(--accent); border-color:#D9930C;
  color:var(--ink); }}
.stButton>button[kind="primary"]:hover{{ background:#E39C10; border-color:var(--ink); }}
.stDownloadButton>button{{ border-color:#C3CEDA; }}
.stButton>button:focus-visible, .stDownloadButton>button:focus-visible,
.stLinkButton>a:focus-visible{{ outline:2px solid var(--info); outline-offset:2px; }}
.stButton>button:disabled, .stDownloadButton>button:disabled{{ opacity:.45; }}

[data-testid="stSidebar"]{{ background:var(--surface); border-right:1px solid var(--rule); }}
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3{{ font-size:.95rem;
  text-transform:uppercase; letter-spacing:.07em; }}
[data-testid="stSidebar"] [data-testid="stMetric"]{{ padding:7px 9px; }}
[data-testid="stSidebar"] [data-testid="stMetricValue"]{{ font-size:1rem; }}

.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea{{
  font-family:var(--body); border-radius:7px; border:1px solid #C9D3DC;
  background:var(--surface); color:var(--ink); }}
.stTextInput input:focus, .stTextArea textarea:focus{{ border-color:var(--info);
  box-shadow:0 0 0 3px rgba(27,108,168,.13); }}
[data-baseweb="select"]>div{{ border-radius:7px; border-color:#C9D3DC;
  background:var(--surface); }}
[data-baseweb="tag"]{{ background:var(--ink) !important; border-radius:5px !important; }}
[data-baseweb="tag"] span{{ font-family:var(--mono); font-size:.76rem; }}

[data-testid="stFileUploaderDropzone"]{{ background:var(--surface);
  border:1.5px dashed #B9C6D2; border-radius:9px; }}
[data-testid="stFileUploaderDropzone"]:hover{{ border-color:var(--accent); }}

[data-testid="stExpander"]{{ border:1px solid var(--rule); border-radius:9px;
  background:var(--surface); }}
[data-testid="stExpander"] summary{{ font-family:var(--display); font-weight:600;
  letter-spacing:.03em; }}
[data-testid="stExpander"] summary p{{ font-family:var(--display); font-weight:600; }}

[data-testid="stDataFrame"], [data-testid="stDataFrameResizable"]{{
  border:1px solid var(--rule); border-radius:9px; }}
[data-testid="stDataFrame"] [role="columnheader"]{{ font-family:var(--display);
  text-transform:uppercase; letter-spacing:.04em; }}

[data-testid="stVerticalBlockBorderWrapper"]:has(>div>[data-testid="stVerticalBlock"]){{
  border-radius:var(--r); }}
div[data-testid="stAlert"]{{ border-radius:9px; border:1px solid var(--rule); }}
hr{{ border-color:var(--rule); margin:1.1rem 0; }}
[data-testid="stProgress"] div[role="progressbar"]>div{{ background:var(--accent); }}
.footnote{{ font-family:var(--mono); font-size:.72rem; color:var(--muted);
  border-top:1px solid var(--rule); padding-top:11px; margin-top:24px; }}
</style>
"""


def inject() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# components
# --------------------------------------------------------------------------- #
def _esc(v: Any) -> str:
    return (str("" if v is None else v).replace("&", "&amp;")
            .replace("<", "&lt;").replace(">", "&gt;"))


def topbar(title: str, subtitle: str, chips: Iterable[dict] | None = None,
           mark: str = "EFL") -> None:
    """Brand + live environment chips. Hazard rule underneath = the signature."""
    bits = []
    for c in (chips or []):
        tone = c.get("tone", "")
        label = f"{_esc(c.get('label',''))} " if c.get("label") else ""
        bits.append(f"<span class='chip {tone}'><span class='dot'></span>"
                    f"{label}<b>{_esc(c.get('value',''))}</b></span>")
    st.markdown(
        f"""<div class="topbar">
  <div class="tb-brand">
    <div class="tb-mark">{_esc(mark)}</div>
    <div>
      <div class="tb-title">{_esc(title)}</div>
      <div class="tb-sub">{_esc(subtitle)}</div>
    </div>
  </div>
  <div class="tb-chips">{''.join(bits)}</div>
</div><div class="hazard"></div>""", unsafe_allow_html=True)


def rail(steps: list[dict]) -> None:
    """
    Numbered process rail — the pick really is a sequence, so the numbers
    carry information: you cannot generate before the plant is confirmed.
    steps: [{"label":..., "value":..., "state":"done|now|todo"}]
    """
    cells = []
    for i, s in enumerate(steps, start=1):
        cells.append(
            f"<div class='st {s.get('state','todo')}'>"
            f"<div class='n'>{i:02d}</div>"
            f"<div class='l'>{_esc(s.get('label',''))}</div>"
            f"<div class='v'>{_esc(s.get('value','—'))}</div></div>")
    st.markdown(f"<div class='rail'>{''.join(cells)}</div>", unsafe_allow_html=True)


def section(title: str, num: str | int | None = None, hint: str = "") -> None:
    n = f"<span class='num'>{_esc(num)}</span>" if num is not None else ""
    h = f"<span class='h'>{_esc(hint)}</span>" if hint else ""
    st.markdown(f"<div class='sec'>{n}<span class='t'>{_esc(title)}</span>{h}</div>",
                unsafe_allow_html=True)


def eyebrow(text: str) -> None:
    st.markdown(f"<div class='eyebrow'>{_esc(text)}</div>", unsafe_allow_html=True)


def stamp(text: str, tone: str = "mute") -> str:
    return f"<span class='stamp {tone}'>{_esc(text)}</span>"


def doc_card(doc_id: str, meta: str, tone: str = "ok", badge: str = "",
             notes: list[str] | None = None) -> None:
    b = stamp(badge, tone) if badge else ""
    li = "".join(f"<li>{_esc(n)}</li>" for n in (notes or []))
    ul = f"<ul>{li}</ul>" if li else ""
    st.markdown(
        f"""<div class="dcard {tone}"><div class="hd">{b}
  <span class="id">{_esc(doc_id)}</span>
  <span class="meta">{_esc(meta)}</span></div>{ul}</div>""",
        unsafe_allow_html=True)


def empty(title: str, body: str = "", icon: str = "📦") -> None:
    st.markdown(f"""<div class="empty"><div class="i">{icon}</div>
<div class="t">{_esc(title)}</div><div class="b">{_esc(body)}</div></div>""",
                unsafe_allow_html=True)


def footnote(text: str) -> None:
    st.markdown(f"<div class='footnote'>{_esc(text)}</div>", unsafe_allow_html=True)
