"""
ui.py — Design system for the OutBound Pick Generator
=====================================================
Direction: a warehouse operations console, not a dashboard.

  Signage type      Barlow Condensed (labels, in the voice of rack signage)
                    Barlow (reading) · IBM Plex Mono (every code, always)
  Codes are mono    LOAD IDs, pallets, locations are read character by character.
  Signature         hazard rule — the striped tape from the warehouse floor,
                    used only under the top bar and on blocked documents.

THEME SAFETY — why there are almost no fixed colours in here
------------------------------------------------------------
Streamlit owns the page background and paints `st.dataframe` on a canvas that
CSS cannot reach. If this file hard-coded "white card / dark text" and the app
ran on a dark theme, half the screen went dark-on-dark.

So every surface, border and muted text colour is mixed from `currentColor`:

    background: color-mix(in srgb, currentColor 5%, transparent)

On a light theme currentColor is near-black → a faint grey tint.
On a dark theme it is near-white → a faint lift.
Body text is never coloured at all — it inherits Streamlit's own text colour,
which is guaranteed to contrast with Streamlit's own background.

Only the four signal hues are fixed (amber / green / red / blue), and they are
always used with a tint of themselves behind them, so the pair travels together.
Result: light theme, dark theme, or a theme the user picks later — all readable.
"""
from __future__ import annotations

from typing import Any, Iterable

import streamlit as st

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 2

# --------------------------------------------------------------------------- #
# the only fixed hues — chosen to sit on both near-white and near-black
# --------------------------------------------------------------------------- #
ACCENT = "#F0A81C"      # hi-vis amber — the one action that matters
ACCENT_INK = "#14202C"  # text on amber, both themes
OK = "#12A06E"
WARN = "#B4740A"
STOP = "#E24A4A"
INFO = "#3E8FD0"

_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Barlow:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root{{
  --accent:{ACCENT}; --accent-ink:{ACCENT_INK};
  --ok:{OK}; --warn:{WARN}; --stop:{STOP}; --info:{INFO};

  /* signal hues pulled toward the live text colour, so they stay legible
     on a white page AND on a dark page without a media query */
  --ok-t:color-mix(in srgb, {OK} 56%, currentColor);
  --warn-t:color-mix(in srgb, {WARN} 56%, currentColor);
  --stop-t:color-mix(in srgb, {STOP} 56%, currentColor);
  --info-t:color-mix(in srgb, {INFO} 56%, currentColor);

  /* neutrals mixed from the live text colour — theme agnostic */
  --t-soft:color-mix(in srgb, currentColor 82%, transparent);
  --t-mute:color-mix(in srgb, currentColor 70%, transparent);
  --fill-1:color-mix(in srgb, currentColor 4%, transparent);
  --fill-2:color-mix(in srgb, currentColor 8%, transparent);
  --fill-3:color-mix(in srgb, currentColor 13%, transparent);
  --line:color-mix(in srgb, currentColor 17%, transparent);
  --line-2:color-mix(in srgb, currentColor 32%, transparent);

  --r:10px;
  --display:'Barlow Condensed','Barlow',system-ui,sans-serif;
  --body:'Barlow',system-ui,-apple-system,'Segoe UI',sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,'SF Mono',monospace;
}}

/* type only — the page background stays Streamlit's */
html, body, .stMarkdown, p, li, label,
[data-testid="stWidgetLabel"] p{{ font-family:var(--body); }}
/* Streamlit's header is FIXED at 3.75rem and content scrolls under it. Its own
   default top padding is 6rem; anything below ~3.75rem clips the app bar. */
.block-container, [data-testid="stMainBlockContainer"]{{
  padding-top:4.6rem; padding-bottom:3rem; max-width:1440px; }}
@media (max-width:640px){{
  .block-container, [data-testid="stMainBlockContainer"]{{ padding-top:4.2rem; }}
}}
h1,h2,h3,h4,h5{{ font-family:var(--display); letter-spacing:.01em; font-weight:700; }}
h2{{ font-size:1.28rem; }} h3{{ font-size:1.08rem; }}
code, kbd{{ font-family:var(--mono); font-size:.86em; color:inherit;
  background:var(--fill-2); padding:1px 5px; border-radius:5px;
  border:1px solid var(--line); }}
[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p,
.stCaption, small{{ color:var(--t-mute) !important; font-size:.815rem; }}

/* ---------- top bar ---------- */
.topbar{{ background:var(--fill-1); border:1px solid var(--line);
  border-radius:var(--r) var(--r) 0 0; padding:14px 18px 13px;
  display:flex; align-items:center; justify-content:space-between; gap:18px;
  flex-wrap:wrap; }}
.tb-brand{{ display:flex; align-items:center; gap:13px; min-width:0; }}
.tb-mark{{ width:40px; height:40px; border-radius:9px; flex:0 0 auto;
  background:var(--accent); color:var(--accent-ink); display:grid;
  place-items:center; font-family:var(--display); font-weight:700;
  font-size:1.02rem; letter-spacing:.02em; }}
.tb-title{{ font-family:var(--display); font-weight:700; font-size:1.36rem;
  line-height:1.05; color:inherit; }}
.tb-sub{{ font-family:var(--mono); font-size:.715rem; color:var(--t-mute);
  margin-top:3px; }}
.tb-chips{{ display:flex; gap:7px; flex-wrap:wrap; align-items:center; }}
.chip{{ display:inline-flex; align-items:center; gap:7px; padding:5px 11px;
  border:1px solid var(--line); border-radius:999px; background:var(--fill-1);
  font-family:var(--mono); font-size:.7rem; color:var(--t-soft);
  white-space:nowrap; }}
.chip b{{ color:inherit; font-weight:600; }}
.chip .dot{{ width:7px; height:7px; border-radius:50%;
  background:color-mix(in srgb, currentColor 45%, transparent); }}
.chip.ok .dot{{ background:var(--ok); box-shadow:0 0 0 3px color-mix(in srgb,{OK} 22%,transparent); }}
.chip.warn .dot{{ background:var(--accent); box-shadow:0 0 0 3px color-mix(in srgb,{ACCENT} 26%,transparent); }}
.chip.stop .dot{{ background:var(--stop); box-shadow:0 0 0 3px color-mix(in srgb,{STOP} 22%,transparent); }}

/* signature: hazard rule — amber dashes, page shows through, so it reads on any theme */
.hazard{{ height:5px; margin-bottom:14px; border-radius:0 0 var(--r) var(--r);
  background:repeating-linear-gradient(115deg,
    var(--accent) 0 12px, transparent 12px 24px); }}

/* ---------- section head ---------- */
.sec{{ display:flex; align-items:baseline; gap:11px; margin:20px 0 8px;
  flex-wrap:wrap; }}
.sec .num{{ font-family:var(--mono); font-size:.74rem; font-weight:600;
  color:var(--t-mute); border:1px solid var(--line); border-radius:6px;
  padding:2px 7px; }}
.sec .t{{ font-family:var(--display); font-weight:700; font-size:1.02rem;
  text-transform:uppercase; letter-spacing:.055em; color:inherit; }}
.sec .h{{ font-size:.83rem; color:var(--t-mute); }}
.eyebrow{{ font-family:var(--display); font-size:.72rem; font-weight:600;
  text-transform:uppercase; letter-spacing:.15em; color:var(--t-mute);
  margin-bottom:5px; }}

/* ---------- step rail ---------- */
.rail{{ display:flex; border:1px solid var(--line); border-radius:var(--r);
  overflow:hidden; margin-bottom:6px; background:var(--fill-1); }}
.rail .st{{ flex:1; padding:11px 14px; border-right:1px solid var(--line);
  min-width:0; }}
.rail .st:last-child{{ border-right:0; }}
.rail .n{{ font-family:var(--mono); font-size:.68rem; font-weight:600;
  color:var(--t-mute); }}
.rail .l{{ font-family:var(--display); font-size:.95rem; font-weight:600;
  text-transform:uppercase; letter-spacing:.05em; color:var(--t-mute);
  line-height:1.15; margin-top:2px; }}
.rail .v{{ font-family:var(--mono); font-size:.71rem; color:var(--t-mute);
  margin-top:3px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.rail .st.done{{ background:color-mix(in srgb, {OK} 11%, transparent); }}
.rail .st.done .l{{ color:inherit; }}
.rail .st.done .n::after{{ content:" ✓"; color:var(--ok-t); font-weight:700; }}
.rail .st.now{{ background:color-mix(in srgb, {ACCENT} 14%, transparent);
  box-shadow:inset 0 -3px 0 var(--accent); }}
.rail .st.now .l, .rail .st.now .n{{ color:inherit; }}

/* ---------- document cards ---------- */
.dcard{{ background:var(--fill-1); border:1px solid var(--line);
  border-left:4px solid var(--line-2); border-radius:8px; padding:10px 14px;
  margin-bottom:7px; }}
.dcard.ok{{ border-left-color:var(--ok); }}
.dcard.warn{{ border-left-color:var(--accent);
  background:color-mix(in srgb, {ACCENT} 8%, transparent); }}
.dcard.stop{{ border-left:0; padding-left:18px; position:relative;
  background:color-mix(in srgb, {STOP} 8%, transparent); }}
.dcard.stop::before{{ content:""; position:absolute; left:0; top:0; bottom:0;
  width:4px; border-radius:8px 0 0 8px;
  background:repeating-linear-gradient(135deg,
    {STOP} 0 5px, transparent 5px 10px); }}
.dcard .hd{{ display:flex; align-items:center; gap:9px; flex-wrap:wrap; }}
.dcard .id{{ font-family:var(--mono); font-weight:600; font-size:.94rem;
  color:inherit; }}
.dcard .meta{{ font-size:.83rem; color:var(--t-mute); }}
.dcard ul{{ margin:6px 0 0 2px; padding-left:16px; }}
.dcard li{{ font-size:.83rem; color:var(--t-soft); margin:1px 0; }}

/* stamps — hue + a tint of the same hue, so the pair travels between themes */
.stamp{{ display:inline-block; font-family:var(--display); font-weight:700;
  font-size:.72rem; text-transform:uppercase; letter-spacing:.11em;
  padding:2px 9px; border-radius:4px; line-height:1.55;
  border:1px solid color-mix(in srgb, currentColor 55%, transparent);
  background:color-mix(in srgb, currentColor 13%, transparent); }}
.stamp.ok{{ color:var(--ok-t); }}
.stamp.stop{{ color:var(--stop-t); }}
.stamp.warn{{ color:var(--warn-t); }}
.stamp.info{{ color:var(--info-t); }}
.stamp.mute{{ color:var(--t-mute); }}
.muted{{ color:var(--t-mute); }}

/* ---------- empty state ---------- */
.empty{{ background:var(--fill-1); border:1px dashed var(--line-2);
  border-radius:var(--r); padding:26px 22px; text-align:center; }}
.empty .i{{ font-size:1.5rem; }}
.empty .t{{ font-family:var(--display); font-weight:700; font-size:1.02rem;
  text-transform:uppercase; letter-spacing:.05em; margin-top:6px; color:inherit; }}
.empty .b{{ font-size:.87rem; color:var(--t-mute); margin-top:4px; }}

/* ---------- streamlit widgets ---------- */
[data-testid="stTabs"] [role="tablist"]{{ gap:2px; border-bottom:1px solid var(--line); }}
[data-testid="stTabs"] [role="tab"]{{ padding:8px 15px; border-radius:8px 8px 0 0; }}
[data-testid="stTabs"] [role="tab"] p{{ font-family:var(--display); font-weight:600;
  font-size:.95rem; letter-spacing:.035em; text-transform:uppercase;
  color:var(--t-mute); }}
[data-testid="stTabs"] [role="tab"]:hover p{{ color:var(--t-soft); }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{ background:var(--fill-2);
  box-shadow:inset 0 -3px 0 var(--accent); }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] p{{ color:inherit; }}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{{ display:none; }}

[data-testid="stMetric"]{{ background:var(--fill-1); border:1px solid var(--line);
  border-radius:9px; padding:11px 13px 9px; }}
[data-testid="stMetricLabel"] p{{ font-family:var(--display); font-weight:600;
  font-size:.74rem; text-transform:uppercase; letter-spacing:.09em;
  color:var(--t-mute); }}
[data-testid="stMetricValue"]{{ font-family:var(--mono); font-weight:600;
  font-size:1.32rem; color:inherit; }}

.stButton>button, .stDownloadButton>button, .stLinkButton>a{{
  font-family:var(--display); font-weight:600; font-size:.95rem;
  letter-spacing:.035em; text-transform:uppercase; border-radius:8px;
  border:1px solid var(--line-2); background:var(--fill-1); color:inherit;
  padding:.42rem 1rem; transition:border-color .12s, background .12s; }}
.stButton>button:hover, .stDownloadButton>button:hover, .stLinkButton>a:hover{{
  border-color:currentColor; background:var(--fill-2); color:inherit; }}
.stButton>button[kind="primary"], .stButton>button[kind="primaryFormSubmit"]{{
  background:var(--accent); border-color:var(--accent); color:var(--accent-ink); }}
.stButton>button[kind="primary"]:hover{{ background:var(--accent);
  border-color:var(--accent-ink); color:var(--accent-ink); filter:brightness(1.06); }}
.stButton>button[kind="primary"] p{{ color:var(--accent-ink); }}
.stButton>button:focus-visible, .stDownloadButton>button:focus-visible,
.stLinkButton>a:focus-visible{{ outline:2px solid var(--info); outline-offset:2px; }}
.stButton>button:disabled, .stDownloadButton>button:disabled{{ opacity:.45; }}

[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3{{ font-size:.95rem;
  text-transform:uppercase; letter-spacing:.07em; }}
[data-testid="stSidebar"] [data-testid="stMetric"]{{ padding:7px 9px; }}
[data-testid="stSidebar"] [data-testid="stMetricValue"]{{ font-size:1rem; }}

.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea{{
  font-family:var(--body); border-radius:7px; border:1px solid var(--line-2);
  color:inherit; }}
.stTextInput input:focus, .stTextArea textarea:focus{{ border-color:var(--info);
  box-shadow:0 0 0 3px color-mix(in srgb, {INFO} 22%, transparent); }}
/* The mail body holds fixed-width tables — a proportional font makes the columns
   look broken even when the text itself is perfectly aligned. Streamlit puts a
   `st-key-<key>` class on any keyed widget, so this hits only those two. */
.st-key-mail_body textarea, .st-key-sh_body textarea{{
  font-family:var(--mono); font-size:.79rem; line-height:1.45;
  white-space:pre; overflow-x:auto; }}
[data-baseweb="select"]>div{{ border-radius:7px; border-color:var(--line-2); }}
[data-baseweb="tag"] span:not([data-testid]):not([translate="no"]){{
  font-family:var(--mono); font-size:.76rem; }}

[data-testid="stFileUploaderDropzone"]{{ background:var(--fill-1);
  border:1.5px dashed var(--line-2); border-radius:9px; }}
[data-testid="stFileUploaderDropzone"]:hover{{ border-color:var(--accent); }}

[data-testid="stExpander"]{{ border:1px solid var(--line); border-radius:9px;
  background:var(--fill-1); }}
[data-testid="stExpander"] summary{{ font-family:var(--display); font-weight:600;
  letter-spacing:.03em; color:inherit; }}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary [data-testid="stMarkdownContainer"]{{
  font-family:var(--display); font-weight:600; color:inherit; }}
[data-testid="stExpander"] summary svg{{ fill:currentColor; }}

[data-testid="stDataFrame"], [data-testid="stDataFrameResizable"]{{
  border:1px solid var(--line); border-radius:9px; }}
[data-testid="stDataFrame"] [role="columnheader"]{{ font-family:var(--display);
  text-transform:uppercase; letter-spacing:.04em; }}

div[data-testid="stAlert"]{{ border-radius:9px; }}
div[data-testid="stAlert"] p{{ color:inherit; }}
hr{{ border-color:var(--line); margin:1.1rem 0; }}
[data-testid="stProgress"] div[role="progressbar"]>div{{ background:var(--accent); }}
.footnote{{ font-family:var(--mono); font-size:.72rem; color:var(--t-mute);
  border-top:1px solid var(--line); padding-top:11px; margin-top:24px; }}

/* Material icons are LIGATURES — the element literally contains the text
   "arrow_right". Any font-family / letter-spacing / text-transform change stops
   the ligature forming and the raw word prints over the label. Keep them out. */
[data-testid="stIconMaterial"], [data-testid^="stExpanderIcon"],
span[translate="no"], .material-icons, .material-symbols-rounded,
.material-symbols-outlined{{
  font-family:'Material Symbols Rounded','Material Symbols Outlined',
    'Material Icons' !important;
  font-weight:400 !important; font-style:normal !important;
  letter-spacing:normal !important; text-transform:none !important;
  font-variant:normal !important; white-space:nowrap !important;
  font-feature-settings:'liga' !important;
  -webkit-font-feature-settings:'liga' !important; }}
</style>
"""


def inject() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)


def theme_type() -> str:
    """Active theme — for the few places Python needs to know (charts, images)."""
    try:
        t = getattr(st.context, "theme", None)
        if t is not None and getattr(t, "type", None) in ("light", "dark"):
            return t.type
    except Exception:
        pass
    try:
        base = st.get_option("theme.base")
        if base in ("light", "dark"):
            return base
    except Exception:
        pass
    return "dark"


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


def muted(text: str) -> str:
    return f"<span class='muted'>{_esc(text)}</span>"


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
