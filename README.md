# VIP Pick Generation System

Python + Streamlit + Google Sheets + GitHub system that turns a **Requirement**
file into two warehouse pick deliverables for the VIP Industries / EFL 3PL flow:

- **INDIA_SO_Pick.xlsx** — WMS upload format (`OutBound MASTER` + `OutBound Detail`)
- **VIP_PICK.xlsx** — pick summary sheet with CBM / time estimate block

The whole pipeline is driven by one validated engine (`pick_engine.py`) and
exposed through a Streamlit UI (`app.py`).

**Flow:** upload **Requirement** + **Inventory_Report** Excel files →
**SKU_MASTER** is read from a Google Sheet (managed/uploaded in the app's
🗂️ SKU_MASTER tab) → before generating, any Requirement material missing from
SKU_MASTER is flagged so it can be added first → generate the two deliverables.

---

## සිංහල — කෙටි විස්තරය

මේ system එක `Requirement` file එක අරගෙන:

1. `SKU_MASTER` එකෙන් එක් එක් Material එකේ **Category (LOOSE/SET)**, **HJ**, **SAP** lookup කරනවා.
2. `Inventory_Report` එකෙන් **box size** (Actual Qty mode), **CBM**, සහ **available qty** ගන්නවා.
3. carton rounding logic එකෙන් **HJ Box Qty / HJ Pcs Qty** calculate කරනවා.
4. **INDIA SO Pick** සහ **VIP PICK** Excel දෙක generate කරනවා.

Input එක Excel upload කරලාවත්, Google Sheet එකකින්වත් දෙන්න පුළුවන්.
Output Excel දෙක download කරගන්න පුළුවන්, ඕනෙ නම් ආපහු Google Sheet එකකට write කරන්නත් පුළුවන්.

---

## Pick logic (authoritative)

For every requirement line the target pick quantity in **pieces** is:

```
target_pcs = Req Qty × HJ ÷ SAP
```

| Case | Rule | Example |
|---|---|---|
| Worked example | Req=2, SAP=1, HJ=2 → pick **4** | 2 × 2 ÷ 1 = 4 |
| **LOOSE** | SAP = HJ → multiplier 1 → pick = **Req Qty** | Req 6 → 6 pcs |
| **SET** | HJ/SAP is the set multiplier | SAP=5, HJ=15 → ×3 |

**Cartons are never split.** A carton holds `Actual Qty` pieces (from the
Inventory_Report). We pick only **whole cartons**, limited by available stock:

```
pick_cartons = min(target_pcs // carton, available // carton)
picked_pcs   = pick_cartons × carton
variance     = target_pcs − picked_pcs    (the part we cannot pick)
```

`HJ Box Qty = pick_cartons`, `HJ Pcs Qty = picked_pcs`. The picked whole-carton
portion goes to the main VIP PICK / INDIA SO outputs.

### Cannot-Pick report (separate)
Any line with `variance > 0` is **also** listed in the **Cannot Pick** report
(a sheet in VIP PICK, a tab in the app, and a Google Sheet worksheet), showing
**Picked Pcs** and **Variance** so you can see how much was picked and how much
couldn't be:

- **Carton split — partial pick** — e.g. Req 23, carton 2 → pick **22** (11
  cartons), **variance 1**.
- **Insufficient stock** — `target_pcs` exceeds available inventory; pick the
  whole cartons that fit, variance = the shortfall.
- **Missing SKU / inventory data** — material not in SKU_MASTER and/or inventory.

### LOAD ID uniqueness
LOAD IDs (= Delivery No) must be globally unique. A **LOAD_ID Registry**
worksheet in the Google Sheet records every LOAD ID ever used; if a new run
produces one that already exists, a suffix is appended — `-A`, `-B`, … `-Z`,
`-AA` — so it never duplicates.

### SKU_MASTER in Google Sheet (CRUD)
SKU_MASTER lives in a Google Sheet worksheet. The app's **🗂️ SKU_MASTER** tab
loads it into an editable grid where rows can be **added, edited, and deleted**.
Two save modes:
- **➕ Merge** — upsert by Material code: existing codes are updated, new codes
  added, and other existing rows are **kept (never deleted)**. Use this when
  uploading a partial/corrected SKU_MASTER.
- **💾 Replace** — overwrite the whole worksheet with the current grid.

Duplicate material codes are reconciled as Category = first, SAP = max, HJ = max.

### Monthly history
Every generated pick is appended to a monthly worksheet named
`History YYYY-MM` (date-stamped, with the deduped LOAD ID per line). The app's
**📅 History** tab lists the months and displays any month's data. A **Run Log**
worksheet records one summary row per run.

### CBM note
The reference VIP_PICK total CBM (125.64) is **not reproducible** from the
supplied inventory snapshot (the engine computes Σ(boxes × Cbm) from the given
`Cbm` column). CBM is therefore computed transparently and **Per Minute CBM**
(default `1/3 ≈ 0.333333`) is an **editable** sidebar assumption.

## Project layout

```
vip-pick-system/
├── app.py                      # Streamlit UI
├── pick_engine.py              # validated calculation + Excel writers
├── gsheet.py                   # Google Sheets read/write helper
├── test_engine.py              # validation harness (166/169 guard)
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/
│   ├── config.toml             # theme
│   └── secrets.toml.example    # Google service-account template
└── sample_data/                # the 5 reference files for local testing
```

---

## Run locally

```bash
pip install -r requirements.txt
# or, on Microsoft Store Python:
python -m pip install -r requirements.txt

streamlit run app.py
# or:
python -m streamlit run app.py
```

Run the validation test any time you touch the engine:

```bash
python test_engine.py
```

---

## Google Sheets setup (optional input/output)

1. Google Cloud Console → create a **service account** → create a **JSON key**.
2. Enable **Google Sheets API** and **Google Drive API** for the project.
3. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`, paste
   your JSON values under `[gcp_service_account]`, and set the `[google_sheet]`
   section — `data_sheet` (the Sheet URL or key), `sku_worksheet`, `auto_save`,
   `load_id_registry`. The app reads the sheet connection from here, **not** the UI.
4. **Share** the target Google Sheet with the service-account `client_email`
   (Editor access).
5. The sidebar shows **✅ Google Sheet — secrets වලින් connected** when both the
   credentials and `[google_sheet] data_sheet` are set. SKU_MASTER worksheet
   defaults to `SKU_MASTER`.
6. Open the **🗂️ SKU_MASTER** tab and click **🆕 Initialize sheet** once — this
   auto-creates every required worksheet with headers (works on a brand-new
   empty sheet). Auto-save also ensures the worksheets exist on each run.

The app reads/creates these worksheets in that one Data Sheet:

| Worksheet | Purpose | Updated |
|---|---|---|
| `SKU_MASTER` | SKU master data (CRUD source) | on Save |
| `VIP PICK`, `CBM Summary`, `OutBound MASTER`, `OutBound Detail`, `LOAD ID QR` | latest run outputs | overwrite each run |
| `Cannot Pick` | carton-split / stock-short lines | overwrite each run |
| `LOAD_ID Registry` | every LOAD ID ever used (uniqueness) | append |
| `History YYYY-MM` | monthly pick history | append |
| `Run Log` | one summary row per run | append |

---

## Deploy to Streamlit Cloud (via GitHub)

1. Create a new GitHub repo and push this folder:
   ```bash
   git init
   git add .
   git commit -m "VIP pick generation system"
   git branch -M main
   git remote add origin https://github.com/<you>/vip-pick-system.git
   git push -u origin main
   ```
   (`.gitignore` already keeps `secrets.toml` and generated `*.xlsx` out of the repo.)
2. Go to **share.streamlit.io** → **New app** → pick the repo, branch `main`,
   main file `app.py`.
3. In **App → Settings → Secrets**, paste your `[gcp_service_account]` block
   (only needed if you use the Google Sheets source).
4. Deploy.

---

## Notes
- All WMS constants (WH_ID `LPGL`, CLIENT_CODE `INM0VIP`, ORDER_TYPE
  `Sales Orders`, etc.) are editable in the sidebar "WMS constants" expander.
- **Reset data** (sidebar 🧹): clear the current session result, or reset the
  Google Sheet data by scope — outputs, monthly history tabs, LOAD_ID Registry,
  Run Log, and optionally SKU_MASTER. Needs an explicit confirm tick; SKU_MASTER
  is left out by default so master data isn't wiped accidentally.
