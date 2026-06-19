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

For every requirement line:

```
mult        = HJ / SAP            (pieces per requirement / order unit)
                LOOSE : SAP == HJ  -> mult = 1     -> pick = Req Qty pieces
                SET   : SAP divides HJ -> mult = set size (e.g. 3)
target_pcs  = Req Qty * mult
available   = sum of Inventory_Report "Actual Qty" for that Item Number
```

We pick **whole order-units**, never exceeding available stock:

```
units_available = available // mult
units_picked    = min(Req Qty, units_available)
HJ Pcs Qty      = units_picked * mult       (pieces -> INDIA SO Detail QTY)
HJ Box Qty      = units_picked
Pcs/Box         = mult
Short Variance  = target_pcs - HJ Pcs Qty   (only if stock is short)
```

| Example | Req | Cat | SAP | HJ | mult | Available | Pick |
|---|---|---|---|---|---|---|---|
| LOOSE | 4 | LOOSE | 1 | 1 | 1 | 4 | **4** |
| SET   | 2 | SET | 1 | 3 | 3 | 186 | **6** (2×3) |
| SET   | 2 | SET | 1 | 2 | 2 | 4 | **4** (2×2) |
| LOOSE | 2 | LOOSE | 1 | 1 | 1 | 28 | **2** |

The match is done by **Requirement Material -> Inventory Item Number**, summing
that item's Actual Qty across all pallets. **Nothing is ever picked beyond the
sum of Actual Qty** (whole order-units only).

### Cannot-Pick report (separate)
A line is reported in **Cannot Pick** (a sheet in VIP PICK, a tab in the app,
and a Google Sheet worksheet) when:

- **Insufficient stock** — `target_pcs` exceeds available; the whole units that
  fit are still picked, and the short part is shown as **Short Variance** with
  **Picked Pcs**.
- **Missing SKU / inventory data** — material not in SKU_MASTER and/or inventory.
  (The app also pre-checks the Requirement and asks you to add missing materials
  to SKU_MASTER before generating.)

### INDIA SO attributes
`GEN_ATTRIBUTE_VALUE1..11` in the OutBound Detail are filled from the
Inventory_Report, matched per item:
Color, Size, Style, Supplier, Plant, Client So, Client So Line, Po Cust Dec,
Customer Ref Number, Item Id, Invoice Number1.

### HighJump-compatible empty cells
Blank cells in the INDIA SO file are written as **truly empty** cells (not empty
strings), so HighJump validations such as `cartonLabel` (000-999) accept them.

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
