# VIP Pick Generation System

Python + Streamlit + Google Sheets + GitHub system that turns a **Requirement**
file into two warehouse pick deliverables for the VIP Industries / EFL 3PL flow:

- **INDIA_SO_Pick.xlsx** — WMS upload format (`OutBound MASTER` + `OutBound Detail`)
- **VIP_PICK.xlsx** — pick summary sheet with CBM / time estimate block

The whole pipeline is driven by one validated engine (`pick_engine.py`) and
exposed through a Streamlit UI (`app.py`). Input can come from uploaded Excel
files **or** directly from Google Sheets.

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

## Validated pick logic

Reverse-engineered from the reference files and validated at **166 / 169
(98.2%) exact match** on `HJ Pcs Qty`:

| Field | Rule |
|---|---|
| `boxsize` | mode of inventory **Actual Qty** for that material (physical carton pack size) |
| `cbm_per_box` | inventory **Cbm** value for that material |
| `avail` | sum of inventory Actual Qty for that material |
| `divisor` | `boxsize` if Category == **LOOSE**, else **SAP** (from SKU_MASTER) for **SET** items |
| `HJ Box Qty` | `floor(Req Qty / divisor)` |
| `HJ Pcs Qty` | `HJ Box Qty × boxsize` |
| `Pcs/Box` | `boxsize` |
| `REMARKS` | `"Shortage"` when `HJ Pcs Qty < Req Qty` (carton-rounding shortfall) |

**SKU_MASTER dedup** (546 duplicate material codes): Category = first, SAP = max, HJ = max.

**INDIA SO `OutBound Detail` QTY == `HJ Pcs Qty`** (confirmed 169/169).

### The 3 mismatches
The 3 non-matching lines are all **Qty == 1 partial-carton manual loose
picks** — cases a human picked a single piece that no deterministic
carton-rounding rule reproduces. They are left to the validated rule; review
the flagged-rows panel in the app if a single-piece line matters for a run.

### CBM note
The reference VIP_PICK total CBM (125.64) is **not reproducible** from the
supplied inventory snapshot (the engine computes 162.31 = Σ(boxes × cbm) from
the given `Cbm` column — likely a different inventory snapshot/CBM table was
used originally). The engine therefore computes CBM **transparently** from the
inventory `Cbm` column and exposes **Per Minute CBM** (default `1/3 ≈ 0.333333`)
as an **editable assumption** in the sidebar so you can match your own pick-rate.

---

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
3. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and paste
   your JSON values under `[gcp_service_account]`.
4. **Share** the target Google Sheet with the service-account `client_email`
   (Editor access).
5. In the app sidebar pick **Google Sheets** as the source and paste the sheet
   URL or key.

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
- Rounding mode is configurable in the sidebar (default **floor**, which gave the
  best match). `round` / `ceil` are available if a future client wants over-pick.
- All WMS constants (WH_ID `LPGL`, CLIENT_CODE `INM0VIP`, ORDER_TYPE
  `Sales Orders`, etc.) are editable in the sidebar "WMS constants" expander.
