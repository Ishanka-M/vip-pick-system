# Donaldson OutBound Pick Generator

Invoice / Delivery Challan **PDF** + **Inventory Report** → pallet-level pick →
Google Sheet ledger + **`OutBound MASTER` / `OutBound Detail`** Excel (Körber One upload).

```
app.py          Streamlit UI (7 tabs)
ui.py           design system — tokens, CSS, topbar / step rail / cards / stamps
.streamlit/     theme config (widget internals follow this, not CSS)
doc_parser.py   Donaldson Invoice + Delivery Challan PDF parser
pick_engine.py  matching · allocation · qty verify · WMS output · Excel · search
pick_pdf.py     Pick sheet + Shortage PDF (QR) · charts · email (.eml / mailto)
sku_master.py   SKU master — dedupe upsert · base-ID search
invoice_register.py  every uploaded invoice — summary + details reports
gsheet.py       Google Sheet DB + API manager (retry · cache · lock · load delete)
```

---

## 1. Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

> **හැම file එකක්ම එකම release එකෙන් ගන්න.** `app.py` විතරක් update කරොත්
> engine එකට තේරෙන්නේ නැති field එකක් යවනවා → Streamlit Cloud එකේ
> `TypeError` එකක්, ඒත් **message එක redact වෙනවා** නිසා මොකක්ද කියලා පේන්නේ නෑ.
> ඒ නිසා හැම module එකකම `API` number එකක් තියෙනවා, mismatch එකක් තිබ්බොත්
> app එක **run වෙන්නේ නෑ** — කොයි file එකද පරණ කියලා නමින්ම කියනවා.

| File | API |
|---|---|
| `gsheet.py` | 7 |
| `invoice_register.py` | 8 |
| `pick_engine.py` · `pick_pdf.py` | 4 |
| `doc_parser.py` | 4 |
| `ui.py` | 3 |
| `sku_master.py` | 2 |

**`.streamlit/config.toml` එක repo එකට push කරන්න අමතක කරන්න එපා.** Dropdown ·
date picker · `st.dataframe` (canvas එකක් — CSS එකට ළඟා වෙන්න බෑ) වගේ widget වල
ඇතුලත colour එන්නේ ඒකෙන් තමයි. ඒක නැත්නම් app එක Streamlit default theme එකට
වැටෙනවා.

`.streamlit/secrets.toml`:

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "xxx@yyy.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"

[google_sheet]
data_sheet  = "https://docs.google.com/spreadsheets/d/<SHEET_KEY>/edit"
auto_save   = true
wh_id       = "INMM01"
client_code = "INM0DONA"
```

Sheet එක service-account email එකට **Editor** විදිහට share කරන්න.
පළවෙනි වතාවට sidebar එකේ **🆕 Initialize worksheets** click කරන්න.

Google Sheet නැතුවත් app එක වැඩ කරනවා — Excel download විතරක් (ledger / duplicate check නෑ).

---

## 2. Flow

1. **Upload** — Invoice / DC PDF (කීයක් හරි) + Inventory Report Excel.
2. **Parse + completeness check** — S.No sequence · `Total Quantity` · `Grand Total`.
   Fail වුණොත් ⛔ — **pick කරන්නේ නෑ**.
   Parse එක වැරදුනොත් *Lines review / edit* table එකෙන් හදන්න පුළුවන්.
3. **Plant confirm** — inventory එකේ තියෙන plant ටික පෙන්නලා confirm ගන්නවා.
   Confirm කරන කල් generate button එක enable වෙන්නේ නෑ.
4. **Pick** — pallet level, FIFO (හෝ තෝරගත්ත strategy එකට).
5. **Qty verify** — Invoice / DC qty එකට **හරියටම** ගැලපෙනවද කියලා check (line · doc total ·
   WMS file total). Fail වුණොත් ඒ document එක reject.
6. **Save + download** — Google Sheet append · **LOAD_ID එකෙන් Excel + PDF** · email.

---

## 3. Rules

| Rule | Implementation |
|---|---|
| **Pick Id gate** | `Pick Id = 0` pallet විතරයි pick කරන්නේ. 0 නොවන එකක් = WMS එකේ pick task එකකට allocate වෙලා (Status එක තාම `Available`) → අයින් කරනවා |
| Item match — base ID විතරයි | `P162400-000-140` → `P162400`. Inventory `P162400-016-140` ගැලපෙනවා |
| Exact item priority | Base ID එක ගැලපුනත් full item number එකට මුල් තැන (option) |
| Plant | Confirm කරපු plant(s) එකෙන් විතරයි |
| Document complete නැත්නම් pick නෑ | Qty total / amount total / S.No sequence mismatch → `INCOMPLETE DOCUMENT` |
| Stock මදි නම් pick නෑ | එක line එකක් මදි වුණත් **මුළු document එකම** reject (`STOCK SHORT`). Partial pick නෑ |
| Duplicate | Batch එකේ + `DOC_REGISTRY` sheet එකේ check → එකක් විතරයි |
| Pallet level save | `QTY_BEFORE → QTY_PICKED → QTY_BALANCE` හැම pallet එකකටම `PALLET_LEDGER` එකට |
| **Pallet cap** | Pallet එකකට **වඩා වැඩියෙන් pick වෙන්නේ නෑ** — allocate කරද්දී cap එකක් + commit කරන්න කලින් `PALLET OVER-PICK` guard එකක් |
| **Next run balance** | පහළ **§ 6 · Stock basis rule** බලන්න |

**Pick strategies**

* `FIFO` — පරණ stock මුලින් (default)
* `SINGLE_PALLET_FIRST` — හැකි නම් එක pallet එකෙන් (pallet touch අඩුයි)
* `LEAST_PALLETS` — ලොකු pallet මුලින්

---

## 4. Output — `OutBound_Upload_*.xlsx`

හැම cell එකක්ම **text** (`number_format = "@"`), හිස් cell **හිස්මයි** (WMS එකට ඕන විදිහට).

### Sheet `OutBound MASTER` (101 columns)

| Column | Value |
|---|---|
| `DISPLAY_ORDER_NUMBER` | Invoice No / DC No |
| `STORE_ORDER_NUMBER` | Invoice No / DC No |
| `CUSTOMER_PO_NUMBER` | Invoice No / DC No |
| `LOAD_ID` | Invoice No / DC No |
| `PROCESSING_CODE` | `NEW` |
| `WH_ID` / `CLIENT_CODE` / `ORDER_TYPE` | sidebar values (`INMM01` / `INM0DONA` / `Sales Orders`) |
| `BACKORDER`, `PARTIAL_ORDER_FLAG`, `SAT_DELIVERY_FLAG`, `REGISTERED_MAIL_FLAG`, `RESTRICTED_MAIL_FLAG`, `COD_FLAG`, `COD_PAY_TYPE`, `COD_OPTION`, `INSURANCE_FLAG`, `SHIP_TO_RESIDENTIAL_FLAG` | `N` |
| අනිත් ඔක්කොම | හිස් |

### Sheet `OutBound Detail` (42 columns)

| Column | Source |
|---|---|
| `DISPLAY_ORDER_NUMBER` | Invoice No / DC No |
| `LINE_NUMBER` | 1, 2, 3, … |
| `DISPLAY_ITEM_NUMBER` | Inventory **Item Number** (WMS එකේ තියෙන හරියටම — trailing `.` වුණත්) |
| `LOT_NUMBER` | Inventory `Lot Number` |
| `QTY` | Invoice / DC Quantity |
| `ORDER_UOM` | Inventory `Uom` |
| `GEN_ATTRIBUTE_VALUE1` | Color |
| `GEN_ATTRIBUTE_VALUE2` | Size |
| `GEN_ATTRIBUTE_VALUE3` | Style |
| `GEN_ATTRIBUTE_VALUE4` | Supplier |
| `GEN_ATTRIBUTE_VALUE5` | **Plant** |
| `GEN_ATTRIBUTE_VALUE6` | Client So |
| `GEN_ATTRIBUTE_VALUE7` | Client So Line |
| `GEN_ATTRIBUTE_VALUE8` | Po Cust Dec |
| `GEN_ATTRIBUTE_VALUE9` | Customer Ref Number |
| `GEN_ATTRIBUTE_VALUE10` | Item Id |
| `GEN_ATTRIBUTE_VALUE11` | Invoice Number 1 |

හිස් attribute එකකට `TBC` (sidebar එකෙන් වෙනස් කරන්න පුළුවන්).

එක document line එකක් pallet කීපයකින් pick වුණොත් — item number + lot + uom + attributes
එකයි නම් **එක detail line එකකට merge** වෙනවා (qty එකතු කරලා). වෙනස් නම් වෙන වෙනම lines.

### `Pick_Report_*.xlsx`

`Doc Summary` · `Pallet Allocation` · `Pallet Balance` · `Shortage` · `Rejected Docs`

---

## 5. Google Sheet worksheets

| Worksheet | Content |
|---|---|
| `OUTBOUND_MASTER` | හැම run එකකම master rows |
| `OUTBOUND_DETAIL` | හැම run එකකම detail rows |
| `PALLET_LEDGER` | pallet · before · picked · balance · row key |
| `DOC_REGISTRY` | process කරපු Invoice / DC numbers (duplicate gate) |
| `REJECTED_LOG` | reject වුණ docs + හේතුව |
| `RUN_LOG` | run summary |

වැරදුනු run එකක් → sidebar **Reset / Undo** එකට `RUN_ID` දාලා delete කරන්න
(ledger + registry එකෙන් අයින් වෙනවා, ඒ නිසා ආපහු pick කරන්න පුළුවන්).

---

## 6. Stock basis rule — QTY_BEFORE same ද?

Upload කරන Inventory එක **fresh ද පරණද** කියලා pallet එකෙන් pallet එකට තීරණය කරනවා.
හැම pallet + item + lot එකකටම `PALLET_LEDGER` එකේ තියෙන
`QTY_BEFORE` (baseline) එකයි, Inventory එකේ **Actual Qty** එකයි compare කරනවා:

| තත්වය | තේරුම | Pick කරන්නේ |
|---|---|---|
| Ledger එකේ නෑ | අලුත් pallet එකක් | `MODE = NEW` → **Actual Qty** |
| `Actual Qty` **==** `QTY_BEFORE` | Inventory report එක තාම refresh වෙලා නෑ (pick එක WMS එකට තාම post වෙලා නෑ) | `MODE = LEDGER BALANCE` → **`QTY_BALANCE`** |
| `Actual Qty` **≠** `QTY_BEFORE` | WMS එක update වෙලා — අලුත් baseline එකක් | `MODE = NEW BASELINE` → **Actual Qty** (ledger එක ආපහු අඩු කරන්නේ නෑ) |

මේකෙන් වළක්වන දේ: එකම inventory file එකෙන් දෙපාරක් pick කරාම **double pick**,
සහ WMS එක update වුණාට පස්සේ **දෙපාරක් අඩු වීම**.

උදාහරණය — pallet `DONALDSON100826-A-108` · `P550945` · Actual 18, කලින් pick 6:

```
Run 1  ledger නෑ                 -> NEW             18 -  6 = 12   ✅
Run 2  එකම inventory file        -> LEDGER BALANCE  12 -  6 =  6   ✅ (18 නෙවෙයි)
Run 3  WMS update, Actual = 12   -> NEW BASELINE    12 -  6 =  6   ✅ (6 නෙවෙයි)
```

බලන්න පුළුවන් තැන් — result එකේ **📊 Stock Basis** tab · **📦 Pallet Balance** tab
(`MODE` column) · Pick Report එකේ `Stock Basis` sheet.
Sidebar → *Pick options* → **Pallet ledger balance logic** off කරොත් හැම වෙලේම Actual Qty.

---

## 7. 🔐 DB Reset

Sidebar → **🧹 Reset / Undo** → **DB Reset**.
Password එකෙන් unlock කරන්න ඕන (default: `Isha@1996`).

| Button | වෙන දේ |
|---|---|
| 🗑️ **Reset** | තෝරගත්ත worksheets විතරක් clear |
| 💣 **FULL DB RESET** | ledger · registry · outputs · rejected · run log ඔක්කොම clear (email book ඉතුරු වෙනවා) |
| ↩️ **Run undo** | `RUN_ID` එකක් විතරක් අයින් කරනවා — වැරදුනු run එකකට |

Header row එක විතරක් ඉතුරු වෙනවා, worksheet delete වෙන්නේ නෑ.

> Password එක code එකේ hardcode වෙලා තියෙනවා. Repo එක private නැත්නම්
> `secrets.toml` එකට දාන්න — code එකට වඩා ඒකට priority තියෙනවා:
> ```toml
> [app]
> reset_password = "…"
> ```

---

## 8. ⚠️ Document check bypass

Parse එක හරි වුණත් total එකක් වැරදියට කියවුනොත් — sidebar එකේ
**Document check bypass** on කරලා pick කරන්න පුළුවන්.
`DOC_REGISTRY` එකේ `DOC_CHECK = MANUAL OVERRIDE — …` කියලා log වෙනවා.
**Stock check එක bypass වෙන්නේ නෑ** — stock මදි නම් හැම වෙලේම reject.

---

## 9. Downloads — LOAD_ID එකෙන්

හැම document එකකටම වෙන වෙනම file:

| File | Content |
|---|---|
| `<LOAD_ID>.xlsx` | ඒ LOAD_ID එකේ `OutBound MASTER` + `OutBound Detail` විතරයි |
| `<LOAD_ID>.pdf` | **Pick sheet (LOAD_ID QR)** + upload කරපු Invoice / DC pages |
| `OutBound_<stamp>.zip` | ඔක්කොම LOAD_ID වල Excel + PDF |
| `OutBound_Upload_<stamp>.xlsx` | ඔක්කොම docs එකම file එකක (කලින් විදිහට) |
| `Pick_Report_<stamp>.xlsx` | Doc Summary · Qty Verification · Allocation · Balance · Shortage · Rejected |

> DC number එකේ `/` තියෙනවා නම් filename එකට `-` දානවා:
> `333/26-27/62` → **`333-26-27-62.pdf`**. Sheet එකේ / QR එකේ තියෙන්නේ ඇත්ත LOAD_ID එකමයි.

---

## 10. Pick Sheet PDF (LOAD_ID QR)

Landscape A4 එකක්:

* **LOAD_ID QR code** එක උඩ දකුණේ (scan කරාම LOAD_ID එක එනවා — HJ/Körber gun එකට)
* Document info — doc no · date · ref · plant · pick date · strategy · qty check
* **PICK DETAILS** — line · item number · description · lot · **pallet · location** ·
  stock · pick qty · balance · `Picked [ ]` tick box
* **QUANTITY VERIFICATION** — doc qty vs picked qty vs diff
* Picked by / Checked by / Loaded by / Remarks sign-off
* ඊට පස්සේ **upload කරපු Invoice / DC PDF එකේ pages ඔක්කොම** (checkbox එකෙන් off කරන්නත් පුළුවන්)

---

## 11. 🔎 Search

ඕනෑම data එකක් — item code · LOAD ID · pallet · location · GRN · lot · plant.
Word කීපයක් දුන්නොත් **ඔක්කොම තියෙන rows විතරයි** (AND search).

හොයන තැන් — current run (document lines · allocation · detail · master · verify · rejected) ·
inventory + balance · Google Sheet (ledger · registry · detail).
හම්බුණ ඒවා CSV එකක් විදිහට download කරන්නත් පුළුවන්.

---

## 12. 📧 Email

Sidebar → **📧 Email settings**

* **To** — save කරපු address book එකෙන් තෝරන්න, නැත්නම් type කරන්න
* **➕ Add** — address එක book එකට. Google Sheet එකේ `APP_SETTINGS` worksheet එකේ
  save වෙනවා, ඒ නිසා next time එකෙත් තියෙනවා
* Cc · From · Signature

Result එකේ **📧 Email** section එකෙන්:

| Button | වෙන දේ |
|---|---|
| ✉️ **Default mail app එකෙන් open** | `mailto:` — default mail app එක subject + body එක්ක open වෙනවා (attachment යන්නේ නෑ) |
| 📎 **Draft (.eml) download** | Double-click කරාම Outlook / Mail එකේ **draft** එකක් විදිහට open වෙනවා — **Excel + PDF attach වෙලාම** (`X-Unsent: 1`) |

Body එකේ තියෙන්නේ LOAD ID · document · plant · lines/qty · pallets · qty check +
pallet-by-pallet pick table එක (plain text + HTML දෙකම).
Subject / body edit කරන්නත් පුළුවන්.

---

## 13. 🚚 Load Manager — LOAD_ID එකෙන් download / delete

Tab **🚚 Loads**. LOAD_ID එකක් type කරන්න, නැත්නම් save කරපු list එකෙන් තෝරන්න.

* **බලන්න** — pick details (ledger) · OutBound Detail · OutBound MASTER
* **Download** — `<LOAD_ID>.xlsx` (WMS upload) · `<LOAD_ID>.pdf` (pick sheet + QR) ·
  pick details CSV
* **Delete** — sidebar password එකෙන් unlock කරලා, LOAD_ID එක ආපහු type කරලා confirm

> Delete කරාම `PALLET_LEDGER` + `DOC_REGISTRY` + master/detail/rejected වලින් අයින් වෙනවා
> → **pallet balance ආපහු එනවා** සහ **ආපහු pick කරන්නත් පුළුවන්**.
> DB එකේ original Invoice / DC PDF එක save වෙන්නේ නෑ, ඒ නිසා මෙතනින් එන PDF එකේ
> තියෙන්නේ pick sheet එක විතරයි.

---

## 14. 🏷️ SKU Master

Tab **🏷️ SKU Master**. Format එක: `Item Number`, `Item Description`
(+ ඕනෑම extra column ගාණක් — ඒවත් save වෙනවා).

| Sub-tab | වැඩේ |
|---|---|
| ⬆️ **Upload / Update** | File එක දාලා preview — 🆕 New · ♻️ Updated · = Unchanged. Save කරාම `SKU_MASTER` worksheet එකට |
| 🔎 **Search** | `07011636` දුන්නම **`07011636-000-440`** හම්බෙනවා. Description එකෙනුත් හොයනවා |
| ✏️ **Edit** | කෙලින්ම edit / අලුත් row. Save කරද්දී duplicate check ආපහු |

**Duplicate නෑ** — key එක `MATCH_KEY` (item number එක clean කරපු එක).
තියෙන item එකක් ආපහු දැම්මොත් **update** වෙනවා, අලුත් row එකක් හැදෙන්නේ නෑ.
හිස් value වලින් තියෙන data overwrite වෙන්නෙත් නෑ — වෙනස් වුණ field මොනවද කියලා
preview එකේ පේනවා.

`ITEM_NUMBER` file එකේ තියෙන හරියටම රැකෙනවා (`#1301` වගේ ඒවත්),
`BASE_ID` සහ `MATCH_KEY` auto calculate වෙනවා.

SKU master එකේ description, pick කරද්දී inventory එකේ description එක හිස් නම්
automatic fill වෙනවා (pick sheet · email · shortage PDF ඔක්කොම).

> **Base ID rule** — suffix එක කියලා ගන්නේ **3-digit කෑලි විතරක්** නම්:
> `07011636-000-440` → `07011636` ✅ · `100409-101` → `100409` ✅ ·
> `05-47174` → `05-47174` (`05` නෙවෙයි — වැරදි match වළක්වන්න)

---

## 15. 🔌 Google Sheet API manage (multi-user)

Sidebar → **🔌 API / Multi-user**

| දේ | කොහොමද |
|---|---|
| **Retry + backoff** | 429 / 5xx වලට exponential backoff + jitter, 5 attempts. Quota error නිසා app එක බිඳෙන්නේ නෑ |
| **Read cache** | TTL එකක් එක්ක (default 45s, slider එකෙන් 0–180). User කීපදෙනෙක් read කරද්දී quota ඉතුරු වෙනවා |
| **Write lock** | `_LOCKS` worksheet එකෙන් soft lock. තව කෙනෙක් save කරමින් නම් රැඳිලා, බැරි වුණොත් 🔒 message එකක් |
| **Duplicate re-check** | Save කරන lock එක ඇතුලේම registry ආපහු කියවනවා — user දෙන්නෙක් එකවර එකම Invoice එක දැම්මොත් එකක් විතරයි යනවා, අනිත් එක "DUPLICATE (other user)" |
| **Stats** | API calls · cache hits · retries · errors · last error |
| **Health check** | Latency + worksheet ටික check |

Sidebar එකේ **👤 User** එකක් දාන්න — lock owner සහ SKU `UPDATED_BY` එකට ඒක යනවා.

---

## 16. ⚠️ Shortage — PDF + Email

Stock මදි නිසා reject වුණ document එකකට:

* **`SHORT_<LOAD_ID>.pdf`** — shortage sheet (QR · short lines · required/available/short ·
  **chart** · මුළු document lines) + **upload කරපු Invoice / DC pages එකම එකට**
* **Shortage email** — ✉️ mailto හෝ 📎 `.eml` draft (shortage PDF attach + chart inline)
* Document කීපයක් නම් 🗜️ ZIP එකක්

---

## 17. 📊 Email charts

Email දෙකේම item details වලට chart එකක් **body එකට inline** යනවා (`cid:` image,
Outlook / Gmail / Apple Mail වල පේනවා):

Sidebar → **Email** → **Chart in the email** එකෙන් style එක තෝරන්න.

**Line by document line** (default) — pick එක ඇත්තටම sequence එකක් නිසා
`x` axis = document line (1, 2, 3, …):

| Email | Series |
|---|---|
| Pick email | **Pick qty** (line එකකට, pallet කීයකින්ද කියලා annotate වෙනවා) + **Balance left on pallets** — දෙකේ magnitude එක හාත්පසින් වෙනස් නිසා balance එක **වෙනම (right) axis** එකකට දාලා තියෙනවා. නැත්නම් pallet එකේ 468 එකෙන් qty 2 flat වෙනවා |
| Shortage email | **Required · Free stock · Short · On pick task** — line එකකට |
| Document කීපයක් | document එකකට line එකක් (max 6), legend එකේ doc numbers |

**Bar by item** — කලින් තිබ්බ එක: pick email එකට *Picked qty by item*,
shortage email එකට *Required vs Available vs Short* grouped bars.

Line එකක් විතරක් තියෙනවා නම් (point එකක් line chart එකකට මදි) **automatic
bar chart එකට වැටෙනවා**. Y සහ X ticks integer විතරයි — qty වලට 2.5 වගේ එකක් නෑ.

Chart එක `.eml` එකේ විතරයි (mailto: වලට image යවන්න බෑ).
Pick email එකේ chart එක result screen එකේ preview කරන්නත් පුළුවන්.

---

## 18. UI

**Direction — warehouse operations console, not a dashboard.**
Ink on light paper (data-heavy screens under warehouse lighting), hi-vis amber for
the one action that matters on each screen, **red = STOP only** (blocked / short),
කිසිම විටක decoration එකක් විදිහට නෙවෙයි.

| | |
|---|---|
| **Theme** | Default එක **dark** (`config.toml`). ඒත් UI එක theme එකකට bind වෙලා නෑ — light එකට මාරු කරත් ඔක්කොම කියවන්න පුළුවන් |
| **Type** | *Barlow Condensed* — labels, rack-signage voice · *Barlow* — reading · *IBM Plex Mono* — **හැම code එකක්ම** (LOAD ID · pallet · location · qty). Code එකක් character by character කියවන නිසා ඒවා හැම වෙලේම mono |
| **Signature** | **Hazard rule** — warehouse floor එකේ striped tape එක. Top bar එක යටත්, block වුණ document card වල වම් පැත්තෙත් විතරයි |
| **Numbers** | `01 → 05` numbering තියෙන්නේ pick එක ඇත්තටම sequence එකක් නිසා — plant confirm නොකර generate කරන්න බෑ |

**UX**

* **Step rail** — Documents → Inventory → Plant → Pick, හැම එකකම live state
  (done ✓ / now / todo) සහ value එක. දැන් කොහෙද ඉන්නේ කියලා එක බැල්මට
* **Top bar chips** — DB · plant · docs · user. Save වෙන්නේ කොහෙද කියලා හැම වෙලේම පේනවා
* **Document stamps** — `ready` · `duplicate` · `blocked` + හේතුව card එකේම
* **Empty states** — හිස් screen එකක් වෙනුවට "ඊළඟට මොකද කරන්නේ" කියන එක
* **Copy** — button එකේ තියෙන නම action එකට සමානයි (`Confirm plant` → toast `Plant confirmed`)
* Focus ring · disabled state · toast · `st.status` progress · mobile දක්වා responsive

### Theme safety — dark සහ light දෙකටම

`ui.py` එකේ colour hardcode වෙලා නෑ. හේතුව: page background එකයි
`st.dataframe` එකයි Streamlit එකට අයිතියි (dataframe එක **canvas** එකක් — CSS
එකෙන් ළඟා වෙන්න බෑ). CSS එකේ "white card / dark text" කියලා තිබ්බොත්, app එක
dark theme එකක run වුණාම අකුරු නොපෙනී යනවා.

ඒ නිසා හැම surface · border · muted text එකක්ම **`currentColor` එකෙන් mix** කරනවා:

```css
background: color-mix(in srgb, currentColor 5%, transparent);
```

* Light theme → currentColor කළුයි → ලා grey tint එකක්
* Dark theme  → currentColor සුදුයි → ලා lift එකක්
* Body text එකට colour එකක් දෙන්නේම නෑ — Streamlit එකේම text colour එක
  inherit වෙනවා, ඒක background එකට contrast වෙනවා කියලා guarantee එකක් තියෙනවා

Signal colour 4ක් විතරයි fixed (amber · green · red · blue). ඒවත් කෙලින්ම
දාන්නේ නෑ — `color-mix(in srgb, var(--ok) 56%, currentColor)` විදිහට **live text
colour එකට ටිකක් අදිනවා**. එතකොට dark එකේ light green, light එකේ dark green.

Contrast check (worst case, හැම theme combination එකකම): **4.61 : 1** —
WCAG AA small text (4.50) pass ✓

Chart — email සහ PDF වලට යන්නේ හැම වෙලේම white version එක (mail client සහ
print), screen එකේ preview එකට විතරයි dark version එක.

### Material icons — ligature වලට අත තියන්න එපා

Streamlit එකේ expander chevron එක වගේ icon ටික **ligature** — element එකේ ඇත්තටම
තියෙන්නේ `arrow_right` කියන **text** එක, `Material Symbols Rounded` font එකෙන් ඒක
arrow එකක් වෙනවා.

ඒ නිසා `summary span`, `button span` වගේ එකකට `font-family` · `letter-spacing` ·
`text-transform` දැම්මොත් ligature එක හැදෙන්නේ නෑ — label එක උඩින්
`arrow_right` කියලා raw text එකක් print වෙනවා.

`ui.py` අගට guard එකක් තියෙනවා:

```css
[data-testid="stIconMaterial"], [data-testid^="stExpanderIcon"],
span[translate="no"], .material-icons, .material-symbols-rounded {
  font-family:'Material Symbols Rounded' !important;
  letter-spacing:normal !important; text-transform:none !important;
  font-feature-settings:'liga' !important;
}
```

`text-transform` සහ `letter-spacing` **inherit** වෙන නිසා, uppercase කරපු button
එකක් ඇතුලේ icon එකක් තිබ්බත් මේකෙන් රැකෙනවා. අලුත් CSS ලියද්දී **bare `span`
selector වලට font එකක් දාන්න එපා** — element එකට direct කරන්න
(`summary p`, `[data-testid="stMarkdownContainer"]` වගේ).

### Top padding — Streamlit header එක fixed

Streamlit එකේ header එක (Share · ⭐ · Manage app) **fixed**, උස `3.75rem`.
Content ඒක යටින් scroll වෙනවා. ඒගොල්ලන්ගේ default `padding-top` එක `6rem`
වෙන්නේ ඒකයි.

`.block-container` එකේ top padding එක `3.75rem` ට වඩා අඩු කරොත් app bar එකේ
උඩ කොටස header එක යටට ගිහින් **කැපෙනවා**. දැන් `4.6rem` — default එකට වඩා
තදයි, ඒත් header එක clear වෙනවා.

```css
.block-container, [data-testid="stMainBlockContainer"]{ padding-top:4.6rem; }
```

---

## 19. Pick Id — දැනටමත් pick task එකකට ගිය stock

Inventory report එකේ **`Pick Id`** column එක:

| Pick Id | තේරුම | App එක කරන දේ |
|---|---|---|
| `0` | free | pick කරනවා |
| `0` නොවන එකක් | WMS එකේ pick task එකකට allocate වෙලා | **අයින් කරනවා** |

**මේක ඇයි වැදගත්:** allocate වුණ pallet වල `Status` එකත් තාම **`Available`**.
ඒ නිසා status filter එකෙන් ඒවා අල්ලන්නේ නෑ — gate එක නැත්නම් **එකම stock එක
දෙපාරක් pick වෙනවා**. Test file එකේ pallet 9ක් (151 units) මේ තත්වයේ තියෙනවා.

* Sidebar → *Pick options* → **Pick Id = 0 only** (default on)
* Plant table එකේ **On pick task** column එකෙන් කීයද අයින් වුණේ කියලා පේනවා
* Result එකේ **On a pick task — excluded** expander එකේ pallet · location · qty ·
  Pick Id ඔක්කොම
* Stock tab එකේ **Pick Id** filter (`FREE` / `ON PICK TASK`) — locked row වල
  `BALANCE` එක 0 කරලා තියෙනවා
* Pick Report එකේ **On Pick Task** sheet එකක්

**Shortage reason එකේ වෙනස:** stock එක ඇත්තටම තියෙනවා ඒත් locked නම්, ඒක
කියනවා —

```
On another pick task — 128 locked (Pick Id 1282815, 1284491, 1284501, 1284542)
Stock short · 128 also locked to a pick task (Pick Id …)
```

Shortage PDF එකේත් **On pick task** column එකක් තියෙනවා. එතකොට "stock නෑ" කියලා
හොයන්න යන්නේ නැතුව, කවුරු හරි ඒක pick task එකකට දාලා තියෙනවා කියලා වහාම පේනවා.

---

## 20. Email — plain-text table alignment

Email එකේ කොටස් 2ක් යනවා: **HTML** එකයි (mail client එකක් පෙන්නන්නේ ඒක)
**plain text** එකයි (fallback). HTML එකේ ඇත්ත `<table>` එකක් තියෙනවා, ඒත් plain
text එකේ column ගැලපෙන්නේ space ගාණෙන් — ඒක කැඩිලා තිබුණා.

**හේතු 2ක්:**

1. Column width `<16`, `<20` වගේ **hardcode වෙලා තිබුණා**. Doc number එකක් හෝ
   pallet id එකක් ඊට වඩා දිග වුණොත් ඊට පස්සේ තියෙන ඔක්කොම එහාට තල්ලු වෙනවා.
   දැන් width එක **data එකෙන්ම calculate වෙනවා** (`_ascii_table`), header එකට
   යටින් `----` rule එකකුත් තියෙනවා.
2. App එකේ preview text area එකේ font එක **proportional** (Barlow) වුණා —
   ඒකෙන් හරියට align වුණ table එකකුත් කැඩිලා පේනවා. දැන් mail body 2ම
   monospace (`st-key-mail_body` · `st-key-sh_body` scoped CSS).

`text/plain` part එකට `format=fixed` header එකත් දාලා තියෙනවා — ඒක respect කරන
client එකක් column straight තියාගන්නවා.

**Document කීපයක් නම්** `Ln` එක document එකකට 1, 2, 3 කියලා restart වෙන නිසා
**Document column එකක්** එකතු වෙනවා. එකක් විතරයි නම් ඒක නෑ — table එක පටුයි.

```
Document      Ln  Item Number       Pallet                 Location          Qty  Balance
------------  --  ----------------  ---------------------  ----------------  ---  -------
333262712337   1  P502639-288-140   DONAL130826-SB-13-1    IMDS01             24      300
333/26-27/62   4  P550576-016-140.  DONA081026-M-2         IMDS01             20      232
```

---

## 21. Release stock from another pick task

Document එකක් block වුණේ **stock එක තියෙනවා, ඒත් තව pick task එකකට allocate
වෙලා** කියන එක නිසා විතරක් නම් —

```
On another pick task — 12 locked (Pick Id 1284465, 1294759)
```

— result එකේ **Release stock** panel එකක් එනවා:

1. මොන documents ද කියලා පෙන්නනවා — short lines · short qty · locked qty · Pick Ids
2. *"The pallets that would be taken"* expander එකේ pallet · location · qty ·
   Pick Id ඔක්කොම
3. **"I have checked the other pick task and this stock is free to take"**
   tick කරන කල් button එක disabled
4. **Release and pick** → ඒ Pick Id ටික විතරක් open කරලා pick එක ආපහු run වෙනවා

**Release වෙන්නේ තෝරගත්ත document එකට, තෝරගත්ත Pick Id වලට විතරයි.** අනිත්
locked pallet ටික ඒ විදිහටම locked. `"*"` දුන්නොත් ඒ document එකට ඔක්කොම open.

**ඇත්තටම stock මදි නම් panel එක එන්නේම නෑ** — `AVAILABLE + ON_PICK_TASK ≥
REQUIRED` හැම line එකකටම හරි ගියොත් විතරයි. Confirmation එකකින් stock හදන්න බෑ.

**Audit** — මේක silent override එකක් නෙවෙයි:

| තැන | මොකද පේන්නේ |
|---|---|
| `DOC_REGISTRY` → `DOC_CHECK` | `RELEASED from pick task 1284465, 1294759` |
| Pick sheet PDF | header එක යටින් red box එකක් — *"RELEASED FROM ANOTHER PICK TASK … Confirm the other task before the load leaves."* |
| Pick email | `Released : taken from pick task …` (text + HTML දෙකේම) |
| On Pick Task report | `RELEASED` column එකේ `YES` |

Result එකේ **Undo the release** button එකෙන් ආපහු ගන්නත් පුළුවන් — release එක
අයින් වෙලා pick එක ආපහු run වෙනවා.

---

## 22. Invoice register

Upload කරන **හැම document එකක්ම** register වෙනවා — pick වුණත් නැතත්,
invoice එකකට **row එකයි**, ස්ථිරවම. Tab: **Register**.

### Summary

| Column | කොහෙන්ද |
|---|---|
| Tax Invoice Date | `Invoice Date` / `Delivery Challan Date` |
| Tax Invoice No. | `Invoice No.` / `Delivery Challan no.` |
| AR Invoice No. | `AR Invoice No.` (invoice වලට විතරයි) |
| Name of Customer | `Ship To / Consignee` · challan එකේ `Name of Consignee(Shipped To)` |
| Qty | මුළු document qty |
| Körber Pick | pick වුණා නම් **Yes** · නැත්නම් **No** + හේතුව remark එකට |
| MRP | Contact Person / Email එක MRP contact එකක් නම් **Yes** |

> Invoice එකේ header එක **column 5ක්** පැත්තට පැත්ත තියෙනවා, ඒ නිසා flat text
> එකේ ඒවා මිශ්‍ර වෙනවා (`438567 438549 Email: …`). Consignee එක Bill To එකෙන්
> වෙන් කරගන්න **x-coordinate band** එකෙන් කියවනවා. Email එකත් line 2කට කැඩිලා
> තියෙනවා (`rahul.sharma1@donaldso` + `n.com`) — ඒක join කරනවා.

### Details

Document line එකකට row එකයි — item · qty · picked qty · **pallet · location ·
lot** ඔක්කොම.

### Körber Pick වෙනස් වෙන විදිහ

```
මුල් upload එක, stock මදි   →  No   + "STOCK SHORT — L2 X006252 (need 999, have 6)"
පස්සේ pick වුණා            →  Yes  + remark එක automatic අයින් වෙනවා (එකම row එක)
Loads tab එකෙන් delete කළා  →  No   + "Load deleted · <time>"
```

**Yes** එකක් duplicate skip එකකින් හොරෙන් **No** වෙන්නේ නෑ — load එක delete
කරොත් විතරයි.

### MRP rule එක වෙනස් කරන්නේ කොහෙන්ද

Register tab → **MRP contacts — the rule behind the Yes / No**.
Default එක `Sharma, Rahul` / `rahul.sharma1@donaldson.com`. Row එකතු කරන්න /
වෙනස් කරන්න පුළුවන් — `APP_SETTINGS` එකේ save වෙනවා, ඊළඟ run එකේ ඉඳන් apply වෙනවා.
Name එක `Sharma, Rahul` හෝ `rahul sharma` — order එකයි punctuation එකයි ගණන් නෑ.
Document එකේ email එකක් තිබ්බොත් email එකට මුල් තැන.

### Download

**Summary** සහ **Details** — **වෙන වෙනම Excel file 2ක්**
(`Invoice_Summary_*.xlsx` · `Invoice_Details_*.xlsx`).
Screen එකේ දාන filter ටික download වෙන file වලටත් apply වෙනවා.
Freeze pane + auto-filter දාලා තියෙනවා.

Worksheets: `INVOICE_SUMMARY` · `INVOICE_DETAIL`. **FULL DB RESET එකෙන් register
එක clear වෙන්නේ නෑ** — scope එකෙන් `register` තෝරොත් විතරයි.

---

## 23. Dashboard — pending vs picked

Tab: **Dashboard**. Invoice register එකෙන් කියවනවා, ඒ නිසා **මුළු history එකම**
— අන්තිම run එක විතරක් නෙවෙයි.

| | |
|---|---|
| Progress bar | picked invoices / total, % එකත් එක්ක |
| Pending invoices · Pending qty | ඉතුරු ප්‍රමාණය — **දෙකම** ගණන් කරනවා. Unit 1000ක pending invoice එකයි, unit 2 බැගින් pending invoice 10යි එකම ප්‍රශ්නයක් නෙවෙයි |
| Oldest pending | ඉතුරු වෙලා තියෙන පරණම එකේ invoice date එකේ ඉඳන් දවස් ගාණ |
| Why they are waiting | pending qty එක **හේතුව අනුව** — stock short · on another pick task · duplicate · document incomplete · not picked yet |
| Who is waiting | pending qty එක customer අනුව, ලොකුම එක උඩින් |
| Oldest first | chase list එක — days · invoice · customer · qty · still to pick · reason |

Remark එකෙන් හේතුව automatic classify වෙනවා (`_REASONS` pattern list එකෙන්).

**Filter** — invoice date range + document type. හැම number එකකටම, download 2ටම
apply වෙනවා.

**Download — හතරම මෙතන**

| File | මොනවද |
|---|---|
| **Summary report (Excel)** | invoice එකකට row එකයි — මුළු register එකම |
| **Details report (Excel)** | document line එකකට row එකයි |
| **Pending report (Excel)** | sheet 5ක්: `Status` · `Pending` · `By reason` · `By customer` · `Picked` |
| **Pending list (CSV)** | chase list එක විතරයි |

Date range + document type filter එක **හතරටම** apply වෙනවා.

> Summary සහ Details report දෙක **තැන් දෙකකින්ම** ගන්න පුළුවන් — data එකම
> එකයි, filter දෙක වෙනස්:
> * **Dashboard** → invoice date range · document type
> * **Register** → search · Körber Pick · MRP

Pending invoice එකක් **තනියම clear වෙනවා** — Pick tab එකෙන් ආපහු upload කරලා
pick වුණාම register එක `Yes` කරනවා, dashboard එකෙන් අයින් වෙනවා.

---

## 24. Duplicates · API quota · multi-user · speed

### Duplicate කිසිම තැනකට යන්නේ නෑ

Invoice එකක් කලින් pick වෙලා තියෙද්දී ආපහු upload කරොත් — ඒක **pending වැඩක්
නෙවෙයි**, ඉවර වුණ වැඩක්. ඒ නිසා **register එකටවත් dashboard එකටවත් යන්නේ නෑ**:

* `R.build()` එකේදීම **register row එකක් හැදෙන්නේ නෑ**
* `merge_summary()` එකේත් guard එකක් — කොහොම හරි එකක් ආවත් store වෙන්නේ නෑ
* `_register_frames()` — Register tab · Dashboard · report හතරම කියවන **එකම
  තැන**. Legacy duplicate row තිබ්බත් ඔක්කොමට filter වෙනවා
* Register tab එකේ **"Clear them"** button එකෙන් sheet එකෙන්ම අයින් කරන්න
  පුළුවන් (summary + ඒවායේ detail lines එක්කම)
* Screen එකේ පේනවා: *"Already picked earlier, left out of the register and the
  dashboard: 333262712337"*

> **`DUPLICATE (batch)` වෙනස්.** ඒක කියන්නේ එකම file එක එකම upload එකේ දෙපාරක්
> දාලා කියන එක විතරයි — **පළවෙනි copy එක ඇත්ත document එක**, ඒකට row එකක්
> ඕන. කලින් version එකේ batch duplicate note එකෙන් ඇත්ත හේතුව
> (`STOCK SHORT` වගේ) overwrite වෙලා, ඒ invoice එකම register එකෙන් වැටිලා
> ගියා. දැන් ඇත්ත හේතුවට මුල් තැන, skip වෙන්නේ
> `already processed` / `other user` දෙකට විතරයි.

### API quota

| | |
|---|---|
| **Token bucket** | Google දෙන්නේ මිනිත්තුවකට request 60ක්. දැන් calls **pace** වෙනවා (default 55/min) — 429 එකක් කාලා seconds ගාණක් backoff වෙනවට වඩා ගොඩක් ලාබයි |
| **Batch read** | `values_batch_get` — worksheet 5ක් කියවන්න **request 1යි**. `read_load` 5 → 1, register 2 → 1, pick run එකේ ledger+registry 2 → 1 |
| **Worksheet list cache** | 5 විනාඩියක් — batch read එකක් සම්පූර්ණයෙන්ම request 1ක් වෙනවා |
| **Single-request write** | `clear()` + `update()` = request 2ක්, ඒ අතරේ sheet එක **හිස්**. දැන් අලුත් rows + පරණ ඒවා ආවරණය වෙන්න blank padding එකක් එකයි `update` එකකින් — atomic, request භාගයක් |
| **Gauge** | Sidebar → API & multi-user: *Quota 12/55 in the last minute* · calls · cache hits · **saved by batching** · retries · errors · paced seconds |

Rate limit එක slider එකෙන් 20–60 අතර වෙනස් කරන්නත් පුළුවන් (user ගොඩක්
එකවර නම් අඩු කරන්න).

### Multiple users

* `_LOCKS` worksheet එකේ lock එක දැන් **request 1කින්** ගන්නවා (කලින් clear +
  update = 2). Lock එකක් ගන්න request 2ක් ඕන නම්, ඒ අතරමැද දෙන්නෙක්ම lock එක
  තමන්ට කියලා හිතන්න පුළුවන් — ඒ window එක වැහුණා
* Register save එකේ merge එක **lock එක ඇතුලේ, fresh read එකකට එරෙහිව** —
  දෙන්නෙක් වෙනස් invoice දෙකක් එකවර pick කරොත් දෙකම register එකට යනවා
* Duplicate re-check එකත් lock එක ඇතුලේම — එකම invoice එක දෙන්නෙක් දැම්මොත්
  එකයි යන්නේ
* Throttle එක **thread-safe** (`threading.Lock`), ඒ නිසා user ගොඩක් එකවර
  වැඩ කරද්දීත් quota එක එකතුවෙන් manage වෙනවා

### Speed — rerun එකකට 1.51s → 0.69s (**2.2×**)

Streamlit එකේ **හැම click එකකටම මුළු script එකම ආපහු run වෙනවා**. ඒ නිසා
rerun එකකට වැය වෙන එක තමයි ඇත්ත speed එක.

| වැඩේ | කලින් | දැන් |
|---|---|---|
| `normalize_inventory` + `plant_summary` | හැම rerun එකකම 127 ms | file එකකට එක පාරයි |
| `stock_view` (2 000 rows) | හැම rerun එකකම 302 ms | file + ledger එකට එක පාරයි |
| Excel download 4ක් build කිරීම | හැම rerun එකකම ~170 ms | content signature එකකට එක පාරයි |
| Google Sheet reads | worksheet එකකට request එකක් | batch — 5 → 1 |

Download cache එකේ key එක frame එකේ **content hash** එකක් — data වෙනස් වුණාම
තනියම rebuild වෙනවා, පරණ file එකක් කවදාවත් යන්නේ නෑ.

---

## 25. Roles — Admin / Dashboard / Packing

App එක open කරාම login screen එකක්. තුනක් තෝරගන්න:

| Role | Access | Password |
|---|---|---|
| **Admin** | හැම එකක්ම — Pick, Dashboard, Register, Loads, SKU master, Search, Stock, History, Reset | ඔව් — default `Isha@1996`, `secrets.toml`→`[app] admin_password` එකෙන් වෙනස් කරන්න පුළුවන් (නැත්නම් reset password එකම) |
| **Dashboard** | **Pending vs picked** සහ **Invoice register** විතරයි — upload / pick / reset කිසිම දෙයක් නෑ | නෑ |
| **Packing** | QR scan interface එක විතරයි | නෑ |

Sidebar එකේ **Switch role / Log out** button එකෙන් ආපහු login screen එකට.
Session එකක් per browser tab — වෙන කෙනෙක් login වෙන එකෙන් මේකට බලපෑමක් නෑ.

---

## 26. Packing — QR scan (mobile camera)

Login → **Packing**. Mobile camera එකෙන්:

1. **Scan QR code** click කරාම phone එකේ native camera app එක open වෙනවා
2. Pick sheet PDF එකේ **LOAD_ID QR** එකට camera එක point කරලා photo එකක් capture කරන්න
3. App එක photo එකෙන් QR එක decode කරලා (OpenCV), `TAX_INVOICE_NO` එකට match කරනවා
4. හම්බුණොත් — `INVOICE_SUMMARY` සහ `INVOICE_DETAIL` දෙකේම **PACKING = Completed**
   ලෙස update වෙනවා, ඒ instant එකේම save
5. **Scan the next one** button එකෙන් ඊළඟ QR එකට

Live video scan එකක් නෙවෙයි — photo capture + decode, ඒක mobile browser හැම එකකම
reliable විදිහට වැඩ කරන approach එක. QR එකක් හම්බුනේ නැත්නම් / LOAD_ID එක register
එකේ නැත්නම් error message එකක්. Session එකේ scan කරපු ලැයිස්තුවක් පහළින්.

---

## 27. Picking / Packing / Dispatch — status columns

`INVOICE_SUMMARY` සහ `INVOICE_DETAIL` දෙකේම තුන් column — `KORBER_PICK`
(මේ app එකෙන්ම pallet allocate කළාද) එකෙන් **independent**, physical warehouse
execution එක track කරන්නේ:

| Column | `Completed` වෙන්නේ කොහොමද |
|---|---|
| **PICKING** | `Pick_Live_status` upload — `Open Pick = 0` · හෝ Invoice sales report match (පහළින්) |
| **PACKING** | Packing role QR scan (§ 26) |
| **DISPATCH** | `Pick_Live_status` upload — `Shipped Pick ≠ 0` හෝ `Total Pick = Shipped Pick` |

**Pending → Completed විතරයි** — කවදාවත් ආපහු Pending වෙන්නේ නෑ (invoice එක ආපහු
upload කළත්, stale report එකක් upload කළත්). Register tab → **Update status**
sub-tab එකෙන් file 2ක් upload කරන්න පුළුවන් (Load Id / Tax Invoice No. එකෙන්
match වෙනවා):

* **Pick_Live_status** (WMS export) — `Load Id`, `Open Pick`, `Total Pick`,
  `Shipped Pick` columns ඕන
* **Invoice sales report** (ERP export) — `Tax Invoice No.`, `Item Code` /
  `Customer Item`, `QTY` columns ඕන. Check වෙන්නේ **ඇත්තටම WMS එකට ගිය data
  එකෙන්ම** — invoice PDF එකෙන් කියෙව්ව qty එකෙන් නෙවෙයි:
  * **`OUTBOUND_MASTER`** → `LOAD_ID` == Tax Invoice No. — මේ invoice එක
    ඇත්තටම pick වෙලා WMS එකට push වෙලාද කියලා. නැත්නම් *"Not in
    OUTBOUND_MASTER — not picked yet"*
  * **`OUTBOUND_DETAIL`** → `DISPLAY_ITEM_NUMBER` (base-ID matched) + `QTY` —
    ඇත්තටම යැව්ව qty එකට upload කරන report එකේ Customer Item / Item Code +
    QTY එක **හරියටම** ගැලපුනොත් ඒ line එකට Picking confirm, invoice එකේ line
    ඔක්කොම confirm වුණාම invoice එකටම Picking = Completed
  * තාම pick එකක් save වෙලා නැත්නම් (`OUTBOUND_MASTER` හිස්) — invoice එකේම
    qty එකට check කරන පරණ විදිහට වැටෙනවා (degraded mode, screen එකේ
    warning එකක් එනවා)

දෙකෙන්ම apply කළාට පස්සේ **row-by-row reconciliation report** එක screen එකේම
(expander එකක) පේනවා, සහ **Excel download button** එකකින් ගන්නත් පුළුවන්
(`Sales_Reconciliation_*.xlsx` · `Pick_Live_Status_Report_*.xlsx`) — apply
කරන හැම එකකටම වෙනම.

Admin සහ Dashboard role දෙකටම මේ upload දෙක Register tab එකේ තියෙනවා.

---

## 28. Invoice Unit Price / Total

Invoice PDF එකේ line table එකේ **Unit Price** සහ **Total** column දෙක, සහ
document එකේ **Total Amount (Incl. Tax)** — දැන් parse වෙලා register එකට save
වෙනවා:

| Register | Column |
|---|---|
| `INVOICE_DETAIL` | `UNIT_PRICE`, `LINE_TOTAL` (line එකේ Total — tax සමග) |
| `INVOICE_SUMMARY` | `TOTAL_INCL_TAX` (document එකේ Total Amount (Incl. Tax)) |

Delivery Challan එකකට Unit Price column එකක් නෑ — `LINE_TOTAL` එකට Total
Amount (incl. tax) column එකේ අගය, `TOTAL_INCL_TAX` එකට Grand Total එක.
