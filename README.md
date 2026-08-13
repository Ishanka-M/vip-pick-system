# Donaldson OutBound Pick Generator

Invoice / Delivery Challan **PDF** + **Inventory Report** → pallet-level pick →
Google Sheet ledger + **`OutBound MASTER` / `OutBound Detail`** Excel (Körber One upload).

```
app.py          Streamlit UI (7 tabs)
doc_parser.py   Donaldson Invoice + Delivery Challan PDF parser
pick_engine.py  matching · allocation · qty verify · WMS output · Excel · search
pick_pdf.py     Pick sheet + Shortage PDF (QR) · charts · email (.eml / mailto)
sku_master.py   SKU master — dedupe upsert · base-ID search
gsheet.py       Google Sheet DB + API manager (retry · cache · lock · load delete)
```

---

## 1. Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

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

| Email | Chart |
|---|---|
| Pick email | **Picked qty by item** — item එකකට qty + pallet ගාණ, ලොකුම එක red |
| Shortage email | **Required vs Available vs Short** — grouped bars |

Chart එක `.eml` එකේ විතරයි (mailto: වලට image යවන්න බෑ).
Pick email එකේ chart එක result screen එකේ preview කරන්නත් පුළුවන්.
