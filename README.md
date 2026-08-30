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
| `gsheet.py` | 10 |
| `invoice_register.py` | 14 |
| `pick_engine.py` · `pick_pdf.py` | 4 |
| `doc_parser.py` | 6 |
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

> **Base ID rule** — separator එකෙන් (hyphen **හෝ space**) කැඩුවම, පස්සේ තියෙන
> කෑලි **ඔක්කොම digit 3ක ගුණාකාර** නම් ඒවා suffix, මුල් කෑල්ල base ID එක:
> `07011636-000-440` → `07011636` ✅ · `100409-101` → `100409` ✅
>
> ERP export එකේ hyphen එක වෙනුවට **space** එකක් එනවා, 3-digit කෑලි එකට
> ඇලිලත් එනවා — ඒවත් suffix, base එක එකමයි:
> `X770132 003710` → `X770132` ✅ · `P951413 000710` → `P951413` ✅ ·
> `P775704   710` → `P775704` ✅ · `X770132-003710` → `X770132` ✅
> (non-breaking space `\xa0` වුණත් අල්ලනවා)
>
> 3ේ ගුණාකාරයක් නෙවෙයි නම් / digit නෙවෙයි නම් **මුළු code එකම** — වැරදි match
> වළක්වන්න: `05-47174` → `05-47174` (digit 5යි) ·
> `1C072323-INL` → `1C072323-INL` (`INL` digit නෙවෙයි)

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

---

## Data mapping — what was checked, and what was wrong

Every frame the app writes was put against the header its worksheet declares.
No column is produced that the header does not know, so nothing is silently
dropped on write. `RUN_ID` and `PROCESSED_AT` are the only header columns not in
the source frames — `save_run` stamps those at write time.

Headers themselves are safe two ways: `init_sheet` writes one for every
registered worksheet, and `sheet_header()` widens an older, shorter header by
appending the missing columns **at the end**, so a value never lands under the
wrong name when a release adds a column.

Two real bugs came out of this pass, both in the reconciliation shipped just
before it:

**1 · The corrections moved no stock at all.**
`ledger_state()` groups the ledger on `ROW_KEY`, falling back to
`PALLET|ITEM|LOT`. The correction rows carried none of those — empty `ROW_KEY`,
empty lot — so they grouped on their own and the pallet's balance never changed.
A silent no-op: the sheet filled up with corrections and the stock stayed
locked. The reconciliation now carries `ROW_KEY`, `LOT_NUMBER`, `LOCATION_ID`,
`PLANT`, `UOM` and `QTY_BEFORE` through from the ledger row it is correcting, and
`reconcile(keys=…)` borrows them from the stock basis for a pallet the system
never chose. `QTY_BEFORE` is carried rather than zeroed, because `ledger_state`
takes the **max** of it as the pallet's baseline and a zero would drag it down.

**2 · Applying the same correction twice moved the stock twice.**
Press *Apply to the ledger*, then *Mark order complete* with the same cached
frame, and a pallet went `picked 20 → 15 → 10` — quietly gaining five units that
were never on the rack. Every correction now carries a deterministic id,
`RECON-<load>-<pallet>-<delta>`, and both writers drop anything the ledger
already holds. A later, genuinely different adjustment has a different delta and
its own id, so it still goes through.

Re-comparing after applying was already safe on its own — the system quantity
reads the corrected figure, the outcome comes back `AGREES` and nothing is
written.

## A cached frame outlives the release that built it

`KeyError: 'HU_SOURCE'`, from a session that had the report loaded before the
column existed. The upload signature is the file name and size — it does not
change when the code does, so nothing re-read the file and the old frame was
still being handed round with the old columns.

The parsed report now carries the module's API number beside it. On every run
the tab checks both that number and that the frame still has every column
`ACTUAL_COLS` declares; if either fails the frame is dropped and the screen says
so, rather than reaching for a column that is not there:

> *The report was read by an earlier version of this app. Upload it again to
> pick up the current columns.*

The reconciliation is checked the same way — one built before `BINDS` existed is
thrown away instead of being displayed.

The module version gate at the top of `app.py` catches a half-updated deploy on
disk. This is the same problem one level in: state that survived the update.

## Which HU is the pallet

`Starting Hu` is the pallet the stock came off, and it is the one the stock file
knows. Where the two columns differ on an outbound pick, `Starting Hu` is in the
inventory **97%** of the time against 34% for `Ending Hu` — and the leftover
provably stays on it: after picking 2 off `DONAL060826-E-13`, the stock file
still shows 25 on `E-13`, while `E-13A` is nowhere in it. So the balance belongs
to `Starting Hu`.

`Ending Hu` is kept as `TO_PALLET` all the same. On a split pick
(`DONAL060826-E-13 → DONAL060826-E-13A`, `DONAL070826-B-17 → DONA-000001`) that
is the HU the picked stock left on, which is what the driver has in his hand.

**Some picks are written with no `Starting Hu` at all** — 130 rows across real
invoice and challan loads in the sample report, and 1,612 movements once the
whole file is read. Those were being dropped, so a genuine pick simply vanished.
The `Ending Hu` now stands in, and `HU_SOURCE` on every row records which column
the pallet came from.

Two things came out of chasing that:

**A missing value is not the word "nan".** Under the string dtype pandas keeps it
as `pd.NA`, so `astype(str)` leaves it `NA` rather than `"nan"`, every
comparison after that goes three-valued, and the blank filter quietly passed
those rows through with a `NaN` pallet. They then survived `PALLET != ""`,
because `NaN != ""` is true.

**A receipt is not a pick.** `PO-342258027` and `3332627/STN-002` are goods
coming in. They have no `Starting Hu` either, so the fallback above would have
turned every receipt into a pick against a load that does not exist. They are
matched on the separator — start, `-` or `/` — and left out.

## A pallet the system never chose

The floor takes what is in front of it. When the report shows a pallet this load
never reserved, it **is** written to the database against that load — a
`PALLET_LEDGER` row with the load in `DOC_NUMBER`, the pallet, and a positive
quantity, plus the movement itself in `ACTUAL_PICKS`.

For that row to do anything it has to reach a stock row. `ledger_state()` groups
on `ROW_KEY`, falling back to `PALLET|ITEM|LOT` — and the report carries no lot,
so a correction built from the report alone has the key `P9|AAA|` while the
inventory holds `P9|AAA|LOT7`. They do not match, and the balance never moves.

So the keys are looked for in three places, best first:

1. the ledger rows for **this load** — the pick it is correcting
2. the **inventory** (`reconcile(keys=…)`, taken from the stock basis) — what the
   balance is actually measured against
3. the **rest of the ledger** — the same pallet under another load still carries
   its `ROW_KEY` and lot

That covers a pallet the system never chose but the warehouse plainly knows.

If none of the three can identify it, the row is still filed against the load,
but `BINDS` reads `no`, the reconciliation counts it, and the screen says so:

> *2 pallet(s) could not be matched to a stock row — they are recorded against
> the load, but they will not change a balance.*

A record that quietly moves nothing is worse than one that says it did nothing.
Loading the inventory report on the Pick tab and comparing again usually binds
them.

## Order complete, and what the floor actually picked

### Complete is not delete

| | Record | Pallet balance |
|---|---|---|
| **Delete** | every row removed — master, detail, ledger, registry | the **whole** reservation comes back, the document can be picked again |
| **Order complete** | every row **kept** — the load left the building and the record has to stand | only the quantity this load reserved but **never used** comes back |

Loads tab → **Order complete**. A load is never completed twice; the corrections
would be applied a second time and the balance would drift.

### The Transactions History Report

The ledger holds the pallets the system *chose*. The picker may well have taken
others. Until the two are put side by side, a pallet nobody touched stays locked
out of every later pick, and a pallet that really was emptied still looks full.

Tab: **Actual picks**. Upload the WMS report and press **Compare with the ledger**.

* `Control Number` carries the load — `INM0DONA-333262712295` is load
  `333262712295`, `INM0DONA-333/26-27/17` is `333/26-27/17`. A revision suffix
  (`…298-A`, `…321A`) matches back to its load; nothing else is guessed.
* `Starting Hu` is the pallet the stock actually came off.
* **A pallet moves rack → picker → staging → dock, and every leg is a row.**
  Summing them multiplies the pick by three or four — one load in the sample
  report reads 467 units flat and **130** once the chains are collapsed back to
  the single movement that took stock out of storage. A genuine second pick off
  the same pallet starts somewhere new, so it survives.

Each pallet lands in one of four outcomes:

| Outcome | Meaning |
|---|---|
| `AGREES` | reserved and taken match |
| `RELEASE` | the system reserved it, the floor never touched it — goes back on the rack |
| `CONSUME` | the floor took it, the system does not know — comes off the rack |
| `ADJUST` | the same pallet, a different quantity |

**Picked twice** lists any pallet appearing on more than one of the matched
loads. Sometimes that is right — a pallet can serve two orders — but it is also
what a double pick looks like.

Corrections are written into the ledger as ordinary rows with a **negative
quantity** where stock is handed back, so the running balance stays a plain sum
and nothing downstream needs to know those rows are different. Apply them once —
either from this tab, or by completing the load and letting that apply them.

Worksheets: `COMPLETED_LOADS` · `ACTUAL_PICKS`. Press **Set up worksheets** once.

## Partial pick — how much of a short line goes out

A short document can be sent two ways. The offer shows both numbers so the
choice is made on the figures, not on a guess:

| Column | What that load would carry |
|---|---|
| `CAN_PICK_NOW` | the complete lines **plus whatever is on the floor** for the short line — that line goes out split |
| `WHOLE_LINES_ONLY` | the complete lines **only** — the short item is left off the load entirely |

**Send whatever is on the floor** (the default) is the old behaviour: 12 asked
for, 4 on the floor, 4 go out and 8 stay owed.

**Send only the lines that are complete** leaves that item off. Nothing of it
ships, all 12 stay owed, and the customer gets whole line items rather than
fragments. `SHORT` on the shortage sheet reflects that — the whole line, not
the balance of a split.

### Leaving an item off by hand

Under *Leave an item off this load* any line can be held back, short or not.
It is recorded as `Left out of this load by the user`, the quantity stays owed
and the document comes back for it on the next pick, exactly like a shortage.

Holding a line back **is** a partial pick — the document is marked `PARTIAL`
and stays open in `DOC_REGISTRY` even without the short-delivery confirmation,
because the load no longer carries the whole document.

If every line is held back, nothing is written. There is no such thing as an
empty order.

### When nothing can be sent at all

`no_partial()` still names the documents where both numbers are zero, and each
one now carries its own `REASONS` — *on another pick task*, *item not in
inventory / plant* — so the next step is on the row rather than in a hunt.

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

## 25. Roles — Admin / Dashboard / Packing / Dispatch

App එක open කරාම login screen එකක්. හතරක් තෝරගන්න:

| Role | Access | Password |
|---|---|---|
| **Admin** | හැම එකක්ම — Pick, Dashboard, Register, Loads, SKU master, Search, Stock, History, Reset | ඔව් — default `Isha@1996`, `secrets.toml`→`[app] admin_password` එකෙන් වෙනස් කරන්න පුළුවන් (නැත්නම් reset password එකම) |
| **Dashboard** | **Pending vs picked** සහ **Invoice register** විතරයි — upload / pick / reset කිසිම දෙයක් නෑ | නෑ |
| **📦 Packing** | Packing station එක විතරයි — LOAD ID එකෙන් `PACKING` complete | නෑ |
| **🚚 Dispatch** | Dispatch station එක විතරයි — LOAD ID එකෙන් `DISPATCH` complete | නෑ |

Sidebar එකේ **Switch role / Log out** button එකෙන් ආපහු login screen එකට.
Session එකක් per browser tab — වෙන කෙනෙක් login වෙන එකෙන් මේකට බලපෑමක් නෑ.

---

## 26. Packing / Dispatch — floor stations

Login → **Packing** හෝ **Dispatch**. දෙකම **එකම screen එක**, update වෙන column
එක විතරයි වෙනස (`PACKING` / `DISPATCH`). Tab දෙකක් — දෙකෙන්ම එකම වැඩේ:

**⌨️ Type the LOAD ID**
* LOAD ID එක type කරලා **Mark … complete** click කරන්න
* QR එක කැඩිලා / light එක අඩු / number එක අතේ තියෙනවා නම් — මේක වේගවත්ම ක්‍රමය
* **Barcode gun එකකුත් මෙතන වැඩ කරනවා** — ඒක id එක type කරලා Enter ගහනවා,
  form එක ඒකෙන්ම submit වෙනවා. Submit වුණාම box එක clear වෙනවා, ඊළඟ එකට readyයි

**📷 Scan the QR code**
1. **Scan QR code** click කරාම phone එකේ native camera එක open වෙනවා
2. Pick sheet PDF එකේ **LOAD_ID QR** එකට point කරලා capture කරන්න
3. OpenCV එකෙන් decode කරලා `TAX_INVOICE_NO` එකට match කරනවා
4. **Scan the next one** button එකෙන් ඊළඟ එකට

හම්බුණොත් — `INVOICE_SUMMARY` සහ `INVOICE_DETAIL` දෙකේම ඒ column එක
**Completed** ලෙස update වෙනවා, ඒ instant එකේම save. LOAD_ID එක register එකේ
නැත්නම් / QR එකක් හම්බුනේ නැත්නම් error එකක්. එකම එක දෙපාරක් දැම්මොත්
*"already Completed"* කියනවා — කිසිම හානියක් නෑ.

Live video scan එකක් නෙවෙයි — photo capture + decode, ඒක mobile browser හැම
එකකම reliable විදිහට වැඩ කරන approach එක. Session එකේ complete කරපු ලැයිස්තුවක්
පහළින්.

---

## 27. Picking / Packing / Dispatch — status columns

`INVOICE_SUMMARY` සහ `INVOICE_DETAIL` දෙකේම තුන් column — `KORBER_PICK`
(මේ app එකෙන්ම pallet allocate කළාද) එකෙන් **independent**, physical warehouse
execution එක track කරන්නේ:

| Column | `Completed` වෙන්නේ කොහොමද |
|---|---|
| **PICKING** | `Pick_Live_status` upload — `Open Pick = 0` · හෝ Invoice sales report match (පහළින්) |
| **PACKING** | Packing station — QR scan හෝ LOAD ID type කිරීම (§ 26) |
| **DISPATCH** | `Pick_Live_status` upload — `Shipped Pick ≠ 0` හෝ `Total Pick = Shipped Pick` · **හෝ** Dispatch station එකෙන් manual (§ 26) |

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

---

## 29. Speed — දෙවෙනි round එක

Register එක ලොකු වෙන්න වෙන්න හෙමින් වුණු තැන් හතරක්. 5 000 invoice /
30 000 detail line එකකින් measure කරපු ඒවා:

| වැඩේ | කලින් | දැන් | |
|---|---|---|---|
| `dashboard()` — Dashboard tab එකේ හැම rerun එකකම | 393 ms | **41 ms** | 10× |
| `details_excel()` download | 9 485 ms | **4 537 ms** | 2× |
| `summary_excel()` download | 1 692 ms | **834 ms** | 2× |
| `merge_details()` — save lock එක ඇතුලේ | 136 ms | **40 ms** | 3× |
| `base_item()` × 30 000 (warm) | 32 ms | **7 ms** | 5× |

**මොකද කළේ:**

* **`parse_date` vectorise කළා** — `.map(parse_date)` කියන්නේ row එකකට
  `pd.to_datetime` call එකක්. දැන් `parse_dates()` එකෙන් column එකම එක පාරට.
  මේක තමයි ලොකුම එක (306 ms → 2.4 ms).
* **Excel writer එක `xlsxwriter`** — openpyxl එකට වඩා දෙගුණයක් වේගවත්, file
  එකත් 3ෙන් 1යි (1.9 MB → 0.6 MB). Install වෙලා නැත්නම් openpyxl එකට
  automatic fall back වෙනවා, ඒ නිසා කැඩෙන්නේ නෑ.
* **Column width එක sample එකකින්** — 30 000 row එකක්ම මනින්නේ නැතුව මුල්
  400න්. Width කියන්නේ cosmetic දෙයක්.
* **`merge_details` එකේ `.iterrows()` අයින් කළා** — indexed lookup එකකට.
  මේක save lock එක ඇතුලේ run වෙන නිසා අනිත් users ලාත් රැඳෙනවා.
* **`base_item` / `clean_item` cache කළා** (`lru_cache`) — SKU master
  10 000+ rows වලට හැම rerun එකකම call වෙනවා.
* **`_register_frames()` එක rerun එකකට එක පාරයි** — කලින් Dashboard tab
  එකයි Register tab එකයි දෙකෙන්ම call වෙලා දෙපාරක් run වුණා.

---

## 30. Register backfill — කලින් pick වුණු invoice

**ප්‍රශ්නය:** Register එක record කරන්න පටන් ගන්නේ ඒක on කරපු මොහොතේ ඉඳන්.
ඊට කලින් pick වුණු හැම document එකකටම `PALLET_LEDGER` එකේ pallet තියෙනවා, ඒත්
`INVOICE_SUMMARY` එකේ row එකක් නෑ. (මේ sheet එකේ: ledger එකේ **138**,
register එකේ **3**.)

**ආපහු upload කරලා හදන්න බෑ** — ඒ document එක `DOC_REGISTRY` එකේ තියෙන නිසා
duplicate gate එකෙන් `DUPLICATE (already processed)` කියලා skip වෙනවා. ඒ නිසා
Register tab → **Backfill** කියලා වෙනම tab එකක්.

### A. Pick history එකෙන් rebuild (file එකක් ඕන නෑ)

**Check what is missing** → ledger එකයි register එකයි compare කරලා කීයක් නැද්ද
කියලා පෙන්නනවා → **Rebuild N row(s)** click කරන්න.

`PALLET_LEDGER` එකෙන් හදනවා — pick එක all-or-nothing නිසා ඒ line එකේ
**picked qty එකම document qty එක**, ඒ නිසා හරියටම reconstruct කරන්න පුළුවන්:

| Register field | කොහෙන්ද |
|---|---|
| Lines · Qty · Picked qty | ledger එකේ `DOC_LINE` / `QTY_PICKED` |
| Item · Base ID · Description · Lot | ledger |
| Pallets · Locations · Plant | ledger (line එකකට කීපයක් නම් `,` වලින්) |
| Run id · Picked at · Source file | ledger |
| Tax invoice date · AR invoice no | `DOC_REGISTRY` |
| **Customer name** | **හිස්** — ඒක තියෙන්නේ PDF එකේ විතරයි (B බලන්න) |

`KORBER_PICK = Yes`, remark එකට *"Backfilled from the pick history"*.
Picking/Packing/Dispatch තුනම `Pending` — ඒවා floor එකේ ඇත්ත තත්වය, ledger
එකෙන් කියන්න බෑ. දෙපාරක් run කරොත් දෙවෙනි එකෙන් 0ක් add වෙනවා (idempotent).

> ඔයාගේ sheet එකෙන් verify කළා — 138ම invoice වල **qty, line count, doc type,
> doc date හතරම DOC_REGISTRY එකට හරියටම ගැලපුණා (138/138)**.

### B. Invoice PDF upload කරලා (customer name එකට)

එකම tab එකේ පහළින් — PDF ටික upload කරලා **Register N document(s)**.
**Pick run එකක් නැතුව** register එකට යනවා. ledger එකේ ඒ document එක තිබ්බොත්
`Yes` විදිහට, pallet/location එක්කම; නැත්නම් `No`.

මේකෙන් customer name එකත් එනවා (PDF එකේ තියෙන නිසා), ඒ නිසා A කරලා පස්සේ B
කරොත් row එක සම්පූර්ණයි.

---

## 31. DOC_REGISTRY column shift — fix එක

`DOC_REGISTRY` sheet එකේ header එක **පරණ 14 column** එකක්, ඒත් code එක
**16ක්** write කරනවා (`WMS_QTY` සහ `VERIFY` පස්සේ එකතු වුණා). `_ensure()`
header එක ලියන්නේ sheet එක **හිස්** නම් විතරයි — ඒ නිසා පරණ header එක
එහෙමම තිබිලා, අලුත් rows වල `PICKED_QTY` එකෙන් පස්සේ **හැම value එකක්ම
column 2ක් වමට** ගියා. ඒකයි `PROCESSED_AT` යටතේ `333-MUMBAI` තිබුණේ
(138න් **120ක්**).

**Fix:** `sheet_header()` — sheet එකේ **ඇත්ත header** එක කියවලා, නැති column
තියෙනවා නම් ඒවා **අගට** එකතු කරනවා (මැදට නෙවෙයි — එතකොට කලින් ලියපු rows වල
alignment එක කැඩෙන්නේ නෑ). Rows ලියන්නේ ඒ order එකට. ඒ නිසා අලුත් column
එකක් add කරාම මේක ආපහු වෙන්නේ නෑ.

> Backfill එක `DOC_REGISTRY` එකේ **මුල් columns විතරයි** කියවන්නේ
> (`DOC_NUMBER … PICKED_QTY`) — shift එක පටන් ගන්නේ ඊට පස්සේ නිසා
> ඒ කොටස විශ්වාසනීයයි. අනිත් හැම දෙයක්ම clean `PALLET_LEDGER` එකෙන්.

---

## 32. Körber dashboard එකෙන් කෙලින්ම Pick_Live_status ගන්න

Excel එකට export කරලා upload කරනවා වෙනුවට, dashboard එකෙන් **කෙලින්ම**
කියවන්න පුළුවන්. Register tab → **Update status** → *"Fetch from the Körber
dashboard instead of uploading"*.

1. **Dashboard URL** එක දාන්න — උදා: `http://130.61.243.161:8081/korber/pick/`
2. **Test** — connect වෙනවද, row කීයක් ආවද, format එක මොකක්ද කියලා පෙන්නනවා
   (ආපු table එකත් පෙන්නනවා)
3. **Fetch & apply** — upload එකට හරියටම සමානව Picking / Dispatch update වෙනවා
4. **Remember this URL** — `APP_SETTINGS` එකේ save වෙනවා, ඊළඟ පාරට type කරන්න ඕන නෑ

Login එකක් ඕන නම් **"It needs a login"** tick කරලා username/password හෝ
bearer token එකක් දෙන්න.

**Format එක auto-detect වෙනවා** — HTML page (table එක auto තෝරගන්නවා), CSV,
JSON (`[...]` හෝ `{"rows":[...]}`), හෝ ඇත්ත `.xlsx`. `Load Id` column එකක්
තියෙන table එක තෝරගන්නවා; නැත්නම් මොන column ආවද කියලා error එකේ පෙන්නනවා.
Rules ටික එකමයි — `Open Pick = 0` → Picking, `Shipped Pick ≠ 0` හෝ
`Total Pick = Shipped Pick` → Dispatch.

> **⚠️ Network එක:** app එක run වෙන machine එකෙන් ඒ host + port එකට reach
> කරන්න පුළුවන් වෙන්න ඕන.
> * **ඔයාලගේම server එකේ (same network)** → වැඩ කරනවා ✅
> * **Streamlit Cloud** → ඒ port එක public internet එකට open නම් විතරයි ❌
>
> Timeout එකක් ආවොත් — app එක run වෙන machine එකේ browser එකෙන් ඒ URL එක
> open වෙනවද කියලා බලන්න.

---

## 33. pandas 3 — `Invalid value '0.0' for dtype 'str'`

Load එකක් delete කරද්දී මේ error එක එනවා නම්:

```
Delete error: Invalid value '0.0' for dtype 'str'.
Value should be a string or missing value, got 'float' instead.
```

**හේතුව:** Google Sheet එකෙන් එන්නේ string විතරයි. **pandas 3.0** ඒවාට
`str` dtype එකක් දෙනවා — ඒ dtype එක number එකක් **භාර ගන්නේ නෑ**. කලින්
(pandas 2) ඒවා `object` වුණා, ඕනම දෙයක් භාර ගත්තා. Streamlit Cloud එකට
pandas 3 ආපු දවසේ ඉඳන් මේක කැඩෙනවා.

මේකෙන් කැඩෙන්නේ delete විතරක් නෙවෙයි — sheet එකෙන් කියවපු frame එකකට
number එකක් ලියන **හැම තැනම**:

| තැන | ලියන්නේ |
|---|---|
| Load delete → `mark_unpicked` | `PICKED_QTY = 0` |
| Backfill → `enrich_from_history` | `PICKED_QTY` (summary + detail) |

**Fix — root එකේම:** `gsheet._frame()` දැන් `dtype=object` එකෙන් frame එක
හදනවා, ඒ නිසා sheet එකෙන් එන හැම frame එකක්ම කලින් වගේම හැසිරෙනවා.
ඒ එක්කම `_writable()` කියලා guard එකකුත් තියෙනවා — number ලියන්න කලින්
ඒ column එක විතරක් widen කරනවා (30 000 row frame එකක් මුළුමනින්ම copy
කරන්නේ නෑ).

---

## 34. Dataflow / data mapping audit — හම්බුණු bug 10ක් සහ fix

"System එක මුල ඉඳන් check කරලා bugs තියෙනවා නම් හදන්න" කියලා කරපු
**systematic** audit එකක ප්‍රතිඵලය. Accident වලින් හම්බුණ ඒවා නෙවෙයි —
හැම boundary එකකම (PDF → DocLine → engine → register → sheet → dashboard)
column mapping එක mechanically cross-check කරලා හම්බුණු ඒවා.

හැම එකකටම `test_dataflow.py` එකේ test එකක් තියෙනවා, ඒ නිසා ආපහු කැඩුනොත්
වහාම හසුවෙනවා.

### Critical — pick එකට කෙලින්ම බලපානවා

**1. Base ID එක layer දෙකක දෙවිදිහකට හැදුනා**

| තැන | `P601560 710` → base | |
|---|---|---|
| Invoice line (`DocLine.base`) | `P601560` | ✅ |
| Inventory (`normalize_inventory`) | `P601560710` | ❌ |
| SKU master | `P601560` | ✅ |

`normalize_inventory` එකේ base එක හදාගත්තේ **clean කරපු** item number
එකෙන් — ඒකෙන් space එක delete වෙලා තිබ්බා, base_item එකට suffix එක
වෙන් කරගන්න **ඕනේ ඒ space එකමයි**. ඒ නිසා invoice එකේ තිබ්බ item එකට
ඒකේම stock එක හම්බුනේ නෑ → **"Item not in inventory / plant" →
STOCK SHORT**. Space තියෙන හැම code එකකටම (SKU_MASTER එකේ 10 377න්
631ක්) මේක වෙනවා.

*Fix:* base එක **raw code** එකෙන් හදනවා (`item_number_raw`), අනිත්
හැම තැනම වගේම.

**2. PDF parser එකම item code එක කඩලා දැම්මා**

`item_code=clean_item(cell)` — parse කරන වෙලාවෙම space එක නැති වුණා.
ඒ නිසා `INVOICE_DETAIL.BASE_ID` එකත් වැරදි වුණා, sales reconciliation
එකත් වැරදි key එකකින් හෙව්වා.

*Fix:* අලුත් `tidy_item()` — upper case + trim, **separator ඒ විදිහටම**.
Match කරද්දී තාමත් `clean_item()` හරහා, ඒ නිසා match key එක වෙනස් නෑ.

*ඒ එක්කම:* engine එකේ pool එකට exact-code fallback එකක් — document
එකේ `P601560710` (කඩන්න බැරි එක) සහ inventory එකේ `P601560 710`,
දෙකම එකම `clean_item` key එකට යනවා, ඒ නිසා දැන් හම්බෙනවා.

### High — data එක වැරදියට පෙන්නනවා

**3. Date parsing — column එකේ පළවෙනි row එකෙන් හැම row එකක්ම කැඩුනා**

`pd.to_datetime(column, dayfirst=True)` කරද්දී pandas **පළවෙනි value
එකෙන් format එකක් අනුමාන කරලා** ඒකට නොගැලපෙන හැම row එකක්ම `NaT`
කරනවා. ඒ නිසා register එකේ උඩම `2026-08-01` වගේ ISO date එකක් තිබ්බොත්
ඊට යටින් තිබ්බ **හැම `01/08/2026` එකක්ම blank** වුණා. ඒ විතරක් නෙවෙයි —
`dayfirst` නිසා `2026-08-01` කියෙව්වේ **8 January** කියලා.

තව: `parse_date` (එකක්) සහ `parse_dates` (column එක) **එකම value එකට
උත්තර දෙකක්** දුන්නා.

*Fix:* format list එකක් **පිළිවෙලට**, තාම parse නොවුණු row වලට විතරක්.
`parse_date` දැන් `parse_dates` එකෙන්ම හදලා — ආපහු වෙන් වෙන්න බෑ.
Speed එකට බලපෑමක් නෑ (5 000 row = 18 ms).

**4. Invoice number එක තැන් දෙකක දෙවිදිහකට match වුණා**

සමහර තැන් `_id_str` (Excel එකෙන් එන `30426013174.0` එකේ `.0` අයින්
කරන එක), සමහර තැන් plain `astype(str)`. Excel හරහා ආපු register එකක්
**එකම invoice එකට row දෙකක්** හැදුවා, dashboard එකට ඒකේම detail line
හම්බුනේ නෑ.

*Fix:* හැම path එකකම `_id_str`, සහ normalize වුණු එකම store වෙනවා.

**5. Sales report එක reconcile කරලා update කරේ නෑ**

Key හදද්දී `_id_str`, ආපහු register එකට ලියද්දී `astype(str)` — float
විදිහට ආපු invoice number වලට **match වුණාට update වුණේ නෑ**.

### Medium — SUMMARY සහ DETAIL එකිනෙකට විරුද්ධ වුණා

**6. Load එකක් delete කරද්දී DETAIL එක තාම "picked" කිව්වා**

`delete_load` / `register_unpick` දෙකම `INVOICE_SUMMARY` විතරයි reset
කරේ. `INVOICE_DETAIL` එකේ pallet, lot, location එහෙමම තිබ්බා — release
වුණු stock එකකට. Summary tab එකේ `No`, Details tab එකේ `Yes`.

*Fix:* අලුත් `mark_unpicked_details()` — pallet/lot/location/qty clear
වෙනවා. **Picking/Packing/Dispatch clear වෙන්නේ නෑ** (ඒවා floor එකේ
වැඩේ, මේ app එකේ allocation එක නෙවෙයි).

**7. Re-upload එකකදී DETAIL එක "No" වුණා, SUMMARY එක "Yes" වුණා**

`merge_summary` දැනටමත් තියෙන `Yes` එක රකිනවා. `merge_details` detail
row ටික **මුළුමනින්ම replace** කරා → allocation එක නැති වුණා.

*Fix:* merge_details එකත් pick එක carry කරගෙන යනවා.

**8. Scan එකක් DETAIL එකට reach වුණේ නෑ**

`scan_status` write එක skip කරේ summary row එක දැනටමත් `Completed`
නම්. ඒත් detail line එකක් තාම `Pending` වෙන්න පුළුවන් (scan එකෙන්
පස්සේ re-upload එකක් / backfill එකක් ආවොත්) → ඒක **හැමදාටම** Pending.

*Fix:* summary එකේ තත්වය නෙවෙයි, **ඇත්තටම වෙනස් වුණාද** කියලා බලනවා.

### Medium — save order

**9. Sheet එක "නෑ" කියන්න කලින්ම register එකට "picked" කියලා ලිව්වා**

`save_run` තමයි "මේ document එක තව කෙනෙක් දැනටමත් save කරලා" කියලා
තීරණය කරන්නේ. ඒත් register එක ලිව්වේ **ඒකට කලින්**. ඒ නිසා ඒ documents
ledger එකක්, registry row එකක්, WMS output එකක් නැති RUN_ID එකක් යටතේ
"picked" විදිහට file වුණා.

*Fix:* save කරලා, skip වුණු ඒවා demote කරලා, ඊට පස්සේ register එකට.

### Low — silent failure modes

**10. `gsheet.py` version gate එකෙන් පිටත තිබ්බා**

`gsheet.py` replace කරන්න අමතක වුණොත් app එක **කිසිම දෙයක් නොකියා**
පරණ code එකෙන් දිගටම දුවනවා — logic bug එකක් වගේ පේනවා. දැන් අනිත්
module වගේම gate එකේ.

*ඒ එක්කම:* sheet header එකේ එකම නම දෙපාරක් තිබ්බොත් `df["COL"]` එකෙන්
Series එකක් වෙනුවට DataFrame එකක් එනවා, ඊට පස්සේ තියෙන හැම
`.astype` / `.map` එකක්ම කැඩෙනවා. `_frame()` දැන් නම් unique කරනවා
(`A`, `A_2`) — පළවෙනි එකට හරි නම එහෙමම තියෙනවා.

### `test_engine.py` → `test_dataflow.py`

පරණ `test_engine.py` එකේ තිබ්බේ `pick_engine.run_pipeline()` කියලා
**මේ app එකේ නැති** API එකක් (HJ/SAP carton multiplier, Requirement
file). Import එකවත් වුණේ නෑ, ඒ නිසා ඒකෙන් **කිසිම දෙයක් check වුණේ නෑ**.

අලුත් `test_dataflow.py` එකේ test 23ක් — ඉහත හැම bug එකකටම එකක්:

```bash
python test_dataflow.py          # හෝ
python -m pytest test_dataflow.py
```

| Module | දැන් API |
|---|---|
| `doc_parser.py` | 7 |
| `pick_engine.py` | 5 |
| `invoice_register.py` | 15 |
| `gsheet.py` | 11 |
| `sku_master.py` | 3 |
| `pick_pdf.py` | 4 |
| `ui.py` | 3 |

**File ඔක්කොම replace කරන්න** — `app.py`, `doc_parser.py`, `pick_engine.py`,
`invoice_register.py`, `gsheet.py`, `sku_master.py`, `ui.py`, `pick_pdf.py`,
`test_dataflow.py`. Login screen එකේ පහළ BUILD එක බලලා verify කරන්න.

---

## 35. Partial pick — "දැන් තියෙන ටික යවමු"

Invoice එකක line එකක් හරි stock short නම් කලින් **මුළු document එකම**
block වුණා (all-or-nothing). Customer එකාට wait කරන්න බැරි වුණාම දැන්
**තියෙන ප්‍රමාණය යවලා, ඉතුරු ටික ණයට තියාගන්න** පුළුවන්.

### පාවිච්චි කරන විදිහ

1. සාමාන්‍ය විදිහට **Generate pick** කරන්න.
2. Stock short නම් Result එකට යටින් **"Send what we have"** කියලා
   section එකක් එනවා — මොන document එකෙන් දැන් කීයක් යවන්න පුළුවන්ද
   කියලා පේනවා.

   | DOC_NUMBER | SHORT_LINES | REQUIRED | AVAILABLE_NOW | STILL_SHORT |
   |---|---|---|---|---|
   | 30426013174 | 1 | 12 | 4 | 8 |

   මේ අංක **short lines වලට විතරයි** — full stock තියෙන lines කොහොමත්
   සම්පූර්ණයෙන් යනවා.
3. Document තෝරලා **"The customer has agreed to a short delivery"**
   tick කරලා **Pick what is available** ඔබන්න.
4. Pick එක ආපහු run වෙනවා. දැන් OutBound MASTER / Detail file වල
   තියෙන්නේ **ඇත්තටම යවන ප්‍රමාණය** — invoice qty එක නෙවෙයි.

Tick box එක **confirmation එක**. ඒක නොදාම partial pick එකක් වෙන්නේ නෑ,
default එක තාමත් all-or-nothing.

### Pick Sheet එකේ

Partial pick එකක Pick Sheet එකේ උඩම **PARTIAL** කියලා පැහැදිලිව
තියෙනවා:

| PICK STATUS | PICKING NOW | ALREADY SENT | STILL OWED |
|---|---|---|---|
| PARTIAL | 14 | 0 | 8 |

> **PARTIAL PICK — THIS IS NOT THE WHOLE DOCUMENT.** Only the quantity
> shown under PICKING NOW is being picked today … do not top them up
> from the invoice.

Floor එකේ කෙනෙක් "අඩුයි නේ" කියලා invoice එකෙන් හදන්න යන එක නවත්තන්න.

### ඉතුරු ටික පස්සේ pick කරන එක

Stock ආවම **ඒම invoice එකම ආපහම upload කරලා Generate pick** කරන්න.
System එකම ඉතුරු ටික විතරක් ගන්නවා:

* `DOC_REGISTRY` එකේ `PICK_STATUS = PARTIAL` නම් ඒ document එක
  **duplicate විදිහට block වෙන්නේ නෑ** — තාම ණයක් තියෙනවා.
* `PALLET_LEDGER` එකෙන් **line එකකට කලින් කීයක් ගියාද** කියලා බලලා ඒක
  අඩු කරනවා.
* සම්පූර්ණ වුණු line එක **හැර යනවා**, අඩුවෙන් ගිය line එකට **ඉතුරු
  ටික විතරක්** ඉල්ලනවා.

උදාහරණයක් — Invoice: AAA 10 + BBB 12 (= 22). BBB තියෙන්නේ 4යි:

| Run | AAA | BBB | OutBound Detail | PICK_STATUS | ණයට |
|---|---|---|---|---|---|
| 1 (partial) | 10 | 4 | **14** | PARTIAL | 8 |
| 2 (stock ආවම) | — | 8 | **8** | FULL | 0 |

Run 2 එකේදී AAA line එක **කොහෙත්ම නැවත pick වෙන්නේ නෑ**.

### Register එකේ — තුන්වෙනි තත්වය

`KORBER_PICK` එකට දැන් **`Partial`** කියලා තත්වයක් තියෙනවා:

| තත්වය | තේරුම |
|---|---|
| `No` | කිසිම දෙයක් pick වෙලා නෑ |
| `Partial` | ටිකක් ගිහින්, ඉතුරු ටික තාම ණයට |
| `Yes` | මුළු document එකම pick වෙලා |

* **ආපස්සට යන්නේ නෑ.** `Partial` එකක් පස්සේ run එකකින් `No` වෙන්නේ නෑ —
  `Yes` වෙන්න විතරයි පුළුවන් (`Partial` → `Yes`). Load එකක් delete
  කරොත් විතරයි `No` වෙන්නේ.
* **INVOICE_DETAIL එකේ line එකකට එකේම උත්තරය.** Partial document එකක
  සම්පූර්ණයෙන් ගිය line එක `Yes`, අඩුවෙන් ගිය එක `Partial`, කිසිසේත්ම
  නොගිය එක `No`.
* **REMARK** එකේ `PARTIAL PICK — 8 of 22 still owed` කියලා තියෙනවා.

### Dashboard එකේ

* **Partial** — ටිකක් ගිහින් තියෙන invoice ගාන.
* **Still to pick** — ඇත්තටම **ඉතුරු** ප්‍රමාණය (මුළු invoice qty එක
  නෙවෙයි). Partial invoice එකක් pending කියලා ගණන් ගන්නවා, ඒත් ණය
  ප්‍රමාණය විතරයි මේකට එකතු වෙන්නේ.
* "Why they are waiting" chart එකේ **Partially picked** කියලා අලුත්
  reason එකක්.

### DOC_REGISTRY එකේ අලුත් column 3ක්

| Column | තේරුම |
|---|---|
| `PICK_STATUS` | `FULL` හෝ `PARTIAL` |
| `TOTAL_PICKED` | Document එකට **මුළුමනින්ම** ගිය ප්‍රමාණය (හැම run එකම) |
| `SHORT_QTY` | තාම ණයට තියෙන ප්‍රමාණය |

`PICKED_QTY` කියන්නේ **ඒ run එකේදී විතරක්** ගිය ප්‍රමාණය.

මේ column තුන sheet එකේ **අන්තිමට add වෙනවා** (§31 බලන්න), ඒ නිසා
තියෙන row වල data එහෙමම තියෙනවා. පරණ row වල `PICK_STATUS` හිස්ය —
ඒවා ඔක්කොම full pick, ඒ නිසා හිස් එකක් = `FULL` විදිහට කියවනවා.

### Qty verify එක

Full pick එකකදී line · document total · WMS file total **තුනම හරියටම**
match වෙන්න ඕන — ඒක වෙනස් වෙලා නෑ.

Partial pick එකකදී අඩුවෙන් යවන එක **තීරණයක්**, error එකක් නෙවෙයි.
ඒ නිසා check වෙන්නේ:

* WMS file එකට ගියේ **ඇත්තටම pallet එකෙන් ගත්ත ප්‍රමාණයමද** ✔
* ඕන ප්‍රමාණයට වඩා **වැඩියෙන් ගත්තේ නැද්ද** ✔

Short line එකට `⚠️ SHORT` කියලා පෙන්නනවා — `❌ MISMATCH` නෙවෙයි.

### Test

`test_dataflow.py` එකේ partial pick එකට test 10ක් තියෙනවා (එකතුව 35):
default all-or-nothing, offer එකේ අංක, WMS file එකේ qty, ණය ගණන,
balance run එකේ line skip වීම, `Partial` ආපස්සට නොයාම, dashboard
ගණන්, pick sheet එකේ PARTIAL banner.

```bash
python test_dataflow.py
```

| Module | දැන් API |
|---|---|
| `doc_parser.py` | 7 |
| `pick_engine.py` | 6 |
| `invoice_register.py` | 16 |
| `gsheet.py` | 12 |
| `pick_pdf.py` | 5 |
| `sku_master.py` | 3 |
| `ui.py` | 3 |

BUILD: `2026-08-21 · partial pick`

---

## 36. "These files are out of date" — version gate එක

App එක පටන් ගද්දී මේක ආවොත්:

```
4 of these files are out of date — replace them and redeploy.
 ❌ pick_engine.py      — found API 5, needs 6
 ❌ pick_pdf.py         — found API 4, needs 5
 ❌ invoice_register.py — found API 15, needs 16
 ❌ gsheet.py           — found API 11, needs 12
 ✅ doc_parser.py — API 7   ✅ sku_master.py — API 3   ✅ ui.py — API 3
```

**තේරුම:** `app.py` අලුත් එකෙන් replace වෙලා, ඒත් ලයිස්තුවේ තියෙන file
තාම **පරණ** ඒවා. App එක crash වෙන්න දෙනවා වෙනුවට කලින්ම නවත්තලා
මොකක්ද replace කරන්න ඕනේ කියලා කියනවා.

**Fix:** release එකේ **file ඔක්කොම** replace කරන්න — flag වුණු ඒවා
විතරක් නෙවෙයි. එකම release එකේ file ටික එකට යන්න ඕන.

### ⚠️ මේ gate එකෙන් **හසු නොවෙන** එක

Gate එකට හසුවෙන්නේ **module එක `app.py` එකට වඩා පරණ නම්** විතරයි.
ඊට **අනිත් පැත්ත** — module අලුත්, `app.py` පරණ — හසුවෙන්නේ නෑ.
එතකොට error එකක් නෑ, අලුත් screen එක **හම්බෙන්නෙත් නෑ**, "මොකුත්ම
වුණේ නෑ" වගේ පේනවා.

ඒ නිසා deploy කරාට පස්සේ **login screen එකේ පහළ BUILD එක** බලන්න:

```
2026-08-21 · partial pick
```

ඒක අලුත් එක නම් `app.py` හරි. Admin sidebar එකේ හැම module එකකේම API
number එකත් තියෙනවා:

```
engine 6 · parser 7 · register 16 · sheet 12 · pdf 5 · sku 3 · ui 3
```

### `gsheet.py` — කලින් හැංගිලා තිබ්බා

කලින් `gsheet.py` check වුණේ **වෙනම, දෙවෙනි gate එකකින්**. පළවෙනි gate
එක `st.stop()` කරන නිසා `gsheet.py` පරණ නම් ඒක **පේන්නේම නෑ** — මුල්
තුන හදලා redeploy කරාට පස්සේ තමයි "දැන් gsheet.py" කියලා එන්නේ.

දැන් ඔක්කොම **එකම ලයිස්තුවේ**, එක screen එකකින් සම්පූර්ණ පිළිතුර
එනවා — up-to-date ඒවත් ✅ දාලා පෙන්නනවා.

---

## 37. PDF නැතුව pick එකක් — manual entry

PDF එක අතේ නැති වෙලාවට Invoice No + Item Code + Qty type කරලා pick එකක්
හදාගන්න පුළුවන්.

### කරන විදිහ

**Pick tab → "No PDF? Enter the invoice by hand"** expander එක.

| Field | |
|---|---|
| Invoice / DC number ★ | LOAD ID එක මේකයි — QR එකට යන්නෙත් මේකමයි |
| Document type | INVOICE / DELIVERY CHALLAN |
| Document date | Default අද |
| Customer | Optional (register එකට යනවා) |
| AR / Order no | Optional |

ඊට යටින් line grid එක:

| Item Code | Description | Qty | UOM |
|---|---|---|---|
| P601560 710 | — | 10 | EA |
| 1C072323 | — | 4 | EA |

* **Item Code** එකට **base ID එක හරි සම්පූර්ණ code එක හරි** දෙන්න
  පුළුවන් — `P601560 710` සහ `P601560` දෙකම එකම stock එක හොයනවා.
* හිස් row ගණන් ගන්නේ නෑ (grid එකට row ඕන තරම් add කරන්න පුළුවන්).
* Inventory report එක upload කරලා තියෙනවා නම් **type කරන ගමන්ම** මොකක්ද
  match වෙන්නේ කියලා පේනවා:

  | Item Code | Base ID | Qty | In stock (free) | |
  |---|---|---|---|---|
  | P601560 710 | P601560 | 10 | 48 | ✅ |
  | 1C072323 | 1C072323 | 4 | 2 | ⚠️ short |

  මේක **බැලීමක් විතරයි** — reservation එකක් නෙවෙයි.

**Add this document** ඔබන්න → document එක run එකට එකතු වෙනවා. ඊට පස්සේ
**සාමාන්‍ය විදිහටම** Generate pick කරන්න.

### මොකද වෙන්නේ

Manual document එකයි PDF එකෙන් ආපු එකයි **පසුව එකම විදිහට හැසිරෙනවා** —
වෙනම code path එකක් නෑ:

* එකම stock check (base ID match, pallet ledger, Pick Id gate)
* එකම OutBound MASTER / Detail file
* එකම Pick Sheet PDF (QR එක්ක)
* `PALLET_LEDGER` · `DOC_REGISTRY` · `INVOICE_SUMMARY` · `INVOICE_DETAIL`
* Stock short නම් **partial pick** එකත් offer වෙනවා (§35)
* Packing / Dispatch QR scan එකත් වැඩ කරනවා

`SOURCE_FILE` එකට `manual entry` කියලා යනවා, ඒ නිසා register එකේ මොනවද
type කරලා ආපු ඒවා කියලා පැහැදිලියි.

### Qty verify එක

PDF එකක `Total Quantity` / `Total Amount` කියලා **cross-check කරන්න
number** තියෙනවා. Manual entry එකකට එහෙම දෙයක් නෑ — ඒ නිසා ඒ check දෙක
**deliberately skip** කරනවා. Verify කරන්නේ **type කරපු ප්‍රමාණයට** pick
එක ගැලපෙනවද කියලා (line · document total · WMS file total).

### පරිස්සම් වෙන්න

Manual entry කියන්නේ **document එකක් නෙවෙයි, document එකක් ගැන කියන
කතාවක්**. Type කරන අංකය වැරදුනොත් system එකට ඒක දැනගන්න ක්‍රමයක් නෑ.
Invoice number එක `DOC_REGISTRY` එකේ දැනටමත් තියෙනවා නම් duplicate gate
එකෙන් නවත්තනවා, ඒත් **qty එකක් වැරදියට type කරොත් ඒක pick වෙනවා**.

---

## 38. Period filter — අද / date range / full data

**Dashboard (Pending vs picked)** එකේ සහ **Invoice register** එකේ දෙකේම
එකම filter එක තියෙනවා, එකම විදිහට වැඩ කරනවා — ඒ නිසා tab දෙකේ අංක
කවදාවත් වෙනස් වෙන්නේ නෑ.

```
[ Today ] [ Yesterday ] [ Last 7 days ] [ Last 30 days ] [ This month ] [ All time ] [ Custom range ]

Date to use: [ Invoice date ▾ ]      Today · 22 Aug 2026      Document type: [ ▾ ]
```

* **Today** — එක click එකක්. වැඩිපුරම අහන ප්‍රශ්නය ඒක නිසා.
* **Custom range** — From / To date box දෙකක් එනවා.
* **All time** — මුළු register එකම (default එක).

### "මොන date එකද" කියන එකත් තෝරන්න පුළුවන්

මේ තුන **වෙනස් ප්‍රශ්න තුනක්**:

| Date to use | අහන්නේ |
|---|---|
| **Invoice date** | "මේ කාලේට customer එකාට invoice කරලා තියෙන්නේ මොනවද" |
| **Picked date** | "අද warehouse එකෙන් ඇත්තටම pick කරේ මොනවද" |
| **Last updated** | "අද මොකක් හරි වුණේ මොනවටද" (scan, status report, pick) |

### හිස් date එකකට මොකද වෙන්නේ

| Column | හිස් නම් |
|---|---|
| Invoice date | **තියාගන්නවා** — date එක PDF එකෙන් කියවගන්න බැරි වුණා කියලා ඒ invoice එකේ වැඩේ ඉවර වෙන්නේ නෑ |
| Picked date / Last updated | **අයින් කරනවා** — ඒ දේ **වුණේම නෑ**, ඒ නිසා ඒක මේ කාලෙන් පිටත |

මේක screen එකේ caption එකකින් කියනවා, ඒ නිසා අංකයක් අඩුවෙන් පේනවනම්
ඇයි කියලා පැහැදිලියි.

### Download ඔක්කොම filter එකට යටත්

Dashboard එකේ **Summary / Details / Pending (Excel) · Pending (CSV)**
හතරම **screen එකේ තියෙන filter එකටම** අදාළව එනවා — යටින් තියෙන caption
එකේ `42 invoices · 118 lines · 9 pending` කියලා තහවුරු කරලා තියෙනවා.

### Register එකේ — filter වෙන්නේ **view එක විතරයි**

Register tab එකේ filter එක Summary / Details table වලට සහ KPI වලට
විතරයි බලපාන්නේ. **Update status** තාමත් **මුළු register එකම** බලනවා.

ඇයි: Packing bench එකේ QR එකක් scan කරද්දී screen එක "Today" කියලා
filter වෙලා තිබ්බොත්, ඊයේ invoice එකක් scan කරාම "not found" කියලා
එන්න බෑ. ඒ නිසා scan එක filter එකට යටත් නෑ.

KPI row එකේ "Invoices" එකට යටින් `138 on file altogether` කියලා පේනවා
— filter එකෙන් කීයක් අයින් වුණාද කියලා දැනගන්න.

### Test

```bash
python test_dataflow.py     # 43 tests
```

Manual entry එකට 4ක්, period filter එකට 4ක් — preset වල date ගණන්,
"Today" tab දෙකේම එකම දේ කියනවද, හිස් date rule එක, summary/details
එකට යනවද කියලා. App එකේ **preset 7 × date basis 3 = 21ම** tab දෙකේම
render කරලා verify කරලා තියෙනවා.

| Module | දැන් API |
|---|---|
| `doc_parser.py` | 8 |
| `pick_engine.py` | 6 |
| `invoice_register.py` | 17 |
| `gsheet.py` | 12 |
| `pick_pdf.py` | 5 |
| `sku_master.py` | 3 |
| `ui.py` | 3 |

BUILD: `2026-08-22 · manual pick + period filter`

---

## 39. "Send what we have" එක නොපෙනෙන්නේ ඇයි

Stock short වුණාට partial pick section එක නොපෙනුනොත් හේතු දෙකයි.

### 1. ඒ document එකෙන් **යවන්න මොකුත්ම නෑ**

Line එකේ `AVAILABLE` එක **0** නම් (screen එකේ `Item not in inventory /
plant`), partial pick එකකින් trolley එකට යන්නේ **0**. ඒක partial pick
එකක් නෙවෙයි — pick එකක්ම නෙවෙයි. ඒ නිසා confirm කරන්න දෙයක් නෑ.

දැන් ඒක **හංගන්නේ නෑ**, පැහැදිලිව කියනවා:

```
⛔ nothing to send
A partial pick would put zero units on the truck for these documents.

DOC_NUMBER      LINES  SHORT_LINES  DOC_QTY  CAN_PICK_NOW  STILL_SHORT
333262712447        1            1        3             0            3
```

**බලන්න ඕන තැන් තුනක්** (`CAN_PICK_NOW` 0 වුණාට කලින්):

| | |
|---|---|
| **Plant** | Step 03 එකේ confirm කරපු plant එකේ නැත්නම් හම්බෙන්නේ නෑ. Item එක වෙන plant එකක තියෙන්න පුළුවන් |
| **Status** | Sidebar එකේ status filter — default එකට `Available` විතරයි ගණන් ගන්නේ |
| **Pick task** | `On a pick task — excluded` කියන expander එක බලන්න. Pick Id ≠ 0 නම් ඒ pallet locked. **Release stock** section එකෙන් release කරන්න පුළුවන් |

### 2. **Bug එකක් තිබ්බා — 2026-08-22b එකේ හදලා**

Line 4ක් තියෙන invoice එකක line 1ක stock නෑ, අනිත් 3ට ඕන තරම් තියෙනවා
කියමු. Partial pick එකකින් **ඒ 3ම යවන්න පුළුවන්**.

ඒත් offer එක හදාගත්තේ **shortage table එකෙන් විතරයි** — ඒකේ තියෙන්නේ
**short lines විතරයි**. ඒ නිසා `AVAILABLE_NOW = 0` කියලා පෙන්නලා
document එක **list එකෙන්ම අයින් වුණා**. Engine එකට ඒක හරියටම pick කරන්න
පුළුවන් වුණා — user එකාට **අහන්නවත් ලැබුනේ නෑ**.

දැන් offer එක හදන්නේ `run_pick` ඇතුලේ, ඇත්ත අංක තියෙන තැනම:

```
CAN_PICK_NOW = (සම්පූර්ණයෙන් allocate වුණු lines)  +  (short lines වල තියෙන ටික)
```

| DOC_NUMBER | LINES | SHORT_LINES | DOC_QTY | ALREADY_SENT | CAN_PICK_NOW | STILL_SHORT |
|---|---|---|---|---|---|---|
| I1 | 4 | 1 | 18 | 0 | **15** | 3 |

`ALREADY_SENT` — කලින් partial pick එකකින් ගිය ප්‍රමාණය (§35).

BUILD: `2026-08-22b · partial offer fix` · `pick_engine.py` API **7**

---

## 40. Pick email එකේ — Doc Qty එකෙන් කීයක් pick කරාද

Partial pick එකකදී email එකෙන් **ඇත්තටම යවන ප්‍රමාණය** කියන්න ඕන.

### කලින් තිබ්බ ප්‍රශ්නය

Email එකේ පහළ තිබ්බේ:

```
Total picked qty: 22
```

ඒත් ඒක ගණන් හැදුවේ **document qty එකෙන්** — pick කරපු ප්‍රමාණයෙන්
නෙවෙයි. Full pick එකකදී දෙකම එකයි, ඒ නිසා අවුලක් තිබ්බේ නෑ. ඒත්
**partial pick** එකකදී 14ක් යවලා email එකෙන් **22ක් යවනවා** කියලා
කිව්වා. Gate එකේදී තර්කයක් හැදෙන්නේ ඒකෙන්.

### දැන් එන email එක

```
Subject: PARTIAL OutBound Pick · LOAD ID 30426013174

PARTIAL PICK — 30426013174. 14 of 25 pcs are being picked;
11 pcs are still owed and will follow when the stock arrives.

LOAD ID       : 30426013174   [PARTIAL]
Document      : INVOICE 30426013174  (22-AUG-2026)
Plant         : PL1
Document qty  : 25 pcs over 3 lines
Picking now   : 14 pcs
Still owed    : 11 pcs
Pallets       : 2
Qty check     : ⚠️ PARTIAL

Line summary — document qty vs picked:
Ln  Item Code    Item Number  Doc Qty  Picked  Short
--  -----------  -----------  -------  ------  -----
1   P601560 710  P601560 710       10      10      0
2   1C072323     1C072323          12       4      8
3   R010077                         3       0      3

Pick details:
Ln  Item Number  Pallet  Location  Qty  Balance
--  -----------  ------  --------  ---  -------
 1  P601560 710  PAL1    A1         10       90
 2  1C072323     PAL2    A2          4        0

Total document qty: 25
Total picked qty  : 14
Still owed        : 11
```

* **Subject** එකට `PARTIAL` කියලා ඉස්සරහින් එනවා
* **Banner** එකක් — HTML එකේ කහ පාට box එකක්
* **Line summary** — line එකකට **Doc Qty · Picked · Short**. Short line
  HTML එකේ කහ පාටින් shade වෙනවා
* **Footer** එකේ අංක තුනම වෙන වෙනම — එකම number එකකින් වැඩ තුනක්
  ගන්නේ නෑ

### Balance run එකේදී

ඉතුරු ටික pick කරද්දී **කලින් ගිය ප්‍රමාණයත්** එනවා:

```
Document qty  : 22 pcs over 2 lines
Picking now   : 8 pcs
Already sent  : 14 pcs (earlier pick)

Line summary — document qty vs picked:
Ln  Item Code  Item Number  Doc Qty  Picked  Short
--  ---------  -----------  -------  ------  -----
1   AAA                          10      10      0
2   BBB        BBB               12      12      0

Total document qty: 22
Total picked qty  : 8
Already sent      : 14 (earlier pick)
Delivered in all  : 22
```

Line summary එකේ `Picked` කියන්නේ **document එකට මුළුමනින්ම ගිය
ප්‍රමාණය** (මේ run එක + කලින්), ඒ නිසා `Short 0` කියලා පේනවා —
document එක දැන් සම්පූර්ණයි.

### Full pick එකකට

Partial නොවුණාම `PARTIAL` කිසිම තැනක නෑ, `Still owed` line එකත් නෑ.
`Doc Qty` සහ `Picked` දෙකම එකයි, `Short` හැම line එකකම 0.

### Pick Sheet PDF එකේ

PDF එකේ **කලින් ඉඳන්ම** `QUANTITY VERIFICATION — Doc Qty vs Picked Qty`
table එකක් තියෙනවා, ඒ එක්කම §35 එකේ `PARTIAL` header block එකත්. ඒ නිසා
PDF එකට වෙනසක් ඕන වුණේ නෑ.

BUILD: `2026-08-23 · partial pick email` · `pick_pdf.py` API **6**
