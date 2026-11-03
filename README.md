<p align="center">
  <img src="assets/logo-128.png" width="96" height="96" alt="ProWebLook Logo" />
</p>

<h1 align="center">ProWebLook for Google Sheets</h1>

<p align="center">
  <strong>Verify email addresses, validate global phone numbers (carrier &amp; line type), and check WhatsApp registration directly inside Google Sheets.</strong><br/>
  Works both as native custom formulas (<code>=PROWEBLOOK_VERIFY_EMAIL(A2)</code>) and as a permanent-value <strong>Bulk Lead Enricher Sidebar</strong>.
</p>

<p align="center">
  <a href="https://proweblook.com">Official Website</a> •
  <a href="https://proweblook.com/user/my_profile">Get Your API Key</a> •
  <a href="https://proweblook.com/openapi.json">OpenAPI 3.0 Spec</a>
</p>

---

## Key Features

- **4 Native Spreadsheet Formulas** with Google Sheets autocomplete, 2D range batching (`UrlFetchApp.fetchAll`), and 6-hour `CacheService` SHA-256 deduplication:
  - `=PROWEBLOOK_VERIFY_EMAIL(A2, [fullDetails])`
  - `=PROWEBLOOK_VALIDATE_PHONE(B2, [countryCode], [fullDetails])`
  - `=PROWEBLOOK_WHATSAPP(B2, [fullDetails])` *(alias: `=PROWEBLOOK_CHECK_WHATSAPP`)*
  - `=PROWEBLOOK_BALANCE([product])`
- **Bulk Lead Enricher Sidebar (`ProWebLook → Open Sidebar`)**:
  - Highlight a column of 10 or 1,000+ leads and write **permanent values + status colors** directly into the adjacent columns — reopening your sheet **never** re-runs formulas or burns extra credits.
  - Built-in **Skip already-verified rows** protection, live progress bar, `Stop` button, and real-time credit balance counters (`Email`, `Phone`, `WhatsApp`).
- **Google Sheets Phone Quirks Handled Automatically**:
  - Restores leading `+` when Google Sheets strips `+14155552671` into numeric `14155552671`.
  - Detects accidental scientific-notation paste (`1.41556E+10`) before calling the API so malformed cells never consume credits.
- **Least-Privilege Security (`@OnlyCurrentDoc`)**:
  - Uses `spreadsheets.currentonly` and restricts external HTTPS requests strictly to `https://proweblook.com/api/v1/zapier/` via `urlFetchWhitelist`.

---

## 60-Second Setup in Any Google Sheet (No CLI Required)

1. Open any Google Sheet and click **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with [`src/Code.gs`](src/Code.gs).
3. Click the **+** icon next to **Files** → choose **HTML** → name it **`Sidebar`** (exact case, without `.html`) → paste [`src/Sidebar.html`](src/Sidebar.html).
4. Click **Save** (`Ctrl+S` / `Cmd+S`), reload your Google Sheet, and click **ProWebLook → Open Sidebar (Connect & Bulk Verify)** in the top menu bar.
5. Paste your API Key from [proweblook.com/user/my_profile](https://proweblook.com/user/my_profile) and click **Connect Account**.

---

## Formula Reference

### 1. Verify Email Address
```excel
=PROWEBLOOK_VERIFY_EMAIL(A2)
```
Returns `"deliverable"`, `"undeliverable"`, or `"unknown"`.

Spill all 3 columns (`[verdict, reason, normalized_email]`):
```excel
=PROWEBLOOK_VERIFY_EMAIL(A2, TRUE)
```

Verify an entire range in parallel (up to 25 uncached rows per formula):
```excel
=PROWEBLOOK_VERIFY_EMAIL(A2:A20)
```

### 2. Validate Phone Number (Carrier & Line Type)
```excel
=PROWEBLOOK_VALIDATE_PHONE(B2)
=PROWEBLOOK_VALIDATE_PHONE(B2, "US")
```
Returns `"valid"`, `"invalid"`, or `"unknown"`.

Spill all 5 columns (`[verdict, linetype, carrier, country_code, e164_number]`):
```excel
=PROWEBLOOK_VALIDATE_PHONE(B2, "US", TRUE)
```

### 3. Check WhatsApp Registration (Sends No Message or OTP)
```excel
=PROWEBLOOK_WHATSAPP(B2)
```
Returns `"on_whatsapp"` or `"not_on_whatsapp"`.

Spill all 4 columns (`[verdict, account_type, formatted_number, country_code]`):
```excel
=PROWEBLOOK_WHATSAPP(B2, TRUE)
```

### 4. Check Remaining Credit Balance (Free — 0 Credits)
```excel
=PROWEBLOOK_BALANCE()
=PROWEBLOOK_BALANCE("email")
```

---

## Developer Deployment via `clasp`

```bash
git clone https://github.com/dubeychandan996/proweblook-google-sheets.git
cd proweblook-google-sheets
npm test

npx @google/clasp login
cp .clasp.json.example .clasp.json
# Paste your Apps Script Project ID into .clasp.json
npm run push
```

See [`MARKETPLACE_SUBMISSION.md`](MARKETPLACE_SUBMISSION.md) for publishing as a public **Google Workspace Marketplace** Editor Add-on.

---

## License

MIT © [ProWebLook](https://proweblook.com)
