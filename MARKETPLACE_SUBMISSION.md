# Google Workspace Marketplace Submission Guide (`ProWebLook`)

This repository is pre-configured for **Google Workspace Marketplace** review with least-privilege scopes (`@OnlyCurrentDoc` / `spreadsheets.currentonly`) and a strict `urlFetchWhitelist`.

---

## 1. Create Standalone Apps Script + Standard GCP Project

1. Open [script.google.com](https://script.google.com) → **New project** → name it **ProWebLook for Google Sheets**.
2. Push `src/` (`Code.gs`, `Sidebar.html`, `appsscript.json`) via `npx @google/clasp push` (or paste manually and enable **Show "appsscript.json" manifest file in editor** under Project Settings).
3. Open [console.cloud.google.com](https://console.cloud.google.com) → create or select a Google Cloud project (`proweblook-workspace-addon`) → copy the **Project Number**.
4. In Apps Script → **Project Settings** → **Google Cloud Platform (GCP) Project** → **Change project** → paste the **Project Number**.

---

## 2. Configure OAuth Consent Screen (Non-Sensitive Scopes!)

In Google Cloud Console → **APIs & Services → OAuth consent screen**:
- **App Name:** `ProWebLook`
- **User Support Email:** `support@proweblook.com`
- **App Logo:** Upload `assets/logo-128.png`
- **Application Home Page:** `https://proweblook.com`
- **Privacy Policy Link:** `https://proweblook.com/generic/terms_conditions` (or your dedicated privacy URL)
- **Terms of Service Link:** `https://proweblook.com/generic/terms_conditions`
- **Authorized Domains:** `proweblook.com`
- **Scopes (all narrow / Editor Add-on standard):**
  - `https://www.googleapis.com/auth/spreadsheets.currentonly`
  - `https://www.googleapis.com/auth/script.container.ui`
  - `https://www.googleapis.com/auth/script.external_request` (restricted by `urlFetchWhitelist` in `appsscript.json` to `https://proweblook.com/api/v1/zapier/`)

### Scope Justification Copy (for Google Trust & Safety)
> - **`spreadsheets.currentonly`**: Required to read selected email/phone cells in the active Google Sheet and write verification verdicts to adjacent cells when invoked by the user. Does not access any other Drive files or spreadsheets.
> - **`script.container.ui`**: Required to display the ProWebLook Lead Verifier sidebar inside Google Sheets where users connect their ProWebLook API key, view credit balances, and run bulk verification.
> - **`script.external_request`**: Required exclusively to call `https://proweblook.com/api/v1/zapier/*` (enforced via `urlFetchWhitelist`) to verify emails, validate phone numbers, check WhatsApp registration, and fetch credit balances.

---

## 3. Enable & Configure Google Workspace Marketplace SDK

1. In Google Cloud Console → **APIs & Services → Library** → enable **Google Workspace Marketplace SDK**.
2. In Apps Script → **Deploy → New deployment** → select **Add-on** → copy the **Deployment Version Number**.
3. Open **Google Workspace Marketplace SDK → App Configuration**:
   - **App Visibility:** `Public`
   - **Installation Settings:** Individual + Admin install
   - **App Integration:** Check **Sheets Add-on** → enter your Apps Script **Script ID** and **Deployment Version Number**.
4. Open **Store Listing**:
   - **Application Name:** `ProWebLook — Email, Phone & WhatsApp Verifier`
   - **Short Description (max 80 chars):** `Verify emails, phone carriers & WhatsApp numbers directly in Google Sheets.`
   - **Category:** `Marketing & Analytics` / `Sales & CRM`
   - **32x32 Icon:** `assets/logo-32.png`
   - **128x128 Icon:** `assets/logo-128.png`
   - **Support URL:** `https://proweblook.com`
5. Click **Publish**!
