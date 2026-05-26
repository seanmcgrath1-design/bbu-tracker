# Bulk EO Creation — Playwright Automation

Automates cloning and releasing Engineering Orders (EOs) in the Verizon order management app. Each script clones a specific source EO with new SPM Project IDs, fills in the target date, checks "Include GC and Shipping", selects DC-GC as the ship method, submits the clone form, and then releases each new EO via Edit → Review → Submit.

---

## First-Time Setup (New Laptop)

### 1. Install Node.js

Download and install from [https://nodejs.org/](https://nodejs.org/) (LTS version). Verify with:

```
node --version
```

### 2. Clone or pull this repository

If you're on a new machine:

```
git clone <repo-url>
cd bbu-tracker
```

If you already have the repo:

```
git pull
```

### 3. Install dependencies

From the `bbu-tracker` folder (not the `Bulk EO Creation` subfolder):

```
npm install
npx playwright install chromium
```

That's it — no additional config needed.

---

## Running a Script

From the `bbu-tracker` folder in VS Code's integrated terminal:

```
node "Bulk EO Creation/run-prolabs-v1-sfp.js"
```

The script will:
1. Prompt for your Verizon password (hidden input)
2. Prompt for Project IDs — paste 1 to 20 IDs, then press Enter twice

**Accepted input formats:**
- One per line
- Comma or space separated on one line
- One continuous string (IDs will be auto-split at every 8 characters)

After the IDs are confirmed, Playwright opens a Chrome window and runs the full automation. The browser stays open at the end so you can review the created EOs. Click **Resume** in the Playwright Inspector to close.

---

## Available Scripts

| Runner | Spec | Source EO | Part |
|--------|------|-----------|------|
| `run-prolabs-v1-sfp.js` | `prolabs-v1-sfp.spec.ts` | E000270779 | ProLabs V1 SFP |
| `run-DWDM.js` | `DWDM.spec.ts` | *(set in spec)* | DWDM |
| `run-Ericsson V1 SFP.js` | `Ericsson V1 SFP.spec.ts` | *(set in spec)* | Ericsson V1 SFP |

---

## Creating a Script for a New Part

Each part needs two files: a **runner** (`.js`) and a **spec** (`.spec.ts`).

### Step 1 — Copy the spec

Copy `prolabs-v1-sfp.spec.ts` and rename it (e.g., `my-new-part.spec.ts`).

Open the new spec and change line 4:

```typescript
const SOURCE_EO = 'E000270779';  // ← replace with the source EO to clone from
```

Also update the target date on lines 7–8 if needed:

```typescript
const TARGET_DATE = '07/17/2026';
const TARGET_DATE_ID = '20260717';
```

### Step 2 — Copy the runner

Copy `run-prolabs-v1-sfp.js` and rename it to match (e.g., `run-my-new-part.js`).

Open the new runner and find line 52. Change `prolabs-v1-sfp.spec.ts` to your new spec filename:

```javascript
const child = spawn('npx', ['playwright', 'test', 'my-new-part.spec.ts', '--headed', '--project=chromium'], {
```

### Step 3 — Run it

```
node "Bulk EO Creation/run-my-new-part.js"
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `Project "chromium" not found` | Run `npx playwright install chromium` from the `bbu-tracker` folder |
| `npm install` errors | Make sure you're in `bbu-tracker`, not inside `Bulk EO Creation` |
| Login page not loading | VPN or network issue; verify you can reach the Verizon app in a browser first |
| Clone window closes before IDs are filled | Usually a page load timing issue; try again — `networkidle` wait handles most cases |
| Submit button stays disabled | The EO may have a validation error; check the open EO manually in the browser |

---

## How the Automation Works (Overview)

1. Navigates to the Verizon order management dashboard and logs in
2. Searches for the source EO and opens it in a new window
3. Clicks the **Clone** button (in the toolbar or the `...` overflow menu)
4. Adds rows for each Project ID, fills the date, checks "Include GC and Shipping", selects DC-GC
5. Clicks **Clone** to submit the form and create the new drafted EOs
6. Opens each new EO link and releases it via **Edit → Review → Submit**
7. Pauses so you can review — click Resume in the Inspector when done

---

## Changelog

| Date | Change |
|------|--------|
| 2026-05-23 | Initial working script for ProLabs V1 SFP (E000270779) |
| 2026-05-23 | Added DWDM and Ericsson V1 SFP scripts |
| 2026-05-26 | Updated runner to accept 1–20 variable Project IDs |
| 2026-05-26 | Added this README |
