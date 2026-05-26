# BBU Tracker — Google Apps Script Project

This repository contains the Google Apps Script source files for the BBU Tracker, managed with [clasp](https://github.com/google/clasp).

---

## Collaborator Setup

### Prerequisites

- [Node.js](https://nodejs.org/) installed
- A Google account with access to the shared Apps Script project

### 1. Install clasp

```bash
npm install -g @google/clasp
```

### 2. Authenticate with Google

```bash
clasp login
```

This opens a browser window — sign in with the Google account that has access to the Apps Script project.

### 3. Clone the repository

```bash
git clone <repo-url>
cd bbu-tracker
```

### 4. Configure clasp

Copy the template to create your local clasp config:

```bash
# Windows
copy .clasp.json.template .clasp.json

# Mac/Linux
cp .clasp.json.template .clasp.json
```

> `.clasp.json` is gitignored so your local path never gets committed.

### 5. Verify the connection

```bash
clasp status
```

You should see the list of local `.gs` files without errors.

---

## Daily Workflow

### Pull latest changes from GitHub, then sync to Apps Script

```bash
git pull
clasp push
```

### Pull from Apps Script to local (if changes were made in the browser editor)

```bash
clasp pull
git add -A
git commit -m "sync: pull changes from Apps Script editor"
git push
```

### Push local changes to Apps Script

```bash
git add -A
git commit -m "feat: description of what changed and why"
git push
clasp push
```

---

## File Overview

| File | Purpose |
|------|---------|
| `Fuze Dump.gs` | Fuze data import logic |
| `Tech Assignment.gs` | Technician assignment handling |
| `CQ status.gs` | CQ status tracking |
| `Master Shading.gs` | Map shading logic |
| `Map_Link.gs` | Map link generation |
| `Missing Coords Reminder.gs` | Reminder for missing coordinates |
| `Proximity_Logic.gs` | Proximity detection logic |
| `Handoff.gs` | Handoff workflow |
| `Live Map.gs` | Live map display |
| `appsscript.json` | Apps Script manifest (OAuth scopes, runtime) |

---

## Commit Message Format

Use prefixes so the changelog and git history stay readable:

| Prefix | Use for |
|--------|---------|
| `feat:` | New feature or behavior |
| `fix:` | Bug fix |
| `sync:` | Syncing between Apps Script editor and GitHub |
| `refactor:` | Code cleanup with no behavior change |
| `docs:` | README, comments, or CHANGELOG updates |

Example:
```
feat: add proximity alert threshold to Proximity_Logic.gs

Increased default radius from 0.5 to 1.0 miles based on dispatch feedback.
Discussed in Claude session 2026-05-21.
```

---

## Bulk EO Creation (Playwright Automation)

The `Bulk EO Creation/` folder contains Playwright scripts that automate cloning and releasing Engineering Orders in the Verizon order management app.

See [Bulk EO Creation/README.md](Bulk%20EO%20Creation/README.md) for setup instructions, how to run scripts, and how to create new ones for different parts.

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for a record of changes by session.
