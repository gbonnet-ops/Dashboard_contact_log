# CLAUDE.md — AI Assistant Guide for Dashboard_contact_log

This file provides context, conventions, and workflows for AI assistants (Claude and others) working in this repository. Keep this file up to date as the project evolves.

---

## Project Overview

**Repository:** `Dashboard_contact_log`
**Status:** Active — core automation script delivered.

This project automates the processing of investment bankers' deal-flow call notes.
Bankers enter free-text notes into a Google Sheet (`Raw_Logs`). A Google Apps Script triggers hourly, sends each unprocessed row to the **OpenAI API (`gpt-4o`)**, and writes structured deal-pipeline data into a second sheet (`Clean_Data`). That clean sheet feeds a **Looker Studio** dashboard.

- **Primary users:** Investment bankers and deal-flow analysts
- **External services:** OpenAI API (`gpt-4o`), Google Sheets, Looker Studio

---

## Technology Stack

| Layer | Technology | Notes |
|-------|-----------|-------|
| Automation engine | Google Apps Script (V8 runtime) | No server required — native to Google Workspace |
| LLM / AI | OpenAI API (`gpt-4o`) | Structured JSON extraction from free-text notes |
| Source of truth | Google Sheets — `Raw_Logs` tab | Manual banker input |
| Structured output | Google Sheets — `Clean_Data` tab | AI-generated, feeds Looker Studio |
| Dashboard | Looker Studio | Connected to `Clean_Data` |
| Secret management | Google `PropertiesService` | API key never in source code |

---

## Directory Structure

```
Dashboard_contact_log/
├── CLAUDE.md                        # This file — AI assistant guide
├── README.md                        # Full setup, deployment & Looker Studio guide
│
└── src/
    └── ContactLogProcessor.gs       # Complete Google Apps Script
                                     # Copy-paste into Apps Script editor to deploy
```

### Key file: `src/ContactLogProcessor.gs`

All logic lives in one `.gs` file for easy deployment into the Apps Script editor.
Internal organization:

| Section | Functions |
|---------|-----------|
| **Main entry point** | `processContactLogs()` |
| **OpenAI API layer** | `_callOpenAIAPI()` |
| **Response validation** | `_validateAndParse()` |
| **Sheet operations** | `_appendToCleanData()`, `_getSheet()` |
| **Secure config** | `_getApiKey()`, `setApiKey()` |
| **One-time setup** | `initializeSheets()`, `createTimeTrigger()`, `retryErrorRows()` |

---

## Development Workflows

> This project runs entirely inside Google Apps Script — there is no local dev server,
> no `npm install`, and no build step. Development = editing `.gs` files in the browser
> editor and testing against a real Google Sheet.

### First-Time Deployment (one-time)

```
1. Open target Google Sheet
2. Extensions → Apps Script
3. Paste contents of src/ContactLogProcessor.gs into Code.gs
4. Run setApiKey()        → stores API key in PropertiesService (then clear the key from code)
5. Run initializeSheets() → creates Raw_Logs + Clean_Data tabs with headers
6. Run createTimeTrigger() → activates hourly auto-processing
```

Full walkthrough: see `README.md`.

### Manual Processing

Run `processContactLogs()` from the Apps Script editor at any time.
Processes all rows where Raw_Logs column E is empty.

### Re-processing Error Rows

Run `retryErrorRows()` to reset all `Erreur`-tagged rows back to empty,
allowing them to be picked up on the next run.

### Editing the Script

1. Edit `src/ContactLogProcessor.gs` in this repository (for version control).
2. Copy the updated content into the Apps Script editor.
3. Test manually by running `processContactLogs()`.
4. Commit changes to this repo with a descriptive message.

---

## Git Conventions

### Branch Naming

| Type | Pattern | Example |
|------|---------|---------|
| Feature | `feature/<short-description>` | `feature/contact-export-csv` |
| Bug fix | `fix/<short-description>` | `fix/duplicate-contact-entries` |
| Chore | `chore/<short-description>` | `chore/update-dependencies` |
| Claude AI work | `claude/<session-id>` | `claude/claude-md-mlzbsn0bwmxc8eez-92GC7` |

### Commit Messages

Follow **Conventional Commits** (`<type>(<scope>): <summary>`):

```
feat(dashboard): add contact search filter
fix(api): handle null phone numbers in log entries
chore(deps): upgrade lodash to 4.17.21
docs(claude): update CLAUDE.md with stack details
test(contacts): add unit tests for deduplication logic
```

Types: `feat`, `fix`, `refactor`, `test`, `docs`, `chore`, `style`, `perf`, `ci`.

- Subject line: imperative mood, ≤ 72 characters, no trailing period
- Body: explain *why*, not just *what*; wrap at 72 chars
- Never force-push to `main` or `master`

### Pull Requests

- Target branch: `main` (or `develop` if a GitFlow model is adopted)
- PR title mirrors the conventional commit format
- Include a short description of what changed and why
- Link related issues with `Closes #<issue-number>`

---

## Code Conventions

### Google Apps Script (V8)

- **Language:** JavaScript ES2019 (V8 runtime) — no TypeScript, no bundler
- **Naming:** `camelCase` for variables and public functions; `_camelCase` (underscore prefix) for private/internal helpers
- **Constants:** `UPPER_SNAKE_CASE` at the top of the file
- Prefer `const`; use `let` only when reassignment is necessary
- No `var` — always `const` or `let`
- Always use `===` (strict equality)
- Guard clauses over deeply nested `if` blocks
- Each function does one thing — keep them short and focused

### API & Data Safety

- **Never** hardcode the API key — always use `PropertiesService`
- **Never** send the full sheet to the API — one row per call
- Validate all LLM output before writing to the sheet (see `_validateAndParse()`)
- Strip accidental markdown fences from LLM responses before `JSON.parse()`

### Nomenclature (domain)

Stick to the exact French/English labels used in the sheet headers and the Claude prompt — do not invent new status names or change capitalization.

| Canonical value | Context |
|----------------|---------|
| `Approché` | First stage of the deal pipeline |
| `Teaser Envoyé` | Teaser document sent |
| `NDA Signé` | Non-disclosure agreement signed |
| `Management Pres` | Management presentation stage |
| `En Due Diligence` | Active due diligence |
| `Offre Soumise` | Offer submitted |
| `Pass/Dropped` | Deal abandoned |
| `Traité` | Row successfully processed (Raw_Logs col E) |
| `Erreur` | Processing failed (Raw_Logs col E) |

---

## Environment Variables

There are no `.env` files — this project uses Google Apps Script `PropertiesService`.

| Property key | Description | How to set |
|-------------|-------------|-----------|
| `OPENAI_API_KEY` | OpenAI API key (starts with `sk-...`) | Run `setApiKey()` once in Apps Script editor |

**Never** commit an actual key to this repository. The `setApiKey()` function in the script is a one-time bootstrap helper — clear the key string from the code immediately after running it. OpenAI keys start with `sk-` and can be generated at platform.openai.com.

---

## Testing Guidelines

Google Apps Script has no native unit-test framework. Testing is done manually:

1. **Happy path:** Add a raw note row to `Raw_Logs` with column E empty, run `processContactLogs()`, verify the row appears in `Clean_Data` with correct values and column E shows `Traité`.
2. **Error path:** Temporarily set an invalid API key, run, verify column E shows `Erreur` and no data is written to `Clean_Data`.
3. **Validation:** Feed edge-case notes (ambiguous status, no date, drop reason) and verify `_validateAndParse()` handles them gracefully.
4. **Re-run safety:** Run `processContactLogs()` twice — already-processed rows must not be duplicated.

---

## AI Assistant Instructions (Claude-Specific)

### Before Making Changes

1. Read the files you intend to modify — never edit code you haven't seen
2. Check this `CLAUDE.md` for project conventions before writing any code
3. Use `TodoWrite` to plan multi-step tasks before starting
4. If requirements are ambiguous, ask a clarifying question before proceeding

### What to Do

- Follow the commit and branch conventions in this file exactly
- Keep solutions minimal — solve the stated problem, nothing more
- Update `CLAUDE.md` in the same commit whenever you add or change a convention
- Mark todos as completed immediately after finishing each step

### What Not to Do

- Do not add unrequested features, comments, docstrings, or refactors
- Do not commit `.env` files, secrets, or large binary files
- Do not push to `main`/`master` directly — always use a feature branch
- Do not guess at missing parameter values; ask the user instead
- Do not introduce security vulnerabilities (SQLi, XSS, command injection, etc.)

### Commit Workflow for Claude

```bash
# Work on the designated branch
git checkout claude/<session-id>

# Stage specific files (never `git add -A` blindly)
git add path/to/changed-file.ts

# Commit with conventional message
git commit -m "feat(contacts): add CSV export endpoint"

# Push
git push -u origin claude/<session-id>
```

---

## Common Tasks (Quick Reference)

All tasks are run inside the **Apps Script editor** (Extensions → Apps Script):

| Task | Function to run |
|------|----------------|
| First-time sheet setup | `initializeSheets()` |
| Store API key securely | `setApiKey()` (then clear the key from code) |
| Activate hourly trigger | `createTimeTrigger()` |
| Process new rows manually | `processContactLogs()` |
| Reset error rows for retry | `retryErrorRows()` |
| View execution logs | Executions tab in Apps Script editor |

---

## Glossary

| Term | Definition |
|------|------------|
| Contact Log | A banker's free-text note from an investor call, entered in `Raw_Logs` |
| Raw_Logs | Google Sheet tab where bankers manually enter call notes |
| Clean_Data | Google Sheet tab populated by the script with AI-structured deal data |
| Deal Status | Current stage of a deal in the M&A pipeline (7 allowed values) |
| Drop Reason | Category explaining why a deal was abandoned (5 allowed values) |
| Traité | French for "Processed" — status written to Raw_Logs col E on success |
| Erreur | French for "Error" — status written to Raw_Logs col E on API failure |
| PropertiesService | Google Apps Script service for secure key/value storage (encrypted at rest) |
| Time-based trigger | Apps Script scheduler that runs `processContactLogs()` every hour |

---

*Last updated: 2026-02-23. Corrected stale Anthropic/Claude API references after migration to OpenAI gpt-4o (commit f25f11f). Update this date whenever this file is meaningfully changed.*
