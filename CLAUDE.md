# CLAUDE.md — AI Assistant Guide for Dashboard_contact_log

This file provides context, conventions, and workflows for AI assistants (Claude and others) working in this repository. Keep this file up to date as the project evolves.

---

## Project Overview

**Repository:** `Dashboard_contact_log`
**Status:** New / bootstrapping phase — no application code yet committed.

This project is a **contact-log dashboard** application. As development progresses, update this section with:
- A one-sentence description of what the app does
- The primary audience / users
- Any external services or integrations it depends on

---

## Repository State

> **Note:** This repository was initialized empty. The sections below are scaffolded templates that should be filled in as the project grows. When adding the first meaningful code, update this file in the same commit.

---

## Technology Stack

*To be filled in once the stack is chosen. Common starting points for a dashboard + contact log:*

| Layer | Likely choice | Notes |
|-------|---------------|-------|
| Frontend | React / Vue / plain HTML | Update when decided |
| Backend | Node.js / Python / Go | Update when decided |
| Database | PostgreSQL / SQLite / MongoDB | Update when decided |
| Styling | Tailwind CSS / Bootstrap | Update when decided |
| Testing | Jest / Pytest / Vitest | Update when decided |
| CI/CD | GitHub Actions / GitLab CI | Update when decided |

When the stack is finalized, replace this table with concrete details and version requirements.

---

## Directory Structure

*Update this section as the project layout is established.*

```
Dashboard_contact_log/
├── CLAUDE.md          # This file — AI assistant guide
├── README.md          # Human-facing project documentation
├── .gitignore         # Files excluded from version control
├── .env.example       # Template for required environment variables
│
├── src/               # Application source code (to be created)
│   ├── ...
│
├── tests/             # Test files (to be created)
│   ├── ...
│
└── docs/              # Additional documentation (optional)
    └── ...
```

---

## Development Workflows

### First-Time Setup

Once project files exist, document the setup steps here. For example:

```bash
# Clone the repository
git clone <repo-url>
cd Dashboard_contact_log

# Install dependencies (adjust for your stack)
npm install          # Node.js
# or
pip install -r requirements.txt   # Python

# Copy environment template and fill in values
cp .env.example .env

# Run the development server
npm run dev
# or
python manage.py runserver
```

### Running Tests

```bash
# Run the full test suite
npm test             # Node.js
# or
pytest               # Python

# Run tests in watch mode
npm run test:watch
```

### Building for Production

```bash
npm run build
# or
python -m build
```

### Linting and Formatting

```bash
# Lint
npm run lint         # Node.js (ESLint / Biome)
# or
flake8 .             # Python

# Format
npm run format       # Prettier
# or
black .              # Python
```

Always run linting and tests before committing.

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

*Finalize these once the language/framework is chosen. Placeholders below.*

### General

- Prefer readability over cleverness
- Keep functions small and single-purpose
- Delete dead code rather than commenting it out
- No secrets or credentials in source code — use environment variables via `.env` (excluded from git)

### JavaScript / TypeScript (if applicable)

- ES modules (`import`/`export`), no CommonJS `require` in new files
- Prefer `const`; use `let` only when reassignment is necessary
- Async/await over raw Promises
- Strict TypeScript mode (`"strict": true` in `tsconfig.json`)
- File names: `kebab-case.ts` for modules, `PascalCase.tsx` for React components

### Python (if applicable)

- Python 3.10+ features are acceptable
- Type hints on all public function signatures
- Follow PEP 8; enforced by `black` + `flake8`
- Use `pathlib.Path` instead of `os.path`

---

## Environment Variables

*Document all required variables here. Never commit real values.*

| Variable | Description | Example |
|----------|-------------|---------|
| `DATABASE_URL` | Database connection string | `postgresql://user:pass@localhost/db` |
| `SECRET_KEY` | App secret / JWT signing key | (generate a random 64-char string) |
| `LOG_LEVEL` | Logging verbosity | `info` |

Add variables to `.env.example` with placeholder values whenever a new one is introduced.

---

## Testing Guidelines

- Write tests alongside new features — not as an afterthought
- Unit tests for pure logic; integration tests for DB/API interactions
- Name test files `<module>.test.ts` (JS) or `test_<module>.py` (Python)
- Tests must pass on the CI pipeline before merging
- Aim for meaningful coverage of edge cases, not just happy paths

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

*Populate this section as the project matures.*

| Task | Command |
|------|---------|
| Start dev server | `TBD` |
| Run tests | `TBD` |
| Lint code | `TBD` |
| Build for production | `TBD` |
| Apply DB migrations | `TBD` |
| Seed test data | `TBD` |

---

## Glossary

*Add domain-specific terms here as they appear in the codebase.*

| Term | Definition |
|------|------------|
| Contact log | A record of communications or interactions associated with a contact |
| Dashboard | The main UI view aggregating contact log data |

---

*Last updated: 2026-02-23. Update this date whenever this file is meaningfully changed.*
