# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

KPMG Audit Assist Suite — a two-part application for audit automation and data reconciliation:
- **Frontend**: React + TypeScript + Vite + shadcn/ui + Tailwind CSS
- **Backend**: FastAPI (Python) + PostgreSQL + SQLAlchemy

## Commands

### Frontend

```bash
# Install dependencies
npm install

# Start dev server (http://localhost:5173)
npm run dev

# Production build
npm run build

# Preview production build locally
npm run preview

# Lint
npm run lint
```

### Backend

```bash
cd backend

# Create and activate virtual environment
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Copy env file and configure
cp .env.example .env

# Run database migrations
alembic upgrade head

# Start server (http://localhost:8000, docs at /docs)
uvicorn app.main:app --reload --port 8000
```

### Database (one-time setup)

```bash
# Apply raw SQL schema (alternative to alembic)
psql -U postgres -d kpmg_audit -f backend/database_schema.sql
```

## Architecture

### Frontend (`src/`)

**Routing** (`src/App.tsx`): Three routes — `/login`, `/onboarding`, and `/` (main app). All authenticated views live under `/`.

**Main layout** (`src/pages/Index.tsx`): Single-page shell with `TopNavbar`, `Sidebar`, and `StatusBar`. The sidebar drives `activeView` state which swaps between:
- `dashboard` → `AuditPanel` (PPT/report automation)
- `reconciliation` → `ReconciliationPanel` (Excel file comparison)
- `templates` → `TemplateRequestPanel`
- `userdetails` → `UserDetailsSection`
- `upload`, `output` → stubs (not yet implemented)

An FS/Non-FS toggle lives in the content area header (not `TopNavbar`). When Non-FS is active, `reconciliation` is hidden and `activeView` resets to `dashboard`.

**API client** (`src/lib/api.ts`): Single `api` object with typed functions for all backend endpoints. Uses `VITE_API_BASE_URL` env var (defaults to `http://localhost:8000`). All file uploads use `FormData`; all other requests use JSON.

**Components**: Organized by feature under `src/components/`:
- `audit/` — file upload, audit type/template selectors, progress display
- `reconciliation/` — two-file slot upload, results display
- `templates/` — template request form
- `layout/` — TopNavbar, Sidebar, StatusBar, UserDetailsSection
- `shared/` — cross-feature components (e.g. `ProcessingDialog`)
- `ui/` — shadcn/ui primitives (do not modify these)

**Import alias**: `@/` maps to `src/` (configured in `tsconfig.json` and `vite.config.ts`).

**Key libraries**: `@tanstack/react-query` for server state; `react-hook-form` + `zod` for form validation; `sonner` for toast notifications; `recharts` for charts.

### Backend (`backend/app/`)

**Entry point** (`main.py`): Registers all routers and sets CORS to allow `localhost:8080`, `5173`, and `3000`.

**Routers** (`app/routers/`):
| File | Prefix | Purpose |
|------|--------|---------|
| `files.py` | `/api/files` | Upload/delete files, stored on disk |
| `ppt_automation.py` | `/api/ppt` | Start/status/cancel audit runs (background stub) |
| `reconciliation.py` | `/api/reconciliation` | Start/status/export reconciliation jobs |
| `auth.py` | `/api/auth` | Register/login using raw SQL (not ORM) |
| `user_onboarding.py` | `/api/users` | Save/retrieve user profiles |
| `template_req.py` | `/api/templates` | Submit template requests with file upload |
| `error_logging.py` | `/api/errors` | Log and retrieve application errors |
| `user_logs.py` | `/api/logs` | Track user activity per module/run |

**Database models** (`models.py`): 8 SQLAlchemy tables — `users`, `user_profiles`, `uploaded_files`, `audit_runs`, `reconciliation_runs`, `template_requests`, `error_logs`, `user_logs`. All primary keys are UUIDs.

**Schemas** (`schemas.py`): Pydantic v2 models mirroring each table with `Create` and `Response` variants.

**Database** (`database.py`): Standard SQLAlchemy session factory. `get_db()` is a FastAPI dependency injected into router handlers. Connection URL from `DATABASE_URL` env var.

`Base.metadata.create_all(bind=engine)` in `main.py` auto-creates all tables on startup — Alembic migrations are an alternative, not a prerequisite.

**Note**: `auth.py` intentionally uses raw SQL (`db.execute(text(...))`) instead of the ORM for user queries.

**No test suite** currently exists for frontend or backend.

### Environment Variables

**Backend** (`.env`):
```
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/kpmg_audit
UPLOAD_DIR=./uploads
```

**Frontend** (`.env.local`, optional):
```
VITE_API_BASE_URL=http://localhost:8000
```
