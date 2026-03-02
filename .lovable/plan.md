

# FastAPI Backend Restructuring Plan

## Overview

Restructure the existing FastAPI backend into 6 clean modules, add 4 new database tables, wire up authentication with direct SQL queries, and update the frontend API client to match.

---

## New Database Tables (in `models.py`)

1. **template_requests** -- stores template request form submissions (file_name, file_path, spoc_email, team_wide, user_id, created_at)
2. **user_profiles** -- stores onboarding form data (user_id FK, team_name, partner_name, client_name, sector, employee_id)
3. **error_logs** -- captures application errors (id, user_id, module, error_message, stack_trace, created_at)
4. **user_logs** -- captures user activity and run history (id, user_id, action, module, run_id, details, created_at)

Existing tables (uploaded_files, audit_runs, reconciliation_runs, users) remain unchanged.

---

## Backend Router Files

### 1. `routers/ppt_automation.py` (replaces current `audit.py`)
- **POST `/api/ppt/start`** -- Receives uploaded file IDs, template path, and output folder path. Creates an `audit_runs` record, launches a background stub task that simulates the 6 processing steps, and logs the run in `user_logs`.
- **GET `/api/ppt/{run_id}/status`** -- Returns current progress and step statuses.
- **POST `/api/ppt/{run_id}/cancel`** -- Cancels a running job.

### 2. `routers/reconciliation.py` (refactored)
- Keep existing endpoints but add `user_logs` tracking on start/complete.
- Stub logic remains (simulated matched/mismatched/missing counts).

### 3. `routers/template_req.py` (new)
- **POST `/api/templates/request`** -- Accepts file upload + spoc_email + team_wide flag. Saves the file, creates a `template_requests` DB record.
- **GET `/api/templates/requests`** -- Lists all template requests (for admin/history view).

### 4. `routers/user_onboarding.py` (new)
- **POST `/api/users/profile`** -- Creates/updates the user profile (team_name, partner_name, client_name, sector, employee_id) in `user_profiles` table.
- **GET `/api/users/profile/{user_id}`** -- Returns the stored profile.

### 5. `routers/error_logging.py` (new)
- **POST `/api/errors/log`** -- Logs an error to the `error_logs` table (module, error_message, stack_trace, user_id).
- **GET `/api/errors`** -- Lists recent errors (for debugging/admin).
- Also provides a utility function `log_error(db, user_id, module, message, trace)` that other routers can import and call in their except blocks.

### 6. `routers/user_logs.py` (new)
- **POST `/api/logs`** -- Records a user activity entry (action, module, run_id, details).
- **GET `/api/logs/{user_id}`** -- Returns the activity history for a specific user.
- Also provides a utility function `log_activity(db, user_id, action, module, run_id, details)` for other routers to call.

### 7. `routers/auth.py` (updated)
- Replace ORM-based queries with raw SQL using `db.execute(text(...))` for user authentication as requested.
- Keep register and login endpoints; use raw SQL INSERT/SELECT instead of `db.query(User)`.

---

## Schema Updates (`schemas.py`)

Add Pydantic models for:
- `TemplateRequestCreate`, `TemplateRequestResponse`
- `UserProfileCreate`, `UserProfileResponse`
- `ErrorLogCreate`, `ErrorLogResponse`
- `UserLogCreate`, `UserLogResponse`
- `PPTAutomationRequest`, `PPTAutomationResponse`

---

## Frontend Changes

### `src/lib/api.ts`
Add new API functions:
- `api.submitTemplateRequest(file, spocEmail, teamWide)` -- calls POST `/api/templates/request`
- `api.saveUserProfile(profile)` -- calls POST `/api/users/profile`
- `api.getUserProfile(userId)` -- calls GET `/api/users/profile/{userId}`
- `api.register(name, email, password)` -- calls POST `/api/auth/register`
- `api.login(email, password)` -- calls POST `/api/auth/login`
- `api.logError(error)` -- calls POST `/api/errors/log`
- `api.getUserLogs(userId)` -- calls GET `/api/logs/{userId}`
- Update `startAudit` to point to `/api/ppt/start`

### `src/pages/Login.tsx`
- Wire the sign-in form to call `api.login()` instead of the mock timeout.
- Wire the register form to call `api.register()` instead of the mock timeout.

### `src/pages/UserOnboarding.tsx`
- On submit, call `api.saveUserProfile()` to persist data in PostgreSQL instead of just localStorage.

### `src/components/templates/TemplateRequestPanel.tsx`
- On submit, call `api.submitTemplateRequest()` to save to the DB.

---

## `main.py` Updates

- Register the new routers: `ppt_automation`, `template_req`, `user_onboarding`, `error_logging`, `user_logs`.
- Remove the old `audit` router import, replace with `ppt_automation`.

---

## File Structure After Changes

```text
backend/app/
  __init__.py
  database.py
  main.py
  models.py          (4 new tables added)
  schemas.py          (new Pydantic models)
  routers/
    __init__.py
    auth.py           (raw SQL queries)
    ppt_automation.py (new, replaces audit.py)
    reconciliation.py (updated with logging)
    template_req.py   (new)
    user_onboarding.py(new)
    error_logging.py  (new)
    user_logs.py      (new)
    files.py          (unchanged)
```

Frontend files modified:
- `src/lib/api.ts`
- `src/pages/Login.tsx`
- `src/pages/UserOnboarding.tsx`
- `src/components/templates/TemplateRequestPanel.tsx`

Old file removed:
- `backend/app/routers/audit.py` (replaced by `ppt_automation.py`)
