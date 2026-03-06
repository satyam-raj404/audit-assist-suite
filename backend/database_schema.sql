-- ============================================================
-- KPMG Audit Tool — PostgreSQL Schema
-- Import: psql -U postgres -d kpmg_audit -f database_schema.sql
-- ============================================================

CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 1. Users
CREATE TABLE IF NOT EXISTS users (
    id          UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name        VARCHAR(255) NOT NULL,
    email       VARCHAR(255) UNIQUE NOT NULL,
    hashed_password TEXT NOT NULL,
    created_at  TIMESTAMP DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);

-- 2. User Profiles (onboarding)
CREATE TABLE IF NOT EXISTS user_profiles (
    id            UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id       UUID UNIQUE NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    team_name     VARCHAR(255) NOT NULL,
    partner_name  VARCHAR(255) NOT NULL,
    client_name   VARCHAR(255) NOT NULL,
    sector        VARCHAR(255) NOT NULL,
    employee_id   VARCHAR(100) NOT NULL,
    created_at    TIMESTAMP DEFAULT NOW(),
    updated_at    TIMESTAMP DEFAULT NOW()
);

-- 3. Uploaded Files
CREATE TABLE IF NOT EXISTS uploaded_files (
    id          UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name        VARCHAR(255) NOT NULL,
    size        INTEGER NOT NULL,
    type        VARCHAR(100) NOT NULL,
    path        TEXT NOT NULL,
    created_at  TIMESTAMP DEFAULT NOW()
);

-- 4. Audit Runs (PPT Automation)
CREATE TABLE IF NOT EXISTS audit_runs (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    audit_type      VARCHAR(100) NOT NULL,
    utility_type    VARCHAR(50),
    report_type     VARCHAR(20),
    excel_file_id   UUID,
    pptx_file_id    UUID,
    month           VARCHAR(20),
    year            VARCHAR(10),
    output_path     TEXT NOT NULL,
    status          VARCHAR(20) DEFAULT 'pending',
    progress        FLOAT DEFAULT 0,
    report_path     TEXT,
    user_id         UUID,
    created_at      TIMESTAMP DEFAULT NOW(),
    updated_at      TIMESTAMP DEFAULT NOW()
);

-- 5. Reconciliation Runs
CREATE TABLE IF NOT EXISTS reconciliation_runs (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    audit_type      VARCHAR(100),
    source_file_id  UUID NOT NULL,
    target_file_id  UUID,
    output_path     TEXT,
    status          VARCHAR(20) DEFAULT 'pending',
    progress        FLOAT DEFAULT 0,
    matched         INTEGER DEFAULT 0,
    mismatched      INTEGER DEFAULT 0,
    missing         INTEGER DEFAULT 0,
    report_path     TEXT,
    user_id         UUID,
    created_at      TIMESTAMP DEFAULT NOW(),
    updated_at      TIMESTAMP DEFAULT NOW()
);

-- 6. Template Requests
CREATE TABLE IF NOT EXISTS template_requests (
    id          UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    file_name   VARCHAR(255) NOT NULL,
    file_path   TEXT NOT NULL,
    spoc_email  VARCHAR(255) NOT NULL,
    team_wide   BOOLEAN DEFAULT FALSE,
    user_id     UUID,
    created_at  TIMESTAMP DEFAULT NOW()
);

-- 7. Error Logs
CREATE TABLE IF NOT EXISTS error_logs (
    id            UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id       UUID,
    module        VARCHAR(100) NOT NULL,
    error_message TEXT NOT NULL,
    stack_trace   TEXT,
    created_at    TIMESTAMP DEFAULT NOW()
);

-- 8. User Logs (activity / history)
CREATE TABLE IF NOT EXISTS user_logs (
    id          UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id     UUID,
    action      VARCHAR(100) NOT NULL,
    module      VARCHAR(100) NOT NULL,
    run_id      UUID,
    details     TEXT,
    created_at  TIMESTAMP DEFAULT NOW()
);
