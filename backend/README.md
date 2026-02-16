# FastAPI Backend for KPMG Audit Tool

## Setup

### 1. Prerequisites
- Python 3.10+
- PostgreSQL 14+

### 2. Create virtual environment
```bash
cd backend
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Set up PostgreSQL
Create a database and update the `.env` file:
```
DATABASE_URL=postgresql://user:password@localhost:5432/kpmg_audit
```

### 5. Run migrations
```bash
alembic upgrade head
```

### 6. Start the server
```bash
uvicorn app.main:app --reload --port 8000
```

The API will be available at `http://localhost:8000`.
API docs at `http://localhost:8000/docs`.

## Project Structure
```
backend/
├── app/
│   ├── main.py          # FastAPI app entry point
│   ├── database.py      # DB connection & session
│   ├── models.py        # SQLAlchemy models
│   ├── schemas.py       # Pydantic schemas
│   └── routers/
│       ├── files.py     # File upload endpoints
│       ├── audit.py     # Audit endpoints
│       └── reconciliation.py  # Reconciliation endpoints
├── alembic/             # DB migrations
├── requirements.txt
├── .env.example
└── README.md
```

## Frontend Connection
The React frontend expects the API at `http://localhost:8000` by default.
Override with the `VITE_API_BASE_URL` environment variable if needed.
