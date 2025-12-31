# GLC Dashboard Backend

FastAPI backend for The Great Lawn Co. TV Delivery Dashboard.

## Quick Start

1. Create virtual environment:
```bash
cd backend
python -m venv venv
```

2. Activate virtual environment:
```bash
# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Copy environment file:
```bash
cp .env.example .env
```

5. Run the server:
```bash
uvicorn app.main:app --reload
```

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | API information |
| `/health` | GET | Health check |
| `/api/v1/schedule` | GET | Get 10-day delivery schedule |
| `/docs` | GET | OpenAPI documentation (Swagger UI) |
| `/redoc` | GET | ReDoc documentation |

## Project Structure

```
backend/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI application
│   ├── config.py            # Configuration management
│   ├── api/
│   │   └── v1/
│   │       └── routes/
│   │           ├── health.py    # Health endpoint
│   │           └── schedule.py  # Schedule endpoint
│   ├── models/
│   │   └── schedule.py      # Pydantic models
│   └── services/
│       └── mock_data.py     # Mock data generator
├── requirements.txt
├── .env.example
└── README.md
```

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `API_HOST` | Server host | `0.0.0.0` |
| `API_PORT` | Server port | `8000` |
| `DEBUG` | Debug mode | `true` |
| `CORS_ORIGINS` | Allowed origins | `*` |
