# CV Bot — Telegram Resume Parser

A production-ready Telegram bot that accepts `.docx` resume files, extracts
structured candidate data, and exports everything to a single Excel file.

## Features

- Accepts multiple `.docx` files per user session
- Extracts: **Name**, **Phone**, **City**, **Age**, **Positions**
- Exports to `candidates.xlsx` with formatted headers
- Per-user in-memory session with automatic temp-file cleanup
- Robust parsing: never crashes on malformed documents
- Ukrainian phone normalization (`+38 0XX XXX-XX-XX`)
- Structured logging

## Requirements

- Python 3.12+
- A Telegram bot token from [@BotFather](https://t.me/BotFather)

## Setup

```bash
# 1. Clone / navigate to the project
cd cvbot

# 2. Create and activate a virtual environment
python3.12 -m venv .venv
source .venv/bin/activate        # Linux / macOS
# .venv\Scripts\activate         # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure environment
cp .env.example .env
# Edit .env and set BOT_TOKEN=<your token>

# 5. Run
python app.py
```

## Project Structure

```
cvbot/
├── app.py                      # Entry point
├── core/
│   ├── config.py               # Environment config
│   └── logging.py              # Logging setup
├── bot/
│   ├── handlers/
│   │   ├── start.py            # /start command
│   │   └── files.py            # File upload + Finish button
│   └── keyboards/
│       └── common.py           # Inline keyboard
├── services/
│   ├── docx_parser.py          # Raw .docx text extraction
│   ├── resume_extractor.py     # Field parsing logic
│   └── excel_exporter.py       # Excel generation
├── requirements.txt
├── .env.example
└── README.md
```

## Excel Output Columns

| Column      | Description                        |
|-------------|------------------------------------|
| Name        | Full name of the candidate         |
| Phone       | Normalized phone number            |
| City        | City of residence                  |
| Age         | Age in years                       |
| Positions   | Desired + past job titles          |
| Source File | Original filename                  |
| Parsed At   | Timestamp of parsing               |

## Supported Resume Format

Optimized for [work.ua](https://www.work.ua) `.docx` exports, with fallback
heuristics for other common Ukrainian resume layouts.
