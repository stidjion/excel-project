# AI CLI Assistant

## Setup
1. Clone the repo:
```bash
git clone <repo_url>
cd <project_folder>

Activate virtual environment:

venv\Scripts\activate   # Windows
source venv/bin/activate   # Mac/Linux

Install dependencies:

pip install -r requirements.txt

api:
uvicorn api:app --reload

Run
python cli/entrypoint.py

---
# Excel Automation Backend

A simple, reliable backend for automating Excel tasks via structured commands (CLI or API). Built as a learning project to master backend architecture, contract design, and execution pipelines.

# Features
- Execute Excel actions: create sheets, add rows, preview data.
- Strict validation for safe execution.
- CLI and FastAPI interfaces.
- Ready for AI integration (natural language → command parsing).

# Architecture
in progress


## Setup
## Setup
1. Clone the repo:
```bash
git clone <repo_url>
cd <project_folder>

Activate virtual environment:

venv\Scripts\activate   # Windows
source venv/bin/activate   # Mac/Linux

Install dependencies:

pip install -r requirements.txt

api:
uvicorn api:app --reload

for testing purposes run
python tests.py

access the cli at
python cli/entrypoint.py

Feel free to fork/contribute!