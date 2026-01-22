from connectors.excel_connector import ExcelConnector
from fastapi import FastAPI
from pydantic import BaseModel
from typing import Optional, Dict, Any

app = FastAPI()

# ---- Request Schema ----

class Command(BaseModel):
    action: str
    params: Optional[Dict[str, Any]] = {}
    mode: Optional[str] = None

connector = ExcelConnector("data.xlsx")

# ---- Dependency ----
# assume `connector` is already initialized elsewhere
# from connectors.excel_connector import ExcelConnector
# connector = ExcelConnector(...)

@app.post("/execute")
def execute_command(command: Command):
     
    try:
        result = connector.execute(
            action=command.action,
            params=command.params,
            mode=command.mode
        )
        return result
    except Exception as e:  
        return {'status': 'error',
                 'message': str(e),
                 'data': None
                }
    except ValueError as ve:
        return {'status': 'error',
                 'message': str(ve),
                 'data': None
                }
    