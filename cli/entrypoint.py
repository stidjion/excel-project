'''
import json
from connectors.excel_connector import ExcelConnector

def main():
    connector = ExcelConnector("data.xlsx")

    print("Excel CLI ready.")
    print("Type JSON commands or 'exit'.")

    while True:
        raw = input(">>> ")

        if raw.lower() in ["exit", "quit"]:
            break

        try:
            command = json.loads(raw)
            action = command.get("action")
            params = command.get("params", {})

            result = connector.execute(action, params)
            print(json.dumps(result, indent=2))

        except json.JSONDecodeError:
            print("Invalid JSON.")
        except Exception as e:
            print("Error:", e)


if __name__ == "__main__":
    main()

'''