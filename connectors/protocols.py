"""
Execution Safety Protocols

This module defines:
- which actions are allowed
- which actions mutate state
- which parameters are required
- which execution modes are valid

This file contains NO business logic.
It only defines rules.
"""

# EXECUTION MODES

EXECUTION_MODES = {"execute", "dry_run"}

# ACTION CATEGORIES

READ_ONLY_ACTIONS = {
    "preview",
    "sum_column",
}

WRITE_ACTIONS = {
    "create_table",
    "add_row",
    "update_cell",
    "create_sheet",
    "set_active_sheet",
}

# ACTION SCHEMAS

ACTION_SCHEMAS = {
    "create_table": {"columns"},
    "add_row": {"value_dict"},
    "update_cell": {"column_name", "row_index", "new_value"},
    "sum_column": {"column_name"},
    "preview": set(),
    "set_active_sheet": {"sheet_name"},
    "create_sheet": {"sheet_name"},
}

action_intent = {
    "create_table": "write_destructive",
    "add_row": "write_safe",
    "update_cell": "write_safe",
    "sum_column": "read",
    "preview": "read",
    "set_active_sheet": "system",
    "create_sheet": "structural",    

}
# PROTOCOL ERRORS

ERROR_UNKNOWN_ACTION = "Unknown action"
ERROR_INVALID_PARAMS = "Params must be an object"
ERROR_MISSING_PARAMS = "Missing required parameters"
ERROR_INVALID_MODE = "Invalid execution mode"
ERROR_WRITE_REQUIRES_MODE = "Write actions require explicit execution mode"
ERROR_NO_EXTRA_PARAMS = "No extra parameters allowed"

class Protocols:

    def __init__(self):
        self.execution_modes = EXECUTION_MODES
        self.read_only_actions = READ_ONLY_ACTIONS
        self.write_actions = WRITE_ACTIONS
        self.action_schemas = ACTION_SCHEMAS
    
    def validate_action(self, action, mode=None):
        if action in self.read_only_actions:
            return True
        if action in self.write_actions:
            if mode in self.execution_modes:
                return True
            else:
                raise ValueError(ERROR_WRITE_REQUIRES_MODE)
        else:
            raise ValueError(ERROR_UNKNOWN_ACTION)
        
    def validate_params(self, action, params):
        if not isinstance(params, dict):
            raise ValueError(ERROR_INVALID_PARAMS)
        
        required = self.action_schemas.get(action, set())
        missing = required - params.keys()
        
        if missing:
            raise ValueError(f"{ERROR_MISSING_PARAMS}: {missing}")
        extra = params.keys() - ACTION_SCHEMAS.get(action, set())

        if extra:
            raise ValueError(ERROR_NO_EXTRA_PARAMS)

        return True