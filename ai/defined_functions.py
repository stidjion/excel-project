create_table_function = {
    "name": "create_table",
    "description": (
        "Creates a fresh table in the active sheet. "
        "Overwrites existing sheet data. "
        "Requires a non-empty list of column names."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "columns": {
                "type": "array",
                "items": {"type": "string"},
                "description": (
                    "List of column names for the new table. "
                    "Must be a non-empty list of strings."
                ),
            },
            "mode": {
                "type": "string",
                "enum": ["execute", "dry_run"],
                "description": "Execution mode for the operation.",
            },
        },
        "required": ["columns"],
    },
}

add_row_function = {
    "name": "add_row",
    "description": (
        "Adds a single row of data to the active table. "
        "Keys must match existing table columns. "
        "Missing columns are allowed and filled as NaN. "
        "Extra columns will cause an error."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "value_dict": {
                "type": "object",
                "additionalProperties": True,
                "description": (
                    "Dictionary mapping column names to values "
                    "for the new row."
                ),
            },
            "mode": {
                "type": "string",
                "enum": ["execute", "dry_run"],
                "description": "Execution mode for the operation.",
            },
        },
        "required": ["value_dict"],
    },
}

update_cell_function = {
    "name": "update_cell",
    "description": (
        "Updates a single cell in the active table. "
        "Row index is zero-based. "
        "Column must already exist. "
        "Overwrites the existing value."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "column_name": {
                "type": "string",
                "description": "Name of the column to update.",
            },
            "row_index": {
                "type": "integer",
                "description": "Zero-based index of the row to update.",
            },
            "new_value": {
                "description": "New value to place in the cell.",
            },
            "mode": {
                "type": "string",
                "enum": ["execute", "dry_run"],
                "description": "Execution mode for the operation.",
            },
        },
        "required": ["column_name", "row_index", "new_value"],
    },
}

sum_column_function = {
    "name": "sum_column",
    "description": (
        "Returns the numeric sum of a column in the active table. "
        "Non-numeric values are ignored. "
        "Strings are not concatenated. "
        "This is a read-only operation and does not modify data."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "column_name": {
                "type": "string",
                "description": "Name of the column to sum.",
            },
        },
        "required": ["column_name"],
    },
}

preview_function = {
    "name": "preview",
    "description": (
        "Shows the top N rows of the active table. "
        "Read-only operation."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "n": {
                "type": "integer",
                "description": (
                    "Number of rows to preview. "
                    "Defaults to 5 if not provided."
                ),
            },
        },
        "required": [],
    },
}

set_active_sheet_function = {
    "name": "set_active_sheet",
    "description": (
        "Switches the active sheet. "
        "The sheet must already exist. "
        "Does not create a new sheet."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "sheet_name": {
                "type": "string",
                "description": "Name of the sheet to activate.",
            },
        },
        "required": ["sheet_name"],
    },
}

create_sheet_function = {
    "name": "create_sheet",
    "description": (
        "Creates a new sheet and switches to it. "
        "Raises an error if the sheet already exists."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "sheet_name": {
                "type": "string",
                "description": "Name of the new sheet to create.",
            },
            "mode": {
                "type": "string",
                "enum": ["execute", "dry_run"],
                "description": "Execution mode for the operation.",
            },
        },
        "required": ["sheet_name"],
    },
}

defined_functions = [
    create_table_function,
    add_row_function,
    update_cell_function,
    sum_column_function,
    preview_function,
    set_active_sheet_function,
    create_sheet_function,
]