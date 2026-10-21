def create_co_po_matrix(mapping_result: dict) -> dict:
    """
    Converts the agent's mapping JSON into a table structure.
    """

    rows = mapping_result.get("co_po_mapping", [])

    po_codes = []

    for row in rows:
        for mapping in row.get("mappings", []):
            po_code = mapping.get("po_code")

            if po_code and po_code not in po_codes:
                po_codes.append(po_code)

    matrix = []

    for row in rows:

        co_code = row.get("co_code")

        values = {
            po: None
            for po in po_codes
        }

        for mapping in row.get("mappings", []):

            po_code = mapping.get("po_code")
            level = mapping.get("level")

            if po_code in values:
                values[po_code] = level

        matrix.append({
            "co_code": co_code,
            **values
        })

    return {
        "columns": ["CO"] + po_codes,
        "rows": matrix
    }


def create_co_pso_matrix(mapping_result: dict) -> dict:
    """
    Converts CO-PSO agent output into a table structure.
    """

    rows = mapping_result.get("co_pso_mapping", [])

    pso_codes = []

    for row in rows:
        for mapping in row.get("mappings", []):

            pso_code = mapping.get("pso_code")

            if pso_code and pso_code not in pso_codes:
                pso_codes.append(pso_code)

    matrix = []

    for row in rows:

        co_code = row.get("co_code")

        values = {
            pso: None
            for pso in pso_codes
        }

        for mapping in row.get("mappings", []):

            pso_code = mapping.get("pso_code")
            level = mapping.get("level")

            if pso_code in values:
                values[pso_code] = level

        matrix.append({
            "co_code": co_code,
            **values
        })

    return {
        "columns": ["CO"] + pso_codes,
        "rows": matrix
    }