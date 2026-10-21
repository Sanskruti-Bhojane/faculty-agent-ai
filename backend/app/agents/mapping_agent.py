import json
import re

from services.gemini_service import generate_text


# ============================================================
# EXTRACT CODES FROM FACULTY DOCUMENTS
# ============================================================

def extract_codes(text: str, prefix: str) -> list[str]:
    """
    Extract codes such as:
        CO1, CO2, CO3
        PO1, PO2, PO3
        PSO1, PSO2, PSO3

    Only codes actually present in the faculty document are used.
    """

    if not text:
        return []

    pattern = (
        rf"(?<![A-Z0-9])"
        rf"{prefix}\s*[-:]?\s*(\d+)"
        rf"(?![A-Z0-9])"
    )

    matches = re.findall(
        pattern,
        text,
        flags=re.IGNORECASE
    )

    codes = []

    for number in matches:

        code = f"{prefix.upper()}{number}"

        if code not in codes:
            codes.append(code)

    return codes


# ============================================================
# CLEAN GEMINI JSON RESPONSE
# ============================================================

def clean_json_response(response_text: str) -> dict:
    """
    Convert Gemini response into a Python dictionary.

    Handles:
    - normal JSON
    - ```json ... ```
    - surrounding explanation text
    """

    if not isinstance(response_text, str):

        if isinstance(response_text, dict):
            return response_text

        raise ValueError(
            "AI response is neither a string nor a dictionary."
        )

    text = response_text.strip()

    # Remove markdown JSON fences
    text = re.sub(
        r"^```json\s*",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"^```\s*",
        "",
        text
    )

    text = re.sub(
        r"\s*```$",
        "",
        text
    )

    text = text.strip()

    # --------------------------------------------------------
    # Direct JSON
    # --------------------------------------------------------

    try:

        result = json.loads(text)

        if isinstance(result, dict):
            return result

    except json.JSONDecodeError:
        pass

    # --------------------------------------------------------
    # JSON object embedded inside text
    # --------------------------------------------------------

    start = text.find("{")
    end = text.rfind("}")

    if start != -1 and end != -1 and end > start:

        json_text = text[
            start:end + 1
        ]

        try:

            result = json.loads(
                json_text
            )

            if isinstance(result, dict):
                return result

        except json.JSONDecodeError:
            pass

    raise ValueError(
        "Gemini returned an invalid JSON response."
    )


# ============================================================
# NORMALIZE MAPPING VALUE
# ============================================================

def normalize_mapping_value(value):
    """
    Convert Gemini mapping values into:

        1 = Low
        2 = Moderate
        3 = High
        None = No meaningful contribution

    Handles both integers and strings:

        1
        "1"
        " 1 "
        2
        "2"
        3
        "3"
        None
        ""
        "null"
    """

    if value is None:
        return None

    # Boolean should never be accepted as a mapping level.
    if isinstance(value, bool):
        return None

    # Already numeric
    if isinstance(value, (int, float)):

        try:
            integer_value = int(value)
        except (TypeError, ValueError):
            return None

        if integer_value in [1, 2, 3]:
            return integer_value

        return None

    # String value
    text = str(value).strip().lower()

    if text in [
        "",
        "null",
        "none",
        "nil",
        "n/a",
        "na",
        "-",
        "blank",
        "no mapping",
        "no meaningful contribution"
    ]:
        return None

    # Remove accidental formatting
    text = text.strip(
        "\"'` "
    )

    # Handle numeric strings
    if text in ["1", "2", "3"]:

        return int(text)

    # Handle labels if Gemini returns them
    if text == "low":
        return 1

    if text == "moderate":
        return 2

    if text == "medium":
        return 2

    if text == "high":
        return 3

    # Handle strings such as "3 - High"
    match = re.match(
        r"^([123])(?:\s*[-:]\s*.*)?$",
        text
    )

    if match:

        return int(
            match.group(1)
        )

    return None


# ============================================================
# NORMALIZE MAPPINGS DICTIONARY
# ============================================================

def normalize_mappings_dict(
    mappings,
    valid_codes
):
    """
    Normalize Gemini's mappings dictionary.

    Example input:

        {
            "PO1": "3",
            "PO2": 2,
            "PO3": "null"
        }

    Output:

        {
            "PO1": 3,
            "PO2": 2,
            "PO3": None
        }

    Only official codes are retained.
    """

    if not isinstance(
        mappings,
        dict
    ):
        mappings = {}

    normalized = {}

    # Create complete rectangular structure
    for code in valid_codes:

        value = None

        # Exact match
        if code in mappings:

            value = mappings[code]

        else:

            # Case-insensitive match
            for key, candidate in mappings.items():

                if (
                    str(key)
                    .strip()
                    .upper()
                    ==
                    code.upper()
                ):

                    value = candidate
                    break

        normalized[
            code
        ] = normalize_mapping_value(
            value
        )

    return normalized


# ============================================================
# NORMALIZE MAPPING
# ============================================================

def normalize_mapping(
    result: dict,
    co_codes: list[str],
    po_codes: list[str],
    pso_codes: list[str]
) -> dict:
    """
    Ensure a complete CO-PO-PSO mapping structure.

    IMPORTANT:
    - COs come from faculty CO document/syllabus.
    - POs come ONLY from official PO document.
    - PSOs come ONLY from faculty PSO document.
    - No attainment calculation is performed.
    """

    if not isinstance(
        result,
        dict
    ):
        raise ValueError(
            "Mapping result must be a dictionary."
        )

    # --------------------------------------------------------
    # CO CODES
    # --------------------------------------------------------

    if not co_codes:

        ai_cos = result.get(
            "co_codes",
            []
        )

        if isinstance(
            ai_cos,
            list
        ):

            co_codes = [
                str(x).strip().upper()
                for x in ai_cos
                if str(x).strip()
            ]

    if not co_codes:

        raise ValueError(
            "No CO codes were found."
        )

    # --------------------------------------------------------
    # OFFICIAL PO CODES
    # --------------------------------------------------------

    if not po_codes:

        raise ValueError(
            "No official PO codes were found in "
            "the uploaded PO document."
        )

    # --------------------------------------------------------
    # PSO CODES
    # --------------------------------------------------------

    if not pso_codes:

        ai_psos = result.get(
            "pso_codes",
            []
        )

        if isinstance(
            ai_psos,
            list
        ):

            pso_codes = [
                str(x).strip().upper()
                for x in ai_psos
                if str(x).strip()
            ]

    if not pso_codes:

        raise ValueError(
            "No PSO codes were found in "
            "the uploaded PSO document."
        )

    # --------------------------------------------------------
    # RAW CO-PO DATA
    # --------------------------------------------------------

    raw_co_po = result.get(
        "co_po_mapping",
        result.get(
            "co_po",
            []
        )
    )

    if not isinstance(
        raw_co_po,
        list
    ):
        raw_co_po = []

    # --------------------------------------------------------
    # RAW CO-PSO DATA
    # --------------------------------------------------------

    raw_co_pso = result.get(
        "co_pso_mapping",
        result.get(
            "co_pso",
            []
        )
    )

    if not isinstance(
        raw_co_pso,
        list
    ):
        raw_co_pso = []

    # --------------------------------------------------------
    # INDEX CO-PO ROWS
    # --------------------------------------------------------

    co_po_lookup = {}

    for row in raw_co_po:

        if not isinstance(
            row,
            dict
        ):
            continue

        co_code = (
            row.get("co_code")
            or row.get("co")
            or row.get("CO")
        )

        if not co_code:
            continue

        co_code = str(
            co_code
        ).strip().upper()

        mappings = row.get(
            "mappings",
            {}
        )

        if not isinstance(
            mappings,
            dict
        ):
            mappings = {}

        co_po_lookup[
            co_code
        ] = mappings

    # --------------------------------------------------------
    # INDEX CO-PSO ROWS
    # --------------------------------------------------------

    co_pso_lookup = {}

    for row in raw_co_pso:

        if not isinstance(
            row,
            dict
        ):
            continue

        co_code = (
            row.get("co_code")
            or row.get("co")
            or row.get("CO")
        )

        if not co_code:
            continue

        co_code = str(
            co_code
        ).strip().upper()

        mappings = row.get(
            "mappings",
            {}
        )

        if not isinstance(
            mappings,
            dict
        ):
            mappings = {}

        co_pso_lookup[
            co_code
        ] = mappings

    # --------------------------------------------------------
    # BUILD COMPLETE CO-PO MATRIX
    # --------------------------------------------------------

    final_co_po = []

    for co in co_codes:

        original_mappings = (
            co_po_lookup.get(
                co,
                {}
            )
        )

        normalized_mappings = (
            normalize_mappings_dict(
                original_mappings,
                po_codes
            )
        )

        final_co_po.append({

            "co_code":
                co,

            "mappings":
                normalized_mappings
        })

    # --------------------------------------------------------
    # BUILD COMPLETE CO-PSO MATRIX
    # --------------------------------------------------------

    final_co_pso = []

    for co in co_codes:

        original_mappings = (
            co_pso_lookup.get(
                co,
                {}
            )
        )

        normalized_mappings = (
            normalize_mappings_dict(
                original_mappings,
                pso_codes
            )
        )

        final_co_pso.append({

            "co_code":
                co,

            "mappings":
                normalized_mappings
        })

    # --------------------------------------------------------
    # JUSTIFICATIONS
    # --------------------------------------------------------

    justifications = result.get(
        "justifications",
        []
    )

    if not isinstance(
        justifications,
        list
    ):
        justifications = []

    # --------------------------------------------------------
    # FINAL RESULT
    # --------------------------------------------------------

    return {

        "co_codes":
            co_codes,

        "po_codes":
            po_codes,

        "pso_codes":
            pso_codes,

        "co_po_mapping":
            final_co_po,

        "co_pso_mapping":
            final_co_pso,

        "justifications":
            justifications,

        "mapping_scale":
            {
                "1": "Low",
                "2": "Moderate",
                "3": "High",
                "null":
                    "No meaningful contribution"
            }

    }


# ============================================================
# MAIN CO-PO-PSO MAPPING AGENT
# ============================================================

def generate_mapping(
    syllabus_text: str,
    co_text: str,
    po_text: str,
    pso_text: str
):
    """
    Generate evidence-based CO-PO-PSO mapping.

    THIS AGENT ONLY DOES:
        CO → PO mapping
        CO → PSO mapping

    It does NOT calculate:
        - CO attainment
        - PO attainment
        - PSO attainment
        - student marks
        - thresholds
        - course exit survey
        - ESE
        - internal assessment

    Official POs are taken only from the uploaded
    official PO document.
    """

    # ========================================================
    # 1. EXTRACT FACULTY-PROVIDED CODES
    # ========================================================

    co_codes = extract_codes(
        co_text,
        "CO"
    )

    # If CO document doesn't contain codes,
    # use syllabus as fallback.
    if not co_codes:

        co_codes = extract_codes(
            syllabus_text,
            "CO"
        )

    po_codes = extract_codes(
        po_text,
        "PO"
    )

    pso_codes = extract_codes(
        pso_text,
        "PSO"
    )

    # ========================================================
    # 2. VALIDATION
    # ========================================================

    if not co_codes:

        raise ValueError(
            "No CO codes found. "
            "Please upload a document containing "
            "CO1, CO2, etc."
        )

    if not po_codes:

        raise ValueError(
            "No official PO codes found. "
            "Please upload the official PO document "
            "containing PO1, PO2, etc."
        )

    if not pso_codes:

        raise ValueError(
            "No PSO codes found. "
            "Please upload the PSO document "
            "containing PSO1, PSO2, etc."
        )

    # ========================================================
    # 3. GEMINI PROMPT
    # ========================================================

    prompt = f"""
You are an Outcome Based Education (OBE) CO-PO-PSO Mapping Agent.

Your ONLY task is to generate an evidence-based:

1. CO → PO mapping
2. CO → PSO mapping

Do NOT calculate attainment.

Do NOT perform:
- student marks analysis
- threshold calculation
- CO attainment
- PO attainment
- PSO attainment
- internal assessment
- ESE calculation
- course exit survey calculation

============================================================
STRICT RULES
============================================================

1. Do NOT invent official POs.

2. Use ONLY the PO codes and statements provided
   in the official PO document.

3. Use ONLY the COs provided by the faculty
   or syllabus.

4. Use ONLY the PSOs provided by the faculty.

5. Do NOT generate new POs.

6. Do NOT generate new PSOs.

7. Do NOT change CO, PO or PSO identifiers.

8. Do NOT map every CO to every PO.

9. Do NOT map every CO to every PSO.

10. Mapping must be based on actual contribution
    between the CO and the PO/PSO.

11. Mapping levels are:

    1 = Low contribution
    2 = Moderate contribution
    3 = High contribution
    null = No meaningful contribution

12. Mapping values MUST be JSON numbers,
    NOT strings.

    Correct:
        "PO1": 3

    Incorrect:
        "PO1": "3"

13. null is allowed.

14. Every CO must contain every official PO
    as a matrix column.

15. Every CO must contain every official PSO
    as a matrix column.

16. If there is no meaningful contribution,
    return null.

17. Return valid JSON only.

18. Do not add explanations outside JSON.

============================================================
FACULTY CO CODES
============================================================

{json.dumps(co_codes)}

============================================================
OFFICIAL PO CODES
============================================================

{json.dumps(po_codes)}

============================================================
FACULTY PSO CODES
============================================================

{json.dumps(pso_codes)}

============================================================
SYLLABUS
============================================================

{syllabus_text}

============================================================
CO DOCUMENT
============================================================

{co_text}

============================================================
OFFICIAL PO DOCUMENT
============================================================

{po_text}

============================================================
PSO DOCUMENT
============================================================

{pso_text}

============================================================
RETURN EXACTLY THIS JSON STRUCTURE
============================================================

{{
    "co_codes": {json.dumps(co_codes)},

    "po_codes": {json.dumps(po_codes)},

    "pso_codes": {json.dumps(pso_codes)},

    "co_po_mapping": [
        {{
            "co_code": "CO1",
            "mappings": {{
                "PO1": 3,
                "PO2": 2,
                "PO3": null
            }}
        }}
    ],

    "co_pso_mapping": [
        {{
            "co_code": "CO1",
            "mappings": {{
                "PSO1": 2,
                "PSO2": null,
                "PSO3": 1
            }}
        }}
    ],

    "justifications": [
        {{
            "co_code": "CO1",
            "po_code": "PO1",
            "justification":
                "Brief evidence-based reason."
        }}
    ],

    "mapping_scale": {{
        "1": "Low",
        "2": "Moderate",
        "3": "High",
        "null": "No meaningful contribution"
    }}
}}

IMPORTANT:
All mapping levels MUST be JSON integers:
1, 2, 3

or JSON null.

Never return:
"1"
"2"
"3"

as strings.
"""

    # ========================================================
    # 4. CALL GEMINI
    # ========================================================

    raw_response = generate_text(
        prompt
    )

    # ========================================================
    # 5. PARSE GEMINI JSON
    # ========================================================

    result = clean_json_response(
        raw_response
    )

    # ========================================================
    # 6. NORMALIZE RESULT
    # ========================================================

    final_result = normalize_mapping(

        result=result,

        co_codes=co_codes,

        po_codes=po_codes,

        pso_codes=pso_codes

    )

    # ========================================================
    # 7. FINAL VALIDATION
    # ========================================================

    if not final_result.get(
        "co_po_mapping"
    ):
        raise ValueError(
            "CO-PO mapping was not generated."
        )

    if not final_result.get(
        "co_pso_mapping"
    ):
        raise ValueError(
            "CO-PSO mapping was not generated."
        )

    # Make sure every CO has the expected PO columns
    for row in final_result[
        "co_po_mapping"
    ]:

        mappings = row.get(
            "mappings",
            {}
        )

        if set(
            mappings.keys()
        ) != set(po_codes):

            raise ValueError(
                f"Incomplete PO mapping for "
                f"{row.get('co_code')}."
            )

    # Make sure every CO has the expected PSO columns
    for row in final_result[
        "co_pso_mapping"
    ]:

        mappings = row.get(
            "mappings",
            {}
        )

        if set(
            mappings.keys()
        ) != set(pso_codes):

            raise ValueError(
                f"Incomplete PSO mapping for "
                f"{row.get('co_code')}."
            )

    return final_result