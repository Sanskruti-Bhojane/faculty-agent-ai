"""
Attainment Agent

Gemini:
- understands syllabus, CO, PO and PSO documents
- identifies CO-PO and CO-PSO articulation
- interprets assessment structure when needed

Python:
- keeps faculty-uploaded marks authoritative
- keeps faculty-uploaded assessment -> CO mapping authoritative
- performs ALL numerical attainment calculations deterministically
- generates final CO, PO and PSO attainment values
"""

import json
import re
from collections.abc import Mapping
from typing import Any, Dict, List

from services.gemini_service import generate_text
from agents.attainment_calculator import calculate_complete_attainment


# ============================================================
# JSON CLEANING
# ============================================================

def _clean_json(text: str) -> Dict[str, Any]:
    """
    Gemini sometimes returns JSON inside ```json ... ```.

    This function removes markdown fences and extracts the
    outermost JSON object.
    """

    text = (text or "").strip()

    text = re.sub(
        r"^```json\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(
        r"^```\s*",
        "",
        text,
    )

    text = re.sub(
        r"\s*```$",
        "",
        text,
    )

    start = text.find("{")
    end = text.rfind("}")

    if start >= 0 and end > start:
        text = text[start:end + 1]

    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            "Gemini returned invalid JSON for the attainment analysis."
        ) from exc


# ============================================================
# BASIC NORMALIZATION
# ============================================================

def _strings(value: Any) -> List[str]:
    """
    Convert a JSON list into a clean list of strings.
    """

    if not isinstance(value, list):
        return []

    result = []

    for item in value:
        value = str(item).strip()

        if value:
            result.append(value)

    return result


def _mapping(
    value: Any,
) -> Dict[str, Dict[str, float]]:
    """
    Normalize CO-PO / CO-PSO articulation.

    Expected structure:

    {
        "CO1": {
            "PO1": 3,
            "PO2": 2
        }
    }
    """

    output: Dict[str, Dict[str, float]] = {}

    if not isinstance(value, Mapping):
        return output

    for co, row in value.items():

        if not isinstance(row, Mapping):
            continue

        co_code = str(co).strip()

        if not co_code:
            continue

        output[co_code] = {}

        for outcome, level in row.items():

            try:
                number = float(level)
            except (TypeError, ValueError):
                continue

            # Only valid articulation levels contribute.
            if number in (1.0, 2.0, 3.0):

                outcome_code = str(outcome).strip()

                if outcome_code:
                    output[co_code][outcome_code] = number

    return output


# ============================================================
# CANONICAL NAME
# ============================================================

def _canonical(value: Any) -> str:
    """
    Convert assessment names into a comparison-safe form.

    Example:
        "MSE-1" -> "mse1"
        "MSE 1" -> "mse1"
        "mse_1" -> "mse1"
    """

    return "".join(
        character.lower()
        for character in str(value or "")
        if character.isalnum()
    )


# ============================================================
# MERGE FACULTY MARKS WITH GEMINI UNDERSTANDING
# ============================================================

def _merge_raw_assessments(
    raw: List[Dict[str, Any]],
    ai_assessments: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Merge Gemini's assessment understanding with the actual
    assessment information extracted from the faculty Excel.

    IMPORTANT:

    Faculty Excel is authoritative.

    Therefore:
    - maximum marks come from Excel
    - student marks come from Excel
    - assessment -> CO mapping comes from Excel
    - Gemini cannot overwrite these values
    """

    result: List[Dict[str, Any]] = []

    # First keep Gemini's assessment understanding.
    for item in ai_assessments or []:

        if not isinstance(item, Mapping):
            continue

        result.append(dict(item))

    # Now merge the authoritative faculty Excel data.
    for item in raw or []:

        if not isinstance(item, Mapping):
            continue

        name = str(
            item.get(
                "name",
                item.get(
                    "assessment_type",
                    "Assessment",
                ),
            )
        ).strip()

        if not name:
            name = "Assessment"

        match = next(
            (
                existing
                for existing in result
                if _canonical(existing.get("name"))
                == _canonical(name)
            ),
            None,
        )

        if match is None:

            match = {
                "name": name,
                "type": (
                    "ESE"
                    if _canonical(name) == "ese"
                    else "IA"
                ),
            }

            result.append(match)

        # ----------------------------------------------------
        # FACULTY EXCEL IS AUTHORITATIVE
        # ----------------------------------------------------

        match["name"] = name

        match["type"] = (
            "ESE"
            if _canonical(name) == "ese"
            else str(
                item.get(
                    "type",
                    "IA",
                )
            )
        )

        match["co_max_marks"] = dict(
            item.get(
                "co_max_marks",
                {},
            ) or {}
        )

        match["co_student_marks"] = dict(
            item.get(
                "co_student_marks",
                {},
            ) or {}
        )

        # Preserve extractor metadata if available.
        for key in (
            "mapping",
            "assessment_mapping",
            "source_columns",
            "source_headers",
            "original_mapping",
        ):
            if key in item:
                match[key] = item[key]

    return result


# ============================================================
# BUILD CALCULATOR INPUT
# ============================================================

def _build_calculation_inputs(
    assessments: List[Dict[str, Any]],
):
    """
    Convert extractor assessment data into the deterministic
    calculator format.

    Example:

    assessment_results = {
        "MSE-1": {
            "CO1": {
                "max_marks": 10,
                "student_marks": [...]
            }
        }
    }

    assessment_mapping = {
        "MSE-1": ["CO1"]
    }
    """

    assessment_results: Dict[str, Dict[str, Any]] = {}
    assessment_mapping: Dict[str, List[str]] = {}

    for assessment in assessments:

        if not isinstance(assessment, Mapping):
            continue

        name = str(
            assessment.get(
                "name",
                "Assessment",
            )
        ).strip()

        if not name:
            name = "Assessment"

        maximum = (
            assessment.get(
                "co_max_marks",
                {},
            )
            or {}
        )

        marks = (
            assessment.get(
                "co_student_marks",
                {},
            )
            or {}
        )

        assessment_results[name] = {}

        mapped_cos: List[str] = []

        # ----------------------------------------------------
        # Collect all COs from maximum marks and student marks.
        # ----------------------------------------------------

        all_cos = []

        for co in list(maximum.keys()) + list(marks.keys()):

            co = str(co).strip()

            if co and co not in all_cos:
                all_cos.append(co)

        for co in all_cos:

            if co not in mapped_cos:
                mapped_cos.append(co)

            values = marks.get(
                co,
                [],
            )

            if not isinstance(values, list):
                values = [values]

            assessment_results[name][co] = {
                "max_marks": maximum.get(co),
                "student_marks": values,
            }

        assessment_mapping[name] = mapped_cos

    return (
        assessment_results,
        assessment_mapping,
    )


# ============================================================
# SURVEY NORMALIZATION
# ============================================================

def _normalize_survey_data(
    survey_ratings=None,
    survey_percentages=None,
    survey_overall_percentage=None,
):
    """
    Support both:

    New extractor format:
    {
        "ratings": {...},
        "percentages": {...},
        "overall_percentage": ...
    }

    and older flat format:
    {
        "CO1": 4.5,
        "CO2": 4.2
    }
    """

    if (
        isinstance(
            survey_ratings,
            Mapping,
        )
        and "ratings" in survey_ratings
    ):

        details = dict(survey_ratings)

        ratings = (
            details.get(
                "ratings",
                {},
            )
            or {}
        )

        percentages = (
            details.get(
                "percentages",
                {},
            )
            or {}
        )

        overall_percentage = details.get(
            "overall_percentage"
        )

        return (
            ratings,
            percentages,
            overall_percentage,
            details,
        )

    ratings = (
        survey_ratings
        if isinstance(
            survey_ratings,
            Mapping,
        )
        else {}
    )

    percentages = (
        survey_percentages
        if isinstance(
            survey_percentages,
            Mapping,
        )
        else {}
    )

    overall_percentage = survey_overall_percentage

    details = {
        "ratings": ratings,
        "percentages": percentages,
        "overall_percentage": overall_percentage,
    }

    return (
        ratings,
        percentages,
        overall_percentage,
        details,
    )


# ============================================================
# MAIN ATTAINMENT AGENT
# ============================================================

def analyze_assessments(
    *,
    syllabus_text: str,
    co_text: str,
    po_text: str,
    pso_text: str,
    assessment_text: str,
    raw_marks_data=None,
    survey_ratings=None,
    survey_percentages=None,
    survey_overall_percentage=None,
    course_code: str = "",
    course_name: str = "",
    semester: str = "",
    academic_year: str = "",
):
    """
    Main orchestration function.

    Gemini:
        Document understanding + articulation reasoning.

    Python:
        All numerical attainment calculations.
    """

    # ========================================================
    # 1. GEMINI REASONING
    # ========================================================

    prompt = f"""
You are the reasoning agent inside an agentic OBE
Faculty Assistance System.

Your task is to understand the uploaded academic documents
and return structured information.

IMPORTANT:
Return ONLY valid JSON.
Do not use markdown.
Do not use ```json fences.

AUTHORITATIVE RULES:

1. CO DOCUMENT
   - Use only the official COs supplied by the faculty.
   - Never invent COs.
   - Never rename COs.

2. PO DOCUMENT
   - The uploaded PO document is official.
   - Never invent POs.
   - Never rename POs.
   - Never add POs.
   - Never remove POs.

3. PSO DOCUMENT
   - The uploaded PSO document is official.
   - Never invent PSOs.
   - Never rename PSOs.
   - Never add PSOs.
   - Never remove PSOs.

4. STUDENT MARKS EXCEL
   - Student marks are authoritative.
   - Assessment names are authoritative.
   - Assessment -> CO mapping is authoritative.
   - Maximum marks are authoritative.
   - Do NOT change the uploaded assessment mapping.
   - Do NOT calculate numerical attainment.

5. COURSE EXIT SURVEY
   - Use the survey only for indirect CO attainment information.
   - Do NOT invent survey responses.

6. CO-PO / CO-PSO ARTICULATION
   Use:
       1 = Low
       2 = Moderate
       3 = High
       blank = No contribution

   Determine articulation using the syllabus,
   CO statements and outcome statements.

7. NUMERICAL CALCULATION
   Do NOT calculate:
   - thresholds
   - percentages
   - attainment levels
   - direct attainment
   - indirect attainment
   - overall IA
   - ESE attainment
   - overall CO attainment
   - PO attainment
   - PSO attainment

   These are calculated by deterministic Python code.

8. DO NOT GENERATE PSO.
   PSO is compulsory and supplied by the faculty.

Return exactly:

{{
    "co_codes": [],
    "po_codes": [],
    "pso_codes": [],
    "co_po_mapping": {{}},
    "co_pso_mapping": {{}},
    "assessments": [],
    "assumptions": []
}}

SYLLABUS:
{syllabus_text}

COURSE OUTCOMES:
{co_text}

OFFICIAL PROGRAM OUTCOMES:
{po_text}

OFFICIAL PROGRAM SPECIFIC OUTCOMES:
{pso_text}

STUDENT MARKS + COURSE EXIT SURVEY:
{assessment_text}
"""

    gemini_response = generate_text(prompt)

    data = _clean_json(
        gemini_response
    )

    # ========================================================
    # 2. EXTRACT OFFICIAL OUTCOME CODES
    # ========================================================

    co_codes = _strings(
        data.get(
            "co_codes"
        )
    )

    po_codes = _strings(
        data.get(
            "po_codes"
        )
    )

    pso_codes = _strings(
        data.get(
            "pso_codes"
        )
    )

    # ========================================================
    # 3. EXTRACT ARTICULATION
    # ========================================================

    co_po_mapping = _mapping(
        data.get(
            "co_po_mapping"
        )
    )

    co_pso_mapping = _mapping(
        data.get(
            "co_pso_mapping"
        )
    )

    # ========================================================
    # 4. MERGE FACULTY MARKS WITH AI ASSESSMENT UNDERSTANDING
    # ========================================================

    assessments = _merge_raw_assessments(
        raw_marks_data
        if isinstance(
            raw_marks_data,
            list,
        )
        else [],
        data.get(
            "assessments",
            [],
        ),
    )

    # ========================================================
    # 5. ENSURE ALL COs FOUND IN EXCEL ARE INCLUDED
    # ========================================================

    for assessment in assessments:

        maximum = (
            assessment.get(
                "co_max_marks",
                {},
            )
            or {}
        )

        marks = (
            assessment.get(
                "co_student_marks",
                {},
            )
            or {}
        )

        for co in list(
            maximum.keys()
        ) + list(
            marks.keys()
        ):

            co = str(co).strip()

            if (
                co
                and co not in co_codes
            ):
                co_codes.append(co)

    # ========================================================
    # 6. BUILD DETERMINISTIC CALCULATION INPUTS
    # ========================================================

    (
        assessment_results,
        assessment_mapping,
    ) = _build_calculation_inputs(
        assessments
    )

    # ========================================================
    # 7. SURVEY DATA
    # ========================================================

    (
        ratings,
        percentages,
        overall_percentage,
        survey_details,
    ) = _normalize_survey_data(
        survey_ratings=survey_ratings,
        survey_percentages=survey_percentages,
        survey_overall_percentage=survey_overall_percentage,
    )

    # ========================================================
    # 8. DETERMINISTIC OBE CALCULATION
    # ========================================================

    calculation = calculate_complete_attainment(
        assessment_results=assessment_results,
        survey_ratings=ratings,
        survey_percentages=percentages,
        survey_overall_percentage=overall_percentage,
        co_po_mapping=co_po_mapping,
        co_pso_mapping=co_pso_mapping,
        assessment_mapping=assessment_mapping,
        co_codes=co_codes,
        po_codes=po_codes,
        pso_codes=pso_codes,
    )

    # ========================================================
    # 9. STUDENT INFORMATION
    # ========================================================

    students = []

    total_students = 0

    if isinstance(
        raw_marks_data,
        Mapping,
    ):

        students = (
            raw_marks_data.get(
                "students",
                [],
            )
            or []
        )

        total_students = (
            raw_marks_data.get(
                "total_students",
                len(students),
            )
            or 0
        )

    # ========================================================
    # 10. FINAL AGENT RESULT
    # ========================================================

    return {

        "analysis_status": "COMPLETED",

        "calculation_status": "SUCCESS",

        # ----------------------------------------------------
        # COURSE
        # ----------------------------------------------------

        "course": {
            "course_code": course_code,
            "course_name": course_name,
            "semester": semester,
            "academic_year": academic_year,
        },

        # ----------------------------------------------------
        # OUTCOME CODES
        # ----------------------------------------------------

        "co_codes": co_codes,

        "po_codes": po_codes,

        "pso_codes": pso_codes,

        # ----------------------------------------------------
        # ARTICULATION MATRICES
        # ----------------------------------------------------

        "co_po_articulation_matrix": co_po_mapping,

        "co_pso_articulation_matrix": co_pso_mapping,

        # ----------------------------------------------------
        # ASSESSMENT DATA
        # ----------------------------------------------------

        "assessment_mapping": assessment_mapping,

        "assessment_results": assessment_results,

        "assessments": assessments,

        # ----------------------------------------------------
        # STUDENTS
        # ----------------------------------------------------

        "students": students,

        "total_students": total_students,

        # ----------------------------------------------------
        # SURVEY
        # ----------------------------------------------------

        "survey_details": survey_details,

        "survey_ratings": ratings,

        "survey_percentages": percentages,

        # ----------------------------------------------------
        # AI ASSUMPTIONS
        # ----------------------------------------------------

        "assumptions": data.get(
            "assumptions",
            [],
        ),

        # ----------------------------------------------------
        # ATTAINMENT
        # ----------------------------------------------------

        "overall_co_attainment":
            calculation.get(
                "overall_co_attainment",
                {},
            ),

        "co_attainment":
            calculation.get(
                "co_attainment",
                {},
            ),

        "direct_co_attainment":
            calculation.get(
                "direct_co_attainment",
                {},
            ),

        "indirect_co_attainment":
            calculation.get(
                "indirect_co_attainment",
                {},
            ),

        "course_indirect_attainment":
            calculation.get(
                "course_indirect_attainment",
                0.0,
            ),

        "overall_ia":
            calculation.get(
                "overall_ia",
                {},
            ),

        "ese_attainment":
            calculation.get(
                "ese_attainment",
                {},
            ),

        # ----------------------------------------------------
        # PO / PSO
        # ----------------------------------------------------

        "po_attainment":
            calculation.get(
                "po_attainment",
                {},
            ),

        "pso_attainment":
            calculation.get(
                "pso_attainment",
                {},
            ),

        "co_po_attainment":
            calculation.get(
                "co_po_attainment",
                {},
            ),

        "co_pso_attainment":
            calculation.get(
                "co_pso_attainment",
                {},
            ),

        # ----------------------------------------------------
        # ASSESSMENT STATISTICS
        # ----------------------------------------------------

        "assessment_statistics":
            calculation.get(
                "assessment_statistics",
                {},
            ),

        # ----------------------------------------------------
        # METHODOLOGY
        # ----------------------------------------------------

        "methodology_used":
            calculation.get(
                "methodology_used",
                {},
            ),
    }