import json
import re

from backend.app.utils.document_parser import extract_document
from services.gemini_service import generate_text


DOCUMENT_AGENT_PROMPT = """
You are the Document Understanding Agent of a Faculty OBE Agentic AI System.

Your job is ONLY to extract and organize information from the uploaded document.

IMPORTANT RULES:

1. NEVER invent official Program Outcomes (POs).
2. NEVER invent PSOs.
3. NEVER create new COs.
4. Extract COs only if they actually appear in the document.
5. Do NOT treat existing CO-PO mapping numbers as newly generated mappings.
6. If the document contains existing mappings, preserve them under existing_mappings.
7. Do not assume a fixed number of COs, POs or PSOs.
8. Do not assume a specific subject.
9. Extract course information when available.
10. Extract course objectives when available.
11. Extract Bloom's Taxonomy levels when available.
12. Extract assessment information when available.
13. Extract attainment methodology, formulas, targets and weightages when available.
14. If information is missing, return an empty list or null.
15. Do NOT generate missing information.
16. Return ONLY valid JSON. Do NOT use markdown fences.

Return exactly this structure:

{
    "course_information": {
        "course_code": null,
        "course_name": null,
        "semester": null,
        "credits": null,
        "course_type": null
    },
    "course_objectives": [],
    "course_outcomes": [],
    "program_outcomes": [],
    "program_specific_outcomes": [],
    "bloom_levels": [],
    "assessments": [],
    "existing_mappings": [],
    "question_co_mappings": [],
    "attainment_methodology": {
        "direct_weightage": null,
        "indirect_weightage": null,
        "target": null,
        "formula": null,
        "attainment_levels": []
    },
    "important_evidence": [],
    "missing_information": []
}
"""


def clean_json_response(response: str) -> str:
    """
    Remove markdown code fences if Gemini adds them.
    """
    response = response.strip()

    response = re.sub(
        r"^```json\s*",
        "",
        response,
        flags=re.IGNORECASE
    )

    response = re.sub(
        r"^```\s*",
        "",
        response
    )

    response = re.sub(
        r"\s*```$",
        "",
        response
    )

    return response.strip()


def analyze_document(file_path: str) -> dict:
    """
    Extract and analyze an uploaded academic document.
    """

    extracted = extract_document(file_path)

    content = json.dumps(
        extracted["content"],
        ensure_ascii=False,
        indent=2
    )

    prompt = f"""
{DOCUMENT_AGENT_PROMPT}

UPLOADED DOCUMENT:

File Name:
{extracted["file_name"]}

File Type:
{extracted["file_type"]}

CONTENT:
{content}
"""

    response = generate_text(prompt)

    cleaned_response = clean_json_response(response)

    try:
        return json.loads(cleaned_response)

    except json.JSONDecodeError as error:

        return {
            "error": "Gemini returned invalid JSON.",
            "json_error": str(error),
            "raw_response": response
        }