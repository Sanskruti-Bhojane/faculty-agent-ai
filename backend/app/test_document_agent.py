import json

from agents.document_agent import analyze_document


FILE_PATH = r"D:\FINAL YEAR\Final Year Project\faculty_agent_final_year\data\uploads\test.pdf"


result = analyze_document(FILE_PATH)

print("\n========== DOCUMENT AGENT RESULT ==========\n")

print(
    json.dumps(
        result,
        indent=4,
        ensure_ascii=False
    )
)

print("\n============================================\n")