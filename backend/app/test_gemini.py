from services.gemini_service import generate_text


prompt = """
You are an OBE assistant for a college Faculty Agent.

Explain Course Outcome (CO) in one simple sentence.
"""

answer = generate_text(prompt)

print("\n========== GEMINI RESPONSE ==========\n")
print(answer)
print("\n=====================================\n")