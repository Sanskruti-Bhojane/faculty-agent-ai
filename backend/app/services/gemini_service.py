import os

from dotenv import load_dotenv
from google import genai

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY is missing from .env")

client = genai.Client(api_key=GEMINI_API_KEY)

MODEL_NAME = "gemini-3.6-flash"


def generate_text(prompt: str) -> str:
    interaction = client.interactions.create(
        model=MODEL_NAME,
        input=prompt
    )

    if not interaction.output_text:
        raise RuntimeError("Gemini returned an empty response.")

    return interaction.output_text.strip()