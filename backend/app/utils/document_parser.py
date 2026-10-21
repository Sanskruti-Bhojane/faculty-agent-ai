from pathlib import Path
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook


def extract_text(file_path: str) -> str:
    """
    Extract text from PDF, DOCX, XLSX, XLS, and TXT files.
    """

    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    extension = path.suffix.lower()

    # ---------------- PDF ----------------
    if extension == ".pdf":
        reader = PdfReader(str(path))

        text = []

        for page in reader.pages:
            page_text = page.extract_text()

            if page_text:
                text.append(page_text)

        return "\n".join(text).strip()

    # ---------------- WORD ----------------
    elif extension == ".docx":
        document = Document(str(path))

        text = []

        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                text.append(paragraph.text)

        # Also read tables
        for table in document.tables:
            for row in table.rows:
                row_text = []

                for cell in row.cells:
                    row_text.append(cell.text.strip())

                text.append(" | ".join(row_text))

        return "\n".join(text).strip()

    # ---------------- EXCEL ----------------
    elif extension in [".xlsx", ".xlsm"]:
        workbook = load_workbook(
            filename=str(path),
            data_only=True
        )

        text = []

        for sheet in workbook.worksheets:

            text.append(f"\n--- SHEET: {sheet.title} ---")

            for row in sheet.iter_rows(values_only=True):

                values = []

                for value in row:
                    if value is not None:
                        values.append(str(value))

                if values:
                    text.append(" | ".join(values))

        return "\n".join(text).strip()

    # ---------------- TEXT ----------------
    elif extension == ".txt":
        return path.read_text(
            encoding="utf-8",
            errors="ignore"
        ).strip()

    # ---------------- UNSUPPORTED ----------------
    else:
        raise ValueError(
            f"Unsupported file format: {extension}. "
            f"Supported formats: PDF, DOCX, XLSX, XLSM, TXT"
        )