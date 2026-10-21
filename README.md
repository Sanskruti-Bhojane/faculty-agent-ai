# Faculty Agent – Agentic AI-Based Faculty Assistance System

An AI-powered faculty assistance system designed to simplify academic outcome mapping and attainment analysis through intelligent document processing, automated workflows, and data-driven reporting.

## Overview

Faculty Agent helps faculty members automate key academic tasks related to Outcome-Based Education (OBE). It processes faculty-provided academic documents, generates CO–PO–PSO mappings for review, and analyzes student performance to support attainment measurement and reporting.

## Key Features

- **Intelligent Document Processing:** Extracts relevant information from academic documents provided by faculty.
- **CO–PO–PSO Mapping:** Generates Course Outcome (CO), Program Outcome (PO), and Program Specific Outcome (PSO) mappings.
- **Faculty Approval Workflow:** Allows faculty to review and approve generated PSOs and outcome mappings before finalization.
- **Assessment Mapping:** Supports mapping assessment questions to relevant Course Outcomes.
- **Attainment Analysis:** Analyzes student marks and course exit survey responses using institutional guidelines.
- **Report Generation:** Supports downloadable academic mapping and attainment reports.

## Technology Stack

- **Programming Language:** Python
- **Backend Framework:** FastAPI
- **Database:** PostgreSQL
- **AI Integration:** Google Gemini API
- **Data Validation:** Pydantic
- **ORM:** SQLAlchemy
- **Document Processing:** PDF and Excel processing libraries
- **API Testing:** Automated Python test scripts

## Project Structure

```text
faculty-agent-ai/
├── backend/
│   └── app/
│       ├── agents/
│       ├── api/
│       ├── core/
│       ├── models/
│       ├── schemas/
│       ├── services/
│       ├── utils/
│       ├── attainment_data_extractor.py
│       └── main.py
├── data/
├── requirements.txt
├── .env.example
├── .gitignore
└── README.md
```

## Installation and Setup

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/faculty-agent-ai.git
cd faculty-agent-ai
```

Replace `YOUR_USERNAME` with your GitHub username.

### 2. Create a Virtual Environment

```bash
python -m venv venv
```

Activate it on Windows:

```bash
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables

Create a `.env` file in the project root and configure the environment variables required by the application.

Example:

```env
GEMINI_API_KEY=your_gemini_api_key
DATABASE_URL=your_postgresql_database_url
```

Use the exact variable names expected by your application. Never commit your actual API key or database credentials.

### 5. Run the Application

From the project root, start the FastAPI server using the appropriate application import path. For example, if `main.py` exposes the FastAPI instance as `app`:

```bash
uvicorn backend.app.main:app --reload
```

If your Python package structure or application entry point differs, adjust the command accordingly.

Once running, open:

- API: `http://127.0.0.1:8000`
- Interactive API documentation: `http://127.0.0.1:8000/docs`

## Use Cases

- Academic outcome mapping for Outcome-Based Education.
- Analysis of student assessment performance.
- CO, PO, and PSO attainment measurement.
- Identification of attainment gaps.
- Preparation of academic reports for faculty review.

## Future Scope

- Support for additional academic document formats.
- Enhanced analytics and attainment visualization.
- Expanded reporting and integration with institutional academic workflows.

## Project Status

Completed academic project developed to support faculty workflows for outcome mapping and attainment analysis.

## Author

**Sanskruti Bhojane**

B.Tech. – Artificial Intelligence and Data Science


---

This project was developed to support faculty members in academic outcome mapping and Outcome-Based Education (OBE) analysis.