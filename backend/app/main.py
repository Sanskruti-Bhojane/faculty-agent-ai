import os
import json
import uuid
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from sqlalchemy import inspect

from core.database import Base, engine, SessionLocal
from models.mapping import MappingRecord

from agents.mapping_agent import generate_mapping
from agents.attainment_agent import analyze_assessments

from utils.extractor import extract_text

from attainment_data_extractor import (
    extract_student_marks,
    build_raw_assessment_data,
    extract_course_exit_survey,
)

from services.report_service import (
    generate_attainment_excel,
    generate_attainment_pdf,
)

from services.mapping_report_service import (
    generate_mapping_excel,
    generate_mapping_pdf,
)


app = FastAPI(
    title="Faculty Agent",
    description="Agentic AI System for CO-PO-PSO Mapping and Attainment",
    version="2.1.0",
)

Base.metadata.create_all(bind=engine)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(BASE_DIR))

UPLOAD_DIR = os.path.join(PROJECT_DIR, "data", "uploads")
EXTRACTED_DIR = os.path.join(PROJECT_DIR, "data", "extracted")
REPORT_DIR = os.path.join(PROJECT_DIR, "data", "reports")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXTRACTED_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {
    ".pdf", ".docx", ".xlsx", ".xls", ".xlsm",
    ".txt", ".csv", ".png", ".jpg", ".jpeg"
}


def save_uploaded_file(upload_file: UploadFile) -> str:
    original_name = upload_file.filename or "uploaded_file"
    extension = os.path.splitext(original_name)[1].lower()

    if extension not in ALLOWED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type: {extension}. "
            f"Allowed types: {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        )

    file_path = os.path.join(
        UPLOAD_DIR,
        f"{uuid.uuid4().hex}{extension}",
    )

    return file_path


async def save_file(upload_file: UploadFile) -> str:
    file_path = save_uploaded_file(upload_file)
    content = await upload_file.read()

    with open(file_path, "wb") as file:
        file.write(content)

    return file_path


def safe_filename(value: str) -> str:
    value = value or "course"
    return (
        value.replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace(":", "_")
    )


def error_response(message: str, status_code: int = 500):
    return JSONResponse(
        status_code=status_code,
        content={"success": False, "error": message},
    )


HOME_HTML = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Faculty Agent</title>
<style>
*{box-sizing:border-box}
body{
    margin:0;
    font-family:Arial,Helvetica,sans-serif;
    background:#f4f7fb;
    color:#172033;
}
.header{
    background:#17365d;
    color:white;
    padding:22px 34px;
}
.header h1{margin:0;font-size:30px}
.header p{margin:6px 0 0;color:#dbe8f7}
.container{max-width:1200px;margin:28px auto;padding:0 20px}
.card{
    background:white;
    border:1px solid #dce4ef;
    border-radius:12px;
    padding:24px;
    margin-bottom:22px;
    box-shadow:0 3px 12px rgba(20,40,80,.05);
}
h2{margin-top:0;color:#17365d}
.grid{
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:16px;
}
.field label{
    display:block;
    font-weight:700;
    margin-bottom:7px;
}
input[type=text],input[type=file],select{
    width:100%;
    padding:11px;
    border:1px solid #cbd5e1;
    border-radius:7px;
    background:white;
}
button,.btn{
    border:0;
    border-radius:7px;
    padding:11px 17px;
    cursor:pointer;
    background:#17365d;
    color:white;
    font-weight:700;
    text-decoration:none;
    display:inline-block;
}
button:hover,.btn:hover{opacity:.9}
.secondary{background:#64748b}
.success{
    background:#dcfce7;
    color:#166534;
    padding:14px 16px;
    border-radius:8px;
    margin-bottom:16px;
}
.error{
    background:#fee2e2;
    color:#991b1b;
    padding:14px 16px;
    border-radius:8px;
    margin-bottom:16px;
}
.loading{
    display:none;
    background:#fff7ed;
    color:#9a3412;
    padding:14px;
    border-radius:8px;
    margin-top:15px;
}
.hidden{display:none}
.actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:18px}
.table-wrap{overflow-x:auto}
table{
    width:100%;
    border-collapse:collapse;
    margin-top:12px;
    background:white;
}
th,td{
    border:1px solid #cbd5e1;
    padding:10px 12px;
    text-align:center;
    white-space:nowrap;
}
th{
    background:#e8eef6;
    font-weight:700;
}
td:first-child,th:first-child{text-align:left}
.status-achieved{color:#166534;font-weight:700}
.status-not{color:#b91c1c;font-weight:700}
.info-grid{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:12px;
}
.info{
    background:#f8fafc;
    border:1px solid #dbe3ee;
    border-radius:9px;
    padding:14px;
}
.info small{display:block;color:#64748b;margin-bottom:5px}
.info strong{font-size:17px}
.scale{
    display:flex;
    gap:12px;
    flex-wrap:wrap;
    margin-top:12px;
}
.scale span{
    background:#f1f5f9;
    border:1px solid #dbe3ee;
    border-radius:6px;
    padding:7px 12px;
}
@media(max-width:800px){
    .grid,.info-grid{grid-template-columns:1fr}
}
</style>
</head>

<body>
<div class="header">
    <h1>Faculty Agent</h1>
    <p>Agentic AI System for CO-PO-PSO Mapping and Attainment</p>
</div>

<div class="container">

<div class="card">
    <h2>Faculty Agent Modules</h2>
    <div class="actions">
        <a class="btn" href="#mapping">CO-PO-PSO Mapping</a>
        <a class="btn secondary" href="/attainment">CO-PO-PSO Attainment</a>
    </div>
</div>

<div class="card" id="mapping">
    <h2>Generate CO-PO-PSO Mapping</h2>
    <p>Upload the faculty-provided syllabus, CO, official PO and official PSO.</p>

    <form id="mappingForm">
        <div class="grid">
            <div class="field">
                <label>Course Code *</label>
                <input type="text" name="course_code" required>
            </div>
            <div class="field">
                <label>Course Name *</label>
                <input type="text" name="course_name" required>
            </div>
            <div class="field">
                <label>Syllabus *</label>
                <input type="file" name="syllabus" required>
            </div>
            <div class="field">
                <label>Course Outcomes (CO) *</label>
                <input type="file" name="co_file" required>
            </div>
            <div class="field">
                <label>Official Program Outcomes (PO) *</label>
                <input type="file" name="po_file" required>
            </div>
            <div class="field">
                <label>Program Specific Outcomes (PSO) *</label>
                <input type="file" name="pso_file" required>
            </div>
        </div>

        <div class="actions">
            <button type="submit">Generate Mapping</button>
        </div>
        <div id="mappingLoading" class="loading">
            AI Agent is analyzing the uploaded documents...
        </div>
    </form>
</div>

<div id="mappingResult" class="card hidden"></div>

</div>

<script>
function esc(value){
    if(value === null || value === undefined) return "";
    return String(value)
        .replaceAll("&","&amp;")
        .replaceAll("<","&lt;")
        .replaceAll(">","&gt;")
        .replaceAll('"',"&quot;");
}

function mappingRows(rows){
    if(!Array.isArray(rows) || !rows.length) return "";
    let codes = [];
    rows.forEach(r=>{
        const m = r && r.mappings ? r.mappings : {};
        Object.keys(m).forEach(k=>{
            if(!codes.includes(k)) codes.push(k);
        });
    });

    let html = "<table><thead><tr><th>CO</th>";
    codes.forEach(c=>html += "<th>"+esc(c)+"</th>");
    html += "</tr></thead><tbody>";

    rows.forEach(r=>{
        const co = r.co_code || r.co || "";
        const m = r.mappings || {};
        html += "<tr><td><b>"+esc(co)+"</b></td>";
        codes.forEach(c=>{
            html += "<td>"+esc(m[c] ?? "")+"</td>";
        });
        html += "</tr>";
    });

    html += "</tbody></table>";
    return html;
}

document.getElementById("mappingForm").addEventListener("submit", async e=>{
    e.preventDefault();

    const loading = document.getElementById("mappingLoading");
    const box = document.getElementById("mappingResult");
    loading.style.display = "block";
    box.classList.add("hidden");

    try{
        const response = await fetch("/generate-mapping",{
            method:"POST",
            body:new FormData(e.target)
        });
        const data = await response.json();

        if(!response.ok || data.success === false){
            throw new Error(data.error || "Mapping generation failed.");
        }

        const coPO = mappingRows(data.co_po_mapping);
        const coPSO = mappingRows(data.co_pso_mapping);

        box.innerHTML = `
            <h2>Generated Mapping</h2>
            <div class="success">Mapping generated successfully.</div>

            <div class="info-grid">
                <div class="info"><small>Mapping ID</small><strong>${esc(data.mapping_id)}</strong></div>
                <div class="info"><small>Approval Status</small><strong>${esc(data.approval_status)}</strong></div>
                <div class="info"><small>Course Code</small><strong>${esc(data.course_code || "")}</strong></div>
                <div class="info"><small>Course Name</small><strong>${esc(data.course_name || "")}</strong></div>
            </div>

            <h3>CO-PO Mapping</h3>${coPO || "<p>No CO-PO mapping returned.</p>"}
            <h3>CO-PSO Mapping</h3>${coPSO || "<p>No CO-PSO mapping returned.</p>"}

            <div class="actions">
                <button onclick="approveMapping(${Number(data.mapping_id)})">Approve Mapping</button>
            </div>

            <div
                id="mappingDownloads"
                class="actions"
                style="display:none;"
            >
                <a
                    class="btn"
                    href="/download-mapping-excel/${Number(data.mapping_id)}"
                >
                    Download Excel
                </a>

                <a
                    class="btn"
                    href="/download-mapping-pdf/${Number(data.mapping_id)}"
                >
                    Download PDF
                </a>
            </div>
        `;

        box.classList.remove("hidden");
    }catch(err){
        box.innerHTML = `<div class="error">${esc(err.message)}</div>`;
        box.classList.remove("hidden");
    }finally{
        loading.style.display = "none";
    }
});

async function approveMapping(id){
    const button = document.querySelector(
        `button[onclick="approveMapping(${id})"]`
    );

    if(button){
        button.disabled = true;
        button.innerText = "Approving...";
    }

    try{
        const response = await fetch(
            "/approve-mapping/" + id,
            {
                method: "POST"
            }
        );

        const data = await response.json();

        if(!response.ok || data.success === false){
            throw new Error(
                data.error || "Approval failed."
            );
        }

        // IMPORTANT:
        // Do NOT reload the page here.
        // Keep the generated mapping visible so the
        // faculty can immediately download it.

        const statusElements =
            document.querySelectorAll(
                ".info strong"
            );

        // Update the visible approval status.
        statusElements.forEach(el=>{
            if(el.textContent.trim() === "PENDING_APPROVAL"){
                el.textContent = "APPROVED";
            }
        });

        if(button){
            button.innerText = "Mapping Approved";
        }

        // Show download buttons after approval.
        const downloadArea =
            document.getElementById("mappingDownloads");

        if(downloadArea){
            downloadArea.style.display = "flex";
        }

        alert(
            "Mapping approved successfully. " +
            "You can now download the CO-PO-PSO Mapping."
        );

    }catch(error){

        if(button){
            button.disabled = false;
            button.innerText = "Approve Mapping";
        }

        alert(error.message);
    }
}
</script>
</body>
</html>
"""


ATTAINMENT_HTML = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Faculty Agent - CO Attainment</title>

<style>
*{box-sizing:border-box}

body{
    margin:0;
    font-family:Arial,Helvetica,sans-serif;
    background:#f4f7fb;
    color:#172033;
}

.header{
    background:#17365d;
    color:white;
    padding:22px 34px;
}

.header h1{
    margin:0;
    font-size:30px;
}

.header p{
    margin:6px 0 0;
    color:#dbe8f7;
}

.container{
    max-width:1450px;
    margin:28px auto;
    padding:0 20px;
}

.card{
    background:white;
    border:1px solid #dce4ef;
    border-radius:12px;
    padding:24px;
    margin-bottom:22px;
    box-shadow:0 3px 12px rgba(20,40,80,.05);
}

h2{
    color:#17365d;
    margin:0 0 18px;
}

h3{
    color:#17365d;
    margin-top:28px;
}

.grid{
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:16px;
}

.field label{
    display:block;
    font-weight:700;
    margin-bottom:7px;
}

input[type=text],
input[type=file]{
    width:100%;
    padding:11px;
    border:1px solid #cbd5e1;
    border-radius:7px;
    background:white;
}

button,.btn{
    border:0;
    border-radius:7px;
    padding:11px 17px;
    cursor:pointer;
    background:#17365d;
    color:white;
    font-weight:700;
    text-decoration:none;
    display:inline-block;
}

.secondary{
    background:#64748b;
}

.actions{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    margin-top:18px;
}

.success{
    background:#dcfce7;
    color:#166534;
    padding:14px;
    border-radius:8px;
    margin-bottom:18px;
}

.error{
    background:#fee2e2;
    color:#991b1b;
    padding:14px;
    border-radius:8px;
    margin-bottom:18px;
}

.loading{
    display:none;
    background:#fff7ed;
    color:#9a3412;
    padding:14px;
    border-radius:8px;
    margin-top:15px;
}

.hidden{
    display:none;
}

.info-grid{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:12px;
    margin-bottom:24px;
}

.info{
    background:#f8fafc;
    border:1px solid #dbe3ee;
    border-radius:9px;
    padding:14px;
}

.info small{
    display:block;
    color:#64748b;
    margin-bottom:5px;
}

.info strong{
    font-size:17px;
}

.table-wrap{
    overflow-x:auto;
    margin:12px 0 25px;
}

table{
    width:100%;
    border-collapse:collapse;
    background:white;
}

th,td{
    border:1px solid #cbd5e1;
    padding:9px 11px;
    text-align:center;
    white-space:nowrap;
}

th{
    background:#e8eef6;
    font-weight:700;
}

td:first-child,
th:first-child{
    text-align:left;
}

.subtle{
    color:#64748b;
    font-size:14px;
}

.achieved{
    color:#166534;
    font-weight:700;
}

.not-achieved{
    color:#b91c1c;
    font-weight:700;
}

.section-title{
    border-bottom:2px solid #17365d;
    padding-bottom:8px;
    margin-top:34px;
}

@media(max-width:900px){
    .grid,.info-grid{
        grid-template-columns:1fr;
    }
}
</style>
</head>

<body>

<div class="header">
    <h1>Faculty Agent</h1>
    <p>Course Outcome (CO) - Attainment</p>
</div>

<div class="container">

<div class="card">

    <h2>Generate CO-PO-PSO Attainment</h2>

    <p class="subtle">
        The Student Marks Excel is authoritative for assessment → CO mapping,
        maximum marks and student marks. Thresholds are calculated automatically.
    </p>

    <form id="attainmentForm">

        <div class="grid">

            <div class="field">
                <label>Course Code *</label>
                <input type="text" name="course_code" required>
            </div>

            <div class="field">
                <label>Course Name *</label>
                <input type="text" name="course_name" required>
            </div>

            <div class="field">
                <label>Semester *</label>
                <input type="text" name="semester" required>
            </div>

            <div class="field">
                <label>Academic Year *</label>
                <input type="text" name="academic_year" required>
            </div>

            <div class="field">
                <label>Syllabus *</label>
                <input type="file" name="syllabus" required>
            </div>

            <div class="field">
                <label>Course Outcomes (CO) *</label>
                <input type="file" name="co" required>
            </div>

            <div class="field">
                <label>Official Program Outcomes (PO) *</label>
                <input type="file" name="po" required>
            </div>

            <div class="field">
                <label>Program Specific Outcomes (PSO) *</label>
                <input type="file" name="pso" required>
            </div>

            <div class="field">
                <label>Student Marks Excel *</label>
                <input type="file" name="student_marks" required>
            </div>

            <div class="field">
                <label>Course Exit Survey *</label>
                <input type="file" name="course_exit_survey" required>
            </div>

        </div>

        <div class="actions">
            <button type="submit">Generate Attainment</button>
            <a class="btn secondary" href="/">Back to Mapping</a>
        </div>

        <div id="loading" class="loading">
            Attainment Agent is processing the uploaded data...
        </div>

    </form>

</div>

<div id="result" class="card hidden"></div>

</div>

<script>

function esc(value){
    if(value === null || value === undefined) return "";
    return String(value)
        .replaceAll("&","&amp;")
        .replaceAll("<","&lt;")
        .replaceAll(">","&gt;")
        .replaceAll('"',"&quot;");
}

function value(obj, keys, fallback=""){
    if(!obj || typeof obj !== "object") return fallback;

    for(const key of keys){
        if(obj[key] !== undefined && obj[key] !== null && obj[key] !== ""){
            return obj[key];
        }
    }

    return fallback;
}

function fmt(v){
    if(v === null || v === undefined || v === ""){
        return "";
    }

    const n = Number(v);

    if(Number.isFinite(n)){
        return Number.isInteger(n) ? String(n) : n.toFixed(2);
    }

    return esc(v);
}

function section(title){
    return `<h2 class="section-title">${esc(title)}</h2>`;
}


/* ============================================================
   COURSE DETAILS
   ============================================================ */

function renderCourse(result){

    const course = result.course || {};

    return `
        <div class="info-grid">

            <div class="info">
                <small>Course Code</small>
                <strong>${esc(course.course_code || "")}</strong>
            </div>

            <div class="info">
                <small>Course Name</small>
                <strong>${esc(course.course_name || "")}</strong>
            </div>

            <div class="info">
                <small>Semester</small>
                <strong>${esc(course.semester || "")}</strong>
            </div>

            <div class="info">
                <small>Academic Year</small>
                <strong>${esc(course.academic_year || "")}</strong>
            </div>

        </div>
    `;
}


/* ============================================================
   ASSESSMENT ATTAINMENT
   ============================================================ */

function renderAssessmentAttainment(result){

    const stats = result.assessment_statistics || {};
    const mapping = result.assessment_mapping || {};
    const assessments = result.assessments || [];
    const ese = result.ese_attainment || {};

    const names = [];

    Object.keys(stats).forEach(name=>{
        if(!names.includes(name)) names.push(name);
    });

    Object.keys(mapping).forEach(name=>{
        if(!names.includes(name)) names.push(name);
    });

    assessments.forEach(item=>{
        const name = item.name || item.assessment;
        if(name && !names.includes(name)) names.push(name);
    });

    // Always include ESE when the deterministic calculator produced it.
    const eseName = names.find(name => {
        const n = String(name || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        return n.includes("ese") || n.includes("endsem") || n.includes("endsemester");
    });

    if(Object.keys(ese).length && !eseName){
        names.push("ESE");
    }

    if(!names.length){
        return `
            ${section("Assessment Attainment")}
            <p class="subtle">
                No assessment statistics were returned by the attainment agent.
            </p>
        `;
    }

    let rows = "";

    names.forEach(name=>{

        const isESE = String(name || "").toLowerCase().replace(/[^a-z0-9]/g, "").includes("ese") ||
                      String(name || "").toLowerCase().replace(/[^a-z0-9]/g, "").includes("endsem");

        const assessmentStats = isESE && Object.keys(ese).length
            ? ese
            : (stats[name] || {});

        const mapped = mapping[name] || [];

        let cos = [];

        if(Array.isArray(mapped)){
            cos = mapped.slice();
        }else if(mapped && typeof mapped === "object"){
            cos = Object.keys(mapped);
        }

        Object.keys(assessmentStats).forEach(co=>{
            if(!cos.includes(co)) cos.push(co);
        });

        // ESE is a single total column mapped to all COs. Show one clean
        // CO1-CO6 row instead of repeating identical statistics six times.
        if(isESE && Object.keys(assessmentStats).length){
            const firstCO = Object.keys(assessmentStats)[0];
            const stat = assessmentStats[firstCO] || {};

            rows += `
                <tr>
                    <td>${esc(name)}</td>
                    <td>CO1-CO6</td>
                    <td>${fmt(value(stat,["threshold","threshold_marks","threshold_mark"],0))}</td>
                    <td>${fmt(value(stat,["total_students","total","students"],0))}</td>
                    <td>${fmt(value(stat,["above_threshold","students_above_threshold","count_above_threshold"],0))}</td>
                    <td>${fmt(value(stat,["percentage","percent"],0))}</td>
                    <td>${fmt(value(stat,["level","attainment_level","level_attained"],0))}</td>
                </tr>
            `;
            return;
        }

        if(!cos.length){
            rows += `
                <tr>
                    <td>${esc(name)}</td>
                    <td colspan="6">No CO statistics available</td>
                </tr>
            `;
            return;
        }

        cos.forEach(co=>{

            const stat = assessmentStats[co] || {};

            const threshold = value(
                stat,
                ["threshold","threshold_marks","threshold_mark"]
            );

            const total = value(
                stat,
                ["total_students","total","students"]
            );

            const above = value(
                stat,
                [
                    "above_threshold",
                    "students_above_threshold",
                    "count_above_threshold"
                ]
            );

            const percentage = value(
                stat,
                ["percentage","percent"]
            );

            const level = value(
                stat,
                ["level","attainment_level","level_attained"]
            );

            rows += `
                <tr>
                    <td>${esc(name)}</td>
                    <td>${esc(co)}</td>
                    <td>${fmt(threshold,0)}</td>
                    <td>${fmt(total,0)}</td>
                    <td>${fmt(above,0)}</td>
                    <td>${fmt(percentage,0)}</td>
                    <td>${fmt(level,0)}</td>
                </tr>
            `;
        });
    });

    return `
        ${section("Assessment Attainment")}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>
                    <th>Assessment</th>
                    <th>CO</th>
                    <th>Threshold Marks</th>
                    <th>Total No. of Students</th>
                    <th>No. of students above threshold</th>
                    <th>Percentage</th>
                    <th>Level Attained</th>
                </tr>
            </thead>

            <tbody>
                ${rows}
            </tbody>

        </table>
        </div>
    `;
}

/* ============================================================
   INDIRECT ATTAINMENT
   ============================================================ */

function renderIndirect(result){

    const details = result.survey_details || {};
    const ratings = result.survey_ratings || {};
    const indirect = result.indirect_co_attainment || {};

    let rows = "";

    const codes = result.co_codes || Object.keys(indirect);

    codes.forEach(co=>{

        const r = indirect[co] || {};
        const surveyRating = value(
            r,
            ["survey_rating","rating"]
        );

        const indirectValue = value(
            r,
            ["indirect_attainment","attainment"]
        );

        rows += `
            <tr>
                <td>${esc(co)}</td>
                <td>${fmt(surveyRating)}</td>
                <td>${fmt(indirectValue)}</td>
            </tr>
        `;
    });

    const overall = details.overall_percentage;

    return `
        ${section("Indirect Attainment (Course Exit Survey)")}

        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>CO</th>
                    <th>Survey Rating</th>
                    <th>Indirect Attainment</th>
                </tr>
            </thead>
            <tbody>
                ${rows}
            </tbody>
        </table>
        </div>

        <p>
            <b>Overall Course Exit Survey Percentage:</b>
            ${fmt(overall)}
        </p>
    `;
}


/* ============================================================
   CO ATTAINMENT
   ============================================================ */

function renderCOAttainment(result){

    const data =
        result.overall_co_attainment ||
        result.co_attainment ||
        {};

    const codes =
        result.co_codes ||
        Object.keys(data);

    let rows = "";

    codes.forEach(co=>{

        const r = data[co] || {};

        rows += `
            <tr>
                <td><b>${esc(co)}</b></td>
                <td>${fmt(value(r,["indirect"]))}</td>
                <td>${fmt(value(r,["direct","direct_ia"]))}</td>
                <td>${fmt(value(r,["overall_ia"]))}</td>
                <td>${fmt(value(r,["overall","overall_co","overall_attainment"]))}</td>
            </tr>
        `;
    });

    return `
        ${section("CO Attainment")}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>
                    <th>CO</th>
                    <th>Indirect</th>
                    <th>Direct</th>
                    <th>Overall IA</th>
                    <th>O.A. (I+E)</th>
                </tr>
            </thead>

            <tbody>
                ${rows}
            </tbody>

        </table>
        </div>
    `;
}

/* ============================================================
   OVERALL CO
   ============================================================ */

function renderOverallCO(result){

    const data =
        result.overall_co_attainment ||
        result.co_attainment ||
        {};

    const codes =
        result.co_codes ||
        Object.keys(data);

    let rows = "";

    codes.forEach(co=>{

        const r = data[co] || {};

        const status = value(r,["status"]);

        rows += `
            <tr>

                <td><b>${esc(co)}</b></td>

                <td>${fmt(value(
                    r,
                    ["overall","overall_attainment","overall_co"]
                ))}</td>

                <td>${fmt(value(
                    r,
                    ["target"]
                ))}</td>

                <td>${fmt(value(
                    r,
                    ["gap"]
                ))}</td>

                <td class="${
                    status === "ACHIEVED"
                    ? "achieved"
                    : "not-achieved"
                }">
                    ${esc(status)}
                </td>

            </tr>
        `;
    });

    return `
        ${section("Overall CO Attained")}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>
                    <th>CO</th>
                    <th>Overall CO Attained</th>
                    <th>CO Target</th>
                    <th>Gap</th>
                    <th>CO Attain or Not</th>
                </tr>
            </thead>

            <tbody>
                ${rows}
            </tbody>

        </table>
        </div>
    `;
}


/* ============================================================
   ARTICULATION MATRIX
   ============================================================ */

function renderMatrix(title, matrix, outcomeCodes){

    matrix = matrix || {};

    const coCodes =
        Object.keys(matrix).length
        ? Object.keys(matrix)
        : [];

    const outcomes =
        outcomeCodes && outcomeCodes.length
        ? outcomeCodes
        : Array.from(
            new Set(
                Object.values(matrix).flatMap(row=>{
                    if(!row || typeof row !== "object") return [];
                    return Object.keys(row);
                })
            )
        );

    if(!coCodes.length || !outcomes.length){
        return `
            ${section(title)}
            <p class="subtle">No matrix data available.</p>
        `;
    }

    let head = `<th>CO</th>`;

    outcomes.forEach(po=>{
        head += `<th>${esc(po)}</th>`;
    });

    let body = "";

    coCodes.forEach(co=>{

        const row = matrix[co] || {};

        body += `<tr><td><b>${esc(co)}</b></td>`;

        outcomes.forEach(po=>{
            body += `<td>${fmt(row[po])}</td>`;
        });

        body += `</tr>`;
    });

    return `
        ${section(title)}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>${head}</tr>
            </thead>

            <tbody>
                ${body}
            </tbody>

        </table>
        </div>
    `;
}


/* ============================================================
   PO / PSO ATTAINMENT
   ============================================================ */

function renderOutcomeAttainment(title, data, codes){

    data = data || {};

    const outcomeCodes =
        codes && codes.length
        ? codes
        : Object.keys(data);

    let rows = "";

    outcomeCodes.forEach(code=>{

        const r = data[code] || {};

        rows += `
            <tr>
                <td><b>${esc(code)}</b></td>
                <td>${fmt(value(
                    r,
                    ["attainment","direct_attainment","direct"]
                ))}</td>
                <td>${fmt(value(
                    r,
                    ["indirect_attainment","indirect"]
                ))}</td>
                <td>${fmt(value(
                    r,
                    ["overall_attainment","overall"]
                ))}</td>
            </tr>
        `;
    });

    return `
        ${section(title)}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>
                    <th>${title.includes("PSO") ? "PSO" : "PO"}</th>
                    <th>Direct</th>
                    <th>Indirect</th>
                    <th>Attained</th>
                </tr>
            </thead>

            <tbody>
                ${rows}
            </tbody>

        </table>
        </div>
    `;
}


/* ============================================================
   METHODOLOGY
   ============================================================ */

function renderMethodology(result){

    const m = result.methodology_used || {};

    const fields = [
        ["IA Threshold", "ia_threshold_percent", "%"],
        ["ESE Threshold", "ese_threshold_percent", "%"],
        ["Level 3 Percentage", "level_3_percent", "%"],
        ["Level 2 Percentage", "level_2_percent", "%"],
        ["Direct IA Weight", "overall_ia_direct_weight", ""],
        ["Indirect IA Weight", "overall_ia_indirect_weight", ""],
        ["Overall IA Weight", "final_ia_weight", ""],
        ["ESE Weight", "final_ese_weight", ""],
        ["CO Target", "co_target", ""]
    ];

    let rows = "";

    fields.forEach(([label,key,suffix])=>{

        let v = m[key];

        if(
            key.includes("weight") &&
            typeof v === "number" &&
            v <= 1
        ){
            v = v * 100;
            suffix = "%";
        }

        rows += `
            <tr>
                <td><b>${esc(label)}</b></td>
                <td>${fmt(v)}${suffix}</td>
            </tr>
        `;
    });

    return `
        ${section("Methodology and Gap Analysis")}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>
                    <th>Parameter</th>
                    <th>Value</th>
                </tr>
            </thead>

            <tbody>
                ${rows}
            </tbody>

        </table>
        </div>
    `;
}


/* ============================================================
   CO GAP ANALYSIS
   ============================================================ */

function renderGapAnalysis(result){

    const data =
        result.overall_co_attainment ||
        result.co_attainment ||
        {};

    const codes =
        result.co_codes ||
        Object.keys(data);

    let rows = "";

    codes.forEach(co=>{

        const r = data[co] || {};
        const status = value(r,["status"]);

        rows += `
            <tr>

                <td><b>${esc(co)}</b></td>

                <td>${fmt(value(
                    r,
                    ["overall","overall_attainment","overall_co"]
                ))}</td>

                <td>${fmt(value(
                    r,
                    ["target"]
                ))}</td>

                <td>${fmt(value(
                    r,
                    ["gap"]
                ))}</td>

                <td class="${
                    status === "ACHIEVED"
                    ? "achieved"
                    : "not-achieved"
                }">
                    ${esc(status)}
                </td>

            </tr>
        `;
    });

    return `
        ${section("CO Gap Analysis")}

        <div class="table-wrap">
        <table>

            <thead>
                <tr>
                    <th>CO</th>
                    <th>Actual</th>
                    <th>Target</th>
                    <th>Gap</th>
                    <th>Status</th>
                </tr>
            </thead>

            <tbody>
                ${rows}
            </tbody>

        </table>
        </div>
    `;
}


/* ============================================================
   FULL RESULT
   ============================================================ */

function renderResult(result){

    return `

        <h2>Course Outcome (CO) - Attainment</h2>

        ${renderCourse(result)}

        ${renderAssessmentAttainment(result)}

        ${renderIndirect(result)}

        ${renderCOAttainment(result)}

        ${renderOverallCO(result)}

        ${renderMatrix(
            "CO-PO Articulation Matrix",
            result.co_po_articulation_matrix,
            result.po_codes
        )}

        ${renderMatrix(
            "CO-PSO Articulation Matrix",
            result.co_pso_articulation_matrix,
            result.pso_codes
        )}

        ${renderOutcomeAttainment(
            "CO-PO Attainment Matrix",
            result.po_attainment || result.co_po_attainment,
            result.po_codes
        )}

        ${renderOutcomeAttainment(
            "CO-PSO Attainment Matrix",
            result.pso_attainment || result.co_pso_attainment,
            result.pso_codes
        )}

        ${renderMethodology(result)}

        ${renderGapAnalysis(result)}

        <div class="actions">

            <a
                class="btn"
                href="${esc(result._reports?.excel || "#")}"
            >
                Download Excel
            </a>

            <a
                class="btn"
                href="${esc(result._reports?.pdf || "#")}"
            >
                Download PDF
            </a>

        </div>
    `;
}


/* ============================================================
   FORM SUBMISSION
   ============================================================ */

document
.getElementById("attainmentForm")
.addEventListener(
    "submit",
    async function(event){

        event.preventDefault();

        const loading =
            document.getElementById("loading");

        const resultBox =
            document.getElementById("result");

        loading.style.display = "block";
        resultBox.classList.add("hidden");

        try{

            const response =
                await fetch(
                    "/generate-attainment",
                    {
                        method:"POST",
                        body:new FormData(event.target)
                    }
                );

            const data =
                await response.json();

            if(
                !response.ok ||
                data.success === false
            ){
                throw new Error(
                    data.error ||
                    "Attainment generation failed."
                );
            }

            const result =
                data.result || {};

            result._reports =
                data.reports || {};

            resultBox.innerHTML = `
                <div class="success">
                    ${esc(
                        data.message ||
                        "Attainment Agent completed successfully."
                    )}
                </div>

                ${renderResult(result)}
            `;

            resultBox.classList.remove("hidden");

            resultBox.scrollIntoView({
                behavior:"smooth"
            });

        }catch(error){

            resultBox.innerHTML = `
                <div class="error">
                    ${esc(error.message)}
                </div>
            `;

            resultBox.classList.remove("hidden");

        }finally{

            loading.style.display = "none";
        }
    }
);

</script>

</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def home():
    return HOME_HTML


@app.get("/attainment", response_class=HTMLResponse)
async def attainment_page():
    return ATTAINMENT_HTML


# ============================================================
# MAPPING
# ============================================================

@app.post("/generate-mapping")
async def generate_mapping_endpoint(
    course_code: str = Form(...),
    course_name: str = Form(...),
    syllabus: UploadFile = File(...),
    co_file: UploadFile = File(...),
    po_file: UploadFile = File(...),
    pso_file: UploadFile = File(...),
):
    db = SessionLocal()

    try:
        syllabus_path = await save_file(syllabus)
        co_path = await save_file(co_file)
        po_path = await save_file(po_file)
        pso_path = await save_file(pso_file)

        mapping_result = generate_mapping(
            extract_text(syllabus_path),
            extract_text(co_path),
            extract_text(po_path),
            extract_text(pso_path),
        )

        if isinstance(mapping_result, str):
            mapping_result = json.loads(mapping_result)

        if not isinstance(mapping_result, dict):
            raise ValueError("Mapping agent returned an invalid structure.")

        record = MappingRecord(
            course_code=course_code,
            course_name=course_name,
            mapping_json=json.dumps(mapping_result),
            status="PENDING_APPROVAL",
        )

        db.add(record)
        db.commit()
        db.refresh(record)

        mapping_result["mapping_id"] = record.id
        mapping_result["approval_status"] = record.status
        mapping_result["course_code"] = course_code
        mapping_result["course_name"] = course_name

        return mapping_result

    except Exception as error:
        db.rollback()
        return error_response(str(error), 500)
    finally:
        db.close()


@app.post("/approve-mapping/{mapping_id}")
async def approve_mapping(mapping_id: int):
    db = SessionLocal()

    try:
        mapping = (
            db.query(MappingRecord)
            .filter(MappingRecord.id == mapping_id)
            .first()
        )

        if not mapping:
            return error_response("Mapping record not found.", 404)

        mapping.status = "APPROVED"
        mapping.approved_at = datetime.utcnow()

        db.commit()
        db.refresh(mapping)

        return {
            "message": "Mapping approved successfully.",
            "mapping_id": mapping.id,
            "status": mapping.status,
            "approved_at": mapping.approved_at.isoformat(),
        }

    except Exception as error:
        db.rollback()
        return error_response(str(error), 500)
    finally:
        db.close()


@app.get("/download-mapping-excel/{mapping_id}")
async def download_mapping_excel(mapping_id: int):
    db = SessionLocal()

    try:
        mapping = (
            db.query(MappingRecord)
            .filter(MappingRecord.id == mapping_id)
            .first()
        )

        if not mapping:
            return error_response("Mapping record not found.", 404)

        if mapping.status != "APPROVED":
            return error_response(
                "Only approved mappings can be downloaded.",
                403,
            )

        data = json.loads(mapping.mapping_json)
        data.update({
            "mapping_id": mapping.id,
            "course_code": mapping.course_code,
            "course_name": mapping.course_name,
            "approval_status": mapping.status,
        })

        filename = (
            safe_filename(mapping.course_code)
            + "_CO_PO_PSO_Mapping.xlsx"
        )
        path = os.path.join(REPORT_DIR, filename)

        generate_mapping_excel(data, path)

        return FileResponse(
            path,
            filename=filename,
            media_type=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception as error:
        import traceback

        traceback.print_exc()

        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": str(error),
                "traceback": traceback.format_exc(),
            },
        )
    finally:
        db.close()


@app.get("/download-mapping-pdf/{mapping_id}")
async def download_mapping_pdf(mapping_id: int):
    db = SessionLocal()

    try:
        mapping = (
            db.query(MappingRecord)
            .filter(MappingRecord.id == mapping_id)
            .first()
        )

        if not mapping:
            return error_response("Mapping record not found.", 404)

        if mapping.status != "APPROVED":
            return error_response(
                "Only approved mappings can be downloaded.",
                403,
            )

        data = json.loads(mapping.mapping_json)
        data.update({
            "mapping_id": mapping.id,
            "course_code": mapping.course_code,
            "course_name": mapping.course_name,
            "approval_status": mapping.status,
        })

        filename = (
            safe_filename(mapping.course_code)
            + "_CO_PO_PSO_Mapping.pdf"
        )
        path = os.path.join(REPORT_DIR, filename)

        generate_mapping_pdf(data, path)

        return FileResponse(
            path,
            filename=filename,
            media_type="application/pdf",
        )

    except Exception as error:
        return error_response(str(error), 500)
    finally:
        db.close()


@app.get("/mapping/{mapping_id}")
async def get_mapping(mapping_id: int):
    db = SessionLocal()

    try:
        mapping = (
            db.query(MappingRecord)
            .filter(MappingRecord.id == mapping_id)
            .first()
        )

        if not mapping:
            return error_response("Mapping record not found.", 404)

        data = json.loads(mapping.mapping_json)
        data.update({
            "mapping_id": mapping.id,
            "course_code": mapping.course_code,
            "course_name": mapping.course_name,
            "approval_status": mapping.status,
            "approved_at": (
                mapping.approved_at.isoformat()
                if mapping.approved_at
                else None
            ),
        })

        return data

    except Exception as error:
        return error_response(str(error), 500)
    finally:
        db.close()


# ============================================================
# ATTAINMENT
# ============================================================

@app.post("/generate-attainment")
async def generate_attainment(
    course_code: str = Form(...),
    course_name: str = Form(...),
    semester: str = Form(...),
    academic_year: str = Form(...),
    syllabus: UploadFile = File(...),
    co: UploadFile = File(...),
    po: UploadFile = File(...),
    pso: UploadFile = File(...),
    student_marks: UploadFile = File(...),
    course_exit_survey: UploadFile = File(...),
):
    try:
        syllabus_path = await save_file(syllabus)
        co_path = await save_file(co)
        po_path = await save_file(po)
        pso_path = await save_file(pso)
        marks_path = await save_file(student_marks)
        survey_path = await save_file(course_exit_survey)

        syllabus_text = extract_text(syllabus_path)
        co_text = extract_text(co_path)
        po_text = extract_text(po_path)
        pso_text = extract_text(pso_path)
        marks_text = extract_text(marks_path)
        survey_text = extract_text(survey_path)

        extracted_marks = extract_student_marks(marks_path)
        if not extracted_marks:
            raise ValueError(
                "Could not read student marks from the uploaded Student Marks Excel."
            )

        raw_marks = build_raw_assessment_data(extracted_marks)
        if not raw_marks:
            raise ValueError(
                "Could not detect assessment → CO mapping, maximum marks "
                "and student marks from the Student Marks Excel."
            )

        survey_data = extract_course_exit_survey(survey_path)
        if not survey_data:
            raise ValueError(
                "Could not detect numbered CO questions and survey responses "
                "from the Course Exit Survey."
            )

        # The extractor may return either the new structured survey format
        # or the older flat CO -> rating format. Pass it unchanged.
        result = analyze_assessments(
            syllabus_text=syllabus_text,
            co_text=co_text,
            po_text=po_text,
            pso_text=pso_text,
            assessment_text=(
                "STUDENT MARKS EXCEL:\n"
                + marks_text
                + "\n\nCOURSE EXIT SURVEY:\n"
                + survey_text
            ),
            raw_marks_data=raw_marks,
            survey_ratings=survey_data,
            course_code=course_code,
            course_name=course_name,
            semester=semester,
            academic_year=academic_year,
        )

        if not isinstance(result, dict):
            raise ValueError("Attainment Agent returned an invalid result.")

        # Defensive compatibility: keep assessment statistics available
        # for the new attainment-result page even if an older agent result
        # omits the field. The actual calculations remain in the attainment agent.
        result["assessment_statistics"] = result.get("assessment_statistics", {})

        result["course"] = {
            "course_code": course_code,
            "course_name": course_name,
            "semester": semester,
            "academic_year": academic_year,
        }

        report_id = uuid.uuid4().hex
        excel_filename = f"attainment_report_{report_id}.xlsx"
        pdf_filename = f"attainment_report_{report_id}.pdf"

        excel_path = os.path.join(REPORT_DIR, excel_filename)
        pdf_path = os.path.join(REPORT_DIR, pdf_filename)

        generate_attainment_excel(result, excel_path)
        generate_attainment_pdf(result, pdf_path)

        return {
            "success": True,
            "message": "Attainment Agent completed successfully.",
            "course_code": course_code,
            "course_name": course_name,
            "report_id": report_id,
            "result": result,
            "reports": {
                "excel": f"/download-attainment-excel/{report_id}",
                "pdf": f"/download-attainment-pdf/{report_id}",
            },
        }

    except Exception as error:
        return error_response(str(error), 500)


@app.get("/download-attainment-excel/{report_id}")
async def download_attainment_excel(report_id: str):
    filename = f"attainment_report_{report_id}.xlsx"
    path = os.path.join(REPORT_DIR, filename)

    if not os.path.exists(path):
        return error_response(
            "Attainment Excel report not found.",
            404,
        )

    return FileResponse(
        path,
        filename=filename,
        media_type=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


@app.get("/download-attainment-pdf/{report_id}")
async def download_attainment_pdf(report_id: str):
    filename = f"attainment_report_{report_id}.pdf"
    path = os.path.join(REPORT_DIR, filename)

    if not os.path.exists(path):
        return error_response(
            "Attainment PDF report not found.",
            404,
        )

    return FileResponse(
        path,
        filename=filename,
        media_type="application/pdf",
    )


@app.get("/health")
async def health():
    try:
        inspector = inspect(engine)
        return {
            "status": "healthy",
            "database": "connected",
            "tables": inspector.get_table_names(),
        }
    except Exception as error:
        import traceback

        print("\n" + "=" * 80)
        print("GENERATE ATTAINMENT ERROR")
        print("=" * 80)
        print("ERROR:", str(error))
        print("TRACEBACK:")
        print(traceback.format_exc())
        print("=" * 80 + "\n")

        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "error": str(error),
                "traceback": traceback.format_exc(),
            },
        )
