"""Faculty Agent - authoritative marks/survey Excel extractor.

Supports the faculty's actual multi-row marksheet format:
row 1: assessment group
row 2: assessment name
row 3: CO mapping
row 4: maximum marks
row 5+: students

Important: repeated CO columns are NOT deduplicated. For example
MSE-1 -> CO1, CO1, CO3 becomes CO1 max=20 and each student's CO1 mark is
sum of the two CO1 columns.
"""
import os
import re
from typing import Any, Dict, List
import pandas as pd

LIKERT = {
    "strongly disagree": 1, "strongly dis agree": 1,
    "disagree": 2, "neutral": 3,
    "neither agree nor disagree": 3,
    "agree": 4, "strongly agree": 5,
}

def clean_text(v: Any) -> str:
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except Exception: pass
    s = str(v).strip()
    return "" if s.lower() in {"nan","none","null"} else s

def is_numeric(v: Any) -> bool:
    try:
        if pd.isna(v): return False
        float(v); return True
    except Exception: return False

def detect_assessment_type(text: Any) -> str:
    s = clean_text(text).lower()
    if re.search(r"end\s*[- ]?sem|end\s*semester|\bese\b|external assessment|university exam", s): return "ESE"
    if re.search(r"mse\s*[- ]?1|mid\s*[- ]?sem(ester)?\s*[- ]?1", s): return "MSE-1"
    if re.search(r"mse\s*[- ]?2|mid\s*[- ]?sem(ester)?\s*[- ]?2", s): return "MSE-2"
    if re.search(r"mse\s*[- ]?3|mid\s*[- ]?sem(ester)?\s*[- ]?3", s): return "MSE-3"
    if "teacher" in s and "assessment" in s: return "Teacher Assessment"
    if re.search(r"\bta\b", s): return "Teacher Assessment"
    if "assignment" in s or "tutorial" in s: return "Assignment"
    if "quiz" in s or "class test" in s or "unit test" in s: return "Quiz"
    if "laboratory" in s or re.search(r"\blab\b|practical", s): return "Laboratory"
    return "Other"

def detect_cos(text: Any) -> List[str]:
    s = clean_text(text).upper().replace("–","-").replace("—","-")
    found=[]
    # A range such as CO1 to CO3 represents CO1, CO2, CO3.
    range_matches = re.findall(r"\bCO\s*[-_ ]?(\d+)\s*(?:TO|-)\s*CO?\s*[-_ ]?(\d+)\b", s)
    for a,b in range_matches:
        for n in range(int(a), int(b)+1):
            found.append(f"CO{n}")
    # Remove the range text before finding individual CO labels so the
    # endpoints are not duplicated. Repeated standalone CO labels remain.
    remainder = re.sub(r"\bCO\s*[-_ ]?\d+\s*(?:TO|-)\s*CO?\s*[-_ ]?\d+\b", " ", s)
    for n in re.findall(r"\bCO\s*[-_ ]?(\d+)\b", remainder):
        found.append(f"CO{int(n)}")
    return found

def read_tabular_file(path: str) -> Dict[str,pd.DataFrame]:
    ext=os.path.splitext(path)[1].lower()
    if ext in {".xlsx",".xlsm",".xls"}:
        book=pd.ExcelFile(path)
        return {s:pd.read_excel(path,sheet_name=s,header=None) for s in book.sheet_names}
    if ext==".csv": return {"CSV":pd.read_csv(path,header=None)}
    raise ValueError(f"Unsupported tabular file type: {ext}")

def _find_header(df):
    for r in range(min(15,len(df))):
        vals=[clean_text(x).lower() for x in df.iloc[r].tolist()]
        roll=any("roll no" in x or "rollno" in x or "roll number" in x or "registration" in x or "enrollment" in x for x in vals)
        name=any(re.search(r"\bname\b",x) for x in vals)
        if roll and name: return r
    return 0

def _find_roll_name(df, header):
    roll=name=None
    for c in range(df.shape[1]):
        txt=" ".join(clean_text(x).lower() for x in df.iloc[max(0,header-2):min(len(df),header+3),c].tolist())
        if roll is None and any(k in txt for k in ("roll no","rollno","roll number","registration","enrollment")): roll=c
        if name is None and re.search(r"\bname\b",txt): name=c
    return roll,name

def _column_metadata(df, header):
    """Build metadata without horizontally propagating CO labels.

    Horizontal fill is appropriate for merged assessment-name cells, but it
    must never copy a CO range from one physical column into another.
    """
    # In the supplied format the four structural rows are: group, assessment,
    # CO mapping, maximum marks. This also works when group/assessment cells
    # are merged because we only use the nearest non-empty assessment label.
    assessment_rows = []
    for r in range(header, min(len(df), header + 2)):
        assessment_rows.append(r)
    cos_row = header + 2
    max_row = header + 3

    metas=[]
    last_assessment=""
    for c in range(df.shape[1]):
        group=clean_text(df.iloc[header,c]) if header < len(df) else ""
        name=clean_text(df.iloc[header+1,c]) if header+1 < len(df) else ""
        if group: last_assessment=group
        if name: last_assessment=name if name else last_assessment
        # Search backwards in the two assessment rows for this column; when
        # merged cells are represented as blanks, use the most recent label.
        assessment_text=" | ".join(x for x in (group,name) if x) or last_assessment
        assessment=detect_assessment_type(assessment_text)
        co_text=clean_text(df.iloc[cos_row,c]) if cos_row < len(df) else ""
        cos=detect_cos(co_text)
        # ESE sometimes has its CO range in the assessment-name cell.
        if not cos and assessment == "ESE":
            cos=detect_cos(assessment_text)
        header_text=" | ".join(x for x in (group,name,co_text) if x)
        metas.append({"column":c,"header":header_text,"assessment":assessment,"cos":cos})
    return metas

def extract_student_marks(path: str) -> Dict[str,Any]:
    students=[]; assessments={}; course_info={}; seen_rolls=set()
    for _,raw in read_tabular_file(path).items():
        df=raw.dropna(axis=0,how="all").dropna(axis=1,how="all").reset_index(drop=True)
        if df.empty: continue
        header=_find_header(df); metas=_column_metadata(df,header)
        roll_col,name_col=_find_roll_name(df,header)
        if roll_col is None or name_col is None: continue

        # Find the actual CO-label row. Prefer the row containing the
        # greatest number of CO-bearing cells; this avoids mistaking the
        # assessment-name row "End Sem Exam CO1-CO6" for the CO row.
        candidates=[]
        for r in range(header+1,min(len(df),header+7)):
            score=sum(1 for x in df.iloc[r].tolist() if detect_cos(x))
            if score: candidates.append((score,r))
        co_row=max(candidates)[1] if candidates else header+2
        max_row=co_row+1; data_start=max_row+1

        max_by_col={c:float(df.iloc[max_row,c]) for c in range(df.shape[1]) if is_numeric(df.iloc[max_row,c])}

        # Build one assessment object and one column mapping per physical column.
        col_specs=[]
        for m in metas:
            c=m["column"]
            if c in (roll_col,name_col): continue
            cos=m["cos"]
            if not cos: continue
            maximum=max_by_col.get(c)
            if maximum is None: continue
            col_specs.append((c,m["assessment"],cos,maximum))
            a=assessments.setdefault(m["assessment"],{"name":m["assessment"],"type":"ESE" if m["assessment"]=="ESE" else "IA","co_max_marks":{},"co_student_marks":{}})
            for co in cos:
                a["co_max_marks"][co]=a["co_max_marks"].get(co,0.0)+maximum
                a["co_student_marks"].setdefault(co,[])

        for r in range(data_start,len(df)):
            roll=clean_text(df.iloc[r,roll_col]); name=clean_text(df.iloc[r,name_col])
            if not roll or not name or roll.lower() in {"total","average","percentage","maximum"}: continue
            numeric=any(is_numeric(df.iloc[r,c]) for c,_,_,_ in col_specs)
            if not numeric: continue
            if roll not in seen_rolls:
                students.append({"roll_no":roll,"name":name}); seen_rolls.add(roll)
            # Each physical CO column contributes to the same CO array; repeated COs are summed per student.
            per_assessment={}
            for c,atype,cos,_ in col_specs:
                if not is_numeric(df.iloc[r,c]): continue
                per_assessment.setdefault(atype,[]).append((c,cos,float(df.iloc[r,c])))
            # Append one value for every mapped assessment/CO for every
            # student. A blank mark is treated as 0 for attainment counting
            # so the denominator remains the real class strength.
            for atype,a in assessments.items():
                items=per_assessment.get(atype,[])
                by_co={}
                for _,cos,val in items:
                    for co in cos: by_co[co]=by_co.get(co,0.0)+val
                for co in a["co_student_marks"]:
                    a["co_student_marks"][co].append(by_co.get(co,0.0))

        # ESE source has one total column; reference methodology applies that total to each mapped CO.
        if "ESE" in assessments:
            a=assessments["ESE"]
            base=next(iter(a["co_student_marks"].values()),[])
            mx=max(a["co_max_marks"].values(),default=60.0)
            for co in a["co_student_marks"]:
                a["co_student_marks"][co]=list(base)
                a["co_max_marks"][co]=mx

    return {"students":students,"assessments":list(assessments.values()),"course_info":course_info,"total_students":len(students)}

def build_raw_assessment_data(extracted):
    return extracted.get("assessments",[]) if isinstance(extracted,dict) else []

def extract_course_exit_survey(path: str) -> Dict[str,Any]:
    for _,raw in read_tabular_file(path).items():
        df=raw.dropna(axis=0,how="all").dropna(axis=1,how="all").reset_index(drop=True)
        if df.empty: continue
        headers=[clean_text(x) for x in df.iloc[0].tolist()]
        ratings={}; percentages={}; counts={}; response_rows=[]
        for r in range(1,len(df)):
            if any(clean_text(x).lower() in LIKERT for x in df.iloc[r].tolist()): response_rows.append(r)
        for c,h in enumerate(headers):
            m=re.match(r"\s*(\d+)\s*[\.)-]",h)
            if not m: continue
            co=f"CO{int(m.group(1))}"; vals=[]; cnt={"Strongly Agree":0,"Agree":0,"Neutral":0,"Disagree":0,"Strongly Disagree":0}
            for r in response_rows:
                v=clean_text(df.iloc[r,c]).lower()
                if v in LIKERT: vals.append(LIKERT[v]); label={5:"Strongly Agree",4:"Agree",3:"Neutral",2:"Disagree",1:"Strongly Disagree"}[LIKERT[v]]; cnt[label]+=1
                elif is_numeric(df.iloc[r,c]) and 1<=float(df.iloc[r,c])<=5: vals.append(float(df.iloc[r,c]))
            if vals:
                avg=sum(vals)/len(vals); ratings[co]=round(avg,4); percentages[co]=round(avg/5*100,4); counts[co]=cnt
        if ratings:
            overall=sum(percentages.values())/len(percentages)
            return {"ratings":ratings,"percentages":percentages,"overall_percentage":round(overall,4),"overall_level":round(overall/100*3,4),"response_counts":counts,"total_response_rows":len(response_rows)}
    return {"ratings":{},"percentages":{},"overall_percentage":0.0,"overall_level":0.0,"response_counts":{},"total_response_rows":0}
