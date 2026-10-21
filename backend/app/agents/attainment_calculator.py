"""Deterministic OBE attainment calculator for Faculty Agent.

The faculty uploads the complete Student Marks Excel and Course Exit Survey Excel.
No attainment-methodology file is required.

The calculator follows the structure used by the OBE Attainment Reference:
- IA threshold = 40% of maximum marks
- IA level: >=80% -> 3, >=60% -> 2, otherwise 1
- ESE threshold = 50% of maximum marks
- Direct IA = 75% of first mapped IA level + 25% of second mapped IA level
- Overall IA = 80% Direct IA + 20% Indirect CO
- Overall CO = 40% Overall IA + 60% ESE level
- Outcome direct attainment = weighted CO attainment by articulation / 3
- Outcome indirect attainment = 3 when the outcome has mapped COs
- Overall outcome = 80% direct + 20% indirect
"""

from typing import Any, Dict, List, Mapping, Optional

IA_THRESHOLD_PERCENT = 40.0
ESE_THRESHOLD_PERCENT = 50.0
LEVEL_3_PERCENT = 80.0
LEVEL_2_PERCENT = 60.0
SURVEY_MAX_RATING = 5.0
MAX_ARTICULATION = 3.0
DIRECT_WEIGHT = 0.80
INDIRECT_WEIGHT = 0.20
IA_WEIGHT = 0.40
ESE_WEIGHT = 0.60
CO_TARGET = 2.10


def _f(v: Any, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except Exception:
        return default


def _r(v: Any, digits: int = 2) -> float:
    return round(_f(v), digits)


def _canon(v: Any) -> str:
    return "".join(c.lower() for c in str(v or "") if c.isalnum())


def _is_ese(name: Any) -> bool:
    c = _canon(name)
    return any(x in c for x in ("ese", "endsem", "endsemester", "externalassessment"))


def level_from_percentage(percent: float) -> int:
    percent = _f(percent)
    if percent >= LEVEL_3_PERCENT:
        return 3
    if percent >= LEVEL_2_PERCENT:
        return 2
    return 1


def _student_stat(raw: Mapping[str, Any], threshold: float) -> Dict[str, Any]:
    marks = raw.get("student_marks", raw.get("marks", []))

    if isinstance(marks, (list, tuple)):
        values = []
        for x in marks:
            try:
                if x is not None and x != "":
                    values.append(float(x))
            except Exception:
                pass

        total = len(values)
        above = sum(1 for x in values if x >= threshold)
    else:
        total = int(_f(raw.get("total_students", raw.get("total", 0))))
        above = int(_f(raw.get("above_threshold", 0)))

    percentage = (above / total * 100.0) if total else 0.0

    return {
        "total_students": total,
        "above_threshold": above,
        "percentage": _r(percentage),
        "threshold": _r(threshold),
        "level": level_from_percentage(percentage) if total else 0,
    }


def calculate_indirect_co_attainment(
    survey_ratings: Optional[Dict[str, Any]] = None,
    survey_percentages: Optional[Dict[str, Any]] = None,
    survey_overall_percentage: Optional[float] = None,
    **_: Any,
) -> Dict[str, Any]:
    """Calculate indirect CO attainment from Course Exit Survey.

    A Likert response is 1..5.
    Question-level indirect attainment = average rating / 5 * 3.

    If the extractor also supplies percentage evidence, it is preserved.
    The reference survey workbook's aggregate calculation is:
        weighted survey percentage / 100 * 3
    """

    ratings = survey_ratings or {}
    percentages = survey_percentages or {}

    out: Dict[str, Any] = {}

    for co, value in ratings.items():
        rating = _f(value)
        pct = _f(percentages.get(co, rating / SURVEY_MAX_RATING * 100.0))

        # Same 0-3 normalization as the supplied survey workbook.
        indirect = pct / 100.0 * MAX_ARTICULATION

        out[str(co)] = {
            "survey_rating": _r(rating),
            "survey_percentage": _r(pct),
            "indirect_attainment": _r(indirect),
            "attainment": _r(indirect),
        }

    # Preserve the aggregate survey calculation as metadata.
    if survey_overall_percentage is not None:
        aggregate = _f(survey_overall_percentage) / 100.0 * MAX_ARTICULATION
        for row in out.values():
            row["survey_overall_percentage"] = _r(survey_overall_percentage)
            row["survey_overall_level"] = _r(aggregate)

    return out


def _normalise_mapping(mapping: Optional[Dict[str, Any]]) -> Dict[str, List[str]]:
    result: Dict[str, List[str]] = {}

    for assessment, values in (mapping or {}).items():
        if isinstance(values, Mapping):
            result[str(assessment)] = [str(k) for k, v in values.items() if _f(v) > 0]
        elif isinstance(values, (list, tuple, set)):
            result[str(assessment)] = [str(x) for x in values]

    return result


def calculate_direct_ia(
    assessment_results: Optional[Dict[str, Any]] = None,
    assessment_mapping: Optional[Dict[str, Any]] = None,
    co_codes: Optional[List[str]] = None,
    **_: Any,
) -> Dict[str, Any]:

    results = assessment_results or {}
    mapping = _normalise_mapping(assessment_mapping)

    by_co: Dict[str, List[Dict[str, Any]]] = {}
    assessment_statistics: Dict[str, Dict[str, Any]] = {}

    for assessment_name, values in results.items():
        if _is_ese(assessment_name) or not isinstance(values, Mapping):
            continue

        mapped_cos = mapping.get(assessment_name, list(values.keys()))

        for co in mapped_cos:
            if co not in values:
                continue

            raw = values[co]
            if not isinstance(raw, Mapping):
                raw = {"student_marks": raw}

            max_marks = _f(raw.get("max_marks", 0))
            threshold = _f(
                raw.get(
                    "threshold",
                    max_marks * IA_THRESHOLD_PERCENT / 100.0,
                )
            )

            stat = _student_stat(raw, threshold)
            assessment_statistics.setdefault(str(assessment_name), {})[str(co)] = stat

            by_co.setdefault(str(co), []).append(
                {"assessment": assessment_name, **stat}
            )

    ordered_cos = list(co_codes or [])
    for co in by_co:
        if co not in ordered_cos:
            ordered_cos.append(co)

    direct: Dict[str, float] = {}
    details: Dict[str, Any] = {}

    for co in ordered_cos:
        entries = by_co.get(co, [])

        primary = _f(entries[0].get("level") if entries else 0)
        secondary = _f(entries[1].get("level") if len(entries) > 1 else 0)

        direct_value = _r(0.75 * primary + 0.25 * secondary)

        direct[co] = direct_value
        details[co] = {
            "primary_level": _r(primary),
            "secondary_level": _r(secondary),
            "direct_attainment": direct_value,
            "assessment_details": entries,
        }

    return {
        "by_co": details,
        "direct": direct,
        "assessment_statistics": assessment_statistics,
    }


def _find_ese_assessment(
    assessment_results: Mapping[str, Any],
) -> Optional[str]:
    """Find the actual ESE / End-Semester assessment name."""
    for name in (assessment_results or {}).keys():
        if _is_ese(name):
            return str(name)
    return None


def calculate_ese(
    assessment_results: Optional[Dict[str, Any]] = None,
    co_codes: Optional[List[str]] = None,
    **_: Any,
) -> Dict[str, Any]:

    results = assessment_results or {}

    ese_name = _find_ese_assessment(results)

    if not ese_name:
        return {}

    raw_assessment = results.get(ese_name, {})
    if not isinstance(raw_assessment, Mapping):
        return {}

    output: Dict[str, Any] = {}
    codes = list(co_codes or raw_assessment.keys())

    # Some older/raw inputs preserve only one ESE CO column. The project's
    # Excel extractor defines one ESE total as applying to all mapped COs, so
    # use that same ESE mark list for any missing CO key.
    fallback_raw = None
    for candidate in raw_assessment.values():
        if isinstance(candidate, Mapping):
            marks = candidate.get("student_marks", candidate.get("marks", []))
            if isinstance(marks, (list, tuple)) and marks:
                fallback_raw = candidate
                break

    for co in codes:
        raw = raw_assessment.get(co)

        if raw is None and fallback_raw is not None:
            raw = dict(fallback_raw)

        if raw is None:
            continue

        if not isinstance(raw, Mapping):
            raw = {"student_marks": raw}

        max_marks = _f(raw.get("max_marks", 60))
        threshold = _f(
            raw.get(
                "threshold",
                max_marks * ESE_THRESHOLD_PERCENT / 100.0,
            )
        )

        output[str(co)] = _student_stat(raw, threshold)

    return output


def _attainment_value(value: Any) -> float:
    if isinstance(value, Mapping):
        return _f(
            value.get(
                "overall",
                value.get(
                    "overall_attainment",
                    value.get("attainment", 0),
                ),
            )
        )
    return _f(value)


def _calculate_outcomes(
    overall_co: Dict[str, Any],
    mapping: Optional[Dict[str, Any]],
    codes: Optional[List[str]],
    course_indirect_attainment: float,
) -> Dict[str, Any]:
    """Calculate direct and overall PO/PSO attainment.

    The Course Exit Survey produces ONE course-level indirect value.
    That exact value is reused for every mapped PO/PSO.
    """
    output: Dict[str, Any] = {}

    for outcome in codes or []:
        contributions = []
        numerator = 0.0
        denominator = 0.0

        for co, row in (mapping or {}).items():
            if not isinstance(row, Mapping) or outcome not in row:
                continue

            articulation = _f(row[outcome])
            if articulation <= 0:
                continue

            co_attainment = _attainment_value(overall_co.get(co, 0))
            numerator += co_attainment * articulation
            denominator += articulation

            contributions.append({
                "co": str(co),
                "co_attainment": _r(co_attainment),
                "mapping": _r(articulation),
                "contribution": _r(co_attainment * articulation),
            })

        if denominator == 0:
            output[str(outcome)] = {
                "attainment": None,
                "direct_attainment": None,
                "average_articulation": None,
                "contributions": [],
                "indirect_attainment": None,
                "overall_attainment": None,
            }
            continue

        direct = numerator / denominator
        indirect = _f(course_indirect_attainment)
        overall = DIRECT_WEIGHT * direct + INDIRECT_WEIGHT * indirect

        output[str(outcome)] = {
            "attainment": _r(direct),
            "direct_attainment": _r(direct),
            "average_articulation": _r(
                denominator / len(contributions) if contributions else 0
            ),
            "contributions": contributions,
            "indirect_attainment": _r(indirect),
            "overall_attainment": _r(overall),
        }

    return output


def calculate_complete_attainment(
    *,
    assessment_results: Optional[Dict[str, Any]] = None,
    survey_ratings: Optional[Dict[str, Any]] = None,
    survey_percentages: Optional[Dict[str, Any]] = None,
    survey_overall_percentage: Optional[float] = None,
    co_po_mapping: Optional[Dict[str, Any]] = None,
    co_pso_mapping: Optional[Dict[str, Any]] = None,
    assessment_mapping: Optional[Dict[str, Any]] = None,
    co_codes: Optional[List[str]] = None,
    po_codes: Optional[List[str]] = None,
    pso_codes: Optional[List[str]] = None,
    target: Optional[Any] = None,
    **_: Any,
) -> Dict[str, Any]:

    assessment_results = assessment_results or {}

    co_codes = list(co_codes or [])
    po_codes = list(po_codes or [])
    pso_codes = list(pso_codes or [])

    discovered = set(co_codes)

    for values in assessment_results.values():
        if isinstance(values, Mapping):
            discovered.update(values.keys())

    if not co_codes:
        co_codes = sorted(discovered)

    if not assessment_mapping:
        assessment_mapping = {
            name: list(values.keys())
            for name, values in assessment_results.items()
            if isinstance(values, Mapping) and not _is_ese(name)
        }

    ia = calculate_direct_ia(
        assessment_results=assessment_results,
        assessment_mapping=assessment_mapping,
        co_codes=co_codes,
    )

    indirect = calculate_indirect_co_attainment(
        survey_ratings=survey_ratings,
        survey_percentages=survey_percentages,
        survey_overall_percentage=survey_overall_percentage,
    )

    # ONE course-level indirect attainment value.
    # Prefer the extractor's aggregate percentage; otherwise average the
    # CO survey percentages. This value is reused everywhere.
    if survey_overall_percentage is not None:
        course_indirect_attainment = _r(
            _f(survey_overall_percentage) / 100.0 * MAX_ARTICULATION
        )
    elif indirect:
        course_indirect_attainment = _r(
            sum(_f(row.get("survey_percentage", 0)) for row in indirect.values())
            / len(indirect) / 100.0 * MAX_ARTICULATION
        )
    else:
        course_indirect_attainment = 0.0

    for row in indirect.values():
        row["course_indirect_attainment"] = course_indirect_attainment
        row["survey_overall_level"] = course_indirect_attainment

    ese = calculate_ese(
        assessment_results=assessment_results,
        co_codes=co_codes,
    )

    # Keep ESE statistics available to the report/web layer as well.
    # The previous calculator intentionally excluded ESE from IA statistics,
    # which caused blank/zero ESE fields in the generated attainment output.
    ese_name = _find_ese_assessment(assessment_results)
    assessment_statistics = dict(ia["assessment_statistics"])
    if ese_name and ese:
        assessment_statistics[ese_name] = ese

    overall_ia: Dict[str, Any] = {}
    overall: Dict[str, Any] = {}

    actual_target = CO_TARGET if target is None else _f(target)

    for co in co_codes:
        direct = _f(ia["direct"].get(co))
        indirect_value = course_indirect_attainment

        overall_ia_value = _r(
            DIRECT_WEIGHT * direct + INDIRECT_WEIGHT * indirect_value
        )

        ese_level = _f(
            ese.get(co, {}).get("level", 0)
        )

        overall_co = _r(
            IA_WEIGHT * overall_ia_value + ESE_WEIGHT * ese_level
        )

        overall_ia[co] = {
            "direct": _r(direct),
            "indirect": _r(indirect_value),
            "overall_ia": overall_ia_value,
            "overall_attainment": overall_ia_value,
        }

        overall[co] = {
            "co": co,
            "direct": _r(direct),
            "indirect": _r(indirect_value),
            "overall_ia": overall_ia_value,
            "ese": _r(ese_level),
            "overall": overall_co,
            "overall_attainment": overall_co,
            "target": _r(actual_target),
            "gap": _r(overall_co - actual_target),
            "status": (
                "ACHIEVED"
                if overall_co >= actual_target
                else "NOT ACHIEVED"
            ),
        }

    po = _calculate_outcomes(
        overall,
        co_po_mapping,
        po_codes,
        course_indirect_attainment,
    )

    pso = _calculate_outcomes(
        overall,
        co_pso_mapping,
        pso_codes,
        course_indirect_attainment,
    )

    return {
        "calculation_status": "SUCCESS",
        "co_attainment": overall,
        "overall_co_attainment": overall,
        "direct_co_attainment": {
            co: {
                "direct_attainment": ia["direct"].get(co, 0)
            }
            for co in co_codes
        },
        "indirect_co_attainment": indirect,
        "course_indirect_attainment": course_indirect_attainment,
        "overall_ia": overall_ia,
        "ese_attainment": ese,
        "po_attainment": po,
        "pso_attainment": pso,
        "co_po_attainment": po,
        "co_pso_attainment": pso,
        "assessment_statistics": assessment_statistics,
        "methodology_used": {
            "ia_threshold_percent": IA_THRESHOLD_PERCENT,
            "ese_threshold_percent": ESE_THRESHOLD_PERCENT,
            "level_3_percent": LEVEL_3_PERCENT,
            "level_2_percent": LEVEL_2_PERCENT,
            "direct_ia_weight": DIRECT_WEIGHT,
            "indirect_ia_weight": INDIRECT_WEIGHT,
            "ia_weight": IA_WEIGHT,
            "ese_weight": ESE_WEIGHT,
            "co_target": actual_target,
        },
    }
