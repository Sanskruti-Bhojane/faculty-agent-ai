
from agents.attainment_calculator import calculate_complete_attainment

expected = {
    "CO1": 1.56,
    "CO2": 1.32,
    "CO3": 1.32,
    "CO4": 1.32,
    "CO5": 1.32,
    "CO6": 1.56,
}

data = calculate_complete_attainment(
    assessment_results={
        "MSE-1": {
            "CO2": {"above_threshold": 33, "total_students": 55},
            "CO3": {"above_threshold": 43, "total_students": 55},
        },
        "MSE-2": {
            "CO5": {"above_threshold": 42, "total_students": 55},
            "CO6": {"above_threshold": 50, "total_students": 55},
        },
        "Teachers Assessment": {
            "CO1": {"above_threshold": 51, "total_students": 55},
            "CO4": {"above_threshold": 41, "total_students": 55},
        },
        "ESE": {
            "CO1": {"above_threshold": 32, "total_students": 55},
            "CO2": {"above_threshold": 32, "total_students": 55},
            "CO3": {"above_threshold": 32, "total_students": 55},
            "CO4": {"above_threshold": 32, "total_students": 55},
            "CO5": {"above_threshold": 32, "total_students": 55},
            "CO6": {"above_threshold": 32, "total_students": 55},
        },
    },
    survey_ratings= {
        "CO1": 5, "CO2": 5, "CO3": 5,
        "CO4": 5, "CO5": 5, "CO6": 5,
    },
    assessment_mapping= {
        "MSE-1": ["CO2", "CO3"],
        "MSE-2": ["CO5", "CO6"],
        "Teachers Assessment": ["CO1", "CO4"],
        "ESE": ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"],
    },
    co_codes= ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"],
    co_po_mapping= {
        "CO1": {"PO1": 3, "PO2": 2, "PO3": 2, "PO5": 1, "PO11": 1},
        "CO2": {"PO1": 3, "PO2": 3, "PO3": 2, "PO4": 2, "PO5": 2, "PO11": 1},
        "CO3": {"PO1": 3, "PO2": 3, "PO3": 2, "PO4": 2},
        "CO4": {"PO1": 3, "PO2": 2, "PO3": 2},
        "CO5": {"PO1": 3, "PO2": 2, "PO3": 2},
        "CO6": {"PO1": 3, "PO2": 2},
    },
    co_pso_mapping= {
        "CO1": {"PSO3": 1},
        "CO2": {"PSO3": 1},
        "CO3": {"PSO3": 1},
    },
    po_codes= ["PO1", "PO2", "PO3", "PO4", "PO5", "PO11"],
    pso_codes= ["PSO3"],
    methodology= {
        "ia_threshold_percent": 40,
        "ese_threshold_percent": 50,
        "level_3_percent": 80,
        "level_2_percent": 60,
        "overall_ia_direct_weight": 0.80,
        "overall_ia_indirect_weight": 0.20,
        "final_ia_weight": 0.40,
        "final_ese_weight": 0.60,
        "ia_assessment_weight": 0.75,
        "ia_secondary_weight": 0.25,
        "co_target": 2.10,
    },
)

print("=" * 60)
print("OBE ATTAINMENT CALCULATOR VALIDATION")
print("=" * 60)

actual = data["overall_co_attainment"]

all_ok = True
for co, exp in expected.items():
    got = actual[co]["overall_attainment"]
    ok = abs(got - exp) < 0.01
    all_ok = all_ok and ok
    print(f"{co}: Expected = {exp:.2f}, Actual = {got:.2f} --> {'PASS' if ok else 'FAIL'}")

print("-" * 60)
print("Direct IA:")
for co in expected:
    print(co, data["overall_ia"][co]["direct"])

print("-" * 60)
print("Overall IA:")
for co in expected:
    print(co, data["overall_ia"][co]["overall_ia"])

print("-" * 60)
print("ESE:")
for co in expected:
    print(co, data["ese_attainment"][co]["level"])

print("-" * 60)
print("RESULT:", "PASS" if all_ok else "FAIL")

if not all_ok:
    raise SystemExit(1)
