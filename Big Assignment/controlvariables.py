import numpy as np
import pandas as pd
import random
from simulation import Simulation

# ============================================================
# SETTINGS — multiple designs
# ============================================================
DESIGNS = [
    # (urgent, strategy, rule)
    (14, 1, 1),
    (16, 3, 3),
    (16, 1, 2),
    (14, 2, 4),
    (12, 3, 1),
    (10, 1, 2),
    (14, 3, 2),
    (10, 2, 3),
]

W_TOTAL = 4083
R = 16

# Optional warm-up deletion
USE_WARMUP = True
WARMUP_WEEKS = 50

OUTPUT_EXCEL = "design_results.xlsx"


def compute_replication_outputs(sim):
    """
    Compute replication-level output X and control variables Y_E, Y_U
    from the current simulation state.

    Returns:
        ov (float): replication objective value
        y_e (int): total elective arrivals
        y_u (int): total urgent arrivals
        el_app_wt (float): elective appointment waiting time
        el_scan_wt (float): elective scan waiting time
        ur_scan_wt (float): urgent scan waiting time
        ot (float): overtime
    """

    # ---------- CONTROL VARIABLES ----------
    y_e = sum(1 for p in sim.patients if p.patientType == 1)
    y_u = sum(1 for p in sim.patients if p.patientType == 2)

    # ---------- PERFORMANCE MEASURES ----------
    if not USE_WARMUP:
        el_app_wt = sim.avgElectiveAppWT
        el_scan_wt = sim.avgElectiveScanWT
        ur_scan_wt = sim.avgUrgentScanWt
        ot = sim.avgOT
    else:
        ss_patients_el = [
            p for p in sim.patients
            if p.scanWeek >= WARMUP_WEEKS and p.patientType == 1 and not p.isNoShow
        ]
        ss_patients_ur = [
            p for p in sim.patients
            if p.scanWeek >= WARMUP_WEEKS and p.patientType == 2 and not p.isNoShow
        ]

        el_app_wt = np.mean([p.getAppWT() for p in ss_patients_el]) if ss_patients_el else 0.0
        el_scan_wt = np.mean([p.getScanWT() for p in ss_patients_el]) if ss_patients_el else 0.0
        ur_scan_wt = np.mean([p.getScanWT() for p in ss_patients_ur]) if ss_patients_ur else 0.0
        ot = np.mean(sim.movingAvgOT[WARMUP_WEEKS:]) if sim.W > WARMUP_WEEKS else 0.0

    ov = sim.weightEl * el_app_wt + sim.weightUr * ur_scan_wt

    return ov, y_e, y_u, el_app_wt, el_scan_wt, ur_scan_wt, ot


def main():
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:

        for urgent, strategy, rule in DESIGNS:
            input_file = f"Big Assignment/Inputs/input-S{strategy}-{urgent}.txt"
            sheet_name = f"S{strategy}-{urgent}-R{rule}"

            print(f"\n=== Running design S{strategy}-{urgent} (rule {rule}) ===")

            sim = Simulation(input_file, W_TOTAL, R, rule)
            sim.setWeekSchedule()

            # Store replication outputs
            X_vals = []       # OV
            YE_vals = []      # elective arrivals
            YU_vals = []      # urgent arrivals

            EL_APP_vals = []
            EL_SCAN_vals = []
            UR_SCAN_vals = []
            OT_vals = []

            replication_rows = []

            print("Running replications...\n")
            for r in range(R):
                sim.resetSystem()
                random.seed(r)
                sim.runOneSimulation()

                ov, y_e, y_u, el_app, el_scan, ur_scan, ot = compute_replication_outputs(sim)

                X_vals.append(ov)
                YE_vals.append(y_e)
                YU_vals.append(y_u)

                EL_APP_vals.append(el_app)
                EL_SCAN_vals.append(el_scan)
                UR_SCAN_vals.append(ur_scan)
                OT_vals.append(ot)

                replication_rows.append({
                    "Replication": r + 1,
                    "OV": ov,
                    "ElAppWT": el_app,
                    "ElScanWT": el_scan,
                    "UrScanWT": ur_scan,
                    "Overtime": ot,
                    "ElectiveArr": y_e,
                    "UrgentArr": y_u,
                })

                print(
                    f"Rep {r+1:2d} | "
                    f"OV={ov:.4f} | "
                    f"ElAppWT={el_app:.2f} | "
                    f"UrScanWT={ur_scan:.2f} | "
                    f"ElectiveArr={y_e} | UrgentArr={y_u}"
                )

            # Convert to arrays
            X = np.array(X_vals, dtype=float)
            YE = np.array(YE_vals, dtype=float)
            YU = np.array(YU_vals, dtype=float)

            # Known means of the control variates
            v_E = 5 * W_TOTAL * sim.lambdaElective
            v_U = W_TOTAL * (4 * sim.lambdaUrgent[0] + 2 * sim.lambdaUrgent[1])

            # Estimate c_E and c_U
            var_YE = np.var(YE, ddof=1)
            var_YU = np.var(YU, ddof=1)

            c_E = 0.0 if var_YE == 0 else np.cov(X, YE, ddof=1)[0, 1] / var_YE
            c_U = 0.0 if var_YU == 0 else np.cov(X, YU, ddof=1)[0, 1] / var_YU

            # Control-variate adjusted replication outputs
            X_cv = X - c_E * (YE - v_E) - c_U * (YU - v_U)

            # Raw estimator summary
            mean_raw = np.mean(X)
            std_raw = np.std(X, ddof=1)
            ci_half_raw = 1.96 * std_raw / np.sqrt(R)

            # Control variate estimator summary
            mean_cv = np.mean(X_cv)
            std_cv = np.std(X_cv, ddof=1)
            ci_half_cv = 1.96 * std_cv / np.sqrt(R)

            reduction_pct = 100 * (1 - std_cv / std_raw) if std_raw > 0 else 0.0

            print("\n" + "=" * 70)
            print("CONTROL VARIATES RESULTS")
            print("=" * 70)
            print(f"Known mean elective arrivals v_E = {v_E:.3f}")
            print(f"Known mean urgent arrivals   v_U = {v_U:.3f}")
            print(f"Estimated c_E = {c_E:.6f}")
            print(f"Estimated c_U = {c_U:.6f}")

            print("\n--- RAW ESTIMATOR (no control variates) ---")
            print(f"Mean OV           = {mean_raw:.6f}")
            print(f"Std dev OV        = {std_raw:.6f}")
            print(f"95% CI half-width = {ci_half_raw:.6f}")
            print(f"95% CI            = [{mean_raw - ci_half_raw:.6f}, {mean_raw + ci_half_raw:.6f}]")

            print("\n--- CV ESTIMATOR (with elective + urgent arrivals) ---")
            print(f"Mean OV           = {mean_cv:.6f}")
            print(f"Std dev OV        = {std_cv:.6f}")
            print(f"95% CI half-width = {ci_half_cv:.6f}")
            print(f"95% CI            = [{mean_cv - ci_half_cv:.6f}, {mean_cv + ci_half_cv:.6f}]")
            print(f"\nVariance reduction (via std dev) = {reduction_pct:.2f}%")

            print("\n--- Extra outputs (raw, for reference) ---")
            print(f"Mean elective appointment WT = {np.mean(EL_APP_vals):.4f}")
            print(f"Mean elective scan WT        = {np.mean(EL_SCAN_vals):.4f}")
            print(f"Mean urgent scan WT          = {np.mean(UR_SCAN_vals):.4f}")
            print(f"Mean overtime                = {np.mean(OT_vals):.4f}")

            # ============================================================
            # WRITE TO EXCEL
            # ============================================================
            rep_df = pd.DataFrame(replication_rows)

            summary_df = pd.DataFrame([
                {"Metric": "Input file", "Value": input_file},
                {"Metric": "Urgent slots", "Value": urgent},
                {"Metric": "Strategy", "Value": strategy},
                {"Metric": "Rule", "Value": rule},
                {"Metric": "W_TOTAL", "Value": W_TOTAL},
                {"Metric": "Replications", "Value": R},
                {"Metric": "Use warm-up", "Value": USE_WARMUP},
                {"Metric": "Warm-up weeks", "Value": WARMUP_WEEKS if USE_WARMUP else 0},
                {"Metric": "Known mean elective arrivals v_E", "Value": v_E},
                {"Metric": "Known mean urgent arrivals v_U", "Value": v_U},
                {"Metric": "Estimated c_E", "Value": c_E},
                {"Metric": "Estimated c_U", "Value": c_U},
                {"Metric": "Raw mean OV", "Value": mean_raw},
                {"Metric": "Raw std dev OV", "Value": std_raw},
                {"Metric": "Raw 95% CI half-width", "Value": ci_half_raw},
                {"Metric": "Raw 95% CI lower", "Value": mean_raw - ci_half_raw},
                {"Metric": "Raw 95% CI upper", "Value": mean_raw + ci_half_raw},
                {"Metric": "CV mean OV", "Value": mean_cv},
                {"Metric": "CV std dev OV", "Value": std_cv},
                {"Metric": "CV 95% CI half-width", "Value": ci_half_cv},
                {"Metric": "CV 95% CI lower", "Value": mean_cv - ci_half_cv},
                {"Metric": "CV 95% CI upper", "Value": mean_cv + ci_half_cv},
                {"Metric": "Variance reduction (%)", "Value": reduction_pct},
                {"Metric": "Mean elective appointment WT", "Value": np.mean(EL_APP_vals)},
                {"Metric": "Mean elective scan WT", "Value": np.mean(EL_SCAN_vals)},
                {"Metric": "Mean urgent scan WT", "Value": np.mean(UR_SCAN_vals)},
                {"Metric": "Mean overtime", "Value": np.mean(OT_vals)},
            ])

            # Write replication table
            rep_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

            # Write summary table a bit lower on same sheet
            summary_start_row = len(rep_df) + 3
            summary_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=summary_start_row)

    print(f"\nExcel file created: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()