import math
import numpy as np
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from simulation import Simulation

# ============================================================
# SETTINGS
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

WARMUP_WEEKS = 50
RUN_WEEKS    = 483
TOTAL_WEEKS  = WARMUP_WEEKS + RUN_WEEKS   # passed to Simulation
R            = 30
OUTPUT_EXCEL = "design_results.xlsx"
INPUT_DIR    = "Big Assignment/Inputs"

# ============================================================
# Styling helpers
# ============================================================
BLUE_FILL   = PatternFill("solid", fgColor="1F4E79")   # dark blue header
LIGHT_FILL  = PatternFill("solid", fgColor="D6E4F0")   # light blue alternating
ORANGE_FILL = PatternFill("solid", fgColor="F4B942")   # accent for summary
GREEN_FILL  = PatternFill("solid", fgColor="E2EFDA")   # corrected column
GREY_FILL   = PatternFill("solid", fgColor="F2F2F2")   # zebra striping

WHITE_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BOLD_FONT   = Font(name="Arial", bold=True, size=10)
REG_FONT    = Font(name="Arial", size=10)
BOLD_BLUE   = Font(name="Arial", bold=True, color="1F4E79", size=10)

thin = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")


def style_header_cell(cell, text, fill=None):
    cell.value = text
    cell.font = WHITE_FONT
    cell.fill = fill or BLUE_FILL
    cell.alignment = CENTER
    cell.border = THIN_BORDER


def style_label_cell(cell, text):
    cell.value = text
    cell.font = BOLD_FONT
    cell.alignment = LEFT
    cell.border = THIN_BORDER


def style_value_cell(cell, value, fmt="#,##0.0000", fill=None):
    cell.value = value
    cell.font = REG_FONT
    cell.number_format = fmt
    cell.alignment = RIGHT
    cell.border = THIN_BORDER
    if fill:
        cell.fill = fill


def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width


# ============================================================
# Helpers
# ============================================================
def safe_avg(values: list) -> float:
    """Average that ignores NaN / Inf (e.g. weeks with 0 patients)."""
    valid = [v for v in values if math.isfinite(v)]
    return sum(valid) / len(valid) if valid else 0.0


# ============================================================
# Compute replication outputs  — uses movingAvg slices (post-warmup)
# so every week gets equal weight, consistent with your existing code.
# ============================================================
def compute_replication_outputs(sim):
    # Slice out only the post-warmup run period
    post_el_app  = sim.movingAvgElectiveAppWT[WARMUP_WEEKS : WARMUP_WEEKS + RUN_WEEKS]
    post_ur_scan = sim.movingAvgUrgentScanWT [WARMUP_WEEKS : WARMUP_WEEKS + RUN_WEEKS]
    post_el_scan = sim.movingAvgElectiveScanWT[WARMUP_WEEKS : WARMUP_WEEKS + RUN_WEEKS]
    post_ot      = sim.movingAvgOT            [WARMUP_WEEKS : WARMUP_WEEKS + RUN_WEEKS]

    el_app_wt  = safe_avg(post_el_app)
    ur_scan_wt = safe_avg(post_ur_scan)
    el_scan_wt = safe_avg(post_el_scan)
    ot         = safe_avg(post_ot)

    ov = sim.weightEl * el_app_wt + sim.weightUr * ur_scan_wt

    # Control variates: total arrivals over the FULL run (warmup + run)
    # so that v_E / v_U match the same time horizon as Y_E / Y_U
    y_e = sum(1 for p in sim.patients if p.patientType == 1)
    y_u = sum(1 for p in sim.patients if p.patientType == 2)

    return ov, y_e, y_u, el_app_wt, el_scan_wt, ur_scan_wt, ot


# ============================================================
# Write one design sheet
# ============================================================
def write_design_sheet(wb, sheet_name, results, meta):
    """
    results : list of dicts with keys Xi, YE_i, YU_i, ElAppWT, ElScanWT, OT
    meta    : dict with v_E, v_U, c_E, c_U, mean_raw, std_raw, ci_half_raw,
                        mean_cv, std_cv, ci_half_cv, reduction_pct, urgent, strategy, rule
    """
    ws = wb.create_sheet(title=sheet_name)

    # ---- Column widths ----
    col_widths = {"A": 6, "B": 16, "C": 16, "D": 16, "E": 22, "F": 6,
                  "G": 28, "H": 18}
    for col, w in col_widths.items():
        set_col_width(ws, col, w)

    # ================================================================
    # SECTION 1 – CONFIGURATION INFO  (row 1-2)
    # ================================================================
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = (f"Control Variates Output  |  Strategy {meta['strategy']}"
                        f"  |  Urgent slots {meta['urgent']}  |  Rule {meta['rule']}")
    title_cell.font  = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    title_cell.fill  = BLUE_FILL
    title_cell.alignment = CENTER
    title_cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 22

    # ================================================================
    # SECTION 2 – MAIN SUMMARY TABLE  (rows 3-11)
    # ================================================================
    r = 3
    ws.merge_cells(f"A{r}:H{r}")
    hdr = ws[f"A{r}"]
    hdr.value = "MAIN SUMMARY"
    hdr.font  = Font(name="Arial", bold=True, color="1F4E79", size=10)
    hdr.fill  = PatternFill("solid", fgColor="D6E4F0")
    hdr.alignment = CENTER
    hdr.border = THIN_BORDER

    summary_rows = [
        # (label_col_A,   value_col_B,       label_col_D,        value_col_E)
        ("Known E[Y_E]  =  v_E",  meta["v_E"],  "Known E[Y_U]  =  v_U",  meta["v_U"]),
        ("Estimated  c_E",        meta["c_E"],  "Estimated  c_U",        meta["c_U"]),
        (None, None, None, None),   # blank separator
        ("Mean  X̄  (raw OV)",     meta["mean_raw"],  "Mean  Ȳ_E  (elective arr)",  np.mean([d["YE_i"] for d in results])),
        ("Mean  Ȳ_U  (urgent arr)", np.mean([d["YU_i"] for d in results]), "Mean  X̄_cv  (corrected)", meta["mean_cv"]),
        (None, None, None, None),
        ("Std dev  (raw)",         meta["std_raw"],   "Std dev  (cv)",             meta["std_cv"]),
        ("95% CI  half-width (raw)", meta["ci_half_raw"], "95% CI  half-width (cv)", meta["ci_half_cv"]),
        ("Variance reduction",     f"{meta['reduction_pct']:.2f}%", "", ""),
    ]

    for i, (la, va, lb, vb) in enumerate(summary_rows, start=r + 1):
        row_idx = i
        ws.row_dimensions[row_idx].height = 17
        if la is None:
            continue
        # Label A
        ca = ws.cell(row=row_idx, column=1)
        style_label_cell(ca, la)
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2)
        # Value B (merged C)
        cb = ws.cell(row=row_idx, column=3)
        fill = ORANGE_FILL if "corrected" in str(la).lower() or "X̄_cv" in str(la) else None
        if isinstance(va, float):
            style_value_cell(cb, va, fill=fill)
        else:
            cb.value = va; cb.font = REG_FONT; cb.alignment = RIGHT; cb.border = THIN_BORDER
            if fill: cb.fill = fill
        ws.merge_cells(start_row=row_idx, start_column=3, end_row=row_idx, end_column=4)

        # spacer E
        ws.cell(row=row_idx, column=5).border = THIN_BORDER

        # Label F
        if lb:
            cf = ws.cell(row=row_idx, column=6)
            style_label_cell(cf, lb)
            ws.merge_cells(start_row=row_idx, start_column=6, end_row=row_idx, end_column=7)
            # Value G
            cg = ws.cell(row=row_idx, column=8)
            fill2 = ORANGE_FILL if "corrected" in str(lb).lower() or "X̄_cv" in str(lb) else None
            if isinstance(vb, float):
                style_value_cell(cg, vb, fill=fill2)
            else:
                cg.value = vb; cg.font = REG_FONT; cg.alignment = RIGHT; cg.border = THIN_BORDER
                if fill2: cg.fill = fill2

    detail_start = r + len(summary_rows) + 3

    # ================================================================
    # SECTION 3 – DETAILED PER-REPLICATION TABLE
    # ================================================================
    ws.merge_cells(f"A{detail_start}:H{detail_start}")
    hdr2 = ws[f"A{detail_start}"]
    hdr2.value = "DETAILED PER-REPLICATION OUTPUT"
    hdr2.font  = Font(name="Arial", bold=True, color="1F4E79", size=10)
    hdr2.fill  = PatternFill("solid", fgColor="D6E4F0")
    hdr2.alignment = CENTER
    hdr2.border = THIN_BORDER

    # Formula explanation row
    expl_row = detail_start + 1
    ws.merge_cells(f"A{expl_row}:H{expl_row}")
    exp = ws[f"A{expl_row}"]
    exp.value = "X_cv,i  =  Xᵢ  −  c_E·(YE_i − v_E)  −  c_U·(YU_i − v_U)        [Slide 45: X̄_cv = X̄ − c·(Ȳ − v)]"
    exp.font  = Font(name="Arial", italic=True, color="595959", size=9)
    exp.alignment = LEFT
    exp.border = THIN_BORDER

    # Column headers
    hdr_row = expl_row + 1
    ws.row_dimensions[hdr_row].height = 20
    headers = ["Rep", "Xᵢ  (OV)", "YE_i  (El.arr)", "YU_i  (Ur.arr)", "Xᵢ_cv  (corrected)"]
    fills   = [BLUE_FILL, BLUE_FILL, BLUE_FILL, BLUE_FILL, PatternFill("solid", fgColor="1A5276")]
    for col_idx, (h, f) in enumerate(zip(headers, fills), start=1):
        c = ws.cell(row=hdr_row, column=col_idx)
        c.value     = h
        c.font      = WHITE_FONT
        c.fill      = f
        c.alignment = CENTER
        c.border    = THIN_BORDER

    # Data rows
    for i, d in enumerate(results):
        data_row = hdr_row + 1 + i
        ws.row_dimensions[data_row].height = 16
        row_fill = GREY_FILL if i % 2 == 0 else None

        # Rep number
        c = ws.cell(row=data_row, column=1)
        c.value = i + 1; c.font = BOLD_FONT; c.alignment = CENTER; c.border = THIN_BORDER
        if row_fill: c.fill = row_fill

        # Xi
        style_value_cell(ws.cell(row=data_row, column=2), d["Xi"],   fill=row_fill)
        # YE_i
        style_value_cell(ws.cell(row=data_row, column=3), d["YE_i"], fmt="#,##0", fill=row_fill)
        # YU_i
        style_value_cell(ws.cell(row=data_row, column=4), d["YU_i"], fmt="#,##0", fill=row_fill)
        # Xi_cv
        style_value_cell(ws.cell(row=data_row, column=5), d["Xi_cv"], fill=GREEN_FILL)

    # Averages footer row
    avg_row = hdr_row + 1 + len(results)
    ws.row_dimensions[avg_row].height = 18

    avg_labels = ["AVG", None, None, None, None]
    avg_vals   = [
        None,
        np.mean([d["Xi"]    for d in results]),
        np.mean([d["YE_i"]  for d in results]),
        np.mean([d["YU_i"]  for d in results]),
        np.mean([d["Xi_cv"] for d in results]),
    ]
    avg_fmts = [None, "#,##0.0000", "#,##0.0", "#,##0.0", "#,##0.0000"]

    for col_idx in range(1, 6):
        c = ws.cell(row=avg_row, column=col_idx)
        c.border = THIN_BORDER
        c.font   = BOLD_FONT
        if col_idx == 1:
            c.value = "AVG"
            c.alignment = CENTER
            c.fill = BLUE_FILL
            c.font = WHITE_FONT
        elif avg_vals[col_idx - 1] is not None:
            c.value        = avg_vals[col_idx - 1]
            c.number_format = avg_fmts[col_idx - 1]
            c.alignment    = RIGHT
            c.fill         = ORANGE_FILL if col_idx == 5 else PatternFill("solid", fgColor="D6E4F0")

    # Freeze panes below header row
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)


# ============================================================
# MAIN
# ============================================================
def main():
    wb = Workbook()
    wb.remove(wb.active)

    for urgent, strategy, rule in DESIGNS:
        input_file = f"{INPUT_DIR}/input-S{strategy}-{urgent}.txt"
        sheet_name = f"S{strategy}-{urgent}-R{rule}"
        print(f"\n=== Running design S{strategy}-{urgent} (rule {rule}) ===")

        sim = Simulation(input_file, TOTAL_WEEKS, R, rule)
        sim.setWeekSchedule()

        X_vals, YE_vals, YU_vals = [], [], []

        for r in range(R):
            sim.resetSystem()
            random.seed(r)
            sim.runOneSimulation()

            ov, y_e, y_u, el_app, el_scan, ur_scan, ot = compute_replication_outputs(sim)

            X_vals.append(ov)
            YE_vals.append(y_e)
            YU_vals.append(y_u)

            print(f"  Rep {r+1:2d} | OV={ov:.5f} | ElAppWT={el_app:.3f} | "
                  f"UrScanWT={ur_scan:.3f} | ElArr={y_e} | UrArr={y_u}")

        X  = np.array(X_vals,  dtype=float)
        YE = np.array(YE_vals, dtype=float)
        YU = np.array(YU_vals, dtype=float)

        # Known means — same horizon as Y (full TOTAL_WEEKS)
        v_E = 5 * TOTAL_WEEKS * sim.lambdaElective
        v_U = TOTAL_WEEKS * (4 * sim.lambdaUrgent[0] + 2 * sim.lambdaUrgent[1])

        # Optimal c  (slide 45: c = Cov(X, Y) / Var(Y))
        var_YE = np.var(YE, ddof=1)
        var_YU = np.var(YU, ddof=1)
        c_E = 0.0 if var_YE == 0 else np.cov(X, YE, ddof=1)[0, 1] / var_YE
        c_U = 0.0 if var_YU == 0 else np.cov(X, YU, ddof=1)[0, 1] / var_YU

        # Corrected values per replication  (slide 45: X_cv = X - c*(Y - v))
        X_cv = X - c_E * (YE - v_E) - c_U * (YU - v_U)

        # Summary statistics
        mean_raw    = float(np.mean(X));    std_raw    = float(np.std(X, ddof=1))
        ci_half_raw = 1.96 * std_raw / np.sqrt(R)
        mean_cv     = float(np.mean(X_cv)); std_cv     = float(np.std(X_cv, ddof=1))
        ci_half_cv  = 1.96 * std_cv / np.sqrt(R)
        reduction   = 100 * (1 - std_cv / std_raw) if std_raw > 0 else 0.0

        print(f"\n  v_E={v_E:.1f}  v_U={v_U:.1f}  c_E={c_E:.6f}  c_U={c_U:.6f}")
        print(f"  Raw  mean={mean_raw:.5f}  std={std_raw:.5f}  CI±{ci_half_raw:.5f}")
        print(f"  CV   mean={mean_cv:.5f}   std={std_cv:.5f}   CI±{ci_half_cv:.5f}")
        print(f"  Variance reduction: {reduction:.2f}%")

        results = [
            {"Xi": float(X[i]), "YE_i": float(YE[i]),
             "YU_i": float(YU[i]), "Xi_cv": float(X_cv[i])}
            for i in range(R)
        ]

        meta = dict(
            urgent=urgent, strategy=strategy, rule=rule,
            v_E=v_E, v_U=v_U, c_E=c_E, c_U=c_U,
            mean_raw=mean_raw, std_raw=std_raw, ci_half_raw=ci_half_raw,
            mean_cv=mean_cv,   std_cv=std_cv,   ci_half_cv=ci_half_cv,
            reduction_pct=reduction,
        )

        write_design_sheet(wb, sheet_name, results, meta)

    wb.save(OUTPUT_EXCEL)
    print(f"\nExcel saved: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
    