import math
import numpy as np
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from simulation import Simulation

# ============================================================
# SETTINGS  — pas hier aan
# ============================================================
DESIGNS = [
    # (urgent_slots, strategy, rule)
    (14, 1, 1),
    (16, 3, 3),
    (16, 1, 2),
    (14, 2, 4),
    (12, 3, 1),
    (10, 1, 2),
    (14, 3, 2),
    (10, 2, 3),
]

WARMUP_WEEKS = 50    # warmup period (bepaald via Welch)
L            = 16    # aantal batches  (= R uit multiple run method)

# Batch lengte:
# Slide 2: "Find L such that |Cov(Xi, Xi+L)| ≈ 0, batch length = 5L"
# We schatten de autocorrelatie-lag L_ac via de post-warmup wekelijkse data
# en stellen M = 5 * L_ac. Dat doen we hieronder automatisch per design.
# Als je een vaste M wil overschrijven, zet FORCE_M op een int (bv. 100).
FORCE_M      = None   # None = automatisch bepalen, int = vaste batchgrootte

OUTPUT_EXCEL = "Big Assignment/Excel Files/batch_means.xlsx"
INPUT_DIR    = "Big Assignment/Inputs"

# ============================================================
# Styling
# ============================================================
BLUE_FILL   = PatternFill("solid", fgColor="1F4E79")
LIGHT_FILL  = PatternFill("solid", fgColor="D6E4F0")
ORANGE_FILL = PatternFill("solid", fgColor="F4B942")
GREEN_FILL  = PatternFill("solid", fgColor="E2EFDA")
GREY_FILL   = PatternFill("solid", fgColor="F2F2F2")
RED_FILL    = PatternFill("solid", fgColor="FCE4D6")

WHITE_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BOLD_FONT  = Font(name="Arial", bold=True, size=10)
REG_FONT   = Font(name="Arial", size=10)

thin = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")


def sh(cell, text, fill=None):
    cell.value = text; cell.font = WHITE_FONT
    cell.fill = fill or BLUE_FILL
    cell.alignment = CENTER; cell.border = THIN_BORDER


def sl(cell, text):
    cell.value = text; cell.font = BOLD_FONT
    cell.alignment = LEFT; cell.border = THIN_BORDER


def sv(cell, value, fmt="#,##0.00000", fill=None):
    cell.value = value; cell.font = REG_FONT
    cell.number_format = fmt; cell.alignment = RIGHT
    cell.border = THIN_BORDER
    if fill: cell.fill = fill


def set_col_width(ws, col, w):
    ws.column_dimensions[col].width = w


# ============================================================
# Helpers
# ============================================================
def safe_avg(values):
    valid = [v for v in values if math.isfinite(v)]
    return sum(valid) / len(valid) if valid else 0.0


def find_autocorr_lag(series):
    """
    Zoek de kleinste lag L_ac waarvoor |autocorr(lag)| < 0.05.
    Slide 2: batchlengte M = 5 * L_ac.
    Geeft minimaal 1 terug.
    """
    arr = np.array([v for v in series if math.isfinite(v)], dtype=float)
    if len(arr) < 10:
        return 1
    arr = arr - arr.mean()
    var = np.var(arr)
    if var == 0:
        return 1
    max_lag = min(len(arr) // 2, 200)
    for lag in range(1, max_lag + 1):
        ac = np.mean(arr[:-lag] * arr[lag:]) / var
        if abs(ac) < 0.005:
            return lag
    return max_lag


# ============================================================
# Run één lange simulatie, geef post-warmup wekelijkse reeks terug
# ============================================================
def run_long_sim(input_file, rule, total_weeks):
    sim = Simulation(input_file, total_weeks, 1, rule)
    sim.setWeekSchedule()
    sim.resetSystem()
    random.seed(0)          # vaste seed voor de lange run
    sim.runOneSimulation()
    return sim


def get_post_warmup_series(sim):
    """Geeft post-warmup wekelijkse OV-reeks terug (één waarde per week)."""
    el_app  = sim.movingAvgElectiveAppWT
    ur_scan = sim.movingAvgUrgentScanWT
    series = []
    for w in range(WARMUP_WEEKS, len(el_app)):
        ov = sim.weightEl * el_app[w] + sim.weightUr * ur_scan[w]
        if math.isfinite(ov):
            series.append(ov)
    return series


# ============================================================
# Batch means berekening
# ============================================================
def compute_batch_means(series, M, L):
    """
    Verdeelt `series` in L batches van M weken.
    Geeft lijst van L batchgemiddelden terug.
    Slide 1: X̄_l = (1/M) * Σ_{k=(l-1)M+1}^{lM} f(Y_k)
    """
    batch_means = []
    for l in range(L):
        batch = series[l * M : (l + 1) * M]
        batch_means.append(safe_avg(batch))
    return batch_means


# ============================================================
# Excel sheet schrijven
# ============================================================
def write_batch_sheet(wb, sheet_name, batch_means, meta):
    ws = wb.create_sheet(title=sheet_name)

    for col, w in zip("ABCDEFG", [6, 18, 18, 18, 18, 14, 14]):
        set_col_width(ws, col, w)

    # titel
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = (f"Batch Mean Method  |  Strategy {meta['strategy']}"
               f"  |  Urgent slots {meta['urgent']}  |  Rule {meta['rule']}")
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    c.fill = BLUE_FILL; c.alignment = CENTER; c.border = THIN_BORDER
    ws.row_dimensions[1].height = 22

    # --- parameters blok ---
    ws.merge_cells("A2:G2")
    h = ws["A2"]
    h.value = "PARAMETERS & METHODE  (Slide 1–3: batch mean method)"
    h.font = Font(name="Arial", bold=True, color="1F4E79", size=10)
    h.fill = LIGHT_FILL; h.alignment = CENTER; h.border = THIN_BORDER

    param_rows = [
        ("Warmup weken (verwijderd)",  meta["warmup"],         "weken"),
        ("Autocorrelatie-lag L_ac",    meta["lag_ac"],         "weken  [|Cov(Xi, Xi+L)| ≈ 0]"),
        ("Batchgrootte M  (= 5·L_ac)", meta["M"],              "weken per batch"),
        ("Aantal batches L",           meta["L"],              "batches"),
        ("Totale run length (na warmup)", meta["run_weeks"],   "weken  (= M × L)"),
        ("Totale simulatielengte",     meta["total_weeks"],    "weken  (warmup + run)"),
    ]

    for i, (label, value, unit) in enumerate(param_rows, start=3):
        ws.row_dimensions[i].height = 16
        sl(ws.cell(row=i, column=1), label)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
        sv(ws.cell(row=i, column=4), value, fmt="#,##0")
        ws.cell(row=i, column=5).value = unit
        ws.cell(row=i, column=5).font = Font(name="Arial", italic=True, color="595959", size=9)
        ws.cell(row=i, column=5).border = THIN_BORDER
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=7)

# --- summary statistieken ---
    sum_start = 3 + len(param_rows) + 1
    ws.merge_cells(f"A{sum_start}:G{sum_start}")
    h2 = ws[f"A{sum_start}"]
    h2.value = "SUMMARY STATISTIEKEN  (Slide 3)"
    h2.font = Font(name="Arial", bold=True, color="1F4E79", size=10)
    h2.fill = LIGHT_FILL; h2.alignment = CENTER; h2.border = THIN_BORDER

    L_val = len(batch_means)
    X_bar = sum(batch_means) / L_val

    # S² = (1/(L-1)) * Σ(X̄_l - X̄)²
    S2      = sum((bm - X_bar) ** 2 for bm in batch_means) / (L_val - 1)
    std_S   = math.sqrt(S2)
    ci_half = 1.96 * math.sqrt(S2 / L_val)

    summary_stats = [
        ("Gemiddelde  X̄  =  (1/L) Σ X̄_l",          X_bar,           "#,##0.00000",   None),
        ("Variantie  S²  =  (1/(L-1)) Σ(X̄_l−X̄)²",  S2,              "#,##0.0000000", None),
        ("Std dev  S  =  √S²",                        std_S,           "#,##0.00000",   None),
        ("95% CI half-width  (1.96·√(S²/L))",         ci_half,         "#,##0.00000",   ORANGE_FILL),
        ("95% CI lower",                               X_bar - ci_half, "#,##0.00000",   GREEN_FILL),
        ("95% CI upper",                               X_bar + ci_half, "#,##0.00000",   GREEN_FILL),
    ]

    for i, (label, value, fmt, fill) in enumerate(summary_stats, start=sum_start + 1):
        ws.row_dimensions[i].height = 16
        sl(ws.cell(row=i, column=1), label)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)
        sv(ws.cell(row=i, column=5), value, fmt=fmt, fill=fill)
        ws.merge_cells(start_row=i, start_column=5, end_row=i, end_column=7)

    # --- per-batch tabel ---
    detail_start = sum_start + len(summary_stats) + 2
    ws.merge_cells(f"A{detail_start}:G{detail_start}")
    h3 = ws[f"A{detail_start}"]
    h3.value = "PER-BATCH DETAIL  —  X̄_l = (1/M) Σ_{k=(l-1)M+1}^{lM} f(Y_k)"
    h3.font = Font(name="Arial", bold=True, color="1F4E79", size=10)
    h3.fill = LIGHT_FILL; h3.alignment = CENTER; h3.border = THIN_BORDER

    hdr_row = detail_start + 1
    ws.row_dimensions[hdr_row].height = 20
    col_headers = ["Batch l", "Week start", "Week einde", "X̄_l (batch gem.)", "X̄_l − X̄", "(X̄_l − X̄)²", ""]
    for ci, hdr in enumerate(col_headers, start=1):
        c = ws.cell(row=hdr_row, column=ci)
        c.value = hdr; c.font = WHITE_FONT; c.fill = BLUE_FILL
        c.alignment = CENTER; c.border = THIN_BORDER

    for i, bm in enumerate(batch_means):
        dr = hdr_row + 1 + i
        ws.row_dimensions[dr].height = 16
        zebra = GREY_FILL if i % 2 == 0 else None
        week_start = WARMUP_WEEKS + i * meta["M"] + 1
        week_end   = WARMUP_WEEKS + (i + 1) * meta["M"]
        diff = bm - X_bar

        c1 = ws.cell(row=dr, column=1)
        c1.value = i + 1; c1.font = BOLD_FONT; c1.alignment = CENTER; c1.border = THIN_BORDER
        if zebra: c1.fill = zebra

        sv(ws.cell(row=dr, column=2), week_start, fmt="#,##0", fill=zebra)
        sv(ws.cell(row=dr, column=3), week_end,   fmt="#,##0", fill=zebra)
        sv(ws.cell(row=dr, column=4), bm,                      fill=zebra)
        sv(ws.cell(row=dr, column=5), diff,                    fill=zebra)
        sv(ws.cell(row=dr, column=6), diff ** 2,               fill=zebra)
        ws.cell(row=dr, column=7).border = THIN_BORDER

    # footer totalen
    avg_row = hdr_row + 1 + L_val
    ws.row_dimensions[avg_row].height = 18
    for ci in range(1, 8):
        c = ws.cell(row=avg_row, column=ci)
        c.border = THIN_BORDER; c.font = BOLD_FONT
    ws.cell(row=avg_row, column=1).value = "AVG/SOM"
    ws.cell(row=avg_row, column=1).fill = BLUE_FILL
    ws.cell(row=avg_row, column=1).font = WHITE_FONT
    ws.cell(row=avg_row, column=1).alignment = CENTER
    sv(ws.cell(row=avg_row, column=4), X_bar,                       fill=ORANGE_FILL)
    sv(ws.cell(row=avg_row, column=5), 0.0,                         fill=LIGHT_FILL)
    sv(ws.cell(row=avg_row, column=6), float(np.sum([(b - X_bar)**2 for b in batch_means])), fill=ORANGE_FILL)

    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)


# ============================================================
# MAIN
# ============================================================
def main():
    import os
    os.makedirs(os.path.dirname(OUTPUT_EXCEL), exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    for urgent, strategy, rule in DESIGNS:
        input_file = f"{INPUT_DIR}/input-S{strategy}-{urgent}.txt"
        sheet_name = f"S{strategy}-{urgent}-R{rule}"
        print(f"\n{'='*60}")
        print(f"Design: {urgent} urgent slots | Strategy {strategy} | Rule {rule}")

        # --- stap 1: pilot run om autocorrelatie-lag te bepalen ---
        # Totale lengte nog onbekend; gebruik een conservatieve lange run voor de pilot
        PILOT_WEEKS = WARMUP_WEEKS + L * 500   # ruim genoeg voor lag-schatting
        sim_pilot = run_long_sim(input_file, rule, PILOT_WEEKS)
        series_pilot = get_post_warmup_series(sim_pilot)

        lag_ac = find_autocorr_lag(series_pilot)
        M = FORCE_M if FORCE_M is not None else max(5 * lag_ac, 10)
        RUN_WEEKS   = M * L
        TOTAL_WEEKS = WARMUP_WEEKS + RUN_WEEKS

        print(f"  Autocorrelatie-lag L_ac = {lag_ac} weken  →  M = 5·L_ac = {M} weken")
        print(f"  L = {L} batches  →  run length = {RUN_WEEKS} weken  (total = {TOTAL_WEEKS})")

        # --- stap 2: echte lange simulatie met juiste lengte ---
        sim = run_long_sim(input_file, rule, TOTAL_WEEKS)
        series = get_post_warmup_series(sim)

        # zorg dat series lang genoeg is
        needed = M * L
        if len(series) < needed:
            print(f"  WAARSCHUWING: series ({len(series)}) korter dan M*L ({needed}). "
                  f"Vergroot PILOT_WEEKS of verklein M/L.")
            series = series + [series[-1]] * (needed - len(series))

        # --- stap 3: batch means ---
        batch_means = compute_batch_means(series, M, L)

        X_bar   = float(np.mean(batch_means))
        var_xbar = float(np.var(batch_means, ddof=1) / L)
        ci_half  = 1.96 * math.sqrt(var_xbar)

        print(f"  X̄ = {X_bar:.5f}  |  CI ± {ci_half:.5f}")
        for i, bm in enumerate(batch_means):
            print(f"    Batch {i+1:2d}: {bm:.5f}")

        meta = dict(
            urgent=urgent, strategy=strategy, rule=rule,
            warmup=WARMUP_WEEKS, lag_ac=lag_ac, M=M, L=L,
            run_weeks=RUN_WEEKS, total_weeks=TOTAL_WEEKS,
        )

        write_batch_sheet(wb, sheet_name, batch_means, meta)

    wb.save(OUTPUT_EXCEL)
    print(f"\nExcel saved → {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()