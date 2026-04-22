import numpy as np
import matplotlib.pyplot as plt
import random
from simulation import Simulation

# ============================================================
# INSTELLINGEN — enkel hier aanpassen
# ============================================================
INPUT_FILE    = "Big Assignment/Inputs/input-S1-14.txt"
W_TOTAL       = 500
R_PILOT       = 30
RULE          = 1
WELCH_WINDOW  = 100  # smoothing window voor Welch

# ============================================================
# PATCH: week-per-week OV bijhouden
# ============================================================
original_run = Simulation.runOneSimulation

def patched_run(self):
    original_run(self)
    weekly_elAppWT = np.zeros(self.W)
    weekly_counts  = np.zeros(self.W)
    for p in self.patients:
        if p.scanWeek == -1:
            break
        if p.patientType == 1 and not p.isNoShow:
            weekly_elAppWT[p.scanWeek] += p.getAppWT()
            weekly_counts[p.scanWeek]  += 1
    safe = np.where(weekly_counts > 0, weekly_counts, 1)
    self.movingAvgOV = (
        self.weightEl * (weekly_elAppWT / safe) +
        self.weightUr * np.array(self.movingAvgUrgentScanWT)
    )

Simulation.runOneSimulation = patched_run

# ============================================================
# PILOT DRAAIEN
# ============================================================
sim = Simulation(INPUT_FILE, W_TOTAL, R_PILOT, RULE)
sim.setWeekSchedule()

all_traj = []

print(f"Pilot: {R_PILOT} replicaties × {W_TOTAL} weken...")
for r in range(R_PILOT):
    sim.resetSystem()
    random.seed(r)
    sim.runOneSimulation()
    all_traj.append(np.array(sim.movingAvgOV))
    print(f"  Replicatie {r+1}/{R_PILOT} klaar", end="\r")

print("\nKlaar!\n")
trajectories = np.array(all_traj)  # (R, W)
weeks = np.arange(1, W_TOTAL + 1)

# ============================================================
# WELCH: warm-up bepalen
# ============================================================
grand_avg    = np.mean(trajectories, axis=0)
welch_smooth = np.convolve(grand_avg, np.ones(WELCH_WINDOW) / WELCH_WINDOW, mode='valid')
weeks_smooth = weeks[WELCH_WINDOW - 1:]

# Automatische schatting: eerste week waar curve blijvend stabiel is
# = verandering < 1% van steady-state niveau voor 10 opeenvolgende weken
end_level   = np.mean(welch_smooth[-50:])
threshold   = 0.01 * end_level
diffs       = np.abs(np.diff(welch_smooth))
warmup_auto = W_TOTAL  # fallback
for i in range(len(diffs) - 10):
    if np.all(diffs[i:i+10] < threshold):
        warmup_auto = int(weeks_smooth[i])
        break

print(f">>> Welch automatische warm-up schatting : week {warmup_auto}")
print(f"    Steady-state niveau (gemiddelde)      : {end_level:.4f}")
print(f"    Bekijk de plots en pas WARMUP_WEEKS aan indien nodig.\n")

# ============================================================
# PLOTS
# ============================================================
cumavg    = np.cumsum(trajectories, axis=1) / weeks
grand_cum = np.mean(cumavg, axis=0)
std_cum   = np.std(cumavg, axis=0, ddof=1)

fig, axes = plt.subplots(2, 1, figsize=(14, 10))

# Plot 1: Welch smoothed curve
ax = axes[0]
ax.plot(weeks, grand_avg, alpha=0.3, color='steelblue', label="Grand average OV (raw)")
ax.plot(weeks_smooth, welch_smooth, color='red', linewidth=2,
        label=f"Welch smoothed (window={WELCH_WINDOW})")
ax.axvline(x=warmup_auto, color='green', linewidth=2, linestyle='--',
           label=f"Warm-up schatting: week {warmup_auto}")
ax.axhline(y=end_level, color='gray', linewidth=1, linestyle=':',
           label=f"Steady-state niveau: {end_level:.4f}")
ax.set_title("Welch's procedure — warm-up bepaling")
ax.set_xlabel("Week")
ax.set_ylabel("OV")
ax.legend()
ax.grid(True, alpha=0.3)

# Plot 2: Cumulatief gemiddelde per replicatie + avg + std
ax = axes[1]
for r in range(R_PILOT):
    ax.plot(weeks, cumavg[r], alpha=0.2, linewidth=0.8, color='steelblue')
ax.plot(weeks, grand_cum, color='black', linewidth=2, label="Gemiddelde over replicaties")
ax.fill_between(weeks,
                grand_cum - std_cum,
                grand_cum + std_cum,
                alpha=0.2, color='black', label="±1 std")
ax.axvline(x=warmup_auto, color='green', linewidth=2, linestyle='--',
           label=f"Warm-up schatting: week {warmup_auto}")
ax.set_title("Cumulatief gemiddelde OV per replicatie")
ax.set_xlabel("Week")
ax.set_ylabel("Cumulatief gemiddelde OV")
ax.legend()
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig("warmup_plot.png", dpi=150)
plt.show()

print("Plot opgeslagen als warmup_plot.png")
print("Bepaal visueel de warm-up grens en vul deze in als WARMUP_WEEKS in de volgende stap.")