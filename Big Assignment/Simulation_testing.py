import numpy as np
import matplotlib.pyplot as plt
import random
from simulation import Simulation

# ============================================================
# INSTELLINGEN — enkel hier aanpassen
# ============================================================
INPUT_FILE = "Big Assignment/Inputs/input-S3-16.txt"
W_TOTAL    = 5000
R_PILOT    = 30
RULE       = 3

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
# PLOT: Cumulatief gemiddelde per replicatie + avg + std
# ============================================================
cumavg = np.cumsum(trajectories, axis=1) / weeks  # (R, W)

grand_avg = np.mean(cumavg, axis=0)  # gemiddelde over replicaties per week
grand_std = np.std(cumavg, axis=0, ddof=1)  # stdev over replicaties per week

fig, ax = plt.subplots(figsize=(14, 5))

# Individuele replicaties
for r in range(R_PILOT):
    ax.plot(weeks, cumavg[r], alpha=0.3, linewidth=0.8, color="steelblue")

# Gemiddelde over replicaties
ax.plot(weeks, grand_avg, color="black", linewidth=2, label="Gemiddelde over replicaties")

# ±1 standaarddeviatie
ax.fill_between(weeks,
                grand_avg - grand_std,
                grand_avg + grand_std,
                alpha=0.2, color="black", label="±1 std")

ax.set_title("Cumulatief gemiddelde OV per replicatie\n"
             "Kijk waar alle lijnen convergeren → warm-up grens")
ax.set_xlabel("Week")
ax.set_ylabel("Cumulatief gemiddelde OV")
ax.legend()
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("warmup_plot.png", dpi=150)
plt.show()

print("Plot opgeslagen als warmup_plot.png")
