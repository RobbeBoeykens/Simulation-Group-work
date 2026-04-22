import numpy as np
import matplotlib.pyplot as plt
import random
from simulation import Simulation

INPUT_FILE = "Big Assignment/Inputs/input-S1-14.txt"
W_TOTAL = 500
R_PILOT = 20
RULE = 1
WELCH_WINDOW = 50    # choose <= W_TOTAL/4

# ------------------------------------------------------------
# Patch: track weekly OV
# ------------------------------------------------------------
original_run = Simulation.runOneSimulation

def patched_run(self):
    original_run(self)

    weekly_elAppWT = np.zeros(self.W)
    weekly_counts = np.zeros(self.W)

    for p in self.patients:
        if p.scanWeek == -1:
            break
        if p.patientType == 1 and not p.isNoShow:
            weekly_elAppWT[p.scanWeek] += p.getAppWT()
            weekly_counts[p.scanWeek] += 1

    safe = np.where(weekly_counts > 0, weekly_counts, 1)
    weekly_elAppWT = weekly_elAppWT / safe

    self.movingAvgOV = (
        self.weightEl * weekly_elAppWT
        + self.weightUr * np.array(self.movingAvgUrgentScanWT)
    )

Simulation.runOneSimulation = patched_run

# ------------------------------------------------------------
# Welch moving average
# ------------------------------------------------------------
def welch_moving_average(y, w):
    T = len(y)
    out = np.zeros(T)
    for t in range(T):
        left = max(0, t - w)
        right = min(T, t + w + 1)
        out[t] = np.mean(y[left:right])
    return out

# ------------------------------------------------------------
# Pilot runs
# ------------------------------------------------------------
sim = Simulation(INPUT_FILE, W_TOTAL, R_PILOT, RULE)
sim.setWeekSchedule()

all_traj = []

print(f"Pilot: {R_PILOT} replications x {W_TOTAL} weeks...")
for r in range(R_PILOT):
    sim.resetSystem()
    random.seed(r)
    sim.runOneSimulation()
    all_traj.append(np.array(sim.movingAvgOV))
    print(f"  Replication {r+1}/{R_PILOT} done", end="\r")

print("\nDone!")

trajectories = np.array(all_traj)   # shape (R, W)
weeks = np.arange(1, W_TOTAL + 1)

# ------------------------------------------------------------
# Welch data
# ------------------------------------------------------------
ybar = np.mean(trajectories, axis=0)
ystd = np.std(trajectories, axis=0, ddof=1)
welch = welch_moving_average(ybar, WELCH_WINDOW)

# ------------------------------------------------------------
# Main Welch plot
# ------------------------------------------------------------
fig, ax = plt.subplots(figsize=(14, 5))
ax.plot(weeks, ybar, color="lightsteelblue", linewidth=1.2, label="Across-replication mean $\\bar{Y}_t$")
ax.fill_between(weeks, ybar - ystd, ybar + ystd, color="lightsteelblue", alpha=0.25, label="±1 std")
ax.plot(weeks, welch, color="black", linewidth=2.5, label=f"Welch moving average (w={WELCH_WINDOW})")

ax.set_title("Welch warm-up plot for OV")
ax.set_xlabel("Week")
ax.set_ylabel("OV")
ax.set_ylim(0.40, 0.50)
ax.grid(True, alpha=0.3)
ax.legend()
plt.tight_layout()
plt.savefig("welch_main_plot.png", dpi=150)
plt.show()

print("Inspect 'welch_main_plot.png' and choose the week where the black curve becomes flat.")