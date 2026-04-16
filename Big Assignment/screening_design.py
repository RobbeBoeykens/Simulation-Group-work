import matplotlib.pyplot as plt
plt.switch_backend("Agg")
import numpy as np
import random
from simulation import Simulation

# --- Patch runOneSimulation to also track per-week AppWT and OV ---
original_runOneSimulation = Simulation.runOneSimulation

def runOneSimulation_patched(self):
    original_runOneSimulation(self)
    
    # Compute per-week elective AppWT (not tracked by default)
    weekly_elAppWT = np.zeros(self.W)
    weekly_counts  = np.zeros(self.W)
    for patient in self.patients:
        if patient.scanWeek == -1:
            break
        if patient.patientType == 1 and not patient.isNoShow:
            weekly_elAppWT[patient.scanWeek] += patient.getAppWT()
            weekly_counts[patient.scanWeek]  += 1
    # Avoid division by zero
    counts_safe = np.where(weekly_counts > 0, weekly_counts, 1)
    weekly_elAppWT = weekly_elAppWT / counts_safe

    # OV per week = 1/168 * elAppWT + 1/9 * urScanWT
    self.movingAvgOV = (
        self.weightEl * weekly_elAppWT +
        self.weightUr * np.array(self.movingAvgUrgentScanWT)
    )

Simulation.runOneSimulation = runOneSimulation_patched


# --- Patch runSimulations to capture per-replication trajectories ---
def runSimulations_patched(self):
    self.setWeekSchedule()
    self.all_OV_trajectories = []   # shape: (R, W)

    print("r \t elAppWT \t elScanWT \t urScanWT \t OT \t OV \n")
    totals = {"elAppWT": 0, "elScanWT": 0, "urScanWT": 0, "OT": 0, "OV": 0}

    for r in range(self.R):
        self.resetSystem()
        random.seed(r)
        self.runOneSimulation()

        ov = self.avgElectiveAppWT * self.weightEl + self.avgUrgentScanWt * self.weightUr
        self.all_OV_trajectories.append(list(self.movingAvgOV))

        totals["elAppWT"]  += self.avgElectiveAppWT
        totals["elScanWT"] += self.avgElectiveScanWT
        totals["urScanWT"] += self.avgUrgentScanWt
        totals["OT"]       += self.avgOT
        totals["OV"]       += ov

        print(f"{r} \t {self.avgElectiveAppWT:.2f} \t\t {self.avgElectiveScanWT:.5f} \t "
              f"{self.avgUrgentScanWt:.2f} \t\t {self.avgOT:.2f} \t {ov:.2f}")

    R = self.R
    print("--------------------------------------------------------------------------------")
    print(f"AVG: \t {totals['elAppWT']/R:.2f} \t\t {totals['elScanWT']/R:.5f} \t "
          f"{totals['urScanWT']/R:.2f} \t\t {totals['OT']/R:.2f} \t {totals['OV']/R:.2f} \n")

Simulation.runSimulations = runSimulations_patched


# --- Run ---
sim = Simulation("Big Assignment/Inputs/input-S1-14.txt", 100, 5, 1)
sim.runSimulations()


# --- Plot ---
trajectories = np.array(sim.all_OV_trajectories)  # shape: (R, W)
weeks = np.arange(1, sim.W + 1)

# Cumulatief moving average per replicatie: MA[w] = mean(OV[0..w])
cumavg = np.cumsum(trajectories, axis=1) / weeks  # shape: (R, W)

fig, ax = plt.subplots(figsize=(14, 6))

colors = plt.cm.tab10(np.linspace(0, 1, sim.R))
for r in range(sim.R):
    ax.plot(weeks, cumavg[r], alpha=0.5, linewidth=1.2,
            color=colors[r], label=f"Rep {r}")

ax.set_title("Moving average OV per replication  —  S1-14")
ax.set_xlabel("Weak (observation)")
ax.set_ylabel("Cumulative average OV")
ax.legend(fontsize=8, ncol=2, title="Replications")
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig("ov_cumavg_S1-14.png", dpi=150)
plt.show()
print("Plot opgeslagen als ov_cumavg_S1-14.png")