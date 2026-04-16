import matplotlib.pyplot as plt
import numpy as np
import random
from simulation import Simulation

WARMUP_WEEKS = 875  # <-- pas dit aan naar jullie visueel bepaalde warm-up

# --- Patch runOneSimulation ---
original_runOneSimulation = Simulation.runOneSimulation

def runOneSimulation_patched(self):
    original_runOneSimulation(self)
    
    weekly_elAppWT = np.zeros(self.W)
    weekly_counts  = np.zeros(self.W)
    for patient in self.patients:
        if patient.scanWeek == -1:
            break
        if patient.patientType == 1 and not patient.isNoShow:
            weekly_elAppWT[patient.scanWeek] += patient.getAppWT()
            weekly_counts[patient.scanWeek]  += 1
    counts_safe = np.where(weekly_counts > 0, weekly_counts, 1)
    weekly_elAppWT = weekly_elAppWT / counts_safe

    self.movingAvgOV = (
        self.weightEl * weekly_elAppWT +
        self.weightUr * np.array(self.movingAvgUrgentScanWT)
    )

Simulation.runOneSimulation = runOneSimulation_patched


# --- Patch runSimulations met warm-up verwijdering ---
def runSimulations_patched(self):
    self.setWeekSchedule()
    self.all_OV_trajectories = []

    totals = {"elAppWT": 0, "elScanWT": 0, "urScanWT": 0, "OT": 0, "OV": 0}
    OV_values = []

    # Per-replicatie steady-state waarden (na warm-up)
    ss_elAppWT = []
    ss_elScanWT = []
    ss_urScanWT = []
    ss_OT = []
    ss_OV = []

    print(f"Warm-up: {WARMUP_WEEKS} weken | Steady-state: {self.W - WARMUP_WEEKS} weken\n")
    print("r \t elAppWT \t elScanWT \t urScanWT \t OT \t OV")
    print("(steady-state waarden per replicatie)\n")

    for r in range(self.R):
        self.resetSystem()
        random.seed(r)
        self.runOneSimulation()

        # Haal patiënten op die ENKEL in steady-state periode vallen
        ss_patients_el = [
            p for p in self.patients
            if p.scanWeek >= WARMUP_WEEKS and p.patientType == 1 and not p.isNoShow
        ]
        ss_patients_ur = [
            p for p in self.patients
            if p.scanWeek >= WARMUP_WEEKS and p.patientType == 2 and not p.isNoShow
        ]

        # Herbereken gemiddelden over steady-state periode
        r_elAppWT  = np.mean([p.getAppWT()  for p in ss_patients_el]) if ss_patients_el else 0
        r_elScanWT = np.mean([p.getScanWT() for p in ss_patients_el]) if ss_patients_el else 0
        r_urScanWT = np.mean([p.getScanWT() for p in ss_patients_ur]) if ss_patients_ur else 0

        # Overtime: gemiddeld over steady-state weken
        r_OT = np.mean(self.movingAvgOT[WARMUP_WEEKS:]) if self.W > WARMUP_WEEKS else 0

        r_OV = self.weightEl * r_elAppWT + self.weightUr * r_urScanWT

        ss_elAppWT.append(r_elAppWT)
        ss_elScanWT.append(r_elScanWT)
        ss_urScanWT.append(r_urScanWT)
        ss_OT.append(r_OT)
        ss_OV.append(r_OV)

        self.all_OV_trajectories.append(list(self.movingAvgOV))

        totals["elAppWT"]  += r_elAppWT
        totals["elScanWT"] += r_elScanWT
        totals["urScanWT"] += r_urScanWT
        totals["OT"]       += r_OT
        totals["OV"]       += r_OV
        OV_values.append(r_OV)

        print(f"{r} \t {r_elAppWT:.2f} \t\t {r_elScanWT:.5f} \t "
              f"{r_urScanWT:.2f} \t\t {r_OT:.2f} \t {r_OV:.2f}")

    R = self.R
    print("\n" + "-"*80)
    print(f"AVG: \t {np.mean(ss_elAppWT):.2f} \t\t {np.mean(ss_elScanWT):.5f} \t "
          f"{np.mean(ss_urScanWT):.2f} \t\t {np.mean(ss_OT):.2f} \t {np.mean(ss_OV):.2f}")

    OV_stdev = np.std(OV_values, ddof=1)
    OV_mean  = np.mean(OV_values)
    ci_half  = 1.96 * OV_stdev / np.sqrt(R)
    print(f"STDEV OV:  {OV_stdev:.4f}")
    print(f"95% CI OV: [{OV_mean - ci_half:.4f}, {OV_mean + ci_half:.4f}]")

    # Sla resultaten op als attributen voor verdere analyse
    self.ss_results = {
        "elAppWT": np.array(ss_elAppWT),
        "elScanWT": np.array(ss_elScanWT),
        "urScanWT": np.array(ss_urScanWT),
        "OT": np.array(ss_OT),
        "OV": np.array(ss_OV),
    }

Simulation.runSimulations = runSimulations_patched


# --- Run ---
sim = Simulation("Big Assignment/Inputs/input-S1-14.txt", 2000, 30, 1)
sim.runSimulations()

# Resultaten beschikbaar als:
# sim.ss_results["OV"]       → array van R steady-state OV waarden
# sim.ss_results["elAppWT"]  → etc.
# --- Plot steady-state trajectories ---
trajectories = np.array(sim.all_OV_trajectories)  # shape: (R, W)
weeks = np.arange(1, sim.W + 1)

# Cumulatief moving average per replicatie
cumavg = np.cumsum(trajectories, axis=1) / weeks  # shape: (R, W)

fig, ax = plt.subplots(figsize=(14, 6))

colors = plt.cm.tab10(np.linspace(0, 1, sim.R))
for r in range(sim.R):
    ax.plot(weeks, cumavg[r], alpha=0.5, linewidth=1.2,
            color=colors[r], label=f"Rep {r}")

# Warm-up grens visualiseren
ax.axvline(x=WARMUP_WEEKS, color='red', linewidth=2, 
           linestyle='--', label=f"Warm-up grens (week {WARMUP_WEEKS})")

ax.set_title("Moving average OV per replicatie  —  S1-14")
ax.set_xlabel("Week (observatie)")
ax.set_ylabel("Cumulatief gemiddelde OV")
ax.legend(fontsize=8, ncol=2, title="Replicaties")
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig("ov_cumavg_S1-14.png", dpi=150)
plt.show()
print("Plot opgeslagen als ov_cumavg_S1-14.png")