"""
File created by Samuel Bakker.
Contains the Simulation-class used in simulation.py.

This is the file you should run.
Due to the "if __name__ == "__main__":"-statement.

Integrated version:
- Welch output structure is kept the same
- Same console output style
- Same Excel output style: Week + Replication columns
- Antithetic variates and control variables are integrated
"""

import re
import math
import random
from functools import cmp_to_key
from openpyxl import Workbook

from helper import Exponential_distribution, Normal_distribution, Bernouilli_distribution
from slot import Slot
from patient import Patient


class Simulation:
    """
    Simulation instance
    """

    inputFileName: str
    D: int = 6
    amountOTSlotsPerDay: int = 10
    S: int = 32 + amountOTSlotsPerDay
    slotLength: float = float(15 / 60)
    lambdaElective: float = 28.345
    meanTardiness: float = 0
    stdevTardiness: float = 2.5
    probNoShow: float = 0.02
    meanElectiveDuration: float = 15
    stdevElectiveDuration: float = 3
    lambdaUrgent: tuple[float] = (2.5, 1.25)
    probUrgentType: tuple[float] = (0.7, 0.1, 0.1, 0.05, 0.05)
    cumulativeProbUrgentType: tuple[float] = (0.7, 0.8, 0.9, 0.95, 1.0)
    meanUrgentDuration: tuple[float] = (15, 17.5, 22.5, 30, 30)
    stdevUrgentDuration: tuple[float] = (2.5, 1, 2.5, 1, 4.5)
    weightEl: float = 1.0 / 168.0
    weightUr: float = 1.0 / 9.0

    W: int
    R: int
    d: int
    s: int
    w: int
    r: int
    rule: int
    weekSchedule: list[list[Slot]]

    patients: list[Patient]
    movingAvgElectiveAppWT: list[float]
    movingAvgElectiveScanWT: list[float]
    movingAvgUrgentScanWT: list[float]
    movingAvgOT: list[float]
    avgElectiveAppWT: float
    avgElectiveScanWT: float
    avgUrgentScanWt: float
    avgOT: float
    numberOfElectivePatientsPlanned: int
    numberOfUrgentPatientsPlanned: int

    def __init__(self, filename: str, W: int, R: int, rule: int) -> None:
        self.patients = list()
        self.inputFileName = filename
        self.W = W
        self.R = R
        self.rule = rule

        self.avgElectiveAppWT = 0
        self.avgElectiveScanWT = 0
        self.avgUrgentScanWt = 0
        self.avgOT = 0
        self.numberOfElectivePatientsPlanned = 0
        self.numberOfUrgentPatientsPlanned = 0

        self.weekSchedule = []
        for d in range(self.D):
            self.weekSchedule.append([Slot() for s in range(self.S)])

        self.movingAvgElectiveAppWT = list()
        self.movingAvgElectiveScanWT = list()
        self.movingAvgUrgentScanWT = list()
        self.movingAvgOT = list()

    # ============================================================
    # Excel output: SAME AS WELCH
    # ============================================================

    def writeWeeklyObjectiveValuesToExcel(
        self,
        weeklyOVPerReplication: list[list[float]],
        filename: str = "test_it.xlsx"
    ) -> None:
        """
        Write the weekly objective values to Excel with weeks in the first column
        and one column per replication.

        Same Excel structure as your Welch code.
        """
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Weekly OV"

        header = ["Week"] + [f"Replication {r + 1}" for r in range(len(weeklyOVPerReplication))]
        worksheet.append(header)

        for week in range(self.W):
            row = [week + 1]
            for replication in range(len(weeklyOVPerReplication)):
                row.append(weeklyOVPerReplication[replication][week])
            worksheet.append(row)

        workbook.save(filename)

    # ============================================================
    # Variance reduction helper methods
    # ============================================================

    @staticmethod
    def sample_mean(values: list[float]) -> float:
        values = [v for v in values if math.isfinite(v)]
        return sum(values) / len(values) if values else 0.0

    @staticmethod
    def sample_variance(values: list[float]) -> float:
        values = [v for v in values if math.isfinite(v)]

        if len(values) <= 1:
            return 0.0

        mean = Simulation.sample_mean(values)
        return sum((x - mean) ** 2 for x in values) / (len(values) - 1)

    @staticmethod
    def sample_covariance(x_values: list[float], y_values: list[float]) -> float:
        paired_values = [
            (x, y)
            for x, y in zip(x_values, y_values)
            if math.isfinite(x) and math.isfinite(y)
        ]

        if len(paired_values) <= 1:
            return 0.0

        xs = [x for x, y in paired_values]
        ys = [y for x, y in paired_values]

        mean_x = Simulation.sample_mean(xs)
        mean_y = Simulation.sample_mean(ys)

        return sum((x - mean_x) * (y - mean_y) for x, y in paired_values) / (len(paired_values) - 1)

    def getWeeklyPatientCounts(self) -> tuple[list[int], list[int]]:
        """
        Control variables:
        - number of elective patients scanned per week
        - number of urgent patients scanned per week
        """
        weeklyElective = [0 for _ in range(self.W)]
        weeklyUrgent = [0 for _ in range(self.W)]

        for patient in self.patients:
            if patient.scanWeek == -1:
                continue
            if patient.scanWeek >= self.W:
                continue

            if patient.patientType == 1:
                weeklyElective[patient.scanWeek] += 1
            elif patient.patientType == 2:
                weeklyUrgent[patient.scanWeek] += 1

        return weeklyElective, weeklyUrgent

    def runOneSimulationWithSeed(self, seed: int, useAntithetic: bool = False) -> dict:
        """
        Runs one simulation with a fixed seed.

        Normal run:
            uses U

        Antithetic run:
            uses 1 - U

        random.randint is also patched because helper.py may use randint(0, 1000)
        to generate uniforms.
        """
        self.resetSystem()
        random.seed(seed)

        if useAntithetic:
            originalRandom = random.random
            originalRandint = random.randint

            def antitheticRandom():
                u = originalRandom()
                return 1.0 - u

            def antitheticRandint(a, b):
                k = originalRandint(a, b)
                return a + b - k

            random.random = antitheticRandom
            random.randint = antitheticRandint

            try:
                self.runOneSimulation()
            finally:
                random.random = originalRandom
                random.randint = originalRandint
        else:
            self.runOneSimulation()

        weeklyOV = []
        for week in range(self.W):
            weeklyOV.append(
                self.movingAvgElectiveAppWT[week] * self.weightEl
                + self.movingAvgUrgentScanWT[week] * self.weightUr
            )

        weeklyElectiveCounts, weeklyUrgentCounts = self.getWeeklyPatientCounts()

        replicationOV = (
            self.avgElectiveAppWT * self.weightEl
            + self.avgUrgentScanWt * self.weightUr
        )

        return {
            "avgElectiveAppWT": self.avgElectiveAppWT,
            "avgElectiveScanWT": self.avgElectiveScanWT,
            "avgUrgentScanWT": self.avgUrgentScanWt,
            "avgOT": self.avgOT,
            "replicationOV": replicationOV,
            "weeklyOV": weeklyOV,
            "weeklyElectiveCounts": weeklyElectiveCounts,
            "weeklyUrgentCounts": weeklyUrgentCounts,
        }

    # ============================================================
    # Original simulation logic
    # ============================================================

    def generatePatients(self) -> None:
        """
        Create new patients and add them to the list of patients for the current simulation object.
        """
        arrivalTimeNext = 0.0
        counter = 0

        for w in range(self.W):
            for d in range(self.D):

                if d < self.D - 1:
                    arrivalTimeNext = 8 + Exponential_distribution(self.lambdaElective) * (17 - 8)

                    while arrivalTimeNext < 17:
                        tardiness = Normal_distribution(self.meanTardiness, self.stdevTardiness) / 60
                        noShow = Bernouilli_distribution(self.probNoShow)
                        duration = Normal_distribution(self.meanElectiveDuration, self.stdevElectiveDuration) / 60

                        self.patients.append(
                            Patient(counter, 1, 0, w, d, arrivalTimeNext, tardiness, noShow, duration)
                        )

                        counter += 1
                        arrivalTimeNext += Exponential_distribution(self.lambdaElective) * (17 - 8)

                lmbd = self.lambdaUrgent[0]
                endTime = 17

                if (d == 3) or (d == 5):
                    lmbd = self.lambdaUrgent[1]
                    endTime = 12

                arrivalTimeNext = 8 + Exponential_distribution(lmbd) * (endTime - 8)

                while arrivalTimeNext < endTime:
                    noShow = 0
                    tardiness = 0
                    scanType = self.getRandomScanType()
                    duration = Normal_distribution(
                        self.meanUrgentDuration[scanType],
                        self.stdevUrgentDuration[scanType]
                    ) / 60

                    self.patients.append(
                        Patient(counter, 2, scanType, w, d, arrivalTimeNext, tardiness, noShow, duration)
                    )

                    counter += 1
                    arrivalTimeNext += Exponential_distribution(lmbd) * (endTime - 8)

    def getRandomScanType(self) -> int:
        r = random.random()

        for idx, prob in enumerate(self.cumulativeProbUrgentType):
            if r < prob:
                return idx

        return len(self.cumulativeProbUrgentType) - 1

    def getNextSlotNrFromTime(self, day: int, patientType: int, time: float) -> int:
        for s in range(self.S):
            if (
                self.weekSchedule[day][s].appTime > time
                and patientType == self.weekSchedule[day][s].patientType
            ):
                return s

        print(f"NO SLOT EXISTS DURING TIME {time} \n")
        exit(0)

    @staticmethod
    def sortPatientsOnAppTime(patient1: Patient, patient2: Patient) -> int:
        if (patient1.scanWeek == -1) and (patient2.scanWeek == -1):
            if patient1.callWeek < patient2.callWeek:
                return -1
            if patient1.callWeek > patient2.callWeek:
                return 1
            if patient1.callDay < patient2.callDay:
                return -1
            if patient1.callDay > patient2.callDay:
                return 1
            if patient1.callTime < patient2.callTime:
                return -1
            if patient1.callTime > patient2.callTime:
                return 1
            if patient1.patientType == 2:
                return -1
            if patient2.patientType == 2:
                return 1
            return 0

        if patient1.scanWeek == -1:
            return 1

        if patient2.scanWeek == -1:
            return -1

        if patient1.scanWeek < patient2.scanWeek:
            return -1
        if patient1.scanWeek > patient2.scanWeek:
            return 1
        if patient1.scanDay < patient2.scanDay:
            return -1
        if patient1.scanDay > patient2.scanDay:
            return 1
        if patient1.appTime < patient2.appTime:
            return -1
        if patient1.appTime > patient2.appTime:
            return 1
        if patient1.patientType == 2:
            return -1
        if patient2.patientType == 2:
            return 1
        if patient1.nr < patient2.nr:
            return -1
        if patient1.nr > patient2.nr:
            return 1

        return 0

    @staticmethod
    def sortPatients(patient1: Patient, patient2: Patient) -> int:
        if patient1.callWeek < patient2.callWeek:
            return -1
        if patient1.callWeek > patient2.callWeek:
            return 1
        if patient1.callDay < patient2.callDay:
            return -1
        if patient1.callDay > patient2.callDay:
            return 1
        if patient1.callTime < patient2.callTime:
            return -1
        if patient1.callTime > patient2.callTime:
            return 1
        if patient1.patientType == 2:
            return -1
        if patient2.patientType == 2:
            return 1

        return 0

    def schedulePatients(self) -> None:
        self.patients = sorted(self.patients, key=cmp_to_key(Simulation.sortPatients))

        week = [0, 0]
        day = [0, 0]
        slot = [0, 0]

        for s in range(self.S):
            if self.weekSchedule[0][s].patientType == 1:
                day[0] = 0
                slot[0] = s
                break

        for s in range(self.S):
            if self.weekSchedule[0][s].patientType == 2:
                day[1] = 0
                slot[1] = s
                break

        previousWeek = 0
        numberOfElectivePerWeek = 0
        numberOfElective = 0

        for patient in self.patients:
            i = patient.patientType - 1

            if week[i] < self.W:
                if patient.callWeek > week[i]:
                    week[i] = patient.callWeek
                    day[i] = 0
                    slot[i] = self.getNextSlotNrFromTime(day[i], patient.patientType, 0)

                elif (patient.callWeek == week[i]) and (patient.callDay > day[i]):
                    day[i] = patient.callDay
                    slot[i] = self.getNextSlotNrFromTime(day[i], patient.patientType, 0)

                if (
                    patient.callWeek == week[i]
                    and patient.callDay == day[i]
                    and patient.callTime >= self.weekSchedule[day[i]][slot[i]].appTime
                ):
                    for s in range(self.S - 1, -1, -1):
                        if self.weekSchedule[day[i]][s].patientType == patient.patientType:
                            slotNr = s
                            break

                    if (patient.patientType == 2) or (
                        patient.callTime < self.weekSchedule[day[i]][slotNr].appTime
                    ):
                        slot[i] = self.getNextSlotNrFromTime(day[i], patient.patientType, patient.callTime)
                    else:
                        if day[i] < self.D - 1:
                            day[i] = day[i] + 1
                        else:
                            day[i] = 0
                            week[i] += 1

                        if week[i] < self.W:
                            slot[i] = self.getNextSlotNrFromTime(day[i], patient.patientType, 0)

                patient.scanWeek = week[i]
                patient.scanDay = day[i]
                patient.slotNr = slot[i]
                patient.appTime = self.weekSchedule[day[i]][slot[i]].appTime

                if patient.patientType == 1:
                    if previousWeek < week[i]:
                        if numberOfElectivePerWeek > 0:
                            self.movingAvgElectiveAppWT[previousWeek] /= numberOfElectivePerWeek

                        numberOfElectivePerWeek = 0
                        previousWeek = week[i]

                    wt = patient.getAppWT()
                    self.movingAvgElectiveAppWT[week[i]] += wt
                    numberOfElectivePerWeek += 1
                    self.avgElectiveAppWT += wt
                    numberOfElective += 1

                found = False
                startD = day[i]
                startS = slot[i] + 1

                for w in range(week[i], self.W):
                    for d in range(startD, self.D):
                        for s in range(startS, self.S):
                            if self.weekSchedule[d][s].patientType == patient.patientType:
                                week[i] = w
                                day[i] = d
                                slot[i] = s
                                found = True
                                break

                        if found:
                            break

                        startS = 0

                    if found:
                        break

                    startD = 0

                if not found:
                    week[i] = self.W

        if numberOfElectivePerWeek > 0:
            self.movingAvgElectiveAppWT[self.W - 1] /= numberOfElectivePerWeek

        if numberOfElective > 0:
            self.avgElectiveAppWT /= numberOfElective

    def runOneSimulation(self) -> None:
        self.generatePatients()
        self.schedulePatients()
        self.patients = sorted(self.patients, key=cmp_to_key(Simulation.sortPatientsOnAppTime))

        prevWeek = 0
        prevDay = -1
        numberOfPatientsWeek = [0, 0]
        numberOfPatients = [0, 0]
        prevScanEndTime = 0
        prevIsNoShow = False

        tard = 0

        for patient in self.patients:
            if patient.scanWeek == -1:
                break

            arrivalTime = patient.appTime + patient.tardiness

            if not patient.isNoShow:
                if (patient.scanWeek != prevWeek) or (patient.scanDay != prevDay):
                    patient.scanTime = arrivalTime
                else:
                    if prevIsNoShow:
                        patient.scanTime = max(
                            self.weekSchedule[patient.scanDay][patient.slotNr].startTime,
                            max(prevScanEndTime, arrivalTime)
                        )
                    else:
                        patient.scanTime = max(prevScanEndTime, arrivalTime)

                wt = patient.getScanWT()

                if patient.patientType == 1:
                    self.movingAvgElectiveScanWT[patient.scanWeek] += wt
                    self.avgElectiveScanWT += wt
                else:
                    self.movingAvgUrgentScanWT[patient.scanWeek] += wt
                    self.avgUrgentScanWt += wt

                numberOfPatientsWeek[patient.patientType - 1] += 1
                numberOfPatients[patient.patientType - 1] += 1

            if (prevDay > -1) and (prevDay != patient.scanDay):
                if (prevDay == 3) or (prevDay == 5):
                    self.movingAvgOT[prevWeek] += max(0, prevScanEndTime - 13)
                    self.avgOT += max(0.0, prevScanEndTime - 13)
                else:
                    self.movingAvgOT[prevWeek] += max(0, prevScanEndTime - 17)
                    self.avgOT += max(0.0, prevScanEndTime - 17)

            if prevWeek != patient.scanWeek:
                if numberOfPatientsWeek[0] > 0:
                    self.movingAvgElectiveScanWT[prevWeek] /= numberOfPatientsWeek[0]

                if numberOfPatientsWeek[1] > 0:
                    self.movingAvgUrgentScanWT[prevWeek] /= numberOfPatientsWeek[1]

                self.movingAvgOT[prevWeek] /= self.D

                numberOfPatientsWeek[0] = 0
                numberOfPatientsWeek[1] = 0

            if patient.isNoShow:
                prevIsNoShow = True

                if (patient.scanWeek != prevWeek) or (patient.scanDay != prevDay):
                    prevScanEndTime = self.weekSchedule[patient.scanDay][patient.slotNr].startTime
            else:
                prevScanEndTime = patient.scanTime + patient.duration
                prevIsNoShow = False

            prevWeek = patient.scanWeek
            prevDay = patient.scanDay
            tard += patient.tardiness

        if numberOfPatientsWeek[0] > 0:
            self.movingAvgElectiveScanWT[self.W - 1] /= numberOfPatientsWeek[0]

        if numberOfPatientsWeek[1] > 0:
            self.movingAvgUrgentScanWT[self.W - 1] /= numberOfPatientsWeek[1]

        self.movingAvgOT[self.W - 1] /= self.D

        if numberOfPatients[0] > 0:
            self.avgElectiveScanWT /= numberOfPatients[0]

        if numberOfPatients[1] > 0:
            self.avgUrgentScanWt /= numberOfPatients[1]

        self.avgOT /= self.D * self.W

    def setWeekSchedule(self) -> None:
        with open(self.inputFileName, 'r', encoding='utf-8-sig') as r:
            slotTypes = list(map(lambda x: re.findall('[0-9]', x), r.readlines()))

            assert len(slotTypes) == 32, "Error: there should be 32 slots (lines) in the file"

            for slotIdx, weekSlot in enumerate(slotTypes):
                assert len(weekSlot) == self.D, f"Error: there should be {self.D} days in the file (columns)"

                for slotDayIdx, inputInteger in enumerate(weekSlot):
                    self.weekSchedule[slotDayIdx][slotIdx].slotType = int(inputInteger)
                    self.weekSchedule[slotDayIdx][slotIdx].patientType = int(inputInteger)

        for d in range(self.D):
            for s in range(32, self.S):
                self.weekSchedule[d][s].slotType = 3
                self.weekSchedule[d][s].patientType = 2

        for d in range(self.D):
            time = 8

            session_elective_count = 0
            session_start_time = 8
            in_afternoon = False

            B = 2
            block_elective_count = 0
            block_start_time = None

            for s in range(self.S):
                if time == 13 and not in_afternoon:
                    in_afternoon = True
                    session_elective_count = 0
                    session_start_time = 13
                    block_elective_count = 0
                    block_start_time = None

                self.weekSchedule[d][s].startTime = time

                if self.weekSchedule[d][s].slotType != 1:
                    self.weekSchedule[d][s].appTime = time
                else:
                    if self.rule == 1:
                        self.weekSchedule[d][s].appTime = time

                    elif self.rule == 2:
                        K = 2

                        if session_elective_count < K:
                            self.weekSchedule[d][s].appTime = session_start_time
                        else:
                            self.weekSchedule[d][s].appTime = time - self.slotLength

                        session_elective_count += 1

                    elif self.rule == 3:
                        if block_elective_count % B == 0:
                            block_start_time = time

                        self.weekSchedule[d][s].appTime = block_start_time
                        block_elective_count += 1

                    elif self.rule == 4:
                        alpha = 0.5
                        self.weekSchedule[d][s].appTime = time - alpha * (self.stdevElectiveDuration / 60)

                time += self.slotLength

                if time == 12:
                    time = 13

    def resetSystem(self) -> None:
        self.patients = list()
        self.avgElectiveAppWT = 0.0
        self.avgElectiveScanWT = 0.0
        self.avgUrgentScanWt = 0.0
        self.avgOT = 0.0
        self.numberOfElectivePatientsPlanned = 0
        self.numberOfUrgentPatientsPlanned = 0

        self.movingAvgElectiveAppWT = []
        self.movingAvgElectiveScanWT = []
        self.movingAvgUrgentScanWT = []
        self.movingAvgOT = []

        for w in range(self.W):
            self.movingAvgElectiveAppWT.append(0.0)
            self.movingAvgElectiveScanWT.append(0.0)
            self.movingAvgUrgentScanWT.append(0.0)
            self.movingAvgOT.append(0.0)

    # ============================================================
    # Integrated antithetic + control variables version
    # Same output format as Welch
    # ============================================================

    def runSimulations(self) -> None:
        """
        Function that runs all simulations.

        This version keeps the same output as Welch:
        - same print table
        - same Excel format

        But now the weekly OV written to Excel is:
        antithetic average + control variate corrected.
        """

        self.setWeekSchedule()

        normalResults = []
        antitheticResults = []

        for r in range(self.R):
            normalResult = self.runOneSimulationWithSeed(seed=r, useAntithetic=False)
            antitheticResult = self.runOneSimulationWithSeed(seed=r, useAntithetic=True)

            normalResults.append(normalResult)
            antitheticResults.append(antitheticResult)

        antitheticAverages = []

        for r in range(self.R):
            normal = normalResults[r]
            anti = antitheticResults[r]

            weeklyOV = []
            weeklyElectiveCounts = []
            weeklyUrgentCounts = []

            for week in range(self.W):
                weeklyOV.append(
                    0.5 * (normal["weeklyOV"][week] + anti["weeklyOV"][week])
                )

                weeklyElectiveCounts.append(
                    0.5 * (
                        normal["weeklyElectiveCounts"][week]
                        + anti["weeklyElectiveCounts"][week]
                    )
                )

                weeklyUrgentCounts.append(
                    0.5 * (
                        normal["weeklyUrgentCounts"][week]
                        + anti["weeklyUrgentCounts"][week]
                    )
                )

            antitheticAverages.append({
                "avgElectiveAppWT": 0.5 * (
                    normal["avgElectiveAppWT"] + anti["avgElectiveAppWT"]
                ),
                "avgElectiveScanWT": 0.5 * (
                    normal["avgElectiveScanWT"] + anti["avgElectiveScanWT"]
                ),
                "avgUrgentScanWT": 0.5 * (
                    normal["avgUrgentScanWT"] + anti["avgUrgentScanWT"]
                ),
                "avgOT": 0.5 * (
                    normal["avgOT"] + anti["avgOT"]
                ),
                "replicationOV": 0.5 * (
                    normal["replicationOV"] + anti["replicationOV"]
                ),
                "weeklyOV": weeklyOV,
                "weeklyElectiveCounts": weeklyElectiveCounts,
                "weeklyUrgentCounts": weeklyUrgentCounts,
            })

        allWeeklyOV = []
        allWeeklyElectiveCounts = []
        allWeeklyUrgentCounts = []

        for r in range(self.R):
            for week in range(self.W):
                allWeeklyOV.append(antitheticAverages[r]["weeklyOV"][week])
                allWeeklyElectiveCounts.append(
                    antitheticAverages[r]["weeklyElectiveCounts"][week]
                )
                allWeeklyUrgentCounts.append(
                    antitheticAverages[r]["weeklyUrgentCounts"][week]
                )

        varElective = Simulation.sample_variance(allWeeklyElectiveCounts)
        varUrgent = Simulation.sample_variance(allWeeklyUrgentCounts)

        if varElective == 0:
            cElective = 0.0
        else:
            cElective = (
                Simulation.sample_covariance(allWeeklyOV, allWeeklyElectiveCounts)
                / varElective
            )

        if varUrgent == 0:
            cUrgent = 0.0
        else:
            cUrgent = (
                Simulation.sample_covariance(allWeeklyOV, allWeeklyUrgentCounts)
                / varUrgent
            )

        expectedElectivePerWeek = 5 * self.lambdaElective
        expectedUrgentPerWeek = (
            4 * self.lambdaUrgent[0]
            + 2 * self.lambdaUrgent[1]
        )

        weeklyOVPerReplication: list[list[float]] = []

        for r in range(self.R):
            correctedWeeklyOV = []

            for week in range(self.W):
                x = antitheticAverages[r]["weeklyOV"][week]
                yElective = antitheticAverages[r]["weeklyElectiveCounts"][week]
                yUrgent = antitheticAverages[r]["weeklyUrgentCounts"][week]

                correctedX = (
                    x
                    - cElective * (yElective - expectedElectivePerWeek)
                    - cUrgent * (yUrgent - expectedUrgentPerWeek)
                )

                correctedWeeklyOV.append(correctedX)

            weeklyOVPerReplication.append(correctedWeeklyOV)

        electiveAppWT: float = 0
        electiveScanWT: float = 0
        urgentScanWT: float = 0
        OT: float = 0
        OV: float = 0

        print("r \t elAppWT \t elScanWT \t urScanWT \t OT \t OV \n")

        for r in range(self.R):
            avgCorrectedOV = Simulation.sample_mean(weeklyOVPerReplication[r])

            electiveAppWT += antitheticAverages[r]["avgElectiveAppWT"]
            electiveScanWT += antitheticAverages[r]["avgElectiveScanWT"]
            urgentScanWT += antitheticAverages[r]["avgUrgentScanWT"]
            OT += antitheticAverages[r]["avgOT"]
            OV += avgCorrectedOV

            print(
                f"{r} \t "
                f"{antitheticAverages[r]['avgElectiveAppWT']:.2f} \t\t "
                f"{antitheticAverages[r]['avgElectiveScanWT']:.5f} \t "
                f"{antitheticAverages[r]['avgUrgentScanWT']:.2f} \t\t "
                f"{antitheticAverages[r]['avgOT']:.2f} \t "
                f"{avgCorrectedOV:.2f}"
            )

        electiveAppWT /= self.R
        electiveScanWT /= self.R
        urgentScanWT /= self.R
        OT /= self.R
        OV /= self.R

        print("--------------------------------------------------------------------------------")
        print(
            f"AVG: \t {electiveAppWT:.2f} \t\t "
            f"{electiveScanWT:.5f} \t "
            f"{urgentScanWT:.2f} \t\t "
            f"{OT:.2f} \t "
            f"{OV:.2f} \n"
        )

        print(f"Control coefficient elective: {cElective:.8f}")
        print(f"Control coefficient urgent:   {cUrgent:.8f}")

        self.writeWeeklyObjectiveValuesToExcel(weeklyOVPerReplication)


if __name__ == "__main__":
    sim = Simulation("Big Assignment/Inputs/input-S3-16.txt", 2000, 5, 3)
    sim.runSimulations()