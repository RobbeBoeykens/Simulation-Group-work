import random
from openpyxl import Workbook

def birthday_probability(nr_days=365, runs=10000, filename="output_ex17.xlsx"):
    random.seed(42)
    feesten = [13, 23, 33, 53]

    wb = Workbook()
    first_sheet = True

    for feest in feesten:
        array_running_average = []
        array_succes = []

        if first_sheet:
            ws = wb.active
            ws.title = f"feest_{feest}"
            first_sheet = False
        else:
            ws = wb.create_sheet(title=f"feest_{feest}")

        ws.append(["Run", "Amount of people", "Succes (1/0)", "Running average"])

        for run in range(1, runs + 1):
            verjaardagen = []

            for people in range(1, feest + 1):
                verjaardagen.append(random.randint(1, nr_days))

            if len(verjaardagen) != len(set(verjaardagen)):
                array_succes.append(1)
            else:
                array_succes.append(0)

            running_average = sum(array_succes) / run
            array_running_average.append(running_average)

            ws.append([run, feest, array_succes[-1], running_average])

        print(f"Feest met {feest} mensen heeft een kans van {array_running_average[-1]:.4f} dat er 2 mensen op dezelfde dag jarig zijn.")

        ws.append([])
        ws.append(["Eindkans", array_running_average[-1]])

    wb.save(filename)
    print(f"Resultaten zijn geschreven naar {filename}")

birthday_probability(runs=10000)
