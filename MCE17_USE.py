import random
from openpyxl import Workbook

def birthday_probability(nr_days=365, runs=10000, filename="output_ex17.xlsx"):
    random.seed(42)
    parties = [13, 23, 33, 53]

    wb = Workbook()
    first_sheet = True

    for party in parties:
        array_running_average = []
        array_succes = []

        if first_sheet:
            ws = wb.active
            ws.title = f"party_{party}"
            first_sheet = False
        else:
            ws = wb.create_sheet(title=f"party_{party}")

        ws.append(["Run", "Amount of people", "Succes (1/0)", "Running average"])

        for run in range(1, runs + 1):
            birthdays = []

            for people in range(1, party + 1):
                birthdays.append(random.randint(1, nr_days))

            if len(birthdays) != len(set(birthdays)):
                array_succes.append(1)
            else:
                array_succes.append(0)

            running_average = sum(array_succes) / run
            array_running_average.append(running_average)

            ws.append([run, party, array_succes[-1], running_average])

        print(f"party with {party} people has a chance of {array_running_average[-1]:.4f} that 2 people share the same birthday.")

        ws.append([])
        ws.append(["End Probability", array_running_average[-1]])

    wb.save(filename)
    print(f"Results written to {filename}")

birthday_probability(runs=10000)
