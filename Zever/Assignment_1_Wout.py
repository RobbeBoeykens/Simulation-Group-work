import random
from openpyxl import Workbook

random.seed(42)

K = 10000
nr_days = 365
party_sizes = [13, 23, 33, 53]

wb = Workbook()

for idx, nr_people in enumerate(party_sizes):
    if idx == 0:
        ws = wb.active
        ws.title = f"party_{nr_people}"
    else:
        ws = wb.create_sheet(title=f"party_{nr_people}")

    ws.append(["Run", "Number of people", "Duplicate birthday (1=yes, 0=no)"])

    for run in range(1, K + 1):
        days = [0] * nr_days
        duplicate_found = 0

        for _ in range(nr_people):
            i2 = random.randint(0, nr_days - 1)
            days[i2] += 1

            if days[i2] > 1:
                duplicate_found = 1
                break

        ws.append([run, nr_people, duplicate_found])

wb.save("output_ex17.xlsx")

print("Results written to output_ex17.xlsx")
