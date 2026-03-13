import random
from openpyxl import Workbook
random.seed(42)

def birthday_probability(nr_days=365, runs=10000, filename="Probabilities365.xlsx"):
    parties = list(range(1, 367))
    
    for party in parties:
        array_running_average = []
        array_succes = []

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
        print(array_running_average[-1])

birthday_probability(runs=10000)
