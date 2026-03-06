import random

random.seed(42)

K = 10000
nr_days = 365
party_sizes = [13, 23, 33, 53]

for nr_people in party_sizes:
    count_success = 0

    for _ in range(K):
        days = [0] * nr_days
        duplicate_found = False

        for _ in range(nr_people):
            i2 = random.randint(0, nr_days - 1)
            days[i2] += 1

            if days[i2] > 1:
                duplicate_found = True
                break

        if duplicate_found:
            count_success += 1

    average = count_success / K
    print(f"For n = {nr_people:2d}, estimated probability = {average:.4f}")