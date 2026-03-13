import random
from openpyxl import Workbook

random.seed(42)


def make_cumulative_matrix(matrix):
    cumulative_matrix = []

    for row in matrix:
        cumulative_row = []
        running_sum = 0

        for value in row:
            running_sum += value
            cumulative_row.append(running_sum) #gewoon SOM?

        cumulative_matrix.append(cumulative_row)

    return cumulative_matrix


def get_next_state(cumulative_row):
    r = random.random()

    for i in range(len(cumulative_row)):
        if r <= cumulative_row[i]:
            return i

    return len(cumulative_row) - 1


def machine_simulation(number_periods=1000, runs=20, filename="output_machine_simulation.xlsx"):

    policies = {
        "policy_0": {
            "transition_matrix": [
                [0.1, 0.9, 0.0, 0.0],
                [0.2, 0.0, 0.8, 0.0],
                [0.5, 0.0, 0.0, 0.5],
                [1.0, 0.0, 0.0, 0.0]
            ],
            "failure_probs": [0.1, 0.2, 0.5, None]  # None = certain fail/replace, cost 1500
        },

        "policy_1": {
            "transition_matrix": [
                [0.1, 0.9, 0.0, 0.0],
                [0.2, 0.0, 0.8, 0.0],
                [0.5, 0.0, 0.0, 0.5],
                [1.0, 0.0, 0.0, 0.0]
            ],
            "failure_probs": [0.1, 0.2, 0.5, "replace"]  # "replace" = voluntary, cost 500
        },

        "policy_2": {
            "transition_matrix": [
                [0.1, 0.9, 0.0],
                [0.2, 0.0, 0.8],
                [1.0, 0.0, 0.0]
            ],
            "failure_probs": [0.1, 0.2, "replace"]
        },

        "policy_3": {
            "transition_matrix": [
                [0.1, 0.9],
                [1.0, 0.0]
            ],
            "failure_probs": [0.1, "replace"]
        }
    }

    wb = Workbook()
    first_sheet = True

    for policy_name, policy_data in policies.items():

        transition_matrix = policy_data["transition_matrix"]
        failure_probs = policy_data["failure_probs"]
        cumulative_matrix = make_cumulative_matrix(transition_matrix)

        average_costs = []

        if first_sheet:
            ws = wb.active
            ws.title = policy_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=policy_name)

        ws.append(["Run", "Average monthly cost", "Running average"])

        for run in range(1, runs + 1):

            current_state = 0
            total_cost = 0

            for period in range(number_periods):

                p = failure_probs[current_state]

                if p == "replace":
                    # voluntary replacement — always 500, no failure risk
                    total_cost += 500
                elif p is None or random.random() < p:
                    # certain failure (None) or sampled failure
                    total_cost += 1500
                # else: survived, cost = 0

                current_state = get_next_state(
                    cumulative_matrix[current_state]
                )

            average_monthly_cost = total_cost / number_periods
            average_costs.append(average_monthly_cost)

            running_average = sum(average_costs) / run

            ws.append([run, average_monthly_cost, running_average])

        ws.append([])
        ws.append(["Final estimate", sum(average_costs) / runs])

        print(f"{policy_name} estimated average monthly cost: {sum(average_costs) / runs:.4f}")

    wb.save(filename)
    print(f"Results written to {filename}")


machine_simulation(number_periods=1000, runs=10000)