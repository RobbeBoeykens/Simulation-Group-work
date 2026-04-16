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
            cumulative_row.append(running_sum)

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
            "cost_vector": [150, 300, 750, 1500]
        },

        "policy_1": {
            "transition_matrix": [
                [0.1, 0.9, 0.0, 0.0],
                [0.2, 0.0, 0.8, 0.0],
                [0.5, 0.0, 0.0, 0.5],
                [1.0, 0.0, 0.0, 0.0]
            ],
            "cost_vector": [150, 300, 750, 500]
        },

        "policy_2": {
            "transition_matrix": [
                [0.1, 0.9,0.0],
                [0.2, 0.0,0.8],
                [1.0,0.0,0.0]
            ],
            "cost_vector": [150, 300, 500]
        },

        "policy_3": {
            "transition_matrix": [
                [0.1,0.9],
                [1,0.0],
            ],
            "cost_vector": [150,500]
        }
    }

    wb = Workbook()
    first_sheet = True

    for policy_name, policy_data in policies.items():

        transition_matrix = policy_data["transition_matrix"]
        cost_vector = policy_data["cost_vector"]
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

                total_cost += cost_vector[current_state]

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


machine_simulation(number_periods=1000, runs=20)