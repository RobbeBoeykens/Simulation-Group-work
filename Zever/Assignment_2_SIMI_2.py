import random
from openpyxl import Workbook

random.seed(42)


def machine_simulation(number_periods=1000, runs=20, filename="output_machine_simulation.xlsx"):

    policies = {
        "policy_0": {
            "failure_probs": [0.1, 0.2, 0.5, None]  # None = certain failure, cost 1500
        },
        "policy_1": {
            "failure_probs": [0.1, 0.2, 0.5, "replace"]  # replace = voluntary, cost 500
        },
        "policy_2": {
            "failure_probs": [0.1, 0.2, "replace"]
        },
        "policy_3": {
            "failure_probs": [0.1, "replace"]
        }
    }

    wb = Workbook()
    first_sheet = True

    for policy_name, policy_data in policies.items():

        failure_probs = policy_data["failure_probs"]
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
                    # voluntary replacement — pay 500, reset to S1
                    total_cost += 500
                    current_state = 0
                else:
                    r = random.random()
                    if p is None or r < p:
                        # certain failure (None) or sampled failure — pay 1500, reset to S1
                        total_cost += 1500
                        current_state = 0
                    else:
                        # survived — no cost, advance to next state
                        current_state += 1

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