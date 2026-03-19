import random
from openpyxl import Workbook


def machine_simulation(number_periods=1000, runs=20, filename="output_machine_simulation3.xlsx"):

    policies = {
        "policy_0": {
            "failure_probs": [0.1, 0.2, 0.5, None],
            "warmup": 1000
        },
        "policy_1": {
            "failure_probs": [0.1, 0.2, 0.5, "replace"],
            "warmup": 1500
        },
        "policy_2": {
            "failure_probs": [0.1, 0.2, "replace"],
            "warmup": 1500
        },
        "policy_3": {
            "failure_probs": [0.1, "replace"],
            "warmup": 1500
        }
    }

    wb = Workbook()
    first_sheet = True

    for policy_name, policy_data in policies.items():

        failure_probs = policy_data["failure_probs"]
        warmup_periods = policy_data["warmup"]
        total_periods = warmup_periods + number_periods

        average_costs = []
        average_costs_warmup = []

        if first_sheet:
            ws = wb.active
            ws.title = policy_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=policy_name)

        ws.append([
            "Run",
            "Avg monthly cost", "Running avg",
            f"Avg monthly cost (skip {warmup_periods} warmup)", "Running avg (warmup)"
        ])

        for run in range(1, runs + 1):

            random.seed(run)  # Elke run krijgt zijn eigen seed
            current_state = 0
            total_cost = 0
            post_warmup_cost = 0

            for period in range(total_periods):

                p = failure_probs[current_state]

                if p == "replace":
                    cost = 500
                    current_state = 0
                else:
                    r = random.random()
                    if p is None or r < p:
                        cost = 1500
                        current_state = 0
                    else:
                        cost = 0
                        current_state += 1

                total_cost += cost
                if period >= warmup_periods:
                    post_warmup_cost += cost

            # Regular stats (over all number_periods — excluding warmup from denominator too)
            avg_cost = total_cost / total_periods
            average_costs.append(avg_cost)
            running_avg = sum(average_costs) / run

            # Warmup-adjusted stats (only post-warmup periods)
            avg_cost_warmup = post_warmup_cost / number_periods
            average_costs_warmup.append(avg_cost_warmup)
            running_avg_warmup = sum(average_costs_warmup) / run

            ws.append([run, avg_cost, running_avg, avg_cost_warmup, running_avg_warmup])

        final_avg = sum(average_costs) / runs
        final_avg_warmup = sum(average_costs_warmup) / runs

        ws.append([])
        ws.append(["Final estimate", final_avg, "", final_avg_warmup])

        print(f"{policy_name} | Full avg: {final_avg:.4f} | Warmup-adjusted avg: {final_avg_warmup:.4f}")

    wb.save(filename)
    print(f"Results written to {filename}")


machine_simulation(number_periods=10000, runs=100)