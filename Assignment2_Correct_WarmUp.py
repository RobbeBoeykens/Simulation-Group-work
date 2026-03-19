import random
import matplotlib.pyplot as plt
from openpyxl import Workbook


def machine_simulation(number_periods=20000, runs=100, filename="output_machine_simulation3.xlsx"):

    policies = {
        "policy_0": {
            "failure_probs": [0.1, 0.2, 0.5, None],
            "warmup": 3500
        },
        "policy_1": {
            "failure_probs": [0.1, 0.2, 0.5, "replace"],
            "warmup": 4000
        },
        "policy_2": {
            "failure_probs": [0.1, 0.2, "replace"],
            "warmup": 3000
        },
        "policy_3": {
            "failure_probs": [0.1, "replace"],
            "warmup": 3500
        }
    }

    wb = Workbook()
    first_sheet = True

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle("Warmup Detection — Running Average Cost per Run (seeds 1–25)", fontsize=14, fontweight="bold")
    axes = axes.flatten()

    for idx, (policy_name, policy_data) in enumerate(policies.items()):

        failure_probs = policy_data["failure_probs"]
        warmup_periods = policy_data["warmup"]

        average_costs_warmup = []

        if first_sheet:
            ws = wb.active
            ws.title = policy_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=policy_name)

        ws.append([
            "Run",
            f"Avg monthly cost (skip {warmup_periods} warmup)", "Running avg (warmup)"
        ])

        ax = axes[idx]

        for run in range(1, runs + 1):

            random.seed(run)
            current_state = 0
            period_costs = []

            for period in range(number_periods):

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

                period_costs.append(cost)

            post_warmup_costs = period_costs[warmup_periods:]
            avg_cost_warmup = sum(post_warmup_costs) / len(post_warmup_costs)
            average_costs_warmup.append(avg_cost_warmup)
            running_avg_warmup = sum(average_costs_warmup) / run

            ws.append([run, avg_cost_warmup, running_avg_warmup])

            # Plot running average per periode voor de eerste 15 seeds
            if run <= 25:
                running_avg_per_period = []
                cumsum = 0
                for t, cost in enumerate(period_costs):
                    cumsum += cost
                    running_avg_per_period.append(cumsum / (t + 1))

                ax.plot(range(1, number_periods + 1), running_avg_per_period,
                        linewidth=0.8, alpha=0.7, label=f"seed {run}")

        final_avg_warmup = sum(average_costs_warmup) / runs

        ws.append([])
        ws.append(["Final estimate", final_avg_warmup])

        print(f"{policy_name} | Warmup-adjusted avg: {final_avg_warmup:.4f}")

        ax.axvline(x=warmup_periods, color="black", linestyle="--", linewidth=1.5, label=f"Warmup: {warmup_periods}")
        ax.axhline(y=final_avg_warmup, color="red", linestyle="--", linewidth=1.2, label=f"Final avg: {final_avg_warmup:.1f}")
        ax.set_title(f"{policy_name}", fontweight="bold")
        ax.set_xlabel("Periode")
        ax.set_ylabel("Running gem. kostprijs")
        ax.legend(fontsize=6, ncol=2)
        ax.grid(True, alpha=0.3)

    wb.save(filename)
    print(f"Results written to {filename}")

    plt.tight_layout()
    plt.show()


machine_simulation(number_periods=20000, runs=100)
