import random

def birthday_probability(nr_days = 365, runs = 10000):
    random.seed(42)
    feesten = [13, 23, 33, 53]
    for feest in feesten:
        array_running_average = []
        array_succes = []
        for run in range(1,runs +1):
            verjaardagen = []
            for  people in range(1, feest +1):
                verjaardagen.append(random.randint(1, nr_days)) #check if this works untill 365
            
            if len(verjaardagen) != len(set(verjaardagen)):
                array_succes.append(1)
            else:
                array_succes.append(0)
             
            array_running_average.append(sum(array_succes)/run)
        print(f"Feest met {feest} mensen heeft een kans van {array_running_average[-1]:.4f} dat er 2 mensen op dezelfde dag jarig zijn.")
        #print(f"array_succes: {array_succes}")
    


birthday_probability(runs=10000)