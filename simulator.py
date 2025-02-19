from project import Project
from scenario import Scenario
from sources2 import Source, SourceManager
import pandas as pd
import time
import random


# SRC QTY
src_1_qts = [0, 1]
src_2_qts = [0, 2, 3, 4]
src_3_qts = [0, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]
src_4_qts = [1]
src_5_qts = [0, 1, 2]
src_6_qts = [0, 3, 4, 5, 6, 7, 8]

# SRC PRT

src_1_prts = [2]
src_2_prts = [3, 4, 5]
src_4_prts = [1]
src_6_prts = [6]

# SRC START DATE

src_1_str_dt = [1, 2, 3, 4, 5, 6, 7, 8]
src_2_str_dt = [1, 2, 3, 4, 5, 6, 7, 8]
src_3_str_dt = [1, 2, 3, 4, 5, 6, 7, 8]
src_4_str_dt = [1]
src_5_str_dt = [1, 2, 3, 4, 5, 6, 7, 8]
src_6_str_dt = [1]

begin = [0, 3, 9, 36, 39, 43]

population  = [['0' for _ in range(53)] for _ in range(10)]


for pop in range(len(population)):

    zeros_sequence = ['0'] * 53

    random_bit = random.randint(0, 1)

    for _ in range(random_bit):
        zeros_sequence[0] = '1'
        zeros_sequence[1] = '2'
        zeros_sequence[2] = str(random.randint(1, 8))

    src_2_qt = random.choice(src_2_qts)
    zeros_sequence[3] = str(src_2_qt)
    src_2_prt = random.choice(src_2_prts)
    src_2_prts.remove(src_2_prt)
    zeros_sequence[4] = str(src_2_prt)

    for i in range(5, 5+src_2_qt):
        zeros_sequence[i] = str(random.choice(src_2_str_dt))

    src_3_qt = random.choice(src_3_qts)
    zeros_sequence[9] = str(src_3_qt)
    src_3_prt = random.choice(src_2_prts)
    src_2_prts.remove(src_3_prt)
    zeros_sequence[10] = str(src_3_prt)

    for i in range(11, 11+src_3_qt):
        zeros_sequence[i] = str(random.choice(src_3_str_dt))


    zeros_sequence[36] = 1
    zeros_sequence[37] = 1
    zeros_sequence[38] = 1

    src_5_qt = random.choice(src_5_qts)
    zeros_sequence[39] = str(src_5_qt)
    src_5_prt = random.choice(src_2_prts)
    src_2_prts.remove(src_5_prt)
    zeros_sequence[40] = str(src_5_prt)

    for i in range(41, 41+src_5_qt):
        zeros_sequence[i] = str(random.choice(src_5_str_dt))


    src_6_qt = random.choice(src_6_qts)
    zeros_sequence[43] = str(src_6_qt)
    src_6_prt = random.choice(src_6_prts)
    zeros_sequence[44] = str(src_6_prt)

    for i in range(45, 45+src_6_qt):
        zeros_sequence[i] = str(random.choice(src_6_str_dt))

    src_2_prts = [3, 4, 5]
    population[pop] = zeros_sequence

source_manager = None
results = []
select_list = []

nb_gens = 4

def mutation(crossed_indivs, threshold_mutate=0.1):
    for i in range(len(crossed_indivs)):
        prob_mut = random.random()
        if prob_mut <= threshold_mutate:
            prio_1, prio_2, prio_3 = random.sample(src_2_prts, 3)
            crossed_indivs[i][begin[1]+1] = prio_1
            crossed_indivs[i][begin[2]+1] = prio_2
            crossed_indivs[i][begin[4]+1] = prio_3
    return crossed_indivs

def crossover(selected_indivs, threshold_crsvr=0.60):
    while len(selected_indivs)<len(population):
        parent_1 = random.randint(0, len(population)//2-1)
        parent_2 = random.randint(0, len(population)//2-1)
        son_1 = selected_indivs[parent_1]
        son_2 = selected_indivs[parent_2]
        prob_src_1 = random.random()
        prob_src_2 = random.random()
        prob_src_3 = random.random()
        prob_src_5 = random.random()
        prob_src_6 = random.random()
        if (prob_src_1 <= threshold_crsvr):
            hold = son_1[begin[0]:begin[1]]
            son_1[begin[0]:begin[1]] = son_2[begin[0]:begin[1]]
            son_2[begin[0]:begin[1]] = hold

        if (prob_src_2 <= threshold_crsvr):
            hold_1, hold_2 = son_1[begin[1]], son_1[begin[1]+2:begin[2]]
            son_1[begin[1]], son_1[begin[1]+2:begin[2]] = son_2[begin[1]], son_2[begin[1]+2:begin[2]]
            son_2[begin[1]], son_2[begin[1]+2:begin[2]] = hold_1, hold_2
        
        if (prob_src_3 <= threshold_crsvr):
            hold_3, hold_4 = son_1[begin[2]], son_1[begin[2]+2:begin[3]]
            son_1[begin[2]], son_1[begin[2]+2:begin[3]] = son_2[begin[2]], son_2[begin[2]+2:begin[3]]
            son_2[begin[2]], son_2[begin[2]+2:begin[3]] = hold_3, hold_4
        
        if (prob_src_5 <= threshold_crsvr):
            hold_5, hold_6 = son_1[begin[4]], son_1[begin[4]+2:begin[5]]
            son_1[begin[4]], son_1[begin[4]+2:begin[5]] = son_2[begin[4]], son_2[begin[4]+2:begin[5]]
            son_2[begin[4]], son_2[begin[4]+2:begin[5]] = hold_5, hold_6
        
        if (prob_src_6 <= threshold_crsvr):
            hold_0 = son_1[begin[5]:]
            son_1[begin[5]:] = son_2[begin[5]:]
            son_2[begin[5]:] = hold_0
        
        choices = [son_1, son_2]
        selected_indivs.append(random.sample(choices, 1)[0])
    return selected_indivs

def selection(ranked_list, population):
    selected_elems = []
    for i in range(len(population)//2):
        selected_elems.append(population[ranked_list[i][1]])
    return selected_elems

def rank_pop(fitness_values):
    indexed_list = [(element, index) for index, element in enumerate(fitness_values)]
    sorted_indexed_list = sorted(indexed_list, key=lambda x: x[0])
    return [[element, index] for element, index in sorted_indexed_list]


def fitness(individual):
    if individual['Energy Fulfillment Ratio (%)'] >= 99 & individual['Critical Load Interruptions (No.)'] <=1:
        return individual['Average Unit Cost ($/kWh)']*0.9 + individual['Estimated Interruption Loss (M $)'] * 0.1
    else:
        return float('inf')



def read_prereq_data():
    try:
        global source_manager
        print('Reading Pre req data')
        Project.read_load_projection("data")
        Project.read_load_solar_data_from_folder('data/load_solar_profile')
        Project.create_load_data()
        source_manager = SourceManager('data/input_data.xlsx')
        return True
    except Exception as e:
        print('Pre req data could not be loaded.')
        return False

def set_baseline_src_config(population):
    
    #configure(self, start_year, end_year,rating, rating_unit, 
    #spin_reserve, priority, min_loading, max_loading):
    
    #5MW solar plant (SRC_4)
    #9 x 1.5MW Captive Gas Generators (SRC_6)
    sources = []
    
    #==========================================
    #5MW existing solar plant SRC_4

    

    for i in range(6):
        for j in range(int(population[begin[i]])):
            src = source_manager.get_source_types_by_name('SRC_'+str(i+1))
            src.configure(start_year=int(population[begin[i]+j+2]), end_year = 12,rating=5, rating_unit='MWh', spin_reserve=0, 
                        priority=int(population[begin[i]+1]), min_loading=0, max_loading=100)
            sources.append(src)
    
    # print(pop)
    # print(len(sources))
    # for sourc in sources:
    #     print(sourc.name, sourc.config)
    

    return sources
    

# Main program
if __name__ == "__main__":

    start_time = time.time()
    if read_prereq_data():
        print('Prereq data is loaded')
        output_filepath = 'data/summary_output.xlsx'
        nex_gen = population
        for k in range(nb_gens+1):
            results = []
            select_list = []
            for i in range(len(nex_gen)):
                src_list = set_baseline_src_config(nex_gen[i])
                print('sources created and configured')
                #BESS non-Em mode: 0 means none, 1 means yes with equal distribution, 2 means yes with selection utilization
                sc = Scenario(
                    name = "Baseline", 
                    client_name = "Engro",
                    selected_sources=src_list,
                    spin_reserve_perc=0,
                    bess_non_emergency_use=2,
                    bess_charge_hours=1,
                    bess_priority_wise_use=True,
                    charge_ratio_night=2.5
                    )
                print('scenario created, starting simulation')
                
                sc.simulate()
                # print('simulation complete.')
                # print('printing yearly summary')
                # sc.write_yearly_data_to_csv2(output_filepath)
                # print('printing complete')
                # print('printing hourly data')
                # #sc.write_hourly_data_to_csv()
                # print('printing complete')
                # print('Scenario KPIs below:')
                # print(sc.scenario_kpis)
                sc.aggregate_power_output_by_source_and_year()
                results.append(sc.scenario_kpis)
                select_list.append(fitness(sc.scenario_kpis))
            if k < nb_gens:
                nex_gen = mutation(crossover(selection(rank_pop(select_list), nex_gen)))

        

    end_time = time.time()
    print(f"Run completed in {end_time - start_time} seconds")
    # print(results)
    # print(rank_pop(select_list))
    print("----------------------------------------------------")
    print(results)
    print("----------------------------------------------------")
    print(select_list)
    print("----------------------------------------------------")
    print(nex_gen)
