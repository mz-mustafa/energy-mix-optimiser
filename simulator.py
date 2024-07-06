from project import Project
from scenario import Scenario
from sources2 import Source, SourceManager
import pandas as pd
import time

source_manager = None

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

def set_baseline_src_config():
    
    #configure(self, start_year, end_year,rating, rating_unit, 
    #spin_reserve, priority, min_loading, max_loading):
    
    #5MW solar plant (SRC_4)
    #9 x 1.5MW Captive Gas Generators (SRC_6)
    sources = []
    
    #==========================================
    #5MW existing solar plant SRC_4
    existing_solar_src_list = []
    for i in range(1,2):
        
        src = source_manager.get_source_types_by_name('SRC_4')
        src.configure(start_year=1, end_year = 12,rating=5, rating_unit='MW', spin_reserve=0, 
                    priority=1, min_loading=0, max_loading=100)
        existing_solar_src_list.append(src)
    
    sources.extend(existing_solar_src_list)
    
    #==========================================
    
    #5MW new solar plant
    
    solar_src_list = []
    for i in range(1,9):
        
        solar_src =  source_manager.get_source_types_by_name('SRC_1')
        solar_src.configure(start_year=1, end_year = 12, rating=5, rating_unit='MW', spin_reserve=0, 
                  priority=1, min_loading=0, max_loading=100)
        solar_src_list.append(solar_src)
    
    sources.extend(solar_src_list)

    solar_src_list = []
    for i in range(1,2):
        
        solar_src =  source_manager.get_source_types_by_name('SRC_1')
        solar_src.configure(start_year=3, end_year = 12, rating=5, rating_unit='MW', spin_reserve=0, 
                  priority=1, min_loading=0, max_loading=100)
        solar_src_list.append(solar_src)
    
    sources.extend(solar_src_list)
    
    solar_src_list = []
    for i in range(1,2):
        
        solar_src =  source_manager.get_source_types_by_name('SRC_1')
        solar_src.configure(start_year=5, end_year = 12, rating=5, rating_unit='MW', spin_reserve=0, 
                  priority=1, min_loading=0, max_loading=100)
        solar_src_list.append(solar_src)
    
    sources.extend(solar_src_list)

    solar_src_list = []
    for i in range(1,2):
        
        solar_src =  source_manager.get_source_types_by_name('SRC_1')
        solar_src.configure(start_year=8, end_year = 12, rating=5, rating_unit='MW', spin_reserve=0, 
                  priority=1, min_loading=0, max_loading=100)
        solar_src_list.append(solar_src)
    
    sources.extend(solar_src_list)
    """
    solar_src_list = []
    for i in range(1,2):
        
        solar_src =  source_manager.get_source_types_by_name('SRC_1')
        solar_src.configure(start_year=10, end_year = 12, rating=5, rating_unit='MW', spin_reserve=0, 
                  priority=1, min_loading=0, max_loading=100)
        solar_src_list.append(solar_src)
    
    sources.extend(solar_src_list)
    """
    
    
    #==========================================
    
    #1.5MW captive generators SRC_6
    captive_src_list = []
    
    for i in range(1,7):
        
        captive_src =  source_manager.get_source_types_by_name('SRC_6')
        captive_src.configure(start_year=1, end_year = 12, rating=1.5, rating_unit='MW', spin_reserve=100, 
                  priority=3, min_loading=10, max_loading=100)
        captive_src_list.append(captive_src)
    
    sources.extend(captive_src_list)
    
    #==========================================
    """
    #PPA HFO generators SRC_2
    hfo_src_list = []
    for i in range(1,2):
        
        hfo_src =  source_manager.get_source_types_by_name('SRC_2')
        hfo_src.configure(start_year=1, end_year = 12, rating=4, rating_unit='MW', spin_reserve=100, 
                  priority=2, min_loading=10, max_loading=100)
        hfo_src_list.append(hfo_src)
    
    sources.extend(hfo_src_list)


    hfo_src_list = []
    for i in range(1,2):
        
        hfo_src =  source_manager.get_source_types_by_name('SRC_2')
        hfo_src.configure(start_year=9, end_year = 12, rating=4, rating_unit='MW', spin_reserve=0, 
                  priority=3, min_loading=25, max_loading=100)
        hfo_src_list.append(hfo_src)
    
    sources.extend(hfo_src_list)
    """
        
    #==========================================
    """
    
    #10MW steam (SRC_5)
    src = source_manager.get_source_types_by_name('SRC_5')
    src.configure(start_year=3, end_year = 12,rating=10, rating_unit='MW', spin_reserve=20, 
                  priority=0, min_loading=0, max_loading=60)
    sources.append(src)
    """
    
    #==========================================
    
    #6 x 0.5MW BESS (SRC_3)
    bess_src_list = []
    for i in range(1,13):
        bess_src = source_manager.get_source_types_by_name('SRC_3')
        bess_src.configure(start_year=1, end_year = 12, rating = 5, 
                           rating_unit='MWh', spin_reserve=0, priority=2, min_loading=0, max_loading=100)
        bess_src_list.append(bess_src)
    sources.extend(bess_src_list)

    bess_src_list = []
    for i in range(1,2):
        bess_src = source_manager.get_source_types_by_name('SRC_3')
        bess_src.configure(start_year=3, end_year = 12, rating = 5, 
                           rating_unit='MWh', spin_reserve=0, priority=2, min_loading=0, max_loading=100)
        bess_src_list.append(bess_src)
    sources.extend(bess_src_list)
    
    bess_src_list = []
    for i in range(1,3):
        bess_src = source_manager.get_source_types_by_name('SRC_3')
        bess_src.configure(start_year=5, end_year = 12, rating = 5, 
                           rating_unit='MWh', spin_reserve=0, priority=2, min_loading=0, max_loading=100)
        bess_src_list.append(bess_src)
    sources.extend(bess_src_list)

    bess_src_list = []
    for i in range(1,3):
        bess_src = source_manager.get_source_types_by_name('SRC_3')
        bess_src.configure(start_year=8, end_year = 12, rating = 5, 
                           rating_unit='MWh', spin_reserve=0, priority=2, min_loading=0, max_loading=100)
        bess_src_list.append(bess_src)
    sources.extend(bess_src_list)

    bess_src_list = []
    for i in range(1,3):
        bess_src = source_manager.get_source_types_by_name('SRC_3')
        bess_src.configure(start_year=10, end_year = 12, rating = 5, 
                           rating_unit='MWh', spin_reserve=0, priority=2, min_loading=0, max_loading=100)
        bess_src_list.append(bess_src)
    sources.extend(bess_src_list)

    return sources
    

# Main program
if __name__ == "__main__":

    start_time = time.time()
    if read_prereq_data():
        print('Prereq data is loaded')
        output_filepath = 'data/summary_output.xlsx'
        src_list = set_baseline_src_config()
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
        print('simulation complete.')
        print('printing yearly summary')
        sc.write_yearly_data_to_csv2(output_filepath)
        print('printing complete')
        print('printing hourly data')
        #sc.write_hourly_data_to_csv()
        print('printing complete')
        print('Scenario KPIs below:')
        print(sc.scenario_kpis)
        sc.aggregate_power_output_by_source_and_year()

    end_time = time.time()
    print(f"Run completed in {end_time - start_time} seconds")

