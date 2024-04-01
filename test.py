from project import Project
from scenario import Scenario
from sources2 import Source, SourceManager
import pandas as pd

source_manager = None

def read_prereq_data():
    try:
        global source_manager
        print('Reading Pre req data')
        Project.read_load_projection("EEM-2/data")
        Project.read_load_solar_data_from_folder('EEM-2/data/load_solar_profile')
        Project.create_load_data()
        source_manager = SourceManager('EEM-2/data/input_data.xlsx')
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
    
    #5MW solar plant (SRC_4)
    src = source_manager.get_source_types_by_name('SRC_4')
    src.configure(start_year=1, end_year = 12,rating=5, rating_unit='MW', spin_reserve=0, 
                  priority=1, min_loading=0, max_loading=100)
    sources.append(src)
    captive_src_list = []
    
    for i in range(1,10):
        
        captive_src =  source_manager.get_source_types_by_name('SRC_6')
        captive_src.configure(start_year=1, end_year = 12, rating=1.5, rating_unit='MW', spin_reserve=20, 
                  priority=2, min_loading=0, max_loading=90)
        captive_src_list.append(captive_src)
    
    sources.extend(captive_src_list)

    return sources

def csv_write(src_list, results):
     
    # Prepare data for DataFrame
    data = []
    # Assuming the first source has all the necessary time periods defined
    first_src_ops_data = src_list[0].ops_data

    for y in first_src_ops_data:
        for m in first_src_ops_data[y]['months']:
            for d in first_src_ops_data[y]['months'][m]['days']:
                for h in first_src_ops_data[y]['months'][m]['days'][d]['hours']:
                    row = [y, m, d, h]
                    for src in src_list:
                        ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                        # Append ops_data values for the source
                        row.extend([
                            ops_data.get('power_capacity', 0),
                            ops_data.get('power_output', 0),
                            ops_data.get('energy_output', 0),
                            ops_data.get('spin_reserve', 0),
                            ops_data.get('status', '')
                        ])
                    # Append results data for the same time period
                    result_data = results[y][m][d][h]
                    row.extend([
                        result_data.get('unserved_power_req', 0),
                        result_data.get('sudden_power_drop', 0),
                        result_data.get('unserved_power_drop', 0),
                        result_data.get('load_shed', 0),
                        result_data.get('log', 0)
                    ])
                    data.append(row)

    # Define column names, including results column names
    column_names = ['Year', 'Month', 'Day', 'Hour']
    for i in range(1, len(src_list) + 1):
        column_names.extend([f'Src_{i}_power_capacity', f'Src_{i}_power_output', f'Src_{i}_energy_output', f'Src_{i}_spin_reserve', f'Src_{i}_status'])
    # Extend column names with results column names
    column_names.extend(['Unserved_Power_Req', 'Sudden_Power_Drop', 'Unserved_Power_Drop', 'Load_Shed', 'Log'])

    # Create DataFrame
    df = pd.DataFrame(data, columns=column_names)

    # Write to CSV
    df.to_csv('src_list_data.csv', index=False)


# Main program
if __name__ == "__main__":
    
    if read_prereq_data():
        print('Prereq data is loaded')

        src_list = set_baseline_src_config()
        print('sources created and configured')
        sc = Scenario(name = "Baseline", client_name = "Engro",selected_sources=src_list)
        print('scenario created, strting simulation')
        
        sc.simulate()
        print('simulation complete.')
        print('printing results')
        csv_write(sc.src_list, sc.results)
        print('printing complete')
def test_print(src_list):
        
    #print(src_list[0].ops_data[1]['months'][5]['days'][10]['hours'][6]['status'])
        print(src_list[5].ops_data[1]['months'][5]['days'][10]['hours'][15]['power_capacity'])
        #print(src_list[0].ops_data[4]['months'][7]['days'][10]['hours'][10]['status'])
        print(src_list[5].ops_data[8]['months'][7]['days'][10]['hours'][10]['power_capacity'])
        print("Year 4 total failures: " , src_list[0].ops_data[4]['year_failures'])
        print("Year 4 total reductions: " , src_list[0].ops_data[4]['year_reductions'])
        print("Year 4 total downtime: " , src_list[0].ops_data[4]['year_downtime'])

"""
src1 = source_manager.get_source_types_by_name('SRC_1')
src2 = source_manager.get_source_types_by_name('SRC_1')
selected_sources = []
selected_sources.append(src1)
selected_sources.append(src2)

for src in selected_sources:

    print(src.metadata['generic_name'])
    print(src.metadata['solar_sudden_drops'])
"""