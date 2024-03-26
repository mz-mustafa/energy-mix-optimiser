import datetime
import openpyxl
import pandas as pd
import math
import random
from project import Project
from itertools import groupby

class Scenario:
    def __init__(self, name, client_name, selected_sources, data_path = "data/"):
        self.name = name
        self.client_name = client_name
        self.timestamp = datetime.datetime.now()
        self.src_list = selected_sources
        self.results = {y: {m: {d: {h: {'unserved_power_req': 0} for h in range(24)} 
                               for d in range(1, 32)} 
                          for m in range(1, 13)} 
                     for y in range(1, 13)}

    
    def has_stable_source(self,year):

        for src in self.src_list:
            # Check if 'stability' in metadata and its value is 'STABLE'
            if src.metadata.get('stability', {}).get('value') == 'STABLE':
                # Check if the source's start year is not later than the given year
                if src.config.get('start_year') <= year:
                    return True  # Found at least one stable source meeting the criteria
        return False  # No stable source found meeting the criteria
    
    def scenario_includes_renewable_src(self, year):
        
        for src in self.src_list:
            
            if src.metadata.get('type', {}).get('value') == 'R':
                # Check if the source's start year is not later than the given year
                if src.config.get('start_year') <= year:
                    return True  # Found at least one stable source meeting the criteria
        return False  # No renewable source found meeting the criteria
    
    def scenario_includes_captive_src(self, year):

        for src in self.src_list:
            
            if src.metadata.get('finance', {}).get('value') == 'CAPTIVE':
                # Check if the source's start year is not later than the given year
                if src.config.get('start_year') <= year:
                    return True  # Found at least one stable source meeting the criteria
        return False  # No renewable source found meeting the criteria

    def calc_src_power_and_energy(self,y,m,d,h,power_req):
        #TO DO probably need to exclude BESS sources here
        for priority, sources in groupby(self.src_list, key=lambda x: x.config['priority']):
            sources = list(sources)
            
            spinning_reserve_req = sources[0].metadata.get('spinning_reserve', {'value': 0})['value']
            current_power_output = 0
            current_power_capacity = 0
            
            for src in sources:
                if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] in [-1, -2, -3]:
                    continue  # Source is not available

                max_loading = src.metadata.get('max_loading', {'value': src.config['rating']})['value'] * src.config['rating']
                power_capacity = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
                
                # Calculate potential contribution without exceeding max_loading or remaining power requirement
                potential_power_output = min(max_loading, power_req - current_power_output, power_capacity)
                
                # Update only if it contributes to meeting power requirement
                if potential_power_output > 0:
                    current_power_output += potential_power_output
                    current_power_capacity += power_capacity

                # Check if power requirement and spinning reserve are met
                if current_power_output >= power_req and (current_power_capacity - current_power_output) >= spinning_reserve_req:
                    # Update ops_data for selected sources and stop selection for this group
                    break

            # Update ops_data for sources considered in this group
            for src in sources:
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = potential_power_output
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = potential_power_output  # Assuming same as power_output
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] = 1
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = power_capacity - potential_power_output
            
            # If the power requirement is met or exceeded, adjust for next group consideration
            power_req = min(0,power_req - current_power_output)
            if power_req == 0:
                
                break  # Exit the function if no more power is needed

        return power_req  # Return the unmet power requirement, if any


    def simulate(self):

        for y in range(1,13):

            for m in range (1,13):

                if m == 2:  # February
                    days = 28
                elif m in [4, 6, 9, 11]:  # April, June, September, November
                    days = 30
                else:  # All other months
                    days = 31

                for d in range (1, days+1):

                    for h in range (0,24):

                        #set the power requirement
                        power_req = Project.load_data[y][m][d][h]
                        #may also include BESS charging
                        for src in self.src_list:
                            if src.metadata['type']['value'] == 'BESS':
                                status = src.ops_data[y][m][d][h]['status']
                                if status == 0:
                                    #Tricket charge
                                    power_req += src.config['rating'] * 0.02
                                elif status == 2:
                                    #Full charge
                                    power_req += src.config['rating'] * 0.5
                        
                        #Consumption of sources, update key results in the scenario
                        self.results[y][m][d][h]['unserved_power_req'] = self.calc_src_power_and_energy(y,m,d,h,power_req)
                        
                        #Can this configuration meet the ramp requirements
                        for src in self.src_list:

                            if src.metadata['type']['value'] != 'BESS' and src.ops_data[y][m][d][h]['status'] == -1:
                                a = 'Todo'
                                #then we need to know the power rating of the source and add it all up.
                                #sources that have failed just now would be what we should have reserve for.
                            if src.metadata['type']['value'] != 'R' and src.ops_data[y][m][d][h]['status'] == 0.5:
                                a = 'Todo'
                                #then find delta between h-1 and h and add it to sudden_drop variable.    

                        #Afther this loop we have the sudden_drop variable and now we need to see if config can serve it.
                        #this check can be in a seperate function.
                        
                        #Use the sorted sources list, form groups.
                        #for the group calc. the spinning reserve
                        #we know the perc of src rating that the group src can take within 4 seconds.
                        #compare this number to the spin reserve
                        #Need to correct, below not fully right.
                        #if spin reserve high, find spin reserve - sudden_drop and load src in group propotionally, exit function
                        #if spin reserve lower, load srcs to full, reduce sudden_drop try with rem src until sudden_drop 0 or end of srcs
                        #if we reach end of srcs and sudden_drop still remains, then we can assume load failure.
                        #record this in the scenario results.
                                
                        #BESS will have a key role in the above.
                        #If BESS present and charging, then use it to gain bonus reduction in sudden_drop equal to its max. op
                        #which is 500kW per unit
                        #but then make its next two hours charging. (so it can't come to the rescue until charged again)


                    #cost calc. at the end, summed to be year level, will then be checked again min off-take.
                    #if min offtake then a differnt charge applied (min-offtake x tariff, otherwise actual consumption x tariff)
                self.calculate_opex(y,m)

                        
    
    def calculate_opex(self,y,m):

        for src in self.src_list:

            if src.ops_data[y][m][1][1] == -3:
                continue
            
            #we may not need this condition at all.
            if src.metadata['finance']['value'] == 'CAPTIVE' and src.metadata['type']['value'] == 'NR':
                
                #src.fixed_opex = src.metadata['fixed_opex']['value'] * src.config[rating] * opex inflation rate if metadata exists else 0
                #src.var_opex = src.ops_data['energy_output'] * src.metadata['var_opex']['value'] * opex inflation if mtadata exists
                #src.fuel_cost = src.[energy op] * src.metadata[op baseline] * metadata[fuel rate] * fuel inflation rate if metadata exists
                #calc depreciation if depriciation rate metadata exists else 0
                #find ppa cost if ppa tariff metadata exists else 0
                #total cost will be sum of all 
                return True
        #after running through the sources, calc number of load failures from results.
        #use that and the per loss value to determine the loss (this can be in propotion to shortfall with critical load?)
        #in results, if there unfed demand then that should also be used (in propotion) to calculate further loss.
                






        
