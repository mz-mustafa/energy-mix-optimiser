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
        self.results = {
            y: {
                m: {
                    d: {
                        h: {
                            'unserved_power_req': 0,
                            'sudden_power_drop' : 0,
                            'unserved_power_drop' : 0,
                            'load_shed' : 0,
                            'log' : 0
                            } for h in range(24)
                        } for d in range(1, 32)
                    } for m in range(1, 13)
                } for y in range(1, 13)}

    
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

    def calc_src_power_and_energy(self, y, m, d, h, power_req):

        sudden_power_drop = 0
        # Sort sources by priority for processing
        sorted_sources = sorted(self.src_list, key=lambda x: x.config['priority'])
        # Group sources by priority
        for priority, group in groupby(sorted_sources, key=lambda x: x.config['priority']):
            sources = list(group)
            if not sources:
                continue  # Skip empty groups
            
            # Use spinning reserve requirements from the first source in the group
            spinning_reserve_req = sources[0].metadata.get('spinning_reserve', {'value': 0})['value']
            current_power_output = 0
            current_power_capacity = 0
            current_spinning_reserve = 0
            
            for src in sources:
                status = src.ops_data[y][m][d][h]['status']
                if status in [-2, -3]:  # Source is not available
                    continue
                
                # Calculate adjusted capacity based on max loading
                power_capacity = src.ops_data[y][m][d][h]['power_capacity']
                max_loading_percentage = src.metadata.get('max_loading', {'value': 1})['value']
                adjusted_capacity = power_capacity * max_loading_percentage
                
                # Include operational sources and simulate output for sources about to fail or reduce output
                if status in [1, -1, 0.5]:
                    if status == -1 or status == 0.5:
                        # Calculate sudden power drop for failing or reducing output sources
                        sudden_power_drop += adjusted_capacity
                        # Assume output remains the same for status 0.5 sources
                        if status == 0.5:
                            adjusted_capacity = src.ops_data[y][m][d][h-1]['power_output']
                    
                    # Calculate contribution proportionally
                    contribution = min(adjusted_capacity, power_req - current_power_output)
                    current_power_output += contribution
                    current_spinning_reserve += power_capacity - contribution  # Update spinning reserve
                    
                    # Record contributions
                    src.ops_data[y][m][d][h]['power_output'] = contribution
                    src.ops_data[y][m][d][h]['energy_output'] = contribution  # Assuming same as power_output
                    src.ops_data[y][m][d][h]['spin_reserve'] = power_capacity - contribution  # Update spinning reserve
            
            # Adjust power requirement based on the total output
            power_req = max(0,power_req - current_power_output)
            
        return power_req, sudden_power_drop



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

                        hourly_results = self.results[y][m][d][h]
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
                        unserved_power_req, sudden_power_drop = self.calc_src_power_and_energy(y,m,d,h,power_req)
                        unserved_power_drop = 0
                        load_shed = 0

                        if unserved_power_req <= 0 and sudden_power_drop > 0:

                            unserved_power_drop,load_shed = self.handle_sudden_power_drop(self, y, m, d, h, sudden_power_drop)

                
                        hourly_results['unserved_power_req'] = unserved_power_req
                        hourly_results['sudden_power_drop'] = sudden_power_drop
                        hourly_results['unserved_power_drop'] = unserved_power_drop
                        hourly_results['load_shed'] = load_shed
                        hourly_results['log'] = self.generate_log(self,y,m,d,h,unserved_power_req, unserved_power_drop,load_shed)


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
    
    def handle_sudden_power_drop(self, y, m, d, h, initial_deficit_power):
        deficit_power = initial_deficit_power
        load_shed = 0

        # Sort sources by block_load_acceptance for effective grouping
        self.src_list.sort(key=lambda src: src.metadata.get('block_load_acceptance', {'value': 0})['value'])
        
        # Group sources by their block_load_acceptance, considering only operational sources
        for block_acceptance, group in groupby(self.src_list, key=lambda src: src.metadata.get('block_load_acceptance', {'value': 0})['value']):
            sources = list(filter(lambda src: src.ops_data[y][m][d][h]['status'] == 1, list(group)))
            if not sources:
                continue  # Skip groups with no operational sources

            src_group_block_acceptance, src_group_spin_reserve = self.distribute_deficit_among_sources(y, m, d, h, sources, deficit_power)

            # Calculate the remaining deficit after distribution
            deficit_power = max(0, deficit_power - src_group_block_acceptance)
            
            if deficit_power <= 0:
                # If the deficit is fully covered, exit the loop
                break

        if deficit_power > 0:
            # If there's still a deficit, consider load shedding
            load_shed = min(self.sheddable_load, deficit_power)
            self.sheddable_load -= load_shed  # Update the sheddable load
            deficit_power -= load_shed  # Adjust the deficit after shedding

        return deficit_power, load_shed

    def distribute_deficit_among_sources(self, y, m, d, h, sources, deficit, block_acceptance):

        src_group_block_acceptance = sum(src.config['rating'] * (block_acceptance / 100) for src in sources)
        src_group_spin_reserve = sum(src.ops_data[y][m][d][h]['spin_reserve'] for src in sources)

        # Calculate how much of the deficit can be covered
        contribution = min(src_group_block_acceptance, deficit, src_group_spin_reserve)

        for src in sources:

            # Calculate each source's contribution based on its block load acceptance
            src_contribution = (src.config['rating'] * (block_acceptance / 100)) / src_group_block_acceptance * contribution
            src.ops_data[y][m][d][h]['power_output'] += src_contribution
            src.ops_data[y][m][d][h]['spin_reserve'] -= src_contribution  # Adjust spin reserve

        return src_group_block_acceptance, src_group_spin_reserve
    
    def generate_log(self, y, m, d, h, unserved_power_req, deficit_power,load_shed):

        # Identifying failed and reduced output sources
        failed_sources = [src for src in self.src_list if src.ops_data[y][m][d][h]['status'] == -1]
        reduced_output_sources = [src for src in self.src_list if src.ops_data[y][m][d][h]['status'] == 0.5]
        
        # Constructing the explanation message
        log_parts = []

        if unserved_power_req > 0:

            full_explanation = f"Total power requirements could not be satisfied. Shortfall = {unserved_power_req} MW"
            return full_explanation
    
        if failed_sources:

            log_parts.append("Failures in sources " + ", ".join([f"{src.config['rating']} {src.config['rating_unit']} {src.name}" for src in failed_sources]))
        
        if reduced_output_sources:

            log_parts.append("sudden reductions in sources " + ", ".join([f"{src.config['rating']} {src.config['rating_unit']} {src.name}" for src in reduced_output_sources]))
        if load_shed > 0:

            log_parts.append(f"{load_shed} MW load was shed")

        # Combine all parts for the final explanation
        full_log = "; ".join(log_parts)
        return full_log






        
