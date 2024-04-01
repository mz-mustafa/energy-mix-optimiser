import datetime
import openpyxl
import pandas as pd
import math
import random
from project import Project
from itertools import groupby

class Scenario:
    def __init__(self, name, client_name, selected_sources):
        self.name = name
        self.client_name = client_name
        self.timestamp = datetime.datetime.now()
        self.src_list = selected_sources
        self.src_list.sort(key=lambda src: src.config['priority'])
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
        #sorted_sources = sorted(self.src_list, key=lambda x: x.config['priority'])
        self.src_list.sort(key=lambda src: src.config['priority'])
        # Group sources by priority
        for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):
            sources = list(group)
            if not sources or sources[0].metadata['type']['value'] == 'BESS':
                continue  # Skip empty groups
            
            # Use spinning reserve requirements from the first source in the group
            spin_reserve_req = sources[0].config['spinning_reserve']
            current_power_output = 0
            current_group_capacity = 0
            current_spinning_reserve = 0
            
            for src in sources:
                status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                if status in [-2, -3]:  # Source is not available
                    continue
                
                # Calculate adjusted capacity based on max loading for current source in group
                src_capacity = src.adjusted_capacity(y,m,d,h)
                if src_capacity == 0:
                    continue
                current_group_capacity += src_capacity
                if status == 0:
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] = 1
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = -0.001
                if current_group_capacity >= power_req:

                    if  current_group_capacity - power_req >= spin_reserve_req * current_group_capacity/100:

                        break
            if current_group_capacity > 0:
                loading_factor = power_req / current_group_capacity if current_group_capacity > 0 else 0
                if loading_factor > 1:
                    loading_factor = 1
                grp_actual_output = 0
                for src in sources:

                    status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                    
                    if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] == -0.001:

                        src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = 0
                        src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = loading_factor * src.adjusted_capacity(y,m,d,h)
                        
                        if status == 1:
        
                            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = src.adjusted_capacity(y,m,d,h) - src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                        
                        #if utilized and failed
                        if status == -1:
                                    
                            sudden_power_drop += src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = 0

                        if status == 0.5:

                            sudden_power_drop += src.ops_data[y]['months'][m]['days'][d]['hours'][h-1]['power_output'] - src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] 
                            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = src.ops_data[y]['months'][m]['days'][d]['hours'][h-1]['power_output'] 
                            

                        grp_actual_output += src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                power_req = max(0,power_req - grp_actual_output)
                if power_req < 0.001: #1kW (math error margin)
                    power_req = 0
                    break
            if power_req == 0:
                break
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
                                status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                                if status == 0:
                                    #Trickel charge
                                    power_req += src.config['rating'] * 0.02
                                elif status == 2:
                                    #Full charge
                                    power_req += src.config['rating'] * 0.5
                        
                        #Consumption of sources, update key results in the scenario
                        unserved_power_req, sudden_power_drop = self.calc_src_power_and_energy(y,m,d,h,power_req)
                        unserved_power_drop = 0
                        load_shed = 0

                        if unserved_power_req <= 0 and sudden_power_drop > 0:

                            unserved_power_drop,load_shed = self.handle_sudden_power_drop(y, m, d, h, sudden_power_drop)

                        hourly_results['unserved_power_req'] = unserved_power_req
                        hourly_results['sudden_power_drop'] = sudden_power_drop
                        hourly_results['unserved_power_drop'] = unserved_power_drop
                        hourly_results['load_shed'] = load_shed
                        hourly_results['log'] = self.generate_log(y,m,d,h,unserved_power_req, unserved_power_drop,load_shed)

                    #TODO min-off take in stat aggregation.
                    #if min offtake then a differnt charge applied (min-offtake x tariff, otherwise actual consumption x tariff)
                #self.calculate_opex(y,m)

                        
    
    def calculate_opex(self,y,m):

        for src in self.src_list:

            if src.ops_data[y]['source_present'] == 0:
                continue
                
            #we may not need this condition at all.
            #if src.metadata['finance']['value'] == 'CAPTIVE' and src.metadata['type']['value'] == 'NR':
            
                
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
        non_critical_load_projection = Project.load_projection[1]['total_load'] - Project.load_projection[1]['critical_load']         
        running_load_factor = Project.load_data[y][m][d][h] / Project.load_projection[1]['total_load']
        if running_load_factor > 1:
            running_load_factor = 1
        
        sheddable_load = non_critical_load_projection * running_load_factor

        # Sort sources by block_load_acceptance for effective grouping
        self.src_list.sort(key=lambda src: src.metadata.get('block_load_acceptance', {'value': 0})['value'], reverse = True)
        
        # Group sources by their block_load_acceptance, considering only operational sources
        for block_acceptance, group in groupby(self.src_list, key=lambda src: src.metadata.get('block_load_acceptance', {'value': 0})['value']):
            
            sources = list(group)

            if not sources: 
                continue
            elif sources[0].metadata['type']['value'] == 'BESS': 
                
                for src in sources:
                    if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == 0:
                        src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] = 1
            
            sources = list(filter(lambda src: src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == 1, sources))

            if not sources:
                continue  # Skip groups with no operational sources

            # Distribute deficit among sources and get updated acceptance and reserves
            if block_acceptance <= 0:
                src_group_block_acceptance = 0
            else:
                src_group_block_acceptance, src_group_spin_reserve = self.distribute_deficit_among_sources(y, m, d, h, sources, deficit_power, block_acceptance)

            # Update the remaining deficit after distribution
            deficit_power -= src_group_block_acceptance

            # Check if deficit is fully managed
            if deficit_power <= 0:
                break

        # Adjust sources with status -1 and 0.5, setting their output and reserve to 0
        for src in self.src_list:
            if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == -1: 
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = 0
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = 0
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = 0
            
            elif src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == 0.5:

                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] 

        # Handle remaining deficit with load shedding
        if deficit_power > 0:

            load_shed = min(sheddable_load, deficit_power)
            deficit_power -= load_shed

        return deficit_power, load_shed

    def distribute_deficit_among_sources(self, y, m, d, h, sources, deficit, block_acceptance):

        src_group_block_acceptance = sum(src.config['rating'] * (block_acceptance / 100) for src in sources)
        src_group_spin_reserve = sum(src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] for src in sources)

        # Calculate how much of the deficit can be covered
        if sources[0].metadata['type']['value'] == 'BESS':
        
            contribution = min(src_group_block_acceptance, deficit)
        else:
            contribution = min(src_group_block_acceptance, deficit, src_group_spin_reserve)

        for src in sources:

            # Calculate each source's contribution based on its block load acceptance
            src_contribution = (src.config['rating'] * (block_acceptance / 100)) / src_group_block_acceptance * contribution
            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] += src_contribution
            
            if src.metadata['type']['value'] == 'BESS':
                
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = 0
                hour = h
                day = d
                month = m
                year = y
                hour = h + 1
                if hour > 23:
                    hour = 0
                    day = d + 1
                    if day > len(src.ops_data[y]['months'][m]['days']):
                        day = 1
                        month = m + 1
                        if month > 12:
                            month = 1
                            year = y+1 if y < 12 else 12
                src.ops_data[year]['months'][month]['days'][day]['hours'][hour]['status'] = 2
                hour = h + 2
                if hour > 23:
                    hour = 0
                    day = d + 1
                    if day > len(src.ops_data[y]['months'][m]['days']):
                        day = 1
                        month = m + 1
                        if month > 12:
                            month = 1
                            year = y+1 if y < 12 else 12
                src.ops_data[year]['months'][month]['days'][day]['hours'][hour]['status'] = 2
            else:
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] -= src_contribution

        return src_group_block_acceptance, src_group_spin_reserve
    
    def generate_log(self, y, m, d, h, unserved_power_req, deficit_power,load_shed):

        # Identifying failed and reduced output sources
        failed_sources = [src for src in self.src_list if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == -1]
        reduced_output_sources = [src for src in self.src_list if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == 0.5]
        
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
        if full_log == '':
            full_log = "Normal Operation"
        return full_log






        
