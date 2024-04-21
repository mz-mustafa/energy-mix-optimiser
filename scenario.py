import csv
import datetime
import openpyxl
import pandas as pd
import math
import random
from project import Project
from itertools import groupby

class Scenario:
    def __init__(self, name, client_name, selected_sources, spin_reserve_perc=20, bess_non_emergency_use = 2):
        self.name = name
        self.client_name = client_name
        self.scenario_kpis = {
            'Average Unit Cost (PKR/kWh)': 0,
            'Energy Fulfillment Ratio (%)': 0,
            'Critical Load Interruptions (No.)': 0,
            'Estimated Interruption Loss (M PKR)': 0
        }
        self.timestamp = datetime.datetime.now()
        self.spinning_reserve_perc = spin_reserve_perc
        #0 means no, 1 means yes with equal distribution, 2 means yes with selection utilization
        self.bess_non_emergency_use = bess_non_emergency_use
        self.src_list = selected_sources
        self.src_list.sort(key=lambda src: src.config['priority'])
        self.hourly_results = {
            y: {
                m: {
                    d: {
                        h: {
                            'power_req' : 0,
                            'unserved_power_req': 0,
                            'sudden_power_drop' : 0,
                            'unserved_power_drop' : 0,
                            'load_shed' : 0,
                            'log' : 0
                            } for h in range(24)
                        } for d in range(1, 32)
                    } for m in range(1, 13)
                } for y in range(1, 13)}
        self.yearly_results = []
    
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

    def advance_hour(self,y,m,d,h, src):

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
        return year, month, day, hour

    def previous_hour(self,y,m,d,h):

        hour = h
        day = d
        month = m
        year = y
        hour = h - 1
        if hour < 0:
            hour = 23
            day = day - 1
            if day == 0:
                if month == 3:
                    day = 28
                elif month in [5, 7, 10, 12]:
                    day = 30
                elif month in [1,2,4,6,8,9,11]:
                    day =31
                month = m-1
                if month == 0:
                    month = 12
                    y = y - 1 if y > 1 else 1
        return year,month,day,hour        

    def calc_src_power_and_energy2(self, y, m, d, h, power_req):

        rem_power_req = power_req
        #rem_spin_reserve_req = power_req * self.spinning_reserve_perc/100
        sudden_power_drop = 0
        #group_list = []
        #total_output = 0
        """
        This iteration over groups is to make sure that source groups that need to deliver
        spinning reserve provide it by running at their minimum loading
        """
        for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):

            sources = list(group)
            if not sources or sources[0].metadata['type']['value'] == 'BESS' or sources[0].config['spinning_reserve'] == 0:
                continue
            
            grp_reserve_req_contrib = power_req * self.spinning_reserve_perc * sources[0].config['spinning_reserve']/(100*100)
            min_load_src_count = 0
            grp_output = 0
            grp_reserve = 0
        
            for src in sources:
                src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                
                if src_hourly_ops_data['status'] in [-2, -3] or src_hourly_ops_data['capacity'] ==0:  # Source is not available
                    continue
                #run src at min load and save status. check if req contrib from group to SR is met. If yes, get next group
                min_src_output = src_hourly_ops_data['capacity'] * src.config['min_loading']/100
                src_hourly_ops_data['power_output'] = min_src_output
                grp_reserve += src_hourly_ops_data['capacity'] - min_src_output
                #0.1 status is to temporarily identify which sources were used to meet initial spin reserve.
                src_hourly_ops_data['status'] = 0.1 if src_hourly_ops_data['status'] == 0 else src_hourly_ops_data['status']
                min_load_src_count +=1
                if grp_reserve >= grp_reserve_req_contrib:
                    break
             
            if grp_reserve > 0:
                min_reserve_on_each_src = grp_reserve_req_contrib / min_load_src_count

                for src in sources:
                    src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                    if src_hourly_ops_data['status'] in [0,-2,-3]:
                        continue
                    #if src_hourly_ops_data['status'] == 0.1:
                    #    src_hourly_ops_data['status'] = 1

                    grp_output += src_hourly_ops_data['power_output']
                    src_hourly_ops_data['mandatory_reserve'] = min_reserve_on_each_src
                    src_hourly_ops_data['reserve'] = src_hourly_ops_data['capacity'] - src_hourly_ops_data['power_output']
                    src_hourly_ops_data['energy_output'] = src_hourly_ops_data['power_output']
                    min_load_src_count -=1
                    if min_load_src_count == 0:
                        break
            #rem_power_req -= grp_output
                        
        """
        Sources that need to deliver SR are now configured.
        The min power they deiver is netted off hourly power requirement.
        Now we iterate over source groups to meet remaining power requirement
        """
        for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):
            
            sources = list(group)
            if not sources or sources[0].metadata['type']['value'] == 'BESS':
                continue
            
            grp_potential_output = 0
            grp_output = 0

            for src in sources:
                
                src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                if src_hourly_ops_data['status'] in [-2,-3] or src_hourly_ops_data['capacity'] == 0:
                    continue
                src_can_provide = src_hourly_ops_data['capacity'] - src_hourly_ops_data['power_output'] - \
                    src_hourly_ops_data['mandatory_reserve']
                if src_can_provide <= 0:
                    continue
                #coming to this block means that source can contribute
                grp_potential_output += src_can_provide
                
                if src_hourly_ops_data['status'] != 0.1:
                    src_hourly_ops_data['power_output'] = -1
                    if src_hourly_ops_data['status'] == 0:
                        src_hourly_ops_data['status'] = 1

                if grp_potential_output > rem_power_req:
                    break
            
            if grp_potential_output > 0:

                loading_factor = rem_power_req / grp_potential_output
                loading_factor = 1 if loading_factor > 1 else loading_factor
                group_actual_output = 0
                for src in sources:
            
                    src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                    if src_hourly_ops_data['status'] in [-2,-3] or src_hourly_ops_data['capacity'] == 0:
                        continue
                    
                    #this sources was not used for initial minimum loading
                    if (src_hourly_ops_data['status'] == 1 and src_hourly_ops_data['power_output'] == -1) or \
                        (src_hourly_ops_data['status'] == 0.1 and src_hourly_ops_data['mandatory_reserve'] > 0):

                        src_hourly_ops_data['power_output'] = 0 if src_hourly_ops_data['power_output'] == -1 else src_hourly_ops_data['power_output']
                        src_hourly_ops_data['power_output'] = loading_factor * (src_hourly_ops_data['capacity'] - \
                                    src_hourly_ops_data['power_output'] - \
                                        src_hourly_ops_data['mandatory_reserve'])
                        src_hourly_ops_data['reserve'] = src_hourly_ops_data['capacity'] - src_hourly_ops_data['power_output']
                        src_hourly_ops_data['energy_output'] = src_hourly_ops_data['power_output']
                        src_hourly_ops_data['status'] = 1 #fine for both cases

                    elif src_hourly_ops_data['status'] == -1 and src_hourly_ops_data['power_output'] == -1:
                        
                        src_hourly_ops_data['power_output'] = 0
                        src_hourly_ops_data['power_output'] = loading_factor * (src_hourly_ops_data['capacity'] - \
                                    src_hourly_ops_data['power_output'] - \
                                        src_hourly_ops_data['mandatory_reserve'])
                        sudden_power_drop += src_hourly_ops_data['power_output']
                        src_hourly_ops_data['energy_output'] = 0
                    
                    elif src_hourly_ops_data['status'] == 0.5 and src_hourly_ops_data['power_output'] == -1:

                        src_hourly_ops_data['power_output'] = 0
                        src_hourly_ops_data['power_output'] = loading_factor * (src_hourly_ops_data['capacity'] - \
                                    src_hourly_ops_data['power_output'] - \
                                        src_hourly_ops_data['mandatory_reserve'])


                        year, month, day, hour = self.previous_hour(y,m,d,h)
                        power_output_prev_hour = src.ops_data[year]['months'][month]['days'][day]['hours'][hour]['power_output']
                        sudden_power_drop +=  power_output_prev_hour - src_hourly_ops_data['power_output']
                        src_hourly_ops_data['energy_output'] = src_hourly_ops_data['power_output']

                    group_actual_output += src_hourly_ops_data['power_output']
                rem_power_req = max(0, rem_power_req - group_actual_output)
                if rem_power_req < 0.01:
                    rem_power_req = 0
                    break
            if rem_power_req == 0:
                break
        
        return rem_power_req, sudden_power_drop

    def bess_non_em_contribution(self,y,m,d,h,rem_power_req):

        bess_sources = [src for src in self.src_list if src.metadata['type']['value']== 'BESS']
        if bess_sources:
            #find total capacity, get loading factor then load each source equally.
            total_bess_cap = 0

            for src in bess_sources:

                src_hourly_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                if src_hourly_data['status'] not in [-1, -2, -3]:

                    total_bess_cap += src_hourly_data['reserve']
            
            if self.bess_non_emergency_use == 1:

                loading_factor = rem_power_req / total_bess_cap if total_bess_cap >= rem_power_req else 1
                for src in bess_sources:
                
                    src_hourly_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                    if src_hourly_data['status'] not in [-1, -2, -3]:

                        src_hourly_data['power_output'] = src_hourly_data['reserve'] * loading_factor
                        src_hourly_data['energy_output'] = src_hourly_data['power_output']
                        src_hourly_data['reserve'] -= src_hourly_data['power_output']
                        src_hourly_data['status'] = 1

            elif self.bess_non_emergency_use == 2:

                for src in bess_sources:
                
                    src_hourly_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                    if src_hourly_data['status'] not in [-1, -2, -3]:

                        src_hourly_data['power_output'] = min(rem_power_req, src_hourly_data['reserve'])
                        src_hourly_data['energy_output'] = src_hourly_data['power_output']
                        src_hourly_data['reserve'] -= src_hourly_data['power_output']
                        src_hourly_data['status'] = 1
                        rem_power_req = max(0,rem_power_req - src_hourly_data['power_output'])
                        if rem_power_req < 0.01:
                            rem_power_req = 0
                            break
            if rem_power_req < 0.01:
                rem_power_req = 0
        return rem_power_req
    
    def simulate(self):

        for y in range(1,13):

            print(f'Simulating Year {y}')
            for m in range (1,13):

                if m == 2:  # February
                    days = 28
                elif m in [4, 6, 9, 11]:  # April, June, September, November
                    days = 30
                else:  # All other months
                    days = 31

                for d in range (1, days+1):

                    for h in range (0,24):

                        hourly_results = self.hourly_results[y][m][d][h]
                        #set the power requirement
                        power_req = Project.load_data[y][m][d][h]
                        hourly_results['power_req'] = power_req

                        #charging_pwr_req = self.set_bess_parameters(y,m,d,h, starting = True)
                        #power_req += charging_pwr_req
                        #Consumption of sources, update key results in the scenario
                        unserved_power_req, sudden_power_drop = self.calc_src_power_and_energy2(y,m,d,h,power_req)
                        #Use bess only if needed
                        if unserved_power_req > 0:

                           unserved_power_req = self.utilize_reserves(y,m,d,h,unserved_power_req)

                        if unserved_power_req > 0 and self.bess_non_emergency_use in [1,2]:
                            unserved_power_req = self.bess_non_em_contribution(y,m,d,h,unserved_power_req)
                        
                        unserved_power_drop = 0
                        load_shed = 0

                        if unserved_power_req <= 0 and sudden_power_drop > 0:

                            unserved_power_drop,load_shed = self.handle_sudden_power_drop(y, m, d, h, sudden_power_drop)

                        _ = self.set_bess_parameters(y,m,d,h, starting = False)

                        hourly_results['unserved_power_req'] = unserved_power_req
                        hourly_results['sudden_power_drop'] = sudden_power_drop
                        hourly_results['unserved_power_drop'] = unserved_power_drop
                        hourly_results['load_shed'] = load_shed
                        hourly_results['log'] = self.generate_log(y,m,d,h,unserved_power_req, unserved_power_drop,load_shed)

                        # Sort sources by priority for processing
                        self.src_list.sort(key=lambda src: src.config['priority'])
        self.aggregate_data_for_reporting()            

    def charge_bess(self,y,m,d,h):

        #for each source group in order of priority.
        #finds its reserve capacity.
        #use entire reserve cap to charge bess.
        #take BESS max. charge rate as a sceanrio parameter i.e. 1 or more hours.
        #power, energy out put of each src in source group to set accordingly.
        #set status and reserve of battery accordingly. (based on 50%)
        #while power of charging sources will be applied forr the whole eyar.

        #after the above we
        pass


    def utilize_reserves(self, y, m, d, h, remaining_demand):

        for src in self.src_list:

            if src.metadata['type']['value'] == 'BESS':
                continue
            src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
            if src_hourly_ops_data['status'] in [-1, -2,-3] or \
                src_hourly_ops_data['capacity'] == 0 or \
                    src_hourly_ops_data['reserve'] == 0:
                continue
            contribution = min(remaining_demand, src_hourly_ops_data['reserve'])
            remaining_demand -= contribution
            src_hourly_ops_data['power_output'] += contribution
            src_hourly_ops_data['energy_output'] += contribution
            src_hourly_ops_data['reserve'] -= contribution
            if remaining_demand < 0.01:
                remaining_demand = 0
                break

        return remaining_demand


    def set_bess_parameters(self, y, m, d, h, starting):

        #assumption that sim starts with full reserve
        #and a status of 0, so charge req is 0
        if y==1 and m == 1 and d == 1 and h == 0 and starting:
            return 0
        
        #extract BESS sources
        bess_sources = [src for src in self.src_list if src.metadata['type']['value'] == 'BESS']
        #iterate over sources
        bess_charging_energy = 0
        for src in bess_sources:

            src_hourly_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
            #i -1 and -2, -3, then capacity and reserve =0
        
            if src_hourly_data['status'] in [-1, -2, -3]:

                src_hourly_data['capacity'] = src_hourly_data['reserve'] = 0
            
            else:

                if starting:
                    #check reserve for previous hour.
                    year, month, day, hour = self.previous_hour(y,m,d,h)
                    src_prev_hour_data = src.ops_data[year]['months'][month]['days'][day]['hours'][hour]
                    #if bess has been trickle charging undisturbed
                    if src_prev_hour_data['status'] == 0:
                        
                        #keep trickle charging
                        src_hourly_data['status'] = 0
                        src_hourly_data['reserve'] = src_prev_hour_data['reserve']
                        bess_charging_energy += src_hourly_data['capacity'] * 0.01

                    #remaining status is 1/ discharging and 2 full charging
                    else:
                        #probably not needed because for discharging this will always be true
                        if src_prev_hour_data['reserve'] < src_prev_hour_data['capacity']:
                            
                            #full charge
                            src_hourly_data['status'] = 2
                            max_charge_energy_av = src_prev_hour_data['capacity']/4
                            charging_req = src_prev_hour_data['capacity'] - src_prev_hour_data['reserve']
                            #this will be charging that can happen during the hour.
                            actual_charging = min(charging_req, max_charge_energy_av)
                            #we are assuming that the bess will be charged throughout, which might not be the case.
                            bess_charging_energy += actual_charging
                            
                            #for now, we add 50% of actual charging to the reserve.
                            #to simulate that the full cap may not be achieved.
                            src_hourly_data['reserve'] = min(src_hourly_data['capacity'], src_prev_hour_data['reserve'] + (max_charge_energy_av/2))
                            
                            if src_hourly_data['reserve'] == src_hourly_data['capacity']:
                                src_hourly_data['status'] == 0

                        else:
                            src_hourly_data['status'] = 0
                            src_hourly_data['reserve'] = src_prev_hour_data['reserve']
                #if the function is ending
                else:
                    #this means that the bess is still in full charge
                    if src_hourly_data['status'] == 2:
                        #and that the reserve can be topped up.
                        year, month, day, hour = self.previous_hour(y,m,d,h)
                        src_prev_hour_data = src.ops_data[year]['months'][month]['days'][day]['hours'][hour]

                        src_hourly_data['reserve'] = min(src_hourly_data['capacity'], 
                                                         src_hourly_data['reserve'] + (
                                                             0.5*src_prev_hour_data['capacity']/4))
                        if src_hourly_data['reserve'] == src_hourly_data['capacity']:
                            src_hourly_data['status'] == 0
                    
                    #there is no need to test for 0 (trickle) and 1 (discharging) conditions
                    #in case of 0, there is no change to what is set in the starting block
                    #in case of 1, other simulate functions changed the state
                    #so it means that reserve has already been reduced and status changed.
        return bess_charging_energy

    
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
            #the block acceptance == 0 condition will filter out all solar sources, which is correct.
            if not sources or sources[0].metadata['block_load_acceptance']['value'] == 0: 
                continue
            #we do this, because BESS is the only instaneous source that can handle sudden deficits
            elif sources[0].metadata['type']['value'] == 'BESS': 
                
                for src in sources:
                    if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] not in [-1, -2, -3]:
                        src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] = 1
            
            sources = list(filter(lambda src: src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == 1, sources))

            if not sources:
                continue  # Skip groups with no operational sources

            # Distribute deficit among sources and get updated acceptance and reserves
            if block_acceptance <= 0:
                src_group_block_acceptance = 0
            else:
                deficit_power = self.distribute_deficit_among_sources(y, m, d, h, sources, deficit_power, block_acceptance)

            # Check if deficit is fully managed
            if deficit_power <= 0:
                break

        # Adjust sources with status -1 and 0.5, setting their output and reserve to 0
        for src in self.src_list:
            if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == -1: 
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = 0
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = 0
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['reserve'] = 0
            

        # Handle remaining deficit with load shedding
        if deficit_power > 0:

            load_shed = min(sheddable_load, deficit_power)
            deficit_power -= load_shed

        return deficit_power, load_shed

    def distribute_deficit_among_sources(self, y, m, d, h, sources, deficit, block_acceptance):

        src_group_block_acceptance = sum(src.config['rating'] * (block_acceptance / 100) for src in sources)

        src_group_reserve = sum(src.ops_data[y]['months'][m]['days'][d]['hours'][h]['reserve'] for src in sources)

        # Calculate how much of the deficit can be covered
        contribution = min(src_group_block_acceptance, deficit, src_group_reserve)
        if contribution == 0:
            return deficit
        
        for src in sources:
            src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
            # Calculate each source's contribution based on its reserve
            src_contribution = (src_hourly_ops_data['reserve']/ src_group_reserve) * contribution
            src_hourly_ops_data['power_output'] += src_contribution
            src_hourly_ops_data['energy_output'] += src_contribution
            src_hourly_ops_data['reserve'] -= src_contribution
            
            if src.metadata['type']['value'] == 'BESS':

                bess_src_consumption = src_hourly_ops_data['capacity'] - src_hourly_ops_data['reserve']
                if bess_src_consumption <= 0.25 * src_hourly_ops_data['capacity']:

                    #we assume that BESS will charged back to full capacity within the hour.
                    src_hourly_ops_data['status'] = 0
                    src_hourly_ops_data['reserve'] = src_hourly_ops_data['capacity']

            deficit = max(0,deficit - src_contribution)

            if deficit <= 0.01:
                deficit = 0
                break
        return deficit
    
    def generate_log(self, y, m, d, h, unserved_power_req, deficit_power,load_shed):

        # Identifying failed and reduced output sources
        failed_sources = [src for src in self.src_list if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == -1]
        reduced_output_sources = [src for src in self.src_list if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] == 0.5]
        
        # Constructing the explanation message
        log_parts = []

        if unserved_power_req > 0:

            full_explanation = f"Total power requirements could not be satisfied. Shortfall = {round(unserved_power_req,3)} MW"
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


    def aggregate_data_for_reporting(self):

        for src in self.src_list:

            src.aggregate_day_stats()
            src.aggregate_month_stats()
            src.aggregate_year_stats()
        self.aggregate_yearly_data_for_csv()
        self.calculate_scenario_kpis()


    def calculate_scenario_kpis(self):
        # Ensure that yearly_data has been populated
        if not self.yearly_results:
            print("Yearly data is not available. Please aggregate yearly data first.")
            return

        avg_unit_cost = sum(year_record['Unit Cost (PKR/kWh)'] for year_record in self.yearly_results) / len(self.yearly_results)
        avg_enr_fulfill = sum(year_record['Energy Fulfilment Ratio (%)'] for year_record in self.yearly_results) / len(self.yearly_results)
        critical_load_interr = sum(year_record['Critical Load Interruptions'] for year_record in self.yearly_results)
        interr_loss = sum(year_record['Estimated Loss due to Interruptions'] for year_record in self.yearly_results)
        load_shed_events = sum(year_record['Non-critical Load shedding events'] for year_record in self.yearly_results)
        # Update the scenario_kpis dictionary with the calculated values
        self.scenario_kpis = {
            'Average Unit Cost (PKR/kWh)': avg_unit_cost,
            'Energy Fulfillment Ratio (%)': avg_enr_fulfill,
            'Critical Load Interruptions (No.)': critical_load_interr,
            'Estimated Interruption Loss (M PKR)': interr_loss,
            'Non-critical Load shedding events': load_shed_events,
        }


    def aggregate_yearly_data_for_csv(self):
        
        self.yearly_results.clear()
        for y in range(1, 13):
            total_energy_req = 0
            unserved_instances = 0
            critical_load_interruptions = 0
            load_shed_events = 0

            # Summing total energy requirements
            for m in range(1, 13):
                if m == 2:  # February
                    days = 28
                elif m in [4, 6, 9, 11]:  # April, June, September, November
                    days = 30
                else:  # All other months
                    days = 31

                for d in range(1, days+1):  # Assuming 31 days for simplicity; adjust as needed
                    for h in range(24):
                        hour_data = self.hourly_results[y][m][d][h]
                        total_energy_req += hour_data['power_req']
                        if hour_data['unserved_power_req'] > 0.01:
                            unserved_instances += 1
                            critical_load_interruptions += 1
                        if hour_data['unserved_power_drop'] > 0.01:
                            critical_load_interruptions += 1
                        if hour_data['load_shed'] > 0:
                            load_shed_events +=1
            # Calculate Energy Fulfilment Ratio (%)
            total_rows = 365 * 24  # Simplified; adjust for actual days in each month/year
            energy_fulfilment_ratio = 100 * (1 - (unserved_instances / total_rows))

            # Calculate Estimated Loss due to Interruptions
            estimated_loss_due_to_interruptions = (critical_load_interruptions * Project.site_data['loss_during_failure'])/1000000

            # Initialize variable for total cost of operation across all sources
            total_cost_of_operation = 0

            # Aggregate total cost of operation from each source for the year
            for src in self.src_list:
                source_year_data = src.ops_data.get(y, {})
                total_cost_of_operation += source_year_data.get('year_cost_of_operation', 0)

            total_cost_m_pkr = estimated_loss_due_to_interruptions + total_cost_of_operation 

            # Calculate Unit Cost (PKR/kWh), ensuring no division by zero
            if total_energy_req > 0:
                unit_cost_pkr_per_kwh = (total_cost_m_pkr * 1000) / total_energy_req
            else:
                unit_cost_pkr_per_kwh = 0  # Avoid division by zero

            year_record = {
            'year': y,
            'total_energy_requirement (MWh)': round(total_energy_req,2),
            'Energy Fulfilment Ratio (%)': round(energy_fulfilment_ratio,2),
            'Critical Load Interruptions': round(critical_load_interruptions,2),
            'Estimated Loss due to Interruptions': round(estimated_loss_due_to_interruptions,2),
            'Non-critical Load shedding events': load_shed_events,
            'Total Cost (M PKR)': round(total_cost_m_pkr,2),
            'Unit Cost (PKR/kWh)': round(unit_cost_pkr_per_kwh,2)
            }

            # Aggregate data for each source
            source_data = []
            for index, src in enumerate(self.src_list, start=1):
                source_year_data = src.ops_data.get(y, {})
                source_energy_output = source_year_data.get('year_energy_output', 0)
                source_op_proportion = source_year_data.get('year_operation_hours', 0) / (365 * 24)
                source_total_cost = source_year_data.get('year_cost_of_operation', 0)
                source_unit_cost = source_year_data.get('year_unit_cost', 0)
                source_name = f"SRC-{index} {src.metadata['generic_name']['value']}"
                
                source_data.append({
                    f'{source_name} energy output (MWh)': round(source_energy_output,2),
                    f'{source_name} year operating proportion (%)': round(source_op_proportion * 100,2),
                    f'{source_name} total cost of operation (M PKR)': round(source_total_cost,2),
                    f'{source_name} unit cost (PKR/kWh)': round(source_unit_cost,2),
                })
            
            year_record.update({k: v for source_dict in source_data for k, v in source_dict.items()})
            self.yearly_results.append(year_record)
        
    def write_yearly_data_to_csv(self, filepath):

        if self.yearly_results:
            keys = self.yearly_results[0].keys()
            with open(filepath, 'w', newline='') as output_file:
                dict_writer = csv.DictWriter(output_file, keys)
                dict_writer.writeheader()
                dict_writer.writerows(self.yearly_results)

    def write_hourly_data_to_csv(self):
     
        # Prepare data for DataFrame
        data = []
        # Assuming the first source has all the necessary time periods defined
        first_src_ops_data = self.src_list[0].ops_data

        for y in first_src_ops_data:
            for m in first_src_ops_data[y]['months']:
                for d in first_src_ops_data[y]['months'][m]['days']:
                    for h in first_src_ops_data[y]['months'][m]['days'][d]['hours']:
                        row = [y, m, d, h]
                        for src in self.src_list:
                            ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                            # Append ops_data values for the source
                            row.extend([
                                round(ops_data.get('capacity', 0),2),
                                round(ops_data.get('power_output', 0),2),
                                round(ops_data.get('energy_output', 0),2),
                                round(ops_data.get('reserve', 0),2),
                                ops_data.get('status', '')
                            ])
                        # Append results data for the same time period
                        result_data = self.hourly_results[y][m][d][h]
                        row.extend([
                            round(result_data.get('power_req', 0),2),
                            round(result_data.get('unserved_power_req', 0),2),
                            round(result_data.get('sudden_power_drop', 0),2),
                            round(result_data.get('unserved_power_drop', 0),2),
                            round(result_data.get('load_shed', 0),3),
                            result_data.get('log', 0)
                        ])
                        data.append(row)

        # Define column names, including results column names
        column_names = ['Year', 'Month', 'Day', 'Hour']
        for i in range(1, len(self.src_list) + 1):
            column_names.extend([f'Src_{i}_power_capacity', f'Src_{i}_power_output', f'Src_{i}_energy_output', f'Src_{i}_spin_reserve', f'Src_{i}_status'])
        # Extend column names with results column names
        column_names.extend(['Power_Req','Unserved_Power_Req', 'Sudden_Power_Drop', 'Unserved_Power_Drop', 'Load_Shed', 'Log'])

        # Create DataFrame
        df = pd.DataFrame(data, columns=column_names)

        # Write to CSV
        df.to_csv('data/hourly_data.csv', index=False)


        
