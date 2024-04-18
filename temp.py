import datetime
import math
import random

import pandas as pd
from source_meta import GasGenMeta
import openpyxl

class Source:
    def __init__(self, n, source_type, src_priority):
        self.n = n
        self.inputs = self.create_input_structure()
        self.outputs = self.create_output_structure()
        self.source_type = source_type
        self.priority = src_priority

    def create_input_structure(self):
        return {
            year: {
                'count_prim_units': 0,
                'rating_prim_units': 0
            }
            for year in range(0, self.n + 1)
        }

    def create_output_structure(self):
        output = {}
        for year in range(0, self.n + 1):
            output[year] = {
                'capital_cost': 0,
                'depreciation_cost': 0,
            }
            for month in range(1, 13):
                output[year][month] = {
                    'energy_output_prim_units': 0,
                    'fixed_opex': 0,
                    'num_pot_failures': 0,
                    'num_failures': 0,
                    'failure_duration': 0,
                    'co2_emissions': 0
                }
                for hour in range(1, 25):
                    output[year][month][hour] = {
                        'power_output_prim_units': 0,
                        'loading_prim_units': 0
                    }
        return output

class GasGenSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n, 'Gas Generator', src_p)
        self.meta = GasGenMeta()
        self.extend_input_structure()
        self.extend_output_structure()

    def extend_input_structure(self):
        # Extend input structure with GasGenSource specific keys
        for year in range(self.n + 1):  # Use range based on n to iterate over years
            self.inputs[year]['rating_backup_units'] = 0
            self.inputs[year]['count_backup_units'] = 0
            self.inputs[year]['perc_rated_output'] = 0

        # These two keys are not associated with a specific year, so they remain the same
        self.inputs['chp_operation'] = False
        self.inputs['fuel_type'] = 'NG'

    def extend_output_structure(self):
        # Extend output structure with GasGenSource specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):  # Use range for months
                self.outputs[year][month]['energy_output_backup_units'] = 0
                self.outputs[year][month]['energy_free_cooling'] = 0
                self.outputs[year][month]['var_opex'] = 0
                self.outputs[year][month]['fuel_charges'] = 0
                for hour in range(1, 25):  # Use range for hours
                    self.outputs[year][month][hour]['power_output_backup_units'] = 0
                    self.outputs[year][month][hour]['loading_backup_units'] = 0

class Scenario:
    def __init__(self, name, client_name, input_file_path='input_data.xlsx', n=5):
        self.name = name
        self.client_name = client_name
        self.timestamp = datetime.datetime.now()
        self.scenario_spec = {}
        self.sources_dict = {}
        self.sources_list = []
        self.ip_site_data = {}
        self.ip_load_data = {}
        self.ip_enr_data = {}
        #OUTPUT DATAFRAMES
        self.power_df = []
        self.energy_df = []
        self.capex_df = []
        self.opex_df = []
        self.emissions_df = []

        def add_source(self, source):

            self.sources_dict[source.source_type] = source
            self.sources_list.append(source)
            self.sources_list.sort(key= lambda src: src.priority)

        def determine_pot_failures(self, src, year, month):

            # for Grid
            # monthly failures are independent of one another.
            if src.source_type == 'Grid' or 'PPA':
                monthly_failures = src.meta.num_failures_year / 12
                lower_bound = 0.75 * monthly_failures
                upper_bound = 1.25 * monthly_failures
                # Randomly round up or down for each bound
                lower_bound = math.ceil(lower_bound) if random.choice([True, False]) else math.floor(lower_bound)
                upper_bound = math.ceil(upper_bound) if random.choice([True, False]) else math.floor(upper_bound)

                # If lower_bound and upper_bound are equal, return one of them
                if lower_bound > upper_bound:
                    return random.randint(upper_bound, lower_bound)
                elif lower_bound < upper_bound:
                    return random.randint(lower_bound, upper_bound)
                else:
                    return lower_bound
            else:
                # for all other sources
                num_failures_so_far = sum(src.outputs[year][m]['num_failures'] for m in range(1, month))
                poss_annual_failures = src.meta.num_failures_year
                remaining_failures = poss_annual_failures - num_failures_so_far
                months_left = 12 - month + 1

                # No more failures needed
                if remaining_failures <= 0:
                    return 0

                # Expected failures this month
                expected_failures = remaining_failures / months_left

                # Randomly decide the number of failures this month
                monthly_failures = 0
                for _ in range(int(expected_failures * 2)):  # Adjust the multiplier for more randomness
                    if random.random() < expected_failures / 2:  # Adjust the divisor for probability
                        monthly_failures += 1

            # Ensure the total failures don't exceed the annual limit

            return min(monthly_failures, remaining_failures)

        def get_gen_pwr_ops(self, source_type, unit_type, current_year):
            # Check if the given source_type exists in the sources dictionary
            if source_type not in self.sources_dict:
                raise ValueError(f"No source of type {source_type} found.")

            source = self.sources_dict[source_type]

            # Check if the unit_type is valid
            if unit_type not in ['PRIMARY', 'BACKUP']:
                raise ValueError(f"Invalid unit type {unit_type}.")

            # Determine which attributes to use based on the unit_type
            if unit_type == 'PRIMARY':
                count_key = 'count_prim_units'
                rating_key = 'rating_prim_units'
            else:  # 'BACKUP'
                count_key = 'count_backup_units'
                rating_key = 'rating_backup_units'
            perc_op_key = 'perc_rated_output'

            # If the source has the 'gas_fuel_type' attribute, then calculate capacity with derating
            if 'gas_fuel_type' in source.inputs:
                fuel_der_fac = self.derating_factor(source.inputs['gas_fuel_type'])
            else:
                fuel_der_fac = 1

            # Calculate total potential power with degradation
            total_pwr_pot = 0
            degradation_rate = source.meta.degradation if hasattr(source.meta, 'degradation') else 0
            total_count = 0
            for year, yr_data in source.inputs.items():
                if isinstance(year, int) and year <= current_year:
                    years_of_operation = current_year - year
                    if perc_op_key in yr_data:
                        perc_op = yr_data[perc_op_key] / 100
                    else:
                        perc_op = 1
                    degradation_factor = 1 - (degradation_rate * years_of_operation / 100)
                    total_pwr_pot += yr_data[count_key] * yr_data[rating_key] * perc_op * \
                                     fuel_der_fac * degradation_factor
                    total_count += yr_data[count_key]

                    if year == current_year:
                        break

            return total_count, total_pwr_pot

        def get_gen_ener_op(self, source_type, current_year, current_month):

            if source_type not in self.sources_dict:
                raise ValueError(f"No source of type {source_type} found.")

            source = self.sources_dict[source_type]

            # NOT NEEDED IN ENERGY BECAUSE WE ALREADY ACCOUNT FOR THIS IN
            """
            if 'gas_fuel_type' in source.inputs:
                fuel_der_fac = self.derating_factor(source.inputs['gas_fuel_type'])
            else:
                fuel_der_fac = 1
            """

            # Retrieve the degradation rate
            degradation_rate = getattr(source.meta, 'degradation', 0)

            # Total Energy Potential considering straight-line degradation
            total_ener_pot = 0
            for year, yr_data in source.inputs.items():
                if isinstance(year, int) and year <= current_year:
                    years_of_operation = current_year - year
                    if 'perc_rated_output' in yr_data:
                        perc_op = yr_data['perc_rated_output'] / 100
                    else:
                        perc_op = 1

                    degradation_factor = 1 - (degradation_rate * years_of_operation / 100)
                    total_ener_pot += (yr_data['count_prim_units'] * yr_data['rating_prim_units'] *
                                       perc_op * degradation_factor) * 24

                    if year == current_year:
                        break

            # Multiply by the number of days in the current month to get the total energy potential for the month
            total_ener_pot *= self.ip_enr_data[current_month]['days']

            return total_ener_pot

        # CALCULATION FUNCTIONS
        def energy_calculation(self):

            for year in range(1, self.n + 1):
                for month in range(1, 13):
                    month_data = {'year': year, 'month': month}
                    print(f"Energy Calc Year {year}, month {month}")
                    # Determine energy requirements
                    prod_enr_req, cool_enr_req, bess_charge_enr_req = self._get_monthly_energy_req(year, month)

                    month_data['Prod Energy Req, MWh'] = prod_enr_req
                    month_data['Cooling Energy Req, MWh'] = cool_enr_req

                    cool_enr_req = max(0, cool_enr_req - self.free_cooling_enr_cal(year, month))
                    month_data['Cooling Energy Req after CHP adj., MWh'] = cool_enr_req
                    month_data['BESS Charging Energy Req, MWh'] = bess_charge_enr_req

                    month_tot_enr_req = prod_enr_req + cool_enr_req + bess_charge_enr_req
                    month_data['Total Energy Req, MWh'] = month_tot_enr_req
                    month_rem_enr_req = month_tot_enr_req

                    critical_load = self.ip_load_data[year]['crit_load_prop'] * self.ip_load_data[year][
                        'max_dem_load_day'] / 100

                    # Energy from renewables
                    for ren_src_name in ['Wind', 'Solar']:

                        if ren_src_name in self.sources_dict:
                            print(f"Finding {ren_src_name} energy")
                            pot_enr_op = self.sources_dict[ren_src_name].calc_output_energy(year, month)

                            # Wind energy func returns daily energy value
                            if ren_src_name == 'Wind':
                                pot_enr_op *= self.ip_enr_data[month]['days']
                            ren_enr_op = min(month_rem_enr_req, pot_enr_op)
                            month_rem_enr_req -= ren_enr_op
                            self.sources_dict[ren_src_name].outputs[year][month][
                                'energy_output_prim_units'] = ren_enr_op
                            month_data[f"{ren_src_name} Output in MWh"] = ren_enr_op

                    month_data["Remaining Energy Demand (after Renewables) MWh"] = month_rem_enr_req

                    for src in self.sources_list:

                        if src.source_type in self.stable_sources(include_backup=False):
                            src_name = src.source_type

                            print(f"Finding {src_name} energy")
                            # Calculate Monthly Failure Probability
                            num_pot_failures = self.determine_pot_failures(src, year, month)
                            month_data[f'{src_name} Potential Failures'] = num_pot_failures
                            if num_pot_failures == 0:
                                month_data[f'{src_name} Failures mitigated'] = 0
                                month_data[f'{src_name} Unavailability, hrs'] = 0

                            else:
                                print(f"Finding failures for {src_name}")
                                # find energy required to cover each failure
                                en_per_fail = src.meta.avg_failure_time * critical_load


                                num_fails_not_cov = num_pot_failures
                                num_failures = num_pot_failures
                                failure_duration = 0
                                # look through other primary sources
                                for alt_src in self.sources_list:
                                    if alt_src.source_type in self.stable_sources(include_backup=True) \
                                            and alt_src.source_type != src.source_type:

                                        alt_src_name = alt_src.source_type
                                        print(f"Checking if {alt_src_name} can provide failure coverage {src_name}")
                                        _, total_cap = self.get_gen_pwr_ops(alt_src_name, 'PRIMARY', year)
                                        # ...and check if these can kick in to cover the failure duration
                                        if total_cap >= critical_load:

                                            print(f"{alt_src_name} does have power cap to backup {src_name}")
                                            # how many failures can be alternate source cover in terms of energy
                                            alt_src_en_pot = self.get_gen_ener_op(alt_src_name, year, month)
                                            alt_src_en_rem = alt_src_en_pot - \
                                                             alt_src.outputs[year][month]['energy_output_prim_units']

                                            alt_src_nfail_cover = math.floor(alt_src_en_rem / en_per_fail)
                                            if not alt_src_nfail_cover:
                                                alt_src_nfail_cover = 0
                                            alt_src_nfail_cover = min(alt_src_nfail_cover, num_fails_not_cov)

                                            # add the failure coverage energy to the alt source's expenditure
                                            # remaining failures are reduced and other sources may cover them (loop)
                                            if alt_src.source_type == 'Grid':
                                                backup_enr_pk, backup_enr_nonpk = \
                                                    self.grid_pk_to_offpk(month, alt_src_nfail_cover * en_per_fail)
                                                alt_src.outputs[year][month]['energy_output_peak'] = backup_enr_pk
                                                alt_src.outputs[year][month]['energy_output_offpeak'] = backup_enr_nonpk

                                            elif alt_src.source_type == 'HFO+Gas Generator':
                                                gas_enr, hfo_enr = alt_src.gas_hfo_enr_op(
                                                    alt_src_nfail_cover * en_per_fail)
                                                alt_src.outputs[year][month]['energy_output_prim_units'] = gas_enr
                                                alt_src.outputs[year][month]['energy_output_prim_units_sec'] = hfo_enr

                                            else:
                                                alt_src.outputs[year][month]['energy_output_prim_units'] += \
                                                    (alt_src_nfail_cover * en_per_fail)
                                            num_fails_not_cov -= alt_src_nfail_cover

                                            # if potential failures have been reduced to zero
                                            # then further sources don't need to tried.
                                            if num_fails_not_cov <= 0:
                                                break
                                # Calculate Instant Backup Potential Power
                                ins_backup_pot_pwr = self.calc_ins_backup_pwr_pot(year, month)

                                if ins_backup_pot_pwr >= critical_load:
                                    num_failures = num_fails_not_cov
                                else:
                                    num_failures = num_pot_failures
                                failure_duration = num_fails_not_cov * src.meta.avg_failure_time

                                month_data[f'{src_name} Failures mitigated'] = num_pot_failures - num_failures
                                month_data[f'{src_name} Unavailability, hrs'] = failure_duration

                                src.outputs[year][month]['num_pot_failures'] = num_pot_failures
                                src.outputs[year][month]['num_failures'] = num_failures
                                src.outputs[year][month]['failure_duration'] = failure_duration

                            # Energy output calculation for stable sources (including failure adjustments)
                            print(f"Finding the energy output for {src_name}")

                            gen_pot_enr_op = self.get_gen_ener_op(src_name, year, month)
                            gen_enr_op = min(month_rem_enr_req, gen_pot_enr_op)
                            month_rem_enr_req -= gen_enr_op
                            if src.source_type == 'Grid':
                                enr_pk, enr_nonpk = self.grid_pk_to_offpk(month, gen_enr_op)
                                src.outputs[year][month]['energy_output_peak'] += enr_pk
                                src.outputs[year][month]['energy_output_offpeak'] += enr_nonpk
                                month_data['Grid Peak Energy, MWh'] = src.outputs[year][month]['energy_output_peak']
                                month_data['Grid Off Peak Energy, MWh'] = src.outputs[year][month][
                                    'energy_output_offpeak']

                            elif src.source_type == 'HFO+Gas Generator':
                                gas_enr, hfo_enr = src.gas_hfo_enr_op(gen_enr_op)
                                src.outputs[year][month]['energy_output_prim_units'] += gas_enr
                                src.outputs[year][month]['energy_output_prim_units_sec'] += hfo_enr
                                month_data['HFO+Gas Gen, Energy from HFO, MWh'] = src.outputs[year][month][
                                    'energy_output_prim_units_sec']
                                month_data['HFO+Gas Gen, Energy from Gas, MWh'] = src.outputs[year][month][
                                    'energy_output_prim_units']

                            else:
                                src.outputs[year][month]['energy_output_prim_units'] += gen_enr_op
                                month_data[f"{src_name} Output in MWh"] = \
                                    self.sources_dict[src_name].outputs[year][month]['energy_output_prim_units']

                    if 'Diesel Generator' in self.sources_dict:
                        print("Finding the energy output for Diesel Generator")
                        src_name = 'Diesel Generator'
                        gen_pot_enr_op = self.get_gen_ener_op(src_name, year, month)
                        gen_enr_op = min(month_rem_enr_req, gen_pot_enr_op)
                        month_rem_enr_req -= gen_enr_op
                        self.sources_dict[src_name].outputs[year][month]['energy_output_prim_units'] += gen_enr_op
                        month_data[f"{src_name} Output in MWh"] = \
                            self.sources_dict[src_name].outputs[year][month]['energy_output_prim_units']

                    month_data['Final Unserved Energy Req in MWh'] = month_rem_enr_req
                    self.energy_df.append(month_data)
                    print(
                        f"Energy data for the year {year}, month {month} determined. Unserved is {month_rem_enr_req} MWh")
            self.energy_df = pd.DataFrame(self.energy_df)

#Can this configuration meet the ramp requirements
                        ramp_power_comp = 0
                        for src in self.src_list:

                            src_day_data = src.ops_data[y][m][d]

                            if src.metadata['type']['value'] != 'BESS' and src_day_data[h]['status'] == -1 and src_day_data[h-1]['status'] == 1:

                                prev_output = src_day_data[h-1]['power_output']
                                ramp_power_comp += prev_output
                                
                                #then we need to know the power rating of the source and add it all up.
                                #sources that have failed just now would be what we should have reserve for.
                            if src.metadata['type']['value'] == 'R' and src.ops_data[y][m][d][h]['status'] == 0.5:
                                
                                #then find delta between h-1 and h and add it to sudden_drop variable.  
                                ramp_power_comp += src_day_data[h-1] - src_day_data[h-1]

    def calc_src_power_and_energy(self,y,m,d,h,power_req):
        #TO DO probably need to exclude BESS sources here
        sudden_power_drop = 0
        for priority, sources in groupby(self.src_list, key=lambda x: x.config['priority']):
            sources = list(sources)
            
            spinning_reserve_req = sources[0].metadata.get('spinning_reserve', {'value': 0})['value']
            current_power_output = 0
            current_power_capacity = 0
            
            for src in sources:
                if src.ops_data[y][m][d][h]['status'] in [-2, -3]:
                    continue  # Source is not available
                    
                max_loading = src.metadata.get('max_loading', {'value': src.config['rating']})['value'] * src.config['rating']
                power_capacity = src.ops_data[y][m][d][h]['power_capacity']
                
                # Calculate potential contribution without exceeding max_loading or remaining power requirement
                potential_power_output = min(max_loading, power_req - current_power_output, power_capacity)
                
                #if source will failin this hour then add its potential output to the sudden drop we must serve and skip the source
                if src.ops_data[y][m][d][h]['status'] == -1:
                    sudden_power_drop += potential_power_output
                    continue

                #if its solar and it will drop output then record its drop
                elif src.ops_data[y][m][d][h]['status'] == 0.5:
                    solar_output_drop = min(0,src.ops_data[y][m][d][h-1]['power_output'] - potential_power_output)
                    sudden_power_drop += solar_output_drop
                    continue

                # Update only if it contributes to meeting power requirement
                if potential_power_output > 0:
                    current_power_output += potential_power_output
                    current_power_capacity += power_capacity

                # Check if power requirement and spinning reserve are met
                if current_power_output >= power_req and (current_power_capacity - current_power_output) >= spinning_reserve_req:
                    # Update ops_data for selected sources and stop selection for this group
                    break
            
            #TODO here, seem to be assigning the same power oputput to all sources in the group, assuming that ehy are equal
            #TODO we are also not checking open assignment if the source if off.
            # Update ops_data for sources considered in this group
            for src in sources:
                src.ops_data[y][m][d][h]['power_output'] = potential_power_output
                src.ops_data[y][m][d][h]['energy_output'] = potential_power_output  # Assuming same as power_output
                src.ops_data[y][m][d][h]['status'] = 1
                src.ops_data[y][m][d][h]['spin_reserve'] = power_capacity - potential_power_output
            
            # If the power requirement is met or exceeded, adjust for next group consideration
            power_req = min(0,power_req - current_power_output)
            if power_req == 0:
                
                break  # Exit the function if no more power is needed

        return power_req, sudden_power_drop  # Return the unmet power requirement, if any
    
    #TO DO add try catch here.
    @classmethod
    def read_load_solar_data_from_folder(cls,folder_path):
        for month in range(1, 13):
            file_name = f'load_{month:02d}.xlsx'
            file_path = os.path.join(folder_path, file_name)

            cls.load_profile[month] = {}
            cls.solar_profile[month] = {}

            try:
                xls = pd.ExcelFile(file_path)
                days_of_month = [sheet for sheet in xls.sheet_names if sheet.isdigit()]

                for day in days_of_month:

                    cls.load_profile[month][int(day)] = {}
                    cls.solar_profile[month][int(day)] = {}

                    data = pd.read_excel(xls, sheet_name=day, usecols=['Total Load (KW)','Solar System (MW)'], skiprows=1, nrows=24)
                    #print(data.columns)
                    if data.isnull().values.any():
                        print(f"Warning: Blank values found in {file_name}, sheet {day}")
                    
                    cls.load_profile[month][int(day)]['Total Load (KW)'] = data['Total Load (KW)'].tolist()
                    cls.solar_profile[month][int(day)]['Solar System (MW)'] = data['Solar System (MW)'].tolist()

                print(f"Successfully read {file_name}. Days found: {len(days_of_month)}")

            except FileNotFoundError:
                print(f"File not found: {file_path}")
                raise
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")
                raise

def read_sources(self):

        # Adjust to read the whole columns B and C, headers are in row 2 (index 1)
        df = pd.read_excel(self.file_path, sheet_name='src', header=1, usecols="B:C")

        # Adjust row slices directly in DataFrame, assuming headers are correctly set at row 2
        attributes_1 = df.iloc[1:20, 0].tolist()  # Adjust slice for 'B' column
        units_1 = df.iloc[1:20, 1].tolist()       # Adjust slice for 'C' column

        # Assuming additional columns for the second range, read them separately if they're non-contiguous
        df2 = pd.read_excel(self.file_path, sheet_name='src', header=1, usecols="J:K")
        attributes_2 = df2.iloc[1:21, 0].tolist()  # Adjust slice for 'J' column
        units_2 = df2.iloc[1:21, 1].tolist()       # Adjust slice for 'K' column

        # Process sources from the first range, modify _process_source_range to handle lists directly
        self._process_source_range(df,attributes_1, units_1, start_col=3, name_row=0, data_start_row=1)

        # Process sources from the second range
        self._process_source_range(df2,attributes_2, units_2, start_col=11, name_row=0, data_start_row=1)        
    
    def _process_source_range(self, df, attributes, units, start_col, name_row, data_start_row):
        # Iterate over columns starting from the specified start_col
        for col_idx, col in enumerate(df.columns[start_col:], start=start_col):
            # Check if the column header is a non-empty string (indicating a source type)
            if pd.isnull(df.iloc[name_row, col_idx]):
                break  # Stop if a blank source type is found
            name = df.iloc[name_row, col_idx]
            values = df.iloc[data_start_row:data_start_row+len(attributes), col_idx].tolist()
            self.sources[name] = Source(name, attributes, units, values)

def calc_src_power_and_energy(self, y, m, d, h, power_req):

        sudden_power_drop = 0
        # Sort sources by priority for processing
        #sorted_sources = sorted(self.src_list, key=lambda x: x.config['priority'])
        # Group sources by priority
        for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):
            sources = list(group)
            if not sources:
                continue  # Skip empty groups
            
            # Use spinning reserve requirements from the first source in the group
            spinning_reserve_req = sources[0].config['spinning_reserve']
            current_power_output = 0
            current_power_capacity = 0
            current_spinning_reserve = 0
            
            for src in sources:
                status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                if status in [-2, -3]:  # Source is not available
                    continue
                
                # Calculate adjusted capacity based on max loading
                power_capacity = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
                max_loading_percentage = src.config['max_loading']
                adjusted_capacity = power_capacity * max_loading_percentage/100
                
                # Include operational sources and simulate output for sources about to fail or reduce output
                if status in [0, -1, 0.5]:
                    if status == -1 or status == 0.5:
                        # Calculate sudden power drop for failing or reducing output sources
                        sudden_power_drop += adjusted_capacity
                        # Assume output remains the same for status 0.5 sources
                        if status == 0.5:
                            adjusted_capacity = src.ops_data[y]['months'][m]['days'][d]['hours'][h-1]['power_output']
                    
                    # Calculate contribution proportionally
                    contribution = min(adjusted_capacity, power_req - current_power_output)
                    current_power_output += contribution
                    current_spinning_reserve += power_capacity - contribution  # Update spinning reserve
                    
                    # Record contributions
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = contribution
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = contribution  # Assuming same as power_output
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = power_capacity - contribution  # Update spinning reserve
            
            # Adjust power requirement based on the total output
            power_req = max(0,power_req - current_power_output)
            
        return power_req, sudden_power_drop

    def read_load_projection(cls, folder_path):
        input_file_path = os.path.join(folder_path, 'input_data.xlsx')
        try:
            # Read 'site_load' worksheet for site_data
            site_load_df = pd.read_excel(input_file_path, sheet_name='site_load')
            # Update site_data dictionary
            for key, value in zip(site_load_df['G'][2:], site_load_df['H'][2:]):
                cls.site_data[key] = value

            # Read range for load_projection
            load_projection_df = pd.read_excel(input_file_path, sheet_name='site_load', usecols="C:D", nrows=12, skiprows=3)
            # Update load_projection dictionary
            for year in range(1, 13):
                cls.load_projection[year] = {
                    'critical_load': load_projection_df.iloc[year-1, 0],
                    'total': load_projection_df.iloc[year-1, 1]
                }

            print("Successfully read input_data and updated dictionaries.")

        except FileNotFoundError:
            print(f"Input data file not found: {input_file_path}")
            raise
        except Exception as e:
            print(f"Error reading input data file: {e}")
            raise

def calc_src_power_and_energy(self, y, m, d, h, power_req):

        sudden_power_drop = 0
        
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

def calc_src_power_and_energy(self, y, m, d, h, power_req):

        sudden_power_drop = 0
        #240405 Instead of doing source wise spinning reserve, there will be scenario level
        spin_reserve_req = power_req * self.spinning_reserve_perc/100
        
        # Group sources by priority
        for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):

            power_req_met = False
            sources = list(group)
            if not sources or sources[0].metadata['type']['value'] == 'BESS':
                continue  # Skip empty groups
            
            # Use spinning reserve requirements from the first source in the group
            #spin_reserve_req = sources[0].config['spinning_reserve']
            
            current_power_output = 0
            current_group_capacity = 0
            current_spinning_reserve = 0
            
            for src in sources:
                status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                if status in [-2, -3]:  # Source is not available
                    continue
                
                # Calculate adjusted capacity based on max loading for current source in group
                src_capacity = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
                if src_capacity == 0:
                    continue
                current_group_capacity += src_capacity
                if status == 0:
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] = 1
                src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = -0.001
                if current_group_capacity >= power_req:

                    #if  current_group_capacity - power_req >= spin_reserve_req * current_group_capacity/100:
                    power_req_met = True
                    break
            #if power_req_met:
                #break

            if current_group_capacity > 0:

                loading_factor = power_req / current_group_capacity if current_group_capacity > 0 else 0

                if loading_factor > 1:
                    loading_factor = 1
                grp_actual_output = 0
                for src in sources:

                    status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                    
                    if src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] == -0.001:

                        #src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = 0
                        src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = loading_factor * src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
                        
                        if status == 1:
        
                            src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                            #src.ops_data[y]['months'][m]['days'][d]['hours'][h]['spin_reserve'] = src.adjusted_capacity(y,m,d,h) - src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                        
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
        #HAVE TO ADD THE BESS STUFF HERE.
        #CHECK IF SCENARIO WANTS TO USE BESS
        #THEN FORM A LIST OF BESS SOURCES
        #ASSIGN POWER, ENERGY AND REDUCE CAPACITY OF NEXT HOUR.
        if power_req != 0:
            bess_sources = [src for src in self.src_list if src.metadata['type']['value'] == 'BESS']
            if bess_sources is not None and self.bess_non_emergency_use:

                bess_total_capacity = sum(src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity'] for src in bess_sources)
                grp_loading_factor = power_req / bess_total_capacity if bess_total_capacity > 0 else 0
                if grp_loading_factor > 1:
                    grp_loading_factor = 1
                grp_actual_output = 0
                for src in bess_sources:

                    status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                    src_capacity = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = src_power_output = grp_loading_factor * src_capacity
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['energy_output'] = src_power_output
                    
                    #get the new hour
                    year, month, day, hour = self.advance_hour(y,m,d,h)
                    src.ops_data[year]['months'][month]['days'][day]['hours'][hour]['power_capacity'] = max(0, src_capacity - src_power_output)
                    grp_actual_output += src_power_output
                power_req = max(0,power_req - grp_actual_output)
                if power_req < 0.001: #1kW (math error margin)
                    power_req = 0
                    
        if power_req == 0:
            self.set_spinning_reserve(y,m,d,h, spin_reserve_req)
        return power_req, sudden_power_drop

    def set_spinning_reserve(self, y,m,d,h, spin_reserve_req):

        rem_spin_reserve_req = spin_reserve_req
        excess_power = 0
        spin_reserve_req_met = False
        for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):

            sources = list(group)
            if sources == None or sources[0].metadata['type']['value'] == 'R':
                continue

            for src in sources:

                status = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status']
                power_output = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                capacity = src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
                #then this source has been used to meet the power requirement.
                if status == 1:
                    rem_spin_reserve_req -= capacity - power_output
                #even if src is not running to meet power requirement but is req to be run just to provide spin reserve.    
                elif status == 0 and src.config['spinning_reserve']:

                    #turn the source on and run the source at its min loading
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['status'] = 1
                    src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] = src.config['min_loading'] * capacity
                    excess_power += src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']
                    rem_spin_reserve_req -= capacity - src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output']

                if rem_spin_reserve_req <=0:
                    rem_spin_reserve_req = 0
                    spin_reserve_req = True
                    break
            if spin_reserve_req:
                break

        if excess_power > 0:

            #then we need to run through each source (starting from least P)
            #and subtract it equally from each group.
            self.src_list.sort(key=lambda src: src.config['priority'],reversed = True)
            for priority, group in groupby(self.src_list, key=lambda x: x.config['priority']):

                sources = list(group)
                total_group_output = sum(src.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_output'] for src in sources)

                #check if the excess power is greater than the group's output. If it is then the group needs to run at min loading.
                group_adjusted_output = max(total_group_output - excess_power, total_group_output)

            return True

        for group in group_list:

            for src in self.src_list:
                src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                src_min_output = src_hourly_ops_data['capacity'] * src.config['min_loading']/ 100
                group['actual_output'] += src_min_output
                group['reserve'] -= group['actual_output']
                group['num_sources_req'] = 1
                rem_power_req = max(0,rem_power_req - group['min_cap'])

                #first take min pwr from sources that have to run
        for group in group_list:

            group['actual_output'] += min(group['min_cap'], rem_power_req)
            rem_power_req = max(0,rem_power_req - group['min_cap'])
            group['reserve'] -= group['actual_output']
                
        
        #then take power by group priority
        #need to respect spin reserve here
        for group in group_list:
            group['actual_output'] += min(group['capacity'],rem_power_req)
            rem_power_req = max(0,rem_power_req - group['actual_output'])
            group['reserve'] -= group['actual_output']
            rem_spin_reserve_req = max(0,rem_spin_reserve_req- group['reserve'])
            if rem_power_req == 0 and rem_spin_reserve_req == 0:
                break


     for src in sources:
                src_hourly_ops_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
                
                if src_hourly_ops_data['status'] in [-2, -3] or src_hourly_ops_data['capacity'] ==0:  # Source is not available
                    continue
                #run src at min load and save status. check if req contrib from group to SR is met. If yes, get next group
                src_hourly_ops_data['power_output'] = src_hourly_ops_data['capacity'] * src.config['min_loading']/100
                total_output += src_hourly_ops_data['power_output']
                #0.1 status is to temporarily identify which sources were used to meet initial spin reserve.
                src_hourly_ops_data['status'] = 0.1 if src_hourly_ops_data['status'] == 0 else src_hourly_ops_data['status']
                src_hourly_ops_data['reserve'] = src_hourly_ops_data['capacity'] - src_hourly_ops_data['power_output']
                grp_reserve_req_contrib = max(0, grp_reserve_req_contrib - src_hourly_ops_data['reserve'])
                if grp_reserve_req_contrib == 0:
                    break