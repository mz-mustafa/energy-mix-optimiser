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

