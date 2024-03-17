from source_meta import SolarMeta, WindMeta, GridMeta, \
    GasGenMeta, HFOGenMeta, TriFuelGenMeta, BESSMeta, DGenMeta, PPAMeta, ExistingGasGenMeta

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



class SolarSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n, 'Solar',src_p)
        self.meta = SolarMeta()

    def calc_output_power(self, current_year, month, hour):

        # Calculate degradation for each year's capacity and then sum up
        degraded_capacity_mw = 0
        degradation_rate = self.meta.degradation if hasattr(self.meta, 'degradation') else 0

        for year, yr_data in self.inputs.items():
            if isinstance(year, int) and year <= current_year:
                years_of_operation = current_year - year
                degradation_factor = 1 - (degradation_rate * years_of_operation/100)
                degraded_capacity_mw += yr_data['count_prim_units'] * yr_data['rating_prim_units'] * degradation_factor

        return degraded_capacity_mw * self.meta.output_data[month][hour]['kw_pv_output_per_mw'] / 1000

    def calc_output_energy(self, current_year, month):

        # Calculate degradation for each year's capacity and then sum up
        degraded_capacity_mw = 0
        degradation_rate = self.meta.degradation if hasattr(self.meta, 'degradation') else 0

        for year, yr_data in self.inputs.items():
            if isinstance(year, int) and year <= current_year:
                years_of_operation = current_year - year
                degradation_factor = 1 - (degradation_rate * years_of_operation/100)
                degraded_capacity_mw += yr_data['count_prim_units'] * yr_data['rating_prim_units'] * degradation_factor

                if year == current_year:
                    break

        return degraded_capacity_mw * self.meta.output_data[month]['mnth_ener_op_per_MW']

class BESSSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'BESS',src_p)
        self.meta = BESSMeta()


class WindSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'Wind', src_p)
        self.meta = WindMeta()

    def calc_output_power(self, current_year, month, hour):
        # Calculate degradation for each year's capacity and then sum up
        degraded_capacity_mw = 0
        degradation_rate = self.meta.degradation if hasattr(self.meta, 'degradation') else 0

        for year, yr_data in self.inputs.items():
            if isinstance(year, int) and year <= current_year:
                years_of_operation = current_year - year
                degradation_factor = 1 - (degradation_rate * years_of_operation/100)
                degraded_capacity_mw += yr_data['count_prim_units'] * yr_data['rating_prim_units'] * degradation_factor

                if year == current_year:
                    break

        wind_speed = self.meta.output_data[month][hour]['windspeed_at_100m']
        output_multiplier = 0
        if wind_speed >= 7.5:
            output_multiplier = 1
        elif wind_speed >= 6:
            output_multiplier = self.meta.output_multiplier_3
        elif wind_speed >= 5:
            output_multiplier = self.meta.output_multiplier_2
        elif wind_speed >= 4:
            output_multiplier = self.meta.output_multiplier_1

        return degraded_capacity_mw * output_multiplier

    def calc_output_energy(self, current_year, month):

        return sum(self.calc_output_power(current_year,month, hour) for hour in range(1, 25))



class GridSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'Grid', src_p)
        self.meta = GridMeta()
        #add output structure here to accomodate peak and off peak units
        self.extend_output_structure()

    def extend_output_structure(self):
        # Extend output structure with Grid Source specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):
                self.outputs[year][month]['energy_output_peak'] = 0
                self.outputs[year][month]['energy_output_offpeak'] = 0
                self.outputs[year][month]['peak_enr_charges'] = 0
                self.outputs[year][month]['offpeak_enr_charges'] = 0
                self.outputs[year][month]['fixed_charges'] = 0

class PPASource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'PPA', src_p)
        self.meta = PPAMeta()
        #add output structure here to accomodate peak and off peak units
        self.extend_output_structure()

    def extend_output_structure(self):
        # Extend output structure with Grid Source specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):
                self.outputs[year][month]['enr_charges'] = 0
                self.outputs[year][month]['fixed_charges'] = 0
    
    def calc_output_power(self, current_year, month, hour):

        # Calculate degradation for each year's capacity and then sum up
        degraded_capacity_mw = 0
        degradation_rate = self.meta.degradation if hasattr(self.meta, 'degradation') else 0

        for year, yr_data in self.inputs.items():
            if isinstance(year, int) and year <= current_year:
                years_of_operation = current_year - year
                degradation_factor = 1 - (degradation_rate * years_of_operation/100)
                degraded_capacity_mw += yr_data['count_prim_units'] * yr_data['rating_prim_units'] * degradation_factor

        return degraded_capacity_mw * self.meta.output_data[month][hour]['kw_pv_output_per_mw'] / 1000

    def calc_output_energy(self, current_year, month):

        # Calculate degradation for each year's capacity and then sum up
        degraded_capacity_mw = 0
        degradation_rate = self.meta.degradation if hasattr(self.meta, 'degradation') else 0

        for year, yr_data in self.inputs.items():
            if isinstance(year, int) and year <= current_year:
                years_of_operation = current_year - year
                degradation_factor = 1 - (degradation_rate * years_of_operation/100)
                degraded_capacity_mw += yr_data['count_prim_units'] * yr_data['rating_prim_units'] * degradation_factor

                if year == current_year:
                    break

        return degraded_capacity_mw * self.meta.output_data[month]['mnth_ener_op_per_MW']

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
            self.inputs[year]['fuel_eff'] = 100

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

class ExistingGasGenSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n, 'Existing Gas Generators', src_p)
        self.meta = ExistingGasGenMeta()
        self.extend_input_structure()
        self.extend_output_structure()

    def extend_input_structure(self):
        # Extend input structure with GasGenSource specific keys
        for year in range(self.n + 1):  # Use range based on n to iterate over years
            self.inputs[year]['rating_backup_units'] = 0
            self.inputs[year]['count_backup_units'] = 0
            self.inputs[year]['perc_rated_output'] = 0 
            self.inputs[year]['fuel_eff'] = 100

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

class HFOGenSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'HFO Generator', src_p)
        self.meta = HFOGenMeta()
        self.extend_input_structure()
        self.extend_output_structure()

    def extend_input_structure(self):
        # Extend input structure with GasGenSource specific keys
        for year in range(self.n + 1):
            self.inputs[year]['rating_backup_units'] = 0
            self.inputs[year]['count_backup_units'] = 0
            self.inputs[year]['perc_rated_output'] = 0
            self.inputs[year]['fuel_eff'] = 100

        self.inputs['chp_operation'] = False
        self.inputs['fuel_type'] = 'HFO'

    def extend_output_structure(self):
        # Extend output structure with GasGenSource specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):
                self.outputs[year][month]['energy_output_backup_units'] = 0
                self.outputs[year][month]['energy_free_cooling'] = 0
                self.outputs[year][month]['var_opex'] = 0
                self.outputs[year][month]['fuel_charges'] = 0
                for hour in range(1, 25):
                    self.outputs[year][month][hour]['power_output_backup_units'] = 0
                    self.outputs[year][month][hour]['loading_backup_units'] = 0


class TrifuelGenSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'HFO+Gas Generator', src_p)
        self.meta = TriFuelGenMeta()
        self.extend_input_structure()
        self.extend_output_structure()

    def extend_input_structure(self):
        # Extend input structure with GasGenSource specific keys
        for year in range(self.n + 1):
            self.inputs[year]['rating_backup_units'] = 0
            self.inputs[year]['count_backup_units'] = 0
            self.inputs[year]['perc_rated_output'] = 0
            self.inputs[year]['fuel_eff'] = 100

        self.inputs['chp_operation'] = False
        self.inputs['fuel_type'] = 'RLNG'
        self.inputs['sec_fuel_type'] = 'HFO'

    def extend_output_structure(self):
        # Extend output structure with GasGenSource specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):
                self.outputs[year][month]['energy_output_prim_units_sec'] = 0
                self.outputs[year][month]['energy_output_backup_units'] = 0
                self.outputs[year][month]['energy_output_backup_units_sec'] = 0
                self.outputs[year][month]['energy_free_cooling'] = 0
                self.outputs[year][month]['var_opex'] = 0
                self.outputs[year][month]['fuel_charges'] = 0
                self.outputs[year][month]['fuel_charges_sec'] = 0
                for hour in range(1, 25):
                    self.outputs[year][month][hour]['power_output_backup_units'] = 0
                    self.outputs[year][month][hour]['loading_backup_units'] = 0

    def gas_hfo_enr_op(self, energy):

        gas_enr = energy * self.meta.gas_op_prop/100
        hfo_enr = energy - gas_enr
        return gas_enr, hfo_enr

class DieselGenSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n,'Diesel Generator', src_p)
        self.meta = DGenMeta()
        self.extend_input_structure()
        self.extend_output_structure()

    def extend_input_structure(self):
        # Extend input structure with GasGenSource specific keys
        for year in range(self.n + 1):
            self.inputs[year]['perc_rated_output'] = 0
            self.inputs[year]['fuel_eff'] = 100
        self.inputs['fuel_type'] = 'Diesel'
        

    def extend_output_structure(self):
        # Extend output structure with Diesel Gen specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):  # Use range for months
                self.outputs[year][month]['var_opex'] = 0
                self.outputs[year][month]['fuel_charges'] = 0