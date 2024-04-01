import pandas as pd
import random
from project import Project

class Source:
    def __init__(self, name, attributes, units, values):
        self.name = name
        self.attributes = attributes
        self.units = units
        self.values = values
        self.metadata = {attr: {'unit': unit, 'value': value} for attr, unit, value in zip(attributes, units, values)}
        self.config = {}
        self.ops_data = {}
        # Status 0 is off, 1 is on, -2 is downtime, -1 is failure, -3 doesn't exist
        # for BESS Status 0 is trickel charge, 1 is discharging, 2 is charging, -1 is downtime, -2 is failure, -3 doesn't exist
        
    def configure(self, start_year, end_year, rating, rating_unit, spin_reserve, priority, min_loading, max_loading):
        # Update the config dictionary with new key-value pairs
        self.config['start_year'] = start_year
        self.config['end_year'] = end_year
        self.config['rating'] = rating
        self.config['rating_unit'] = rating_unit
        self.config['priority'] = priority
        self.config['min_loading'] = min_loading
        self.config['max_loading'] = max_loading
        self.config['spinning_reserve'] = spin_reserve
        # Example calculation for 'capex', assumes 'capital_cost_baseline' is in metadata and 'inflation_rate' is available
        self.config['capex'] = rating * self.metadata.get('capital_cost_baseline', {'value': 0})['value'] * (1 + Project.inflation_rate)**(start_year-1)
        self.ops_data = self._initialize_years()
        self.update_power_capacity()
        self.seed_failures()
        self.seed_solar_reductions()
        self.aggregate_failure_reduction_stats()
    
    def display_info(self):
        for attr, info in self.data.items():
            print(f"{attr} ({info['unit']}): {info['value']}")

    def _initialize_years(self):
        years_data = {}
        for year in range(1, 13):
            if year >= self.config['start_year'] and year <= self.config['end_year']:
                exists = True
            else:
                exists = False
            year_dict = {
                'source_present': 1 if exists else 0,
                'year_failures': 0,
                'year_reductions' : 0,
                'year_downtime': 0,
                'year_energy_output': 0,
                'year_cost_of_operation': 0,
                'year_fuel_cost': 0,
                'year_fixed_opex': 0,
                'year_var_opex': 0,
                'year_depreciation': 0,
                'year_ppa_cost': 0,
                'year_operation_hours': 0,
                'months': {}
            }
            year_dict['months'] = self._initialize_months(year, exists)
            years_data[year] = year_dict
        return years_data

    def _initialize_months(self, year, exists):
        months_data = {}
        for month in range(1, 13):
            days_in_month = 28 if month == 2 else 30 if month in [4, 6, 9, 11] else 31
            month_dict = {
                'month_failures': 0,
                'month_reductions' : 0,
                'month_downtime': 0,
                'month_energy_output': 0,
                'month_cost_of_operation': 0,
                'month_fuel_cost': 0,
                'month_fixed_opex': 0,
                'month_var_opex': 0,
                'month_ppa_cost': 0,
                'month_operation_hours': 0,
                'days': {}
            }
            month_dict['days'] = self._initialize_days(year, month, exists, days_in_month)
            months_data[month] = month_dict
        return months_data

    def _initialize_days(self, year, month, exists, days_in_month):
        days_data = {}
        for day in range(1, days_in_month + 1):
            day_dict = {
                'avg_power_output': 0,
                'min_power_output': 0,
                'max_power_output': 0,
                'day_energy_output': 0,
                'failure_events': 0,
                'reduction_events' : 0,
                'operation_hours': 0,
                'downtime': 0,
                'hours': self._initialize_hours(exists)
            }
            days_data[day] = day_dict
        return days_data

    def _initialize_hours(self, exists):
        hours_data = {}
        status = 0 if exists else -3
        for hour in range(24):
            hours_data[hour] = {
                'power_capacity': 0,
                'power_output': 0,
                'energy_output': 0,
                'spin_reserve' : 0,
                'status': status
            }
        return hours_data

    def seed_solar_reductions(self):
        # Ensure this function only applies to renewable sources
        if self.metadata['type']['value'] != 'R':
            return
        
        daily_hours_to_flag = self.metadata.get('solar_sudden_drops', {'value': 0})['value']

        if daily_hours_to_flag <= 0:
            return  # Skip if no disturbances are to be seeded

        for year, year_data in self.ops_data.items():
            if year_data.get('source_present') == 0:
                continue  # Skip non-existent years for this source
            
            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    
                    candidate_hours = []  # hours eligible for being flagged
                    
                    for hour in range(1,24):
                        
                        # Check for a negative power output delta between h and h-1
                        if hour > 0 and day_data['hours'][hour]['power_capacity'] < day_data['hours'][hour - 1]['power_capacity']:
                            candidate_hours.append(hour)
                    
                    # Randomly select hours to flag, up to the daily limit
                    hours_to_flag = random.sample(candidate_hours, min(len(candidate_hours), daily_hours_to_flag))
                    
                    for hour in hours_to_flag:
                        day_data['hours'][hour]['status'] = 0.5  # Flag as sudden power reduction 
    
    def seed_failures(self):

        annual_fails = self.metadata['num_annual_fails']['value']
        if annual_fails <= 0:
            return
        for year, year_data in self.ops_data.items():

            if year_data.get('source_present') == 0:
                continue  # Skip non-existent years for this source
            days_of_year = []
            # Generate a flat list of all day-hour combinations
            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    for hour in range(1,24):
                        days_of_year.append((month, day, hour))
            
            # Randomly select days for failures, ensuring unique days
            failure_days = random.sample(days_of_year, annual_fails)

            for month, day, fail_hour in failure_days:
                # Mark the failure hour
                year_data['months'][month]['days'][day]['hours'][fail_hour]['status'] = -1
                
                # Apply downtime for subsequent hours
                downtime = self.metadata['downtime_per_fail']['value'] - 1
                while downtime > 0:
                    fail_hour += 1
                    if fail_hour >= 24:
                        fail_hour = 0
                        day += 1
                        # Move to the next month if the days exceed the month's length
                        if day > len(year_data['months'][month]['days']):
                            day = 1
                            month += 1
                            # Reset to January if the month exceeds the year
                            if month > 12:
                                month = 1
                    year_data['months'][month]['days'][day]['hours'][fail_hour]['status'] = -2
                    downtime -= 1

    def aggregate_failure_reduction_stats(self):
        # Iterate through all levels of data to update day, month, and year aggregates
        for year, year_data in self.ops_data.items():

            if year_data.get('source_present') == 0:
                continue  # Skip non-existent years for this source

            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    # Count the failure occurrences and downtime at the day level
                    failure_occurrence = sum(1 for hour in day_data['hours'].values() if hour['status'] == -1)
                    downtime = sum(1 for hour in day_data['hours'].values() if hour['status'] == -2) + failure_occurrence
                    reductions = sum(1 for hour in day_data['hours'].values() if hour['status'] == 0.5)
                    # Update day level data
                    day_data['failure_events'] = failure_occurrence
                    day_data['downtime'] = downtime
                    day_data['reduction_events'] = reductions
                    
                    # Accumulate counts for the month level data
                    month_data['month_failures'] += failure_occurrence
                    month_data['month_downtime'] += downtime
                    month_data['month_reductions'] += reductions
                
                # Accumulate counts for the year level data
                year_data['year_failures'] += month_data['month_failures']
                year_data['year_downtime'] += month_data['month_downtime']
                year_data['year_reductions'] += month_data['month_reductions']    

    def update_power_capacity(self):
        for year, year_data in self.ops_data.items():

            if year_data.get('source_present') == 0:
                continue  # Skip non-existent years for this source
            
            years_of_operation = year - self.config['start_year'] 

            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    for hour, hour_data in day_data['hours'].items():
                        power_capacity = 0

                        if self.metadata['type']['value'] == 'NR' and self.metadata['finance']['value'] == 'PPA':
                            power_capacity = self.config['rating']

                        elif self.metadata['type']['value'] == 'NR' and self.metadata['finance']['value'] == 'CAPTIVE':
                            # Check for 'annual_degradation' key in metadata
                            if 'annual_degradation' in self.metadata:
                                annual_degradation_rate = self.metadata['annual_degradation']['value']
                                # Apply annual degradation
                                degraded_rating = self.config['rating'] * ((1 - (annual_degradation_rate/100)) ** years_of_operation)
                                power_capacity = degraded_rating

                        elif self.metadata['type']['value'] == 'R':
                            # Use solar profile for renewable sources
                            # Assuming the Project class and solar_profile structure allows this direct access
                            solar_output = Project.solar_profile[month][day][hour]
                            power_capacity = (solar_output / 5) * self.config['rating']

                        # Update power_capacity for the hour
                        hour_data['power_capacity'] = power_capacity

    #TODO add degradation here, check if exists in metadata.
    def adjusted_capacity(self,y,m,d,h):

        src_capacity = self.ops_data[y]['months'][m]['days'][d]['hours'][h]['power_capacity']
        max_loading_percentage = self.config['max_loading']
        return src_capacity * max_loading_percentage/100


    def aggregate_day_level(self):

        for year in self.ops_data:
            for month in self.ops_data[year]['months']:
                for day in self.ops_data[year]['months'][month]['days']:
                    hours = self.ops_data[year]['months'][month]['days'][day]['hours'].values()

                    # Directly calculate metrics without intermediate steps
                    self.ops_data[year]['months'][month]['days'][day].update({
                        'avg_power_output': sum(hour['power_output'] for hour in hours) / 24,
                        'min_power_output': min(hour['power_output'] for hour in hours),
                        'max_power_output': max(hour['power_output'] for hour in hours),
                        'day_energy_output': sum(hour['energy_output'] for hour in hours),
                        'failure_events': sum(hour['status'] == -1 for hour in hours),
                        'reduction_events': sum(hour['status'] == 0.5 for hour in hours),
                        'operation_hours': sum(hour['status'] == 1 for hour in hours),
                        'downtime': sum(hour['status'] == -2 for hour in hours),
                    })

    def aggregate_month_level(self):

        for year in self.ops_data:
            for month in self.ops_data[year]['months']:
                month_data = self.ops_data[year]['months'][month]['days']
                
                # Calculate month-level metrics using list comprehensions
                self.ops_data[year]['months'][month].update({
                    'month_failures': sum(day['failure_events'] for day in month_data.values()),
                    'month_reductions': sum(day['reduction_events'] for day in month_data.values()),
                    'month_downtime': sum(day['downtime'] for day in month_data.values()),
                    'month_energy_output': sum(day['day_energy_output'] for day in month_data.values()),
                    'month_operation_hours': sum(day['operation_hours'] for day in month_data.values()),
                    'month_fuel_cost': sum(day['day_energy_output'] for day in month_data.values()) * self.metadata['fuel_consumption']['value'] * self.metadata['fuel_cost']['value'],
                    'month_fixed_opex': self.config['rating'] * self.metadata['opex_baseline_fixed']['value'],
                    'month_var_opex': sum(day['day_energy_output'] for day in month_data.values()) * self.metadata['opex_baseline_var']['value'],
                    'month_ppa_cost': (sum(day['day_energy_output'] for day in month_data.values()) * self.metadata['tariff_baseline_var']['value']) + (self.config['rating'] * self.metadata['tariff_baseline_fixed']['value']),
                })
                
                # After calculating the individual components, update 'month_cost_of_operation'
                month_costs = self.ops_data[year]['months'][month]
                month_costs['month_cost_of_operation'] = (month_costs['month_fuel_cost'] +
                                                        month_costs['month_fixed_opex'] +
                                                        month_costs['month_var_opex'] +
                                                        month_costs['month_ppa_cost'])

    def aggregate_year_level(self):
        for year in self.ops_data:
            year_data = self.ops_data[year]['months']

            depreciation =  self.config['capex']/ self.metadata['useful_life']['value']
            
            # Direct assignment of year-level metrics using list comprehensions
            self.ops_data[year].update({
                'year_failures': sum(month_data['month_failures'] for month_data in year_data.values()),
                'year_reductions': sum(month_data['month_reductions'] for month_data in year_data.values()),
                'year_downtime': sum(month_data['month_downtime'] for month_data in year_data.values()),
                'year_energy_output': sum(month_data['month_energy_output'] for month_data in year_data.values()),
                'year_cost_of_operation': sum(month_data['month_cost_of_operation'] for month_data in year_data.values()),
                'year_fuel_cost': sum(month_data['month_fuel_cost'] for month_data in year_data.values()),
                'year_fixed_opex': sum(month_data['month_fixed_opex'] for month_data in year_data.values()),
                'year_var_opex': sum(month_data['month_var_opex'] for month_data in year_data.values()),
                'year_depreciation': depreciation,
                'year_ppa_cost': sum(month_data['month_ppa_cost'] for month_data in year_data.values()),
                'year_operation_hours': sum(month_data['month_operation_hours'] for month_data in year_data.values()),
            })

class SourceManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.source_types = {}
        try:
            self.read_sources()
        except Exception as e:
            print(f"Failed to load source data: {e}")
            raise

    def read_sources(self):

        # Load the DataFrame using Excel's row 2 as headers (header=1 in Pandas)
        df = pd.read_excel(self.file_path, sheet_name='src', header=1)

        # Extract attributes and units for the first set of sources (SRC_1 to SRC_5)
        attributes = df['ATTRIBUTE'].tolist()
        units = df['UNIT'].tolist()

        # Extract attributes and units for the second set of sources (SRC_6), using renamed column headers
        attributes_captive = df['ATTR_CAPTIVE'].tolist()  # Use dropna() to ignore empty cells
        units_captive = df['UNIT_CAPTIVE'].tolist()

        # Source names are in columns D to H for the first 5 sources
        source_columns = list(df.columns[3:8])  # Adjust if there are more sources

        # Initialize Source objects for SRC_1 to SRC_5
        for source_column in source_columns:
            values = df[source_column].tolist()
            self.source_types[source_column] = Source(source_column, attributes, units, values)

        # Column L (index 11) for SRC_6, using the second set of attributes and units
        values_captive = df[df.columns[11]].tolist()
        self.source_types[df.columns[11]] = Source(df.columns[11], attributes_captive, units_captive, values_captive)

    def get_source_types_by_name(self, name):

        if name in self.source_types:
            # Get the template source object
            template_src = self.source_types[name]
            new_src = Source(template_src.name, template_src.attributes, template_src.units, template_src.values)
            return new_src
        else:
            return None
    
    #don't think we need this.
    """
    def select_source_types(source_manager):
        print("Available sources:")
        for name in source_manager.sources.keys():
            print(name)
        
        selected_sources = []
        while True:
            choice = input("Enter source name to select (or 'done' to finish): ")
            if choice == 'done':
                break
            if choice in source_manager.sources:
                config = collect_source_config()
                source_instance = Source(choice, **source_manager.sources[choice].metadata, config=config)
                selected_sources.append(source_instance)
            else:
                print("Source not found.")
        return selected_sources
    """
def collect_source_config():

    return {
        'start_year': input("Start Year: "),
        'rating': input("Rating: "),
        'rating_unit': input("Rating Unit: "),
        'priority': input("Priority: "),
        'min_loading' : input("Min Loading"),
        'max_loading' : input("Max Loading"),
        'spinning_reserve' : input("Spinning Reserve"),
    }