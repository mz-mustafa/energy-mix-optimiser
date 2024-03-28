import pandas as pd
import random
from project import Project

class Source:
    def __init__(self, name, attributes, units, values, start_year, rating, rating_unit,reserve_perc,priority,min_loading,max_loading):
        self.name = name
        #TODO put reserve_perc in metadata
        #TODO put solar_sudden_drops in metadata
        self.metadata = {attr: {'unit': unit, 'value': value} for attr, unit, value in zip(attributes, units, values)}
        self.config = {
            'start_year' : start_year,
            'rating' : rating,
            'rating_unit' : rating_unit,
            'priority' : priority,
            'min_loading': min_loading,
            'max_loading': max_loading,
        }
        self.ops_data = self._initialize_years()
        # Status 0 is off, 1 is on, -1 is downtime, -2 is failure, -3 doesn't exist
        # for BESS Status 0 is trickel charge, 1 is discharging, 2 is charging, -1 is downtime, -2 is failure, -3 doesn't exist
        
    
    def display_info(self):
        for attr, info in self.data.items():
            print(f"{attr} ({info['unit']}): {info['value']}")

    def _initialize_years(self):
        years_data = {}
        for year in range(1, 13):
            is_future = year < self.config['start_year']
            year_dict = {
                'source_present': 0 if is_future else 1,
                'year_potential_failures': 0,
                'year_failures_mitigated': 0,
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
            year_dict['months'] = self._initialize_months(year, is_future)
            years_data[year] = year_dict
        return years_data

    def _initialize_months(self, year, is_future):
        months_data = {}
        for month in range(1, 13):
            days_in_month = 28 if month == 2 else 30 if month in [4, 6, 9, 11] else 31
            month_dict = {
                'month_potential_failures': 0,
                'month_failures_mitigated': 0,
                'month_downtime': 0,
                'month_energy_output': 0,
                'month_cost_of_operation': 0,
                'month_fuel_cost': 0,
                'month_fixed_opex': 0,
                'month_var_opex': 0,
                'month_depreciation': 0,
                'month_ppa_cost': 0,
                'month_operation_hours': 0,
                'days': {}
            }
            month_dict['days'] = self._initialize_days(year, month, is_future, days_in_month)
            months_data[month] = month_dict
        return months_data

    def _initialize_days(self, year, month, is_future, days_in_month):
        days_data = {}
        for day in range(1, days_in_month + 1):
            day_dict = {
                'avg_power_output': 0,
                'min_power_output': 0,
                'max_power_output': 0,
                'day_energy_output': 0,
                'failure_occurrence': 0,
                'failure_mitigation': 0,
                'operation_hours': 0,
                'downtime': 0,
                'hours': self._initialize_hours(is_future)
            }
            days_data[day] = day_dict
        return days_data

    def _initialize_hours(self, is_future):
        hours_data = {}
        status = -3 if is_future else 0
        for hour in range(24):
            hours_data[hour] = {
                'power_capacity': 0,
                'power_output': 0,
                'energy_output': 0,
                'status': status
            }
        return hours_data

    def seed_solar_disturbances(self):
        # Ensure this function only applies to renewable sources
        if self.metadata['type']['value'] != 'R':
            print("This function is only applicable to renewable sources.")
            return
        
        for year, year_data in self.ops_data.items():
            if year_data.get('source_present') == 0:
                continue  # Skip non-existent years for this source
            
            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    daily_hours_to_flag = self.metadata.get('solar_sudden_drops', {'value': 0})['value']
                    if daily_hours_to_flag <= 0:
                        continue  # Skip if no disturbances are to be seeded
                    
                    candidate_hours = []  # Hours eligible for being flagged
                    
                    for hour in range(24):
                        if day_data['hours'][hour]['status'] != 1:
                            continue  # Skip hours not in operation
                        
                        # Check for a negative power output delta between h and h-1
                        if hour > 0 and day_data['hours'][hour]['power_output'] < day_data['hours'][hour - 1]['power_output']:
                            candidate_hours.append(hour)
                    
                    # Randomly select hours to flag, up to the daily limit
                    hours_to_flag = random.sample(candidate_hours, min(len(candidate_hours), daily_hours_to_flag))
                    
                    for hour in hours_to_flag:
                        day_data['hours'][hour]['status'] = 0.5  # Flag as sudden power reduction 
    #TODO make sure hour 0 is never seeded
    #TODO only seed hours for years in which source is available (check year level attribute)
    def seed_availabilty(self):
        for year, year_data in self.ops_data.items():
            days_of_year = []
            # Generate a flat list of all day-hour combinations
            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    for hour in range(24):
                        days_of_year.append((month, day, hour))
            
            # Randomly select days for failures, ensuring unique days
            failure_days = random.sample(days_of_year, self.metadata['num_annual_fails']['value'])

            for month, day, fail_hour in failure_days:
                # Mark the failure hour
                year_data['months'][month]['days'][day]['Hours'][fail_hour]['Status'] = -2
                
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
                    year_data['months'][month]['days'][day]['Hours'][fail_hour]['Status'] = -1
                    downtime -= 1
        self.aggregate_availability()
    #TODO skip years in which source is not present
    def aggregate_availability(self):
        # Iterate through all levels of data to update day, month, and year aggregates
        for year, year_data in self.ops_data.items():
            for month, month_data in year_data['months'].items():
                for day, day_data in month_data['days'].items():
                    # Count the failure occurrences and downtime at the day level
                    failure_occurrence = sum(1 for hour in day_data['Hours'].values() if hour['Status'] == -2)
                    downtime = sum(1 for hour in day_data['Hours'].values() if hour['Status'] == -1) + failure_occurrence
                    
                    # Update day level data
                    day_data['Failure_occurrence'] = failure_occurrence
                    day_data['Downtime'] = downtime
                    
                    # Accumulate counts for the month level data
                    month_data['month_potential_failures'] += failure_occurrence
                    month_data['month_downtime'] += downtime
                
                # Accumulate counts for the year level data
                year_data['year_potential_failures'] += month_data['month_potential_failures']
                year_data['year_downtime'] += month_data['month_downtime']   

    def update_power_capacity(self):
        for year, year_data in self.ops_data.items():
            #TODO instead of calc, just use year level ops data parameter
            # Calculate years of operation differently, considering each year as present during iteration
            years_of_operation = year - self.config['start_year']
            
            if years_of_operation < 0:
                continue  # Skip years before the source's start year

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
                                degraded_rating = self.config['rating'] * ((1 - annual_degradation_rate) ** years_of_operation)
                                power_capacity = degraded_rating

                        elif self.metadata['type']['value'] == 'R':
                            # Use solar profile for renewable sources
                            # Assuming the Project class and solar_profile structure allows this direct access
                            solar_output = Project.solar_profile[month][day][hour]
                            power_capacity = (solar_output / 5) * self.config['rating']

                        # Update power_capacity for the hour
                        hour_data['power_capacity'] = power_capacity


class SourceManager:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sources = {}
        try:
            self.read_sources()
        except Exception as e:
            print(f"Failed to load source data: {e}")
            raise

    def read_sources(self):
        df = pd.read_excel(self.file_path, sheet_name='src_ppa')
        
        # Read the first set of attribute names and units
        attributes_1 = df['B'][1:18].tolist()  # Adjust based on actual data range
        units_1 = df['C'][1:18].tolist()

        # Process sources from the first range (D2 and rightwards)
        self._process_source_range(df, attributes_1, units_1, start_col=3, name_row=0, data_start_row=1)

        # Read the second set of attribute names and units
        attributes_2 = df['J'][1:19].tolist()  # Adjust for the new range
        units_2 = df['K'][1:19].tolist()

        # Process sources from the second range (L2 and rightwards)
        self._process_source_range(df, attributes_2, units_2, start_col=11, name_row=0, data_start_row=1)

    def _process_source_range(self, df, attributes, units, start_col, name_row, data_start_row):
        # Iterate over columns starting from the specified start_col
        for col_idx, col in enumerate(df.columns[start_col:], start=start_col):
            # Check if the column header is a non-empty string (indicating a source type)
            if pd.isnull(df.iloc[name_row, col_idx]):
                break  # Stop if a blank source type is found
            name = df.iloc[name_row, col_idx]
            values = df.iloc[data_start_row:data_start_row+len(attributes), col_idx].tolist()
            self.sources[name] = Source(name, attributes, units, values)

    def get_source_by_name(self, name):
        return self.sources.get(name)
    
    def select_sources(source_manager):
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

def collect_source_config():
    # Collect configuration data from the user
    # This is a placeholder; implement according to your application's needs
    return {
        'start_year': input("Start Year: "),
        'rating': input("Rating: "),
        'rating_unit': input("Rating Unit: "),
        'reserve_perc': input("Reserve Percentage: ")
    }
