import pandas as pd
import os

class Project:

    load_profile = {}  
    solar_profile = {}  
    site_data = {}
    load_projection = {}
    load_data = {}
    inflation_rate = 0 

    #TO DO add try catch here.
    @classmethod
    def read_load_projection(cls, folder_path):
        input_file_path = os.path.join(folder_path, 'input_data.xlsx')
        try:
            # Read 'site_load' worksheet for site_data
            site_load_df = pd.read_excel(input_file_path, sheet_name='site_load',header=2,nrows=3)
            print(site_load_df.columns)

            # Update site_data dictionary
            for key, value in zip(site_load_df['site_details_attribute'], site_load_df['value']):
                cls.site_data[key] = value

            # Read range for load_projection
            load_projection_df = pd.read_excel(input_file_path, sheet_name='site_load', usecols=["critical_load", "total_load"], nrows=12, skiprows=2)
        
            for index, row in load_projection_df.iterrows():
                year = index + 1  # Adjusting index to match your year numbering
                cls.load_projection[year] = {
                    'critical_load': row['critical_load'],
                    'total_load': row['total_load']
                    }

            print("Successfully read input_data and updated dictionaries.")

        except FileNotFoundError:
            print(f"Input data file not found: {input_file_path}")
            raise
        except Exception as e:
            print(f"Error reading input data file: {e}")
            raise
    
    @classmethod
    def read_load_solar_data_from_folder(cls, folder_path):

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

                    # Use column letters directly for usecols
                    data = pd.read_excel(xls, sheet_name=day, usecols="B:C", skiprows=1, nrows=24)
                    #print(data.columns)

                    data = data.replace('-', 0)

                    if data.isnull().values.any():
                        print(f"Warning: Blank values found in {file_name}, sheet {day}")

                    data = data.map(lambda x: x / 1000 if pd.notnull(x) else x)

                    # Assign lists to respective profiles
                    # Assuming 'B' is 'Total Load (KW)' and 'C' is 'Solar System (MW)' after adjusting for header and skiprows
                    cls.load_profile[month][int(day)] = data.iloc[:, 0].tolist()  # First column in the specified range
                    cls.solar_profile[month][int(day)] = data.iloc[:, 1].tolist()  # Second column in the specified range

                print(f"Successfully read {file_name}. Days found: {len(days_of_month)}")

            except FileNotFoundError:
                print(f"File not found: {file_path}")
                raise
            except Exception as e:
                print(f"Error processing file {file_name}: {e}")
                raise

    #TO DO add try catch here.
    @classmethod
    def create_load_data(cls):
        reference_total_load = cls.load_projection[1]['total_load']  # Total load of year 1 as reference

        # Iterate through each year in load_projection
        for year in range(1, 13):
            multiplier = cls.load_projection[year]['total_load'] / reference_total_load
            cls.load_data[year] = {}

            # Iterate through each month, day, and hour in load_profile to scale values
            for month, days in cls.load_profile.items():
                cls.load_data[year][month] = {}
                for day, hours in days.items():
                    cls.load_data[year][month][day] = [value * multiplier for value in hours]

        print("Load data creation completed.")

# The `create_load_data` method calculates the multiplier for each year based on the total_load value from `load_projection`.
# It then scales the values in `load_profile` by this multiplier for each year, creating a new structure in `load_data`.
