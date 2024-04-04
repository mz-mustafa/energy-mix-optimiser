from project import Project
import pandas as pd

def dict_to_csv(filename="load_output.csv"):
    # Prepare lists for each column
    months, days, hours, power_outputs = [], [], [], []
    
    for month, days_data in Project.load_profile.items():
        for day, hourly_data in days_data.items():
            for hour, power_output in enumerate(hourly_data):
                # Append data for each column
                months.append(month)
                days.append(day)
                hours.append(hour)
                power_outputs.append(power_output)
    
    # Create a DataFrame
    df = pd.DataFrame({
        'Month': months,
        'Day': days,
        'Hour': hours,
        'Power Output': power_outputs
    })
    
    # Write the DataFrame to CSV
    df.to_csv(filename, index=False)
    print(f"Data successfully written to {filename}")


if __name__ == "__main__":
        
        #global source_manager
        #print('Reading Pre req data')
        Project.read_load_projection("data")
        Project.read_load_solar_data_from_folder('data/load_solar_profile')
        Project.create_load_data()
        #print(Project.solar_profile)
        dict_to_csv()
        #print(Project.load_data)


