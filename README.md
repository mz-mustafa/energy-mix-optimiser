# Energy Mix Optimizer

An AI-powered energy system simulation framework that optimizes generation portfolios using genetic algorithms. The tool models hourly power dispatch, battery management, and source failures over multi-year horizons to evaluate reliability and economic performance of diverse energy mixes.

## Features

- **AI Optimization**: Uses genetic algorithms to find optimal generation mixes
- **Hourly Simulation**: Models power dispatch on an hour-by-hour basis over 12 years 
- **Multi-Source Integration**: Handles conventional generators, renewable energy, and battery storage
- **Reliability Analysis**: Tracks power outages, load shedding, and critical load interruptions
- **Economic Modeling**: Calculates comprehensive costs including capital, operational, and maintenance expenses
- **Battery Management**: Simulates advanced charging/discharging strategies
- **Failure Simulation**: Models equipment failures and maintenance downtime

## Components

### Project Class
Manages load profiles and forecasts, processes data from Excel files, and scales energy consumption patterns for future years.

### Source Class
Represents different power generation sources with configurable parameters:
- Source type (conventional, renewable, battery storage)
- Capacity and rating
- Financial model (PPA or captive)
- Operational constraints
- Failure rates and maintenance requirements

### Scenario Class
Core simulation engine that:
- Dispatches power sources according to priority
- Manages spinning reserves
- Handles sudden power drops and equipment failures
- Optimizes battery storage utilization
- Calculates key performance indicators

### Simulator
Implements genetic algorithm optimization to explore different generation portfolios and identify optimal configurations.

## How to Use

1. Prepare input data files:
   - Load profiles in Excel format
   - Source definitions with technical and economic parameters
   - Project configuration settings

2. Configure simulation parameters:
   ```python
   sc = Scenario(
       name="Baseline",
       client_name="Client",
       selected_sources=src_list,
       spin_reserve_perc=20,
       bess_non_emergency_use=2,
       bess_charge_hours=1,
       bess_priority_wise_use=True,
       charge_ratio_night=30
   )
   ```

3. Run the simulation:
   ```python
   sc.simulate()
   ```

4. Analyze results:
   ```python
   sc.aggregate_data_for_reporting()
   sc.write_yearly_data_to_csv2(output_filepath)
   print(sc.scenario_kpis)
   ```

## Example Output

The simulation generates comprehensive reports including:
- Yearly summary of system performance
- Hourly operational logs
- Key performance indicators:
  - Average Unit Cost ($/kWh)
  - Energy Fulfillment Ratio (%)
  - Critical Load Interruptions
  - Estimated Interruption Loss (M $)
  - Non-critical Load shedding events

## Requirements
- Python 3.7+
- pandas
- openpyxl
- matplotlib (for visualization)
