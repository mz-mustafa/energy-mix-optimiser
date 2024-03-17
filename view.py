import streamlit as st
import re
import datetime
from sources import GridSource, SolarSource, WindSource, GasGenSource, \
    HFOGenSource, TrifuelGenSource, BESSSource, DieselGenSource, PPASource, ExistingGasGenSource
from scenario_sorted import Scenario
from utilities import write_results_to_outputs, generate_excel_in_memory, has_duplicate_values
import pandas as pd


def main():
    # Title and Logo
    st.title("KPWS- Energy Economic Modeling")

    st.write(f"## Client: Pakistan Cables Limited")

    # Scenario Inputs
    st.write("## Scenario Inputs")

    st.write("### 1. Select study length in years")
    n = st.slider(label='', min_value=5, max_value=10)

    st.write("### Select source types to be included in your scenario:")

    # Source Toggles
    grid_req = st.checkbox('Grid Supply')
    ppa_req = st.checkbox('PPA Supply')
    solar_req = st.checkbox('Solar PV')
    gas_gen_req = st.checkbox('Gas Generator')
    existing_gas_gen_req = st.checkbox('Existing Gas Generators')
    hfo_gen_req = st.checkbox('HFO Generator')
    tri_fuel_gen_req = st.checkbox('Tri-fuel FO+Gas generator')
    wind_req = st.checkbox('Wind')
    bess_req = st.checkbox('BESS')
    diesel_req = st.checkbox('Diesel Generator')
    priority_list = list(range(1,9))
    grid_priority, gas_priority, existing_gas_priority, hfo_priority, tri_priority, diesel_priority = (-1,) * 6
    solar_priority, wind_priority, ppa_priority = (0,) * 3
    # Display Options based on toggles
    if grid_req:

        st.write("#### Grid Supply Parameters")
        grid_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 2,key="grid_priority")
        grid_input = { f"Year {i}": { 'Sanctioned Load Added (MW)': 0.0 }
                       for i in range(n+1) }

        grid_input_df = st.data_editor(grid_input)
        #check if all values are not zero
        if all(grid_data['Sanctioned Load Added (MW)'] == 0.0 for grid_data in grid_input_df.values()):
            st.warning("All values for 'Sanctioned Load Added (MW)' are zero. Please adjust or "
                       "remove Grid from scenario. Note that Year 0 sanctioned load is the current one.")

    if ppa_req:

        st.write("#### PPA Supply Parameters")
        #ppa_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 1,key="ppa_priority")
        ppa_input = { f"Year {i}": { 'Supply Capacity Added (MW)': 0.0 }
                       for i in range(n+1) }

        ppa_input_df = st.data_editor(ppa_input)
        #check if all values are not zero
        if all(ppa_data['Supply Capacity Added (MW)'] == 0.0 for ppa_data in ppa_input_df.values()):
            st.warning("All values for 'Supply Capacity Added (MW)' are zero. Please adjust or "
                       "remove PPA from scenario. Note that Year 0 supply capacity is the current one.")


    if solar_req:
        st.write("#### On-grid Solar PV Parameters")
        #solar_priority = st.selectbox("Source Priority to Feed Load:", priority_list, key="solar_priority")
        solar_input = {f"Year {i}": {'Solar Capacity Added (MW)': 0.0}
                       for i in range(n + 1)}

        solar_input_df = st.data_editor(solar_input, key='solar_input')
        if all(solar_data['Solar Capacity Added (MW)'] == 0.0 for solar_data in solar_input_df.values()):
            st.warning("All input values for Solar PV are zero. Please adjust or "
                       "remove Solar PV from scenario. Note that Year 0 capacity represents current investment")

    if wind_req:
        st.write("#### Wind Parameters:")
        #wind_priority = st.selectbox("Source Priority to Feed Load:", priority_list, key="wind_priority")
        wind_input = {f"Year {i}": {'2MW Turbines Added': 0}
                       for i in range(n + 1)}

        wind_input_df = st.data_editor(wind_input, key='wind_input')
        if all(wind_data['2MW Turbines Added'] == 0 for wind_data in wind_input_df.values()):
            st.warning("All input values for Wind are zero. Please adjust or "
                       "remove Wind from scenario. Note that Year 0 quantity represents current investment.")

    if gas_gen_req:
        st.write("#### Gas Generator Parameters")
        #gas_base_rating_pu = st.number_input("Base Power Rating in MW, per set:", step=0.05, key="gas_base_rating_pu")
        #gas_chp_cooling = st.checkbox('CHP Operation with Absorption Chiller(s)?', key="gas_chp_cooling")
        gas_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 0,key="gas_priority")
        gas_fuel_type = st.selectbox("Choose Fuel type:", Scenario.available_gas_types(), key="gas_fuel_type")

        gas_gen_input = {
            f"Year {year}": {
                "Qty of Primary Units": 0,
                "Rating of Primary Units": 0.0,
                "% of rated output after derating": 100,
                "Fuel Efficiency Rating": 100
            }
            for year in range(n + 1)
        }
        gas_gen_input_df = st.data_editor(gas_gen_input, key='gas_gen_input')

        if all(gas_data['Qty of Primary Units'] == 0 for gas_data in gas_gen_input_df.values()):
            st.warning("All Primary Gas Genset quantities are zero. Please adjust or "
                       "remove Gas Generators from scenario. Note that Year 0 Quantities are current investment.")

    if existing_gas_gen_req:
            st.write("#### Existing Gas Generator Parameters: 2 x 1MW")
            #gas_base_rating_pu = st.number_input("Base Power Rating in MW, per set:", step=0.05, key="gas_base_rating_pu")
            #gas_chp_cooling = st.checkbox('CHP Operation with Absorption Chiller(s)?', key="gas_chp_cooling")
            existing_gas_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 0,key="existing_gas_priority")
            existing_gas_fuel_type = st.selectbox("Choose Fuel type:", Scenario.available_gas_types(), key="existing_gas_fuel_type")

            existing_gas_gen_input = {
                f"Year {year}": {
                    #"Qty of Primary Units": 0,
                    #"Rating of Primary Units": 0.0,
                    "% of rated output after derating": 100,
                    "Fuel Efficiency Rating": 100
                }
                for year in range(n + 1)
            }
            existing_gas_gen_input_df = st.data_editor(gas_gen_input, key='existing_gas_gen_input')

            if all(existing_gas_data['Qty of Primary Units'] == 0 for existing_gas_data in existing_gas_gen_input_df.values()):
                st.warning("All Existing Gas Genset quantities are zero. Please adjust or "
                        "remove Existing Gas Generators from scenario. Note that Year 0 Quantities are current investment.")


    if hfo_gen_req:
        st.write("#### HFO Generator Parameters")
        #hfo_base_rating_pu = st.number_input("Base Power Rating in MW, per set:", step=0.05, key = "hfo_base_rating_pu")
        #hfo_chp_cooling = st.checkbox('CHP Operation with Absorption Chiller?', key="hfo_chp_cooling")
        hfo_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 0,key="hfo_priority")

        hfo_gen_input = {
            f"Year {year}": {
                "Qty of Primary Units": 0,
                "Rating of Primary Units": 0.0,
                "% of rated output after derating": 100,
                "Fuel Efficiency Rating": 100
            }
            for year in range(n + 1)
        }
        hfo_gen_input_df = st.data_editor(hfo_gen_input, key='hfo_gen_input')

        if all(hfo_data['Qty of Primary Units'] == 0 for hfo_data in hfo_gen_input_df.values()):
            st.warning("All Primary HFO Genset quantities are zero. Please adjust or "
                       "remove HFO Generators from scenario. Note that Year 0 Quantities represent current investment.")

    if tri_fuel_gen_req:
        st.write("#### Tri-Fuel Generator Parameters")
        #trifuel_base_rating_pu = st.number_input("Base Power Rating in MW, per set:", step=0.05, key="trifuel_base_rating_pu")
        #trifuel_chp_cooling = st.checkbox('CHP Operation with Absorption Chiller(s)?', key="trifuel_chp_cooling")
        tri_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 0,key="tri_priority")
        trifuel_fuel_type = st.selectbox("Choose Fuel type:", Scenario.available_gas_types(), key="trifuel_fuel_type")

        trifuel_gen_input = {
            f"Year {year}": {
                "Qty of Primary Units": 0,
                "Rating of Primary Units": 0.0,
                "% of rated output after derating": 100,
                "Fuel Efficiency Rating": 100
            }
            for year in range(n + 1)
        }
        trifuel_gen_input_df = st.data_editor(trifuel_gen_input, key='trifuel_gen_input')

        if all(tri_data['Qty of Primary Units'] == 0 for tri_data in trifuel_gen_input_df.values()):
            st.warning("All Primary Trifuel Genset quantities are zero. Please adjust or "
                       "remove Trifuel Generators from scenario. Note that Year 0 Quantities represent current investment.")

    if diesel_req:
        st.write("#### Diesel Generator Parameters")
        diesel_priority = st.selectbox("Source Priority to Feed Load:", priority_list, index = 3, key="diesel_priority")

        diesel_gen_input = {
            f"Year {year}": {
                "Qty of Primary Units": 0,
                "Rating of Primary Units": 0.0,
                "% of rated output after derating": 100,
                "Fuel Efficiency Rating": 100
            }
            for year in range(n + 1)
        }
        diesel_gen_input_df = st.data_editor(diesel_gen_input, key='diesel_gen_input')

        if all(diesel_data['Qty of Primary Units'] == 0 for diesel_data in diesel_gen_input_df.values()):
            st.warning("All Diesel Genset quantities are zero. Please adjust or "
                       "remove Diesel Generators from scenario. Note that Year 0 Quantities represent current investment.")

    if bess_req:
        st.write("#### BESS Parameters")
        bess_input = {f"Year {i}": {'0.5MWh BESS Units Added': 0}
                       for i in range(n + 1)}
        bess_input_df = st.data_editor(bess_input, key='bess_input')
        if all(bess_data['0.5MWh BESS Units Added'] == 0 for bess_data in bess_input_df.values()):
            st.warning("All input values for BESS are zero. Please adjust or "
                       "remove BESS from scenario. Note that Year 0 quantity represents current investment.")


    st.write("#### Give your scenario a name")
    sc_name = st.text_input('Scenario Name',placeholder='e.g. Grid, Solar, Gas Gen 5Y', label_visibility='collapsed')

    submit_button = st.button('Submit')

    # add the created sources to scenario
    if submit_button:
        print('Submit Button Pressed')
        p_list = [grid_priority, ppa_priority, gas_priority, existing_gas_priority, hfo_priority, tri_priority, diesel_priority]
        if has_duplicate_values(p_list):
            st.warning("Source priorities are not defined. You may have assigned the same priority to multiple" 
                       "sources. Please correst and submit the scenario again")
        else:
            sc = Scenario(sc_name, 'Pakistan Cables Limited', n=n)
            sc.timestamp = datetime.datetime.now()
            print("Scenario created")

            if grid_req and grid_priority > 0:
                grid = GridSource(n,grid_priority)
                for year, grid_data in grid_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    grid.inputs[y]['rating_prim_units'] = grid_data['Sanctioned Load Added (MW)']
                    grid.inputs[y]['count_prim_units'] = 1 if grid_data['Sanctioned Load Added (MW)'] != 0 else 0

                sc.add_source(grid)
                print("Grid source added to scenario")

            if ppa_req:
                ppa = PPASource(n, ppa_priority)
                for year, ppa_data in ppa_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    ppa.inputs[y]['rating_prim_units'] = ppa_data['Supply Capacity Added (MW)']
                    ppa.inputs[y]['count_prim_units'] = 1 if ppa_data['Supply Capacity Added (MW)'] != 0 else 0

                sc.add_source(ppa)
                print("ppa source added to scenario")

            if solar_req:
                solar = SolarSource(n, solar_priority)
                for year, solar_data in solar_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    solar.inputs[y]['rating_prim_units'] = solar_data['Solar Capacity Added (MW)']
                    solar.inputs[y]['count_prim_units'] = 1 if solar_data['Solar Capacity Added (MW)'] != 0 else 0

                sc.add_source(solar)
                print("solar source added to scenario")

            if wind_req:
                wind = WindSource(n, wind_priority)
                for year, wind_data in wind_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    wind.inputs[y]['rating_prim_units'] = 2
                    wind.inputs[y]['count_prim_units'] = wind_data['2MW Turbines Added']

                sc.add_source(wind)
                print("wind source added to scenario")

            if bess_req:
                bess = BESSSource(n, 0)
                for year, bess_data in bess_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    bess.inputs[y]['rating_prim_units'] = 0.5
                    bess.inputs[y]['count_prim_units'] = bess_data['0.5MWh BESS Units Added']
                print(bess.inputs[4]['count_prim_units'])
                sc.add_source(bess)
                print("bess source added to scenario")

            if gas_gen_req and gas_priority > 0:
                gas_gen = GasGenSource(n, gas_priority)
                gas_gen.inputs['chp_operation'] = True
                gas_gen.inputs['gas_fuel_type'] = gas_fuel_type
                for year, gas_data in gas_gen_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    gas_gen.inputs[y]['count_prim_units'] = gas_data['Qty of Primary Units']
                    gas_gen.inputs[y]['rating_prim_units'] = gas_data['Rating of Primary Units']
                    gas_gen.inputs[y]['perc_rated_output'] = gas_data['% of rated output after derating']
                    gas_gen.inputs[y]['fuel_eff'] = gas_data["Fuel Efficiency Rating"]

                sc.add_source(gas_gen)
                print("gas source added to scenario")

            if existing_gas_gen_req and existing_gas_priority > 0:
                existing_gas_gen = ExistingGasGenSource(n, existing_gas_priority)
                existing_gas_gen.inputs['chp_operation'] = True
                existing_gas_gen.inputs['gas_fuel_type'] = existing_gas_fuel_type
                for year, existing_gas_data in existing_gas_gen_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    existing_gas_gen.inputs[y]['count_prim_units'] = 2
                    existing_gas_gen.inputs[y]['rating_prim_units'] = 1
                    existing_gas_gen.inputs[y]['perc_rated_output'] = existing_gas_data['% of rated output after derating']
                    existing_gas_gen.inputs[y]['fuel_eff'] = existing_gas_data["Fuel Efficiency Rating"]

                sc.add_source(existing_gas_gen)
                print("existing gas source added to scenario")

            if hfo_gen_req and hfo_priority > 0:
                hfo_gen = HFOGenSource(n, hfo_priority)
                hfo_gen.inputs['chp_operation'] = True
                for year, hfo_data in hfo_gen_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    hfo_gen.inputs[y]['count_prim_units'] = hfo_data['Qty of Primary Units']
                    hfo_gen.inputs[y]['rating_prim_units'] = hfo_data['Rating of Primary Units']
                    hfo_gen.inputs[y]['perc_rated_output'] = hfo_data['% of rated output after derating']
                    hfo_gen.inputs[y]['fuel_eff'] = hfo_data["Fuel Efficiency Rating"]

                sc.add_source(hfo_gen)
                print("hfo source added to scenario")

            if tri_fuel_gen_req and tri_priority > 0:
                trifuel_gen = TrifuelGenSource(n, tri_priority)
                trifuel_gen.inputs['chp_operation'] = True
                trifuel_gen.inputs['gas_fuel_type'] = trifuel_fuel_type
                for year, tri_data in trifuel_gen_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    trifuel_gen.inputs[y]['count_prim_units'] = tri_data['Qty of Primary Units']
                    trifuel_gen.inputs[y]['rating_prim_units'] = tri_data['Rating of Primary Units']
                    trifuel_gen.inputs[y]['perc_rated_output'] = tri_data['% of rated output after derating']
                    trifuel_gen.inputs[y]['fuel_eff'] = tri_data["Fuel Efficiency Rating"]

                sc.add_source(trifuel_gen)
                print("trifuel source added to scenario")

            if diesel_req and diesel_priority > 0:
                diesel_gen = DieselGenSource(n, diesel_priority)
                for year, diesel_data in diesel_gen_input_df.items():
                    y = int(re.search(r'\d+', year).group())
                    diesel_gen.inputs[y]['count_prim_units'] = diesel_data['Qty of Primary Units']
                    diesel_gen.inputs[y]['rating_prim_units'] = diesel_data['Rating of Primary Units']
                    diesel_gen.inputs[y]['perc_rated_output'] = diesel_data['% of rated output after derating']
                    diesel_gen.inputs[y]['fuel_eff'] = diesel_data["Fuel Efficiency Rating"]
                
                sc.add_source(diesel_gen)
                print("diesel source added to scenario")

            sc.generate_results()
            sc.generate_summaries()

            st.write("#### Summary Outcomes")
            st.dataframe(sc.summary_df.round(1), hide_index=True)

            st.write("#### Power Summary")
            st.write("###### Shown cases either have unserved demand or are randomly picked for each year")
            st.dataframe(sc.power_summary_df.round(1), hide_index=True)

            st.write("#### Energy Summary")
            st.dataframe(sc.energy_summary_df.round(1), hide_index=True)

            st.write("#### CAPEX Summary")
            st.dataframe(sc.capex_df.round(1), hide_index=True)

            st.write("#### OPEX Summary")
            st.dataframe(sc.opex_summary_df.round(1), hide_index=True)

            st.write("#### Emissions Summary")
            st.dataframe(sc.emissions_summary_df.round(1), hide_index=True)


            print("Now creating in memory excel file for download")
            excel_file = generate_excel_in_memory(sc)
            st.download_button(
                label="Download Excel File",
                data=excel_file,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

def create_kpi_card(column, title, value):
    with column:
        st.markdown(f"## {title}")
        st.markdown(f"<h2 style='color: red;'>{value}</h2>", unsafe_allow_html=True)

main()
