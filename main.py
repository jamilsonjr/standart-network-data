#%%
import pandas as pd
import os
from classes import Network
from extractor import Extractor
def main():
    #    # Excel file location
    file_location = 'data/raw/Data_Example.xlsx'
    # create one dataframe for the Information sheet information in the excel file
    raw_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Information')
    # create one dataframe for the General_Information sheet information in the excel file
    raw_general_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='General_Information')
    # create one dataframe for the Peers_info sheet information in the excel file
    raw_peers_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Peers_Info')
    # create one dataframe for the Load sheet information in the excel file
    raw_load_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Load')
    # create one dataframe for the Generator sheet information in the excel file
    raw_generator_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Generator')
    # create one dataframe for the Storage sheet information in the excel file
    raw_storage_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Storage')
    # create one dataframe for the Vehicle sheet information in the excel file
    raw_vehicle_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Vehicle')
    # create one dataframe for the CStation sheet information in the excel file
    raw_charging_station_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='CStation')
    # create one dataframe for the Network_Info sheet information in the excel file
    raw_network_info_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Network_Info')
    # create one dataframe for the General_Information sheet information in the excel file
    raw_general_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='General_Information')
    # Extract useful info from the raw dataframes.
    extractor = Extractor()
    # Create network object
    network = Network(
        simulation_periods = extractor.create_simulation_periods(raw_general_information_sheet),
        periods_duration_min = extractor.create_periods_duration_min(raw_general_information_sheet),
        objective_functions_list = extractor.create_objective_functions_list(raw_general_information_sheet),
        info = extractor.create_network_info(raw_network_info_sheet),
        vehicle_list = extractor.create_vehicles_list(raw_vehicle_sheet),
        load_list = extractor.create_loads_list(raw_load_sheet),
        generator_list = extractor.create_generators_list(raw_generator_sheet),
        storage_list = extractor.create_storages_list(raw_storage_sheet),
        charging_station_list = extractor.create_charging_stations_list(raw_charging_station_sheet),
        peer_list = extractor.create_peers_list(raw_peers_sheet)
    )
    # Create the pandapower network.
    network.create_pandapower_model() # Property name: net_model.
    # Plot the network.
    network.plot_network()    
if __name__ == "__main__":
    main()
# TODO: Merge info sheet to general info shee
# TODO: Get profiles and  create ground info.
# TODO: Decide which problems to solve first, regression? Which techniques to use?