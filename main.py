#%%
import pandas as pd
import os
from classes import Network
from extractor import Extractor
import pandapower as pp
# def main():
#     # Excel file location
#     file_location = 'data/raw/Data_Example.xlsx'
#     # create one dataframe for the Information sheet information in the excel file
#     raw_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Information')
#     # create one dataframe for the General_Information sheet information in the excel file
#     raw_general_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='General_Information')
#     # create one dataframe for the Network_info sheet information in the excel file
#     raw_network_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Network_Info')
#     # create one dataframe for the Peers_info sheet information in the excel file
#     raw_peers_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Peers_Info')
#     # create one dataframe for the Load sheet information in the excel file
#     raw_load_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Load')
#     # create one dataframe for the Generator sheet information in the excel file
#     raw_generator_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Generator')
#     # create one dataframe for the Storage sheet information in the excel file
#     raw_storage_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Storage')
#     # create one dataframe for the Vehicle sheet information in the excel file
#     raw_vehicle_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Vehicle')
#     # create one dataframe for the CStation sheet information in the excel file
#     raw_charging_station_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='CStation')
#     # create one dataframe for the Network_Info sheet information in the excel file
#     raw_network_info_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Network_Info')
#     # Extract useful info from the raw dataframes.
#     extractor = Extractor()
#     peers = extractor.create_peers_list(raw_peers_sheet)
#     vehicles = extractor.create_vehicles_list(raw_vehicle_sheet)
#     charging_staions = extractor.create_charging_stations_list(raw_charging_station_sheet)
#     generators = extractor.create_generators_list(raw_generator_sheet)
#     storages = extractor.create_storages_list(raw_storage_sheet)
#     loads = extractor.create_loads_list(raw_load_sheet)
#     network_info = extractor.create_network_info_dict(raw_network_info_sheet)
#     # Create the pandapower network.
    
# if __name__ == "__main__":
#     main()
#    # Excel file location
file_location = 'data/raw/Data_Example.xlsx'
# create one dataframe for the Information sheet information in the excel file
raw_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Information')
# create one dataframe for the General_Information sheet information in the excel file
raw_general_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='General_Information')
# create one dataframe for the Network_info sheet information in the excel file
raw_network_sheet = pd.read_excel(os.getcwd() + '\\' + file_location, sheet_name='Network_Info')
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
# %% Doing now
# create pandapower network
# Create empty network.
net = pp.create_empty_network()
# Create the busses.
branch_info = network.info.branch_info
for bus_num in branch_info.component_n.values:
    pp.create_bus(net, vn_kv=20.0, index=bus_num)
# External grid.
pp.create_ext_grid(net, bus=1, vm_pu=1.0)
# Create tge loads from network.load_list.
for i, load in enumerate(network.load_list):
    name = load.id
    bus = load.internal_bus_location
    p_mw = load.power_contracted_kw / 1000
    q_mvar  = p_mw * load.tg_phi
    pp.create_load(net, bus=bus, p_mw=p_mw, q_mvar=q_mvar, name=name)
# Create generators from network.generator_list.
for i, generator in enumerate(network.generator_list):
    bus = generator.internal_bus_location
    p_mw = generator.p_max_kw / 1000
    q_mvar = generator.q_max_kvar / 1000
    pp.create_sgen(net, bus=bus, p_mw=p_mw, q_mvar=q_mvar, name=name)    
# Create the storage from network.storage_list.
for i, storage in enumerate(network.storage_list):
    bus = storage.internal_bus_location
    p_mw = storage.p_charge_max_kw / 1000
    max_e_mwh = storage.energy_capacity_kvah / 1000
    pp.create_storage(net, bus=bus, p_mw=p_mw, max_e_mwh=max_e_mwh, name=name)
#%%
# Merge info sheet to general info shee