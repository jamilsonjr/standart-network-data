#%%
import pandas as pd
import os
import classes
import re
# this part goes do main.py
# Excel file location
file_loacation = 'data/raw/Data_Example.xlsx'
# create one dataframe for the Information sheet information in the excel file
raw_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Information')
# create one dataframe for the General_Information sheet information in the excel file
raw_general_information_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='General_Information')
# create one dataframe for the Network_info sheet information in the excel file
raw_network_info_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Network_Info')
# create one dataframe for the Peers_info sheet information in the excel file
raw_peers_info_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Peers_Info')
# create one dataframe for the Load sheet information in the excel file
raw_load_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Load')
# create one dataframe for the Generator sheet information in the excel file
raw_generator_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Generator')
# create one dataframe for the Storage sheet information in the excel file
raw_storage_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Storage')
# create one dataframe for the Vehicle sheet information in the excel file
raw_vehicle_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='Vehicle')
# create one dataframe for the CStation sheet information in the excel file
raw_cstation_sheet = pd.read_excel(os.getcwd() + '\\' + file_loacation, sheet_name='CStation')

##############################################################################
#################### Extractor Functions: Peers ##############################
##############################################################################
# Function that takes as input raw_network_info_sheet and outputs a list of Peer class objects initialized with the data from raw_network_info_sheet.
class Extractor: 
    def create_peers_list(self, df):
        # create n empty list to store Peer objects
        peers = []
        peers_static_info_df, peers_dynamic_info = self.create_peers_info(df)
        for i in peers_static_info_df.index.values:
            new_peer = classes.Peer(id=peers_static_info_df.loc[i]['peer_id'],
                            type_of_peer=peers_static_info_df.loc[i]['type_of_peer'],
                            type_of_contract=peers_static_info_df.loc[i]['type_of_contract'],
                            )
            new_peer.set_profile(buy_price_mu=peers_dynamic_info['buy_price_mu'].loc[i],
                                sell_price_mu=peers_dynamic_info['sell_price_mu'].loc[i])
            peers.append(new_peer)
        return peers
    def create_vehicles_list(self, df):
        # Create empty list to store Vehicle objects.
        vehicles = []
        vehicles_static_info_df, vehicles_dynamic_info = self.create_vehicles_info(df)
        for i in vehicles_static_info_df.index.values:
            new_vehicle = classes.Vehicle(id=vehicles_static_info_df.loc[i]['vehicle_id'],
                                          owner=vehicles_static_info_df.loc[i]['owner'],
                                          manager=vehicles_static_info_df.loc[i]['manager'],
                                          type_of_vehicle=vehicles_static_info_df.loc[i]['type_of_vehicle'],
                                          type_of_contract=vehicles_static_info_df.loc[i]['type_of_contract'],
                                          energy_capacity_max_kWh=vehicles_static_info_df.loc[i]['e_capacity_max_kwh'],
                                          p_charge_max_kW=vehicles_static_info_df.loc[i]['p_charge_max_kw'],
                                          p_discharge_max_kW=vehicles_static_info_df.loc[i]['p_discharge_max_kw'],
                                          charge_efficiency_percent=vehicles_static_info_df.loc[i]['charge_efficiency_percent'],
                                          discharge_efficiency_percent=vehicles_static_info_df.loc[i]['discharge_efficiency_percent'],
                                          initial_state_SOC_percent=vehicles_static_info_df.loc[i]['initial_state_soc_percent'],
                                          minimun_technical_SOC_percent=vehicles_static_info_df.loc[i]['minimum_technical_soc_percent']
                                          )
            new_vehicle.set_profile(
                arrive_time_period = vehicles_dynamic_info['arrive_time_period'].iloc[i],
                departure_time_period = vehicles_dynamic_info['departure_time_period'].iloc[i],
                place = vehicles_dynamic_info['place'].iloc[i],
                used_soc_percent_arriving = vehicles_dynamic_info['used_soc_percent_arriving'].iloc[i],
                soc_percent_arriving = vehicles_dynamic_info['soc_percent_arriving'].iloc[i],
                soc_required_percent_exit = vehicles_dynamic_info['soc_required_percent_exit'].iloc[i],
                p_charge_max_constracted_kw = vehicles_dynamic_info['p_charge_max_constracted_kw'].iloc[i],
                p_discharge_max_constracted_kw = vehicles_dynamic_info['p_discharge_max_constracted_kw'].iloc[i],
                charge_price = vehicles_dynamic_info['charge_price'].iloc[i],
                discharge_price = vehicles_dynamic_info['discharge_price'].iloc[i]
                )
            vehicles.append(new_vehicle)
        return vehicles 
    ############################# Aux Fucntions ###################################
    # Extract useful info from the raw_peer_info_sheet into a new dataframe.
    def create_peers_info(self, df):
        # Static Data
        # create a pandas series to store the id of the peer, and only keep integer values
        peer_id = df['Peer ID'].filter(regex='^\d+$').dropna().drop(1).reset_index(drop=True).rename('peer_id')
        # for each value value of peer id, tranforme into a str and concant 'peer_' + str(value)
        peer_id = peer_id.apply(lambda x: 'peer_' + str(x))
        # create a pandas series to store the type of peer.
        # find the index of the all the values equal to 'Type of Peer' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'
        type_of_peer = df[df['Unnamed: 2'].str.contains('Type of Peer').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'], axis=1).rename(columns={'Unnamed: 3': 'type_of_peer'})
        # find the index of the all the values equal to 'Type of Contract' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'
        type_of_contract = df[df['Unnamed: 2'].str.contains('Type of Contract').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'type_of_contract'})
        # Concat peer_id, type_of_peer and type_of_contract into a new dataframe.
        peers_static_info_df = pd.concat([peer_id, type_of_peer, type_of_contract], axis=1)
        
        # Profiles
        # Get the index of the column 'Total Time (h)' in the list of column names
        total_time_index = df.columns.get_loc('Total Time (h)')
        # Get the indexes of all the columns after 'Total Time (h)'.
        # This will be used to get the values of the columns that we want to keep
        # and store themss in a list.
        columns_to_keep_profiles = df.columns[total_time_index+1:].tolist()
        # Find the where the values equal to 'Buy Price (m.u.)' in the column 'Total Time (h)' and get the all rows, and all the columns from time index until the end.
        buy_price_mu = df[df['Total Time (h)'].apply(lambda x: x == 'Buy Price (m.u.)')].iloc[:][columns_to_keep_profiles].reset_index(drop=True)
        # Find the where the values equal to 'Sell Price (m.u.)' in the column 'Total Time (h)' and get the all rows, and all the columns from time index until the end.
        sell_price_mu = df[df['Total Time (h)'].apply(lambda x: x == 'Sell Price (m.u.)')].iloc[:][columns_to_keep_profiles].reset_index(drop=True)
        # Join all the previous series into one dataframe, with the peer_id as index.
        peers_dynamic_info = {'buy_price_mu': buy_price_mu, 'sell_price_mu': sell_price_mu} 
        return peers_static_info_df, peers_dynamic_info
    # Extract useful info from the raw_vehicle_info_sheet into a new dataframe, the same way as in create_peers_info.
    def create_vehicles_info(self, df):
        # Static Data
        # create a pandas series to store the id of the vehicle, and only keep integer values
        vehicle_id = df['Electric Vehicle ID'].filter(regex='^\d+$').dropna().drop(1).reset_index(drop=True).rename('vehicle_id')
        # for each value value of vehicle id, tranforme into a str and concant 'vehicle_' + str(value)
        vehicle_id = vehicle_id.apply(lambda x: 'vehicle_' + str(x))
        # create a pandas series to store the type of vehicle.
        # find the index of the all the values equal to 'Type of Vehicle' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'
        type_of_vehicle = df[df['Unnamed: 2'].str.contains('Type of Vehicle').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'], axis=1).rename(columns={'Unnamed: 3': 'type_of_vehicle'})
        # create a pandas series to store the owner.
        # find the index of the all the values equal to 'Owner' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'
        owner = df[df['Unnamed: 2'].str.contains('Owner').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'], axis=1).rename(columns={'Unnamed: 3': 'owner'})
        # create a pandas series to store the type of contract.
        # find the index of the all the values equal to 'Type of Contract' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'
        type_of_contract = df[df['Unnamed: 2'].str.contains('Type of Contract').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'type_of_contract'})
        # find the index of the all the values equal to 'Manager' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        manager = df[df['Unnamed: 2'].str.contains('Manager').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'manager'})
        # find the index of the all the values equal to 'E Capacity Max (kWh)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        e_capacity_max_kwh = df[df['Unnamed: 2'].str.contains('Capacity Max').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'e_capacity_max_kwh'})
        # find the index of the all the values equal to 'P Charge Max (kW)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        p_charge_max_kw = df[df['Unnamed: 2'].str.contains('Charge Max').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'p_charge_max_kw'})
        # find the index of the all the values equal to 'P Discharge Max (kW)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        p_discharge_max_kw = df[df['Unnamed: 2'].str.contains('Discharge Max').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'p_discharge_max_kw'})
        # find the index of the all the values equal to 'Charge Efficiency (%)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        charge_efficiency_percent = df[df['Unnamed: 2'].str.contains('Charge Efficiency').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'charge_efficiency_percent'})
        # find the index of the all the values equal to 'Discharge Efficiency (%)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        discharge_efficiency_percent = df[df['Unnamed: 2'].str.contains('Discharge Efficiency').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'discharge_efficiency_percent'})
        # find the index of the all the values equal to 'Initial State SOC (%)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        initial_state_soc_percent = df[df['Unnamed: 2'].str.contains('Initial State SOC').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'initial_state_soc_percent'})
        # find the index of the all the values equal to 'Minimun Technical SOC (%)' in the column 'Unnamed 2'
        # and then use the index to get the value of the column 'Unnamed 3'.
        minimum_technical_soc_percent = df[df['Unnamed: 2'].str.contains('Minimun Technical SOC').fillna(False)]['Unnamed: 3'].reset_index().drop(['index'],axis=1).rename(columns={'Unnamed: 3':'minimum_technical_soc_percent'})
        # Concat vehicle_id, type_of_vehicle, manager, e_capacity_max_kwh, p_charge_max_kw, p_discharge_max_kw, charge_efficiency_percent, discharge_efficiency_percent, initial_state_soc_percent, minimum_technical_soc_percent into a new dataframe.
        vehicles_static_info_df = pd.concat([owner, vehicle_id, type_of_vehicle, type_of_contract, manager, e_capacity_max_kwh, p_charge_max_kw, p_discharge_max_kw, charge_efficiency_percent, discharge_efficiency_percent, initial_state_soc_percent, minimum_technical_soc_percent], axis=1)
        
        # Dynamic Data
        # Get the index of the column 'Total Time (h)' in the list of column names
        event_time_column = raw_vehicle_sheet.columns.get_loc('Unnamed: 5')
        columns_to_keep_profiles = raw_vehicle_sheet.columns[event_time_column+1:].tolist()
        event_time_values = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'] ==  'Event'].iloc[:][columns_to_keep_profiles].values.tolist()[0]
        # Find the where the values equal to 'Arrive time period' in the column 'Unnamed: 5' and get the all rows, and all the columns from time index until the end,
        # example: buy_price_mu = df[df['Total Time (h)'].apply(lambda x: x == 'Buy Price (m.u.)')].iloc[:][columns_to_keep_profiles].reset_index(drop=True)
        arrive_time_period = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Arrive time period').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        arrive_time_period.columns = event_time_values
        # Same as above, but for Departure time period
        departure_time_period = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Departure time period').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        departure_time_period.columns = event_time_values
        # Same as above, but for Place
        place = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Place').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        place.columns = event_time_values
        # Same as above, but for Used SOC (%) Arriving
        used_soc_percent_arriving = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Used SOC').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        used_soc_percent_arriving.columns = event_time_values
        # Same as above, but for SOC (%) Arriving
        soc_percent_arriving = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('SOC Arriving').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        soc_percent_arriving.columns = event_time_values
        # Same as above, but for SOC Required (%) Exit
        soc_required_percent_exit = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('SOC Required').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        soc_required_percent_exit.columns = event_time_values
        # Same as above, but for Pcharge Max contracted [kW]
        p_charge_max_constracted_kw = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Pcharge Max contracted').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        p_charge_max_constracted_kw.columns = event_time_values
        # Same as above, but for PDcharge Max contracted [kW]
        p_discharge_max_constracted_kw = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('PDcharge Max contracted').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        p_discharge_max_constracted_kw.columns = event_time_values
        # Same as above, but for Charge Price
        charge_price = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Charge Price').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        charge_price.columns = event_time_values
        # Same as above, but for Disharge Price
        discharge_price = raw_vehicle_sheet[raw_vehicle_sheet['Unnamed: 5'].str.contains('Disharge Price').fillna(False)][columns_to_keep_profiles].reset_index(drop=True)
        discharge_price.columns = event_time_values
        # Rename all columns of discharge_efficiency_percent with the values of the list event_time_values.
        # This will be used to create the dataframe with the values of the discharge efficiency.
        vehicles_dynamic_info = {
        'arrive_time_period': arrive_time_period,
        'departure_time_period': departure_time_period,
        'place': place,
        'used_soc_percent_arriving': used_soc_percent_arriving,
        'soc_percent_arriving': soc_percent_arriving,
        'soc_required_percent_exit': soc_required_percent_exit,
        'p_charge_max_constracted_kw': p_charge_max_constracted_kw,
        'p_discharge_max_constracted_kw': p_discharge_max_constracted_kw,
        'charge_price': charge_price,
        'discharge_price': discharge_price
        }
        return vehicles_static_info_df, vehicles_dynamic_info        
#%% Create instance of the class Extrator
extractor = Extractor()
peers = extractor.create_peers_list(raw_peers_info_sheet)
vehicles = extractor.create_vehicles_list(raw_vehicle_sheet)
#%% 
# function parameter:
df = raw_vehicle_sheet
# Get static info names from raw_generator_sheet.
# Get all values of the column 'Unnamed: 2', remove empty values, removed duplicate values, and convert to list.
static_info_names = df['Unnamed: 2'].dropna().unique().tolist()
static_info_names[0] = df.columns[0]
# Get the index of the column 'Total Time (h)' in the list of column names
# Get all values of the column 'Total Time (h)', remove empty values, removed duplicate values, and convert to list.
if static_info_names[0] == 'Electric Vehicle ID': 
    dynamic_info_names = df['Unnamed: 5'].dropna().unique().tolist()
    dynamic_info_names[:2] = []
else:    
    dynamic_info_names = df['Total Time (h)'].dropna().unique().tolist()
    dynamic_info_names[:2] = []

# Static Data
info_dict = {}
# iterate through the static_info_names and for each name create a new dataframe with info extracted from the raw information sheets.
for i, static_info_name in enumerate(static_info_names):
    if i == 0: # The the ID
        # Create the new dataframe with the extracted information.
        # Substitute space by underscore in the name of the column.
        _df_name = static_info_name.lower().replace(' ', '_') 
        info_dict[_df_name] = pd.DataFrame()
        info_dict[_df_name] = df[static_info_names[0]].filter(regex='^\d+$').dropna().drop(1).reset_index(drop=True).rename(re.sub('[^0-9a-zA-Z]+', '_', static_info_name).rstrip('_'))
    else: # The rest of the static info
        # All characters of _df_name are converted to lower case, all charaters that are not numbers or letters are removed, all white spaces are replaced by underscores, and if the last character is an underscore, it is removed.
        _df_name = re.sub('[^0-9a-zA-Z]+', '_', static_info_name).rstrip('_').lower()
        info_dict[_df_name] = pd.DataFrame()
        info_dict[_df_name] = df[df['Unnamed: 2'].str.contains(static_info_name.split('(')[0]).fillna(False)]['Unnamed: 3'].reset_index().drop(['index'], axis=1).rename(columns={'Unnamed: 3': _df_name})   
# Dynamic Data
if static_info_names[0] == 'Electric Vehicle ID':
    total_time_index = df.columns.get_loc('Unnamed: 5')
else:
    total_time_index = df.columns.get_loc('Total Time (h)')
# Get the indexes of all the columns after 'Total Time (h)'.
# This will be used to get the values of the columns that we want to keep
# and store themss in a list.
columns_to_keep_profiles = df.columns[total_time_index+1:].tolist() 
for dynamic_info_name in dynamic_info_names:
    _dynamic_info_name = dynamic_info_name.replace('.', '')
    _df_name = re.sub('[^0-9a-zA-Z]+', '_', _dynamic_info_name).rstrip('_').lower()
    info_dict[_df_name] = pd.DataFrame()
    info_dict[_df_name] = df[df.iloc[:,5].apply(lambda x: x  == dynamic_info_name)].iloc[:][columns_to_keep_profiles].reset_index(drop=True) 
# %%
