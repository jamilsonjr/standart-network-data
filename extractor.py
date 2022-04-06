#%%
import pandas as pd
import os
import classes
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
#
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
            new_peer.set_buy_sell_prices(buy_price_mu=peers_dynamic_info['buy_price_mu'].loc[i],
                                         sell_price_mu=peers_dynamic_info['sell_price_mu'].loc[i])
            peers.append(new_peer)
        return peers
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

# Create instance of the class Extrator
extractor = Extractor()
#%%
peers = extractor.create_peers_list(raw_peers_info_sheet)
#%%
peers_static_info_df, peers_dynamic_info = extractor.create_peers_info(raw_peers_info_sheet)
