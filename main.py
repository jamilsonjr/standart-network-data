# python main boiler plate code
import pandas as pd
import os
import extractor

def main():
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
    

if __name__ == "__main__":
    main()