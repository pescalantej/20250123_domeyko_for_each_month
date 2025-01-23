# %%
# LIBRARIES AND MODULES
import datetime
from pathlib import Path

# %% 
# PERIOD
START_DATE = "01-12-2024"
END_DATE = "01-01-2025"

# %%
# PARK INFO
PARK = "DOM"
N_CABINS = 22
N_INVERTERS_PER_CABIN = 4
N_INVERTERS = N_CABINS*N_INVERTERS_PER_CABIN

# %%
# EXTENTIONS
EXTENTIONS = ["xls", "csv"]

# %%
# GENERAL PROPOUSE
DATE_OBJECT = datetime.datetime.strptime(START_DATE, "%d-%m-%Y")
FORMATED_DATE = DATE_OBJECT.strftime("%m_%Y")
OUTPUT_FORMAT_AND_EXTENTION = f"_{FORMATED_DATE}_{PARK}.{EXTENTIONS[0]}"

# INPUT AGGREGATION PERIODS
INPUT_AGG_PERIOD = 1
INPUT_METERS_AGG_PERIOD = 15

# OUTPUT AGGREGATION PERIODS
OUTPUT_AGG_PERIOD_1M = 1
OUTPUT_AGG_PERIOD_15M = 15
OUTPUT_AGG_PERIOD_1H = 60
OUTPUT_AGG_PERIOD_1D = 1440

# %%
# FOLDERS
# RAW DATA FOLDERS
ROOT_PATH = Path(f"{DATE_OBJECT.year}_{DATE_OBJECT.month:02}")
INPUT_FOLDER_PATH_RAW = ROOT_PATH / "01_raw_data"
INPUT_FOLDER_PATH_INVERTERS = INPUT_FOLDER_PATH_RAW / "01_inverters"
INPUT_FOLDER_PATH_SENSORS_METEO = INPUT_FOLDER_PATH_RAW / "02_sensors"
INPUT_FOLDER_PATH_GENERACION = INPUT_FOLDER_PATH_RAW / "03_generacion"
INPUT_FOLDER_PATH_METERS = INPUT_FOLDER_PATH_RAW / "04_meters"
INPUT_FOLDER_PATH_PRMTE = INPUT_FOLDER_PATH_RAW / "05_prmte"
INPUT_FOLDER_SOLAR_GIS = INPUT_FOLDER_PATH_RAW / "06_solar_gis"

# PROCESSED DATA FOLDER
OUTPUT_FOLDER_PROCESSED_DATA = ROOT_PATH / "02_processed_data"

# PROCESSED FILE NAMES
OUTPUT_FILE_NAMES_INVERTERS_PRODUCTION = [
    "01_psn_inverters_production.xlsx",
]

OUTPUT_FILE_NAMES_SENSORS = [
    "02_psn_sensors.xlsx"
]

OUTPUT_FILE_NAME_METERS =[
    "04_psn_meters.xlsx"
]

OUTPUT_FILE_NAME_PRMTE = [
    "05_psn_prmte.xlsx"
]

# %% 
# ALL INVERTERS
INVERTERS_KW_SCADA_TO_TAG = {
    'PN1_S11_AN10028': 'Cabin 1 inverter 1 [kW]',
    'PN1_S11_AN20028': 'Cabin 1 inverter 2 [kW]',
    'PN1_S11_AN30028': 'Cabin 1 inverter 3 [kW]',
    'PN1_S11_AN40028': 'Cabin 1 inverter 4 [kW]',
    'PN1_S12_AN10028': 'Cabin 2 inverter 1 [kW]',
    'PN1_S12_AN20028': 'Cabin 2 inverter 2 [kW]',
    'PN1_S12_AN30028': 'Cabin 2 inverter 3 [kW]',
    'PN1_S12_AN40028': 'Cabin 2 inverter 4 [kW]',
    'PN1_S13_AN10028': 'Cabin 3 inverter 1 [kW]',
    'PN1_S13_AN20028': 'Cabin 3 inverter 2 [kW]',
    'PN1_S13_AN30028': 'Cabin 3 inverter 3 [kW]',
    'PN1_S13_AN40028': 'Cabin 3 inverter 4 [kW]',
    'PN1_S14_AN10028': 'Cabin 4 inverter 1 [kW]',
    'PN1_S14_AN20028': 'Cabin 4 inverter 2 [kW]',
    'PN1_S14_AN30028': 'Cabin 4 inverter 3 [kW]',
    'PN1_S14_AN40028': 'Cabin 4 inverter 4 [kW]',
    'PN1_S15_AN10028': 'Cabin 5 inverter 1 [kW]',
    'PN1_S15_AN20028': 'Cabin 5 inverter 2 [kW]',
    'PN1_S15_AN30028': 'Cabin 5 inverter 3 [kW]',
    'PN1_S15_AN40028': 'Cabin 5 inverter 4 [kW]',
    'PN1_S21_AN10028': 'Cabin 6 inverter 1 [kW]',
    'PN1_S21_AN20028': 'Cabin 6 inverter 2 [kW]',
    'PN1_S21_AN30028': 'Cabin 6 inverter 3 [kW]',
    'PN1_S21_AN40028': 'Cabin 6 inverter 4 [kW]',
    'PN1_S22_AN10028': 'Cabin 7 inverter 1 [kW]',
    'PN1_S22_AN20028': 'Cabin 7 inverter 2 [kW]',
    'PN1_S22_AN30028': 'Cabin 7 inverter 3 [kW]',
    'PN1_S22_AN40028': 'Cabin 7 inverter 4 [kW]',
    'PN1_S23_AN10028': 'Cabin 8 inverter 1 [kW]',
    'PN1_S23_AN20028': 'Cabin 8 inverter 2 [kW]',
    'PN1_S23_AN30028': 'Cabin 8 inverter 3 [kW]',
    'PN1_S23_AN40028': 'Cabin 8 inverter 4 [kW]',
    'PN1_S24_AN10028': 'Cabin 9 inverter 1 [kW]',
    'PN1_S24_AN20028': 'Cabin 9 inverter 2 [kW]',
    'PN1_S24_AN30028': 'Cabin 9 inverter 3 [kW]',
    'PN1_S24_AN40028': 'Cabin 9 inverter 4 [kW]',
    'PN1_S33_AN10028': 'Cabin 10 inverter 1 [kW]',
    'PN1_S33_AN20028': 'Cabin 10 inverter 2 [kW]',
    'PN1_S33_AN30028': 'Cabin 10 inverter 3 [kW]',
    'PN1_S33_AN40028': 'Cabin 10 inverter 4 [kW]',
    'PN1_S34_AN10028': 'Cabin 11 inverter 1 [kW]',
    'PN1_S34_AN20028': 'Cabin 11 inverter 2 [kW]',
    'PN1_S34_AN30028': 'Cabin 11 inverter 3 [kW]',
    'PN1_S34_AN40028': 'Cabin 11 inverter 4 [kW]',
    'PN1_S41_AN10028': 'Cabin 12 inverter 1 [kW]',
    'PN1_S41_AN20028': 'Cabin 12 inverter 2 [kW]',
    'PN1_S41_AN30028': 'Cabin 12 inverter 3 [kW]',
    'PN1_S41_AN40028': 'Cabin 12 inverter 4 [kW]',
    'PN1_S42_AN10028': 'Cabin 13 inverter 1 [kW]',
    'PN1_S42_AN20028': 'Cabin 13 inverter 2 [kW]',
    'PN1_S42_AN30028': 'Cabin 13 inverter 3 [kW]',
    'PN1_S42_AN40028': 'Cabin 13 inverter 4 [kW]',
    'PN1_S43_AN10028': 'Cabin 14 inverter 1 [kW]',
    'PN1_S43_AN20028': 'Cabin 14 inverter 2 [kW]',
    'PN1_S43_AN30028': 'Cabin 14 inverter 3 [kW]',
    'PN1_S43_AN40028': 'Cabin 14 inverter 4 [kW]',
    'PN1_S44_AN10028': 'Cabin 15 inverter 1 [kW]',
    'PN1_S44_AN20028': 'Cabin 15 inverter 2 [kW]',
    'PN1_S44_AN30028': 'Cabin 15 inverter 3 [kW]',
    'PN1_S44_AN40028': 'Cabin 15 inverter 4 [kW]',
    'PN1_S45_AN10028': 'Cabin 16 inverter 1 [kW]',
    'PN1_S45_AN20028': 'Cabin 16 inverter 2 [kW]',
    'PN1_S45_AN30028': 'Cabin 16 inverter 3 [kW]',
    'PN1_S45_AN40028': 'Cabin 16 inverter 4 [kW]',
    'PN1_S51_AN10028': 'Cabin 17 inverter 1 [kW]',
    'PN1_S51_AN20028': 'Cabin 17 inverter 2 [kW]',
    'PN1_S51_AN30028': 'Cabin 17 inverter 3 [kW]',
    'PN1_S51_AN40028': 'Cabin 17 inverter 4 [kW]',
    'PN1_S52_AN10028': 'Cabin 18 inverter 1 [kW]',
    'PN1_S52_AN20028': 'Cabin 18 inverter 2 [kW]',
    'PN1_S52_AN30028': 'Cabin 18 inverter 3 [kW]',
    'PN1_S52_AN40028': 'Cabin 18 inverter 4 [kW]',
    'PN1_S53_AN10028': 'Cabin 19 inverter 1 [kW]',
    'PN1_S53_AN20028': 'Cabin 19 inverter 2 [kW]',
    'PN1_S53_AN30028': 'Cabin 19 inverter 3 [kW]',
    'PN1_S53_AN40028': 'Cabin 19 inverter 4 [kW]',
    'PN1_S54_AN10028': 'Cabin 20 inverter 1 [kW]',
    'PN1_S54_AN20028': 'Cabin 20 inverter 2 [kW]',
    'PN1_S54_AN30028': 'Cabin 20 inverter 3 [kW]',
    'PN1_S54_AN40028': 'Cabin 20 inverter 4 [kW]',
    'PN1_S31_AN10028': 'Cabin 21 inverter 1 [kW]',
    'PN1_S31_AN20028': 'Cabin 21 inverter 2 [kW]',
    'PN1_S31_AN30028': 'Cabin 21 inverter 3 [kW]',
    'PN1_S31_AN40028': 'Cabin 21 inverter 4 [kW]',
    'PN1_S32_AN10028': 'Cabin 22 inverter 1 [kW]',
    'PN1_S32_AN20028': 'Cabin 22 inverter 2 [kW]',
    'PN1_S32_AN30028': 'Cabin 22 inverter 3 [kW]',
    'PN1_S32_AN40028': 'Cabin 22 inverter 4 [kW]'
}
    
INVERTERS_OPERATIONS_1M_TO_15M = {
    value: 'mean' for value in list(INVERTERS_KW_SCADA_TO_TAG.values())
}

INVERTERS_KW_TO_KWH = {
    value: value[:-1] + 'h]' for value in list(INVERTERS_KW_SCADA_TO_TAG.values())
}

INVERTERS_OPERATIONS_15M_TO_1H_1D = {
    value: 'sum' for value in list(INVERTERS_KW_TO_KWH.values())
}

INVERTERS_MWH_SCADA_TO_TAG_ = {
    # to be completed
}

# %%
# METEO
METEO_SCADA_TO_TAG = {
    'PN1_S00_AN00005':'Pyranometer H. 01 (WS) cabin 01 [W/m2]',
    'PN1_S00_AN00006':'Pyranometer H. 02 (WS) cabin 01 [W/m2]',
    'PN1_S00_AN00007':'Pyranometer H. 03 (WS) cabin 01 [W/m2]',
    'PN1_S11_AN00001':'Pyranometer POA cabin 01 [W/m2]',
    'PN1_S14_AN00001':'Pyranometer POA cabin 04 [W/m2]',
    'PN1_S22_AN00001':'Pyranometer POA cabin 07 [W/m2]',
    'PN1_S33_AN00001':'Pyranometer POA cabin 10 [W/m2]',
    'PN1_S41_AN00001':'Pyranometer POA cabin 12 [W/m2]',
    'PN1_S44_AN00001':'Pyranometer POA cabin 15 [W/m2]',
    'PN1_S52_AN00001':'Pyranometer POA cabin 18 [W/m2]',
    'PN1_S31_AN00001':'Pyranometer POA cabin 21 [W/m2]',
    'PN1_S00_AN00008':'Pyranometer difuse (WS) cabin 01 [W/m2]',
    'PN1_S00_AN00003':'Ambient temp. (WS) cabin 01 [°C]',
    'PN1_S14_AN00002':'Module temp. cabin 04 [°C]',
    'PN1_S22_AN00002':'Module temp. cabin 07 [°C]',
    'PN1_S33_AN00002':'Module temp. cabin 10 [°C]',
    'PN1_S44_AN00002':'Module temp. cabin 15 [°C]',
    'PN1_S52_AN00002':'Module temp. cabin 18 [°C]'
}

METEO_TAG_W_TO_TAG_WH = {
    value: value.replace("[W/m2]", "[Wh/m2]") if "[W/m2]" in value else value for value in METEO_SCADA_TO_TAG.values()
}

METEO_OPERATIONS_1M_TO_15M = {
    value: 'mean' for value in METEO_SCADA_TO_TAG.values()
}

METEO_OPERATIONS_15M_TO_1H_1D = {
    value: "sum" if "[Wh/m2]" in value else "mean" for value in METEO_TAG_W_TO_TAG_WH.values()
}

# %%
# PRMTE
PRMTE_INDEXES_TO_REMOVE = [3, 6, 7, 8, 9, 10, 11, 12, 13]