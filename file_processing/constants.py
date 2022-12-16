import os
import re
import json

parent_directory = os.path.dirname(os.path.abspath(__file__))
parent_directory = re.sub('\\\\file_processing', '', parent_directory)
SETTINGS_PATH = os.path.join(parent_directory, 'data', 'settings.json')

with open(SETTINGS_PATH, 'r') as file:
    settings = json.load(file)

ACCESS_PATH = settings["Access Database Path"]
RAW_DATA_PATH = settings["Raw Data Path"]
RAW_DATA_SHEET = settings["Raw Data Sheet Name"]
ACCESS_TBL = settings["Import Access Table Name"]
ACCESS_FORM = settings["Access Form Name"]
SF_UPLOAD = settings["SF Upload File Name"]
UDB_UPLOAD = settings["UDB Upload File Name"]
UDB_EXCLUDE = settings["UDB Exclude File Name"]
DEMO_DEST_PATH = settings["Demo Folder Destination Path"]
SOP_PATH = settings["Demo SOP Path"]