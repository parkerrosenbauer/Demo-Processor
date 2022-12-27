import os
import pandas as pd
from file_processing.demo import Demo
import file_processing.constants as demo_c
import file_processing.file_paths as const

ARCHIVE_PATH = const.ARCHIVE_PATH
COUNTS_SHEET = "Upload_Counts"
UDB_SHEET = "UDB_Uploads"
SFDC_SHEET = "SFDC_Uploads"
RAW_SHEET = "Raw_Data"


class ArchiveMgr:
    def __init__(self, demo: Demo):
        self.demo = demo
        self.raw = None
        self.sfdc = None
        self.udb = None
        self.counts = None

    def append_raw(self) -> None:
        """"""

        # append raw data
        self.raw = pd.read_excel(ARCHIVE_PATH, sheet_name=RAW_SHEET)
        data_name = os.path.basename(demo_c.RAW_DATA_PATH)
        raw_path = os.path.join(self.demo.destination_path, data_name)
        new_data = pd.read_excel(raw_path, sheet_name=demo_c.RAW_DATA_SHEET)
        new_data = new_data[['Attended', 'Last Name', 'First Name', 'Email Address', 'State/Province', 'Phone',
                             'Organization', 'Job Title', 'Unsubscribed']]
        new_data["Date"] = self.demo.demo_date
        new_data["Type"] = "HC Demo"
        new_data = pd.concat([self.raw, new_data], ignore_index=True)
        new_data["Date"] = pd.to_datetime(new_data["Date"])
        new_data["Date"] = new_data["Date"].dt.strftime('%m/%d/%Y')
        with pd.ExcelWriter(ARCHIVE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            new_data.to_excel(writer, sheet_name=RAW_SHEET, index=False)
        del new_data
        self.raw = None

    def append_sfdc(self):
        """"""
        self.sfdc = pd.read_excel(ARCHIVE_PATH, sheet_name=SFDC_SHEET)
        new_data = pd.read_excel(self.demo.sf_path, sheet_name=self.demo.sf_upload)
        new_data = new_data[list(set(self.sfdc.columns) & set(new_data.columns))]
        new_data["Date"] = self.demo.demo_date
        new_data["Type"] = "HC Demo"
        new_data = pd.concat([self.sfdc, new_data], ignore_index=True)
        new_data["Date"] = pd.to_datetime(new_data["Date"])
        new_data["Date"] = new_data["Date"].dt.strftime('%m/%d/%Y')
        with pd.ExcelWriter(ARCHIVE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            new_data.to_excel(writer, sheet_name=SFDC_SHEET, index=False)
        del new_data
        self.sfdc = None

    def append_udb(self):
        """"""
        self.udb = pd.read_excel(ARCHIVE_PATH, sheet_name=UDB_SHEET)
        new_data = pd.read_excel(self.demo.udb_path, sheet_name=self.demo.udb_upload)
        new_data["Date"] = self.demo.demo_date
        new_data["Type"] = "HC Demo"
        new_data = new_data[list(set(self.udb.columns) & set(new_data.columns))]
        new_data = pd.concat([self.udb, new_data], ignore_index=True)
        new_data["Date"] = pd.to_datetime(new_data["Date"])
        new_data["Date"] = new_data["Date"].dt.strftime('%m/%d/%Y')
        with pd.ExcelWriter(ARCHIVE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            new_data.to_excel(writer, sheet_name=UDB_SHEET, index=False)
        del new_data
        self.udb = None

    def append_counts(self):
        """"""
        self.counts = pd.read_excel(ARCHIVE_PATH, sheet_name=COUNTS_SHEET)
        demo_counts = self.demo.counts
        attendee_counts = {
            "Date": self.demo.demo_date,
            "Type": "HC Demo",
            "Pub_Code": self.demo.pub,
            "Initial_Count": demo_counts.retrieve_one("a_initial_count"),
            "Internal_Records": demo_counts.retrieve_one("a_internal_records"),
            "SF_TrackingCode": demo_counts.retrieve_one("tmattendee_code"),
            "SF_Count": demo_counts.retrieve_one("tmattendee_count"),
            "UDB_TrackingCode": demo_counts.retrieve_one("attendee_code"),
            "UDB_Uploaded_Count": demo_counts.retrieve_one("attendee_count"),
            "UDB_MasterSupp": demo_counts.retrieve_one("a_mastersupp"),
            "UDB_IsActiveFalse": demo_counts.retrieve_one("a_activefalse"),
            "UDB_HardBounce": demo_counts.retrieve_one("a_hardbounce"),
            "SF_New_Leads": demo_counts.retrieve_one("a_new"),
            "SF_Updated_Leads": demo_counts.retrieve_one("a_lead_update"),
            "SF_Updated_Contact": demo_counts.retrieve_one("a_contact_update"),
            "Converted": demo_counts.retrieve_one("a_converted"),
            "SF_Dead": 0,
            "Left_Dead": 0,
            "Flipped_Open": demo_counts.retrieve_one("flipped_open"),
            "Contact_no_Lead": demo_counts.retrieve_one("a_contact_no_lead"),
            "Null_Phone": demo_counts.retrieve_one("a_null_phone"),
            "Merged": demo_counts.retrieve_one("a_merged"),
            "BadEmail": demo_counts.retrieve_one("a_bad_email"),
        }
        nonattendee_counts = {
            "Date": self.demo.demo_date,
            "Type": "HC Demo",
            "Pub_Code": self.demo.pub,
            "Initial_Count": demo_counts.retrieve_one("na_initial_count"),
            "Internal_Records": demo_counts.retrieve_one("na_internal_records"),
            "SF_TrackingCode": demo_counts.retrieve_one("tmnonattendee_code"),
            "SF_Count": demo_counts.retrieve_one("tmnonattendee_count"),
            "UDB_TrackingCode": demo_counts.retrieve_one("nonattendee_code"),
            "UDB_Uploaded_Count": demo_counts.retrieve_one("nonattendee_count"),
            "UDB_MasterSupp": demo_counts.retrieve_one("na_mastersupp"),
            "UDB_IsActiveFalse": demo_counts.retrieve_one("na_activefalse"),
            "UDB_HardBounce": demo_counts.retrieve_one("na_hardbounce"),
            "SF_New_Leads": demo_counts.retrieve_one("na_new"),
            "SF_Updated_Leads": demo_counts.retrieve_one("na_lead_update"),
            "SF_Updated_Contact": demo_counts.retrieve_one("na_contact_update"),
            "Converted": demo_counts.retrieve_one("na_converted"),
            "SF_Dead": 0,
            "Left_Dead": demo_counts.retrieve_one("left_dead"),
            "Flipped_Open": 0,
            "Contact_no_Lead": demo_counts.retrieve_one("na_contact_no_lead"),
            "Null_Phone": demo_counts.retrieve_one("na_null_phone"),
            "Merged": demo_counts.retrieve_one("na_merged"),
            "BadEmail": demo_counts.retrieve_one("na_bad_email"),
        }
        attendee_data = pd.DataFrame(attendee_counts, index=[0])
        nonattendee_data = pd.DataFrame(nonattendee_counts, index=[0])
        new_data = pd.concat([self.counts, attendee_data], ignore_index=True)
        new_data = pd.concat([new_data, nonattendee_data], ignore_index=True)
        new_data["Date"] = pd.to_datetime(new_data["Date"])
        new_data["Date"] = new_data["Date"].dt.strftime('%m/%d/%Y')
        with pd.ExcelWriter(ARCHIVE_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            new_data.to_excel(writer, sheet_name=COUNTS_SHEET, index=False)

        self.counts = None
