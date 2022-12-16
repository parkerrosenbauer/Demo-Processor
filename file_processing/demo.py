import os
import datetime
import pandas as pd
import access_interface.access as access
import file_processing.constants as demo_c
import file_processing.validation as v
from logs.log import logger

DEMO_INFO_PATH = r"\\CT-FS10\BLR_Share\Marketing\_Database Management\PRosenbauer\demo_automation\data\Demo_Info.csv"
logger.debug(DEMO_INFO_PATH)
DEMO_INFO = pd.read_csv(DEMO_INFO_PATH)


class Demo:
    def __init__(self, date: datetime.datetime, demo_type: str = None):
        """Initialize Demo.

        :param date: date of demo
        :type date: str
        :param demo_type: type of demo if there is more than one for given date (default None)
        :type demo_type: str
        """
        logger.debug("Demo's date %s", date)

        self.demo_date = date
        demo_info = DEMO_INFO[(DEMO_INFO["Webinar Date"] == self.demo_date.strftime('%#m/%#d/%Y'))]

        if demo_type is None:
            self.demo_type = demo_info["Demo Type"].iloc[0]
        else:
            self.demo_type = demo_type
            demo_info = demo_info[demo_info["Demo Type"] == demo_type]

        self.sf_attend = demo_info["Tracking Code"].iloc[0]
        self.sf_non_attend = demo_info["Tracking Code"].iloc[1]
        self.udb_attend = demo_info["Tracking Code"].iloc[2]
        self.udb_non_attend = demo_info["Tracking Code"].iloc[3]
        self.pub = demo_info["Pub Code"].iloc[0]

        self.sf_upload = demo_c.SF_UPLOAD
        self.udb_upload = demo_c.UDB_UPLOAD
        self.udb_exclude = demo_c.UDB_EXCLUDE
        self.destination_folder = f"{self.demo_type}Demo-{self.demo_date.strftime('%m%d%y')}"
        self.destination_path = os.path.join(demo_c.DEMO_DEST_PATH, self.destination_folder)
        self.sf_file = f"{self.demo_type}-{self.demo_date.strftime('%m%d%y')}-{self.sf_upload}.xlsx"
        self.sf_path = os.path.join(self.destination_path, self.sf_file)
        self.udb_file = f"{self.demo_type}-{self.demo_date.strftime('%m%d%y')}-{self.udb_upload}.xlsx"
        self.udb_path = os.path.join(self.destination_path, self.udb_file)
        self.exclude_file = f"{self.demo_type}-{self.demo_date.strftime('%m%d%y')}-{self.udb_exclude}.xlsx"
        self.exclude_path = os.path.join(self.destination_path, self.exclude_file)

        self.flip_to_open = []

        self.counts = v.Validation(self.demo_type, self.demo_date)

    @property
    def demo_date(self):
        return self._demo_date

    @demo_date.setter
    def demo_date(self, date: datetime.datetime):
        """Validates date of demo.

        :param date: date of demo
        :type date: datetime.datetime
        :return: None
        :rtype: None
        """
        valid_date = DEMO_INFO["Webinar Date"].str.contains(date.strftime('%#m/%#d/%Y')).any()
        if valid_date:
            self._demo_date = date
        else:
            raise ValueError(f"There is no demo scheduled for {date}")

    def run_through_access(self) -> None:
        """Process demo through Access.

        :return: None
        :rtype: None
        """
        cnxn = access.MSAccess(demo_c.ACCESS_PATH)
        try:
            cnxn.run_access_query("delete_leadimportfile")
        except Exception as e:
            logger.info(str(e))
        cnxn.upload_table(demo_c.RAW_DATA_PATH, demo_c.RAW_DATA_SHEET, demo_c.ACCESS_TBL)
        cnxn.form_fill_run(demo_c.ACCESS_FORM,
                           self.demo_type,
                           self.sf_attend,
                           self.sf_non_attend,
                           self.udb_attend,
                           self.udb_non_attend,
                           self.pub)
        cnxn.download_to_excel(self.sf_upload, self.sf_path)
        cnxn.download_to_excel(self.udb_upload, self.udb_path)
        cnxn.download_to_excel(self.udb_exclude, self.exclude_path)
        cnxn.download_to_excel("No_Master_Org_Match", os.path.join(r"\\CT-FS10\BLR_Share\Marketing\_Database "
                                                                   r"Management\PRosenbauer\UpdateMasterOrg",
                                                                   f"{self.demo_date}.xlsx"))
