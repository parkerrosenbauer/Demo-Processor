from __future__ import annotations
import pyodbc
import warnings
import pandas as pd
import win32com.client as win32
from accessdb import to_accessdb


class MSAccess:
    def __init__(self, db_path: str):
        """Initialize MSAccess

        :param db_path: path to Access Database
        :type db_path: str
        """
        self.path = db_path
        self.conn_str = (
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path}'
        )

    def download_to_excel(self, tbl_name: str, destination: str, sheet="") -> None:
        """Download Access table as an Excel sheet.

        :param tbl_name: name of table in Access to download
        :type tbl_name: str
        :param destination: path to download table to
        :type destination: str
        :param sheet: what to name the sheet in Excel (default tbl_name)
        :type sheet: str
        :return: None
        :rtype: None
        """
        sql = f"SELECT * FROM {tbl_name}"
        data = self.run_select_sql(sql_query=sql, method="df")
        pd.io.formats.excel.ExcelFormatter.header_style = None
        if sheet == "":
            sheet = tbl_name
        data.to_excel(destination, sheet_name=sheet, index=False)

    def form_fill_run(self, form: str, *fields: str) -> None:
        """Update the Access form and run the queries.

        :param form: name of form in Access
        :type form: str
        :param fields: fields required to fill the form
        :type fields: str
        :return: None
        :rtype: None
        """
        try:
            cnxn = win32.Dispatch('Access.Application')
            db = cnxn.OpenCurrentDatabase(self.path)

            cnxn.Visible = True

            cnxn.DoCmd.OpenForm(form)

            cnxn.Forms(form).Fill_Form(*fields)
            cnxn.Forms(form).RunForm_Click()
        finally:
            cnxn.DoCmd.CloseDatabase()
            cnxn.Quit()

    def run_select_sql(self, sql_query: str, method="print") -> None | pd.DataFrame:
        """Return the SELECT query.

        :param sql_query: sql query
        :type sql_query: str
        :param method: way to return the data (print or df) (default print)
        :type method: str
        :return: data from select query
        :rtype: None | pd.DataFrame
        """
        cnxn = pyodbc.connect(self.conn_str)

        if method == "print":
            cursor = cnxn.execute(sql_query)
            for row in cursor:
                print(row)
            cursor.close()
        elif method == "df":
            data = pd.read_sql(sql_query, cnxn)
            return data
        else:
            warnings.warn("That is not a valid method.")

        cnxn.close()

    def run_sql(self, sql_query: str) -> None:
        """Run the SQL query in Access.

        :param sql_query: sql query
        :type sql_query: str
        :return: None
        :rtype: None
        """
        cnxn = pyodbc.connect(self.conn_str)
        cursor = cnxn.execute(sql_query)
        cnxn.commit()
        cursor.close()
        cnxn.close()
        if "select" in sql_query.lower():
            warnings.warn("To see the output of the SELECT statement, use run_select_sql(sql_query) instead")

    def run_access_query(self, access_query: str) -> None:
        """Run the predefined Access query.

        :param access_query: name of query in Access
        :type access_query: str
        :return: None
        :rtype: None
        """
        cnxn = pyodbc.connect(self.conn_str)
        sql = f'\u007bCALL {access_query}\u007d'
        cursor = cnxn.execute(sql)
        cnxn.commit()
        cursor.close()
        cnxn.close()

    def upload_table(self, file_path: str, file_sheet: str, tbl_name: str) -> None:
        """Upload Excel file to Access.

        :param file_path: path to Excel file to upload
        :type file_path: str
        :param file_sheet: name of sheet in file to upload
        :type file_sheet: str
        :param tbl_name: what to name the uploaded table in Access
        :type tbl_name: str
        :return: None
        :rtype: None
        """
        data = pd.read_excel(file_path, sheet_name=file_sheet)
        for col in data.columns:
            if len(col) > 25:
                data.rename(columns={col: col[0:25]}, inplace=True)
        data = data.loc[:, ~data.apply(lambda x: x.duplicated(), axis=1).all()].copy()
        data.to_accessdb(self.path, tbl_name)
