import access_interface.access as access
import file_processing.constants as c
import pandas as pd
import os
import re

file_path = r'\\CT-FS10\BLR_Share\Marketing\_Database Management\Leads DB List Upload Requests\Healthcare\PRODUCT ' \
            r'DEMOS\2023 Demos\SelectCoderDemo-011223'
upload = pd.read_excel(r'\\CT-FS10\BLR_Share\Marketing\_Database Management\Leads DB List Upload '
                         r'Requests\Healthcare\PRODUCT DEMOS\2023 '
                         r'Demos\SelectCoderDemo-011223\SelectCoder-011223-SFDC_Upload.xlsx', sheet_name="SFDC_Upload")

files = os.listdir(file_path)
r = re.compile(".*SF.*Validation.*")
sfdc_validation_file = sorted(list(filter(r.match, files)), reverse=True)[0]

sfdc_validation_path = os.path.join(file_path, sfdc_validation_file)
validation = pd.read_excel(sfdc_validation_path, sheet_name=1)

valid = pd.merge(validation, upload, how="left", left_on=["Last Name", "Email"], right_on=["LastName", "Email"])
valid = valid[["Stage", "Converted Date", "Lead Owner", "AG", "SFDC ID (18 digit)", "Tracking Code", "Email"]]
valid = valid[valid["Lead Owner"].notnull()]
valid.drop_duplicates(subset=["SFDC ID (18 digit)"], inplace=True)

new_data = pd.merge(upload, valid, how="left", on=["Email"])
new_data = new_data.fillna('')
new_data.loc[new_data["Existing Lead ID"] == '', "Existing Lead ID"] = new_data["SFDC ID (18 digit)"]

print(new_data["Existing Lead ID"])