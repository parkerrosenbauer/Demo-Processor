import os
import re
import pandas as pd
from win32com.client import constants, gencache
from logs.log import logger
import file_processing.demo as demo
import file_processing.constants as demo_c

pd.io.formats.excel.ExcelFormatter.header_style = None


# --------------------- INITIAL COUNTS --------------------- #
def initial_counts(demo_obj: demo.Demo) -> None:
    """Find initial and internal counts.

    :param demo_obj: current demo object
    :type demo_obj: demo.Demo
    :return: None
    :rtype: None
    """
    internal_emails = ["@datacenterhawk.com", "@acdis.org", "@blr.com", "@bluepointleadership.com", "@ccmi.com",
                       "@decisionhealth.com", "@h3.group", "@hci.org", "@hcpro.com", "@mleesmith.com",
                       "@simplifycompliance.com", "@ucg.com"]
    pattern = '|'.join(internal_emails)

    data = pd.read_excel(demo_c.RAW_DATA_PATH, sheet_name=demo_c.RAW_DATA_SHEET)
    a_initial = data[data["Attended"] == "Yes"]
    na_initial = data[data["Attended"] == "No"]
    a_internal = a_initial[a_initial["Email Address"].str.contains(pattern)]
    na_internal = na_initial[na_initial["Email Address"].str.contains(pattern)]
    demo_obj.counts.update_counts(a_initial_count=len(a_initial),
                                  na_initial_count=len(na_initial),
                                  a_internal_records=len(a_internal),
                                  na_internal_records=len(na_internal))


# --------------------- DESTINATION FOLDER --------------------- #
def create_destination(path: str) -> None:
    """ Create directory.

    :param path: path of new directory
    :type path: str
    :return: None
    :rtype: None
    """
    try:
        os.mkdir(path)
    except FileExistsError:
        logger.warning("The folder for this demo already exists. This could indicate the demo was already processed.")


# --------------------- CREATE PIVOT TABLES --------------------- #
def pivot_table(file: str, sheet: str, tables: list) -> None:
    """Create pivot table(s) in file.

    :param file: Excel path
    :type file: str
    :param sheet: data sheet
    :type sheet: str
    :param tables: list of dictionaries with table info
    :type tables: list
    :return: None
    :rtype: None
    """
    excel = gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    try:
        wb = excel.Workbooks.Open(file)
    except Exception as e:
        excel.Quit()
        raise e

    ws1 = wb.Sheets(sheet)

    ws2_name = 'Summary'
    wb.Sheets.Add().Name = ws2_name

    ws2 = wb.Sheets(ws2_name)

    for table in tables:
        table_name = f"Table{table['pt_num']}"
        pc = wb.PivotCaches().Create(SourceType=constants.xlDatabase, SourceData=ws1.UsedRange)

        pc.CreatePivotTable(TableDestination=f'{ws2_name}!R{table["pt_row"]}C{table["pt_col"]}', TableName=table_name)

        ws2.Select()
        ws2.Cells(table["pt_row"], table["pt_col"]).Select()

        for i, field in enumerate(table["pt_fields"]):
            ws2.PivotTables(table_name).AddDataField(Field=ws2.PivotTables(table_name).PivotFields(field[0]),
                                                     Caption=field[1], Function=field[2])

        for field_list, field_r in ((table["pt_filters"], constants.xlPageField),
                                    (table["pt_rows"], constants.xlRowField)):
            for i, value in enumerate(field_list):
                ws2.PivotTables(table_name).PivotFields(value).Orientation = field_r
                ws2.PivotTables(table_name).PivotFields(value).Position = i + 1

        ws2.PivotTables(table_name).ShowValuesRow = True
        ws2.PivotTables(table_name).ColumnGrand = True

    wb.Save()


# --------------------- SFDC PRE-VALIDATION --------------------- #
def sfdc_pre_val(demo_obj: demo.Demo) -> None:
    """Clean the sfdc file for review.

    :param demo_obj: current demo object
    :type demo_obj: demo.Demo
    :return: None
    :rtype: None
    """
    # put the salesforce leads into a variable
    sfdc = pd.read_excel(demo_obj.sf_path, sheet_name=demo_obj.sf_upload)
    sfdc = sfdc.fillna('')

    # validates name proper casing
    sfdc.LastName.str.title()
    sfdc.FirstName.str.title()
    sfdc.loc[sfdc.LastName == '', 'LastName'] = '[Unknown]'

    # if the phone number is empty, it tries to populate it with the existing lead phone. If both fields are empty,
    # it stores the record separately to be uploaded into the NullPhone tab
    sfdc.loc[sfdc["PhoneNumber"] == '', 'PhoneNumber'] = sfdc["Existing Lead Phone"]
    null_phone = sfdc[sfdc["PhoneNumber"] == '']
    sfdc = sfdc[sfdc["PhoneNumber"] != '']
    sfdc = sfdc.drop(['Existing Lead Phone'], axis=1)

    # deletes the phone ext column if empty
    if (sfdc['PhoneExt'] == '').all():
        sfdc = sfdc.drop(['PhoneExt'], axis=1)

    # replaces the company name with the master company name based off domain
    # if the master company name is left blank, the existing company name will be marked with review needed
    sfdc.loc[sfdc["Master Name"] != '', 'Company'] = sfdc['Master Name']
    sfdc.loc[sfdc["Master Name"] == '', 'Company'] = 'REVIEW ' + sfdc.Company
    sfdc = sfdc.drop(['Master Name'], axis=1)

    """cleans the secondary description of duplicates - note, it's not a perfect system and this field will need to 
    be reviewed, but in essence, if a secondary description is X / Y / Z, it will compare there are no duplicated 
    terms within X, then Y, then Z, then it will make sure that X does not equal Y and does not equal Z (and do the 
    same for each other section) finally, it will make sure that X is not contained with Y or Z, if it is, 
    X is removed for an example of the above line, if the description is "HCPro Library Card / Library Card", 
    Library Card is within HCPro Library Card, so that is dropped, giving the result: "HCPro Library Card" """

    def uniquify(string, splitter=" "):
        output = []
        seen = set()
        for word in string.split(splitter):
            if word not in seen or word == '/':
                output.append(word)
                seen.add(word)

        if splitter != " ":
            for item in output:
                for thing in seen:
                    if item != thing and item in thing:
                        try:
                            output.remove(item)
                        except ValueError:
                            pass

        return splitter.join(output)

    for idx, desc in sfdc["Current Secondary Description"].items():
        try:
            sfdc.at[idx, "Current Secondary Description"] = uniquify(desc, splitter=" / ")
            sections = sfdc.at[idx, "Current Secondary Description"].split(' / ')
            new_line = []
            for section in sections:
                new_line.append(uniquify(section))
            sfdc.at[idx, "Current Secondary Description"] = ' / '.join(new_line)
        except AttributeError:
            pass

    # checks each record that is marked dead in salesforce and determines if they were an attendee or not
    # if not, it stores the record separately to be uploaded to the DeadNonAttendee tab, if they were
    # an attendee, the record's email will be presented, so they can be manually flipped to open
    dead = sfdc[(sfdc["Dead Reason"] != '') & (sfdc["Current Marketing Note"].str.contains("Cold"))]
    sfdc = sfdc[(sfdc["Dead Reason"] == '') | (sfdc["Current Marketing Note"].str.contains("Warm"))]
    demo_obj.flip_to_open = list(
        sfdc[(sfdc["Dead Reason"] != '') & (sfdc["Current Marketing Note"].str.contains("Warm"))].Email)

    logger.info(demo_obj.flip_to_open)

    # separates the records with a contact id but no lead id and stores them separately to be uploaded
    # to the ContactNoLead tab
    cnl = sfdc[(sfdc["Existing Contact ID"] != '') & (sfdc["Existing Lead ID"] == '')]
    sfdc = sfdc[(sfdc["Existing Contact ID"] == '') | (sfdc["Existing Lead ID"] != '')]

    # removes unnecessary columns
    dfs = [sfdc, null_phone, dead, cnl]
    for df in dfs:
        df.drop(["Prior Marketing Note", "Prior Sales Note", "Prior Description", "Prior Secondary Description",
                 "Prior Lead Status",
                 "LastActivityDate", "Existing Lead Owner", "Existing Lead Owner ID", "ID"], axis=1, inplace=True)

    # update validation counts
    demo_obj.counts.update_counts(left_dead=len(dead),
                                  flipped_open=len(demo_obj.flip_to_open),
                                  null_phone=len(null_phone),
                                  contact_no_lead=len(cnl))

    # reformat the Excel file and separates it into the correct sheets
    with pd.ExcelWriter(demo_obj.sf_path, engine='openpyxl', mode='w') as writer:
        sfdc.to_excel(writer, sheet_name=demo_obj.sf_upload, index=False)
        if cnl.shape[0] > 0:
            cnl.to_excel(writer, sheet_name='ContactNoLead', index=False)
        if null_phone.shape[0] > 0:
            null_phone.to_excel(writer, sheet_name='NullPhone', index=False)
        if dead.shape[0] > 0:
            dead.to_excel(writer, sheet_name='DeadNonAttendee', index=False)

            # create validation pivot tables
    pts = [
        {
            "pt_num": 1,
            "pt_row": 2,
            "pt_col": 1,
            "pt_rows": ['AG', 'Current Owner'],
            "pt_filters": [],
            "pt_fields": [['Current Owner', 'Count of Current Owner', -4112]]
        },
        {
            "pt_num": 2,
            "pt_row": 20,
            "pt_col": 1,
            "pt_rows": ['PubCode', 'TrackingCode'],
            "pt_filters": [],
            "pt_fields": [['TrackingCode', 'Count of TrackingCode', -4112]]
        },
        {
            "pt_num": 3,
            "pt_row": 30,
            "pt_col": 1,
            "pt_rows": ['Current Sales Note', 'Current Marketing Note'],
            "pt_filters": [],
            "pt_fields": [['Current Marketing Note', 'Count of Current Marketing Note', -4112]]
        },
        {
            "pt_num": 4,
            "pt_row": 2,
            "pt_col": 4,
            "pt_rows": ['Silverpop Sync', 'Email'],
            "pt_filters": [],
            "pt_fields": [['Silverpop Sync', 'Count of Silverpop Sync', -4112]]
        }
    ]

    pivot_table(demo_obj.sf_path, demo_obj.sf_upload, pts)


# --------------------- UDB PRE-VALIDATION --------------------- #
def udb_pre_val(demo_obj: demo.Demo) -> None:
    """Clean the udb file using the cleaned sfdc file for review.

    :param demo_obj: current demo object
    :type demo_obj: Demo
    :return: None
    :rtype: None
    """
    sfdc = pd.read_excel(demo_obj.sf_path, sheet_name=demo_obj.sf_upload)
    udb = pd.read_excel(demo_obj.udb_path, sheet_name=demo_obj.udb_upload)
    cols = list(udb.columns)

    # merge udb and sfdc data on email
    udb = pd.merge(udb, sfdc, how="left", on="Email", suffixes=("", "_sf"))
    udb = udb.fillna('')

    # transfer the clean data from the sfdc file to the udb file
    def transfer_sfdc(udb_col, sfdc_col, review=False):
        nonlocal udb
        # where the salesforce field is not blank, fill in the udb field
        udb.loc[udb[sfdc_col] != '', udb_col] = udb[sfdc_col]
        if review:
            # if review is true, it will mark the field as needing to be reviewed
            udb.loc[udb[sfdc_col] == '', udb_col] = 'REVIEW ' + udb[udb_col]

    transfer_sfdc('FirstName', 'FirstName_sf')
    transfer_sfdc('LastName', 'LastName_sf')
    transfer_sfdc('State', 'State_sf')
    transfer_sfdc('PhoneNumber', 'PhoneNumber_sf')
    transfer_sfdc('Company', 'Company_sf', review=True)
    transfer_sfdc('CustomerTitle', 'CustomerTitle_sf')
    if 'PhoneExt_sf' in udb.columns:
        transfer_sfdc('PhoneExt', 'PhoneExt_sf')
    elif (udb['PhoneExt'] == '').all():
        udb = udb.drop(['PhoneExt'], axis=1)

    # remove [Unknown] from file
    udb = udb.replace(to_replace=r"[Unknown]", value="")

    # remove / from title field
    udb.CustomerTitle = udb.CustomerTitle.replace(to_replace='\/', value=' ', regex=True)
    udb.CustomerTitle = udb.CustomerTitle.replace('\s+', value=' ', regex=True)

    # removing the sf columns
    for col in udb.columns:
        if col not in cols:
            udb = udb.drop([col], axis=1)

    # update validation counts
    exclude = pd.read_excel(demo_obj.exclude_path, sheet_name=demo_obj.udb_exclude)

    demo_obj.counts.update_counts(attendee_code=demo_obj.udb_attend,
                                  nonattendee_code=demo_obj.udb_non_attend,
                                  udb_excluded=len(exclude),
                                  udb_uploaded=len(udb))

    # save the Excel file
    with pd.ExcelWriter(demo_obj.udb_path, engine='openpyxl', mode='w') as writer:
        udb.to_excel(writer, sheet_name=demo_obj.udb_upload, index=False)

    # TODO create validation pivot tables
    pts = [
        {
            "pt_num": 1,
            "pt_row": 2,
            "pt_col": 1,
            "pt_rows": ['OppProduct', 'TrackingCode'],
            "pt_filters": [],
            "pt_fields": [['TrackingCode', 'Count of TrackingCode', -4112]]
        },
        {
            "pt_num": 2,
            "pt_row": 8,
            "pt_col": 1,
            "pt_rows": ['SalesNotes', 'MarketingNotes'],
            "pt_filters": [],
            "pt_fields": [['MarketingNotes', 'Count of MarketingNotes', -4112]]
        },
        {
            "pt_num": 3,
            "pt_row": 14,
            "pt_col": 1,
            "pt_rows": ['LeadSource', 'Site', 'ParentCompanyID', 'NewsletterIDs'],
            "pt_filters": [],
            "pt_fields": []
        }
    ]

    pivot_table(demo_obj.udb_path, demo_obj.udb_upload, pts)


# --------------------- SFDC POST-VALIDATION --------------------- #
def sfdc_post_val(demo_obj: demo.Demo) -> None:
    """Prepare sfdc file for upload.

    :param demo_obj: current demo object
    :type demo_obj: Demo
    :return: None
    :rtype: None
    """
    # reads in the manually approved data
    sfdc = pd.read_excel(demo_obj.sf_path, sheet_name=demo_obj.sf_upload)
    try:
        cnl = pd.read_excel(demo_obj.sf_path, sheet_name="ContactNoLead")
    except ValueError:
        cnl = None

    sfdc = sfdc.fillna('')
    if cnl is not None:
        cnl = cnl.fillna('')
        cnl = cnl.drop(['Existing Lead ID'], axis=1)

    review_count = (sfdc["Company"].str.contains("REVIEW")).sum()
    if review_count > 0:
        logger.warning("Data may not have been manually reviewed, REVIEW found in Company column")

    # separates the data into new leads, lead updates, and contact updates
    new = sfdc[(sfdc["Existing Contact ID"] == '') & (sfdc["Existing Lead ID"] == '')]
    lead_update = sfdc[sfdc["Existing Lead ID"] != '']
    contact_update = sfdc[sfdc["Existing Contact ID"] != '']
    if cnl is not None:
        contact_update = pd.concat([contact_update, cnl], ignore_index=True)

    new = new.fillna('')
    lead_update = lead_update.fillna('')
    contact_update = contact_update.fillna('')

    # remove necessary columns for upload
    # new
    new_cols = ['Existing Lead ID', 'Existing Contact ID', 'Domain', 'AG', 'Current Secondary Description',
                'Dead Reason']
    for col in new_cols:
        try:
            new = new.drop([col], axis=1)
        except KeyError:
            logger.info('%s missing from New', col)

    if 'PhoneExt' in new.columns and (new['PhoneExt'] == '').all():
        new = new.drop(['PhoneExt'], axis=1)

    # lead update
    lead_cols = ['Existing Contact ID', 'Domain', 'AG', 'Dead Reason', 'Existing Lead Company']
    for col in lead_cols:
        try:
            lead_update = lead_update.drop([col], axis=1)
        except KeyError:
            logger.info('%s missing from LeadUpdate', col)

    if 'PhoneExt' in lead_update.columns and (lead_update['PhoneExt'] == '').all():
        lead_update = lead_update.drop(['PhoneExt'], axis=1)

    # contact update
    contact_cols = ['Existing Lead ID', 'LastName', 'FirstName', 'Email', 'Domain', 'State', 'PhoneNumber', 'Company',
                    'Existing Lead Company', 'CustomerTitle', 'AG',
                    'Silverpop Sync', 'Record Type ID', 'LeadSource', 'Current Lead Status', 'Dead Reason',
                    'Current Owner', 'Current Owner ID', 'Dead Reason', 'PhoneExt']
    for col in contact_cols:
        try:
            contact_update = contact_update.drop([col], axis=1)
        except KeyError:
            logger.info('%s missing from ContactUpdate', col)

    # add new sheets to the Excel file
    with pd.ExcelWriter(demo_obj.sf_path, engine='openpyxl', mode='a') as writer:
        if new.shape[0] > 0:
            new.to_excel(writer, sheet_name='New', index=False)
        if lead_update.shape[0] > 0:
            lead_update.to_excel(writer, sheet_name="LeadUpdate", index=False)
        if contact_update.shape[0] > 0:
            contact_update.to_excel(writer, sheet_name="ContactUpdate", index=False)

    # save the new sheets as CSVs
    if new.shape[0] > 0:
        new_path = os.path.join(demo_obj.destination_path,
                                f"{demo_obj.demo_type}-"
                                f"{demo_obj.demo_date.strftime('%m%d%y')}-{demo_obj.sf_upload}-New.csv")
        new.to_csv(new_path, index=False)
    if lead_update.shape[0] > 0:
        lead_path = os.path.join(demo_obj.destination_path,
                                 f"{demo_obj.demo_type}-"
                                 f"{demo_obj.demo_date.strftime('%m%d%y')}-{demo_obj.sf_upload}-LeadUpdate.csv")
        lead_update.to_csv(lead_path, index=False)
    if contact_update.shape[0] > 0:
        contact_path = os.path.join(demo_obj.destination_path,
                                    f"{demo_obj.demo_type}"
                                    f"-{demo_obj.demo_date.strftime('%m%d%y')}-{demo_obj.sf_upload}-ContactUpdate.csv")
        contact_update.to_csv(contact_path, index=False)


# --------------------- UDB POST-VALIDATION --------------------- #
def udb_post_val(demo_obj: demo.Demo) -> None:
    """Prepare udb file for upload

    :param demo_obj: current demo object
    :type demo_obj: Demo
    :return: None
    :rtype: None
    """
    udb = pd.read_excel(demo_obj.udb_path, sheet_name=demo_obj.udb_upload)
    udb = udb.fillna('')

    review_count = (udb["Company"].str.contains("REVIEW")).sum()
    if review_count > 0:
        logger.warning("Data may not have been manually reviewed, REVIEW found in Company column")

    # save the file as a CSV
    udb_csv_path = os.path.join(demo_obj.destination_path,
                                f"{demo_obj.demo_type}-"
                                f"{demo_obj.demo_date.strftime('%m%d%y')}-{demo_obj.udb_upload}.csv")
    udb.to_csv(udb_csv_path, index=False)


# --------------------- VALIDATION COUNTS --------------------- #
def validation_counts(demo_obj: demo.Demo) -> None:
    """Find counts after sfdc records upload.

    :param demo_obj: current demo object
    :type demo_obj: demo.Demo
    :return: None
    :rtype: None
    """
    files = os.listdir(demo_obj.destination_path)
    r = re.compile(".*SF.*Validation.*")
    sfdc_validation_file = sorted(list(filter(r.match, files)), reverse=True)[0]

    sfdc_validation_path = os.path.join(demo_obj.destination_path, sfdc_validation_file)
    validation = pd.read_excel(sfdc_validation_path, sheet_name=1)
    upload = pd.read_excel(demo_obj.sf_path, sheet_name=demo_obj.sf_upload)

    valid = pd.merge(validation, upload, how="left", left_on=["Last Name", "Email"], right_on=["LastName", "Email"])
    valid = valid[["Stage", "Converted Date", "Lead Owner", "AG", "SFDC ID (18 digit)", "Tracking Code"]]
    valid = valid[valid["Lead Owner"].notnull()]
    valid.drop_duplicates(subset=["SFDC ID (18 digit)"], inplace=True)

    updated_records = len(valid[(valid["AG"] == "Active") & (valid["Converted Date"].isnull()) &
                                (valid["Stage"].isnull())]["Lead Owner"])
    requested_records = len((valid[(valid["AG"] != "Active") & (valid["Converted Date"].isnull()) &
                                   (valid["Stage"].isnull())]["Lead Owner"]))

    assignment = list(valid.AG.unique())
    assignment = [str(x) for x in assignment]
    if "nan" in assignment:
        assignment.remove("nan")
    assignment = sorted(assignment, reverse=True)[0]
    if "BDR" in assignment:
        assignment = assignment.split()[0]
    else:
        assignment += " by territory"

    # tracking code counts
    nc_validation = valid[(valid["Converted Date"].isnull()) & (valid["Stage"].isnull())]
    attendee_count = len(nc_validation[nc_validation["Tracking Code"].str.contains("AC")])
    nonattendee_count = len(nc_validation[nc_validation["Tracking Code"].str.contains("BC")])
    total_count = attendee_count + nonattendee_count
    c_validation = valid[valid["Converted Date"].notnull()]
    a_convert = len(c_validation[c_validation["Tracking Code"].str.contains("AC")])
    na_convert = len(c_validation[c_validation["Tracking Code"].str.contains("BC")])

    # add to validation counts
    demo_obj.counts.update_counts(a_converted=a_convert,
                                  na_converted=na_convert,
                                  updated_leads=str(updated_records),
                                  as_requested=str(requested_records),
                                  requested_assign=assignment,
                                  tmattendee_code=demo_obj.sf_attend,
                                  tmattendee_count=attendee_count,
                                  tmnonattendee_code=demo_obj.sf_non_attend,
                                  tmnonattendee_count=nonattendee_count,
                                  total=total_count,
                                  converted=a_convert+na_convert)


# --------------------- GENERATE EMAIL --------------------- #
def generate_email(demo_obj: demo.Demo) -> None:
    """Generates demo communication from template.

    :param demo_obj: current demo object
    :type demo_obj: demo.Demo
    :return: None
    :rtype: None
    """
    with open(r"\\CT-FS10\BLR_Share\Marketing\_Database "
              r"Management\PRosenbauer\demo_automation\file_processing\files\demo_template.txt", 'r') as file:
        data = file.read()

        data = data.replace("[demo type]", demo_obj.demo_type)

        for item, value in demo_obj.counts.retrieve_all().items():
            data = data.replace(f"[{item}]", str(value))

    with open(r"\\CT-FS10\BLR_Share\Marketing\_Database "
              r"Management\PRosenbauer\demo_automation\file_processing\files\generated_email.txt", 'w') as file:
        file.write(data)

    with open(r"\\CT-FS10\BLR_Share\Marketing\_Database "
              r"Management\PRosenbauer\demo_automation\file_processing\files\generated_email.txt", 'r+') as file:
        lines = file.readlines()
        file.seek(0)
        for line in lines:
            if "- 0 " not in line:
                file.write(line)
        file.truncate()
    os.startfile(r"\\CT-FS10\BLR_Share\Marketing\_Database "
                 r"Management\PRosenbauer\demo_automation\file_processing\files\generated_email.txt")