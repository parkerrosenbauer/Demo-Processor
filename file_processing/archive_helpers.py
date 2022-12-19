import pandas as pd
from file_processing import demo
from logs.log import logger


def attend_nonattend_counts(path: str, sheet: str, tc_col: str) -> tuple:
    """Separates a count of leads into attendees and nonattendees then adds them to the verification counts.

    :param path: path to the desired file
    :type path: str
    :param sheet: sheet of the desired file
    :type sheet: str
    :param tc_col: name of the tracking code column
    :type tc_col: str
    :return: attend count, nonattend count
    :rtype: tuple
    """
    try:
        data = pd.read_excel(path, sheet_name=sheet)
    except ValueError:
        logger.warning("%s sheet was not found", sheet)
        return 0, 0
    else:
        attend = len(data[data[tc_col].str.contains("AC")])
        nonattend = len(data[data[tc_col].str.contains("BC")])
        return attend, nonattend


def sfdc_counts(demo_obj: demo.Demo) -> None:
    """Updates the sfdc demo counts for the archive.

    :param demo_obj: Demo object
    :type demo_obj: Demo
    :return: None
    :rtype: None
    """
    new = attend_nonattend_counts(demo_obj.sf_path, "New", "TrackingCode")
    lead_update = attend_nonattend_counts(demo_obj.sf_path, "LeadUpdate", "TrackingCode")
    contact_update = attend_nonattend_counts(demo_obj.sf_path, "ContactUpdate", "TrackingCode")
    contact_no_lead = attend_nonattend_counts(demo_obj.sf_path, "ContactNoLead", "TrackingCode")
    nullphone = attend_nonattend_counts(demo_obj.sf_path, "NullPhone", "TrackingCode")

    demo_obj.counts.update_counts(a_new=new[0],
                                  na_new=new[1],
                                  a_lead_update=lead_update[0],
                                  na_lead_update=lead_update[1],
                                  a_contact_update=contact_update[0],
                                  na_contact_update=contact_update[1],
                                  a_contact_no_lead=contact_no_lead[0],
                                  na_contact_no_lead=contact_no_lead[1],
                                  a_null_phone=nullphone[0],
                                  na_null_phone=nullphone[1])

    # sf count if non-attendees are excluded
    if int(demo_obj.counts.retrieve_one("tmnonattendee_count")) == 0 \
            and int(demo_obj.counts.retrieve_one("sf_excluded") > 0):
        demo_obj.counts.update_counts(tmnonattendee_count=demo_obj.counts.retrieve_one("sf_excluded"))


def udb_counts(demo_obj: demo.Demo) -> None:
    """Updates the udb demo counts for the archive.

    :param demo_obj:
    :type demo_obj:
    :return:
    :rtype:
    """
    udb = attend_nonattend_counts(demo_obj.udb_path, demo_obj.udb_upload, "TrackingCode")
    exclude = pd.read_excel(demo_obj.exclude_path, sheet_name=demo_obj.udb_exclude)
    a_exclude = exclude[exclude["TrackingCode"].str.contains("AC")]
    a_ms_count = len(a_exclude[a_exclude["MasterSuppression"].notnull()])
    a_iaf_count = len(a_exclude[a_exclude["IsActiveFalse"].notnull()])
    na_exclude = exclude[exclude["TrackingCode"].str.contains("BC")]
    na_ms_count = len(na_exclude[na_exclude["MasterSuppression"].notnull()])
    na_iaf_count = len(na_exclude[na_exclude["IsActiveFalse"].notnull()])

    demo_obj.counts.update_counts(attendee_count=udb[0],
                                  nonattendee_count=udb[1],
                                  a_mastersupp=a_ms_count,
                                  na_mastersupp=na_ms_count,
                                  a_activefalse=a_iaf_count,
                                  na_activefalse=na_iaf_count)


