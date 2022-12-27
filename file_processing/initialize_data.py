import json
import file_processing.demo as demo
import file_processing.file_paths as demo_paths


def initialize() -> None:
    """Initialize json file with demo_info.

    :return: None
    :rtype: None
    """
    data = {
        "a_initial_count": 0,
        "na_initial_count": 0,
        "a_internal_records": 0,
        "na_internal_records": 0,
        "a_null_phone": 0,
        "na_null_phone": 0,
        "a_contact_no_lead": 0,
        "na_contact_no_lead": 0,
        "a_new": 0,
        "na_new": 0,
        "a_lead_update": 0,
        "na_lead_update": 0,
        "a_contact_update": 0,
        "na_contact_update": 0,
        "flipped_open": 0,
        "left_dead": 0,
        "a_converted": 0,
        "na_converted": 0,
        "updated_leads": 0,
        "as_requested": 0,
        "requested_assign": "",
        "tmattendee_code": "",
        "tmattendee_count": 0,
        "tmnonattendee_code": "",
        "tmnonattendee_count": 0,
        "total": 0,
        "attendee_count": 0,
        "nonattendee_count": 0,
        "attendee_code": "",
        "nonattendee_code": "",
        "a_mastersupp": 0,
        "na_mastersupp": 0,
        "a_activefalse": 0,
        "na_activefalse": 0,
        "a_bad_email": 0,
        "na_bad_email": 0,
        "a_merged": 0,
        "na_merged": 0,
        "contact_no_lead": 0,
        "null_phone": 0,
        "converted": 0,
        "udb_excluded": 0,
        "udb_uploaded": 0,
        "sf_excluded": 0,
        "a_hardbounce": 0,
        "na_hardbounce": 0,
    }
    demo_id = [
        f"{row['Demo Type']} ({row['Webinar Date']})"
        for idx, row in demo.DEMO_INFO.iterrows()
    ]
    if "NaN" in demo_id:
        demo_id.remove("NaN")

    count_dicts = {
        idx: data
        for idx in demo_id
    }
    with open(demo_paths.VALIDATION_COUNTS, 'w') as file:
        json.dump(count_dicts, file)
