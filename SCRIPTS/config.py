import os

mapping = {
    0: 'Withdrawn',
    1: 'Inactive',
    2: 'Retiree',
    3: 'Active'
}


ORDER_HDR_PATH = os.path.join(
        "C:\\",
        "PROJECTS",
        "INBOUD_IDOC",
        "INPUT",
        "SAP",
        "downloaded_data_ZE1ORDRHDR.xlsx",
    )

WORK_HDR_PATH = os.path.join(
        "C:\\",
        "PROJECTS",
        "INBOUD_IDOC",
        "INPUT",
        "SAP",
        "downloaded_data_ZWORK_ORDER_HDR.xlsx",
    )


OP2_PATH = os.path.join(
        "C:\\",
        "PROJECTS",
        "INBOUD_IDOC",
        "INPUT",
        "SAP",
        "downloaded_data_ZE1OPERATION2.xlsx",
    )

OP1_PATH = os.path.join(
        "C:\\",
        "PROJECTS",
        "INBOUD_IDOC",
        "INPUT",
        "SAP",
        "downloaded_data_ZE1OPERATION1.xlsx",
    )


PA0001_PATH = os.path.join(
        "c:\\", "PROJECTS", "INBOUD_IDOC", "INPUT", "SAP", "PA001_DATA.xlsx"
    )