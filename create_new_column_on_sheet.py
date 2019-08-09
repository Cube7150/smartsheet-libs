import smartsheet

# import pandas as pd
# import datetime
# import re

# Read Environment Variable for envs['APIKEY']
envs = {x: y for x, y in [x.strip().split("=") for x in open("c:\\cube.envs", "r").readlines()]}

# Create Smartsheet Client Object
ss_client = smartsheet.Smartsheet(envs['APIKEY'])

# Define new column information for adding
new_column_info = {
    'name': 'STAGE',
    'index': 0,
    'type': 'PICKLIST',
    'width': 60,
    'locked': False
}

target_sid = '6374109023102852'


def colname_checker(ss, colname):
    ss_titles = [x.title for x in ss.columns]
    return colname in ss_titles


def create_column(ss, colinfo):
    if not colname_checker(ss, colinfo['name']):
        col = ss_client.models.Column(
            {
                'title': colinfo['name'], 'type': colinfo['type'], 'width': colinfo['width'],
                'locked': colinfo['locked'], 'index': colinfo['index'], 'hidden': True,
                'options': ['S1',
                            'S2',
                            'S3',
                            'S4',
                            'S5',
                            'S6']
            }
        )
        res = ss.add_columns(col)
        result = res.message
    else:
        result = 'PASS'
    return result


def get_sheet(sid):
    return ss_client.Sheets.get_sheet(sid)


sheet_list = ss_client.Sheets.list_sheets(include_all=True)
tot_tot = list()
s_data = sheet_list.to_dict()['data']
for ss in s_data:
    if ss['name'].find("_프로젝트") > -1 and ss['name'].find("데모용") < 0:
        # print(ss['id'], ss['name'])
        ss = get_sheet(ss['id'])
        aa = create_column(ss, new_column_info)
        print(ss.permalink, ss.name, aa)
