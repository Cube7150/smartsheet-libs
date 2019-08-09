import smartsheet
import pandas as pd
import datetime
import re


def sheets_to_dataframe(sheetobject):
    datacols = [x.title for x in sheetobject.columns]
    datalists = list()
    for r in sheetobject.rows:
        datalist = list()
        for c in r.cells:
            try:
                if c.display_value == None and c.value == None:
                    datalist.append('')
                elif c.display_value != None:
                    datalist.append(c.display_value)
                elif c.value != None:
                    if re.match("\d{4}\-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z*", c.value) == None:
                        datalist.append(c.value)
                    else:
                        matchval = re.match("\d{4}\-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z*", c.value).string.replace("Z", '')
                        datalist.append(datetime.datetime.strptime(matchval, "%Y-%m-%dT%H:%M:%S"))
                else:
                    datalist.append("NOTAVAILABLE")
            except:
                datalist.append("ERROR!")
        datalists.append(datalist)
    return pd.DataFrame(datalists, columns=datacols)


def parse_dataframe(df):
    df.columns = COLSET_
    return df


def make_splits(df, prefixcount):
    return pd.concat(
        [parse_dataframe(df.filter(regex="IDNUM|^" + str(x) + "_")) for x in range(1, prefixcount + 1)]).sort_values(
        by="사번").reset_index(drop=True)


# Read Environment Variable for envs['APIKEY']
envs = {x: y for x, y in [x.strip().split("=") for x in open("c:\\cube.envs", "r").readlines()]}

ss_client = smartsheet.Smartsheet(envs['APIKEY'])

# Column postfix name set example (1_PCODE, 1_PORTION, 1_CATCODE, 2_PCODE, 2_PORTION, 2_CATCODE)
COLSET_ = ['사번', '프로젝트코드', '분류코드', '비중', '업무세부', '등급']

sid = '494368973973380'

sheet = ss_client.Sheets.get_sheet(sid)

df = sheets_to_dataframe(sheet)
df2 = df[df['IDNUM'] != ""]

make_splits(df2, 11).to_excel("c:\\temp\\abcabc.xlsx", index=None)
