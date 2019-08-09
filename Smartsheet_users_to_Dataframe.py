import smartsheet
import pandas as pd


# Make Dataframe for users information from ss_client
def get_users(ss_client):
    response = ss_client.Users.list_users(include_all=True)
    tmplist = list()
    for d in response.data:
        tmplist.append([y for x, y in d.to_dict().items()])
    headlist = [x for x, y in d.to_dict().items()]
    return pd.DataFrame(tmplist, columns=headlist)


# Read Environment Variable for envs['APIKEY']
envs = {x: y for x, y in [x.strip().split("=") for x in open("c:\\cube.envs", "r").readlines()]}
ss_client = smartsheet.Smartsheet(envs['APIKEY'])

#If I have any Dataframe variable, Power BI can catch them
#In this case, Smartsheet_Users variable will show your Power BI python datasource.
Smartsheet_Users = get_users(ss_client)


