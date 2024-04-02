import datetime
import hashlib
import requests
import sys
import json
import urllib3
import pandas

#Suppress Insecure HTTPS Request Warning
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def TokenGenerator():
    timestamp = (int(datetime.datetime.now().timestamp()) | 0)
    hash1 = hashlib.sha256()
    userNonce = 'YOUR_USERNONCE_HERE'
    userKey = 'YOUR_USERKEY_HERE'
    integrationIdentifier = ''
    hash1.update(bytes(str(timestamp), "utf-8"))
    hash1.update(bytes(userKey, "utf-8"))
    encrypted = hash1.hexdigest()

    if (len(integrationIdentifier) > 0):
        authToken = str(userNonce) + ":" + str(timestamp) + ":" + str(encrypted) + ":" + str(integrationIdentifier)
        return authToken
    else:
        authToken = str(userNonce) + ":" + str(timestamp) + ":" + str(encrypted)
        return authToken

#SERVER IP
serverip = "YOUR_SERVER_IP_HERE"

#generate authorization token for login POST request
authToken1 = TokenGenerator()
loginquery = {
    'username': 'YOUR_AVIGILON_USERNAME',
    'password': 'YOUR_AVIGILON_PASSWORD',
    'authorizationToken': str(authToken1),
    'clientName': 'APICLIENT'
    }

#initial login request, return session key / headers
s = requests.Session()
try:
    response = s.post('https://' + serverip + ':8443/mt/api/rest/v1/login', json=loginquery, verify=False)
    sessionkey = response.json()
    print("Login:")
    print(response.json())
    logoutsessionkey = {'session': str(sessionkey["result"]["session"])}
    sessionheaders = {'x-avg-session': str(sessionkey["result"]["session"])}
    s.headers.update(sessionheaders)
except s.Timeout as error:
    print("Exception in login: ")
    print(error)

#Logout Function
def logout(logoutkey):
    try:
        response = s.post('https://' + serverip + ':8443/mt/api/rest/v1/logout', json=logoutkey, verify=False)
        print("Logout: ")
        print(response.json())
    except Exception as error:
        print("Exception in logout: ")
        print(error)


#get IDs from LPR List
print("Fetching ID's. . . .")
try:
    response = s.get('https://' + serverip + ':8443/mt/api/rest/v1/lpr/lists', verify=False)
    watchidlist = []
    for i in range(len((response.json())["result"]["watchlists"])):
        watchidlist.append((response.json())["result"]["watchlists"][i]['id'])
except Exception as error:
    print("Exception in fetching IDs: ")
    print(error)
    logout(logoutsessionkey)

#Iterate through all watchlist ID's to pull all information
newLPRDict = {'Name': [], 'Description': [], 'License Plate': []}
print("Fetching and formatting LPR list. . . .")
try:
    for watchID in watchidlist:   

        response = s.get('https://' + serverip + ':8443/mt/api/rest/v1/lpr/list', params={'id': str(watchID)}, verify=False)
        response = response.json()
        newLPRDict['Name'].append(response['result']['watchlist']['name'])
        newLPRDict['Description'].append(response['result']['watchlist']['description'])

        if len(response['result']['watchlist']['watches']) == 0:
            newLPRDict['License Plate'].append('')
        elif len(response['result']['watchlist']['watches']) > 1:
            platestring = ""
            for pIndex in range(len(response['result']['watchlist']['watches'])):
                if pIndex == len(response['result']['watchlist']['watches']) - 1:
                    platestring = platestring + str(response['result']['watchlist']['watches'][pIndex]['licensePlate'])
                else:
                    platestring = platestring  + str(response['result']['watchlist']['watches'][pIndex]['licensePlate']) + ", "
            newLPRDict['License Plate'].append(platestring)
        else:
            newLPRDict['License Plate'].append(response['result']['watchlist']['watches'][0]['licensePlate'])
            
except Exception as error:
    print("Exception in fetching / formatting list: ")
    print(error)
    logout(logoutsessionkey)

#save into excel file
print("Generating Spreadsheet. . . .")
now = datetime.datetime.now()
date_time = now.strftime("%m-%d-%Y at %H:%M")
datedict = {'Last Updated': [str(date_time)]}
filename = f"LicensePlateWatchList.xlsx"


df1 = pandas.DataFrame.from_dict(datedict)
lprDataFrame = pandas.DataFrame.from_dict(newLPRDict)

#sort list by name
df2 = lprDataFrame.sort_values(by='Name', ascending=True)

writer = pandas.ExcelWriter(filename, mode='w', engine="xlsxwriter")
df1.to_excel(writer, sheet_name='License Plate Watch List', startrow=0, header=True, index=False)
df2.to_excel(writer, sheet_name='License Plate Watch List', startrow=4, header=False, index=False)

#get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['License Plate Watch List']

#get dimensions of dataframe
(max_row, max_col) = df2.shape

#create list of column headers to use in add_table()
column_settings = []
for header in df2.columns:
    column_settings.append({'header': header})

#add_table(firstrow, firstcol, lastrow, lastcol, format options) --> Zero Indexed
worksheet.add_table(3, 0, max_row + 2, max_col - 1, {'columns': column_settings})
worksheet.set_column(0, max_col - 1, 50)

writer.close()

#logout of rest api session
logout(logoutsessionkey) 
print("Done!")
