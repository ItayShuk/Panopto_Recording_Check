import glob
import os
import webbrowser

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from panopto_oauth2 import PanoptoOAuth2
import urllib3
import requests
from datetime import datetime, timedelta
from urllib.parse import quote
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import config
import socket
import pandas as pd
import time
import pandas_read_xml


# python to EXE:
# https://towardsdatascience.com/how-to-easily-convert-a-python-script-to-an-executable-file-exe-4966e253c7e9
# auto-py-to-exe

def send_mail(subject, body):
    """
    Mailing done by Outlook web
    """
    email_sender = config.USER

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = ", ".join(config.TO_SEND)
    msg['Subject'] = subject

    msg.attach(MIMEText(f'{body}\n ', 'plain'))
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login(config.USER, config.PASSWORD)
        text = msg.as_string()
        server.sendmail(email_sender, config.TO_SEND, text)
        print('email sent')
        server.quit()
    except socket.error:
        print("SMPT server connection error")
    return True


def authorization(requests_session, oauth2):
    # Go through authorization
    access_token = oauth2.get_access_token_authorization_code_grant()
    # Set the token as the header of requests
    requests_session.headers.update({'Authorization': 'Bearer ' + access_token})


def update_client():
    global client
    client = gspread.authorize(creds)


def get_data():

    webbrowser.open_new_tab(config.SHNATON)
    sheet = None
    while sheet is None:
        input("Should we continue?\n")
        for filename in os.listdir():
            if filename.endswith(".xml"):
                from lxml import etree

                with (open(filename, 'r',encoding='utf-8')) as f:
                    doc = etree.parse(f)

                namespaces = {'o': 'urn:schemas-microsoft-com:office:office',
                              'x': 'urn:schemas-microsoft-com:office:excel',
                              'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

                L = []
                ws = doc.xpath('/ss:Workbook/ss:Worksheet', namespaces=namespaces)
                if len(ws) > 0:
                    tables = ws[0].xpath('./ss:Table', namespaces=namespaces)
                    if len(tables) > 0:
                        rows = tables[0].xpath('./ss:Row', namespaces=namespaces)
                        for row in rows:
                            tmp = []
                            cells = row.xpath('./ss:Cell/ss:Data', namespaces=namespaces)
                            for cell in cells:
                                #                print(cell.text);
                                tmp.append(cell.text)
                            L.append(tmp)
                sheet = pd.DataFrame(L)
                sheet.columns = sheet.iloc[0]
                print("Done")
        if sheet is None:
            print("No xml file found, please try again")
    data = sheet[["MO_From", "MO_To", "HA_Name", "GR_CO_id"]]
    data = data.iloc[1:]
    data[["MO_From","MO_To"]] = data[["MO_From","MO_To"]].astype(int)
    data = data.sort_values(["MO_From", "MO_To", "HA_Name", "GR_CO_id"], ascending=True)
    data.reset_index(drop=True, inplace=True)
    return data


def maintain(requests_session, data):
    for i in range(8, 20):  # FROM 8:00 to 20:00
        current_data = data.loc[(data["MO_From"] <= i) & (data["MO_To"] > i)]  # Select lectures that occurs now
        current_data.reset_index(drop=True, inplace=True)
        while True:
            current_time = datetime.now()
            print("Waiting for hour check " + str(current_time))
            if i < current_time.hour:  # adjusting to current hour
                break
            if current_time.minute == 0:
                time.sleep(15)
                parse_and_check(requests_session, current_data)
                time.sleep(60)
                break
            time.sleep(60 - current_time.second)  # waiting between checks
    print("END OF SERVICE - have a nice day")


def parse_and_check(requests_session, data):
    result_df = data.drop_duplicates(subset=["HA_Name"], keep='first')
    result_df = result_df[result_df["HA_Name"].isin(config.SERVERS.keys())]
    check_if_servers_record(requests_session, result_df["HA_Name"])


def check_if_servers_record(requests_session, remote_records):
    non_working_servers = []
    for remote_recorder in remote_records:
        recorder = config.SERVERS[remote_recorder]
        url = config.BASE_URL + "remoteRecorders/search?searchQuery={0}".format(quote(recorder))
        print('Calling GET {0}'.format(url))
        resp = requests_session.get(url=url).json()
        if resp and resp["Results"] and resp["Results"][0] and resp["Results"][0]["State"] == 2:
            print("@@@@@@@@@@@@@@ IS RECORDING @@@@@@@@@@@@@@@")
        else:
            print("REMOTE DOESNT RECORD: " + remote_recorder)
            non_working_servers.append(remote_recorder)
    if len(non_working_servers) != 0:
        send_mail("Servers to check" + str(datetime.now()),
                  "Update on servers time: " + str(datetime.now()) + "\nList of servers to check: " + str(
                      non_working_servers))
    else:
        send_mail("All servers record properly", "")


def main():
    global folders, client, creds, requests_session
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(config.GOOGLE_JSON, scope)
    client = gspread.authorize(creds)
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    requests_session = requests.Session()
    # Load OAuth2 logic
    oauth2 = PanoptoOAuth2(config.PANOPTO_SERVER_NAME, config.PANOPTO_CLIEND_ID, config.PANOPTO_SECRET, False)
    authorization(requests_session, oauth2)
    data = get_data()
    maintain(requests_session, data)


if __name__ == '__main__':
    main()
