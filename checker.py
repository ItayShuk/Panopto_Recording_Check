import gspread
from oauth2client.service_account import ServiceAccountCredentials
from panopto_folders import PanoptoFolders
from panopto_oauth2 import PanoptoOAuth2
import urllib3
import requests
import json
from datetime import datetime, timedelta
from urllib.parse import quote
from dateutil import parser, rrule
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import config
import socket
import sessions
from sessions import PanoptoSessions
import argparse
import win32com.client
import pandas as pd
import xml.etree.ElementTree as ETree

outlook = win32com.client.Dispatch("Outlook.Application")


# python to EXE:
# https://towardsdatascience.com/how-to-easily-convert-a-python-script-to-an-executable-file-exe-4966e253c7e9


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


def get_remote_recorders():
    tree = ETree.parse(r"C:\Users\Itayshu\Desktop\Panopto_Recording_Check\sample.xml")
    # get the root of the tree
    root = tree.getroot()

    # return the DataFrame


    return ["קפלן"]


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
    recorders = get_remote_recorders()
    # check_if_servers_record(requests_session,recorders)

    # schedule.every(1).hours.do(update_client)
    # schedule.every(1).hours.do(authorization, requests_session, oauth2)
    # schedule.every(5).seconds.do(while_waiting)
    # while True:
    #     schedule.run_pending()


def check_if_servers_record(requests_session, remote_records):
    resp_list = []
    for remote_recorder in remote_records:
        recorder = config.SERVERS[remote_recorder]
        url = config.BASE_URL + "remoteRecorders/search?searchQuery={0}".format(quote(recorder))
        print('Calling GET {0}'.format(url))
        resp_list.append(requests_session.get(url=url).json())
    for resp in resp_list:
        if resp["Results"][0]["State"] != 2:
            print("REMOTE DOESNT RECORD")
        print("IS RECORDING")


if __name__ == '__main__':
    main()
