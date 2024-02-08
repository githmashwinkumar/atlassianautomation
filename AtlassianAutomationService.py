import time
import random
from pathlib import Path

import win32api
import win32con
import win32evtlog
import win32evtlogutil
import win32security

from SMWinservice import SMWinservice
import logging

import mimetypes
import requests
import json
from requests.auth import HTTPBasicAuth
import time
from datetime import datetime

class AtlassianAutomationService(SMWinservice):
    _svc_name_ = "AtlassianAutomationService"
    _svc_display_name_ = "XXX Atlassian Automation Service"
    _svc_description_ = "Service for automating processes in Atlassian"

    def start(self):
        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def removeAttachments(self):
        msg = ""
        keyName = "EPS"
        url = "https://XXX/wiki/rest/api/space/" + keyName
        auth = HTTPBasicAuth('', '')
        response = requests.request("GET", url, auth=auth)
        try:
            if (response.status_code != 404):
                space = response.json()
                expandable = space['_expandable']
                homePage = expandable['homepage']
                homePageID = homePage[-9:]
                url1 = 'https://XXX/wiki/rest/api/content/' + str(homePageID) + '/child/page/'
                auth1 = HTTPBasicAuth('', '')
                headers1 = {"Accept": "application/json"}
                response1 = requests.request("GET", url1, headers=headers1)
                res1 = json.loads(response1.text)
                results = res1['results']
                for page in results:
                    pageID = page['id']
                    now = datetime.now()
                    url2 = 'https://XXX/wiki/rest/api/content/search?&cql=type=attachment and space=' + keyName + ' and created <= now("-24h") '
                    auth2 = HTTPBasicAuth('', '')
                    headers2 = {"Accept": "application/json"}
                    response2 = requests.request("GET", url2, headers=headers2)
                    res2 = json.loads(response2.text)
                    try:
                        results2 = res2['results']
                        i = 0
                        limit = 5

                        while i < limit:
                            # print( str(i) + " - " + str(results2[i]['id']) + " - " + str(results2[i]['title']))
                            msg = results2[i]['title']
                            id = str(results2[i]['id'])
                            delURL = "https://XXX/wiki/rest/api/content/" + id
                            auth2 = HTTPBasicAuth('', '')
                            delresponse = requests.request( "DELETE",delURL,auth=auth2 )
                            i = i + 1
                            descr = ["Attachment with this ID wsa removed - " + msg]
                            data = "Application\0Data".encode("ascii")
                            ph = win32api.GetCurrentProcess()
                            th = win32security.OpenProcessToken(ph, win32con.TOKEN_READ)
                            win32evtlogutil.ReportEvent("AtlassianAutomationService", 4098, eventCategory=5,
                                                        eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE, strings=descr, data=data, sid=win32security.GetTokenInformation(th, win32security.TokenUser)[0])
                    except:
                        x = 1
        except:
            y = 1

    def main(self):
        i = 0
        while self.isrunning:
            random.seed()
            x = random.randint(1, 1000000)
            ph = win32api.GetCurrentProcess()
            th = win32security.OpenProcessToken(ph, win32con.TOKEN_READ)
            my_sid = win32security.GetTokenInformation(th, win32security.TokenUser)[0]

            applicationName = "AtlassianAutomationService"
            eventID = 4098
            category = 5  # Shell
            myType = win32evtlog.EVENTLOG_INFORMATION_TYPE

            try:
                self.removeAttachments()
            except:
                i = 0

            time.sleep(5)


if __name__ == '__main__':
    AtlassianAutomationService.parse_command_line()