from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

import os

class Auth:
    def __init__(self, SCOPES):
        self.SCOPES = SCOPES
            
    def get_credentials(self):
        store = file.Storage(os.path.join(os.path.dirname(__file__).replace("/","\\"),"credentials.json"))
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(os.path.join(os.path.dirname(__file__).replace("/","\\"),"client_secret.json"), self.SCOPES)
            creds = tools.run_flow(flow, store)
        return creds
