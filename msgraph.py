import requests
import json
import sys
from zipfile import ZipFile
import io
import base64
import re
import mimetypes
import os
import hashlib
import urllib.request
import email
from email.policy import default
import boto3

LOCAL_DEV = False

FILE_BUCKET = 'NAME OF BUCKET FOR OUTPUT FILES, e.g "file-bucket"'

AUTHORIZE_URL = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize"
TOKEN_URL = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
CALLBACK_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"

MESSAGE_URL = "https://graph.microsoft.com/beta/me/messages"
INBOX_URL = "https://graph.microsoft.com/beta/me/mailFolders/Inbox/messages"
MY_GROUPS_URL = "https://graph.microsoft.com/beta/me/memberOf"

class MSGraph:
    client_id = None
    client_secret = None
    token_path = None
    access_token = None
    refresh_token = None
    api_call_headers = {}
    filestub = None
    bucket_name = FILE_BUCKET

    def __init__(self, token_path, client_id = None, client_secret = None, token_bucket = None, tokens_key = 'tokens.json'):
        self.client_id = client_id
        self.client_secret = client_secret
        self.token_path = token_path
        if LOCAL_DEV:
            self.filestub = "./tmp"
        else:
            self.filestub = "/tmp"
        # Try to open tokens file
        try:
            tokens = {}
            with open(self.token_path, 'r') as tokens_file:
                tokens = json.load(tokens_file)
        except FileNotFoundError:
            # No token file; request initial token
            if LOCAL_DEV:
                # Print instructions for manual authentication and token request
                if self.client_id and self.client_secret:
                    authorization_redirect_url = AUTHORIZE_URL + '?response_type=code&client_id=' + self.client_id + '&redirect_uri=' + CALLBACK_URI + '&scope=openid offline_access'
                    print("---  " + authorization_redirect_url + "  ---")
                    authorization_code = input('code: ')
                    data = {'grant_type': 'authorization_code', 'code': authorization_code, 'redirect_uri': CALLBACK_URI}
                    print("requesting access token")
                    access_token_response = requests.post(TOKEN_URL, data=data, allow_redirects=False, auth=(self.client_id, self.client_secret))
                    tokens = json.loads(access_token_response.text)
                else:
                    print("ERROR: Client ID and Client Secret required.")  
            else:
                # Cloud environment
                print("\nERROR: No token file.\n") # No console, so no output # TODO: Make auth request and handle response?
                exit(1)
        
        # By now we should have a tokens object
        try:
            self.access_token = tokens['access_token']
            self.refresh_token = tokens['refresh_token']
        except KeyError:
            print("\nERROR: Token file is missing a token.\n") # TODO: Which one? Handle it
            exit(1)
        
        # Perform authentication with refresh token (even if we got an initial token before) and get updated tokens
        data = {'grant_type': 'refresh_token', 'refresh_token': self.refresh_token, 'redirect_uri': CALLBACK_URI}
        access_token_response = requests.post(TOKEN_URL, data=data, allow_redirects=False, auth=(self.client_id, self.client_secret))
        tokens = json.loads(access_token_response.text)
        try:
            self.access_token = tokens['access_token']
            self.refresh_token = tokens['refresh_token']
        except:
            print("\nERROR: Request to update tokens resulted in bad response:")
            print("\n" + json.dumps(tokens) + "\n")
            exit(1)

        # Save the updated token
        with open(self.token_path, 'w') as tokens_file:
            json.dump(tokens, tokens_file)
        if not LOCAL_DEV:
            s3 = boto3.client('s3')
            s3.upload_file(self.token_path, token_bucket, tokens_key) 
        
        # Set headers for API calls
        self.api_call_headers = {'Authorization': 'Bearer ' + self.access_token}
    
    def extract_zip(self, filekey):
        f = open(os.path.join(self.filestub, filekey), "rb")
        with ZipFile(f, 'r') as inZip:
            inZip.setpassword(b'[PASSWORD]')    # TODO: If you know a password to try
            inZip.extractall(path=self.filestub)
        f.close()
        items = []
        for item in inZip.infolist():
            # get filename from zip archive
            result = re.search(r'filename=(?:\"|\')(.*)(?:\"|\')\s+compress_type', str(item))
            if result:
                filename = result.group(1)
                with open(os.path.join(self.filestub, filename), "rb") as infile:
                    content = infile.read()
                    hasher = hashlib.md5(content)
                    checksum = hasher.hexdigest()
                    with open(os.path.join(self.filestub, checksum), 'wb') as fp:
                        fp.write(content)
                    os.remove(os.path.join(self.filestub, filename))
                    mimetype = mimetypes.guess_type(os.path.join(self.filestub, urllib.request.pathname2url(filename)))[0]
                    if not LOCAL_DEV:
                        s3 = boto3.client('s3')
                        s3.upload_file(os.path.join(self.filestub, checksum), self.bucket_name, checksum)
                    node = {}
                    node["filename"] = filename
                    node["md5"] = checksum
                    node["mimetype"] = mimetype
                    if mimetype == 'application/zip':
                        node["children"] = self.extract_zip(os.path.join(self.filestub, checksum))
                    if mimetype == 'message/rfc822':
                        with open(os.path.join(self.filestub, checksum), 'r') as fp:
                            child_message = email.message_from_file(fp, policy=default)
                            node["metadata"] = {
                                "to" : child_message["to"],
                                "from" : child_message["from"],
                                "subject" : child_message["subject"]
                            }
                            node = self.handle_message(child_message, node)
                    items.append(node)
        return items
    
    def get_last_unread_message_id(self):
        api_url = INBOX_URL + "?$filter=isRead ne true"
        message = requests.get(api_url, headers = self.api_call_headers).json() # TODO: What happens if response doesn't have 'value', or empty value?
        return message['value'][0]['id']
    
    def get_message_by_id(self, _message_id):
        api_url = MESSAGE_URL + "/" + _message_id + "/$value"   # Ask for raw MIME conttent
        message = requests.get(api_url, headers = self.api_call_headers).content
        return message
    
    def handle_message(self, _message, _message_tree):
        message = _message
        message_tree = _message_tree
        children = []
        counter = 1

        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            filename = part.get_filename()
            if not filename:
                ext = mimetypes.guess_extension(part.get_content_type())
                if not ext:
                    ext = '.bin'
                filename = 'part-%03d%s' % (counter, ext)
            node = {
                "filename" : filename,
                "mimetype" : part.get_content_type()
                }
            counter += 1
            content = part.get_content()
            checksum = ""
            if isinstance(content, bytes):
                with io.BytesIO(content) as f:
                    hasher = hashlib.md5(f.read())
                    checksum = hasher.hexdigest()
                    with open(os.path.join(self.filestub, checksum), 'wb') as fp:
                        fp.write(content)
                    if (part.get_content_type() == 'application/zip'):
                        node["children"] = self.extract_zip(checksum)

            elif isinstance(content, str):
                hasher = hashlib.md5(content.encode())
                checksum = hasher.hexdigest()
                fp = open(os.path.join(self.filestub, checksum), 'w')
                fp.write(content)

            elif isinstance(content, email.message.EmailMessage):
                with io.BytesIO(content.as_bytes()) as f:
                    hasher = hashlib.md5(f.read())
                    checksum = hasher.hexdigest()
                    fp = open(os.path.join(self.filestub, checksum), 'wb')
                    fp.write(content.as_bytes())
                node["metadata"] = {}
                node["metadata"]["subject"] = content["subject"]
                node["metadata"]["to"] = content["to"]
                node["metadata"]["from"] = content["from"]

            else:
                print("Something's gone wrong.")
            
            node["md5"] = checksum
            if LOCAL_DEV:
                node["python_type"] = str(type(content))
            else:
                s3 = boto3.client('s3')
                s3.upload_file(os.path.join(self.filestub, checksum), self.bucket_name, checksum)
            children.append(node)
        message_tree["children"] = children
        return message_tree

    def get_top_message_tree(self):
        message_id = self.get_last_unread_message_id()
        raw_message = self.get_message_by_id(message_id)
        try:
            os.mkdir(self.filestub)
        except:
            pass
        with io.BytesIO(raw_message) as f:
            hasher = hashlib.md5(f.read())
            checksum = hasher.hexdigest()
            self.filestub = os.path.join(self.filestub, checksum)
            try:
                os.mkdir(self.filestub)
            except:
                pass
        with open(os.path.join(self.filestub, checksum), 'wb') as f:
            f.write(raw_message)
        
        message_tree = {}
        message_tree["filename"] = checksum
        message_tree["md5"] = checksum
        with open(os.path.join(self.filestub, checksum), 'r') as f:
            message = email.message_from_file(f, policy=default)
            message_tree["mimetype"] = message.get_content_type()
            message_tree["metadata"] = {
                "to" : message["to"],
                "from" : message["from"],
                "subject" : message["subject"]
            }
            message_tree = self.handle_message(message, message_tree)
        self.mark_read(message_id)  # TODO: validate response code
        return message_tree

    def mark_read(self, _message_id):
        api_url = MESSAGE_URL + "/" + _message_id
        self.api_call_headers = {'Authorization': 'Bearer ' + self.access_token, 'Content-Type': 'application/json'}
        response = requests.patch(api_url, json={"isRead":True}, headers=self.api_call_headers)
        return response.status_code

if __name__ == "__main__":
    LOCAL_DEV = True

    TOKEN_PATH = "tokens.json"
    client_id = "[YOUR O365 CLIENT ID]"
    client_secret = "[YOUR O365 CLIENT SECRET]"

    msgraph = MSGraph(TOKEN_PATH, client_id, client_secret)
    message_tree = msgraph.get_top_message_tree()
    
    print(json.dumps(message_tree))
    