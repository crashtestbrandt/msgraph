import boto3
import base64
import json
import os
from msgraph import MSGraph

BUCKET_NAME = '[BUCKET WITH MSGRAPH TOKEN FILE, e.g "tokens-bucket"]'
TOKEN_FILE = 'tokens.json'
FILESTUB = '/tmp'
SECRET_NAME = '[NAME OF O365 CREDS IN SECRETSMANAGER]'
REGION_NAME = '[REGION NAME, e.g. "us-east-1"]'

def get_secret(_secret):

    secret_name = _secret
    region_name = REGION_NAME

    # Create a Secrets Manager client
    session = boto3.session.Session()
    client = session.client(
        service_name='secretsmanager',
        region_name=region_name
    )
    get_secret_value_response = client.get_secret_value(
        SecretId=secret_name
    )
    if 'SecretString' in get_secret_value_response:
        secret = get_secret_value_response['SecretString']
        return secret
    else:
        decoded_binary_secret = base64.b64decode(get_secret_value_response['SecretBinary'])
        return decoded_binary_secret

def get_graph_token():
    KEY = TOKEN_FILE
    s3 = boto3.resource('s3')
    try:
        token_path = os.path.join(FILESTUB, TOKEN_FILE)
        #token_path = FILESTUB + TOKEN_FILE
        s3.Bucket(BUCKET_NAME).download_file(KEY, os.path.join(FILESTUB, TOKEN_FILE))
    except:
        print("ERROR: Could not get token file.")
        raise
    return token_path


def lambda_handler(event, context):
    token_path = get_graph_token()
    creds = json.loads(get_secret(SECRET_NAME))
    msgraph = MSGraph(token_path, creds["client_id"], creds["client_secret"])
    message_tree = msgraph.get_top_message_tree()

    return {
        "message_tree" : message_tree
    }
