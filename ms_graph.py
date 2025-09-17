import os
import sys
import threading
import webbrowser
import msal
import requests
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from dotenv import load_dotenv

MS_GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0'
auth_code = None

# ---------------- AUTH ----------------
class OAuthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        query = urlparse(self.path).query
        params = parse_qs(query)
        if 'code' in params:
            auth_code = params['code'][0]
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'Authorization code received! You can close this window.')
        else:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b'No code found in URL.')

def start_local_server():
    server_address = ('', 8000)
    httpd = HTTPServer(server_address, OAuthHandler)
    print("Listening on port 8000 for redirect...")
    httpd.handle_request()
    httpd.server_close()

def get_access_token(application_id, client_secret, scopes):
    global auth_code
    client = msal.ConfidentialClientApplication(
        client_id=application_id,
        client_credential=client_secret,
        authority='https://login.microsoftonline.com/consumers/'
    )

    if os.path.exists('refresh_token.txt'):
        with open('refresh_token.txt', 'r') as file:
            refresh_token = file.read().strip()
        token_response = client.acquire_token_by_refresh_token(refresh_token, scopes=scopes)
        if "access_token" in token_response:
            return token_response['access_token']

    server_thread = threading.Thread(target=start_local_server, daemon=True)
    server_thread.start()

    auth_url = client.get_authorization_request_url(scopes, redirect_uri='http://localhost:8000')
    webbrowser.open(auth_url)
    print("Please complete login in browser...")

    while auth_code is None:
        pass

    token_response = client.acquire_token_by_authorization_code(
        code=auth_code,
        scopes=scopes,
        redirect_uri='http://localhost:8000'
    )

    if 'refresh_token' in token_response:
        with open('refresh_token.txt', 'w') as file:
            file.write(token_response['refresh_token'])

    if "access_token" in token_response:
        return token_response['access_token']
    else:
        print("Error obtaining token:", token_response)
        return None

# ---------------- FUNCTIONS ----------------
def fetch_inbox_emails(access_token, top=10):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(f"{MS_GRAPH_BASE_URL}/me/mailFolders/Inbox/messages?$top={top}", headers=headers)
    if response.status_code == 200:
        emails = response.json().get('value', [])
        for mail in emails:
            print(From: {mail.get('from', {}).get('emailAddress', {}).get('address','?')}, Subject: {mail.get('subject','No Subject')}")
    else:
        print("Error fetching emails:", response.text)

def fetch_all_emails(access_token, top=10):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(f"{MS_GRAPH_BASE_URL}/me/messages?$top={top}", headers=headers)
    if response.status_code == 200:
        emails = response.json().get('value', [])
        for mail in emails:
            print(From: {mail.get('from', {}).get('emailAddress', {}).get('address','?')}, Subject: {mail.get('subject','No Subject')}")
    else:
        print("Error fetching all emails:", response.text)

def search_emails(access_token, query, top=10):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(f"{MS_GRAPH_BASE_URL}/me/messages?$search=\"{query}\"&$top={top}", headers=headers)
    if response.status_code == 200:
        emails = response.json().get('value', [])
        for mail in emails:
            print(From: {mail['from']['emailAddress']['address']}, Subject: {mail['subject']}")
    else:
        print("Error searching emails:", response.text)

def send_email(access_token, to_email, subject, body, attachment_path=None):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": to_email}}]
    }

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            content_bytes = f.read()
        import base64
        attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(attachment_path),
            "contentBytes": base64.b64encode(content_bytes).decode("utf-8")
        }
        message["attachments"] = [attachment]

    data = {"message": message, "saveToSentItems": "true"}
    response = requests.post(f"{MS_GRAPH_BASE_URL}/me/sendMail", headers=headers, json=data)
    if response.status_code == 202:
        print("Email sent successfully!")
    else:
        print("Error sending email:", response.text)

def download_attachments(access_token, message_id, save_dir="attachments"):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(f"{MS_GRAPH_BASE_URL}/me/messages/{message_id}/attachments", headers=headers)
    if response.status_code == 200:
        attachments = response.json().get('value', [])
        os.makedirs(save_dir, exist_ok=True)
        for att in attachments:
            if att['@odata.type'] == "#microsoft.graph.fileAttachment":
                filename = os.path.join(save_dir, att['name'])
                with open(filename, "wb") as f:
                    import base64
                    f.write(base64.b64decode(att['contentBytes']))
                print(f"Downloaded: {filename}")
    else:
        print("Error downloading attachments:", response.text)

def create_folder(access_token, folder_name):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    data = {"displayName": folder_name}
    response = requests.post(f"{MS_GRAPH_BASE_URL}/me/mailFolders", headers=headers, json=data)
    if response.status_code == 201:
        print(f"Folder '{folder_name}' created!")
    else:
        print("Error creating folder:", response.text)

def reply_email(access_token, message_id, body):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    data = {"message": {"body": {"contentType": "Text", "content": body}}}
    response = requests.post(f"{MS_GRAPH_BASE_URL}/me/messages/{message_id}/createReply", headers=headers, json=data)
    if response.status_code == 201:
        reply_id = response.json()['id']
        requests.post(f"{MS_GRAPH_BASE_URL}/me/messages/{reply_id}/send", headers=headers)
        print("Reply sent!")
    else:
        print("Error replying:", response.text)

def create_draft(access_token, to_email, subject, body):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    data = {"subject": subject, "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to_email}}]}
    response = requests.post(f"{MS_GRAPH_BASE_URL}/me/messages", headers=headers, json=data)
    if response.status_code == 201:
        print(f"Draft created: {response.json()['id']}")
    else:
        print("Error creating draft:", response.text)

def send_draft(access_token, draft_id):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.post(f"{MS_GRAPH_BASE_URL}/me/messages/{draft_id}/send", headers=headers)
    if response.status_code == 202:
        print("Draft sent successfully!")
    else:
        print("Error sending draft:", response.text)

def delete_email(access_token, message_id):
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.delete(f"{MS_GRAPH_BASE_URL}/me/messages/{message_id}", headers=headers)
    if response.status_code == 204:
        print("Email deleted!")
    else:
        print("Error deleting email:", response.text)

def move_email(access_token, message_id, folder_id):
    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
    data = {"destinationId": folder_id}
    response = requests.post(f"{MS_GRAPH_BASE_URL}/me/messages/{message_id}/move", headers=headers, json=data)
    if response.status_code == 201:
        print(f"Email moved to folder {folder_id}")
    else:
        print("Error moving email:", response.text)

# ---------------- MAIN ----------------
def main():
    load_dotenv()
    APPLICATION_ID = os.getenv('APPLICATION_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    SCOPES = ['User.Read', 'Mail.ReadWrite', 'Mail.Send']

    access_token = get_access_token(APPLICATION_ID, CLIENT_SECRET, SCOPES)
    if not access_token:
        print("Failed to get token")
        return

    if len(sys.argv) < 2:
        print("Usage: python ms_graph.py <command> [args]")
        return

    cmd = sys.argv[1]

    if cmd == "inbox":
        fetch_inbox_emails(access_token, top=10)
    elif cmd == "all":
        fetch_all_emails(access_token, top=10)
    elif cmd == "search":
        search_emails(access_token, sys.argv[2])
    elif cmd == "send":
        send_email(access_token, sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "send_attach":
        send_email(access_token, sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    elif cmd == "download_attach":
        download_attachments(access_token, sys.argv[2])
    elif cmd == "create_folder":
        create_folder(access_token, sys.argv[2])
    elif cmd == "reply":
        reply_email(access_token, sys.argv[2], sys.argv[3])
    elif cmd == "draft":
        create_draft(access_token, sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "send_draft":
        send_draft(access_token, sys.argv[2])
    elif cmd == "delete":
        delete_email(access_token, sys.argv[2])
    elif cmd == "move":
        move_email(access_token, sys.argv[2], sys.argv[3])
    else:
        print("Unknown command")

if __name__ == "__main__":
    main()
