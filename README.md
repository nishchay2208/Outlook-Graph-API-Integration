# Outlook Graph API Integration in Python

A full-featured Python project to integrate Outlook with Microsoft Graph API.  
This project demonstrates authentication, retrieving emails, searching, sending (with attachments), downloading attachments, folder management, replying, drafts, deleting, and moving emails between folders using Microsoft Graph API.

---

## Features

1. **User Authentication**
   - Authenticate Microsoft/Outlook users via OAuth2.
   - Supports `refresh_token` for persistent sessions.

2. **Email Retrieval**
   - Fetch latest emails from Inbox or any folder.
   - Fetch emails with filters, e.g., unread, specific sender, or subject search.

3. **Email Sending**
   - Send emails to any recipient.
   - Support for sending attachments.
   - Draft emails and send them later.

4. **Email Management**
   - Reply to emails.
   - Delete emails.
   - Move emails to specific folders.
   - Create new folders in Outlook.

5. **Attachment Handling**
   - Download attachments from emails.
   - Support for multiple attachments.

---

## Prerequisites

- Python 3.8 or higher
- Git installed
- Microsoft 365 account
- Registered application in Azure Portal with:
  - `APPLICATION_ID` (Client ID)
  - `CLIENT_SECRET` (Client Secret)
  - `TENANT_ID`
  - Required Graph API permissions:
    - `User.Read`
    - `Mail.ReadWrite`
    - `Mail.Send`

---

## Setup Instructions

1. **Clone the repository**
```bash
git clone https://github.com/nishchay2208/Outlook-Graph-API-Integration.git
cd Outlook-Graph-API-Integration
