# Email Parser for Job Applications
## Problem
![problem](https://i.imgur.com/rE5qa24.jpeg)

When I was trying to land an internship this year, I wanted to keep track of applications so I could plot a fancy graph like people from *r/CS* and say:
> 'Hey, took me just a shy thousand applications to get there!'

The process of applying, copying information, and putting it inside the table is definitely not a high ROI activity. Sometimes it can be depressing since most emails are just "Well, you're great, but not great enough" to the point that I don't want to fill out this table, and instead just focus on applying as much as possible, as early as possible.

What's why I implemented this parser

## Supported Platforms
Currently script is capable of only reading Gmail letters, and putting/updating information from Unread messages inside the Google Sheets. If this will get a huge interest, I will consider adding support to other platforms like Outlook.
- Gmail
- Google Sheets
## Requirements for Python
```
pip3 install openai
pip3 install Pillow
pip3 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib gspread nltk
```

## How To Run
## Step 1: Set Up a Google Cloud Project
Go to the Google Cloud Console
Make sure you're logged in with your Google account.
Create a New Project:
- Click on the project dropdown at the top of the page.
- Click "New Project."
- Give your project a name and click "Create."

Enable the APIs:
- Once your project is created, navigate to the "API & Services" dashboard.
- Click on "Enable APIs and Services."
- Search for "Google Sheets API" and click "Enable."
- Search for "Gmail API" and click "Enable."
  
## Step 2: Create OAuth 2.0 Credentials
Create Credentials:
- Go to the "Credentials" tab in the "API & Services" dashboard.
- Click "Create Credentials" and select "OAuth 2.0 Client IDs."

Configure the Consent Screen:
- Choose "External" for user type if prompted.
- Ensure that your app is set to "Testing" mode and your email is in the list!
- No need to fill out scope information.
- Fill in the necessary fields (app name, email, etc.).
- Click "Save and Continue" (you can skip the Scopes and Test Users steps for now).

Create OAuth Client ID:
- After configuring the consent screen, choose "Application type" as "Desktop app."
- Give it a name (e.g., "Python Parser").
- Click "Create."

Download the credentials.json File:
- Once the client ID is created, you'll see a "Download" button.
- Click the "Download" button to download the credentials.json file.
- Save this file in the same directory as your Python script and rename it 'credentials.json'.

  
## Step 3: Run the Script
> [!WARNING]
> Script right now works only with preprocessed emails (max 800) that were appended to one email and sent as 1 request to ChatGPT 4o-mini. I'm in the process of improving this step to make it more accessible 

Make sure you created a spreadsheet in Google Sheets or imported one - replace the spreadsheet ID and name with yours. It has to follow this format:
```
Job Name | Company | Date | Status(rejected,pending,invite to interview)
```

The first time you run your script, it will prompt you to log in via a browser window.
After logging in, the script will generate a token.pickle or token_sheets.pickle file, which stores your credentials for future use.
After updating the spreadsheet, the GPT will also generate an image of you after applying to X number of position for Y job(-s)!
