import os
import time
import pickle
from openai import OpenAI
from openai import RateLimitError
import gspread
from PIL import Image
from io import BytesIO
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
''''For future: Revise batch. 
                Analyze unread emails through GMAIL API using keywords to lower the volume of request for batch
                Search for "not selected" / "not moving forward" etc - REJECTION
                Search for "application submitted" etc - PENDING
                Search for "happy blah blah" - Invite to Interview
                                                                                                                '''


# Set your OpenAI API key 'YOUR_API_KEY'
OPENAI_API_KEY = 'YOUR_API_KEY'
client = OpenAI(api_key=OPENAI_API_KEY)
# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']

def authenticate_gmail():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('gmail', 'v1', credentials=creds)
    return service

def authenticate_google_sheets():
    creds = None
    if os.path.exists('token_sheets.pickle'):
        with open('token_sheets.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token_sheets.pickle', 'wb') as token:
            pickle.dump(creds, token)
    gc = gspread.authorize(creds)
    return gc

def get_emails(service):
    results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
    messages = results.get('messages', [])
    
    email_data = []
    for msg in messages:
        msg = service.users().messages().get(userId='me', id=msg['id']).execute()
        headers = msg['payload']['headers']
        subject = next(header['value'] for header in headers if header['name'] == 'Subject')
        date = next(header['value'] for header in headers if header['name'] == 'Date')
        email_body = msg['snippet']
        
        email_data.append({
            'subject': subject,
            'date': date,
            'body': email_body
        })
    return email_data

def analyze_email_with_gpt(subject, body):
    prompt = f"""
    I have received the following email with the subject "{subject}":
    {body}
    
    Please analyze this email and tell me the following:
    1. Is it a job application confirmation or an update on a job application?
    2. If it's a confirmation, extract the position name, company name, and the date of application submission.
    3. If it's an update, determine whether it's a rejection or an invite to an interview.

    Your response should follow this format:
    Confirmation/Rejection/Invite to Interview; Position Name; Company Name; Date of Application(DD/MM/YYYY)
    """
    #time.sleep(5)  
    response = client.chat.completions.create(
        model="gpt-4o-mini",   
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt},
        ],
        max_tokens=240
    )
    
    return response.choices[0].message['content'].strip()

def process_response(response):
    parts = response.split(';')
    if len(parts) == 4:
        status = parts[0].strip()
        position = parts[1].strip()
        company = parts[2].strip()
        date = parts[3].strip()
        return {
            "status": status,
            "position": position,
            "company": company,
            "date": date
        }
    else:
        return {"error": "Unexpected response format"}
    
def extract_information(email_data, worksheet):
    extracted_data = []
    for email in email_data:
        subject = email['subject']
        body = email['body']
        
        rawAnalysis = analyze_email_with_gpt(subject, body)
        analysis = process_response(rawAnalysis)
        if "Confirmation" in analysis['status']:
            # Parse confirmation details from GPT response
            position = analysis.get('position', 'N/A')
            company = analysis.get('company', 'N/A')
            date_submitted = analysis.get('date', 'N/A')
            extracted_data.append([position, company, date_submitted, "pending"])
        
        elif "Invite to Interview" in analysis['status']:
            # Parse confirmation details from GPT response
            position = analysis.get('position', 'N/A')
            company = analysis.get('company', 'N/A')
            update_application_status(worksheet, position, company,  1)
        
        elif "Rejection" in analysis['status']:
            # Parse confirmation details from GPT response
            position = analysis.get('position', 'N/A')
            company = analysis.get('company', 'N/A')
            update_application_status(worksheet, position, company, -1)
    
    return extracted_data

def update_application_status(worksheet, position_name, company_name, status_flag): 
    # status_flag = 1 if invite else rejection
    # Assume the position name is in column 1 and company name is in column 2
    position_col = 1
    company_col = 2
    status_col = 4

    # Iterate through rows to find a match for both position name and company name
    cell_to_update = None
    for row in worksheet.get_all_values():
        if row[position_col - 1] == position_name and row[company_col - 1] == company_name:
            cell_to_update = row
            break
    
    if cell_to_update:
        # Find the row number of the cell_to_update
        row_num = worksheet.find(cell_to_update[0]).row
        if status_flag == 1:
            worksheet.update_cell(row_num, status_col, 'Invite to Interview')
        else:          
            worksheet.update_cell(row_num, status_col, 'rejected')
    else:
        print(f"No matching row found for position '{position_name}' and company '{company_name}'")

def generate_dalle_image(num_applications, area_of_jobs):
    prompt = f"Generate an image of a person who applied to {num_applications} applications in {area_of_jobs} - make it as crazy and absurd as you want."

    response = client.images.generate(
        model="dall-e-3",
        prompt=prompt,
        n=1,
        size="1024x1024",
        quality="standard",
    )
    
    image_url = response.data[0].url
    
    return image_url

def add_image_to_sheet(gc, spreadsheet_id, image_url):
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.sheet1
    
    # Add the image URL to the end of the Google Sheet
    last_row = len(worksheet.get_all_values()) + 1
    worksheet.update_cell(last_row, 1, "Generated Image:")
    worksheet.update_cell(last_row, 2, image_url)

def store_in_google_sheets(gc, extracted_data, spreadsheet_id, sheet_name):
    sh = gc.open_by_key(spreadsheet_id)
    worksheet = sh.worksheet(sheet_name)
    
    for row in extracted_data:
        if row[0] is not None:  # Skip entries that aren't new submissions
            worksheet.append_row(row, value_input_option='RAW')

    # Generate image after storing all the data
    num_applications = len(worksheet.get_all_values()) - 1
    area_of_jobs = extracted_data[0][0] if extracted_data else "various fields"
    image_url = generate_dalle_image(num_applications, area_of_jobs)

    # Add the generated image to the sheet
    add_image_to_sheet(gc, spreadsheet_id, image_url)

def main():
    
    service = authenticate_gmail()
    gc = authenticate_google_sheets()

    email_data = get_emails(service)
    
    # Replace with your actual Spreadsheet ID and Sheet name
    spreadsheet_id = '1kYTfc94xGKN79UjL_Sg3S8GLyYaQkIqdPyhPlSgV2m8'
    sheet_name = 'Internships'
    
    sh = gc.open_by_key(spreadsheet_id)
    worksheets = sh.worksheets()
    # Print all worksheet names to confirm the correct names
    print("Available worksheets:")
    for ws in worksheets:
        print(ws.title)
    #worksheet = sh.worksheet(sheet_name)
    worksheet = gc.open("Internships").sheet1
    extracted_data = extract_information(email_data, worksheet)
    
    store_in_google_sheets(gc, extracted_data, spreadsheet_id, sheet_name)

if __name__ == '__main__':
    main()
