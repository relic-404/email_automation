from Google import create_service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import date
import random
import csv
import requests
import os.path
import time
import wootric_api as woo
import client_list_google as cl_google


CLIENT_SECRET_FILE = 'credentials.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']
service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# Generates random number to choose template that will be used
ran_num = random.randint(1, 3)

# create variable for survey data from wootric_api.py
survey__data = woo.rows
date_auto_sent = date.today()
date_auto_sent = date_auto_sent.strftime("%m/%d/%y")

# create variable for client list data from google sheets
cl_google.client_list_data()
client_sheet = cl_google.new_list
client_sheet = client_sheet[0]['values'][1:]


# PUT request variables
wootric_api = woo.api_base_url
headers_test = woo.headers
wootric_api_endpont = woo.endpoint



surveys_being_sent = [] # Empty list to add surveys being sent out

temp_lst = [['FIRM ID:','EMAIL:','SCORE:','DATE AUTO-EMAIL SENT:', 'DATE SURVEY SUBMITTED:', 'Assigned CSM:', 'TEMPLATE USED:', 'TYPE OF SCORE:']]



# loops through survey_data function(wootric_api.py)

def email_automate():
    for i in survey__data[1:]:
        # opens csv to call client list and match emails:
        with open('client_list.csv') as loopfile:
                client_data = csv.DictReader(loopfile, dialect='excel')
                for client in client_data:
                    client_wootric_email=i[4] # Client email that is pulled from survey data

                    if client_wootric_email == client['Email']:
                        print('\nVerified, Emails match correctly', 'EMAIL:', i[4], 'SCORE:', i[1], 'DATE:', i[3], 'FIRM ID', client['ID'])
                        client_temp = client['First_Name']
                        client_capital= client_temp.capitalize() # capitalize the first letter
                        for data in client_sheet:
                            firm_id_goog_sheet = data[0]
                            if firm_id_goog_sheet == client['ID']: # If firm ids match from client list and from survey data AND the csm on the google sheet is not Constance, send emails
                                #if data[8] != 'Constance':
                                    print('IDS MATCH!', data[0], client['ID'])
                                    print('CSM for client:', data[8])
                                    ran_num = random.randint(1, 3)
                                    score=int(i[1]) # The client submitted score
                                    survey_text_none = i[2] # Comments on survey
                                    # Promoter Templates
                                    promoter_temp_one=f"Hi {client_capital}," f"""\n\nThank you for taking the time to respond to our recent survey asking your likelihood to recommend us to others. \n\nYou scored us "{score}" out of 10.
                                    \nWe appreciate your feedback and wanted to reach out to follow up. Would you be able to provide any additional feedback? We would like to know what features you have been enjoying and the reasoning behind your score. This helps identify what we are doing right and continue to ensure that we provide the most for our clients!
                                    \nFurthermore, we can discuss any other questions or concerns that you may have. We could also schedule a call to discuss this feedback. If so, let me know and I will send over a link that will allow us to schedule a call.
                                    \nPlease let us know if you do have any questions, we will be happy to assist. 
                                    \nThank you!"""
                                    promoter_temp_two=f"Hi {client_capital}," f"""\n\nThank you for taking the time to submit your survey response and giving us a positive score of {score} out of 10!
                                    \nWe are always trying to improve the software for our client's so we would like to better understand your feedback from a user's perspective. This feedback will be invaluable and will provide insight into our client's such as understanding what we can do better and what we are currently doing right. 
                                    \nIf you have availability, we would love to schedule a time to chat with you! Please let us know and we can provide a Calendly link which will allow you to select a time that best fits your availability. We can use this call to discuss your feedback, questions, or any concerns. However, if you would not like to schedule a call. Any information that you can provide in response to this email will be greatly appreciated!
                                    \nThank you!"""
                                    promoter_temp_three=f"Hi {client_capital}," f"""\n\nThank you for submitting a survey response and giving us a positive score of “{score}” out of 10!
                                    \nYour feedback is essential to help us continue to better the product for our clients. If you could please provide any additional feedback, we would greatly appreciate it!
                                    \nOn top of that, your feedback helps us identify areas of improvement and what we are doing right. If you are interested in scheduling some time to discuss any feedback, please let us know and we can send over a Calendly link that will allow you to schedule a time of your choosing. In addition, we can use this time to go over any questions or concerns that you may have.
                                    \nIf this is something you would like to discuss, please let us know!
                                    \nThank you!"""
                                    

                                    # Passive Templates
                                    passives_temp_one=f"Hi {client_capital}," f"""\n\nThank you for submitting a survey response and giving us a positive score of “{score}” out of 10!
                                    \nWe greatly appreciate you taking the time to submit your response and we would like to know more regarding your submitted score. This helps us identify what areas we can improve on and what we're also doing right.
                                    \nIf you would like to schedule a call with us, please let us know and we can send a Calendly link that will let you schedule a time. We can use this meeting to address your feedback and any questions or concerns that you may have.
                                    \nThank you!"""
                                    passives_temp_two=f"Hi {client_capital}," f"""\n\nThank you for submitting a survey response and giving us a positive score of “{score}” out of 10!
                                    \nYour feedback is essential to helping us better the software for our users. Would you be willing to schedule a call with us so we can gather any additional feedback that you may have?
                                    \nFurthermore, we could use the meeting to address any questions or concerns that you may have. If this is something you are interested in, please let us know and we can send a link that will allow you to schedule some time with us.
                                    \nThank you!"""
                                    passives_temp_three=f"Hi {client_capital}," f"""\n\nThank you for submitting a survey response and giving us a score of “{score}” out of 10!
                                    \nWe appreciate you taking the time to submit a response and would like to gather additional information. This will be beneficial to help us identify what we are doing right and what we can improve on as well.
                                    \nIf you would like, we can schedule a call as well. Please let us know and we will send over a Calendly link that will let you choose a time that works best for you.
                                    \nThank you!"""


                                    # Detractor Templates
                                    detractor_temp_one=f"{client_capital}," f"""\n\nThank you for taking the time to submit your survey response.
                                    \nWe received a score of {score} out of 10.
                                    \nTaking this score into mind, we understand that the likelihood of you recommending us to others would be least likely. We would like to understand the reasoning behind your score and gather any additional feedback that you may have. We want to ensure we improve the software as best as we are able to and your feedback will be beneficial insight into what our client's are seeking.
                                    \nIf you have availability, we would love to schedule a time to chat with you! Please let us know and we can provide a Calendly link which will allow you to select a time that best fits your availability. We can use this call to discuss your feedback, questions, or any concerns. However, if you would not like to schedule a call. Any information that you can provide in response to this email will be greatly appreciated.
                                    \nThank you!"""
                                    detractor_temp_two=f"Hi {client_capital}," f"""\n\nWe received your most recent survey response and would like to say thank you for taking the time to submit your feedback.
                                    \nYou scored us a  “{score}” out of 10 and it is unfortunate to receive such a score. However, we do understand that we can improve upon this and we would like to gather additional feedback from you.
                                    \nIf you would like, we could schedule a meeting to go over any feedback, concerns, and questions that you may have. Your feedback will be insightful and will help us understand what our client’s would like to see and how we can better the experience for our users.
                                    \nPlease let us know if you would like to speak and we will gladly get back to you with a scheduling link that will allow you to select a time of your choosing.
                                    \nIf you have any questions or concerns, please do not hesitate to reach out.
                                    \nThank you!"""
                                    detractor_temp_three=f"Hi {client_capital}," f"""\n\nThank you for taking the time to submit your survey response.
                                    \nWe received a score of “{score}” out of 10. Gauging from the submitted score, we understand that there might be some concerns. Your survey response is greatly appreciated as it will help us better understand how our users are feeling about the software.
                                    \nKnowing this, we would love to schedule a meeting with you to hear out any concerns that you may have!
                                    \nIf this is something you would like to discuss, feel free to reply back to this email and let us know! We'll follow up with a provided link that will allow you to schedule a time of your choosing.
                                    \nWe look forward to hearing from you.
                                    \nThank you!"""
                                    
                                    if survey_text_none == None:
                                        
                                        # If the score is greater than 8 (9,10), it will choose a random number 1:3 and choose a  Promoter template
                                        if (score>8):
                                            if ran_num == 1:
                                                emailMsg = promoter_temp_one
                                                temp_name = "Promoter Template #1"
                                            if ran_num == 2:
                                                emailMsg = promoter_temp_two
                                                temp_name = "Promoter Template #2"
                                            if ran_num == 3:
                                                emailMsg = promoter_temp_three
                                                temp_name = "Promoter Template #3"
                                            # Send out email
                                            mimeMessage = MIMEMultipart()
                                            mimeMessage['to'] = client_wootric_email # Recipient
                                            mimeMessage['from'] = 'support@meruscase.com' # Sends email as support
                                            mimeMessage['subject'] = 'Thank you for your feedback!' # Email Subject line
                                            mimeMessage.attach(MIMEText(emailMsg, 'plain'))
                                            raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
                                            message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
                                            print(f'Sending out Email to {client_wootric_email}', message) # prints the email as confirmation
                                            data_field = f'{i[0]}?completed=true'
                                            endpoint_path = f"/responses/{data_field}"
                                            endpoint = f"{wootric_api}{endpoint_path}"
                                            print('SURVEY MARKED AS COMPLETED')
                                            r = requests.put(endpoint, headers=headers_test)
                                            new_lst = [client['ID'],client_wootric_email, i[1], date_auto_sent, i[3], data[8], temp_name, "PROMOTER"]
                                            temp_lst.append(new_lst)


                                        # if the score is less than 7 (6:0), choose a Detractor Template
                                        if (score<7):
                                            if ran_num == 1:
                                                emailMsg = detractor_temp_one
                                                temp_name = "Detractor Template #1"
                                            if ran_num == 2:
                                                emailMsg = detractor_temp_two
                                                temp_name = "Detractor Template #2"
                                            if ran_num == 3:
                                                emailMsg = detractor_temp_three
                                                temp_name = "Detractor Template #3"

                                            # Send out email
                                            mimeMessage = MIMEMultipart()
                                            mimeMessage['to'] = client_wootric_email # Recipient
                                            mimeMessage['from'] = 'support@meruscase.com' # Sends email as support
                                            mimeMessage['subject'] = 'Thank you for your feedback!' # Email Subject Line
                                            mimeMessage.attach(MIMEText(emailMsg, 'plain'))
                                            raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
                                            message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
                                            print(f'Sending out Email to {client_wootric_email}', message)
                                            print(i[0])
                                            data_field = f'{i[0]}?completed=true'
                                            endpoint_path = f"/responses/{data_field}"
                                            endpoint = f"{wootric_api}{endpoint_path}"
                                            requests.put(endpoint, headers=headers_test)
                                            print('SURVEY MARKED AS COMPLETED')
                                            new_lst = [client['ID'],client_wootric_email, i[1], date_auto_sent, i[3], data[8], temp_name, "DETRACTOR"]
                                            temp_lst.append(new_lst)

                                        # if the score is greater than 6 and the score is less than 9. (7,8), choose a Passive Template
                                        if (score>6 and score <9):
                                            if ran_num == 1:
                                                emailMsg = passives_temp_one
                                                temp_name = "Passive Template #1"
                                            if ran_num == 2:
                                                emailMsg = passives_temp_two
                                                temp_name = "Passive Template #2"
                                            if ran_num == 3:
                                                emailMsg = passives_temp_three
                                                temp_name = "Passive Template #3"

                                            # Send out email
                                            mimeMessage = MIMEMultipart()
                                            mimeMessage['to'] = client_wootric_email # Recipient
                                            mimeMessage['from'] = 'support@meruscase.com' # Sends email as support
                                            mimeMessage['subject'] = 'Thank you for your feedback!' # Email Subject Line
                                            mimeMessage.attach(MIMEText(emailMsg, 'plain'))
                                            raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
                                            message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
                                            print(f'Sending out Email to {client_wootric_email}', message)
                                            print(i[0])
                                            data_field = f'{i[0]}?completed=true'
                                            endpoint_path = f"/responses/{data_field}"
                                            endpoint = f"{wootric_api}{endpoint_path}"
                                            requests.put(endpoint, headers=headers_test)
                                            print('SURVEY MARKED AS COMPLETED')
                                            new_lst = [client['ID'],client_wootric_email, i[1], date_auto_sent, i[3], data[8], temp_name, "PASSIVE"]
                                            temp_lst.append(new_lst)

        # close csv file
        loopfile.close()

def list_of_emails_sent():
    temp_survey_data = survey__data[1:]
    num_of_emails = 0

    for i in temp_survey_data:
        if i[2] == None:
            surveys_being_sent.append(temp_survey_data[num_of_emails])
            num_of_emails+=1      
    return surveys_being_sent

email_automate()

def update_email_data(temp_lst):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'keys_woo.json'

    credentials = None
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1b4GG3cbZe7uZLPTBx8F9aMnMFNg01UQ4jcKyZHX3uG4'
    service = build('sheets', 'v4', credentials=credentials)
    # Initialize rows with an initial header row

    # Range, value render, date/time for Google Sheets API
    range_ = "Automated Emails"
    value_render_option = 'UNFORMATTED_VALUE'
    date_time_render_option = 'FORMATTED_STRING'


    # send values to the body
    body = {
        'values':temp_lst
    }

    # update values in spreadsheet
    result = service.spreadsheets().values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_,
        valueInputOption='RAW',body=body).execute()



    # output how many cells were updated
    print('{0} cells updated.'.format(result.get('updatedCells')))
    print('{0} cells updated.'.format(result.get('updatedCells')))

update_email_data(temp_lst)

print("THIS IS THE TEMP LIST\n\n", temp_lst)

