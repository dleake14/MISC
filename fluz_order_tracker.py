import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
import googleapiclient.discovery
import time 
import email
import base64
import re
from datetime import datetime
import pymysql
import ast
from selenium import webdriver

# This line sets the permissions for what it can/can't do in gmail.
SCOPES = ['https://mail.google.com/']

driver = webdriver.Chrome()

# This opens the text file with a dictionary containing AWS login credentials.
file = open('aws_info.txt', 'r')
aws_login_creds = file.read()
connection = ast.literal_eval(aws_login_creds)

# This is the first function that is ran in order to create the gmail connection and set service variable. 
def get_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('gmail', 'v1', credentials=creds)
    return service

# This function searches gmail folder for the search string and returns the selected email id's. 
def search_messages(service, user_id, search_string):
    try:
        search_id = service.users().messages().list(userId = user_id, q = search_string).execute()
        number_results = search_id['resultSizeEstimate']
        final_list = []
        if number_results > 0:
            message_ids = search_id['messages']
            for ids in message_ids:
                # print('Message found!')
                final_list.append(ids['id'])
                # final_list.append(ids)
            return final_list
        else:
            print('There were no results for your search.')
            return ""
    except:
        print('An error occured:')

# This function gets the found email information and changes the data type to a readable datatype.
def get_message(service, user_id, msg_id):
    try:
        for item in msg_id:
            # Makes the connection and GETS the emails in RAW format. 
            message = service.users().messages().get(userId = user_id, id = item, format = 'raw').execute()
            # Changes format from RAW to ASCII
            msg_raw = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            # Changes format type again
            msg_str = email.message_from_bytes(msg_raw)
            # This line checks if the content is if multipart (plaintext and html) or single part
            content_types = msg_str.get_content_maintype()
            if content_types == 'multipart':
                # Part1 is plaintext and part2 is html text
                part1, part2 = msg_str.get_payload()
                raw_email = part1.get_payload()
                remove_char = ["|", "=20", "=C2=A0"]
                for i in remove_char:
                    raw_email = raw_email.replace(i, "")
                raw_email = "".join([s for s in raw_email.strip().splitlines(True) if s.strip()])
                return str(raw_email)
            else:
                return msg_str.get_payload()
    except:
        print('An error has occured during the get_message function.')

# This function pulls the order date/time/site/card_used from the returned data in get_messages
def get_data(email):
    dateFound = 'no'
    for line in email.splitlines():
        if 'you earned' in line.lower():
            site = line.split('at ')[1]
            site = site.split(".")[0]
        if dateFound == 'yes':
            # This line finds the Date line and separates the time and date and converts it 
            date_time = line
            rawDate = date_time.split(',')[0]
            month_dict = {'January': '01', 'Feburary': '02', 'March': '03', 'April': '04',
                    'May': '05', 'June': '06', 'July': '07', 'August': '08', 
                    'September': '09', 'October': '10', 'November': '11', 'December': '12'}
            month = rawDate.split(" ")[0]
            day = rawDate.split(' ')[1]
            if len(day) < 2:
                day = '0' + day
            correct_date = str(str(month_dict[month]) + '/' + day + '/' + str(datetime.now().year)[-2:])

            # This section will conver the time correctly to military time. 
            time = date_time.split(str(datetime.now().year))[1].strip()
            if time[-2:] == "AM" and time[:2] == "12":
                correct_time = "00" + time[2:-2]
            elif time[-2:] == "AM":
                correct_time = time[:-2]
            elif time[-2:] == "PM" and time[:2] == "12":
                correct_time = time[:-2]
            else: 
                hour = int(time.split(':')[0]) + 12
                minute = time.split(':')[1][0:2]
                correct_time = str(hour) + ":" + str(minute)
            dateFound = 'no'
        if "purchase date" in line.lower():
            dateFound = "yes"
        if "****" in line:
            card_used = line[-4:]
        if 'order #' in line.lower():
            id = line.split('#')[1]
            id = id.strip()
    return id, correct_date, correct_time, site, card_used

# This function grabs the URL from the email for the VGC
def get_URL(service, msg_id):
    try:
        for item in msg_id:
            message = service.users().messages().get(userId = 'me', id = item, format = 'raw').execute()
            msg_raw = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            msg_str = email.message_from_bytes(msg_raw)
            content_types = msg_str.get_content_maintype()
            use_next = 'no'
            end = 'no'
            if content_types == 'multipart':
                part1, part2 = msg_str.get_payload()
                part2 = part2.get_payload()
                for line in part2.splitlines():
                    if 'for visiting' in line.lower():
                        first_half = line.split('3D"')[1]
                        first_half = first_half.strip('=')
                        use_next = 'yes'
                    if use_next == 'yes':
                        if 'token' in line:
                            second_half = line.split('"')[0]
                            end = 'yes'
                            break
                    if end == 'yes':
                        break
                url = first_half + second_half
                return url
            else:
                return msg_str.get_payload()
    except:
        print('An error has occured during the URL function.')

# This function gets the code from verification email.
def get_code(service, msg_id):
    try:
        for item in msg_id:
            message = service.users().messages().get(userId = 'me', id = item, format = 'raw').execute()
            msg_raw = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            msg_str = email.message_from_bytes(msg_raw)
            content_types = msg_str.get_content_maintype()
            if content_types == 'multipart':
                part1, part2 = msg_str.get_payload()
                part1 = part1.get_payload()
                for line in part1.splitlines():
                    if 'is:' in line:
                        code = line.split('is:')[1]
                        code = code.strip()
                        return code
                print('no code found')
            else:
                return msg_str.get_payload()
    except:
        print('An error has occured during the get_code function.')

# This function takes 3 variables and adds them to the 'orders' table in the fluz db.
def update_orders(id, date, time, site, cardUsed, card_num, exp_date, sec_code):
    con = pymysql.connect(connection['host'], connection['username'], connection['password'], connection['db'])
    cursor = con.cursor()

    insert_statement="""
    INSERT INTO fluz (id, date, time, site, cardUsed, card_num, exp_date, sec_code)
    VALUES (""" + "'" + id + "', '" + date + "', '" + time + "', '" + site + "', '" + cardUsed + "', '" + card_num + "', '" + exp_date + "', '" + sec_code + "')"

    cursor.execute(insert_statement)
    con.commit() 

    select_statement = """
    SELECT * FROM orders
    """
    cursor.execute(select_statement)

    results = cursor.fetchall()
    for row in results:
        print(row)
    
    con.close()

if __name__ == "__main__":
    service = get_service()
    last_email = ''
    email1 = 'recent purchase'
    email2 = 'Virtual Visa'
    user_id = 'me'
    run = 1
    last1 = ''
    last2 = ''
    new1 = 'no'
    new2 = 'no'
    new_email1 = ''
    new_email2 = ''
    while run > 0:
        # searching for the first email from FLUZ with order info. 
        new_email1 = search_messages(service, user_id, email1)
        if last1 == new_email1:
            new1 = 'no'
        else:
            new1 = 'yes'
        if new1 == 'yes':
            if len(new_email1) > 0:
                raw_email = get_message(service, user_id, new_email1)
                id = raw_email[0]
                date_ordered = raw_email[1]
                time_ordered = raw_email[2]
                site_ordered = raw_email[3]
                card_used = raw_email[4]
                last1 = new_email1
                print(date_ordered)

        # Searching for the 2nd email with the code. 
        new_email2 = search_messages(service, user_id, email2)
        if last2 == new_email2:
            new2 = 'no'
        else:
            new2 = 'yes'
        if new2 == 'yes':
            if len(new_email2) > 0:
                url = get_URL(service, new_email2)
                driver.get(url)
                #code = str(input('code: '))
                code = ''
                while code == '':
                    code_email = search_messages(service, user_id, 'Verification')
                    if len(code_email) > 0:
                        code = get_code(service, code_email)
                    time.sleep(5)
                box = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div/div[1]/div/div/div/div/div[2]/form[1]/div[1]/div/input')
                box.send_keys(code)
                box.send_keys(u'\ue007')
                time.sleep(1)
                box = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[1]/div/div/div[2]/form/div[2]/div[2]/div[1]/label')
                box.click()
                box = driver.find_element_by_xpath('//*[@id="btn-continue"]')
                box.click()
                time.sleep(1)
                card_num = driver.find_element_by_xpath('//*[@id="cardCollapsableContent"]/div[1]/div[1]/div[1]/span').text
                print(card_num)
                exp_date = driver.find_element_by_xpath('//*[@id="cards__wrapper"]/div/div/div[2]/div/div[2]/div[1]').text
                # exp_date = exp_date.strip('Expiration Date')
                exp_date = exp_date.splitlines()[1]
                print(exp_date)
                sec_code = driver.find_element_by_xpath('//*[@id="cardCollapsableContent"]/div[1]/div[1]/div[2]').text
                # sec_code = sec_code.strip('CVV')
                sec_code = sec_code.splitlines()[1]
                print(sec_code)
                run = 2
        print('looping')
        time.sleep(5)
        if run > 1:
            update_orders(id, date_ordered, time_ordered, site_ordered, card_used, card_num, exp_date, sec_code)
            run = 1 
            print('The order has been updated to the AWS RDS Database')
            time.sleep(10)





