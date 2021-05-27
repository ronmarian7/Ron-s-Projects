from __future__ import print_function
import os
import os.path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import csv
import smtplib
from email.message import EmailMessage
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from time import sleep
import datefinder



def create_event(start_time_str, summary, description=None, location=None):
  
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    calendar_list = service.calendarList().list().execute()
    tasks_calender_id = calendar_list['items'][1]['id']

    start_time = list(datefinder.find_dates(start_time_str))[0].date()

    end_time = start_time
    timezone = 'Asia/Jerusalem'

    event = {
        'summary': summary,
        'start': {
            'date': start_time.strftime("%Y-%d-%m"),
            'timeZone': timezone,
        },
        'end': {
            'date': end_time.strftime("%Y-%d-%m"),
            'timeZone': timezone,
        },
    }
    service.events().insert(calendarId=tasks_calender_id, body=event).execute()


def send_mail(subject, body):
    EMAIL_ADDRESS = "ronmarian7@gmail.com" #Enter Email address 
    EMAIL_PASSWORD = "1234" #Enter Email address password 
    # Build the message block  
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = 'ronmarian7@gmail.com'
    msg.set_content(body)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD) # logging using the Credentials above
        smtp.send_message(msg)


def diff_copy(): # finding the diff between the file to see if any new tasks added
    # IF THERE ARE ANY CHANGES SEND AN EMAIL WITH THE TASK INFORMATION AND ADD THE TASK TO GOOGLE CALENDAR
    course, task, date, time = '', '', '', ''
    with open(input_file1, 'r', encoding='utf-8') as t1, open(input_file2, 'r', encoding='utf-8') as t2:
        fileone = t1.readlines()
        filetwo = t2.readlines()
        with open(output_path, 'w', encoding='utf-8') as outFile:
            for line in filetwo:
                if line not in fileone:
                    outFile.write(line)
    with open(output_path, 'r', encoding='utf-8') as outFile:
        out_file = outFile.readlines()
        if len(out_file) != 0:
            for line in out_file:
                course, task, date, time = line.split(',', 3)
                subject = f'{course} {task}'
                body = f' פורסמה {task} בקורס {course} להגשה בתאריך {date} בשעה {time}'
                send_mail(subject, body) # CALL THE MAIL SENDER FUNC
                create_event(date, f'{task}-{course}') # CALL THE GOOGLE CALENDAR FUNC TO MAKE EVENT
    with open(input_file1, 'w', encoding='utf-8') as t1:
        t1.writelines(filetwo)

# enter the paths
input_file1 = "PATH FILE 1" # THE TASKS FILE FROM PREV RUN
input_file2 = "PATH FILE 2" # THE TASKS FILE FROM CUR RUN
output_path = "PATH OUT FILE" # THE FILE TO WRITE THE DIFF BETWEEN THE FILES


options = webdriver.ChromeOptions()
options.page_load_strategy = 'normal'
# if you want the chrome browser to work in the background un comments the line below
#options.add_argument('--headless')
driver = webdriver.Chrome('PATH OF THE CHROME DRIVER\\chromedriver.exe', options= options) # ENTER THE PATH OF THE CHROME DRIVER
driver.get('https://moodle.tau.ac.il/login/index.php')
sleep(10) # LETTING THE PAGE TO LOAD
try:
    driver.switch_to.frame(driver.find_element_by_id('content'))
    driver.switch_to.frame(driver.find_element_by_id('credentials'))
    user_bar = driver.find_element_by_name("Ecom_User_ID")
    user_bar.send_keys('USER NAME')
    pin_bar = driver.find_element_by_name("Ecom_User_Pid")
    pin_bar.send_keys('ID_NUMBER')
    pass_bar = driver.find_element_by_name("Ecom_Password")
    pass_bar.send_keys('PASSWORD')
    # FINDING THE LOGGING BUTTON 
    login_btn = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/input')
    # click the login
    login_btn.click()
    sleep(40)

    time_table_era = WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, "page-container-2"))) # LOOKING FOR THE CONTAINER OF THE TASKS IN THE PAGE
    out_table = time_table_era.find_elements_by_class_name('list-group')
    dates = time_table_era.find_elements_by_tag_name('h5')
    dates_list = []
    for i in range(len(dates)):
        dates_list.append(dates[i].text.strip()) # making a nested list containing the tasks by dates
    temp_task = []
    for i in range(len(out_table)):
        in_table = []
        in_table.append(out_table[i].find_elements_by_css_selector('h6.event-name'))
        temp_task.append(in_table)
    tasks = time_table_era.find_elements_by_tag_name('h6')
    tasks_list = []
    for i in range(0, len(tasks), 2):
        if "'" in tasks[i].text:
            task = tasks[i].text.split("'", 2)[1].strip()
            tasks_list.append(task)
        elif "is due" in tasks[i].text:
            task = tasks[i].text.split("is due")[0].strip()
            tasks_list.append(task)
        else:
            tasks_list.append(tasks[i].text)
    courses = time_table_era.find_elements_by_tag_name('small')
    courses_list = []
    times_list = []
    for i in range(0, len(courses), 2):
        course = courses[i].text.split('-', 1)[1].strip(' "')
        courses_list.append(course)
        times_list.append(courses[i + 1].text)
    temp_task = []
    for i in range(len(out_table)):
        in_table = []
        in_table.append(out_table[i].find_elements_by_css_selector('h6.event-name'))
        temp_task.append(in_table)
    # making a list containing the dates with multi
    new_dates_list = []
    for i in range(len(temp_task)):
        for j in range(len(temp_task[i][0])):
            new_dates_list.append(dates_list[i])
    with open(input_file2, 'w', newline='', encoding='utf-8') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Course', 'Task', 'Date', 'Time'])
        for i in range(len(courses_list)):
            csv_writer.writerow([f'{courses_list[i]}', f'{tasks_list[i]}', f'{new_dates_list[i]}', f'{times_list[i]}'])
    diff_copy() # RUNNINGS THE FUNC TO SEE IF THERE IS ANY DIFF
except Exception as e:
    print(e)
    with open("Error File.txt", 'w', encoding='utf-8') as t1: # IF THERE IS ANY ERROR MAKE A FILE AND WRITE THE MESSAGE
        t1.write(f"{e}")
    raise e
finally:
    driver.quit()
