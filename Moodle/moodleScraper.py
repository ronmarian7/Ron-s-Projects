from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import csv
import smtplib
import os
from email.message import EmailMessage
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from datetime import date, timedelta
import datefinder
import datetime
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import schedule
import time

from selenium.webdriver.support.wait import WebDriverWait


def create_event(start_time_str, summary, description=None, location=None):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """

    # If modifying these scopes, delete the file token.json.
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
#    EMAIL_ADDRESS = os.environ.get('EMAIL_ADDRESS')
#   EMAIL_PASSWORD = os.environ.get('EMAIL_PASSWORD')
    EMAIL_ADDRESS = "ronmarian7@gmail.com"
    EMAIL_PASSWORD = "oxcezloendyagyff" # old
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = 'ronmarian7@gmail.com'
    msg.set_content(body)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)


def diff_copy():
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
                send_mail(subject, body)
                create_event(date, f'{task}-{course}')
    with open(input_file1, 'w', encoding='utf-8') as t1:
        t1.writelines(filetwo)


input_file1 = "C:\\Users\\ronma\\PycharmProjects\\newmoodlescrap\\temp.csv"
input_file2 = "C:\\Users\\ronma\\PycharmProjects\\newmoodlescrap\\ClassHw.csv"
output_path = "C:\\Users\\ronma\\PycharmProjects\\newmoodlescrap\\out.csv"

def job():
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'normal'
    #options.add_argument('--headless')
    driver = webdriver.Chrome('C:\\Users\\ronma\PycharmProjects\\webscrapingproject\\chromedriver.exe', options= options)
    driver.get('https://moodle.tau.ac.il/login/index.php')
    sleep(10)

    try:
        driver.switch_to.frame(driver.find_element_by_id('content'))
        driver.switch_to.frame(driver.find_element_by_id('credentials'))
        user_bar = driver.find_element_by_name("Ecom_User_ID")
        user_bar.send_keys('ronmarian')
        pin_bar = driver.find_element_by_name("Ecom_User_Pid")
        pin_bar.send_keys('316593839')
        pass_bar = driver.find_element_by_name("Ecom_Password")
        pass_bar.send_keys('Ronm6800534')
        # click the login
        login_btn = driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/input')
        login_btn.click()
        sleep(40)
        # click on dashboard
        #driver.find_element_by_xpath('//*[@id="nav-drawer"]/nav[2]/a[2]/div/div/span[2]').click() ## erised
        #sleep(20)


        # getting the table scoop
        #time_table_era = driver.find_element_by_class_name('paged-content-page-container')
        #time_table_era = WebDriverWait(driver, 60).until(
        #    EC.visibility_of_element_located((By.CLASS_NAME, "paged-content-page-container")))
        #out_table = time_table_era.find_elements_by_class_name('list-group')

        time_table_era = WebDriverWait(driver, 60).until(
            EC.visibility_of_element_located((By.ID, "page-container-2")))
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
        diff_copy()
    except Exception as e:
        print(e)
        with open("Error File.txt", 'w', encoding='utf-8') as t1:
            t1.write(f"{e}")
        raise e
    finally:
        driver.quit()


#schedule.every(5).minutes.do(job)
#schedule.every().hour.do(job)
#schedule.every().day.at('13:58').do(job)
# schedule.every(5).to(10).minutes.do(job)
# schedule.every().monday.do(job)
# schedule.every().wednesday.at("13:15").do(job)
# schedule.every().minute.at(":17").do(job)

#while True:
 #   schedule.run_pending()
  #  time.sleep(1) # wait one minute

job()