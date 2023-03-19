from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
import sqlite3
import time
import pandas as pd

conn = sqlite3.connect('xchange_last_run.db')
cursor = conn.cursor()


BASE_URL = "http://zambrero.xchangelive.com.au/XchangeLive.aspx"

data = open("config.txt", "r")
for x in data:
    if 'start_date' in x:
        start_date = x.replace('start_date = ', '').replace('\n', '')
    if 'download_path' in x:
        download_path = x.replace('download_path = ', '').replace('\n', '')
    if 'username' in x:
        username = x.replace('username = ', '').replace('\n', '')
    if 'password' in x:
        password = x.replace('password = ', '').replace('\n', '')
    if 'iteration' in x:
        iteration = int(x.replace('iteration = ', '').replace('\n', ''))


def format_date(date, increment_by=0):
    date = datetime.strptime(date, '%d/%m/%Y')  # convert to datetime
    date = date + timedelta(days=increment_by)  # add 1 day
    date = date.strftime('%d/%m/%Y')  # convert back to string
    return date


def format_date_for_filename(date):
    start_datetime = datetime.strptime(date, "%d/%m/%Y")
    current_datetime = datetime.now()

    # Append current time to the start date
    start_datetime = start_datetime.replace(
        hour=current_datetime.hour,
        minute=current_datetime.minute,
        second=current_datetime.second
    )
    return start_datetime.strftime("%d%m%Y%H%M%S")



cursor.execute("SELECT * FROM last_run")
result = cursor.fetchone()
start_date = result[0]


def login(page, context):
    time.sleep(2)
    if page.title() == 'Login':
        print('Logging in ...')
        page.fill('input#TbUserName', username)
        page.fill('input#TbPassword', password)
        page.click('input#BtnLogin')
        print('Password entered')
        context.storage_state(path="auth.json")


with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(storage_state="auth.json", viewport={'width': 1280, 'height': 1000})
    page = context.new_page()
    page.goto(BASE_URL)

    login(page, context)
    time.sleep(2)

    try:
        page.get_by_text("Sales Reports").click()
        for j in range(iteration):
            start_date = format_date(start_date)
            end_date = format_date(start_date, 1)
            print("Iteration: " + str(j + 1))
            print("Start Date: " + start_date)
            print("End Date: " + end_date)
            print("----------------------")
            time.sleep(3)
            page.get_by_text("Product Mix By Location").click()

            page.frame_locator("iframe[name=\"main\"]").get_by_label("AU NSW Casula").check()
            page.frame_locator("iframe[name=\"main\"]").locator("#btnnext__4").click()
            print('AU NSW Casula checked, Next button clicked')
            time.sleep(3)
            page.frame_locator("iframe[name=\"main\"]").get_by_label("Select All").check()
            print('Select All checked')
            time.sleep(3)
            page.frame_locator("iframe[name=\"main\"]").locator("#btnnext__4").click()
            print('Next button clicked')
            time.sleep(3)
            page.frame_locator("iframe[name=\"main\"]").locator("#DDQuickStartDate_input").fill(start_date)
            print('Start Date filled')
            page.frame_locator("iframe[name=\"main\"]").locator("#DDQuickEndDate_input").fill(end_date)
            print('End Date filled')

            page.frame_locator("iframe[name=\"main\"]").locator("#btnfinish__4").click()
            print('Finish button clicked')

            print("Network idle")
            time.sleep(5)
            for i in range(0, 40):
                print(i)
                time.sleep(1)
                try:
                    with page.expect_download() as download_info:
                        page.frame(name="wacframe").get_by_role("button", name="Download").click()
                    download = download_info.value
                    download.save_as(download_path + "/downloaded.xls")
                    print("File Downloaded")
                    break
                except:
                    print("Not found")
            # Replace 'file.xlsx' with the name of your Excel file
            excel_file = pd.ExcelFile(download_path + '/downloaded.xls')

            # Replace 'Sheet3' with the name of the third sheet in your Excel file
            df = excel_file.parse('Table')
            df['Start Date'] = start_date

            file_name = download_path + '/zambreropos_date' + format_date_for_filename(start_date) + '.csv'

            # Save the file with the generated name
            df.to_csv(file_name, sep='|', index=False)
            start_date = end_date
            cursor.execute("UPDATE last_run SET date = ?", (end_date,))
    except:
        conn.commit()
        print("Error")


conn.commit()
conn.close()
print("Value updated in 'last_run' table.")
print('Execution complete.')
