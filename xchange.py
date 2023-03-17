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


def format_date(date, increment_by=0):
    date = datetime.strptime(date, '%d/%m/%Y')  # convert to datetime
    date = date + timedelta(days=increment_by)  # add 1 day
    date = date.strftime('%d/%m/%Y')  # convert back to string
    return date


cursor.execute("SELECT * FROM last_run")
result = cursor.fetchone()
start_date = result[0]

start_date = format_date(start_date)
end_date = format_date(start_date, 1)


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

    page.get_by_text("Sales Reports").click()
    page.get_by_text("Product Mix By Location").click()

    page.frame_locator("iframe[name=\"main\"]").get_by_label("AU NSW Casula").check()
    page.frame_locator("iframe[name=\"main\"]").locator("#btnnext__4").click()
    time.sleep(1)
    page.frame_locator("iframe[name=\"main\"]").get_by_label("Select All").check()
    time.sleep(2)
    page.frame_locator("iframe[name=\"main\"]").locator("#btnnext__4").click()
    time.sleep(2)
    page.frame_locator("iframe[name=\"main\"]").locator("#DDQuickStartDate_input").fill(start_date)
    page.frame_locator("iframe[name=\"main\"]").locator("#DDQuickEndDate_input").fill(end_date)

    page.frame_locator("iframe[name=\"main\"]").locator("#btnfinish__4").click()

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

    # Replace 'output_file.txt' with the name of the file you want to create

    now = datetime.now()
    file_name = download_path + '/zambreropos_date' + now.strftime("%d%m%Y%H%M%S") + '.csv'

    # Save the file with the generated name
    df.to_csv(file_name, sep='|', index=False)


cursor.execute("UPDATE last_run SET date = ?", (end_date,))
conn.commit()
conn.close()
print("Value updated in 'last_run' table.")
print('Execution complete.')
