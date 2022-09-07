import argparse
from pathlib import Path
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import yaml
from datetime import datetime as dt
from typing import Tuple
import time
from datetime import timedelta

class MInfo():
    def __init__(self, meeting_info):
        self.__dict__.update(meeting_info)

def process_date(target_date: str | None=None) -> dt:
    # if target date is None then set it as today
    date_fmt='%Y-%m-%d'
    if target_date is None:
        target_date = dt.today()
    else:
        # check datetime format
        try:
            target_date = dt.strptime(target_date, date_fmt)
        except ValueError:
            raise ValueError(f'Incorrect data format, should be {date_fmt}')
    return target_date

def get_datetime(datetime: Tuple[str]):
    date, time, ampm = datetime
    if ampm in ['오전', '오후']:
        ampm = {'오전': 'AM', '오후': 'PM'}.get(ampm)
    t = dt.strptime(' '.join([date, time, ampm]), '%Y/%m/%d %I:%M:%S %p')    
    return t

def zoom_split(line: str) -> Tuple[int|str]:
    *tkn_ids, line = line.split(maxsplit=3)
    meeting_id = int(''.join(tkn_ids))
    *meeting_start, line = line.split(maxsplit=3)
    meeting_dt = get_datetime(meeting_start)
    meeting_num_attendance = int(line)
    return (meeting_id, meeting_dt, meeting_num_attendance)

def match_meeting(line: str, m_info: MInfo, target_date: dt):
    # today must be a format of %Y-%m-%d
    date_fmt='%Y-%m-%d'
    t_date = dt.strftime(target_date, date_fmt)
    m_id, m_dt, m_na = zoom_split(line)
    date = m_dt.strftime(date_fmt)
    
    return (t_date == date) and (m_info.start_hour == m_dt.hour) and (m_info.id == m_id)

def set_month(driver, datetime):
    calender = driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[1]')
    current_month_year = calender.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[1]/div').text
    current_month_year = dt.strptime(current_month_year.replace('월', ''), '%m %Y')

    number_to_click = current_month_year.month - datetime.month
    for i in range(number_to_click):
        calender.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[1]/a[1]').click()
        time.sleep(1)

    calender_days = driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/table/tbody').find_elements(By.TAG_NAME, 'a')
    for i, c in enumerate(calender_days):
        print(c.text, end=' ')
        if int(c.text) == datetime.day:
            break
    c.click()

def main(settings, m_info, target_date: dt):
    start_date = target_date - timedelta(days=30)

    driver = uc.Chrome() 
    driver.get('https://zoom.us/account/my/report')
    # login to google
    driver.find_element(By.XPATH, '//*[@id="login-btn-group"]/div[1]/a[3]').click()
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="identifierId"]')))
    driver.find_element(By.XPATH, '//*[@id="identifierId"]').send_keys(settings['email'])
    driver.find_element(By.XPATH, '//*[@id="identifierNext"]/div/button/span').click()
    time.sleep(3)
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="password"]')))
    driver.find_element(By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input').send_keys(settings['pass'])
    driver.find_element(By.XPATH, '//*[@id="passwordNext"]/div/button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="onetrust-close-btn-container"]/button').click()
    time.sleep(1)

    # always search from -30 days from today
    driver.find_element(By.XPATH, '//*[@id="searchMyForm"]/div/button[1]').click()  # set start 
    time.sleep(1)
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]')))
    set_month(driver, start_date)

    driver.find_element(By.XPATH, '//*[@id="searchMyForm"]/div/button[2]').click()  # set end 
    time.sleep(1)
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ui-datepicker-div"]')))
    set_month(driver, target_date)

    driver.find_element(By.XPATH, '//*[@id="searchMyButton"]').click()
    time.sleep(3)


    # filter columns
    dropdown_button = driver.find_element(By.XPATH, '//*[@id="meetingDropdownMenu"]/button')
    dropdown_button.click()
    time.sleep(1)
    dropdown_list = dropdown_button.find_element(By.XPATH, '//*[@id="meetingDropdownMenu"]/ul')\
        .find_elements(By.CLASS_NAME, 'zoom_dropdownlist')

    col_list = ['회의 ID', '시작 시간', '참가자']
    for d in dropdown_list:
        f_name = d.find_element(By.TAG_NAME, 'label').text
        f_box = d.find_element(By.TAG_NAME, 'input')
        if (f_name in col_list):
            if not f_box.is_selected():
                f_box.click()
        else:
            f_box.click()
    driver.find_element(By.XPATH, '//*[@id="meetingDropdownMenu"]/button').click()
    time.sleep(3)

    # table
    a = driver.find_element(By.XPATH, '//*[@id="meeting_list"]/tbody')
    meetings = []
    for i, line in enumerate(a.text.splitlines()):
        if match_meeting(line, m_info, target_date=target_date):
            meetings.append(i)
    assert len(meetings) == 1, 'must be one single matched'
    m_idx = meetings[0]
    a.find_elements(By.TAG_NAME, 'tr')[m_idx].find_element(By.TAG_NAME, 'a').click()
    time.sleep(3)
    # get the attandances
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="selectUnique"]')
        )
    )
    driver.find_element(By.XPATH, '//*[@id="selectUnique"]').click()
    driver.find_element(By.XPATH, '//*[@id="btnExportParticipants"]').click()
    time.sleep(5)
    driver.close()
    print('Exproted')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--lec', required=True, type=str)  # BKMS2, P4DS
    parser.add_argument('--target', default='', type=str)  # %Y-%m-%d
    args = parser.parse_args()

    main_path = Path().resolve()
    with (main_path / 'settings.yml').open('r') as file:
        settings = yaml.load(file, yaml.FullLoader)
    meeting_info = {}

    for k, v in settings['meeting_information'].items():
        meeting_info[k] = MInfo(v)

    m_info = meeting_info[args.lec]
    target_date = None if args.target == '' else args.target
    target_date = process_date(target_date)
    main(settings, m_info, target_date=target_date)