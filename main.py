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

class MInfo():
    def __init__(self, meeting_info):
        self.__dict__.update(meeting_info)

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

def match_meeting(line: str, m_info: MInfo, today: str | None=None):
    # today must be a format of %Y-%m-%d
    date_fmt='%Y-%m-%d'
    m_id, m_dt, m_na = zoom_split(line)
    if today is None:
        today = dt.today().strftime(date_fmt)
    else:
        today = dt.strptime(today, date_fmt)
    date = m_dt.strftime(date_fmt)
    return (today == date) and (m_info.start_hour == m_dt.hour) and (m_info.id == m_id)

def main(settings, m_info, today=None):

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
    driver.find_element(By.XPATH, '//*[@id="meeting_list"]/tbody')
    driver.find_element(By.XPATH, '//*[@id="meetingDropdownMenu"]/button').click()
    # filter columns
    dropdown_list = driver.find_element(By.XPATH, '//*[@id="meetingDropdownMenu"]/button')\
        .find_element(By.XPATH, '//*[@id="meetingDropdownMenu"]/ul')\
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
        if match_meeting(line, m_info, today=today):
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
    parser.add_argument('--today', default='', type=str)
    args = parser.parse_args()


    main_path = Path().resolve()
    with (main_path / 'settings.yml').open('r') as file:
        settings = yaml.load(file, yaml.FullLoader)
    meeting_info = {}

    for k, v in settings['meeting_information'].items():
        meeting_info[k] = MInfo(v)

    m_info = meeting_info[args.lec]
    today = None if args.today == '' else args.today
    main(settings, m_info, today=today)