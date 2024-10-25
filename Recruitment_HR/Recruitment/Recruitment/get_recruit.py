import logging
import pandas as pd
from datetime import datetime, timedelta
import time

from selenium import webdriver
from selenium.webdriver.support.select import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import exceptions
import glob

import RecruitSetting
from RecruitSetting import *

def get_data_web(dateint_obj):
    try:
        option = webdriver.ChromeOptions()
        pref = {'download.default_directory': ori_rec}
        option.add_experimental_option('prefs', pref)
        option.add_argument('ignore-certificate-errors')
        option.add_argument("--no-sandbox")
        option.add_argument("--disable-dev-shm-usage")
        option.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options= option)
        driver.implicitly_wait(30)
    except Exception as e:
        print(e)
    
    login_ins = RecruitSetting.Login()
    login_ins.login()
    login_ins.webdri()
    driver.get(login_ins.recruitment)
    print("Open webdriver")
    driver. maximize_window()
    driver.find_element(By.NAME, 'username').send_keys(login_ins.user)
    time.sleep(2)
    driver.find_element(By.NAME, 'pwd').send_keys(login_ins.password + Keys.ENTER)
    # 24/06/2024 00:00 - 24/06/2024 23:59
    find_excelfile = [file for file in os.listdir(celendar_path) if file.endswith('.xlsx')]
    if find_excelfile:
        first_file = os.path.join(celendar_path, find_excelfile[0])
    monthstr_obj = datetime.now().strftime('%b')
    df = pd.read_excel(first_file, sheet_name= 'Sheet1')
    for col in df.columns:
        if col == monthstr_obj:
            print(f"Column '{col}' matches the current month.")
            # print(df[col])
            for index, row in df.iterrows():
                if pd.notna(row[col]):
                    start_date, end_date = row[col].split('-')
                    start_date, end_date = int(start_date), int(end_date)
                    
                    if start_date <= dateint_obj.day <= end_date:
                        print(f"Current day {dateint_obj} is in the range {row[col]} for week {row['Week']}")
                        break
                        
    format = "%d/%m/%Y"
    start_date = datetime(day= start_date, month= dateint_obj.month, year= dateint_obj.year).strftime(format)
    end_date = datetime(day= end_date, month= dateint_obj.month, year= dateint_obj.year).strftime(format)
    print(f"{start_date} - {end_date}")
    driver.find_element(By.ID, 'reservationtime').clear()
    driver.find_element(By.ID, 'reservationtime').send_keys(start_date + ' 00:00 - ' + end_date + ' 23:59' + Keys.ENTER)
    # print(nowstr_obj + ' ' + '00:00' + ' ' + '-' + ' ' + weekstr_obj + '23:59 ')
    driver.find_element(By.ID, 'reservationtime').send_keys(Keys.ENTER)
    driver.find_element(By.ID, 'searchTable').click()
    driver.find_element(By.XPATH, '//*[@id="recruitTable_wrapper"]/div[1]/button').click()
    wait = 1
    while wait == 1:
        wait = driver.execute_script('return jQuery.active;')
        time.sleep(0.5)
    time.sleep(3)
    logging.debug('Downloading')
    time.sleep(10)
    check_file_dw = glob.glob(os.path.join(ori_rec, '*.xlsx'))
    if check_file_dw:
        print('Download Success.')
    else:
        print('Download Fail.')
    logging.info('Quit web driver')
    driver.close()
    return dateint_obj
# if __name__ == '__main__':
#     get_data_web(datetime.now() - timedelta(days=5))

