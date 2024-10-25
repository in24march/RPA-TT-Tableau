import os
from pathlib import Path

class Login:
    def __init__(self) -> None:
        pass
    
    def login(self):
        self.user = "tharathn"
        self.password = "Julli@772244"
        
    def webdri(self):
        self.recruitment = "https://chatbot.cloud.ais.th/recruit"

class SorcingCH:
    def __init__(self) -> None:
        pass
    
    def chanel(self):
        self.list_ch = ['Sheet1', 'Facebook', 'Jobthai', 'TikTok', 'GFGJ', 'Other']
        self.fillter_ch = ['jobthai', 'facebook', 'tiktok']
    
    def map(self):
        self.s_map  = {
            'Jobthai' : 'Jobthai', 
            'Facebook' : 'Facebook', 
            'TikTok': 'TikTok',
            'เพื่อนหรือคนรู้จักที่ทำงานใน AIS/ACC' : 'GFGJ'
            }
    def sheet(self):
        self.sheetdata = {'02_Application' : 'Application',
                          '02_Interview_Pass' : 'Final Interview by HR',
                          '02_Hiring' : 'Hiring',
                          '02_Pre-training' : 'Pre-training ',
                          }
    def date_edit(self):
        self.date_range = ['02_Application',
                           '02_Interview_Pass',
                           '02_Hiring',
                           '02_Pre-training',
                           '02_Sourcing_Channel_All',
                           '02_Sourcing_Channel_BKK',
                           '02_Sourcing_Channel_NMA',]
class find_file:
    def __init__(self, path):
        self.path = path
        
    def find_excel(self):
        path_file = self.path
        find_excel = [file for file in os.listdir(path_file) if file.endswith('.xlsx')]
        if find_excel:
            first_file = os.path.join(path_file, find_excel[0])
            return first_file
        else:
            return None
    def find_ex_time(self):
        path_file = self.path
        find_excelfile = [file for file in os.listdir(path_file) if file.endswith('.xlsx')]
        if find_excelfile:
            first_file_path = max(
            (os.path.join(path_file, file) for file in find_excelfile),
            key=os.path.getmtime
            )
            return first_file_path
        else:
            return None

dir_path = os.path.dirname(os.path.abspath(__file__)) + os.sep
ori_rec = dir_path + 'Recruitment' + os.sep
celendar_path = dir_path + 'Calendar' + os.sep
RCm_filter = dir_path + 'RCM Filter' + os.sep
master_path = dir_path + 'master' + os.sep
User_path = dir_path + 'User_add' + os.sep
RCm_total = dir_path + 'RCM_total' + os.sep
Path(ori_rec).mkdir(parents= True, exist_ok= True)
Path(User_path).mkdir(parents= True, exist_ok= True)
Path(RCm_filter).mkdir(parents= True, exist_ok= True)
Path(RCm_total).mkdir(parents= True, exist_ok= True)