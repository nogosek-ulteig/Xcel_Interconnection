import os

user_env = os.getlogin()

username = 'joseph.h.nogosek@xcelenergy.com'
password = ''
name = 'Joseph Nogosek'

path_to_driver = os.path.join('C:\\Users', user_env, 'Documents', 'Python','chromedriver.exe')
path_to_MN_CR_tracker = os.path.join('C:\\Users', user_env, 'Ulteig Engineers, Inc', 'Xcel DER System Impact Studies - DER Reviews - NSP','MN Completeness Review Tracker.xlsx')
path_to_MN_initials_tracker = os.path.join('C:\\Users', user_env, 'Ulteig Engineers, Inc', 'Xcel DER System Impact Studies - DER Reviews - NSP','MN Initial Review Screen Tracker.xlsx')
path_to_MN_preapps_tracker = os.path.join('C:\\Users', user_env, 'Ulteig Engineers, Inc', 'Xcel DER System Impact Studies - DER Reviews - NSP','MN Preapplication Report Tracker.xlsx')
path_to_WI_CR_tracker = os.path.join('C:\\Users', user_env, 'Ulteig Engineers, Inc', 'Xcel DER Work - DER Reviews - NSP WI','WI Completeness Review Tracker.xlsx')
path_to_CO_CR_tracker = os.path.join('C:\\Users', user_env, 'Ulteig Engineers, Inc', 'Xcel DER Work - DER Reviews - PSCo','CO Completeness Review Tracker.xlsx')
path_to_downloads = os.path.join('C:\\Users', user_env, 'Downloads')
