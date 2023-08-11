# Made by Joe Nogosek; joe.nogosek@ulteig.com
# 3/22/2022

from selenium import webdriver
import time

twoFA = input("Enter six digit 2FA code: ")

driver = webdriver.Chrome(executable_path=r"C:\Users\joe.nogosek\Documents\Python\chromedriver.exe")
driver.get('https://xcelenergy.my.salesforce.com/?ec=302&startURL=%2Fhome%2Fhome.jsp')

xcel_login = driver.find_element_by_xpath('//button[normalize-space()="Xcel Energy CORP credentials"]').click()

username = '239665'
user_box = driver.find_element_by_id('username')
user_box.send_keys(username)

password = 'airdoc1Ee'
pass_box = driver.find_element_by_id('password')
pass_box.send_keys(password)

sign_on_button = driver.find_element_by_css_selector("a[onclick^='postOk']").click()

passcode_box = driver.find_element_by_name('pf.pass')
passcode_box.send_keys('2846' + str(twoFA))

submit_button = driver.find_element_by_xpath("//button[contains(@onclick,'postOk')]").click()

# Put the IA that is being processed here and pull it from the folder name
driver.find_element_by_id('phSearchInput').send_keys("IA37295")
driver.find_element_by_id('phSearchButton').click()

time.sleep(2)
# Click the top result from case number selector
driver.find_element_by_css_selector("a[target='_top']").click()
time.sleep(5)

driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[10]/div[1]/div[1]/div[2]/div[1]/a[2]").click()
time.sleep(3)

j=0
for i in range(2,25):
    if driver.find_element_by_xpath(f"/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[3]/div[1]/div[2]/table[1]/tbody[1]/tr[" + str(i) + "]/td[2]").get_attribute("innerText") == 'One-Line Diagram' or driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[3]/div[1]/div[2]/table[1]/tbody[1]/tr[" + str(i) + "]/td[2]").get_attribute("innerText") == 'Site Plan':
         driver.find_element_by_xpath(f"/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[3]/div[1]/div[2]/table[1]/tbody[1]/tr[" + str(i) + "]/th[1]/a[1]").click()
         time.sleep(3)
         driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[9]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[2]/td[1]/a[2]").click()
         time.sleep(5)
         driver.back()
         time.sleep(1)
         j+=1
         if j == 2:
             break
driver.back()
time.sleep(1)


# reports_button = driver.find_element_by_css_selector("[title*='Reports Tab']").click()
#
# time.sleep(5)
# driver.find_element_by_css_selector("div[class='nameFieldContainer descrContainer']").click()
#
# driver.find_element_by_css_selector("[title*='Export Details']").click()
#
# driver.find_element_by_css_selector("[title*='Export']").click()
#
# driver.close()
# Uncheck everything that is checked if verify complete

# Click the 'Go to List' page for checks
# driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[17]/div[1]/div[1]/div[2]/div[1]/a[2]").click()
# # 3.21
# isChecked = driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/table[1]/tbody[1]/tr[1]/td[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[2]/td[5]/img[1]").get_attribute("alt")
# print(isChecked)
