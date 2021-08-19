from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.alert import Alert
from bs4 import BeautifulSoup
import pandas as pd
import sys

option = webdriver.ChromeOptions()
option.add_argument("--incognito")
option.add_argument("--headless")
option.add_argument("--disable-gpu")

username = input("Username: ")
password = input("Password: ")

driver = webdriver.Chrome("chromedriver")
driver.get("https://wish.wis.ntu.edu.sg/pls/webexe/ldap_login.login?w_url=https://wish.wis.ntu.edu.sg/pls/webexe/aus_stars_planner.main")
driver.find_element_by_id("UID").send_keys(username)
driver.find_elements_by_xpath("//input[@type='submit']")[0].click()
driver.find_element_by_id("PW").send_keys(password)
driver.find_elements_by_xpath("//input[@type='submit']")[0].click()

# wait the ready state to be complete
WebDriverWait(driver=driver, timeout=10).until(
    lambda x: x.execute_script("return document.readyState === 'complete'")
)

#Open Plan 3
try:
    Select(driver.find_element_by_name("plan_no")).select_by_visible_text("Plan 3")
except:
    holder = input("Login failed. Hit Enter to continue...")
    sys.exit()

driver.find_elements_by_xpath("//input[@value='Load']")[0].click()
Alert(driver).accept()

#Open details of Plan 3 courses
links = driver.find_elements_by_xpath("//a[contains(@href, 'javascript:view_subject')]")
for link in links:
    link.click()

#Save schedule
schedule_dfs = []
#Before reverse: initial tab, last opened tab...first opened tab
tab_handles = driver.window_handles
#inplace reverse and returns None
tab_handles.reverse()
for tab_handle in tab_handles[:-1]:
    driver.switch_to.window(tab_handle)
    soup = BeautifulSoup(driver.page_source, "lxml")
    tables = soup.find_all("table")
    tables = pd.read_html(str(tables))
    course = [tables[0].iat[0, 0].replace("[+] ", "")]
    try:
        schedule = tables[1].iloc[1:,:].reset_index(drop=True)
        course = pd.Series(course, name="Course")
        schedule = pd.concat([course, schedule], axis=1)
        schedule_dfs.append(schedule)
    except IndexError:
        message = "Schedule for " + course[0] + " not available."
        print(message)

driver.quit()

combined_df = pd.concat(schedule_dfs)
combined_df.rename(columns={0: "Index", 1: "Type", 2: "Group", 3: "Day", 4: "Time", 5: "Venue", 6: "Remark"},
                   inplace=True)

#xlsxwriter
writer = pd.ExcelWriter("Schedule.xlsx", engine="xlsxwriter")
combined_df.to_excel(writer, index=False)
workbook = writer.book
worksheet = writer.sheets["Sheet1"]
#set column width
worksheet.set_column(0, 0, 7)
worksheet.set_column(1, 1, 6)
worksheet.set_column(2, 2, 11)
worksheet.set_column(3, 3, 6)
worksheet.set_column(4, 4, 4)
worksheet.set_column(5, 5, 11)
worksheet.set_column(6, 6, 11)
worksheet.set_column(7, 7, 20)
writer.save()

holder = input("Process finished. Hit Enter to continue...")
