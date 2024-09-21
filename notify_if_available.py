# import
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import time
import smtplib
from email.mime.text import MIMEText
from pretty_html_table import build_table

# setup
import os
from pyvirtualdisplay import Display
Display(visible = 0, size = (1920, 1080)).start()

# window
driver = webdriver.Chrome()
driver.implicitly_wait(3)
achains = ActionChains(driver)

# url
driver.maximize_window()
driver.get("https://ais.usvisa-info.com/en-ca/niv/users/sign_in")

# user
elem = driver.find_element(By.ID, "user_email")
elem.send_keys(os.getenv("PORTAL_USER"))

# pass
elem = driver.find_element(By.ID, "user_password")
elem.send_keys(os.getenv("PORTAL_PASS"))

# terms
elem = driver.find_element(By.ID, "policy_confirmed")
achains.move_to_element(elem).click().perform()

# submit
elem = driver.find_element(By.NAME, "commit")
achains.move_to_element(elem).click().perform()

# continue
elem = driver.find_element(By.XPATH, ".//a[@class='button primary small']")
elem.click()

# reschedule 1
elem = driver.find_element(By.LINK_TEXT, "Reschedule Appointment")
elem.click()

# reschedule 2
time.sleep(3)
elems = driver.find_elements(By.LINK_TEXT, "Reschedule Appointment")
elems[1].click()

# acc.
date_df = pd.DataFrame(columns = ["consulate", "closest_date", "report_time"])
consulates = ["Calgary", "Halifax", "Montreal", "Ottawa", "Quebec City", "Toronto", "Vancouver"]

# consulate
for consulate in consulates: 
    elem = Select(driver.find_element(By.NAME, "appointments[consulate_appointment][facility_id]"))
    elem.select_by_visible_text(consulate)
    
    # calendar
    elem = driver.find_element(By.NAME, "appointments[consulate_appointment][date]")
    try: elem.click()
    except: continue
    
    # soup
    while(1):
        soup = BeautifulSoup(driver.page_source, "html.parser").find_all("td", attrs = {"data-handler": "selectDay"})
        # availability
        if len(soup) > 0: break
        # next month
        elem = driver.find_element(By.XPATH, ".//a[@title='Next']")
        elem.click()
    
    # record
    closest_date = datetime.strptime(soup[0]["data-year"] + "-" + soup[0]["data-month"] + "-" + soup[0].get_text(), "%Y-%m-%d").strftime("%Y-%m-%d")
    if closest_date < "2027-12-08": date_df = pd.concat([date_df, pd.DataFrame([[consulate, closest_date, (datetime.now() - timedelta(hours = 4)).strftime("%Y-%m-%d %I:%M %p")]], columns = date_df.columns)], ignore_index = True)

# email - from, to, body
sender_email = "shithi30@gmail.com"
# receiver_email = ["maitra.shithi.aust.cse@gmail.com", "shithi30@outlook.com", "Purnabchowdhury@gmail.com"]
receiver_email = ["maitra.shithi.aust.cse@gmail.com", "shithi30@outlook.com"]
body = '''Please find the earliest posted empty slots by consulates.''' + build_table(date_df, "green_dark", font_size = "12px", text_align = "left") + '''<br>Thanks,<br>Shithi Maitra<br>Ex Asst. Manager, CS Analytics<br>Unilever BD Ltd.<br>'''

# email - object
html_msg = MIMEText(body, "html")
html_msg["Subject"], html_msg["From"], html_msg["To"] = "US Visa Appt.", "Shithi Maitra", ""

# email - send
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(sender_email, os.getenv("EMAIL_PASS"))
    if date_df.shape[0] > 0: server.sendmail(sender_email, receiver_email, html_msg.as_string())
