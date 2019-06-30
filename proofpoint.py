# By Chase Moxley Created on 2/21/2018 version 1.0
###This is used to automate meterics from proofpoint



from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import smtplib
from configparser import ConfigParser
import os

if __name__ == '__main__':
# Config for the chrome webdriver
     chrome_options = Options()
     chrome_options.add_argument("--headless")
     chrome_options.add_argument("--window-size=1920x1080")
# set var for username/password

     stuff2 = os.path.abspath('config.ini')
     with open(stuff2, 'r') as stuff:
          config = ConfigParser()
          config.readfp(stuff)
usernameStr = config.get('auth','username')
passwordStr = config.get('auth','password')
path1 = config.get('instance','path1')
url_1 = config.get('instance','url_1')
url_2 = config.get('instance','url_2')
url_2 = config.get('instance','url_2')
uno = config.get('email','uno')
smtp1 = config.get('email','smtp')
from1 = config.get('email','from')
# Chrome web browser set it to a var
browser = webdriver.Chrome(chrome_options=chrome_options,executable_path= path1)
# starts chrome browser and launches URL
browser.get((url_1))
#this passes the username/password vars to the id with the webpage. It also selects the login and clicks the button
username = browser.find_element_by_id('user')
username.send_keys(usernameStr)
passwd = browser.find_element_by_id('pass')
passwd.send_keys(passwordStr)
nextButton = browser.find_element_by_id('login')
nextButton.click()
# Goes to the Proofpoint report that is in text 
browser.get((url_2))
# gets page source and sets it to a var
text = browser.page_source
#closes browser
browser.quit()
# assigns new var after cleaning up some of the text with the source
text2 = text.replace("<html xmlns=\"http://www.w3.org/1999/xhtml\"><head></head><body><pre style=\"word-wrap: break-word; white-space: pre-wrap;\">","").replace("Top Actions","").replace("</pre></body></html>","")
# assigns the cleaned up version to a new var
string = text2
#  the string var gets split and assigned to a new var 
splitstring = string.split()

listolists = []
for i in range(0, len(splitstring), 2):
     x = []
     x.append(splitstring[i])
     x.append(splitstring[i + 1])
     listolists.append(x)
# Create dateframe with pandas with the listolists and assign columns
df22 = pd.DataFrame(listolists, columns=['Type','Volume'])
# create index in pandas 
df22 = df22.set_index('Type')
#select cells with the dataframe
cell01 = df22.at['reject','Volume']
cell02 = df22.at['discard','Volume']
#add cells togeather and declare them as ints. 
total = int(cell01) + int(cell02)
# Print results
co_cell01 = str(cell01)
co_cell02 = str(cell02)
co_total = str(total)
#Send email with results
from_addr = from1
to_addr = uno
msg = MIMEMultipart()
msg['From'] = from_addr
msg['To'] =  to_addr
msg['Subject'] = 'Proofpoint Metrics'
body = "Rejected: %s + Discarded %s = Total %s" %(co_cell01,co_cell02,co_total)
msg.attach(MIMEText(body, 'plain'))
text55 = msg.as_string()
server = smtplib.SMTP(smtp1, 25)
server.sendmail(from_addr,to_addr, text55 )
server.quit()
