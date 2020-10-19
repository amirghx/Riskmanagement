import requests
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import shutil

# find path we are wotking on
path = os.getcwd()

# create data folder

Created_path = path + '\data'
try:
    os.mkdir(Created_path)
except OSError:
    print("Creation of the directory %s failed" % path)
else:
    print("Successfully created the directory %s " % path)

# download Asas excel file from site

url = 'https://asemanetf.com/Download/DownloadNavChartList?exportType=Excel'
r = requests.get(url, allow_redirects=True)
open(Created_path + '\Asas Detail.xlsx', 'wb').write(r.content)

url = 'http://mellatetf.com/Download/DownloadNavChartList?exportType=Excel'
r = requests.get(url, allow_redirects=True)
open(Created_path + '\ofoghmelat Detail.xlsx', 'wb').write(r.content)

url = 'https://iran-kfunds3.ir/Download/DownloadNavChartList?exportType=Excel'
r = requests.get(url, allow_redirects=True)
open(Created_path + '\kardan Detail.xlsx', 'wb').write(r.content)

url = 'https://iranetf.ir/Download/DownloadNavChartList?exportType=Excel'
r = requests.get(url, allow_redirects=True)
open(Created_path + '\Firooze Detail.xlsx', 'wb').write(r.content)

url = 'https://atlasetf.com/Download/DownloadNavChartList?exportType=Excel'
r = requests.get(url, allow_redirects=True)
open(Created_path + '\Atlas Detail.xlsx', 'wb').write(r.content)

op = Options()
op.add_argument('--disable-notifications')
op.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

# Opening the page
driver.get("http://www.hbagahfund.com/Home/Nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"Agas details.xlsx"))
# Closing the webdriver
driver.close()

driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

driver.get("http://www.scetf.ir/Home/Nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"karis details.xlsx"))
# Closing the webdriver
driver.close()

driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

driver.get("http://sarv.fund/Home/Nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"sarv details.xlsx"))
# Closing the webdriver
driver.close()

driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

driver.get("http://www.armanmesetf.com/home/nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"Atimes details.xlsx"))
# Closing the webdriver
driver.close()

driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

driver.get("http://www.bazreomidfund.ir/Home/Nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"Bazr details.xlsx"))
# Closing the webdriver
driver.close()

driver = webdriver.Chrome(options=op)

# Setting Chrome to trust downloads
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': Created_path}}
command_result = driver.execute("send_command", params)

driver.implicitly_wait(5)

driver.get("http://fardaetf.tadbirfunds.com/Home/Nav")
time.sleep(10)
# Click on the button and wait for 10 seconds
driver.find_element_by_xpath('//*[@class="dt-button buttons-excel buttons-html5"]').click()
time.sleep(10)
filename = max([Created_path + "\\" + f for f in os.listdir(Created_path)], key=os.path.getctime)
shutil.move(filename, os.path.join(Created_path, r"Almas details.xlsx"))
# Closing the webdriver
driver.close()
