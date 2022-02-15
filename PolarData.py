from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import datetime
import glob
import os
import pandas as pd
from dotenv import load_dotenv   #for python-dotenv method
load_dotenv()      

now = datetime.datetime.now()
year = '{:04d}'.format(now.year)
month = '{:02d}'.format(now.month)
day = '{:02d}'.format(now.day)

POLAR_PATH =  os.environ.get('POLAR_PATH')
POLAR_EMAIL = os.environ.get('POLAR_EMAIL')
POLAR_PASSWORD = os.environ.get('POLAR_PASSWORD')
DOWNLOADS_PATH = os.environ.get('DOWNLOADS_PATH')

def CheckForSetup():
    return os.path.exists('.env')

def CreateData(user,passw,dpath,ppath):
    f = open(".env","x")
    f.write('POLAR_EMAIL='+user+'\nPOLAR_PASSWORD='+passw+'\nDOWNLOADS_PATH='+dpath+'\nPOLAR_PATH='+ppath)
    return 0


SESSIONS = 0
files = {}
def FindAndClick(driver,locator,value):
    # Find button
    button = driver.find_element(locator,value)
    # Click button
    button.click()
    
def setURL():
    return f"https://teampro.polar.com/diary?period=day&date={year}-{month}-{day}"

def startWeb():
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : DOWNLOADS_PATH}
    chrome_options.add_experimental_option('prefs', prefs)

    s=Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=s,options=chrome_options)

def getLinks(driver):
    driver.get(setURL())
    driver.set_window_position(0, 0)
    driver.set_window_size(300, 300)
    #let page load
    time.sleep(2)
    #find user and password box
    id_box = driver.find_element(By.ID, "email")
    password_box = driver.find_element(By.ID, "password")

    id_box.send_keys(POLAR_EMAIL)
    password_box.send_keys(POLAR_PASSWORD)
    FindAndClick(driver,By.ID,'login')

    time.sleep(2)

    driver.switch_to.new_window('tab')
    driver.get(setURL())

    time.sleep(3)

    links = []
    elements = driver.find_elements(By.TAG_NAME,"a")
    for e in elements:
        if 'min' in e.text:
             links.append(e.get_attribute("href"))
    return links 

def exportLinks(driver,links):
    time.sleep(3)
    print(str(SESSIONS),' training sessions are being downloaded')
    k = 1
    for link in links:
        driver.switch_to.new_window('tab')
        driver.get(link)
        FindAndClick(driver,By.CLASS_NAME,'export')
        time.sleep(2)
        # Find export buttons
        export_buttons = driver.find_elements(By.CLASS_NAME,'btn.btn-primary.session')

        for i in export_buttons:
            if i.text == 'Export as csv':
                i.click()   
        print('Training session',k,'is being downloaded')
        os.chdir(DOWNLOADS_PATH)
        while (True):
            if len(glob.glob(f'{year}{month}10_*_Lacrosse_Rugby.csv'))>0:
                break
            time.sleep(2)
        formatSession(k)
        k+=1

def formatSession(k):
    print('Training session',k,'is being reformated')
    export_df = pd.read_csv(POLAR_PATH+'Roster.csv',)
    export_df[['Total distance [m]','Training load score','Muscle load', 'Cardio load']] = 0
    os.chdir(DOWNLOADS_PATH)
    for file in glob.glob(f'{year}{month}{day}_*_Lacrosse_Rugby.csv'):
        import_df = pd.read_csv(DOWNLOADS_PATH+file)
    
    import_df = import_df[['Player name','Total distance [m]','Training load score','Muscle load', 'Cardio load']]
    export_df = import_df.append(export_df[~export_df['Player name'].isin(import_df['Player name'].values.tolist())])
    export_df.sort_values('Player name',inplace=True)
    export_df.reset_index(drop=True, inplace = True)

    #export file as excel file
    
    excelfile = f'MLAX-{year}-{month}-{day}-Session-{k}.xlsx'
    export_df.to_excel(POLAR_PATH+excelfile, index = False)
    print('Training session',k,'is being removed from csv files')
    os.remove(DOWNLOADS_PATH+file) 
    print('Training session',k,'has been converted to EXCEL')
    files[excelfile] = import_df.shape[0]

def Start():
    driver = startWeb()
    links = getLinks(driver)
    SESSIONS = len(links)
    if(SESSIONS>0):
        exportLinks(driver,links)
    else:
        print("No Polar data has not been uploaded today") 
    driver.quit() 
    print('ENDED SUCCESSFULLY')
    return files
