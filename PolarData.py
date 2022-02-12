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
load_dotenv()                    #for python-dotenv method

now = datetime.datetime.now()
year = '{:04d}'.format(now.year)
month = '{:02d}'.format(now.month)
day = '{:02d}'.format(now.day)

POLAR_EMAIL = os.environ.get('POLAR_EMAIL')
POLAR_PASSWORD = os.environ.get('POLAR_PASSWORD')
DOWNLOADS_PATH = os.environ.get('DOWNLOADS_PATH')
SESSIONS = 0

Roster = ['Arthur Miller',
        'Beck Gozdenovich',
        'Billy Kroeger',
        'Brandon Wasitowski',
        'Brendan Falatko',
        'Brendan Geissel',
        'Charlie Robbins',
        'Cole Ward',
        'Collin Fogerty',
        'Colton Johnson',
        'Colton Jones',
        'Darren Romaine',
        'Devon Ford',
        "Dominick O'Melia",
        'Eric Sherman',
        'Ethan Kaban',
        'Gage Adams',
        'Garrett Muscatella',
        'Gavyn Willson',
        'Guiseppe Chiovera',
        'Isaac Vanzomeren',
        'Jack Bowie',
        'Jack Mahony',
        'Jake Baumgardt',
        'Joey Schwarz',
        'Josh Failia',
        'Keegan Ford',
        'Liam Brown',
        'Logan Hone',
        'Lynch Raby',
        'Matt Bowerman',
        'Max Wilson',
        'Myles Hickey',
        'Nick Gutierrez',
        'Owen Corry',
        'Richie Nardone',
        'Teddy Grimley',
        'Tyler Nolan',
        'Will Labartino',
        'Xavier Ritter',
        'Yanni Kalas'
         ]
def FindAndClick(locator,value):
    # Find button
    button = driver.find_element(locator,value)
    # Click button
    button.click()
def setURL():
    return f"https://teampro.polar.com/diary?period=day&date={year}-{month}-10"

def startWeb():
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : DOWNLOADS_PATH}
    chrome_options.add_experimental_option('prefs', prefs)

    s=Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=s,options=chrome_options)

def getLinks(driver):
    driver.get(setURL())
    #let page load
    time.sleep(2)
    #find user and password box
    id_box = driver.find_element(By.ID, "email")
    password_box = driver.find_element(By.ID, "password")

    id_box.send_keys(POLAR_EMAIL)
    password_box.send_keys(POLAR_PASSWORD)
    FindAndClick(By.ID,'login')

    time.sleep(2)

    driver.switch_to.new_window('tab')
    driver.get(setURL())

    time.sleep(2)

    links = []
    elements = driver.find_elements(By.TAG_NAME,"a")
    for e in elements:
        if 'min' in e.text:
             links.append(e.get_attribute("href"))
    return links 
def exportLinks(links):
    print(str(SESSIONS),' training sessions are being downloaded')
    k = 1
    for link in links:
        
        driver.switch_to.new_window('tab')
        driver.get(link)
        FindAndClick(By.CLASS_NAME,'export')
        time.sleep(2)
        # Find export buttons
        export_buttons = driver.find_elements(By.CLASS_NAME,'btn.btn-primary.session')

        for i in export_buttons:
            if i.text == 'Export as csv':
                i.click()   
        os.chdir(DOWNLOADS_PATH)
        #check if file is downloaded every 2 seconds to optimize runtime.
        #max runtime of this portion is 200 seconds. 
        print('Training session',k,'is being downloaded')
        for z in range(0,100):
            if len(glob.glob(f'{year}{month}10_*_Lacrosse_Rugby.csv'))>0:
                break
            time.sleep(2)
        formatSession(k)
        k+=1

def formatSession(k):
    print('Training session',k,'is being reformated')
    export_df = pd.DataFrame(Roster,columns = ['Player name'])
    export_df[['Total distance [m]','Training load score','Muscle load', 'Cardio load']] = 0

    os.chdir(DOWNLOADS_PATH)
    for file in glob.glob(f'{year}{month}10_*_Lacrosse_Rugby.csv'):
        import_df = pd.read_csv( DOWNLOADS_PATH+file)
    import_df = import_df[['Player name','Total distance [m]','Training load score','Muscle load', 'Cardio load']]
    export_df = import_df.append(export_df[~export_df['Player name'].isin(import_df['Player name'].values.tolist())])
    export_df.sort_values('Player name',inplace=True)
    export_df.reset_index(drop=True, inplace = True)
    #export file as excel file
    excelfile = f'LAX{year}{month}{day}{k}.xlsx'
    export_df.to_excel(r'C:\\Users\\Rich\\Desktop\\polar\\ExcelFiles\\'+excelfile, index = False)
    print('Training session',k,' is being removed from csv files')
    os.remove(DOWNLOADS_PATH+file) 
    print('Training session',k,'has been converted to EXCEL')


driver = startWeb()
links = getLinks(driver)
SESSIONS = len(links)
if(len(links)>0):
    exportLinks(links)
else:
    print("No Polar data has not been uploaded today") 
time.sleep(5)
driver.quit() 
print('ENDED SUCCESSFULLY')
