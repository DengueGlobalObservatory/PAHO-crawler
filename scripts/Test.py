import undetected_chromedriver as uc
#from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome.service import Service
import time
import os
from datetime import datetime

def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False   
    while not got_file:
        try: 
            # Use glob to get the current file name
            currentFile = max([default_dir + "/" + f for f in os.listdir(default_dir)], key=os.path.getctime)

            #if len(currentFiles) > 0:
            #currentFile = currentFiles[0]  # Assuming there's only one file matching the pattern
            got_file = True
        except Exception as e:
            print("File has not finished downloading")
            time.sleep(5)

    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)

    # Move the file
    os.rename(currentFile, fileDestination)
    print(f"Moved file to {fileDestination}")


today = datetime.now().strftime('%Y%m%d%H%m') # current date and time

# set directory
# default_dir = 'C:/Users/AhyoungLim/Downloads'  
github_workspace = os.path.join(os.getenv('GITHUB_WORKSPACE'), 'data')
default_dir = os.getcwd()

today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
downloadPath = os.path.join(github_workspace, today_directory_name)
os.makedirs(downloadPath, exist_ok=True) # create a new directory 


chrome_options = uc.ChromeOptions()

prefs = {
    "download.default_directory": os.getcwd()
}

chrome_options.add_experimental_option("prefs", prefs)

# using undetected-chromedriver
driver = uc.Chrome(headless=True, use_subprocess=False, options = chrome_options)
driver.get('https://worldhealthorg.shinyapps.io/dengue_global/')

print(driver.title)

# Click and download --------------------------------------------------------------
# Wait for the "I accept" button to be clickable and then click it
accept_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, "closeModal"))
)
driver.execute_script("arguments[0].scrollIntoView();", accept_button)
driver.execute_script("arguments[0].click();", accept_button)

#accept_button.click()


# Find and click the "Download Data" link in the menu
download_link = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, "//a[@data-value='dl_data']"))
)
driver.execute_script("arguments[0].scrollIntoView();", download_link)
driver.execute_script("arguments[0].click();", download_link)

time.sleep(5)

# Click the button to download all data
download_all_data_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "dl_all_data"))
)
driver.execute_script("arguments[0].scrollIntoView();", download_all_data_button)
driver.execute_script("arguments[0].click();", download_all_data_button)

time.sleep(5)

# Use the move_to_download_folder function to move the downloaded file
newFileName = "newFileName"  # Base filename
fileExtension = '.csv'  # File extension
move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)

