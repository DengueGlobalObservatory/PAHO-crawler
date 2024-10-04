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
import subprocess
import re
import sys

#from webdriver_manager.chrome import ChromeDriverManager
#driver_executable_path = ChromeDriverManager().install()
#https://github.com/ultrafunkamsterdam/undetected-chromedriver/issues/1904

# Function to get the installed Chrome version
def get_chrome_version():
    try:
        if sys.platform == "win32":
            # Command to retrieve Chrome version from Windows registry
            output = subprocess.check_output(
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                shell=True, 
                text=True
            )
            version = re.search(r'\s+version\s+REG_SZ\s+(\d+)\.', output)
        else:  # Assuming Linux or other Unix-like systems
            # Try different commands to retrieve Chrome version on Linux
            for command in ['google-chrome --version', 'google-chrome-stable --version', 'chromium-browser --version']:
                try:
                    output = subprocess.check_output(command, shell=True, text=True)
                    version = re.search(r'\b(\d+)\.', output)
                    if version:
                        break
                except subprocess.CalledProcessError:
                    continue
            else:
                raise RuntimeError("Could not determine Chrome version")

        if version:
            return int(version.group(1))
        else:
            raise ValueError("Could not parse Chrome version")
    except Exception as e:
        raise RuntimeError("Failed to get Chrome version") from e

# Get the major version of Chrome installed
chrome_version = get_chrome_version()

def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False   
    while not got_file:
        try: 
            # Use glob to get the current file name
            currentFile = max([default_dir + "/" + f for f in os.listdir(default_dir)], key=os.path.getctime)

           # Ensure the file exists before proceeding
            if os.path.exists(currentFile):
                got_file = True
            else:
                raise FileNotFoundError("File not found. Retrying...")
            
        except Exception as e:
            print("File has not finished downloading")
            time.sleep(10)

    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)

    # Move the file
    os.rename(currentFile, fileDestination)
    print(f"Moved file to {fileDestination}")
 

def download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today):
    """Download and rename the file for the given week number."""
    
    # Wait for the week number to update
    weeknum_div = wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
    )
    weeknum = int(weeknum_div.text)  # Convert to integer for comparison

    # print(f"Week Number: {weeknum}")

    # Find and click the download button at the bottom of the dashboard
    download_button = wait.until(
        EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
    )
    download_button.click()
    
    time.sleep(5)

    # Find and click the crosstab button (in a pop up window)
    crosstab_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
    )
    crosstab_button.click()
    
    time.sleep(5)

    # Find and select the CSV option
    #csv_div = wait.until(
    #    EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='radio'][value='csv']"))
    #)
    csv_div = shadow_doc2.find_element(By.CSS_SELECTOR, "input[type='radio'][value='csv']")
    driver.execute_script("arguments[0].scrollIntoView();", csv_div)
    driver.execute_script("arguments[0].click();", csv_div)
    time.sleep(5)

    # Find and click the export button
    export_button = shadow_doc2.find_element(By.CSS_SELECTOR, '[data-tb-test-id="export-crosstab-export-Button"]')
    export_button.click()
    print("Downloading CSV file")
    time.sleep(5)

    # Use the move_to_download_folder function to move the downloaded file
    downloadPath = downloadPath
    default_dir = default_dir
    newFileName = f"PAHO_{year}_W{weeknum}_{today}"  # Base filename
    fileExtension = '.csv'  # File extension

    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def iterate_weekly(): 
    year = "2023" # choose year to download
    today = datetime.now().strftime('%Y%m%d%H%m') # current date and time

    # set directory
    # github_workspace = os.path.join(os.getenv('GITHUB_WORKSPACE'), 'data')
    # default_dir = os.getcwd()
    github_workspace = 'C:/Users/AhyoungLim/Dropbox/WORK/OpenDengue/PAHO-crawler/data'
    default_dir = 'C:/Users/AhyoungLim/Downloads'  

    today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
    downloadPath = os.path.join(github_workspace, today_directory_name)
    os.makedirs(downloadPath, exist_ok=True) # create a new directory 

    # set chrome download directory
    chrome_options = uc.ChromeOptions()
    prefs = {"download.default_directory": os.getcwd()}
    chrome_options.add_experimental_option("prefs", prefs)

    # using undetected-chromedriver
    driver = uc.Chrome(headless=False, use_subprocess=False, options = chrome_options, version_main=chrome_version)     
    driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')

    #driver = webdriver.Chrome(service=Service(), options=chrome_options)  # Ensure chrome_options is defined
    #driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')
    
    # Define wait outside the loop
    wait = WebDriverWait(driver, 20)

    # First iframe
    iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
    iframe_locator = (By.XPATH, "//div[contains(@class, 'vizTab')]//iframe[@src='" + iframe_src + "']")
    iframe = wait.until(EC.presence_of_element_located(iframe_locator))
    driver.switch_to.frame(iframe)

    # Grab the shadow element
    shadow = driver.execute_script('return document')

    # Get the iframe inside shadow element of first iframe
    iframe2 = shadow.find_element(By.XPATH, "//body/iframe")
    driver.switch_to.frame(iframe2)
    shadow_doc2 = driver.execute_script('return document')

    iframe_page_title = driver.title
    print(iframe_page_title)    

    if iframe_page_title != "PAHO/WHO Data - National Dengue fever cases": 
        print("Wrong access")
        driver.quit()

    time.sleep(3)

    # find the year tab
    year_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId13')))

    # find the dropdown button within the year tab
    dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
    dd_open = year_tab.find_element(*dd_locator)
    dd_open.click()   
    
    # select year 2024
    y2024_xpath = '//div[contains(@class, "facetOverflow")]/a[text()="2024"]/preceding-sibling::input'
    shadow_doc2.find_element(By.XPATH, y2024_xpath).click()
    
    # select the year of interest
    # year_xpath = f'//div[contains(@class, "facetOverflow")]//a[text()="{year}"]/preceding-sibling::input'
    # shadow_doc2.find_element(By.XPATH, year_xpath).click()

    # close the dropdown menu
    dd_close = wait.until(
        EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))
    )
    dd_close.click()

    time.sleep(3)
        
    # Initial call to download_and_rename (for week 53 only)
    print(f"Processing Week Number: 53")
    download_and_rename(wait, shadow_doc2, 53, default_dir, downloadPath, driver, year, today)
    

    weeknum = 53  # Initialize weeknum outside the loop
    # loop for downloading and renaming files
    while weeknum > 0:
        print(f"Processing Week Number: {weeknum-1}")
        decrement_button = wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'tableauArrowDec')]")))
        decrement_button.click()
        time.sleep(3)

        # Update weeknum after decrementing
        weeknum -= 1

        # Pass updated weeknum to download_and_rename
        download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today)

        if weeknum == 1:
            print("Reached week 1, breaking the loop.")
            break

    driver.quit()

iterate_weekly()