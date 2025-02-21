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
 


# set directory
github_workspace = os.path.join(os.getenv('GITHUB_WORKSPACE'), 'data')
default_dir = os.getcwd()
# when running locally:
#github_workspace = 'C:/Users/AhyoungLim/Dropbox/WORK/OpenDengue/PAHO-crawler/data'
#default_dir = 'C:/Users/AhyoungLim/'  

today_directory_name = f"Image_DL_{datetime.now().strftime('%Y%m%d')}"
downloadPath = os.path.join(github_workspace, today_directory_name)
os.makedirs(downloadPath, exist_ok=True) # create a new directory 

# set chrome download directory
chrome_options = uc.ChromeOptions()
prefs = {"download.default_directory": os.getcwd()}
chrome_options.add_experimental_option("prefs", prefs)

# using undetected-chromedriver
driver = uc.Chrome(headless=True, use_subprocess=False, options = chrome_options, version_main=chrome_version)     
driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')


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

image = shadow_doc2.find_element(By.TAG_NAME, "img")

# Resize the browser window to match the iframe's dimensions
driver.set_window_size(image.size['width']*1.2, image.size['height']*1.2)

# Now take the screenshot of the image 
image.screenshot("image_in_iframe_resized.png")

# Use the move_to_download_folder function to move the downloaded file
today = datetime.now().strftime('%Y%m%d%H%m') # current date and time
newFileName = f"PAHO_ByLastAvailableEpiWeek_{today}"  # Base filename
fileExtension = '.png'  # File extension

move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)

 
driver.quit()

