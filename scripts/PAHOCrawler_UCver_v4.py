import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import os
from datetime import datetime
import subprocess
import re
import sys

# Define a global default timeout variable
DEFAULT_WAIT_TIMEOUT = 60 # Seconds

# Function to get the installed Chrome version
def get_chrome_version():
    try:
        if sys.platform == "win32":
            output = subprocess.check_output(
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                shell=True,
                text=True
            )
            version = re.search(r'\s+version\s+REG_SZ\s+(\d+)\.', output)
        else:
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
        raise RuntimeError(f"Failed to get Chrome version: {e}")

# Get the major version of Chrome installed
chrome_version = get_chrome_version()

# Improved move_to_download_folder function
def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    start_time = time.time()
    timeout = 180 # 3 minutes timeout for download
    while not got_file and (time.time() - start_time) < timeout:
        try:
            files = [os.path.join(default_dir, f) for f in os.listdir(default_dir) if os.path.isfile(os.path.join(default_dir, f))]

            downloading_files = [f for f in files if f.endswith('.crdownload') or f.endswith('.tmp') or f.endswith('.part') or f.endswith('.download')]

            if downloading_files:
                print(f"Waiting for download to complete: {downloading_files[0]}")
                time.sleep(5)
                continue

            finished_files = [f for f in files if not f.endswith('.crdownload') and not f.endswith('.tmp') and not f.endswith('.part') and not f.endswith('.download')]
            if not finished_files:
                if not files and (time.time() - start_time) > 10:
                     raise FileNotFoundError("No files found in the download directory yet.")
                else:
                    time.sleep(5)
                    continue

            currentFile = max(finished_files, key=os.path.getctime)

            if os.path.exists(currentFile) and not (currentFile.endswith('.crdownload') or currentFile.endswith('.tmp') or currentFile.endswith('.part') or currentFile.endswith('.download')):
                got_file = True
            else:
                raise FileNotFoundError("File not found or still downloading. Retrying...")

        except FileNotFoundError as e:
            print(f"File not yet fully downloaded or visible: {e}. Retrying in 10 seconds...")
            time.sleep(10)
        except Exception as e:
            print(f"An unexpected error occurred while waiting for file: {e}. Retrying in 10 seconds...")
            time.sleep(10)

    if not got_file:
        raise TimeoutError(f"Download did not complete within {timeout} seconds.")

    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)
    os.rename(currentFile, fileDestination)
    print(f"Moved file from {currentFile} to {fileDestination}")


# `shadow_doc2` parameter now consistently refers to the driver object within the iframe
def download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver_main_page_ignored, year, today):
    """Download and rename the file for the given week number."""

    print(f"Attempting to download for Week Number: {weeknum}")

    # Scroll and then interact with the week number slider.
    # The week number slider is at the top, so it should already be visible,
    # but let's be explicit for consistency.
    weeknum_div_locator = (By.CLASS_NAME, "sliderText")
    weeknum_div = wait.until(EC.presence_of_element_located(weeknum_div_locator))
    shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", weeknum_div)
    time.sleep(0.5) # Give it a moment to scroll

    weeknum = int(weeknum_div.text)
    print(f"Confirmed Week Number Displayed: {weeknum}")

    # The 'Tabla Anual' tab is usually at the top, but let's ensure it's in view
    table_tab_locator = (By.XPATH, "//span[@class='tabLabel' and text()='Tabla Anual']")
    try:
        table_tab = wait.until(EC.element_to_be_clickable(table_tab_locator))
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", table_tab)
        time.sleep(0.5)
        table_tab.click()
        print("Clicked 'Tabla Anual' tab.")
        time.sleep(3)
    except Exception as e:
        print(f"Could not click 'Tabla Anual' tab: {e}. Ensure this tab is present and visible.")
        raise

    # Download button is at the bottom, so we definitely need to scroll to it.
    download_button_id = "download-ToolbarButton"
    download_button = wait.until(EC.presence_of_element_located((By.ID, download_button_id)))

    try:
        print(f"Attempting to click download button ({download_button_id}) via JS...")
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_button)
        time.sleep(0.5)
        shadow_doc2.execute_script("arguments[0].click();", download_button)
        print("Clicked download button successfully via JS.")
        time.sleep(5)
    except Exception as e:
        print(f"Failed to click download button ({download_button_id}) via JS: {e}")
        screenshot_path = os.path.join(downloadPath, f"debug_download_button_fail_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path)
        print(f"Screenshot taken to help debug: {screenshot_path}")
        raise

    # Crosstab button appears in a pop-up, so it needs to be made visible
    crosstab_button_locator = (By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]')
    crosstab_button = wait.until(EC.element_to_be_clickable(crosstab_button_locator))
    shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", crosstab_button)
    time.sleep(0.5)
    crosstab_button.click()
    print("Clicked Crosstab button.")
    time.sleep(5)

    # CSV option is in the pop-up, ensure it's in view
    csv_div_locator = (By.CSS_SELECTOR, "input[type='radio'][value='csv']")
    csv_div = wait.until(EC.presence_of_element_located(csv_div_locator))
    shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", csv_div)
    time.sleep(0.5)
    shadow_doc2.execute_script("arguments[0].click();", csv_div)
    print("Selected CSV option.")
    time.sleep(5)

    # Export button is in the pop-up, ensure it's in view
    export_button_locator = (By.CSS_SELECTOR, '[data-tb-test-id="export-crosstab-export-Button"]')
    export_button = wait.until(EC.presence_of_element_located(export_button_locator))

    try:
        print("Attempting to scroll export button into view and click via JS...")
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_button)
        time.sleep(0.5)
        shadow_doc2.execute_script("arguments[0].click();", export_button)
        print("Downloading CSV file")
        time.sleep(5)
    except Exception as e:
        print(f"Failed to click export button even with scrolling and JS click: {e}")
        screenshot_path = os.path.join(downloadPath, f"debug_export_button_click_fail_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path)
        print(f"Screenshot taken to help debug: {screenshot_path}")
        raise

    downloadPath = downloadPath
    default_dir = default_dir
    newFileName = f"PAHO_2023_2025_W{weeknum}_{today}"
    fileExtension = '.csv'

    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def iterate_weekly():

    year = "(All)"
    today = datetime.now().strftime('%Y%m%d%H%M')

    github_workspace = os.getenv('GITHUB_WORKSPACE')
    if github_workspace:
        base_data_path = os.path.join(github_workspace, 'data')
    else:
        base_data_path = os.path.join(os.getcwd(), 'data')
        print(f"GITHUB_WORKSPACE environment variable not set. Using: {base_data_path}")

    default_dir = os.path.join(os.getcwd(), "temp_downloads")
    os.makedirs(default_dir, exist_ok=True)

    today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
    downloadPath = os.path.join(base_data_path, today_directory_name)
    os.makedirs(downloadPath, exist_ok=True)

    chrome_options = uc.ChromeOptions()
    prefs = {
        "download.default_directory": default_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--single-process")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-setuid-sandbox")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--no-zygote")


    driver = uc.Chrome(headless=True, use_subprocess=False, options=chrome_options, version_main=chrome_version)
    driver.get('https://www.paho.org/en/arbo-portal/dengue-data-and-analysis/dengue-analysis-country')

    wait = WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT)

    iframe_page_title = driver.title
    print(f"Current page title: {iframe_page_title}")

    if "Dengue: analysis by country" not in iframe_page_title:
        print("Wrong access: Page title does not match expected. Exiting.")
        driver.quit()
        sys.exit(1)
    else:
        print("Page title confirmed.")

    time.sleep(7)

    iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008_NAC_ES_Indicadores_reporte_semanal.asp"
    iframe_locator = (By.XPATH, f"//iframe[@src='{iframe_src}']")

    try:
        print(f"Attempting to switch to iframe with src: {iframe_src}")
        iframe_element_on_main_page = wait.until(EC.presence_of_element_located(iframe_locator))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", iframe_element_on_main_page)
        time.sleep(2)

        driver.switch_to.frame(iframe_element_on_main_page)
        print("Switched to iframe successfully.")

        shadow_doc2 = driver

    except TimeoutException as e:
        print(f"Timeout: Iframe with src '{iframe_src}' not found within {DEFAULT_WAIT_TIMEOUT} seconds. Error: {e}")
        screenshot_path = os.path.join(downloadPath, f"debug_iframe_timeout_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot taken: {screenshot_path}")
        driver.quit()
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred while switching to iframe: {e}")
        driver.quit()
        sys.exit(1)

    # Confirm dashboard elements loaded within the iframe
    try:
        time.sleep(7)
        dashboard_viewport = wait.until(EC.presence_of_element_located((By.ID, "dashboard-viewport")))
        print("Dashboard viewport found inside iframe, presuming dashboard loaded correctly.")

        # --- REMOVED: Blanket scroll to bottom, replaced by targeted scrolls ---
        # print("Scrolling to the bottom of the iframe content...")
        # shadow_doc2.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # time.sleep(3)
        # print("Scrolled to bottom.")

    except TimeoutException as e:
        print(f"Timeout: Dashboard viewport not found inside iframe within {DEFAULT_WAIT_TIMEOUT} seconds. Error: {e}")
        screenshot_path = os.path.join(downloadPath, f"debug_dashboard_viewport_timeout_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path)
        print(f"Screenshot taken: {screenshot_path}")
        driver.quit()
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred after switching to iframe: {e}")
        driver.quit()
        sys.exit(1)

    time.sleep(5)

    year_filter_container_id = 'tabZoneId16'

    try:
        # Year filter is at the top, scroll to it if necessary
        year_filter_container = wait.until(EC.presence_of_element_located((By.ID, year_filter_container_id)))
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", year_filter_container)
        time.sleep(0.5)

        print(f"Year filter container ({year_filter_container_id}) is present in iframe DOM.")

        year_dropdown_opener_locator = (By.XPATH, f"//div[@id='{year_filter_container_id}']//div[contains(@class, 'tabComboBoxNameContainer')]")
        year_dropdown_opener = wait.until(EC.visibility_of_element_located(year_dropdown_opener_locator))
        print(f"Year dropdown opener is visible within iframe.")

        shadow_doc2.execute_script("arguments[0].click();", year_dropdown_opener)
        print("Clicked 'tabComboBoxNameContainer' div to open Year dropdown via JS.")
        time.sleep(3)

    except TimeoutException as e:
        screenshot_path = os.path.join(downloadPath, f"debug_year_dropdown_iframe_timeout_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path)
        print(f"TimeoutException: Failed to open year dropdown within iframe. Screenshot saved to {screenshot_path}")
        print(f"Error details: {e}")
        driver.quit()
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred while trying to open year dropdown within iframe: {e}")
        driver.quit()
        sys.exit(1)


    try:
        y2024_input = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@class, "facetOverflow")]//a[text()="2024"]/preceding-sibling::input')
        ))
        shadow_doc2.execute_script("arguments[0].click();", y2024_input)
        print("Selected year 2024.")
        time.sleep(2)

        y2023_input = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@class, "facetOverflow")]//a[text()="2023"]/preceding-sibling::input')
        ))
        shadow_doc2.execute_script("arguments[0].click();", y2023_input)
        print("Selected year 2023.")
        time.sleep(2)
    except Exception as e:
        print(f"Error selecting years (within iframe): {e}. Ensure the year filter pop-up is visible and elements are correct.")
        driver.quit()
        sys.exit(1)

    try:
        dd_close = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))
        )
        shadow_doc2.execute_script("arguments[0].click();", dd_close)
        print("Closed year dropdown menu.")
    except Exception:
        print("Could not find or click 'tab-glass' to close dropdown (within iframe). Attempting to click body.")
        shadow_doc2.find_element(By.TAG_NAME, 'body').click()
        print("Clicked body to close dropdown.")
    time.sleep(3)


    print(f"Processing Initial Week Number: 53")
    download_and_rename(wait, shadow_doc2, 53, default_dir, downloadPath, driver, year, today)

    weeknum = 53
    while weeknum > 0:
        print(f"Processing Week Number: {weeknum-1}")

        # Decrement button - LOCATOR WITHIN IFRAME CONTEXT, ensure it's in view
        decrement_button_locator = (By.XPATH, "//div[@id='tabZoneId7']//div[contains(@class, 'tableauArrowDec')]")
        decrement_button = wait.until(EC.element_to_be_clickable(decrement_button_locator))
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", decrement_button)
        time.sleep(0.5)

        shadow_doc2.execute_script("arguments[0].click();", decrement_button)
        time.sleep(3)

        weeknum -= 1

        download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today)

        if weeknum == 1:
            print("Reached week 1, breaking the loop.")
            break

    driver.switch_to.default_content()

    driver.quit()
    print("Script finished successfully.")

if __name__ == "__main__":
    try:
        iterate_weekly()
    except Exception as e:
        print(f"An error occurred during script execution: {e}")
        sys.exit(1)
