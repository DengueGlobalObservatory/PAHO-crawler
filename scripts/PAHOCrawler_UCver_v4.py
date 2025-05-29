import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from datetime import datetime
import subprocess
import re
import sys

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

# --- Improved move_to_download_folder function ---
def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    start_time = time.time()
    timeout = 180 # 3 minutes timeout for download
    while not got_file and (time.time() - start_time) < timeout:
        try:
            files = [os.path.join(default_dir, f) for f in os.listdir(default_dir) if os.path.isfile(os.path.join(default_dir, f))]

            # Filter out temporary download files (e.g., .crdownload, .tmp, .part, .download)
            downloading_files = [f for f in files if f.endswith('.crdownload') or f.endswith('.tmp') or f.endswith('.part') or f.endswith('.download')]

            if downloading_files:
                print(f"Waiting for download to complete: {downloading_files[0]}")
                time.sleep(5)
                continue

            finished_files = [f for f in files if not f.endswith('.crdownload') and not f.endswith('.tmp') and not f.endswith('.part') and not f.endswith('.download')]
            if not finished_files:
                # No completed files found yet, but check if there are any files at all.
                if not files and (time.time() - start_time) > 10: # Only raise if no files for a bit
                     raise FileNotFoundError("No files found in the download directory yet.")
                else: # Wait more if files are still appearing/downloading
                    time.sleep(5)
                    continue

            currentFile = max(finished_files, key=os.path.getctime)

            if os.path.exists(currentFile) and not (currentFile.endswith('.crdownload') or currentFile.endswith('.tmp') or currentFile.endswith('.part') or currentFile.endswith('.download')):
                got_file = True
            else:
                # File found but still looks like a temp file, or not fully ready
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

    # Wait for the week number to update (using shadow_doc2, which is the driver in iframe)
    weeknum_div = wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
    )
    weeknum = int(weeknum_div.text)  # Convert to integer for comparison
    print(f"Confirmed Week Number Displayed: {weeknum}")

    # Step 1: Click the "Tabla Anual" tab (Annual Table)
    # This tab is within the iframe
    try:
        table_tab = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='tabLabel' and text()='Tabla Anual']"))
        )
        table_tab.click()
        print("Clicked 'Tabla Anual' tab.")
        time.sleep(3) # Give it a moment to load the table view
    except Exception as e:
        print(f"Could not click 'Tabla Anual' tab: {e}. Ensure this tab is present and visible.")
        raise # Critical step, re-raise error if tab not found

    # Find and click the download button at the bottom of the dashboard
    # Using the ID you provided for robustness, and clicking via JS
    download_button_id = "download-ToolbarButton"
    download_button = wait.until(
        EC.presence_of_element_located((By.ID, download_button_id))
    )

    try:
        print(f"Attempting to click download button ({download_button_id}) via JS...")
        # Scroll to ensure it's in view, then click via JS
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_button)
        time.sleep(0.5) # Short pause after scrolling
        shadow_doc2.execute_script("arguments[0].click();", download_button)
        print("Clicked download button successfully via JS.")
        time.sleep(5)
    except Exception as e:
        print(f"Failed to click download button ({download_button_id}) via JS: {e}")
        # Take a screenshot if this critical step fails
        screenshot_path = os.path.join(downloadPath, f"debug_download_button_fail_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path) # Screenshot within the iframe context
        print(f"Screenshot taken to help debug: {screenshot_path}")
        raise # Re-raise the exception to stop execution


    # Find and click the crosstab button (in a pop up window) - LOCATOR REMAINS THE SAME
    crosstab_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
    )
    crosstab_button.click()
    print("Clicked Crosstab button.")
    time.sleep(5)

    # Find and select the CSV option - LOCATOR REMAINS THE SAME
    csv_div = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='radio'][value='csv']"))
    )
    shadow_doc2.execute_script("arguments[0].scrollIntoView();", csv_div)
    shadow_doc2.execute_script("arguments[0].click();", csv_div)
    print("Selected CSV option.")
    time.sleep(5)

    # Find and click the export button - LOCATOR REMAINS THE SAME
    export_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[data-tb-test-id="export-crosstab-export-Button"]')))

    # Click via JavaScript for robustness (similar to download button)
    try:
        print("Attempting to scroll export button into view and click via JS...")
        shadow_doc2.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_button)
        time.sleep(0.5) # Small pause after scrolling
        shadow_doc2.execute_script("arguments[0].click();", export_button)
        print("Downloading CSV file")
        time.sleep(5)
    except Exception as e:
        print(f"Failed to click export button even with scrolling and JS click: {e}")
        screenshot_path = os.path.join(downloadPath, f"debug_export_button_click_fail_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path) # Screenshot within iframe context
        print(f"Screenshot taken to help debug: {screenshot_path}")
        raise # Re-raise the exception to stop execution


    # Use the move_to_download_folder function to move the downloaded file
    downloadPath = downloadPath
    default_dir = default_dir
    newFileName = f"PAHO_2023_2025_W{weeknum}_{today}"
    fileExtension = '.csv'

    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def iterate_weekly():

    year = "(All)"
    today = datetime.now().strftime('%Y%m%d%H%M')

    # set directory
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

    # set chrome download directory
    chrome_options = uc.ChromeOptions()
    prefs = {
        "download.default_directory": default_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # Add more robust arguments for headless execution on CI/CD
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
    chrome_options.add_argument("--ignore-certificate-errors") # May help with SSL errors on some sites
    chrome_options.add_argument("--no-zygote") # Sometimes improves stability in Linux containers


    # using undetected-chromedriver
    driver = uc.Chrome(headless=True, use_subprocess=False, options=chrome_options, version_main=chrome_version)
    driver.get('https://www.paho.org/en/arbo-portal/dengue-data-and-analysis/dengue-analysis-country')

    wait = WebDriverWait(driver, 60) # Increased timeout for robustness

    iframe_page_title = driver.title
    print(f"Current page title: {iframe_page_title}")

    # FIX: Corrected page title check to match actual title from output
    if "Dengue: analysis by country" not in iframe_page_title:
        print("Wrong access: Page title does not match expected. Exiting.")
        driver.quit()
        sys.exit(1)
    else:
        print("Page title confirmed.")

    time.sleep(7) # Increased sleep after main page load to allow JavaScript to process

    # --- Simplified Iframe Switching ---
    iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008_NAC_ES_Indicadores_reporte_semanal.asp"
    iframe_locator = (By.XPATH, f"//iframe[@src='{iframe_src}']")

    try:
        print(f"Attempting to switch to iframe with src: {iframe_src}")
        iframe = wait.until(EC.presence_of_element_located(iframe_locator))
        driver.switch_to.frame(iframe)
        print("Switched to iframe successfully.")

        # Assign driver to shadow_doc2 to maintain original script's variable name for iframe context
        shadow_doc2 = driver

    except TimeoutException as e:
        print(f"Timeout: Iframe with src '{iframe_src}' not found within {wait.timeout} seconds. Error: {e}")
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
        wait.until(EC.presence_of_element_located((By.ID, "dashboard-viewport")))
        print("Dashboard viewport found inside iframe, presuming dashboard loaded correctly.")
    except TimeoutException as e:
        print(f"Timeout: Dashboard viewport not found inside iframe within {wait.timeout} seconds. Error: {e}")
        screenshot_path = os.path.join(downloadPath, f"debug_dashboard_viewport_timeout_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        shadow_doc2.save_screenshot(screenshot_path) # Use shadow_doc2 for screenshot within iframe
        print(f"Screenshot taken: {screenshot_path}")
        driver.quit()
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred after switching to iframe: {e}")
        driver.quit()
        sys.exit(1)

    time.sleep(5) # Added extra sleep after dashboard is confirmed present

    # --- Find and click the year dropdown opener ---
    # LOCATOR WITHIN IFRAME CONTEXT, using JS click for robustness
    year_filter_container_id = 'tabZoneId16'

    try:
        wait.until(EC.presence_of_element_located((By.ID, year_filter_container_id)))
        print(f"Year filter container ({year_filter_container_id}) is present in iframe DOM.")

        year_dropdown_opener_locator = (By.XPATH, f"//div[@id='{year_filter_container_id}']//div[contains(@class, 'tabComboBoxNameContainer')]")
        year_dropdown_opener = wait.until(EC.visibility_of_element_located(year_dropdown_opener_locator))
        print(f"Year dropdown opener is visible within iframe.")

        # Click via JavaScript
        shadow_doc2.execute_script("arguments[0].click();", year_dropdown_opener)
        print("Clicked 'tabComboBoxNameContainer' div to open Year dropdown via JS.")
        time.sleep(3) # Increased sleep after clicking to allow dropdown to fully render

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


    # select year 2024 - LOCATOR WITHIN IFRAME CONTEXT
    try:
        y2024_input = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@class, "facetOverflow")]//a[text()="2024"]/preceding-sibling::input')
        ))
        shadow_doc2.execute_script("arguments[0].click();", y2024_input) # Use shadow_doc2 for click
        print("Selected year 2024.")
        time.sleep(2) # Increased sleep

        # select year 2023
        y2023_input = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//div[contains(@class, "facetOverflow")]//a[text()="2023"]/preceding-sibling::input')
        ))
        shadow_doc2.execute_script("arguments[0].click();", y2023_input) # Use shadow_doc2 for click
        print("Selected year 2023.")
        time.sleep(2) # Increased sleep
    except Exception as e:
        print(f"Error selecting years (within iframe): {e}. Ensure the year filter pop-up is visible and elements are correct.")
        driver.quit()
        sys.exit(1)

    # close the dropdown menu - LOCATOR WITHIN IFRAME CONTEXT
    try:
        dd_close = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))
        )
        shadow_doc2.execute_script("arguments[0].click();", dd_close) # Use shadow_doc2 for click
        print("Closed year dropdown menu.")
    except Exception:
        print("Could not find or click 'tab-glass' to close dropdown (within iframe). Attempting to click body.")
        shadow_doc2.find_element(By.TAG_NAME, 'body').click() # Fallback to clicking body
        print("Clicked body to close dropdown.")
    time.sleep(3) # Keep sleep after closing dropdown


    # Initial call to download_and_rename (for week 53 only)
    # Pass `shadow_doc2` (which is now the driver object in the iframe)
    print(f"Processing Initial Week Number: 53")
    # Note: `driver_main_page_ignored` parameter is just a placeholder to match original signature.
    # We pass `driver` again, but it's not used inside download_and_rename.
    download_and_rename(wait, shadow_doc2, 53, default_dir, downloadPath, driver, year, today)

    weeknum = 53  # Initialize weeknum outside the loop
    while weeknum > 0:
        print(f"Processing Week Number: {weeknum-1}")

        # Decrement button - LOCATOR WITHIN IFRAME CONTEXT
        decrement_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//div[@id='tabZoneId7']//div[contains(@class, 'tableauArrowDec')]")
        ))
        shadow_doc2.execute_script("arguments[0].click();", decrement_button) # Use shadow_doc2 for click
        time.sleep(3)

        # Update weeknum after decrementing
        weeknum -= 1

        # Pass updated weeknum to download_and_rename
        download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today)

        if weeknum == 1:
            print("Reached week 1, breaking the loop.")
            break

    # VERY IMPORTANT: Switch back to default content if you need to interact with main page elements later
    # This is good practice before quitting, or if you had more steps on the main page.
    driver.switch_to.default_content()

    driver.quit()
    print("Script finished successfully.")

# Execute the main function
if __name__ == "__main__":
    try:
        iterate_weekly()
    except Exception as e:
        print(f"An error occurred during script execution: {e}")
        sys.exit(1)
