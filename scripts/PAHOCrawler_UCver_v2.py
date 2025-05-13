import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException # Import TimeoutException
import time
import os
from datetime import datetime, date # Added date for isocalendar
import subprocess
import re
import sys
import glob # Added for move_to_download_folder
import shutil # Added for move_to_download_folder fallback

# Function to get the installed Chrome version
def get_chrome_version():
    try:
        if sys.platform == "win32":
            # Command to retrieve Chrome version from Windows registry
            output = subprocess.check_output(
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                shell=True,
                text=True,
                stderr=subprocess.DEVNULL # Suppress stderr for this command
            )
            version_match = re.search(r'\s+version\s+REG_SZ\s+(\d+)\.', output)
        elif sys.platform == "darwin": # macOS
            # Command for macOS
            chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            if os.path.exists(chrome_path):
                output = subprocess.check_output([chrome_path, '--version'], text=True)
                version_match = re.search(r'Google Chrome (\d+)\.', output)
            else: # Try Chromium if Chrome is not found
                chromium_path = "/Applications/Chromium.app/Contents/MacOS/Chromium"
                if os.path.exists(chromium_path):
                    output = subprocess.check_output([chromium_path, '--version'], text=True)
                    version_match = re.search(r'Chromium (\d+)\.', output)
                else:
                    # Try finding Chrome via mdfind
                    try:
                        chrome_app_path_output = subprocess.check_output(
                            ['mdfind', 'kMDItemCFBundleIdentifier == "com.google.Chrome"'],
                            text=True
                        ).strip()
                        if chrome_app_path_output: # Check if mdfind returned a path
                            # mdfind can return multiple paths, take the first one
                            first_path = chrome_app_path_output.split('\n')[0]
                            chrome_executable_path = os.path.join(first_path, "Contents/MacOS/Google Chrome")
                            if os.path.exists(chrome_executable_path):
                                output = subprocess.check_output([chrome_executable_path, '--version'], text=True)
                                version_match = re.search(r'Google Chrome (\d+)\.', output)
                            else:
                                raise FileNotFoundError(f"Chrome found by mdfind at {first_path}, but executable missing at {chrome_executable_path}.")
                        else:
                             raise FileNotFoundError("Chrome or Chromium not found on macOS using mdfind.")
                    except (subprocess.CalledProcessError, FileNotFoundError) as e_mdfind:
                        print(f"mdfind for Chrome failed: {e_mdfind}")
                        raise FileNotFoundError("Chrome or Chromium not found on macOS.")

        else:  # Assuming Linux or other Unix-like systems
            version_match = None
            # Try different commands to retrieve Chrome version on Linux
            for command in ['google-chrome --version', 'google-chrome-stable --version', 'chromium-browser --version', 'chromium --version']:
                try:
                    output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'\b(\d+)\.', output)
                    if version_match:
                        break
                except subprocess.CalledProcessError:
                    continue
            if not version_match:
                raise RuntimeError("Could not determine Chrome version on Linux/Unix.")

        if version_match:
            return int(version_match.group(1))
        else:
            raise ValueError("Could not parse Chrome version from output.")
    except Exception as e:
        print(f"Detailed error getting Chrome version: {e}")
        env_chrome_version = os.getenv('CHROME_VERSION_MAJOR')
        if env_chrome_version and env_chrome_version.isdigit():
            print(f"Using CHROME_VERSION_MAJOR from environment: {env_chrome_version}")
            return int(env_chrome_version)
        print("Failed to automatically detect Chrome version, returning a common default (e.g., 124). Set CHROME_VERSION_MAJOR env var to override.")
        return 124

try:
    chrome_version = get_chrome_version()
    print(f"Detected Chrome major version: {chrome_version}")
except RuntimeError as e:
    print(f"Error in get_chrome_version: {e}")
    print("Attempting to proceed with a default Chrome version (e.g. 124 for uc.Chrome).")
    chrome_version = 124 # Fallback version


def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    max_retries = 18 # Retry for up to 3 minutes (18 * 10 seconds)
    retries = 0
    currentFile = None

    print(f"Monitoring download directory: {default_dir} for new file (filename: {newFileName}{fileExtension})...")

    time.sleep(15) # Initial wait for download to start and file to appear

    while not got_file and retries < max_retries:
        try:
            all_current_files_in_dir = [os.path.join(default_dir, f) for f in os.listdir(default_dir)
                                        if os.path.isfile(os.path.join(default_dir, f))]

            if not all_current_files_in_dir:
                if retries % 3 == 0:
                    print(f"No files found in download directory: {default_dir}. Waiting... (Attempt {retries+1}/{max_retries})")
                time.sleep(10)
                retries += 1
                continue

            potential_files = [f for f in all_current_files_in_dir if not f.endswith('.crdownload') and not f.endswith('.tmp')]

            if not potential_files:
                if retries % 3 == 0:
                    crdownload_files = [f for f in all_current_files_in_dir if f.endswith('.crdownload')]
                    if crdownload_files:
                        print(f"Download in progress: {os.path.basename(crdownload_files[0])}. Waiting... (Attempt {retries+1}/{max_retries})")
                    else:
                        print(f"No completed files found yet. Found: {[os.path.basename(f) for f in all_current_files_in_dir]}. Waiting... (Attempt {retries+1}/{max_retries})")
                time.sleep(10)
                retries += 1
                continue

            currentFile = max(potential_files, key=os.path.getctime)
            time.sleep(2) # Wait for filesystem to catch up

            if os.path.exists(currentFile) and os.path.getsize(currentFile) > 0:
                print(f"Downloaded file identified: {os.path.basename(currentFile)} (Size: {os.path.getsize(currentFile)} bytes)")
                got_file = True
            else:
                if not os.path.exists(currentFile):
                    print(f"Identified file {os.path.basename(currentFile)} does not exist. Retrying...")
                elif os.path.getsize(currentFile) == 0:
                     print(f"Identified file {os.path.basename(currentFile)} is empty. Assuming download is not complete. Retrying...")
                currentFile = None
                time.sleep(10)
                retries += 1

        except Exception as e:
            print(f"Error checking download directory: {e}. Retrying... (Attempt {retries+1}/{max_retries})")
            time.sleep(10)
            retries += 1

    if not got_file or currentFile is None:
        print(f"Failed to get the downloaded file '{newFileName}{fileExtension}' from {default_dir} after {(max_retries * 10)} seconds.")
        placeholder_filename = newFileName + "_DOWNLOAD_FAILED" + fileExtension
        placeholder_filepath = os.path.join(downloadPath, placeholder_filename)
        try:
            with open(placeholder_filepath, 'w') as f:
                f.write("Download failed or file was empty/not found in the source directory.")
            print(f"Created a placeholder error file: {placeholder_filepath}")
        except Exception as e_placeholder:
            print(f"Failed to create placeholder error file: {e_placeholder}")
        return

    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)
    try:
        print(f"Attempting to move {os.path.basename(currentFile)} to {fileDestination}")
        shutil.move(currentFile, fileDestination)
        print(f"Moved file to {fileDestination}")
    except Exception as e:
        print(f"Error moving file {currentFile} to {fileDestination} using shutil.move: {e}")
        placeholder_filename = newFileName + "_MOVE_FAILED" + fileExtension
        placeholder_filepath = os.path.join(downloadPath, placeholder_filename)
        try:
            with open(placeholder_filepath, 'w') as f:
                f.write(f"Download succeeded but move failed. Original file: {currentFile}, Error: {e}")
            print(f"Created a placeholder error file for move failure: {placeholder_filepath}")
        except Exception as e_placeholder_move:
            print(f"Failed to create placeholder error file for move failure: {e_placeholder_move}")


def download_and_rename(wait, driver, weeknum_for_filename, default_dir, downloadPath, year_str, today_timestamp):
    """Download and rename the file for the given week number."""
    print(f"Attempting to download data for year {year_str}, expected week (for filename): {weeknum_for_filename}")

    actual_week_on_page = weeknum_for_filename
    try:
        time.sleep(3)
        weeknum_div = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
        )
        actual_week_on_page = int(weeknum_div.text)
        print(f"Actual week number on page: {actual_week_on_page}. Using this for filename.")
    except Exception as e:
        print(f"Could not read week number from page: {e}. Using expected week for filename: {weeknum_for_filename}")

    try:
        download_button = wait.until(EC.element_to_be_clickable((By.ID, "download-ToolbarButton")))
        driver.execute_script("arguments[0].click();", download_button)
        print("Clicked main download button.")
        time.sleep(10)

        print("Waiting for download options dialog to become visible...")
        download_options_dialog_xpath = "//div[(@role='dialog' or contains(@class, 'tableauModalDialog')) and .//button[@data-tb-test-id='DownloadCrosstab-Button']]"
        try:
            wait.until(EC.visibility_of_element_located((By.XPATH, download_options_dialog_xpath)))
            print("Download options dialog is visible.")
        except TimeoutException:
            print("Timeout: Download options dialog (or one containing crosstab button) did not become visible.")
            return

        crosstab_button_selector = '[data-tb-test-id="DownloadCrosstab-Button"]'
        print(f"Waiting for crosstab button: {crosstab_button_selector}")
        crosstab_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, crosstab_button_selector)))
        driver.execute_script("arguments[0].click();", crosstab_button)
        print("Clicked crosstab button.")
        time.sleep(5)

        csv_radio_selector = "input[type='radio'][value='csv']"
        print(f"Waiting for CSV radio button: {csv_radio_selector}")
        csv_div = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, csv_radio_selector)))
        driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", csv_div)
        print("Selected CSV option.")
        time.sleep(3)

        export_button_selector = '[data-tb-test-id="export-crosstab-export-Button"]'
        print(f"Waiting for export button: {export_button_selector}")
        export_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector)))
        driver.execute_script("arguments[0].click();", export_button)
        print(f"Downloading CSV file for year {year_str}, week {actual_week_on_page}...")

    except TimeoutException as te:
        print(f"A timeout occurred during the download process for week {actual_week_on_page}, year {year_str}: {te}")
        try:
            close_buttons = driver.find_elements(By.XPATH, "//button[contains(@aria-label, 'Close') or contains(@class, 'close') or .//span[contains(text(),'×')]]")
            for btn in close_buttons:
                if btn.is_displayed() and btn.is_enabled():
                    print("Attempting to click a close button to dismiss dialog...")
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(1)
                    break
        except Exception as e_close:
            print(f"Could not find or click a close button for the dialog: {e_close}")
        return

    except Exception as e:
        print(f"An unexpected error occurred during download preparation for week {actual_week_on_page}, year {year_str}: {e}")
        return

    newFileName = f"PAHO_{year_str}_W{actual_week_on_page}_{today_timestamp}"
    fileExtension = '.csv'
    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def click_filter_checkbox_robust(wait, driver, text_label, filter_type="year"):
    """More robustly clicks a filter checkbox in Tableau, trying input then label."""
    xpath_input = f'//div[contains(@class, "facetOverflow")]//a[text()="{text_label}"]/preceding-sibling::input[@type="checkbox"]'
    xpath_label = f'//div[contains(@class, "facetOverflow")]//a[text()="{text_label}"]'

    try:
        print(f"Attempting to click {filter_type} filter INPUT for: {text_label}")
        checkbox_input = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_input)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", checkbox_input)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", checkbox_input)
        print(f"Clicked {filter_type} filter INPUT for: {text_label} using JS.")
        time.sleep(1.5)
        return True
    except Exception as e_input:
        print(f"Could not click {filter_type} filter INPUT for '{text_label}': {e_input}. Trying LABEL.")
        try:
            print(f"Attempting to click {filter_type} filter LABEL for: {text_label}")
            checkbox_label = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_label)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", checkbox_label)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", checkbox_label)
            print(f"Clicked {filter_type} filter LABEL for: {text_label} using JS.")
            time.sleep(1.5)
            return True
        except Exception as e_label:
            print(f"Could not click {filter_type} filter LABEL for '{text_label}': {e_label}")
            return False


def iterate_weekly():
    years_to_process = ["2023", "2024", "2025"]
    today_timestamp = datetime.now().strftime('%Y%m%d%H%M%S')

    base_data_dir = os.getenv('GITHUB_WORKSPACE', '.')
    data_dir_path = os.path.join(base_data_dir, 'data')
    default_download_dir_for_browser = os.path.join(os.getcwd(), "chrome_downloads")
    os.makedirs(default_download_dir_for_browser, exist_ok=True)
    print(f"Chrome default download directory set to: {default_download_dir_for_browser}")

    final_download_path_base = os.path.join(data_dir_path, f"DL_{datetime.now().strftime('%Y%m%d')}")
    os.makedirs(final_download_path_base, exist_ok=True)
    print(f"Base storage path for downloaded files: {final_download_path_base}")

    chrome_options = uc.ChromeOptions()
    prefs = {
        "download.default_directory": default_download_dir_for_browser,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    # REMOVED: chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    # REMOVED: chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    # REMOVED: chrome_options.add_experimental_option('useAutomationExtension', False)


    print(f"Initializing Chrome with version_main={chrome_version}")
    driver = uc.Chrome(headless=True, use_subprocess=True, options=chrome_options, version_main=chrome_version)
    # REMOVED: driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')
    print("Page loaded.")
    wait = WebDriverWait(driver, 45)

    iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
    iframe_locator = (By.XPATH, f"//div[contains(@class, 'vizTab')]//iframe[@src='{iframe_src}']")
    print("Waiting for the first iframe...")
    iframe1 = wait.until(EC.presence_of_element_located(iframe_locator))
    driver.switch_to.frame(iframe1)
    print("Switched to the first iframe.")
    time.sleep(2)

    print("Waiting for the nested (Tableau) iframe...")
    tableau_iframe_selectors = [
        (By.XPATH, "//iframe[@title='viz']"),
        (By.ID, "viz_iframe"),
        (By.XPATH, "//iframe[contains(@id,'viz_iframe')]"),
        (By.TAG_NAME, "iframe")
    ]
    iframe2 = None
    # Store the current window handle in case we need to switch back to default content from an iframe for debugging
    # main_window_handle = driver.current_window_handle # Not strictly needed if only dealing with frames

    for i, (by, selector_val) in enumerate(tableau_iframe_selectors):
        try:
            print(f"Attempting to find Tableau iframe with: {by}='{selector_val}'")
            iframe2_element = wait.until(EC.presence_of_element_located((by, selector_val)))
            driver.switch_to.frame(iframe2_element)
            print("Found and switched to nested (Tableau) iframe.")
            iframe2 = iframe2_element
            break
        except Exception as e_iframe:
            print(f"Attempt {i+1} to find Tableau iframe with {by}='{selector_val}' failed: {e_iframe}")
            # Ensure we are in the context of the first iframe before trying the next selector for the second iframe
            # This assumes the structure is Page -> iframe1 -> iframe2
            # If a switch to iframe2_element fails, we might still be in iframe1 or default_content.
            # To be safe, explicitly switch back to iframe1 before next attempt to find iframe2 within it.
            # However, if the wait.until fails, the context doesn't change.
            # If driver.switch_to.frame(iframe2_element) fails, then an error occurs.
            # The crucial part is that if a find for iframe2 fails, we are still in iframe1's context.
            if i < len(tableau_iframe_selectors) - 1:
                print("Retrying to find Tableau iframe from the context of the first iframe.")
                # No need to switch back explicitly if wait.until for iframe2_element failed,
                # as the context would not have changed to iframe2_element.
            else: # All attempts failed
                print("CRITICAL: Could not find or switch to the Tableau iframe. Exiting.")
                driver.quit()
                return
    if iframe2 is None:
        print("CRITICAL: Tableau iframe not found after all attempts. Exiting.")
        driver.quit()
        return

    time.sleep(5)

    region_filter_tab_id = 'tabZoneId9'
    print(f"Waiting for region filter control (ID: {region_filter_tab_id})...")
    wait.until(EC.visibility_of_element_located((By.ID, region_filter_tab_id)))
    region_dropdown_button_selector = (By.CSS_SELECTOR, f'#{region_filter_tab_id} span.tabComboBoxButton')
    wait.until(EC.element_to_be_clickable(region_dropdown_button_selector)).click()
    print("Clicked to open region selection dropdown.")
    time.sleep(2)
    if click_filter_checkbox_robust(wait, driver, "(All)", filter_type="region"):
        print("Successfully selected (All) regions.")
    else:
        print("WARNING: Failed to select (All) regions. Proceeding, but data might be incomplete.")

    try:
        open_dropdown_button = driver.find_element(By.CSS_SELECTOR, f'#{region_filter_tab_id} span.tabComboBoxButton.tabDropDownButtonOpen')
        if open_dropdown_button.is_displayed():
             open_dropdown_button.click()
             print("Closed region dropdown by clicking the open button again.")
        else:
            raise Exception("Dropdown button not in 'open' state.")
    except:
        try:
            wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))).click()
            print("Closed region selection dropdown via tab-glass.")
        except Exception as e_close_region:
            print(f"Could not close region dropdown cleanly: {e_close_region}. Attempting to continue by blurring active element.")
            try:
                active_element = driver.switch_to.active_element
                if active_element:
                    driver.execute_script("arguments[0].blur();", active_element)
            except Exception as e_blur:
                print(f"Failed to blur active element: {e_blur}")
    time.sleep(5)


    for year_to_download in years_to_process:
        print(f"\nProcessing Year: {year_to_download}")

        current_year_download_path = os.path.join(final_download_path_base, year_to_download)
        os.makedirs(current_year_download_path, exist_ok=True)
        print(f"Download path for {year_to_download}: {current_year_download_path}")

        year_filter_tab_id = 'tabZoneId13'
        print(f"Waiting for year filter control (ID: {year_filter_tab_id})...")
        wait.until(EC.visibility_of_element_located((By.ID, year_filter_tab_id)))
        year_dropdown_button_selector = (By.CSS_SELECTOR, f'#{year_filter_tab_id} span.tabComboBoxButton')
        wait.until(EC.element_to_be_clickable(year_dropdown_button_selector)).click()
        print("Clicked to open year selection dropdown.")
        time.sleep(2)

        print(f"Clearing previous year selections and selecting only {year_to_download}...")
        click_filter_checkbox_robust(wait, driver, "(All)")
        click_filter_checkbox_robust(wait, driver, "(All)")

        if not click_filter_checkbox_robust(wait, driver, year_to_download):
            print(f"CRITICAL: Failed to select target year {year_to_download}. Skipping this year.")
            try:
                open_year_dropdown_button = driver.find_element(By.CSS_SELECTOR, f'#{year_filter_tab_id} span.tabComboBoxButton.tabDropDownButtonOpen')
                if open_year_dropdown_button.is_displayed(): open_year_dropdown_button.click()
            except:
                try: wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))).click()
                except: pass
            time.sleep(1)
            continue

        try:
            open_year_dropdown_button = driver.find_element(By.CSS_SELECTOR, f'#{year_filter_tab_id} span.tabComboBoxButton.tabDropDownButtonOpen')
            if open_year_dropdown_button.is_displayed():
                open_year_dropdown_button.click()
                print("Closed year dropdown by clicking the open button again.")
            else: raise Exception("Year dropdown button not in 'open' state.")
        except:
            try:
                wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))).click()
                print("Closed year selection dropdown via tab-glass.")
            except Exception as e_close_year:
                print(f"Could not close year dropdown cleanly: {e_close_year}. Attempting to blur.")
                try:
                    active_element = driver.switch_to.active_element
                    if active_element: driver.execute_script("arguments[0].blur();", active_element)
                except: pass
        time.sleep(5)

        initial_weeknum = 0
        try:
            weeknum_div_initial = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "sliderText")))
            initial_weeknum = int(weeknum_div_initial.text)
            print(f"Initial week on page for year {year_to_download} (slider): {initial_weeknum}")
        except Exception as e:
            print(f"Could not read initial week number for {year_to_download}: {e}. Calculating fallback.")
            year_int = int(year_to_download)
            last_day = date(year_int, 12, 31)
            iso_year, iso_week, _ = last_day.isocalendar()
            if iso_year == year_int:
                 initial_weeknum = iso_week
            else:
                 initial_weeknum = date(year_int, 12, 24).isocalendar()[1]

            print(f"Calculated fallback initial week for {year_to_download}: {initial_weeknum}")


        current_loop_week = initial_weeknum
        if current_loop_week == 0:
            print(f"ERROR: Initial week number is 0 for year {year_to_download}. Cannot proceed for this year.")
            continue

        download_and_rename(wait, driver, current_loop_week, default_download_dir_for_browser, current_year_download_path, year_to_download, today_timestamp)

        decrement_button_xpath = "//*[contains(@class, 'tableauArrow') and contains(@class, 'Dec')]"
        while current_loop_week > 1:
            print(f"Preparing to decrement week from {current_loop_week} for year {year_to_download}...")
            try:
                decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, decrement_button_xpath)))
                driver.execute_script("arguments[0].click();", decrement_button)
                print(f"Clicked decrement week button. Current week was {current_loop_week}.")
                current_loop_week -= 1
                time.sleep(7)
            except Exception as e_dec:
                print(f"Error clicking decrement button for year {year_to_download} (week {current_loop_week}) or week already at 1: {e_dec}")
                break

            download_and_rename(wait, driver, current_loop_week, default_download_dir_for_browser, current_year_download_path, year_to_download, today_timestamp)

            if current_loop_week == 1:
                print(f"Reached week 1 for year {year_to_download}, finishing download cycle for this year.")
                break

        print(f"--- Finished processing year: {year_to_download} ---")

    print("\nAll specified years processed.")
    driver.quit()

if __name__ == "__main__":
    try:
        iterate_weekly()
    except Exception as e:
        print(f"An unhandled error occurred in the main script: {e}")
        import traceback
        traceback.print_exc()

