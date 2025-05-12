import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import time
import os
from datetime import datetime
import subprocess
import re
import sys
import shutil # For robust file moving

# Function to get the installed Chrome version
def get_chrome_version():
    """
    Attempts to get the installed Chrome major version.
    More robust for different OS and common installation paths.
    Returns None if version cannot be determined, allowing uc.Chrome to auto-detect.
    """
    try:
        version_match = None
        if sys.platform == "win32":
            reg_paths = [
                r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Google\Chrome\Update" /v LastKnownVersionString /reg:32',
                r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Google\Chrome\Update" /v LastKnownVersionString /reg:64',
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version'
            ]
            for path in reg_paths:
                try:
                    output = subprocess.check_output(path, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'REG_SZ\s+(\d+)\.', output)
                    if version_match: break
                except (subprocess.CalledProcessError, FileNotFoundError): continue
        elif sys.platform == "darwin": # macOS
             commands = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome --version']
             for command in commands:
                try:
                    output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'Google Chrome (\d+)\.', output)
                    if version_match: break
                except (subprocess.CalledProcessError, FileNotFoundError): continue
        else:  # Linux
            commands = [
                'google-chrome-stable --version',
                'google-chrome --version',
                'chromium-browser --version',
                'chromium --version'
            ]
            for command in commands:
                try:
                    output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'\b(\d+)\.', output) # More general regex for version numbers
                    if version_match: break
                except (subprocess.CalledProcessError, FileNotFoundError): continue

        if version_match:
            version = int(version_match.group(1))
            print(f"Detected Chrome major version: {version}")
            return version
        else:
            print("Could not automatically determine Chrome version. uc.Chrome will attempt auto-detection.")
            return None # Let uc.Chrome try to find it

    except Exception as e:
        print(f"Warning: Failed to get Chrome version due to: {e}. uc.Chrome will attempt auto-detection.")
        return None

# Get the major version of Chrome installed (can be None)
chrome_version_detected = get_chrome_version()

def move_to_download_folder(default_dir_path, final_download_path, new_file_name_stem, file_extension):
    """
    Waits for a download to complete in default_dir_path, then moves the latest file
    to final_download_path with the new name.
    """
    got_file = False
    max_wait_time = 180  # Increased wait time for download (3 minutes)
    check_interval = 5
    start_time = time.time()
    downloaded_file_path = None

    print(f"Waiting for download in '{default_dir_path}' for up to {max_wait_time}s...")

    while not got_file and (time.time() - start_time) < max_wait_time:
        try:
            # List files, ignore temporary download files
            candidate_files = [
                os.path.join(default_dir_path, f) for f in os.listdir(default_dir_path)
                if not f.lower().endswith(('.crdownload', '.tmp', '.part')) and \
                   os.path.isfile(os.path.join(default_dir_path, f))
            ]

            if not candidate_files:
                time.sleep(check_interval)
                continue

            # Get the latest (by modification time) candidate file
            latest_file = max(candidate_files, key=os.path.getmtime)

            # Check if the file is still being written (basic size check)
            # Allow a short time for file system to catch up
            time.sleep(1) # Brief pause before size check
            initial_size = os.path.getsize(latest_file)
            time.sleep(2) # Wait a moment to see if size changes
            current_size = os.path.getsize(latest_file)

            if initial_size == current_size and current_size > 0:
                print(f"Detected downloaded file: {os.path.basename(latest_file)} (Size: {current_size} bytes)")
                downloaded_file_path = latest_file
                got_file = True
            else:
                print(f"File '{os.path.basename(latest_file)}' size changed (Initial: {initial_size}, Current: {current_size}) or is empty. Re-evaluating...")
                time.sleep(check_interval - 3) # Adjust remaining sleep time

        except FileNotFoundError:
            print(f"Download directory '{default_dir_path}' not found during check. Retrying...")
            time.sleep(check_interval)
        except Exception as e:
            print(f"Error checking download directory: {e}. Retrying...")
            time.sleep(check_interval)

    if not got_file or not downloaded_file_path:
        print(f"Error: Download did not complete or file not detected within {max_wait_time} seconds in '{default_dir_path}'.")
        raise TimeoutError(f"Download timeout or file not found in '{default_dir_path}'.")

    final_file_name = new_file_name_stem + file_extension
    final_destination_path = os.path.join(final_download_path, final_file_name)
    os.makedirs(os.path.dirname(final_destination_path), exist_ok=True)

    try:
        print(f"Attempting to move '{os.path.basename(downloaded_file_path)}' to '{final_destination_path}'")
        if os.path.exists(final_destination_path):
            print(f"Warning: File '{final_destination_path}' already exists. Overwriting.")
            os.remove(final_destination_path)
        shutil.move(downloaded_file_path, final_destination_path) # Use shutil.move for robustness
        print(f"Successfully moved file to '{final_destination_path}'")
    except Exception as e:
        print(f"Error moving file from '{downloaded_file_path}' to '{final_destination_path}': {e}")
        raise

def download_and_rename(wait, shadow_doc2_context, weeknum_for_file, default_dir_path, final_download_path, driver_instance, year_filter_val, today_timestamp_str):
    """Downloads and renames the file for the given week number."""
    print("-" * 10 + f" Starting download for Year Filter '{year_filter_val}', Week {weeknum_for_file} " + "-" * 10)

    try:
        print("Locating and clicking main download button (download-ToolbarButton)...")
        download_button = wait.until(EC.element_to_be_clickable((By.ID, "download-ToolbarButton")))
        driver_instance.execute_script("arguments[0].scrollIntoView(true);", download_button)
        time.sleep(0.5) # Brief pause after scroll
        download_button.click()
        print("Clicked main download button.")
        time.sleep(3) # Allow dialog to appear

        print("Locating and clicking 'Crosstab' button in dialog...")
        crosstab_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]')))
        crosstab_button.click()
        print("Clicked 'Crosstab' button.")
        time.sleep(3) # Allow next dialog/options to appear

        print("Locating and selecting 'CSV' option...")
        csv_radio_selector = "input[type='radio'][value='csv']"
        time.sleep(1) # Allow elements to render in dialog
        # Use shadow_doc2_context to find element, then driver_instance to execute script on it
        csv_div = shadow_doc2_context.find_element(By.CSS_SELECTOR, csv_radio_selector)
        driver_instance.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", csv_div)
        print("Selected 'CSV' option.")
        time.sleep(2) # Allow selection to register

        print("Locating and clicking final 'Download' (Export) button...")
        export_button_selector = '[data-tb-test-id="export-crosstab-export-Button"]'
        export_button = shadow_doc2_context.find_element(By.CSS_SELECTOR, export_button_selector)
        driver_instance.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", export_button)
        print(f"Clicked Export/Download for Week {weeknum_for_file}. Waiting for file...")

        # Construct filename based on your original script's logic
        new_file_name_stem = f"PAHO_2014_2024_W{weeknum_for_file:02d}_{today_timestamp_str}"
        file_extension = '.csv'

        move_to_download_folder(default_dir_path, final_download_path, new_file_name_stem, file_extension)
        print(f"Successfully downloaded and processed file for Week {weeknum_for_file}.")

    except TimeoutException as te:
        print(f"Timeout during download process for week {weeknum_for_file}: {te}")
        driver_instance.save_screenshot(f"err_download_timeout_wk{weeknum_for_file}.png")
        raise
    except Exception as e:
        print(f"Error during download process for week {weeknum_for_file}: {type(e).__name__} - {e}")
        driver_instance.save_screenshot(f"err_download_generic_wk{weeknum_for_file}.png")
        raise

def iterate_weekly():
    year_filter_value = "(All)"
    today_timestamp = datetime.now().strftime('%Y%m%d%H%M')

    # Determine if running in GitHub Actions
    is_github_actions = os.getenv('GITHUB_ACTIONS') == 'true'
    print(f"Running in GitHub Actions: {is_github_actions}")

    # --- Directory Setup ---
    # Base directory for all data output
    github_workspace_env = os.getenv('GITHUB_WORKSPACE', os.getcwd()) # Default to current dir if not in GHA
    base_output_dir = os.path.join(github_workspace_env, 'data_output')

    # Temporary directory for Chrome downloads
    if is_github_actions:
        temp_download_dir_for_chrome = os.path.abspath(os.path.join(github_workspace_env, 'temp_chrome_downloads'))
    else: # Local execution
        temp_download_dir_for_chrome = os.path.abspath(os.path.join(os.getcwd(), "temp_chrome_downloads"))

    os.makedirs(temp_download_dir_for_chrome, exist_ok=True)
    print(f"Chrome configured to download files to: {temp_download_dir_for_chrome}")

    # Final destination for processed files (dated subfolder)
    today_dated_folder_name = f"OD_DL_{datetime.now().strftime('%Y%m%d')}"
    final_file_destination_path = os.path.join(base_output_dir, today_dated_folder_name)
    os.makedirs(final_file_destination_path, exist_ok=True)
    print(f"Processed files will be moved to: {final_file_destination_path}")

    driver = None
    try:
        print("Setting up undetected-chromedriver...")
        chrome_options = uc.ChromeOptions()
        prefs = {
            "download.default_directory": temp_download_dir_for_chrome,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0
        }
        chrome_options.add_experimental_option("prefs", prefs)

        if is_github_actions:
            print("Applying GitHub Actions specific Chrome options (headless, no-sandbox, etc.).")
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--proxy-server='direct://'")
            chrome_options.add_argument("--proxy-bypass-list=*")
            chrome_options.add_argument("--start-maximized")

        driver_version_main = chrome_version_detected
        driver = uc.Chrome(
            options=chrome_options,
            version_main=driver_version_main,
            headless=is_github_actions,
            use_subprocess=True
        )
        if not is_github_actions:
            driver.maximize_window()

        print("Navigating to PAHO data page...")
        driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')

        wait = WebDriverWait(driver, 45)

        print("Switching to the first iframe (main viz)...")
        iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
        iframe_locator = (By.CSS_SELECTOR, f"iframe[src='{iframe_src}']")
        wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
        print("Switched to first iframe.")
        time.sleep(5)

        print("Looking for nested iframe within the first iframe's content...")
        nested_iframe_locator = (By.TAG_NAME, "iframe")
        wait.until(EC.frame_to_be_available_and_switch_to_it(nested_iframe_locator))
        print("Switched to the nested iframe (vizcontent).")
        time.sleep(5)

        shadow_doc2_context = driver.execute_script('return document')

        iframe_page_title = driver.title
        print(f"Title of nested iframe content: {iframe_page_title}")
        if "PAHO/WHO Data" not in iframe_page_title:
            raise Exception(f"Wrong iframe content loaded. Title: {iframe_page_title}")
        print("Successfully accessed the main dashboard content.")
        time.sleep(5)

        print(f"Selecting year filter value: '{year_filter_value}'...")
        year_tab_id = 'tabZoneId13'
        year_tab = wait.until(EC.visibility_of_element_located((By.ID, year_tab_id)))
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = year_tab.find_element(*dd_locator)
        dd_open.click()
        print("Clicked year dropdown.")
        time.sleep(2)

        year_xpath = f'//div[contains(@class, "facetOverflow")]//a[text()="{year_filter_value}"]/preceding-sibling::input'
        year_element_for_click = shadow_doc2_context.find_element(By.XPATH, year_xpath)
        driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", year_element_for_click)
        print(f"Selected year '{year_filter_value}'.")
        time.sleep(1)

        print("Closing year dropdown...")
        dd_close_locator = (By.CLASS_NAME, "tab-glass")
        dd_close = wait.until(EC.element_to_be_clickable(dd_close_locator))
        dd_close.click()
        print("Closed year dropdown.")
        time.sleep(5)

        # --- SET EPI WEEK TO 53 ---
        print("-" * 30)
        print("Ensuring Epidemiological Week is set to 53...")
        TARGET_WEEK_TO_SET = 53
        # Increased timeout specifically for these sensitive interactions
        WEEK_INTERACTION_TIMEOUT = 40 # Increased from 30

        SLIDER_TEXT_LOCATOR_WEEK = (By.CSS_SELECTOR, ".sliderText")
        WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR = (By.ID, "dijit_form_Button_3")
        SEARCH_INPUT_TEXT_FIELD_LOCATOR = (By.ID, "dijit_form_ComboBox_0")

        current_week_value_read = -1
        try:
            print(f"Attempting to read current week from: {SLIDER_TEXT_LOCATOR_WEEK}")
            slider_text_elements = wait.until(EC.presence_of_all_elements_located(SLIDER_TEXT_LOCATOR_WEEK))
            visible_slider_text_element = next((elem for elem in slider_text_elements if elem.is_displayed()), None)
            if visible_slider_text_element:
                current_text = visible_slider_text_element.text.strip()
                cleaned_text = "".join(filter(str.isdigit, current_text))
                if cleaned_text:
                    current_week_value_read = int(cleaned_text)
                    print(f"Current week detected as: {current_week_value_read}")
            else: print(f"No visible {SLIDER_TEXT_LOCATOR_WEEK} found.")
        except TimeoutException: print(f"Timeout waiting for {SLIDER_TEXT_LOCATOR_WEEK} to read initial value.")
        except Exception as e_read: print(f"Error reading initial week value: {e_read}")

        if current_week_value_read != TARGET_WEEK_TO_SET:
            print(f"Current week {current_week_value_read if current_week_value_read != -1 else 'unknown'} is not {TARGET_WEEK_TO_SET}. Updating...")

            # --- Step 1: Click Search Activator ---
            search_activator_clicked_successfully = False
            try:
                print(f"Locating week search activator: {WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}")
                # First, wait for presence, then visibility, then try to click
                search_activator = WebDriverWait(driver, WEEK_INTERACTION_TIMEOUT).until(
                    EC.presence_of_element_located(WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR)
                )
                print("Search activator present. Scrolling into view...")
                driver.execute_script("arguments[0].scrollIntoView(true);", search_activator)
                time.sleep(0.5) # Pause after scroll

                print("Waiting for search activator to be clickable...")
                search_activator = WebDriverWait(driver, WEEK_INTERACTION_TIMEOUT).until(
                    EC.element_to_be_clickable(WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR)
                )
                print("Search activator clickable. Attempting standard click...")
                search_activator.click()
                search_activator_clicked_successfully = True
                print("Clicked week search activator using standard click.")

            except TimeoutException as e_activator_timeout:
                print(f"Timeout waiting for search activator ({WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}) to be clickable: {e_activator_timeout}")
                driver.save_screenshot(f"err_timeout_activator_wk{TARGET_WEEK_TO_SET}.png")
                # Fallback to JavaScript click if standard click times out on clickability
                try:
                    print("Standard click failed/timed out for activator. Attempting JavaScript click...")
                    search_activator_js = driver.find_element(*WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR) # Re-find for JS
                    driver.execute_script("arguments[0].click();", search_activator_js)
                    search_activator_clicked_successfully = True
                    print("Clicked week search activator using JavaScript click.")
                except Exception as e_js_click_activator:
                    print(f"JavaScript click on activator also failed: {e_js_click_activator}")
                    driver.save_screenshot(f"err_js_click_activator_wk{TARGET_WEEK_TO_SET}.png")
                    raise # Re-raise the original or JS error to stop
            except Exception as e_activator_other:
                print(f"Other error clicking week search activator ({WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}): {e_activator_other}")
                driver.save_screenshot(f"err_other_search_activator_wk{TARGET_WEEK_TO_SET}.png")
                raise # Re-raise

            if not search_activator_clicked_successfully:
                # This block should not be reached if exceptions above are raised, but as a safeguard:
                print("Critical error: Search activator button was not successfully clicked.")
                raise Exception("Failed to click search activator button.")

            time.sleep(3) # Increased pause for search input to appear/activate after click

            # --- Step 2: Interact with Search Input ---
            try:
                print(f"Interacting with week search input: {SEARCH_INPUT_TEXT_FIELD_LOCATOR}")
                search_input = WebDriverWait(driver, WEEK_INTERACTION_TIMEOUT).until(
                    EC.element_to_be_clickable(SEARCH_INPUT_TEXT_FIELD_LOCATOR)
                )
                print("Search input field is clickable.")
                search_input.click() # Standard click first
                search_input.clear()
                search_input.send_keys(str(TARGET_WEEK_TO_SET))
                print(f"Typed '{TARGET_WEEK_TO_SET}'.")
                time.sleep(0.5)
                search_input.send_keys(Keys.ENTER)
                print("Sent Keys.ENTER.")
                time.sleep(5) # Allow filter to apply fully
            except TimeoutException as e_input_timeout:
                print(f"Timeout interacting with week search input ({SEARCH_INPUT_TEXT_FIELD_LOCATOR}): {e_input_timeout}")
                driver.save_screenshot(f"err_timeout_search_input_wk{TARGET_WEEK_TO_SET}.png")
                raise
            except ElementNotInteractableException as e_input_interact:
                print(f"ElementNotInteractableException with week search input ({SEARCH_INPUT_TEXT_FIELD_LOCATOR}): {e_input_interact}")
                # Try JS to set value and then simulate Enter if standard send_keys fails
                try:
                    print("Standard send_keys failed for input. Attempting JavaScript value set + Enter event...")
                    search_input_js = driver.find_element(*SEARCH_INPUT_TEXT_FIELD_LOCATOR) # Re-find
                    driver.execute_script(f"arguments[0].value = '{TARGET_WEEK_TO_SET}';", search_input_js)
                    # Create and dispatch a KeyboardEvent for Enter
                    driver.execute_script("""
                        var inputElement = arguments[0];
                        var event = new KeyboardEvent('keydown', {
                            key: 'Enter',
                            code: 'Enter',
                            keyCode: 13,
                            which: 13,
                            bubbles: true,
                            cancelable: true
                        });
                        inputElement.dispatchEvent(event);
                    """, search_input_js)
                    print(f"Set value to '{TARGET_WEEK_TO_SET}' and dispatched Enter event using JavaScript.")
                    time.sleep(5) # Allow filter to apply
                except Exception as e_js_input:
                    print(f"JavaScript interaction with search input also failed: {e_js_input}")
                    driver.save_screenshot(f"err_js_interact_search_input_wk{TARGET_WEEK_TO_SET}.png")
                    raise # Re-raise the JS error or original ElementNotInteractableException
            except Exception as e_input_other:
                print(f"Other error with week search input ({SEARCH_INPUT_TEXT_FIELD_LOCATOR}): {e_input_other}")
                driver.save_screenshot(f"err_other_search_input_wk{TARGET_WEEK_TO_SET}.png")
                raise

            # --- Step 3: Optional Verification ---
            try:
                print(f"Verifying week update by checking {SLIDER_TEXT_LOCATOR_WEEK} for '{TARGET_WEEK_TO_SET}'...")
                WebDriverWait(driver, WEEK_INTERACTION_TIMEOUT).until(
                    EC.text_to_be_present_in_element_located(SLIDER_TEXT_LOCATOR_WEEK, str(TARGET_WEEK_TO_SET)))
                print(f"Verification successful: Week appears to be set to {TARGET_WEEK_TO_SET}.")
            except TimeoutException:
                print(f"Verification WARNING: Slider text did not update to {TARGET_WEEK_TO_SET} within timeout.")
        else:
            print(f"Week is already correctly set to {TARGET_WEEK_TO_SET}.")
        print("Week setting process finished.")
        print("-" * 30)
        # --- END SET EPI WEEK ---

        weeknum_for_loop = TARGET_WEEK_TO_SET
        print(f"--- Starting Download Loop for Year Filter '{year_filter_value}' from Week {weeknum_for_loop} ---")

        download_and_rename(wait, shadow_doc2_context, weeknum_for_loop, temp_download_dir_for_chrome, final_file_destination_path, driver, year_filter_value, today_timestamp)

        while weeknum_for_loop > 1:
            print("-" * 20)
            target_decrement_week = weeknum_for_loop - 1
            print(f"Attempting to decrement to Week Number: {target_decrement_week}")
            try:
                decrement_locator = (By.XPATH, "//*[contains(@class, 'tableauArrowDec') or contains(@class, 'dijitSliderDecrementIconH')]")
                decrement_button = wait.until(EC.element_to_be_clickable(decrement_locator))
                decrement_button.click()
                print(f"Clicked decrement button. Aiming for week {target_decrement_week}.")
                time.sleep(6)
            except Exception as e_dec:
                print(f"Error clicking decrement button for week {target_decrement_week}: {e_dec}")
                driver.save_screenshot(f"err_decrement_week_{target_decrement_week}.png")
                break
            weeknum_for_loop = target_decrement_week
            download_and_rename(wait, shadow_doc2_context, weeknum_for_loop, temp_download_dir_for_chrome, final_file_destination_path, driver, year_filter_value, today_timestamp)
        print(f"--- Finished Download Loop for Year Filter '{year_filter_value}' ---")

    except Exception as e:
        print(f"An critical error occurred in iterate_weekly: {type(e).__name__} - {e}")
        if driver:
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            # Save screenshot in the current working directory (which is workspace root in GHA)
            screenshot_path = os.path.join(os.getcwd(), f"critical_error_main_process_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Screenshot saved: {screenshot_path}")
            except Exception as scr_e:
                print(f"Could not save screenshot: {scr_e}")
    finally:
        if driver:
            print("Closing WebDriver...")
            driver.quit()
        print("Script finished.")

if __name__ == "__main__":
    iterate_weekly()
