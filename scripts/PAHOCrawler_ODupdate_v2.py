import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys # Added for Keys.ENTER
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException # Added for specific error handling
import time
import os
from datetime import datetime
import subprocess
import re
import sys

# Function to get the installed Chrome version
def get_chrome_version():
    try:
        version_match = None
        if sys.platform == "win32":
            # Try more robust registry queries for default Chrome install locations
            reg_paths = [
                r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Google\Chrome\Update" /v LastKnownVersionString /reg:32',
                r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Google\Chrome\Update" /v LastKnownVersionString /reg:64',
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version' # Fallback
            ]
            for path in reg_paths:
                try:
                    output = subprocess.check_output(path, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'REG_SZ\s+(\d+)\.', output)
                    if version_match:
                        break
                except (subprocess.CalledProcessError, FileNotFoundError):
                    continue
        else:  # Assuming Linux or macOS
            commands = [
                '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome --version', # macOS default
                'google-chrome --version',
                'google-chrome-stable --version',
                'chromium-browser --version',
                'chromium --version'
            ]
            for command in commands:
                try:
                    output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'\b(\d+)\.', output)
                    if version_match:
                        break
                except (subprocess.CalledProcessError, FileNotFoundError):
                    continue

        if version_match:
            print(f"Detected Chrome major version: {version_match.group(1)}")
            return int(version_match.group(1))
        else:
            raise ValueError("Could not parse Chrome version from any attempted method.")

    except Exception as e:
        raise RuntimeError(f"Failed to get Chrome version: {e}") from e

# Get the major version of Chrome installed
try:
    chrome_version = get_chrome_version()
except RuntimeError as e:
    print(f"Error getting Chrome version: {e}")
    print("Please ensure Chrome is installed or specify version_main in uc.Chrome() manually.")
    # Decide if to exit or use a default
    # For now, let's allow uc.Chrome to try and auto-detect if this fails
    chrome_version = None # uc.Chrome will try to autodetect if version_main is None or not provided

def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    max_wait_time = 180  # Increased wait time for download (3 minutes)
    check_interval = 5   # How often to check for the file
    start_time = time.time()
    currentFile = None
    processed_files = set() # Keep track of files already considered

    print(f"Waiting for download in '{default_dir}' to be named '{newFileName}{fileExtension}'...")

    while not got_file and (time.time() - start_time) < max_wait_time:
        try:
            # List files, ignore temp download files and already processed ones
            files_in_dir = [
                f for f in os.listdir(default_dir)
                if not f.lower().endswith(('.crdownload', '.tmp')) and \
                   os.path.isfile(os.path.join(default_dir, f)) and \
                   os.path.join(default_dir, f) not in processed_files
            ]

            if not files_in_dir:
                # print(f"No new completed files found yet in '{default_dir}'. Waiting {check_interval}s...")
                time.sleep(check_interval)
                continue

            # Get the latest (by modification time) new, completed file
            latest_file_path = None
            latest_mod_time = 0
            for f_name in files_in_dir:
                f_path = os.path.join(default_dir, f_name)
                mod_time = os.path.getmtime(f_path) # Use modification time
                if mod_time > latest_mod_time:
                    latest_mod_time = mod_time
                    latest_file_path = f_path

            if not latest_file_path: # Should not happen if files_in_dir is populated
                time.sleep(check_interval)
                continue

            currentFile = latest_file_path
            processed_files.add(currentFile) # Mark as considered

            # Check if the file is still being written (basic size check)
            initial_size = os.path.getsize(currentFile)
            time.sleep(2) # Wait a moment to see if size changes
            current_size = os.path.getsize(currentFile)

            if initial_size == current_size and current_size > 0: # File size stable and not empty
                print(f"Detected downloaded file: {os.path.basename(currentFile)} (Size: {current_size} bytes)")
                got_file = True
            else:
                print(f"File '{os.path.basename(currentFile)}' size changed or is empty (Initial: {initial_size}, Current: {current_size}). Re-evaluating...")
                got_file = False # Ensure it re-evaluates or picks another if this one is unstable
                processed_files.remove(currentFile) # Allow it to be re-checked
                time.sleep(check_interval - 2)


        except FileNotFoundError:
            print(f"Download directory '{default_dir}' not found during check. Retrying...")
            time.sleep(check_interval)
        except Exception as e:
            print(f"Error checking download directory: {e}. Retrying...")
            time.sleep(check_interval)

    if not got_file or not currentFile:
        print(f"Error: Download did not complete or file not detected within {max_wait_time} seconds.")
        raise TimeoutError("Download timeout exceeded or file not found.")

    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)
    os.makedirs(os.path.dirname(fileDestination), exist_ok=True)

    try:
        if os.path.exists(fileDestination):
            print(f"Warning: File '{fileDestination}' already exists. Overwriting.")
            os.remove(fileDestination)
        os.rename(currentFile, fileDestination)
        print(f"Moved '{os.path.basename(currentFile)}' to '{fileDestination}'")
    except OSError as oe:
        print(f"OSError moving file from '{currentFile}' to '{fileDestination}': {oe}")
        try:
            import shutil
            shutil.move(currentFile, fileDestination) # Try shutil.move as a fallback
            print(f"Moved file using shutil.move to '{fileDestination}'")
        except Exception as sh_e:
            print(f"Shutil.move also failed: {sh_e}")
            raise # Re-raise the original error or the shutil error

def download_and_rename(wait, shadow_doc2_context, weeknum_for_file, default_dir_path, final_download_path, driver_instance, year_str, today_str):
    """Download and rename the file for the given week number."""
    print("-" * 10 + f" Starting download for Year '{year_str}', Week {weeknum_for_file} " + "-" * 10)
    DOWNLOAD_DIALOG_TIMEOUT = 30

    # It's good practice to use the passed weeknum_for_file for filename consistency
    # rather than re-reading from sliderText here, as sliderText might lag or be in transition.
    # weeknum_div = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "sliderText")))
    # actual_week_on_slider = int("".join(filter(str.isdigit, weeknum_div.text.strip())))
    # if actual_week_on_slider != weeknum_for_file:
    # print(f"Note: Slider shows week {actual_week_on_slider}, but downloading for requested week {weeknum_for_file}")

    try:
        print("Locating and clicking main download button (download-ToolbarButton)...")
        download_button = wait.until(
            EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
        )
        driver_instance.execute_script("arguments[0].scrollIntoView(true);", download_button)
        time.sleep(0.5)
        download_button.click()
        print("Clicked main download button.")
        time.sleep(3)

        print("Locating and clicking 'Crosstab' button in dialog...")
        crosstab_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
        )
        crosstab_button.click()
        print("Clicked 'Crosstab' button.")
        time.sleep(3)

        # CSV option and Export button are within shadow_doc2_context (likely iframe's document)
        print("Locating and selecting 'CSV' option...")
        # The find_element for shadow content should be on the shadow_doc2_context directly if it's a shadow root
        # or via driver.execute_script if it's just a document object.
        # Given the original script structure, we assume shadow_doc2_context is the correct search base.
        csv_radio_selector = "input[type='radio'][value='csv']"
        # Wait for presence using driver and then get element via execute_script or direct find if possible.
        # For elements possibly inside shadow DOM, direct wait on shadow_doc2_context is not standard.
        # We'll use a short pause then find, assuming it's quick.
        time.sleep(1) # Allow elements to render
        csv_div = shadow_doc2_context.find_element(By.CSS_SELECTOR, csv_radio_selector)
        driver_instance.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", csv_div)
        print("Selected 'CSV' option.")
        time.sleep(2)

        print("Locating and clicking final 'Download' (Export) button...")
        export_button_selector = '[data-tb-test-id="export-crosstab-export-Button"]'
        export_button = shadow_doc2_context.find_element(By.CSS_SELECTOR, export_button_selector)
        driver_instance.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", export_button)
        print(f"Clicked Export/Download for Week {weeknum_for_file}. Waiting for file...")

        # Filename construction (matches original script's hardcoded year span)
        # The `year_str` parameter here is from the main function (e.g., "(All)")
        # but the filename is hardcoded, so we keep it.
        newFileName = f"PAHO_2014_2024_W{weeknum_for_file:02d}_{today_str}"
        fileExtension = '.csv'

        move_to_download_folder(default_dir_path, final_download_path, newFileName, fileExtension)
        print(f"Successfully downloaded and processed file for Week {weeknum_for_file}.")

    except TimeoutException as te:
        print(f"Timeout during download process for week {weeknum_for_file}: {te}")
        # Consider adding screenshot here
        raise
    except Exception as e:
        print(f"Error during download process for week {weeknum_for_file}: {type(e).__name__} - {e}")
        # Consider adding screenshot here
        raise

def iterate_weekly():
    # Year variable from original script
    year_filter_value = "(All)" # choose year to download, matches original script
    today_timestamp = datetime.now().strftime('%Y%m%d%H%M') # current date and time

    # Set directories
    github_workspace_env = os.getenv('GITHUB_WORKSPACE')
    if github_workspace_env:
        base_dir_for_data = os.path.join(github_workspace_env, 'data')
    else:
        # Fallback for local execution, place 'data' folder where script is, or user's preferred location
        base_dir_for_data = os.path.join(os.getcwd(), 'data')
        # base_dir_for_data = 'C:/Users/AhyoungLim/Dropbox/WORK/OpenDengue/PAHO-crawler/data' # Example

    # Chrome's default download directory (can be temporary)
    # This MUST be an absolute path.
    default_download_dir_for_chrome = os.path.abspath(os.path.join(os.getcwd(), "temp_downloads"))
    os.makedirs(default_download_dir_for_chrome, exist_ok=True)
    print(f"Chrome configured to download files to: {default_download_dir_for_chrome}")

    # Final destination for processed files
    today_dated_folder_name = f"OD_DL_{datetime.now().strftime('%Y%m%d')}"
    final_file_destination_path = os.path.join(base_dir_for_data, today_dated_folder_name)
    os.makedirs(final_file_destination_path, exist_ok=True)
    print(f"Processed files will be moved to: {final_file_destination_path}")

    driver = None
    try:
        print("Setting up undetected-chromedriver...")
        chrome_options = uc.ChromeOptions()
        prefs = {"download.default_directory": default_download_dir_for_chrome,
                 "download.prompt_for_download": False,
                 "safeBrowse.enabled": True}
        chrome_options.add_experimental_option("prefs", prefs)
        # chrome_options.add_argument('--headless') # Uncomment for headless
        # chrome_options.add_argument('--disable-gpu') # Usually with headless

        driver_version_main = chrome_version if chrome_version else None # Use detected version if available
        driver = uc.Chrome(headless=False, use_subprocess=False, options=chrome_options, version_main=driver_version_main)
        driver.maximize_window()

        print("Navigating to PAHO data page...")
        driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')

        wait = WebDriverWait(driver, 30) # Increased default wait time

        print("Switching to the first iframe (main viz)...")
        iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
        iframe_locator = (By.CSS_SELECTOR, f"iframe[src='{iframe_src}']") # More robust selector
        # Wait for frame to be available and switch
        wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
        print("Switched to first iframe.")
        time.sleep(3) # Allow iframe content to load

        print("Looking for nested iframe within the first iframe's content...")
        nested_iframe_locator = (By.TAG_NAME, "iframe")
        # Wait for the nested iframe to be present AND switch to it
        wait.until(EC.frame_to_be_available_and_switch_to_it(nested_iframe_locator))
        print("Switched to the nested iframe (vizcontent).")
        time.sleep(3)

        # This context is what shadow_doc2 represented, for year/CSV interactions.
        # We will use 'driver' now as it's correctly focused.
        # For clarity, if a specific shadow root is needed from this point, it would be:
        # shadow_host = driver.find_element(By.CSS_SELECTOR, "host-element-selector")
        # shadow_root_context = driver.execute_script('return arguments[0].shadowRoot', shadow_host)
        # But based on your working year selection, direct interaction after frame switch was okay.
        # Let's assume the current 'driver' context is what 'shadow_doc2' effectively was.
        # So, where `shadow_doc2.find_element` was used, we'll now ensure driver is in the right frame
        # and then use `driver.find_element` or `wait.until(EC.element_to_be_clickable(...))`
        # For `download_and_rename`, we'll pass the `driver` object as the search context for shadow elements.
        # No, the user's `shadow_doc2 = driver.execute_script('return document')` is what they used.
        # This gets the document object of the current frame.
        shadow_doc2_context = driver.execute_script('return document')


        iframe_page_title = driver.title
        print(f"Title of nested iframe content: {iframe_page_title}")
        if "PAHO/WHO Data" not in iframe_page_title:
            print(f"Error: Unexpected content in the nested iframe. Title: {iframe_page_title}")
            raise Exception("Wrong iframe content loaded.")
        print("Successfully accessed the main dashboard content.")
        time.sleep(5)

        print(f"Selecting year filter value: '{year_filter_value}'...")
        year_tab_id = 'tabZoneId13'
        year_tab = wait.until(EC.visibility_of_element_located((By.ID, year_tab_id)))
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = year_tab.find_element(*dd_locator)
        dd_open.click()
        print("Clicked year dropdown.")
        time.sleep(1)

        # Select the specified year (e.g., "(All)") using the shadow_doc2_context
        year_xpath = f'//div[contains(@class, "facetOverflow")]//a[text()="{year_filter_value}"]/preceding-sibling::input'
        # Wait for element to be present using WebDriverWait on the driver, then find with shadow_doc2_context
        # This is a bit tricky. A simpler way is to ensure the element is there via wait, then operate.
        # Let's try waiting for the element to be clickable using the main driver context first.
        year_input_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, year_xpath))
        )
        # If it's found via driver, it means it's not strictly needing shadow_doc2_context for location by Selenium's wait
        # But original script used shadow_doc2_context.find_element.
        # For JS click, we pass the element found by Selenium's main driver context (if it works) or found via shadow_doc2_context
        # For now, keeping the original interaction pattern:
        # Wait for items to load in dropdown
        time.sleep(1) # Give dropdown items time to load
        year_element_for_click = shadow_doc2_context.find_element(By.XPATH, year_xpath)
        driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", year_element_for_click)

        print(f"Selected year '{year_filter_value}'.")
        time.sleep(0.5)

        print("Closing year dropdown...")
        dd_close_locator = (By.CLASS_NAME, "tab-glass")
        dd_close = wait.until(EC.element_to_be_clickable(dd_close_locator))
        dd_close.click()
        print("Closed year dropdown.")
        time.sleep(3)

        # --- SET EPI WEEK TO 53 ---
        print("-" * 30)
        print("Ensuring Epidemiological Week is set to 53...")
        TARGET_WEEK_TO_SET = 53
        WEEK_INTERACTION_TIMEOUT = 20

        SLIDER_TEXT_LOCATOR_WEEK = (By.CSS_SELECTOR, ".sliderText")
        WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR = (By.ID, "dijit_form_Button_3")
        SEARCH_INPUT_TEXT_FIELD_LOCATOR = (By.ID, "dijit_form_ComboBox_0")

        current_week_value_read = -1
        try:
            print(f"Attempting to read current week from: {SLIDER_TEXT_LOCATOR_WEEK}")
            # Wait for presence of *any* sliderText, then check visibility for the correct one
            slider_text_elements = wait.until(EC.presence_of_all_elements_located(SLIDER_TEXT_LOCATOR_WEEK))
            visible_slider_text_element = None
            for elem in slider_text_elements:
                if elem.is_displayed(): # Check visibility
                    visible_slider_text_element = elem
                    break

            if visible_slider_text_element:
                current_text = visible_slider_text_element.text.strip()
                cleaned_text = "".join(filter(str.isdigit, current_text))
                if cleaned_text:
                    current_week_value_read = int(cleaned_text)
                    print(f"Current week detected as: {current_week_value_read}")
                else:
                    print(f"Could not parse digits from slider text: '{current_text}'. Proceeding to update.")
            else:
                print(f"No visible {SLIDER_TEXT_LOCATOR_WEEK} found. Proceeding to set week to {TARGET_WEEK_TO_SET}.")
        except TimeoutException:
            print(f"Timeout waiting for {SLIDER_TEXT_LOCATOR_WEEK}. Assuming week needs to be set.")
        except Exception as e_read:
            print(f"Could not read initial week value: {e_read}. Will attempt to set to {TARGET_WEEK_TO_SET}.")

        if current_week_value_read != TARGET_WEEK_TO_SET:
            if current_week_value_read != -1:
                print(f"Current week {current_week_value_read} is not {TARGET_WEEK_TO_SET}. Updating...")
            else:
                print(f"Attempting to set week to {TARGET_WEEK_TO_SET}.")

            try:
                print(f"Locating and clicking week search activator: {WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}")
                search_activator = wait.until(EC.element_to_be_clickable(WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR))
                search_activator.click()
                print("Clicked week search activator.")
                time.sleep(1.5)
            except TimeoutException:
                print(f"Error: Could not click Week Search Activator ({WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}).")
                driver.save_screenshot("err_click_week_search_activator.png")
                raise

            try:
                print(f"Interacting with week search input: {SEARCH_INPUT_TEXT_FIELD_LOCATOR}")
                search_input = wait.until(EC.element_to_be_clickable(SEARCH_INPUT_TEXT_FIELD_LOCATOR))
                search_input.click()
                search_input.clear()
                search_input.send_keys(str(TARGET_WEEK_TO_SET))
                print(f"Typed '{TARGET_WEEK_TO_SET}'.")
                time.sleep(0.5)
                search_input.send_keys(Keys.ENTER)
                print("Sent Keys.ENTER.")
                time.sleep(3) # Allow filter to apply
            except TimeoutException:
                print(f"Timeout: Week search input ({SEARCH_INPUT_TEXT_FIELD_LOCATOR}) not clickable.")
                driver.save_screenshot("err_timeout_week_search_input.png")
                raise
            except ElementNotInteractableException as e:
                print(f"ElementNotInteractable: {e}")
                driver.save_screenshot("err_interactable_week_search_input.png")
                raise

            try: # Optional verification
                print(f"Verifying week update to {TARGET_WEEK_TO_SET}...")
                WebDriverWait(driver, WEEK_INTERACTION_TIMEOUT).until(
                    EC.and_(
                        EC.visibility_of_element_located(SLIDER_TEXT_LOCATOR_WEEK),
                        EC.text_to_be_present_in_element_located(SLIDER_TEXT_LOCATOR_WEEK, str(TARGET_WEEK_TO_SET))
                    )
                )
                print(f"Verification successful: Week is {TARGET_WEEK_TO_SET}.")
            except TimeoutException:
                print(f"Verification WARNING: Timeout waiting for slider text to update to {TARGET_WEEK_TO_SET}.")
                # driver.save_screenshot("warn_verify_timeout.png") # screenshot for warning
        else:
            print(f"Week is already correctly set to {TARGET_WEEK_TO_SET}.")
        print("Week setting process finished.")
        print("-" * 30)
        # --- END SET EPI WEEK ---

        # --- Download Loop ---
        weeknum_for_loop = TARGET_WEEK_TO_SET # Start from the week we just ensured

        print(f"--- Starting Download Loop for Year Filter '{year_filter_value}' from Week {weeknum_for_loop} ---")

        # Initial download for the starting week (e.g., 53)
        print(f"Initial download for Week Number: {weeknum_for_loop}")
        # Pass shadow_doc2_context which is `driver.execute_script('return document')` from the correct iframe
        download_and_rename(wait, shadow_doc2_context, weeknum_for_loop, default_download_dir_for_chrome, final_file_destination_path, driver, year_filter_value, today_timestamp)


        # Loop downwards from week (e.g., 53-1) down to 1
        while weeknum_for_loop > 1:
            print("-" * 20)
            # The week we want to *change to* and then *download*
            target_decrement_week = weeknum_for_loop - 1
            print(f"Attempting to decrement to Week Number: {target_decrement_week}")

            try:
                # More specific XPaths for decrement button might be needed if too generic
                decrement_locator = (By.XPATH, "//*[contains(@class, 'tableauArrowDec') or contains(@class, 'dijitSliderDecrementIconH') or contains(@title, 'Decrement')]")
                decrement_button = wait.until(EC.element_to_be_clickable(decrement_locator))
                decrement_button.click()
                print(f"Clicked decrement button. Aiming for week {target_decrement_week}.")
                time.sleep(4) # Allow time for week change and data update
            except TimeoutException:
                print(f"Error: Could not find or click the decrement button to reach week {target_decrement_week}.")
                driver.save_screenshot(f"err_decrement_week_{target_decrement_week}.png")
                break # Stop loop if decrement fails

            weeknum_for_loop = target_decrement_week # Update current week after successful decrement
            download_and_rename(wait, shadow_doc2_context, weeknum_for_loop, default_download_dir_for_chrome, final_file_destination_path, driver, year_filter_value, today_timestamp)

        print(f"--- Finished Download Loop for Year Filter '{year_filter_value}' ---")

    except Exception as e:
        print(f"An critical error occurred in iterate_weekly: {type(e).__name__} - {e}")
        if driver:
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"critical_error_main_process_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Screenshot of critical error saved to: {screenshot_path}")
            except Exception as scr_e:
                print(f"Could not save screenshot during critical error handling: {scr_e}")
    finally:
        if driver:
            print("Closing WebDriver...")
            driver.quit()
        print("Script finished.")

if __name__ == "__main__":
    iterate_weekly()
