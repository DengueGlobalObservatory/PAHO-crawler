import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import time
import os
from datetime import datetime
import subprocess
import re
import sys
import shutil # For robust file moving
import glob # For cleaning up temporary files

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
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version' # User's original
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
                    version_match = re.search(r'Google Chrome (\d+)\.', output) # Regex for macOS output
                    if version_match: break
                except (subprocess.CalledProcessError, FileNotFoundError): continue
        else:  # Linux
            commands = [ # User's original list
                'google-chrome --version',
                'google-chrome-stable --version',
                'chromium-browser --version',
                'chromium --version' # Added for completeness
            ]
            for command in commands:
                try:
                    output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'\b(\d+)\.', output) # General regex for version numbers
                    if version_match: break
                except (subprocess.CalledProcessError, FileNotFoundError): continue

        if version_match:
            version = int(version_match.group(1))
            print(f"Detected Chrome major version: {version}")
            return version
        else:
            # If specific commands fail, try a more generic approach for Linux if possible
            if sys.platform.startswith('linux'):
                try: # Check common path
                    output = subprocess.check_output(['/usr/bin/google-chrome-stable', '--version'], text=True, stderr=subprocess.DEVNULL)
                    version_match = re.search(r'\b(\d+)\.', output)
                    if version_match:
                        version = int(version_match.group(1))
                        print(f"Detected Chrome major version (fallback): {version}")
                        return version
                except (subprocess.CalledProcessError, FileNotFoundError):
                    pass
            print("Could not automatically determine Chrome version using common methods. uc.Chrome will attempt auto-detection.")
            return None

    except Exception as e:
        print(f"Warning: Failed to get Chrome version due to: {e}. uc.Chrome will attempt auto-detection.")
        return None

# Get the major version of Chrome installed (can be None)
chrome_version_detected = get_chrome_version()


def move_to_download_folder(default_dir_path, final_download_path, new_file_name_stem, file_extension):
    """
    Waits for a download to complete in default_dir_path.
    First, renames it to a temporary contextual name in default_dir_path.
    Then, moves this temporary file to final_download_path with its final name.
    """
    got_file = False
    max_wait_time = 180  # Max time to wait for download (3 minutes)
    check_interval = 5   # How often to check for the file
    start_time = time.time()
    raw_downloaded_file_path = None # Path to the initially downloaded file (e.g., "W By Last...")

    print(f"Waiting for download in '{default_dir_path}' for up to {max_wait_time}s...")

    while not got_file and (time.time() - start_time) < max_wait_time:
        try:
            candidate_files = [
                os.path.join(default_dir_path, f) for f in os.listdir(default_dir_path)
                if not f.lower().endswith(('.crdownload', '.tmp', '.part')) and \
                   os.path.isfile(os.path.join(default_dir_path, f)) and \
                   not f.endswith("_TEMPRAW" + file_extension) # Ignore already temp-renamed files from previous failed attempts
            ]

            if not candidate_files:
                time.sleep(check_interval)
                continue

            latest_file = max(candidate_files, key=os.path.getmtime)

            time.sleep(1)
            initial_size = os.path.getsize(latest_file)
            time.sleep(2)
            current_size = os.path.getsize(latest_file)

            if initial_size == current_size and current_size > 0:
                print(f"Detected raw downloaded file: {os.path.basename(latest_file)} (Size: {current_size} bytes)")
                raw_downloaded_file_path = latest_file
                got_file = True
            else:
                print(f"File '{os.path.basename(latest_file)}' size changed or is empty. Re-evaluating...")
                time.sleep(check_interval - 3)

        except FileNotFoundError:
            print(f"Download directory '{default_dir_path}' not found during check. Retrying...")
            time.sleep(check_interval)
        except Exception as e:
            print(f"Error checking download directory: {e}. Retrying...")
            time.sleep(check_interval)

    if not got_file or not raw_downloaded_file_path:
        print(f"Error: Raw download did not complete or file not detected within {max_wait_time} seconds in '{default_dir_path}'.")
        raise TimeoutError(f"Raw download timeout or file not found in '{default_dir_path}'.")

    # --- First stage: Rename to a temporary contextual name ---
    temp_contextual_filename_stem = new_file_name_stem + "_TEMPRAW"
    temp_contextual_filepath = os.path.join(default_dir_path, temp_contextual_filename_stem + file_extension)

    try:
        print(f"Renaming raw file '{os.path.basename(raw_downloaded_file_path)}' to temporary contextual name '{os.path.basename(temp_contextual_filepath)}'")
        if os.path.exists(temp_contextual_filepath): # Should not happen if logic is correct, but as a safeguard
            print(f"Warning: Temporary contextual file '{temp_contextual_filepath}' already exists. Removing it.")
            os.remove(temp_contextual_filepath)
        os.rename(raw_downloaded_file_path, temp_contextual_filepath)
        print(f"Successfully renamed to temporary: {temp_contextual_filepath}")
    except Exception as e_rename_temp:
        print(f"Error renaming raw file to temporary contextual name: {e_rename_temp}")
        # If this rename fails, the original raw file is still there. The finally block in iterate_weekly will try to clean it.
        raise # Re-raise, as we can't proceed with this file.

    # --- Second stage: Move the temporary contextual file to its final destination with the final name ---
    final_file_name = new_file_name_stem + file_extension # This is the name without _TEMPRAW
    final_destination_path = os.path.join(final_download_path, final_file_name)
    os.makedirs(os.path.dirname(final_destination_path), exist_ok=True)

    try:
        print(f"Attempting to move temporary file '{os.path.basename(temp_contextual_filepath)}' to final destination '{final_destination_path}'")
        if os.path.exists(final_destination_path):
            print(f"Warning: Final destination file '{final_destination_path}' already exists. Overwriting.")
            os.remove(final_destination_path)
        shutil.move(temp_contextual_filepath, final_destination_path)
        print(f"Successfully moved file to final destination: '{final_destination_path}'")
    except Exception as e_move_final:
        print(f"Error moving temporary contextual file to final destination: {e_move_final}")
        # If this move fails, the temp_contextual_filepath (with week info) remains in default_dir_path.
        # The finally block in iterate_weekly will attempt to clean it.
        raise # Re-raise to signal that this specific download processing failed.


def download_and_rename(wait, shadow_doc2_context, weeknum_for_file, default_dir_path, final_download_path, driver_instance, year_label_for_filename, today_timestamp_str):
    """Downloads and renames the file for the given week number.
    The weeknum_for_file should be the *actual* week observed on the dashboard.
    """
    print("-" * 10 + f" Starting download for Year(s) '{year_label_for_filename}', Observed Week {weeknum_for_file} " + "-" * 10)

    try:
        print("Locating and clicking main download button (download-ToolbarButton)...")
        download_button = wait.until(EC.element_to_be_clickable((By.ID, "download-ToolbarButton")))
        driver_instance.execute_script("arguments[0].scrollIntoView(true);", download_button)
        time.sleep(0.5)
        download_button.click()
        print("Clicked main download button.")
        time.sleep(3)

        print("Locating and clicking 'Crosstab' button in dialog...")
        crosstab_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]')))
        crosstab_button.click()
        print("Clicked 'Crosstab' button.")
        time.sleep(3)

        print("Locating and selecting 'CSV' option...")
        csv_radio_selector = "input[type='radio'][value='csv']"
        time.sleep(1)
        csv_div = shadow_doc2_context.find_element(By.CSS_SELECTOR, csv_radio_selector)
        driver_instance.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", csv_div)
        print("Selected 'CSV' option.")
        time.sleep(2)

        print("Locating and clicking final 'Download' (Export) button...")
        export_button_selector = '[data-tb-test-id="export-crosstab-export-Button"]'
        export_button = shadow_doc2_context.find_element(By.CSS_SELECTOR, export_button_selector)
        driver_instance.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", export_button)
        print(f"Clicked Export/Download for Observed Week {weeknum_for_file}. Waiting for file...")

        # Construct filename stem (used for both temp and final names)
        new_file_name_stem = f"PAHO_Y{year_label_for_filename}_W{weeknum_for_file:02d}_{today_timestamp_str}"
        file_extension = '.csv'

        move_to_download_folder(default_dir_path, final_download_path, new_file_name_stem, file_extension)
        print(f"Successfully downloaded and processed file for Observed Week {weeknum_for_file}.")

    except TimeoutException as te:
        print(f"Timeout during download process for observed week {weeknum_for_file}: {te}")
        driver_instance.save_screenshot(f"err_download_timeout_wk{weeknum_for_file}.png")
        raise
    except Exception as e:
        print(f"Error during download process for observed week {weeknum_for_file}: {type(e).__name__} - {e}")
        driver_instance.save_screenshot(f"err_download_generic_wk{weeknum_for_file}.png")
        raise

def iterate_weekly():

    # Define target years for selection
    target_years_to_select = ["2023", "2024", "2025"] # User wants these selected
    # Create a label for filenames based on selected years
    year_label_for_filename = "-".join(target_years_to_select)

    today_timestamp = datetime.now().strftime('%Y%m%d%H%M')

    is_github_actions = os.getenv('GITHUB_ACTIONS') == 'true'
    print(f"Running in GitHub Actions: {is_github_actions}")

    # --- Directory Setup (User's original structure with robustness) ---
    github_workspace_env = os.getenv('GITHUB_WORKSPACE', os.getcwd())
    # User's original path for final data storage
    final_data_storage_base = os.path.join(github_workspace_env, 'data')

    # Temporary directory for Chrome downloads (must be absolute)
    default_download_dir_for_chrome = os.path.abspath(os.getcwd())
    print(f"Chrome configured to download files to: {default_download_dir_for_chrome}")

    # Final destination for processed files (dated subfolder in user's specified path)
    today_dated_folder_name = f"DL_{datetime.now().strftime('%Y%m%d')}" # User's original folder naming
    final_file_destination_path = os.path.join(final_data_storage_base, today_dated_folder_name)
    os.makedirs(final_file_destination_path, exist_ok=True)
    print(f"Processed files will be moved to: {final_file_destination_path}")

    driver = None
    # Define the known raw download filename pattern from Tableau
    raw_tableau_download_filename = "W By Last Available EpiWeek.csv"
    # Define pattern for our temporary contextual files
    temp_contextual_file_pattern = "*_TEMPRAW.csv"


    try:
        print("Setting up undetected-chromedriver...")
        chrome_options = uc.ChromeOptions()
        prefs = {
            "download.default_directory": default_download_dir_for_chrome, # Must be absolute
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
            use_subprocess=False # User's original setting
        )
        if not is_github_actions:
            driver.maximize_window()

        print("Navigating to PAHO data page...")
        driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')

        wait = WebDriverWait(driver, 45) # Increased default wait from user's 20

        print("Switching to the first iframe (main viz)...")
        iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
        # Using user's original XPATH, but CSS selector is often more robust if structure allows
        iframe_locator = (By.XPATH, f"//div[contains(@class, 'vizTab')]//iframe[@src='{iframe_src}']")
        wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
        print("Switched to first iframe.")
        time.sleep(5)

        print("Looking for nested iframe within the first iframe's content...")
        nested_iframe_locator = (By.XPATH, "//body/iframe") # User's original XPATH
        wait.until(EC.frame_to_be_available_and_switch_to_it(nested_iframe_locator))
        print("Switched to the nested iframe (vizcontent).")
        time.sleep(5)

        shadow_doc2_context = driver.execute_script('return document') # User's method for context

        iframe_page_title = driver.title
        print(f"Title of nested iframe content: {iframe_page_title}")
        if "PAHO/WHO Data" not in iframe_page_title: # More flexible check
            raise Exception(f"Wrong iframe content loaded. Title: {iframe_page_title}")
        print("Successfully accessed the main dashboard content.")
        time.sleep(5)

        # --- Select Target Years (2023, 2024, 2025) and Verify ---
        print(f"Attempting to select and verify years: {', '.join(target_years_to_select)}...")
        year_tab_id = 'tabZoneId13'
        year_tab = wait.until(EC.visibility_of_element_located((By.ID, year_tab_id)))

        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open_button = year_tab.find_element(*dd_locator)
        dd_open_button.click()
        print("Clicked year dropdown open button.")
        time.sleep(5)
        print("Taking screenshot: year_dropdown_opened.png")
        driver.save_screenshot("year_dropdown_opened.png")

        # Attempt to select each target year if not already selected
        for year_str_to_select in target_years_to_select:
            year_xpath_select = f'//div[contains(@class, "facetOverflow")]//a[text()="{year_str_to_select}"]/preceding-sibling::input'
            try:
                print(f"Processing year for selection: {year_str_to_select}")
                # Wait for the specific checkbox to be present using the main driver context
                WebDriverWait(driver, 30).until( # Increased timeout
                    EC.presence_of_element_located((By.XPATH, year_xpath_select))
                )
                year_checkbox = shadow_doc2_context.find_element(By.XPATH, year_xpath_select)

                if not year_checkbox.is_selected():
                    print(f"Checkbox for year {year_str_to_select} is not selected. Clicking...")
                    driver.execute_script("arguments[0].click();", year_checkbox)
                    time.sleep(0.7)
                    print(f"Clicked to select year {year_str_to_select}.")
                else:
                    print(f"Year {year_str_to_select} is already selected.")
            except TimeoutException:
                print(f"Timeout: Checkbox for target year {year_str_to_select} not found. This year might not be available.")
                driver.save_screenshot(f"err_year_checkbox_timeout_{year_str_to_select}.png")
                raise Exception(f"Target year {year_str_to_select} checkbox not found in dropdown. Cannot proceed.")
            except Exception as e_year_select:
                print(f"Error processing year {year_str_to_select} for selection: {e_year_select}")
                driver.save_screenshot(f"err_year_select_{year_str_to_select}.png")
                raise

        # Verification step
        print("Verifying year selections...")
        all_required_years_confirmed_selected = True
        for year_str_to_verify in target_years_to_select:
            year_xpath_verify = f'//div[contains(@class, "facetOverflow")]//a[text()="{year_str_to_verify}"]/preceding-sibling::input'
            try:
                # Re-find the element for verification
                year_checkbox_verify = shadow_doc2_context.find_element(By.XPATH, year_xpath_verify)
                if year_checkbox_verify.is_selected():
                    print(f"Verified: Year {year_str_to_verify} is selected.")
                else:
                    print(f"Verification FAILED: Year {year_str_to_verify} is NOT selected.")
                    all_required_years_confirmed_selected = False
                    driver.save_screenshot(f"err_year_verify_not_selected_{year_str_to_verify}.png")
            except Exception as e_verify:
                print(f"Verification Error: Could not find/check checkbox for year {year_str_to_verify}: {e_verify}")
                all_required_years_confirmed_selected = False
                driver.save_screenshot(f"err_year_verify_find_{year_str_to_verify}.png")

        if not all_required_years_confirmed_selected:
            error_message = f"Critical Error: Not all target years ({', '.join(target_years_to_select)}) were confirmed as selected. Quitting."
            print(error_message)
            raise Exception(error_message)

        print("All target years successfully selected and verified.")

        print("Closing year dropdown...")
        dd_close_locator = (By.CLASS_NAME, "tab-glass")
        dd_close_button = wait.until(EC.element_to_be_clickable(dd_close_locator))
        dd_close_button.click()
        print("Closed year dropdown.")
        time.sleep(5)

        # --- ADJUST EPI WEEK TO STARTING WEEK (53) USING SLIDER BUTTONS ---
        print("-" * 30)
        print("Adjusting Epidemiological Week to 53 using slider buttons...")
        TARGET_START_WEEK = 53
        SLIDER_TEXT_LOCATOR_WEEK = (By.CSS_SELECTOR, ".sliderText")
        INCREMENT_BUTTON_LOCATOR = (By.XPATH, "//*[contains(@class, 'tableauArrowInc') or contains(@class, 'dijitSliderIncrementIconH')]")
        DECREMENT_BUTTON_LOCATOR = (By.XPATH, "//*[contains(@class, 'tableauArrowDec') or contains(@class, 'dijitSliderDecrementIconH')]")

        max_slider_adjust_attempts = 70
        slider_adjust_attempts = 0
        current_week_value_read = -1

        while slider_adjust_attempts < max_slider_adjust_attempts:
            try:
                slider_text_elements = WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located(SLIDER_TEXT_LOCATOR_WEEK)
                )
                visible_slider_text_element = next((elem for elem in slider_text_elements if elem.is_displayed()), None)

                if not visible_slider_text_element:
                    print("Error: Slider text element for week not visible. Cannot proceed.")
                    driver.save_screenshot(f"err_slider_text_not_visible_adj_attempt_{slider_adjust_attempts}.png")
                    raise Exception("Slider text for week not visible during adjustment.")

                current_week_text = visible_slider_text_element.text.strip()
                cleaned_text = "".join(filter(str.isdigit, current_week_text))

                if not cleaned_text:
                    print(f"Warning: Could not parse digits from slider text '{current_week_text}'. Retrying...")
                    time.sleep(3)
                    slider_adjust_attempts += 1
                    continue

                current_week_value_read = int(cleaned_text)
                print(f"Current week on slider: {current_week_value_read}")

                if current_week_value_read == TARGET_START_WEEK:
                    print(f"Successfully reached target start week {TARGET_START_WEEK}.")
                    break
                elif current_week_value_read < TARGET_START_WEEK:
                    print(f"Current week {current_week_value_read} < {TARGET_START_WEEK}. Clicking increment...")
                    action_button = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable(INCREMENT_BUTTON_LOCATOR)
                    )
                    action_button.click()
                elif current_week_value_read > TARGET_START_WEEK:
                    print(f"Current week {current_week_value_read} > {TARGET_START_WEEK}. Clicking decrement...")
                    action_button = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable(DECREMENT_BUTTON_LOCATOR)
                    )
                    action_button.click()

                # Wait for the text to change to the new expected value or at least to be different from the old one.
                # This is more robust than a fixed sleep.
                expected_week_after_click = current_week_value_read + 1 if current_week_value_read < TARGET_START_WEEK else current_week_value_read -1
                if current_week_value_read == TARGET_START_WEEK -1 and current_week_value_read < TARGET_START_WEEK : # about to hit target
                    expected_week_after_click = TARGET_START_WEEK
                elif current_week_value_read == TARGET_START_WEEK +1 and current_week_value_read > TARGET_START_WEEK: # about to hit target
                    expected_week_after_click = TARGET_START_WEEK

                print(f"Waiting for slider text to update from {current_week_value_read} (expected: {expected_week_after_click})...")
                WebDriverWait(driver, 15).until( # Wait up to 15s for text to change
                    EC.text_to_be_present_in_element(SLIDER_TEXT_LOCATOR_WEEK, str(expected_week_after_click))
                )
                print(f"Slider text updated after click.")
                time.sleep(1) # Short additional pause after text confirms update

                slider_adjust_attempts += 1

            except TimeoutException as e_slider_timeout:
                print(f"Timeout during week adjustment (attempt {slider_adjust_attempts}): {e_slider_timeout}")
                driver.save_screenshot(f"err_timeout_slider_adj_wk{slider_adjust_attempts}.png")
                # If slider text doesn't update as expected, could be an issue.
                # Try to re-read and continue if possible, or raise if stuck.
                if slider_adjust_attempts > 5 : # Check if stuck after a few tries
                    try:
                        stuck_check_text = "".join(filter(str.isdigit, driver.find_element(*SLIDER_TEXT_LOCATOR_WEEK).text.strip()))
                        if stuck_check_text and int(stuck_check_text) == current_week_value_read:
                            print("Slider text seems stuck. Raising error.")
                            raise
                    except: pass # Ignore if re-read fails, main loop will retry or fail
                slider_adjust_attempts += 1 # Count as an attempt even on timeout if we retry
                if slider_adjust_attempts >= max_slider_adjust_attempts: raise
                print("Retrying slider adjustment step...")
                continue

            except Exception as e_slider_err:
                print(f"Error during week adjustment (attempt {slider_adjust_attempts}): {e_slider_err}")
                driver.save_screenshot(f"err_slider_adj_wk{slider_adjust_attempts}.png")
                if slider_adjust_attempts >= max_slider_adjust_attempts -1 : raise
                time.sleep(3)
                slider_adjust_attempts +=1

        if slider_adjust_attempts >= max_slider_adjust_attempts and current_week_value_read != TARGET_START_WEEK:
            print(f"Error: Max attempts ({max_slider_adjust_attempts}) to set week to {TARGET_START_WEEK}. Last read: {current_week_value_read}")
            raise Exception(f"Failed to set week to {TARGET_START_WEEK} using slider buttons.")

        print("Week adjustment process (using slider buttons) finished.")
        print("-" * 30)
        # --- END ADJUST EPI WEEK ---

        # --- Download Loop ---
        weeknum_for_loop_control = TARGET_START_WEEK # This controls the loop iterations

        print(f"--- Starting Download Loop for Year(s) '{year_label_for_filename}' from Week {weeknum_for_loop_control} ---")

        # Initial download for the starting week
        # Read the *actual* current week from slider for the first download
        try:
            print("Reading slider text for initial download...")
            initial_slider_text_elem = WebDriverWait(driver, 15).until(EC.visibility_of_element_located(SLIDER_TEXT_LOCATOR_WEEK))
            initial_observed_week_str = "".join(filter(str.isdigit, initial_slider_text_elem.text.strip()))
            if not initial_observed_week_str: raise ValueError("Cannot read initial week for download.")
            observed_week_for_download = int(initial_observed_week_str)
            print(f"Initial download for Observed Week Number: {observed_week_for_download}")
            if observed_week_for_download != TARGET_START_WEEK:
                print(f"WARNING: Slider shows {observed_week_for_download} but expected {TARGET_START_WEEK} for initial download.")
        except Exception as e_init_read:
            print(f"Error reading initial week for download: {e_init_read}. Defaulting to TARGET_START_WEEK ({TARGET_START_WEEK}) for filename.")
            observed_week_for_download = TARGET_START_WEEK # Fallback

        download_and_rename(wait, shadow_doc2_context, observed_week_for_download, default_download_dir_for_chrome, final_file_destination_path, driver, year_label_for_filename, today_timestamp)

        # Decrement and download for subsequent weeks
        current_loop_iteration_week = observed_week_for_download # Start decrementing from the actually observed week

        while current_loop_iteration_week > 1:
            print("-" * 20)
            target_decrement_week_expected = current_loop_iteration_week - 1
            print(f"Loop expects to decrement to Week Number: {target_decrement_week_expected}")

            try:
                decrement_button_actual_locator = DECREMENT_BUTTON_LOCATOR
                decrement_button = wait.until(EC.element_to_be_clickable(decrement_button_actual_locator))

                week_before_decrement_text_elem = WebDriverWait(driver, 15).until(EC.visibility_of_element_located(SLIDER_TEXT_LOCATOR_WEEK))
                week_before_decrement_str = "".join(filter(str.isdigit, week_before_decrement_text_elem.text.strip()))
                if not week_before_decrement_str: raise ValueError("Cannot read week before decrement.")
                # week_before_decrement = int(week_before_decrement_str) # Not strictly needed for logic if we wait for target
                print(f"Week before clicking decrement: {week_before_decrement_str}")

                decrement_button.click()
                print(f"Clicked decrement button. Waiting for slider to update to {target_decrement_week_expected}...")

                WebDriverWait(driver, 20).until(
                     EC.text_to_be_present_in_element(SLIDER_TEXT_LOCATOR_WEEK, str(target_decrement_week_expected))
                )
                print(f"Slider text confirmed updated to: {target_decrement_week_expected}")
                time.sleep(2)

                observed_week_for_download = target_decrement_week_expected

            except Exception as e_dec:
                print(f"Error during decrement or waiting for text update for week {target_decrement_week_expected}: {e_dec}")
                driver.save_screenshot(f"err_decrement_week_{target_decrement_week_expected}.png")
                try:
                    stuck_week_elem = driver.find_element(*SLIDER_TEXT_LOCATOR_WEEK)
                    stuck_week_val = stuck_week_elem.text.strip()
                    print(f"Slider seems stuck or did not update as expected. Currently shows: {stuck_week_val}")
                except:
                    print("Could not even re-read slider text after decrement error.")
                break

            current_loop_iteration_week = target_decrement_week_expected
            download_and_rename(wait, shadow_doc2_context, observed_week_for_download, default_download_dir_for_chrome, final_file_destination_path, driver, year_label_for_filename, today_timestamp)

        print(f"--- Finished Download Loop for Year(s) '{year_label_for_filename}' ---")

    except Exception as e:
        print(f"An critical error occurred in iterate_weekly: {type(e).__name__} - {e}")
        if driver:
            timestamp_err = time.strftime("%Y%m%d-%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"critical_error_main_process_{timestamp_err}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Screenshot saved: {screenshot_path}")
            except Exception as scr_e:
                print(f"Could not save screenshot: {scr_e}")
    finally:
        if driver:
            print("Closing WebDriver...")
            # Enhanced cleanup in finally block
            print(f"Attempting cleanup in temporary download directory: {default_download_dir_for_chrome}")
            # 1. Clean up the known raw Tableau download filename
            raw_tableau_file_path_to_check = os.path.join(default_download_dir_for_chrome, raw_tableau_download_filename)
            if os.path.exists(raw_tableau_file_path_to_check):
                try:
                    print(f"Attempting to clean up raw Tableau downloaded file: {raw_tableau_file_path_to_check}")
                    os.remove(raw_tableau_file_path_to_check)
                    print(f"Successfully cleaned up: {raw_tableau_file_path_to_check}")
                except Exception as e_cleanup_raw:
                    print(f"Warning: Could not clean up raw Tableau file '{raw_tableau_file_path_to_check}': {e_cleanup_raw}")

            # 2. Clean up any of our temporary contextual files
            # The pattern for these files is `*_TEMPRAW.csv`
            # We need to use glob to find them in the default_download_dir_for_chrome
            temp_files_to_clean = glob.glob(os.path.join(default_download_dir_for_chrome, temp_contextual_file_pattern))
            if temp_files_to_clean:
                print(f"Found temporary contextual files to clean: {temp_files_to_clean}")
                for temp_file in temp_files_to_clean:
                    try:
                        print(f"Attempting to clean up temporary contextual file: {temp_file}")
                        os.remove(temp_file)
                        print(f"Successfully cleaned up: {temp_file}")
                    except Exception as e_cleanup_temp:
                        print(f"Warning: Could not clean up temp contextual file '{temp_file}': {e_cleanup_temp}")
            else:
                print("No temporary contextual files found for cleanup.")

            driver.quit()
        print("Script finished.")

if __name__ == "__main__":
    iterate_weekly()
