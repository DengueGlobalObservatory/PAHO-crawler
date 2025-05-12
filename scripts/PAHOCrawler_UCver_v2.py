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
    Waits for a download to complete in default_dir_path, then moves the latest file
    to final_download_path with the new name.
    """
    got_file = False
    max_wait_time = 180  # Max time to wait for download (3 minutes)
    check_interval = 5   # How often to check for the file
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
                time.sleep(check_interval - 3)

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
        shutil.move(downloaded_file_path, final_destination_path)
        print(f"Successfully moved file to '{final_destination_path}'")
    except Exception as e:
        print(f"Error moving file from '{downloaded_file_path}' to '{final_destination_path}': {e}")
        raise

def download_and_rename(wait, shadow_doc2_context, weeknum_for_file, default_dir_path, final_download_path, driver_instance, year_label_for_filename, today_timestamp_str):
    """Downloads and renames the file for the given week number."""
    print("-" * 10 + f" Starting download for Year(s) '{year_label_for_filename}', Week {weeknum_for_file} " + "-" * 10)

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
        print(f"Clicked Export/Download for Week {weeknum_for_file}. Waiting for file...")

        # Construct filename
        new_file_name_stem = f"PAHO_Y{year_label_for_filename}_W{weeknum_for_file:02d}_{today_timestamp_str}"
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

        # --- Select Multiple Target Years (2023, 2024, 2025) ---
        print(f"Attempting to select years: {', '.join(target_years_to_select)}...")
        year_tab_id = 'tabZoneId13' # User's original ID
        year_tab = wait.until(EC.visibility_of_element_located((By.ID, year_tab_id)))

        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton') # User's original selector
        dd_open_button = year_tab.find_element(*dd_locator)
        dd_open_button.click()
        print("Clicked year dropdown open button.")
        time.sleep(3) # ** INCREASED PAUSE for dropdown to render **
        print("Taking screenshot: year_dropdown_opened.png")
        driver.save_screenshot("year_dropdown_opened.png")


        # Find all year options (input checkboxes and their corresponding 'a' tag for text)
        all_year_inputs_xpath = '//div[contains(@class, "facetOverflow")]//div[contains(@class, "valueSection")]//input[@type="checkbox"]'
        all_year_labels_xpath = '//div[contains(@class, "facetOverflow")]//div[contains(@class, "valueSection")]//a'

        print(f"Waiting for year options to be present using XPath: {all_year_inputs_xpath}")
        # ** INCREASED TIMEOUT for year options **
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, all_year_inputs_xpath)))
        print("Year options are present. Fetching elements...")
        time.sleep(1.5) # Extra pause for all items to fully render after presence

        year_input_elements = shadow_doc2_context.find_elements(By.XPATH, all_year_inputs_xpath)
        year_label_elements = shadow_doc2_context.find_elements(By.XPATH, all_year_labels_xpath)

        if not year_input_elements or len(year_input_elements) != len(year_label_elements):
            print(f"Error: Could not find year options or mismatch in inputs ({len(year_input_elements)}) and labels ({len(year_label_elements)}).")
            driver.save_screenshot("err_year_options_mismatch.png")
            # Decide if to raise error or continue if some years might be selectable
            # For now, let's try to proceed if any elements were found, but log a clear warning.
            if not year_input_elements:
                 raise Exception("No year input elements found in the dropdown.")
        else:
            print(f"Found {len(year_input_elements)} year options in dropdown.")
            for i in range(len(year_input_elements)):
                year_input = year_input_elements[i]
                # Handle potential StaleElementReferenceException if DOM changes during iteration
                try:
                    year_text_element = year_label_elements[i]
                    year_text = year_text_element.text.strip()
                    if not year_text: # If 'a' tag has no direct text, try to get it from a child span or title
                        year_text = year_text_element.get_attribute("title") or \
                                    driver.execute_script("return arguments[0].innerText;", year_text_element)
                        year_text = year_text.strip() if year_text else "UNKNOWN_YEAR"


                    is_selected_by_script = year_input.is_selected()

                    if year_text in target_years_to_select:
                        if not is_selected_by_script:
                            print(f"Selecting year: {year_text}")
                            driver.execute_script("arguments[0].click();", year_input)
                            time.sleep(0.7) # Pause between clicks
                        else:
                            print(f"Year {year_text} is already selected.")
                    else:
                        if is_selected_by_script:
                            print(f"Deselecting year: {year_text}")
                            driver.execute_script("arguments[0].click();", year_input)
                            time.sleep(0.7)
                except Exception as e_year_item:
                    print(f"Warning: Error processing year item {i}: {e_year_item}. Skipping this item.")
                    driver.save_screenshot(f"err_processing_year_item_{i}.png")
                    continue # Continue to the next year item

            print("Finished processing year selections.")

        print("Closing year dropdown...")
        dd_close_locator = (By.CLASS_NAME, "tab-glass") # User's original
        dd_close_button = wait.until(EC.element_to_be_clickable(dd_close_locator))
        dd_close_button.click()
        print("Closed year dropdown.")
        time.sleep(5) # Allow filters to apply

        # --- ADJUST EPI WEEK TO STARTING WEEK (53) USING SLIDER BUTTONS ---
        print("-" * 30)
        print("Adjusting Epidemiological Week to 53 using slider buttons...")
        TARGET_START_WEEK = 53
        SLIDER_TEXT_LOCATOR_WEEK = (By.CSS_SELECTOR, ".sliderText")
        INCREMENT_BUTTON_LOCATOR = (By.XPATH, "//*[contains(@class, 'tableauArrowInc') or contains(@class, 'dijitSliderIncrementIconH')]")
        DECREMENT_BUTTON_LOCATOR = (By.XPATH, "//*[contains(@class, 'tableauArrowDec') or contains(@class, 'dijitSliderDecrementIconH')]")

        max_slider_adjust_attempts = 70
        slider_adjust_attempts = 0
        current_week_value_read = -1 # Initialize

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

                time.sleep(3)
                slider_adjust_attempts += 1

            except TimeoutException as e_slider_timeout:
                print(f"Timeout during week adjustment (attempt {slider_adjust_attempts}): {e_slider_timeout}")
                driver.save_screenshot(f"err_timeout_slider_adj_wk{slider_adjust_attempts}.png")
                raise
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
        weeknum_for_loop = TARGET_START_WEEK
        print(f"--- Starting Download Loop for Year(s) '{year_label_for_filename}' from Week {weeknum_for_loop} ---")

        # Initial download for the starting week
        print(f"Initial download for Week Number: {weeknum_for_loop}")
        download_and_rename(wait, shadow_doc2_context, weeknum_for_loop, default_download_dir_for_chrome, final_file_destination_path, driver, year_label_for_filename, today_timestamp)

        # Decrement and download for subsequent weeks
        while weeknum_for_loop > 1:
            print("-" * 20)
            target_decrement_week = weeknum_for_loop - 1
            print(f"Attempting to decrement to Week Number: {target_decrement_week}")
            try:
                decrement_button_actual_locator = DECREMENT_BUTTON_LOCATOR
                decrement_button = wait.until(EC.element_to_be_clickable(decrement_button_actual_locator))
                decrement_button.click()
                print(f"Clicked decrement button. Aiming for week {target_decrement_week}.")
                time.sleep(6)
            except Exception as e_dec:
                print(f"Error clicking decrement button for week {target_decrement_week}: {e_dec}")
                driver.save_screenshot(f"err_decrement_week_{target_decrement_week}.png")
                break

            weeknum_for_loop = target_decrement_week
            download_and_rename(wait, shadow_doc2_context, weeknum_for_loop, default_download_dir_for_chrome, final_file_destination_path, driver, year_label_for_filename, today_timestamp)

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
            driver.quit()
        print("Script finished.")

if __name__ == "__main__":
    iterate_weekly()
