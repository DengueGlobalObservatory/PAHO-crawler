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

def move_to_download_folder(default_chrome_download_dir, final_destination_path, new_file_name_stem, file_extension):
    """
    Waits for a new file in default_chrome_download_dir, then moves and renames it.
    """
    got_file = False
    max_wait_seconds = 300  # Wait up to 5 minutes for the download
    check_interval_seconds = 5
    time_waited = 0
    downloaded_file_path = None

    print(f"Watching for download in: {default_chrome_download_dir}")
    initial_files = set(os.listdir(default_chrome_download_dir))

    while not got_file and time_waited < max_wait_seconds:
        time.sleep(check_interval_seconds)
        time_waited += check_interval_seconds
        current_files = set(os.listdir(default_chrome_download_dir))
        new_files = current_files - initial_files

        if new_files:
            # Find the most recently created/modified file among the new ones
            # This handles cases where multiple files might appear if script is run quickly
            latest_file = ""
            latest_time = 0
            for f_name in new_files:
                f_path = os.path.join(default_chrome_download_dir, f_name)
                # Ignore temporary download files
                if f_name.endswith(('.crdownload', '.tmp', '.part')):
                    print(f"Download in progress (temp file: {f_name}). Continuing to wait...")
                    # Reset new_files so we don't process this temp file yet
                    new_files = set()
                    break # Break from inner loop to re-check files after sleep

                # Check if it's a regular file and get its modification time
                if os.path.isfile(f_path):
                    mod_time = os.path.getmtime(f_path)
                    if mod_time > latest_time:
                        latest_time = mod_time
                        latest_file = f_path

            if latest_file and not latest_file.endswith(('.crdownload', '.tmp', '.part')):
                downloaded_file_path = latest_file
                print(f"Detected new file: {downloaded_file_path}")

                # Small delay to ensure file writing is complete
                time.sleep(2)

                # Verify file size is not zero (optional, but good check)
                if os.path.getsize(downloaded_file_path) > 0:
                    got_file = True
                else:
                    print(f"File {downloaded_file_path} is empty. Waiting a bit longer...")
                    # Reset downloaded_file_path so we don't try to move an empty file yet
                    downloaded_file_path = None
                    initial_files.add(os.path.basename(latest_file)) # Add it back to initial if empty for now
                    got_file = False # Ensure we keep waiting
                    new_files = set() # Clear new_files
            elif not new_files: # This means the only new file was a temp file
                pass # Handled by the inner break, will continue outer loop
            else: # No valid new file found yet
                 print(f"Still waiting for download... ({time_waited}/{max_wait_seconds}s)")

        else: # No new files detected
            print(f"Still waiting for download... ({time_waited}/{max_wait_seconds}s)")

    if not got_file or not downloaded_file_path:
        print(f"Error: File download timed out or file not found after {max_wait_seconds} seconds.")
        return False

    # Create new file name
    final_file_destination = os.path.join(final_destination_path, new_file_name_stem + file_extension)
    os.makedirs(os.path.dirname(final_file_destination), exist_ok=True) # Ensure destination directory exists

    try:
        os.rename(downloaded_file_path, final_file_destination)
        print(f"Moved file to {final_file_destination}")
        return True
    except Exception as e:
        print(f"Error moving file from {downloaded_file_path} to {final_file_destination}: {e}")
        return False

def download_for_current_week(wait, driver, shadow_doc2_context, default_chrome_download_dir, final_download_path, year_identifier_for_filename, current_datetime_str):
    """
    Performs the download actions for the currently displayed week.
    Reads the week number from the page for the filename.
    """
    try:
        # Get the actual week number displayed on the page
        weeknum_div = wait.until(
            EC.visibility_of_element_located((By.CLASS_NAME, "sliderText"))
        )
        week_text = weeknum_div.text
        # Extract number, assuming format like "53" or "Week 53"
        match = re.search(r'\d+', week_text)
        if not match:
            print(f"Error: Could not parse week number from text: '{week_text}'")
            return False
        actual_week_on_page = int(match.group())
        print(f"Confirmed Week Number on page for download: {actual_week_on_page}")

        download_button = wait.until(
            EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
        )
        download_button.click()
        time.sleep(3) # Allow download modal to open

        crosstab_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
        )
        crosstab_button.click()
        time.sleep(3) # Allow crosstab options to load

        # The CSV and Export buttons are within the driver's current context (shadow_doc2_context)
        # Using explicit waits for these elements within the modal
        csv_radio_button_selector = "div[role='dialog'] input[type='radio'][value='csv']"
        csv_div = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, csv_radio_button_selector))
        )
        # Scroll into view if necessary, then click
        driver.execute_script("arguments[0].scrollIntoView(true);", csv_div)
        driver.execute_script("arguments[0].click();", csv_div) # JS click can be more reliable for custom controls
        # csv_div.click() # Standard click
        time.sleep(1)

        export_button_selector = "div[role='dialog'] button[data-tb-test-id='export-crosstab-export-Button']"
        export_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector))
        )
        export_button.click()
        print(f"Downloading CSV for Year(s) {year_identifier_for_filename}, Week {actual_week_on_page}...")
        time.sleep(5) # Give a moment for download to initiate

        new_file_name_stem = f"PAHO_{year_identifier_for_filename}_W{actual_week_on_page}_{current_datetime_str}"
        file_extension = '.csv'

        if not move_to_download_folder(default_chrome_download_dir, final_download_path, new_file_name_stem, file_extension):
            print(f"Failed to complete download and move for Week {actual_week_on_page}.")
            return False
        return True

    except Exception as e:
        print(f"An error occurred during download_for_current_week: {e}")
        import traceback
        traceback.print_exc()
        return False

def set_week_on_dashboard(wait, driver, target_week):
    """
    Adjusts the week slider on the dashboard to the target_week.
    Returns True if successful, False otherwise.
    """
    max_attempts = 70  # Max clicks to prevent infinite loops (e.g., from week 1 to 53 or vice versa)
    for attempt in range(max_attempts):
        try:
            weeknum_div = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "sliderText")))
            current_week_text = weeknum_div.text
            match = re.search(r'\d+', current_week_text)
            if not match:
                print(f"Error: Could not parse current week from '{current_week_text}' during set_week attempt.")
                return False # Cannot determine current state
            current_week_on_page = int(match.group())

            print(f"Attempt {attempt+1}: Current week on page: {current_week_on_page}, Target: {target_week}")

            if current_week_on_page == target_week:
                print(f"Successfully set week to {target_week}.")
                return True

            action_taken = False
            if current_week_on_page < target_week:
                # Need to increment. Assumes an increment button exists.
                # IF NO INCREMENT BUTTON: This logic path needs to be rethought.
                # For now, assuming 'tableauArrowInc' is a placeholder for the actual incrementor.
                increment_button_xpath = "//*[contains(@class, 'tableauArrowInc')]" # YOU MUST VERIFY THIS XPATH
                try:
                    increment_button = wait.until(EC.element_to_be_clickable((By.XPATH, increment_button_xpath)))
                    increment_button.click()
                    action_taken = True
                    print("Clicked increment.")
                except:
                    print("Warning: Increment button not found or clickable. Cannot increase week.")
                    # If only decrement is available, and we are below target, we cannot reach it.
                    return False # Cannot reach target if only decrement exists and current < target

            elif current_week_on_page > target_week:
                decrement_button_xpath = "//*[contains(@class, 'tableauArrowDec')]"
                decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, decrement_button_xpath)))
                decrement_button.click()
                action_taken = True
                print("Clicked decrement.")

            if action_taken:
                time.sleep(3) # Wait for the week display to update
            else: # Should not happen if current_week != target_week and buttons exist
                print("No action taken in set_week logic, current matches target or button issue.")
                return current_week_on_page == target_week


        except Exception as e:
            print(f"Error during set_week_on_dashboard (attempt {attempt+1}): {e}")
            time.sleep(2) # Wait a bit before retrying

    print(f"Failed to set week to {target_week} after {max_attempts} attempts.")
    return False

def iterate_weekly():
    # --- Configuration ---
    # Define the three years you want to select
    target_years_to_select = ["2025", "2024", "2023"] # <<< ADJUST AS NEEDED
    # Create a string for filenames, e.g., "2022-2024"
    year_identifier_for_filename = f"{min(target_years_to_select)}-{max(target_years_to_select)}"

    current_datetime_str = datetime.now().strftime('%Y%m%d%H%M')

    # Use a consistent directory for downloads by Chrome AND for your script to find files
    # Fallback to user's Downloads if specific path is not writable/found
    # preferred_download_dir = 'C:/Users/AhyoungLim/Downloads' # Your original
    preferred_download_dir = os.path.expanduser('~/Downloads') # More portable default

    # Check if preferred_download_dir is writable, else use a temp dir in CWD
    if os.path.exists(preferred_download_dir) and os.access(preferred_download_dir, os.W_OK):
        default_chrome_download_dir = preferred_download_dir
    else:
        default_chrome_download_dir = os.path.join(os.getcwd(), "chrome_downloads")
        os.makedirs(default_chrome_download_dir, exist_ok=True)
        print(f"Warning: Preferred download directory '{preferred_download_dir}' not accessible. Using '{default_chrome_download_dir}' instead.")

    print(f"Chrome will download to: {default_chrome_download_dir}")

    # Final destination for processed files
    # github_workspace_base = 'C:/Users/AhyoungLim/Dropbox/WORK/OpenDengue/PAHO-crawler/data' # Your original
    github_workspace_base = os.path.join(os.getcwd(), 'data') # Example relative path

    today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
    final_download_path_for_files = os.path.join(github_workspace_base, today_directory_name)
    os.makedirs(final_download_path_for_files, exist_ok=True)
    print(f"Organized files will be saved in: {final_download_path_for_files}")

    chrome_options = uc.ChromeOptions()
    prefs = {
        "download.default_directory": default_chrome_download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safeBrowse.enabled": True # Good practice
    }
    chrome_options.add_experimental_option("prefs", prefs)
    # chrome_options.add_argument('--headless') # Optional: for running without UI
    # chrome_options.add_argument('--no-sandbox') # Optional: if running as root/in Docker
    # chrome_options.add_argument('--disable-dev-shm-usage') # Optional: for some Linux environments

    driver = None
    try:
        print(f"Initializing Chrome driver (detected version: {chrome_version if chrome_version else 'auto'})...")
        driver = uc.Chrome(headless=False, use_subprocess=True, options=chrome_options, version_main=chrome_version)
        driver.maximize_window() # Helps with element visibility

        main_page_url = 'https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html'
        driver.get(main_page_url)
        wait = WebDriverWait(driver, 40) # Increased wait time for robustness

        # --- Navigate to Tableau iFrame ---
        iframe_src_outer = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
        iframe_locator_outer = (By.XPATH, f"//div[contains(@class, 'vizTab')]//iframe[@src='{iframe_src_outer}']")
        print("Waiting for outer iframe...")
        iframe_outer_element = wait.until(EC.presence_of_element_located(iframe_locator_outer))
        driver.switch_to.frame(iframe_outer_element)
        print("Switched to outer iframe.")

        # This is the document context of the outer iframe, not a shadow DOM
        # shadow_doc_outer = driver.execute_script('return document')

        iframe_locator_inner = (By.XPATH, "//body/iframe") # Assuming it's the direct child iframe in body
        # Alternative if it has an ID or specific attribute: (By.ID, "tableauViz")
        print("Waiting for inner Tableau iframe...")
        iframe_inner_element = wait.until(EC.presence_of_element_located(iframe_locator_inner))
        driver.switch_to.frame(iframe_inner_element)
        print("Switched to inner Tableau iframe.")

        # This is the document context of the Tableau viz iframe
        # shadow_doc2_tableau_viz = driver.execute_script('return document')

        iframe_page_title = driver.title
        print(f"Title of inner Tableau iframe: {iframe_page_title}")
        if "PAHO/WHO Data" not in iframe_page_title and "Tableau" not in iframe_page_title : # Broader check
            print(f"Warning: Unexpected iframe title: {iframe_page_title}. This might indicate a problem.")
            # Consider quitting if title is critical:
            # driver.quit()
            # return

        time.sleep(5) # Allow Tableau viz to fully render elements

        # --- Year Selection ---
        print(f"Attempting to select years: {', '.join(target_years_to_select)}")
        year_tab_filter_area = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId13')))

        dropdown_button_in_year_tab = year_tab_filter_area.find_element(By.CSS_SELECTOR, 'span.tabComboBoxButton')
        wait.until(EC.element_to_be_clickable(dropdown_button_in_year_tab)).click()
        print("Opened year selection dropdown.")
        time.sleep(2) # Wait for dropdown items to become visible/interactive

        # Logic to select multiple years:
        # Assumes clicking a year checkbox toggles its state.
        # If selecting one year deselects others, this logic needs to change (e.g. single year processing loop)

        # First, potentially clear existing selections if necessary (Tableau specific)
        # This is complex as "clear" might not exist or might behave differently.
        # A common pattern is to click an "(All)" to deselect if it's checked, then select specifics.
        # For now, we assume clicking the desired years will achieve the target state.

        successfully_selected_years = 0
        for year_str in target_years_to_select:
            try:
                # XPath to find the input checkbox preceding the link with the year text
                year_checkbox_xpath = f'//div[contains(@class, "facetOverflow")]//a[text()="{year_str}"]/preceding-sibling::input[@type="checkbox"]'
                year_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, year_checkbox_xpath)))

                # Click to select. If it's already selected, this might deselect it.
                # For robust selection, you might check `is_selected()` first,
                # but this can be unreliable with custom JS controls.
                # A direct click is often what's needed.
                driver.execute_script("arguments[0].click();", year_checkbox) # JS click for reliability
                print(f"Clicked checkbox for year {year_str}.")
                successfully_selected_years +=1
                time.sleep(0.5) # Brief pause between clicks
            except Exception as e:
                print(f"Could not find or click year {year_str} checkbox: {e}")

        if successfully_selected_years != len(target_years_to_select):
            print(f"Warning: Attempted to select {len(target_years_to_select)} years, but only interacted with {successfully_selected_years} checkboxes.")
        else:
            print(f"Successfully interacted with checkboxes for years: {', '.join(target_years_to_select)}.")

        # Close the year dropdown (e.g., by clicking the "glass pane" or an "Apply" button if one exists)
        try:
            # Try clicking the overlay/glass pane first
            glass_pane_to_close_dropdown = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
            glass_pane_to_close_dropdown.click()
            print("Closed year dropdown by clicking glass pane.")
        except:
            # Fallback: Try clicking the dropdown button again to toggle it closed
            print("Could not close dropdown via glass pane, trying to click dropdown button again.")
            try:
                dropdown_button_in_year_tab.click()
                print("Closed year dropdown by clicking dropdown button again.")
            except Exception as e_close:
                print(f"Failed to close year dropdown: {e_close}. Proceeding cautiously.")

        time.sleep(5) # Wait for data to update after year selection

        # --- Week Setting to Target (e.g., 53) ---
        target_start_week = 53
        print(f"Attempting to set dashboard week to {target_start_week}...")
        if not set_week_on_dashboard(wait, driver, target_start_week):
            print(f"Critical Error: Failed to set dashboard week to {target_start_week}. Aborting script.")
            if driver: driver.quit()
            return
        print(f"Dashboard week successfully set to {target_start_week}.")

        # --- Loop for Downloading Files (from target_start_week down to 1) ---
        # The week is now confirmed to be at target_start_week.
        # The loop will perform one download, then decrement, then download, etc.

        current_processing_week = target_start_week
        while current_processing_week >= 1:
            print(f"\n--- Processing week: {current_processing_week} ---")

            # Download for the current week
            # Pass `driver` as shadow_doc2_context because modals are in this frame
            if not download_for_current_week(wait, driver, driver, default_chrome_download_dir, final_download_path_for_files, year_identifier_for_filename, current_datetime_str):
                print(f"Failed to download data for week {current_processing_week}. Skipping this week.")
                # Decide on error: continue to next week, or break? For now, continue.

            if current_processing_week == 1:
                print("Reached week 1. Download process complete.")
                break # Exit loop

            # Decrement week on the dashboard for the next iteration
            print(f"Attempting to decrement week on dashboard from {current_processing_week}...")
            try:
                decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'tableauArrowDec')]")))
                decrement_button.click()
                time.sleep(4) # Crucial: Wait for week to update on page AND for data to potentially reload

                # Verify week changed (optional but good for debugging)
                weeknum_div_after_dec = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "sliderText")))
                week_text_after_dec = weeknum_div_after_dec.text
                match_after_dec = re.search(r'\d+', week_text_after_dec)
                if match_after_dec:
                    actual_week_after_dec = int(match_after_dec.group())
                    print(f"Dashboard week is now: {actual_week_after_dec} (expected around {current_processing_week-1})")
                    if actual_week_after_dec >= current_processing_week : # Check if decrement failed
                         print(f"Warning: Decrement button clicked, but week did not decrease as expected (is {actual_week_after_dec}, was {current_processing_week}).")
                         # Potentially add logic here to break or retry if decrement fails repeatedly.
                else:
                    print("Warning: Could not read week number after decrementing.")


            except Exception as e_dec:
                print(f"Error clicking decrement button or reading week after decrement for week {current_processing_week-1}: {e_dec}")
                print("Stopping further downloads due to decrement failure.")
                break

            current_processing_week -= 1 # Decrement our loop counter

        print("All targeted weeks processed.")

    except Exception as e:
        print(f"An unhandled error occurred in the main iterate_weekly function: {e}")
        import traceback
        traceback.print_exc()
        if driver:
            try:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                screenshot_path = os.path.join(os.getcwd(), f"error_screenshot_{timestamp}.png")
                driver.save_screenshot(screenshot_path)
                print(f"Screenshot saved to {screenshot_path}")
            except Exception as se:
                print(f"Failed to save screenshot: {se}")

    finally:
        if driver:
            print("Closing the browser.")
            driver.quit()

if __name__ == '__main__':
    iterate_weekly()
