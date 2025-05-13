import undetected_chromedriver as uc
#from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#from selenium.webdriver.common.action_chains import ActionChains # Not used in original snippet
#from selenium.webdriver.common.keys import Keys # Not used in original snippet

#from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome.service import Service
import time
import os
from datetime import datetime
import subprocess
import re # Ensure re is imported for regex operations
import sys

#from webdriver_manager.chrome import ChromeDriverManager
#driver_executable_path = ChromeDriverManager().install()
#https://github.com/ultrafunkamsterdam/undetected-chromedriver/issues/1904

# Function to get the installed Chrome version (No changes from your original)
def get_chrome_version():
    try:
        if sys.platform == "win32":
            # Command to retrieve Chrome version from Windows registry
            output = subprocess.check_output(
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                shell=True,
                text=True,
                stderr=subprocess.DEVNULL # Suppress stderr
            )
            version = re.search(r'\s+version\s+REG_SZ\s+(\d+)\.', output)
        else:  # Assuming Linux or other Unix-like systems
            version = None
            # Try different commands to retrieve Chrome version on Linux
            for command in ['google-chrome --version', 'google-chrome-stable --version', 'chromium-browser --version']:
                try:
                    output = subprocess.check_output(command, shell=True, text=True, stderr=subprocess.DEVNULL)
                    version = re.search(r'\b(\d+)\.', output)
                    if version:
                        break
                except (subprocess.CalledProcessError, FileNotFoundError):
                    continue
            if not version: # Moved this check outside the loop
                 # raise RuntimeError("Could not determine Chrome version") # Soften this for UC
                 print("Warning: Could not determine Chrome version via common commands. UC will attempt to use a compatible driver.")
                 return None


        if version:
            return int(version.group(1))
        else:
            # raise ValueError("Could not parse Chrome version") # Soften this for UC
            print("Warning: Could not parse Chrome version. UC will attempt to use a compatible driver.")
            return None
    except Exception as e:
        # raise RuntimeError("Failed to get Chrome version") from e # Soften this for UC
        print(f"Warning: Failed to get Chrome version due to: {e}. UC will attempt to use a compatible driver.")
        return None

# Get the major version of Chrome installed
chrome_version = get_chrome_version()

# move_to_download_folder (No changes from your original)
def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    time_waited = 0
    max_wait_interval = 300 # Wait up to 5 minutes
    check_interval = 10 # Check every 10 seconds

    while not got_file and time_waited < max_wait_interval:
        try:
            # Use glob to get the current file name
            # Ensure default_dir is not empty before calling max
            list_of_files = [os.path.join(default_dir, f) for f in os.listdir(default_dir) if os.path.isfile(os.path.join(default_dir, f))]
            if not list_of_files:
                print(f"Download directory '{default_dir}' is empty. Waiting...")
                time.sleep(check_interval)
                time_waited += check_interval
                continue

            currentFile = max(list_of_files, key=os.path.getctime)

            # Ensure the file exists and is not a temporary download file before proceeding
            if os.path.exists(currentFile) and not currentFile.endswith(('.crdownload', '.tmp', '.part')):
                 # Optional: check if file size is greater than 0 and stable
                time.sleep(2) # Short pause to ensure writing is finished
                if os.path.getsize(currentFile) > 0:
                    got_file = True
                else:
                    print(f"File {currentFile} found but is empty. Waiting...")
                    # To avoid repeatedly picking the same empty file,
                    # we might need more sophisticated logic or hope it gets written to.
                    # For now, just wait.
                    time.sleep(check_interval)
                    time_waited += check_interval
            else:
                if not os.path.exists(currentFile):
                    # This case should be rare if list_of_files was populated
                    print(f"File {currentFile} not found after listing. Retrying...")
                else: # It's a temporary file
                    print(f"Download in progress (temp file: {os.path.basename(currentFile)}). Waiting...")
                time.sleep(check_interval)
                time_waited += check_interval

        except FileNotFoundError: # If default_dir itself doesn't exist
            print(f"Error: Download directory '{default_dir}' not found. Please check the path.")
            return # Exit if download directory is invalid
        except Exception as e:
            print(f"File has not finished downloading or other error: {e}. Waiting...")
            time.sleep(check_interval)
            time_waited += check_interval

    if not got_file:
        print(f"Failed to get downloaded file from '{default_dir}' after {max_wait_interval} seconds.")
        return

    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)
    os.makedirs(os.path.dirname(fileDestination), exist_ok=True) # Ensure destination directory exists

    # Move the file
    try:
        os.rename(currentFile, fileDestination)
        print(f"Moved file to {fileDestination}")
    except Exception as e:
        print(f"Error moving file {currentFile} to {fileDestination}: {e}")


# download_and_rename (No changes from your original, uses weeknum from page for filename)
def download_and_rename(wait, shadow_doc2, weeknum_loop_counter, default_dir, downloadPath, driver, year_for_filename, today):
    """Download and rename the file. weeknum_loop_counter is for loop control, actual week for filename is read from page."""

    # Wait for the week number to update on the page and use this for the filename
    weeknum_div_page = wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
    )
    actual_week_on_page = int(re.search(r'\d+', weeknum_div_page.text).group()) # Extract number
    print(f"Data on page corresponds to Week: {actual_week_on_page}")

    # Find and click the download button at the bottom of the dashboard
    download_button = wait.until(
        EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
    )
    download_button.click()

    time.sleep(5) # Original sleep

    # Find and click the crosstab button (in a pop up window)
    crosstab_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
    )
    crosstab_button.click()

    time.sleep(5) # Original sleep

    # Find and select the CSV option
    # Using shadow_doc2 as in your original script for elements within this modal
    csv_div = shadow_doc2.find_element(By.CSS_SELECTOR, "input[type='radio'][value='csv']")
    driver.execute_script("arguments[0].scrollIntoView(true);", csv_div) # Ensure visibility
    driver.execute_script("arguments[0].click();", csv_div) # Use JS click for reliability
    time.sleep(5) # Original sleep

    # Find and click the export button
    export_button = shadow_doc2.find_element(By.CSS_SELECTOR, '[data-tb-test-id="export-crosstab-export-Button"]')
    export_button.click()
    print("Downloading CSV file...") # Changed print message slightly
    time.sleep(5) # Original sleep

    # Use the move_to_download_folder function to move the downloaded file
    # Filename uses actual_week_on_page
    newFileName = f"PAHO_{year_for_filename}_W{actual_week_on_page}_{today}"
    fileExtension = '.csv'

    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def iterate_weekly():
    # --- MODIFICATION: Define the three years you want to select ---
    # year = "2023" # Original: choose year to download
    target_selected_years = ["2025", "2024", "2023"] # <<< ADJUST THESE THREE YEARS AS NEEDED
    # Create a string for filenames, e.g., "2024_2023_2022"
    year_for_filename = "_".join(sorted(target_selected_years, reverse=True))


    today = datetime.now().strftime('%Y%m%d%H%M') # current date and time

    # set directory (No changes from your original paths)
    github_workspace = 'C:/Users/AhyoungLim/Dropbox/WORK/OpenDengue/PAHO-crawler/data'
    default_dir = 'C:/Users/AhyoungLim/Downloads'

    today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
    downloadPath = os.path.join(github_workspace, today_directory_name)
    os.makedirs(downloadPath, exist_ok=True) # create a new directory

    # set chrome download directory
    chrome_options = uc.ChromeOptions()
    # --- Ensure consistent download directory for Chrome and move_to_download_folder ---
    prefs = {"download.default_directory": default_dir} # Use the same default_dir
    chrome_options.add_experimental_option("prefs", prefs)

    driver = None # Initialize driver to None for finally block
    try:
        # using undetected-chromedriver
        print(f"Initializing Chrome driver (UC version: {chrome_version if chrome_version else 'auto'})...")
        driver = uc.Chrome(headless=False, use_subprocess=False, options = chrome_options, version_main=chrome_version if chrome_version else None)
        driver.maximize_window()
        driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')

        wait = WebDriverWait(driver, 30) # Increased default wait time slightly

        # First iframe
        iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
        iframe_locator = (By.XPATH, "//div[contains(@class, 'vizTab')]//iframe[@src='" + iframe_src + "']")
        print("Waiting for outer iframe...")
        iframe = wait.until(EC.presence_of_element_located(iframe_locator))
        driver.switch_to.frame(iframe)
        print("Switched to outer iframe.")

        # Grab the shadow element (document of the current iframe)
        # shadow = driver.execute_script('return document') # Not used directly for iframe2 finding

        # Get the iframe inside shadow element of first iframe
        print("Waiting for inner Tableau iframe...")
        iframe2 = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe"))) # Assuming it's the only/first iframe in body
        driver.switch_to.frame(iframe2)
        shadow_doc2 = driver.execute_script('return document') # This is the document of the Tableau iframe
        print("Switched to inner Tableau iframe.")

        iframe_page_title = driver.title
        print(f"Inner iframe title: {iframe_page_title}")

        if "PAHO/WHO Data" not in iframe_page_title and "Tableau" not in iframe_page_title:
            print("Warning: Unexpected iframe title. May not be on the correct page.")
            # driver.quit() # Decide if to quit
            # return

        time.sleep(3) # Original sleep

        # find the year tab
        year_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId13')))

        # find the dropdown button within the year tab
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = year_tab.find_element(*dd_locator)
        dd_open.click()
        print("Opened year dropdown.")
        time.sleep(1) # Wait for dropdown to populate

        # --- MODIFICATION: Select the defined target years ---
        print(f"Attempting to select years: {', '.join(target_selected_years)}")
        selected_count = 0
        for year_to_select in target_selected_years:
            try:
                year_xpath = f'//div[contains(@class, "facetOverflow")]/a[text()="{year_to_select}"]/preceding-sibling::input'
                # Using shadow_doc2.find_element as in your original script
                year_checkbox = shadow_doc2.find_element(By.XPATH, year_xpath)
                # Click the checkbox - assuming it toggles selection
                # For more robustness, one might check if it's already selected,
                # but for simplicity and to match original, direct click.
                year_checkbox.click()
                print(f"Clicked year: {year_to_select}")
                selected_count += 1
                time.sleep(0.5) # Small pause between clicks
            except Exception as e:
                print(f"Warning: Could not find or click year '{year_to_select}'. Error: {e}")

        if selected_count == len(target_selected_years):
            print(f"Successfully interacted with checkboxes for all {len(target_selected_years)} target years.")
        else:
            print(f"Warning: Interacted with {selected_count} of {len(target_selected_years)} target year checkboxes.")

        # close the dropdown menu
        dd_close = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass"))
        )
        dd_close.click()
        print("Closed year dropdown.")
        time.sleep(5) # Wait for data to potentially update based on year selection

        # --- MODIFICATION: Week checking and setting step ---
        print("--- Ensuring week is set to 53 before starting downloads ---")
        target_initial_week = 53
        max_attempts_set_week = 70 # Max clicks to prevent infinite loop (e.g. 52 up + 52 down)
        attempt_count = 0
        successfully_set_week = False

        while attempt_count < max_attempts_set_week:
            try:
                weeknum_div_check = wait.until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "sliderText")) # Use visibility for text reading
                )
                current_week_on_page_text = weeknum_div_check.text
                match = re.search(r'\d+', current_week_on_page_text)
                if not match:
                    print(f"Error: Could not parse week number from '{current_week_on_page_text}'. Skipping week set attempt.")
                    break
                current_week_on_page = int(match.group())
                print(f"Week check attempt {attempt_count + 1}: Dashboard week is {current_week_on_page}. Target: {target_initial_week}")

                if current_week_on_page == target_initial_week:
                    print(f"Dashboard week is correctly set to {target_initial_week}.")
                    successfully_set_week = True
                    break
                elif current_week_on_page < target_initial_week:
                    # !!! IMPORTANT: You MUST verify the selector for the INCREMENT button !!!
                    increment_button_xpath = "//*[contains(@class, 'tableauArrowInc')]" # <<< USER MUST VERIFY THIS XPATH
                    print(f"Week {current_week_on_page} < {target_initial_week}. Clicking increment (XPATH: {increment_button_xpath})")
                    try:
                        increment_button = wait.until(EC.element_to_be_clickable((By.XPATH, increment_button_xpath)))
                        increment_button.click()
                    except Exception as e_inc:
                        print(f"ERROR: Increment button not found or not clickable using XPATH: '{increment_button_xpath}'. {e_inc}")
                        print("Cannot increase week. If week remains below target, download will start from current week or fail if logic expects 53.")
                        # If we can't increment, and we are below target, we can't reach it.
                        successfully_set_week = False # Mark as not successfully set
                        break # Exit week setting loop
                else: # current_week_on_page > target_initial_week
                    print(f"Week {current_week_on_page} > {target_initial_week}. Clicking decrement.")
                    decrement_button_check = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'tableauArrowDec')]")))
                    decrement_button_check.click()

                time.sleep(2.5) # Wait for UI to update after click
                attempt_count += 1
            except Exception as e:
                print(f"Error during week check/set (attempt {attempt_count + 1}): {e}")
                attempt_count += 1
                time.sleep(1) # Wait a bit before retrying

        if not successfully_set_week:
            print(f"Warning: Could not definitively set dashboard week to {target_initial_week} after {attempt_count} attempts.")
            # Add a final check of the week to see where we ended up
            try:
                weeknum_div_final = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "sliderText")))
                final_week_text = weeknum_div_final.text
                final_week_match = re.search(r'\d+', final_week_text)
                if final_week_match:
                     print(f"Proceeding with dashboard week at: {final_week_match.group()}. The loop will still iterate 53 times if `weeknum` remains 53.")
                else:
                    print("Could not read final week on page. Loop will proceed based on its counter.")
            except:
                print("Could not read final week on page after setting attempts.")

        time.sleep(3) # Final pause after attempting to set week

        # --- Original loop for downloading and renaming files ---
        # The 'weeknum' variable here controls the number of iterations (53 times)
        # The actual week number for the filename is read from the page in download_and_rename
        loop_iteration_count = 53  # This is the `weeknum` from your original loop, signifies iterations

        # We start the loop, assuming the page is at or near week 53 (or highest available if setting failed).
        # The first action in the loop is to click decrement, so the first download will be for week 52 (if page was at 53).
        # OR, if the first data to download is for week 53 itself, then download_and_rename should be called *before* the first decrement.
        # Your original code: click decrement THEN download. So it downloads 52, 51, ..., 0 (effectively).
        # If week 0 is not valid, the loop should be `while loop_iteration_count > 1` and download week 1.
        # Let's adjust to match common "download week N... then week 1"

        # If initial state after setting is week 53, and you want to download week 53 first:
        print(f"\n--- Starting download iterations. Expecting to iterate {loop_iteration_count} times. ---")
        print(f"Initial download for the current week on page (should be around 53).")
        download_and_rename(wait, shadow_doc2, loop_iteration_count, default_dir, downloadPath, driver, year_for_filename, today)

        for i in range(loop_iteration_count - 1, 0, -1): # Iterate for weeks 52 down to 1
            current_target_iteration_week = i # This represents the week number we are targeting in this iteration (e.g. 52, 51 ... 1)
            print(f"\nProcessing for (target) Week Number: {current_target_iteration_week} (Iteration {loop_iteration_count - i}/{loop_iteration_count-1} of decrements)")

            decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'tableauArrowDec')]")))
            decrement_button.click()
            print("Clicked decrement for next week's data.")
            time.sleep(3) # Original sleep, allow data to update

            # The weeknum parameter passed here (current_target_iteration_week) is mostly for logging or loop control.
            # download_and_rename reads the actual week from the page for the filename.
            download_and_rename(wait, shadow_doc2, current_target_iteration_week, default_dir, downloadPath, driver, year_for_filename, today)

        print("Finished all download iterations.")

    except Exception as e_main:
        print(f"An error occurred in the main script: {e_main}")
        import traceback
        traceback.print_exc()
        if driver:
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                driver.save_screenshot(f"error_screenshot_{ts}.png")
                print(f"Screenshot saved as error_screenshot_{ts}.png")
            except:
                pass # ignore screenshot error if driver is problematic
    finally:
        if driver:
            print("Closing browser.")
            driver.quit()

iterate_weekly()
