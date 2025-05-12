import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys # Need this import
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException # Ensure these are imported
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
            # Using a more robust registry query for default Chrome install location
            try:
                 output = subprocess.check_output(
                     r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Google\Chrome\Update" /v LastKnownVersionString /reg:32',
                     shell=True, text=True, stderr=subprocess.DEVNULL
                 )
            except subprocess.CalledProcessError:
                 # Try 64-bit location if 32-bit fails
                 output = subprocess.check_output(
                     r'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Google\Chrome\Update" /v LastKnownVersionString /reg:64',
                     shell=True, text=True, stderr=subprocess.DEVNULL
                 )
            except Exception:
                 # Fallback to user-specific install location if machine-wide fails
                 output = subprocess.check_output(
                    r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                    shell=True, text=True, stderr=subprocess.DEVNULL
                 )
            version_match = re.search(r'REG_SZ\s+(\d+)\.', output)
        else:  # Assuming Linux or macOS
            version_match = None
            # Try different commands to retrieve Chrome version
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
            if not version_match:
                 raise RuntimeError("Could not determine Chrome version using common commands.")

        if version_match:
            print(f"Detected Chrome major version: {version_match.group(1)}")
            return int(version_match.group(1))
        else:
            raise ValueError("Could not parse Chrome version from output.")
    except Exception as e:
        raise RuntimeError(f"Failed to get Chrome version: {e}") from e

# Get the major version of Chrome installed
try:
    chrome_version = get_chrome_version()
except RuntimeError as e:
    print(f"Error getting Chrome version: {e}")
    print("Please ensure Chrome is installed or specify version manually.")
    # Handle error appropriately, e.g., exit or ask for input
    sys.exit(1)


# Function to move downloaded file
def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    """
    Waits for a download to complete in default_dir, then moves the latest file
    to downloadPath with the new name.
    """
    got_file = False
    max_wait_time = 120  # Maximum time to wait for download in seconds (e.g., 2 minutes)
    start_time = time.time()
    currentFile = None

    print(f"Waiting for download in '{default_dir}'...")
    while not got_file and (time.time() - start_time) < max_wait_time:
        try:
            # Find files, ignore temporary download files (.crdownload)
            files_in_dir = [f for f in os.listdir(default_dir) if not f.lower().endswith('.crdownload') and os.path.isfile(os.path.join(default_dir, f))]
            if not files_in_dir:
                # print("No completed files found yet...")
                time.sleep(5)
                continue

            # Get the latest completed file based on creation time
            currentFile = max([os.path.join(default_dir, f) for f in files_in_dir], key=os.path.getctime)

            # Check if the file is still being written (basic size check)
            initial_size = os.path.getsize(currentFile)
            time.sleep(2) # Wait a moment to see if size changes
            current_size = os.path.getsize(currentFile)

            if initial_size == current_size and current_size > 0: # File size stable and not empty
                 print(f"Detected downloaded file: {os.path.basename(currentFile)} (Size: {current_size} bytes)")
                 got_file = True
            else:
                 print(f"File '{os.path.basename(currentFile)}' still downloading or empty (Size: {current_size})...")
                 time.sleep(5) # Wait longer if size changed

        except FileNotFoundError:
             print("Download directory check failed. Retrying...")
             time.sleep(5)
        except Exception as e:
            print(f"Error checking download directory: {e}. Retrying...")
            time.sleep(10)

    if not got_file:
        print(f"Error: Download did not complete within {max_wait_time} seconds.")
        raise TimeoutError("Download timeout exceeded.")

    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)

    # Ensure the destination directory exists
    os.makedirs(os.path.dirname(fileDestination), exist_ok=True)

    # Move the file, handle potential overwrite or errors
    try:
        if os.path.exists(fileDestination):
            print(f"Warning: File '{fileDestination}' already exists. Overwriting.")
            os.remove(fileDestination)
        os.rename(currentFile, fileDestination)
        print(f"Moved file to {fileDestination}")
    except OSError as oe:
        print(f"Error moving file from '{currentFile}' to '{fileDestination}': {oe}")
        # Consider copying and deleting as an alternative if rename fails across drives
        try:
            import shutil
            shutil.move(currentFile, fileDestination)
            print(f"Moved file using shutil.move to {fileDestination}")
        except Exception as sh_e:
            print(f"Shutil.move also failed: {sh_e}")
            raise


# Function to download and rename file
def download_and_rename(wait, shadow_doc2, weeknum_expected, default_dir, downloadPath, driver, year, today):
    """Download and rename the file for the given week number."""
    print("-" * 10 + f" Starting download for Year {year}, Expected Week {weeknum_expected} " + "-" * 10)
    TIMEOUT_SECONDS = 30 # Increased timeout for download interactions

    # Optional: Verify current week number again right before download if needed
    # try:
    #     weeknum_div = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "sliderText")))
    #     current_weeknum_text = weeknum_div.text.strip()
    #     current_weeknum = int("".join(filter(str.isdigit, current_weeknum_text)))
    #     if current_weeknum != weeknum_expected:
    #          print(f"WARNING: Expected week {weeknum_expected} but found {current_weeknum} just before download!")
    #          # Decide how to handle: proceed, error out, or try to fix
    #     else:
    #          print(f"Confirmed week {current_weeknum} before download.")
    # except Exception as e_verify:
    #     print(f"Warning: Could not re-verify week number before download: {e_verify}")

    # Use the week number passed to the function for the filename
    actual_weeknum_for_file = weeknum_expected

    try:
        # Find and click the download button at the bottom of the dashboard
        print("Locating and clicking main download button...")
        download_button = wait.until(
            EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
        )
        # Scroll into view just in case
        driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
        time.sleep(0.5)
        download_button.click()
        print("Clicked main download button.")
        time.sleep(3) # Allow dialog to appear

        # Find and click the crosstab button (in the popup dialog)
        print("Locating and clicking 'Crosstab' button...")
        crosstab_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
        )
        crosstab_button.click()
        print("Clicked 'Crosstab' button.")
        time.sleep(3) # Allow next dialog/options to appear

        # Find and select the CSV option using JavaScript click
        # This interaction happens within the shadow_doc2 context according to original script
        print("Locating and selecting 'CSV' option...")
        # Wait for the element to be present within shadow_doc2 context
        csv_radio_selector = "input[type='radio'][value='csv']"
        # Note: WebDriverWait needs the driver, not the shadow_doc element directly for waiting.
        # We need to execute find_element within the shadow_doc context.
        # A simple time.sleep might be okay here if waits are problematic across contexts.
        time.sleep(1) # Give shadow content time to render
        csv_div = shadow_doc2.find_element(By.CSS_SELECTOR, csv_radio_selector)
        # Use JS click for potentially tricky radio buttons
        driver.execute_script("arguments[0].scrollIntoView(true);", csv_div)
        driver.execute_script("arguments[0].click();", csv_div)
        print("Selected 'CSV' option.")
        time.sleep(2) # Allow selection to register

        # Find and click the export button using shadow_doc2 context
        print("Locating and clicking final 'Download' (Export) button...")
        export_button_selector = '[data-tb-test-id="export-crosstab-export-Button"]'
        export_button = shadow_doc2.find_element(By.CSS_SELECTOR, export_button_selector)
        # Scroll and use JS click for robustness
        driver.execute_script("arguments[0].scrollIntoView(true);", export_button)
        driver.execute_script("arguments[0].click();", export_button)
        print(f"Clicked Export/Download for Week {actual_weeknum_for_file}. Waiting for file...")
        # Download starts here, wait time is handled in move_to_download_folder

        # Define filename components
        newFileName = f"PAHO_{year}_W{actual_weeknum_for_file:02d}_{today}" # Format weeknum with leading zero
        fileExtension = '.csv'

        # Move and rename the downloaded file
        move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)
        print(f"Successfully downloaded and renamed file for Week {actual_weeknum_for_file}.")

    except TimeoutException as te:
        print(f"Error: Timeout during download process for week {actual_weeknum_for_file}: {te}")
        # Consider adding screenshot here: driver.save_screenshot(...)
        raise # Re-raise to indicate failure for this week
    except Exception as e:
        print(f"Error during download process for week {actual_weeknum_for_file}: {type(e).__name__} - {e}")
        # Consider adding screenshot here: driver.save_screenshot(...)
        raise # Re-raise to indicate failure for this week


# Main function to iterate through weeks
def iterate_weekly():

    year = "2023" # Choose year to download
    today = datetime.now().strftime('%Y%m%d%H%M') # Current date and time string

    # Set download directory
    # Use environment variable for flexibility in GitHub Actions / local runs
    base_data_dir = os.getenv('GITHUB_WORKSPACE', os.getcwd()) # Default to current dir if not in GHA
    downloadPathBase = os.path.join(base_data_dir, 'data') # Store data in a 'data' subfolder

    # Define the default download directory for Chrome (can be temporary)
    # Using os.getcwd() might place downloads in the repo root during Actions
    # safer to use a dedicated temp dir if possible, but os.getcwd() matches original script
    default_dir = os.getcwd()
    print(f"Chrome default download directory set to: {default_dir}")

    # Create dated subfolder within the data directory
    today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
    downloadPath = os.path.join(downloadPathBase, today_directory_name)
    os.makedirs(downloadPath, exist_ok=True) # Create destination directory if it doesn't exist
    print(f"Final files will be moved to: {downloadPath}")

    # --- Undetected Chromedriver Setup ---
    driver = None # Initialize driver
    try:
        print("Setting up undetected-chromedriver...")
        chrome_options = uc.ChromeOptions()
        # Set download preferences
        prefs = {
             "download.default_directory": default_dir,
             "download.prompt_for_download": False, # Disable prompt
             "download.directory_upgrade": True,
             "safeBrowse.enabled": True # Keep safe Browse enabled
        }
        chrome_options.add_experimental_option("prefs", prefs)
        # Add headless option if needed
        # chrome_options.add_argument('--headless')
        # chrome_options.add_argument('--disable-gpu')

        driver = uc.Chrome(headless=False, # Set to True for headless runs
                          use_subprocess=False,
                          options=chrome_options,
                          version_main=chrome_version)

        print("Navigating to PAHO data page...")
        # Navigate to the page containing the first iframe
        driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')

        # Define wait with a reasonable timeout
        wait = WebDriverWait(driver, 30) # Increased timeout

        # --- Handle IFrames ---
        print("Switching to the first iframe (main viz)...")
        iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
        # Using CSS selector for potentially more robust iframe location
        iframe_locator = (By.CSS_SELECTOR, f"iframe[src='{iframe_src}']")
        iframe = wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))
        print("Switched to first iframe.")
        time.sleep(2) # Allow iframe content to load

        # --- Shadow DOM / Second IFrame Handling ---
        # The original script uses 'document' and then finds another iframe.
        # This implies the viz might be nested. Let's replicate carefully.
        print("Looking for nested iframe within the first iframe's content...")
        # Execute script in the current frame context to find the nested iframe
        # Using find_element directly on driver should work if it's not shadow DOM
        nested_iframe_locator = (By.TAG_NAME, "iframe") # Assuming only one iframe inside
        # Wait for the nested iframe to be present
        iframe2 = wait.until(EC.presence_of_element_located(nested_iframe_locator))
        driver.switch_to.frame(iframe2)
        print("Switched to the nested iframe.")
        time.sleep(2) # Allow nested frame content to load

        # Get the document context *after* switching to the second iframe
        # This context is used for shadow DOM interactions if needed, based on original script
        # Although the original script assigns 'document', it uses it like the driver context
        # for shadow_doc2.find_element. Let's keep the logic but clarify.
        # 'shadow_doc2' will represent the context of this second iframe for interactions.
        # For interactions needing shadow DOM within this iframe, you'd get the shadowRoot first.
        # Based on user's working year selection, direct find_element on shadow_doc2 works,
        # which implies shadow_doc2 might actually be the shadow root or behaves like it.
        # Let's assume shadow_doc2 IS the search context needed for year/CSV download based on original script.
        # We get it *after* switching frame.
        # shadow_doc2 = driver.execute_script('return document.body') # Example, get body as search context
        # For safety and clarity, let's just use 'driver' which IS now focused on iframe2
        # and pass 'driver' where 'shadow_doc2' was used, unless specific shadow DOM required.
        # Re-evaluating: The year selection `shadow_doc2.find_element` worked.
        # Let's assume `driver.execute_script('return document')` gives the necessary context.
        shadow_doc2 = driver.execute_script('return document') # Keep original logic's variable name

        iframe_page_title = driver.title
        print(f"Title of nested iframe content: {iframe_page_title}")

        # Check if we are in the right place
        if "PAHO/WHO Data" not in iframe_page_title: # More flexible check
             print("Error: Unexpected content in the nested iframe. Title: " + iframe_page_title)
             driver.quit()
             return # Exit function

        print("Successfully accessed the main dashboard content.")
        time.sleep(5) # Longer pause for dashboard elements like filters to fully render

        # --- Select Year ---
        print(f"Selecting year '{year}'...")
        # Wait for the year tab container element to be visible
        year_tab_id = 'tabZoneId13' # Assuming this ID is stable
        year_tab = wait.until(EC.visibility_of_element_located((By.ID, year_tab_id)))
        print("Year tab container found.")

        # Find and click the dropdown button within the year tab
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = year_tab.find_element(*dd_locator) # Find within year_tab context
        dd_open.click()
        print("Clicked year dropdown.")
        time.sleep(1) # Allow dropdown to open

        # Deselect 2024 (using shadow_doc2 context as per original script)
        # Ensure element is clickable
        print("Deselecting 2024...")
        y2024_xpath = '//div[contains(@class, "facetOverflow")]//a[text()="2024"]/preceding-sibling::input'
        y2024_input = WebDriverWait(driver, 10).until(
             EC.element_to_be_clickable((By.XPATH, y2024_xpath))
             )
        # Use JavaScript click for potentially tricky elements
        driver.execute_script("arguments[0].click();", y2024_input)
        print("Deselected 2024.")
        time.sleep(0.5)

        # Ensure 2023 is selected (using shadow_doc2 context as per original script)
        print("Ensuring 2023 is selected...")
        y2023_xpath = '//div[contains(@class, "facetOverflow")]//a[text()="2023"]/preceding-sibling::input'
        y2023_input = WebDriverWait(driver, 10).until(
             EC.element_to_be_clickable((By.XPATH, y2023_xpath))
             )
        # Check if it's already selected before clicking
        if not y2023_input.is_selected():
             driver.execute_script("arguments[0].click();", y2023_input)
             print("Selected 2023.")
        else:
             print("2023 was already selected.")
        time.sleep(0.5)

        # Close the dropdown menu by clicking the overlay glass pane
        print("Closing year dropdown...")
        dd_close_locator = (By.CLASS_NAME, "tab-glass") # Click overlay to close
        dd_close = wait.until(EC.element_to_be_clickable(dd_close_locator))
        dd_close.click()
        print("Closed year dropdown.")
        time.sleep(3) # Allow filter to apply

        # --- Select all countries (Commented out as per original script) ---
        # print("Ensuring 'All' countries are selected (Original code commented out)...")
        # ...

        # <<< NEW SECTION START: Ensure Epidemiological Week is 53 >>>
        # Uses 'wait' (driver context of iframe2), not 'shadow_doc2', based on HTML analysis
        print("-" * 30)
        print("Ensuring Epidemiological Week is set to 53...")
        TARGET_WEEK_SET = 53
        WEEK_TIMEOUT_SECONDS = 20 # Specific timeout for this section

        # Locators for the week filter interaction
        SLIDER_TEXT_LOCATOR_WEEK = (By.CSS_SELECTOR, ".sliderText") # For reading current week
        WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR = (By.ID, "dijit_form_Button_3") # Search icon
        SEARCH_INPUT_TEXT_FIELD_LOCATOR = (By.ID, "dijit_form_ComboBox_0") # Input field

        current_slider_value = -1 # Default to trigger update

        try:
            # Attempt to read current week. Might fail if sliderText is hidden initially.
            print(f"Attempting to read current week from: {SLIDER_TEXT_LOCATOR_WEEK}")
            # Use WebDriverWait for robustness, check visibility
            visible_slider_text_elems = wait.until(EC.presence_of_all_elements_located(SLIDER_TEXT_LOCATOR_WEEK))
            displayed_elem = None
            for elem in visible_slider_text_elems:
                if elem.is_displayed():
                    displayed_elem = elem
                    break

            if displayed_elem:
                current_slider_value_text = displayed_elem.text.strip()
                cleaned_text = "".join(filter(str.isdigit, current_slider_value_text))
                if cleaned_text:
                    current_slider_value = int(cleaned_text)
                    print(f"Current week detected as: {current_slider_value}")
                else:
                    print(f"Could not parse digits from slider text: '{current_slider_value_text}'. Proceeding to update.")
            else:
                print(f"Element {SLIDER_TEXT_LOCATOR_WEEK} not found or not visible. Proceeding to set week to {TARGET_WEEK_SET}.")

        except TimeoutException:
            print(f"Timeout waiting for {SLIDER_TEXT_LOCATOR_WEEK} to read initial value. Assuming week needs to be set.")
        except Exception as e_read:
            print(f"Could not read initial week value (element {SLIDER_TEXT_LOCATOR_WEEK}): {e_read}. Will attempt to set to {TARGET_WEEK_SET} anyway.")

        if current_slider_value != TARGET_WEEK_SET:
            if current_slider_value != -1:
                print(f"Current week {current_slider_value} is not {TARGET_WEEK_SET}. Updating...")
            else:
                print(f"Attempting to set week to {TARGET_WEEK_SET}.")

            # Step 1: Click the "Search" icon/button
            try:
                print(f"Locating and clicking the week search activator button: {WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}")
                search_activator_button = wait.until(
                    EC.element_to_be_clickable(WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR)
                )
                search_activator_button.click()
                print("Clicked the week search activator button.")
                time.sleep(1.5) # Pause for search input to appear/activate
            except TimeoutException:
                print(f"Error: Could not find or click the Week Search Activator Button ({WEEK_SEARCH_ACTIVATOR_BUTTON_LOCATOR}).")
                driver.save_screenshot("error_clicking_week_search_activator.png")
                raise # Stop script if this fails

            # Step 2: Interact with the search input field
            try:
                print(f"Waiting for the week search input field ({SEARCH_INPUT_TEXT_FIELD_LOCATOR}) to be clickable...")
                search_input_field = wait.until(
                    EC.element_to_be_clickable(SEARCH_INPUT_TEXT_FIELD_LOCATOR)
                )
                print("Week search input field located.")
                search_input_field.click()
                search_input_field.clear()
                search_input_field.send_keys(str(TARGET_WEEK_SET))
                print(f"Typed '{TARGET_WEEK_SET}'.")
                time.sleep(0.5)
                search_input_field.send_keys(Keys.ENTER)
                print("Sent Keys.ENTER.")
                time.sleep(3) # Pause for filter to apply fully before verification/next step
            except TimeoutException:
                print(f"TimeoutException: Week search input field ({SEARCH_INPUT_TEXT_FIELD_LOCATOR}) was not clickable after activating search.")
                driver.save_screenshot("error_timeout_week_search_input.png")
                raise
            except ElementNotInteractableException as e:
                print(f"ElementNotInteractableException with week search input: {e}")
                driver.save_screenshot("error_interactable_week_search_input.png")
                raise

            # Step 3: Optional Verification (Check if sliderText updates)
            try:
                print(f"Verifying week update by checking {SLIDER_TEXT_LOCATOR_WEEK} for '{TARGET_WEEK_SET}'...")
                WebDriverWait(driver, WEEK_TIMEOUT_SECONDS).until(
                    EC.and_(
                        EC.visibility_of_element_located(SLIDER_TEXT_LOCATOR_WEEK),
                        EC.text_to_be_present_in_element_located(SLIDER_TEXT_LOCATOR_WEEK, str(TARGET_WEEK_SET))
                    )
                )
                print(f"Verification successful: Week appears to be set to {TARGET_WEEK_SET}.")
            except TimeoutException:
                print(f"Verification WARNING: Timeout waiting for slider text to update to {TARGET_WEEK_SET}. Proceeding anyway.")
                driver.save_screenshot("warning_verification_timeout.png")
            except Exception as e_verify:
                print(f"Verification WARNING: Error during verification step: {e_verify}. Proceeding anyway.")

        else: # current_slider_value == TARGET_WEEK_SET
            print(f"Week is already correctly set to {TARGET_WEEK_SET}.")

        print("Week setting process finished.")
        print("-" * 30)
        # <<< NEW SECTION END >>>

        # --- Download Loop ---
        # Start download process, assuming week is now 53
        weeknum_start = TARGET_WEEK_SET # Start from the week we just set

        print(f"--- Starting Download Loop for Year {year} from Week {weeknum_start} ---")

        # Initial download for week 53
        print(f"Processing Week Number: {weeknum_start}")
        download_and_rename(wait, shadow_doc2, weeknum_start, default_dir, downloadPath, driver, year, today)

        # Loop downwards from week 52 to 1
        weeknum = weeknum_start
        while weeknum > 1:
            print("-" * 20)
            print(f"Processing Week Number: {weeknum - 1}")
            # Locate and click the decrement arrow
            try:
                decrement_locator = (By.XPATH, "//*[contains(@class, 'tableauArrowDec') or contains(@class, 'dijitSliderDecrement')]")
                decrement_button = wait.until(EC.element_to_be_clickable(decrement_locator))
                decrement_button.click()
                print(f"Clicked decrement button for week {weeknum - 1}.")
                # Wait for the dashboard to potentially update after changing the week
                # A longer sleep might be needed if data reload is slow
                time.sleep(4) # Allow time for week change and potential data update
            except TimeoutException:
                print(f"Error: Could not find or click the decrement button for week {weeknum - 1}.")
                driver.save_screenshot(f"error_clicking_decrement_week_{weeknum-1}.png")
                # Decide whether to break or continue
                break # Stop the loop if decrement fails

            # Update weeknum after successfully decrementing
            weeknum -= 1

            # Download for the new week
            download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today)

            # Check if we need to break (redundant if while condition is weeknum > 1)
            # if weeknum == 1:
            #     print("Reached week 1, loop will terminate.")
            #     break
        print(f"--- Finished Download Loop for Year {year} ---")

    except Exception as e:
        print(f"An error occurred during the main iterate_weekly process: {type(e).__name__} - {e}")
        if driver:
             timestamp = time.strftime("%Y%m%d-%H%M%S")
             screenshot_path = os.path.join(os.getcwd(), f"error_main_process_{timestamp}.png")
             try:
                  driver.save_screenshot(screenshot_path)
                  print(f"Screenshot saved to: {screenshot_path}")
             except Exception as scr_e:
                  print(f"Could not save screenshot during error handling: {scr_e}")

    finally:
        if driver:
            print("Closing WebDriver...")
            driver.quit()
        print("Script finished.")

# Run the main function
if __name__ == "__main__":
    iterate_weekly()
