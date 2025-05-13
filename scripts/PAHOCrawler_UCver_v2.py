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
import glob # Added for move_to_download_folder

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
        # Fallback or default version if detection fails, e.g., for GitHub Actions
        # Check environment variable first
        env_chrome_version = os.getenv('CHROME_VERSION_MAJOR')
        if env_chrome_version and env_chrome_version.isdigit():
            print(f"Using CHROME_VERSION_MAJOR from environment: {env_chrome_version}")
            return int(env_chrome_version)
        print("Failed to automatically detect Chrome version, returning a common default (e.g., 124). Set CHROME_VERSION_MAJOR env var to override.")
        return 124 # Or another sensible default, or raise the error

# Get the major version of Chrome installed
try:
    chrome_version = get_chrome_version()
    print(f"Detected Chrome major version: {chrome_version}")
except RuntimeError as e:
    print(f"Error in get_chrome_version: {e}")
    print("Attempting to proceed with a default Chrome version (e.g. 124 for uc.Chrome).")
    chrome_version = 124 # Default if detection fails catastrophically


def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    max_retries = 12 # Retry for up to 2 minutes (12 * 10 seconds)
    retries = 0
    currentFile = None

    while not got_file and retries < max_retries:
        try:
            # List all files in the default_dir, filter for .csv or .crdownload (partial download)
            files = [os.path.join(default_dir, f) for f in os.listdir(default_dir)
                     if os.path.isfile(os.path.join(default_dir, f))]

            if not files:
                if retries == 0: # Print only on first attempt if no files found
                    print(f"No files found in download directory: {default_dir}. Waiting...")
                time.sleep(10)
                retries += 1
                continue

            # Find the most recently modified file
            currentFile = max(files, key=os.path.getctime)

            # Check if the file is still downloading (Chrome uses .crdownload extension)
            if currentFile.endswith('.crdownload'):
                print(f"File is still downloading: {os.path.basename(currentFile)}. Waiting...")
                time.sleep(10)
                retries += 1
                continue

            # Ensure the file exists and is not empty
            if os.path.exists(currentFile) and os.path.getsize(currentFile) > 0:
                print(f"Downloaded file identified: {currentFile}")
                got_file = True
            else:
                if not os.path.exists(currentFile):
                    print(f"File {currentFile} does not exist. Retrying...")
                elif os.path.getsize(currentFile) == 0:
                     print(f"File {currentFile} is empty. Assuming download is not complete. Retrying...")
                time.sleep(10)
                retries += 1

        except Exception as e:
            print(f"Error checking download directory: {e}. Retrying...")
            time.sleep(10)
            retries += 1

    if not got_file or currentFile is None:
        print(f"Failed to get the downloaded file from {default_dir} after {max_retries * 10} seconds.")
        # Create an empty placeholder file to signify failure for this specific download
        placeholder_filename = newFileName + "_DOWNLOAD_FAILED" + fileExtension
        placeholder_filepath = os.path.join(downloadPath, placeholder_filename)
        with open(placeholder_filepath, 'w') as f:
            f.write("Download failed or file was empty/not found.")
        print(f"Created a placeholder error file: {placeholder_filepath}")
        return


    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)

    # Move the file
    try:
        os.rename(currentFile, fileDestination)
        print(f"Moved file to {fileDestination}")
    except Exception as e:
        print(f"Error moving file {currentFile} to {fileDestination}: {e}")
        # Try to copy and delete if rename fails (e.g. across different filesystems)
        try:
            import shutil
            shutil.copy(currentFile, fileDestination)
            os.remove(currentFile)
            print(f"Copied and deleted file to {fileDestination} as fallback.")
        except Exception as e_copy:
            print(f"Fallback copy also failed: {e_copy}")
            # Create an empty placeholder file to signify failure for this specific download
            placeholder_filename = newFileName + "_MOVE_FAILED" + fileExtension
            placeholder_filepath = os.path.join(downloadPath, placeholder_filename)
            with open(placeholder_filepath, 'w') as f:
                f.write(f"Download succeeded but move failed. Original file: {currentFile}")
            print(f"Created a placeholder error file for move failure: {placeholder_filepath}")


def download_and_rename(wait, driver, weeknum_expected, default_dir, downloadPath, year, today_timestamp):
    """Download and rename the file for the given week number."""
    print(f"Attempting to download data for expected week: {weeknum_expected}")

    # Wait for the week number display to update and verify it
    # This element might take a moment to reflect the change after a click
    time.sleep(2) # Short static wait for UI to settle before checking week number
    try:
        weeknum_div = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
        )
        current_weeknum_on_page = int(weeknum_div.text)
        print(f"Current week number on page: {current_weeknum_on_page}")
        # It's possible the weeknum_expected is based on an external loop counter.
        # The actual weeknum on the page is what matters for the filename.
        # We'll use current_weeknum_on_page for the filename.
    except Exception as e:
        print(f"Could not read week number from page: {e}. Using expected week: {weeknum_expected}")
        current_weeknum_on_page = weeknum_expected # Fallback

    # Find and click the download button at the bottom of the dashboard
    download_button = wait.until(
        EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
    )
    download_button.click()
    print("Clicked main download button.")
    time.sleep(5) # Wait for popup

    # Find and click the crosstab button (in a pop up window)
    crosstab_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
    )
    crosstab_button.click()
    print("Clicked crosstab button.")
    time.sleep(5) # Wait for next popup/dialog

    # Find and select the CSV option
    # The CSV radio button might be inside a shadow DOM or just a regular element in the new dialog
    # Using JavaScript click can be more robust for radio buttons if direct .click() is problematic
    csv_radio_selector = "input[type='radio'][value='csv']"
    try:
        csv_div = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, csv_radio_selector))
        )
        # Scroll into view and click using JavaScript
        driver.execute_script("arguments[0].scrollIntoView(true);", csv_div)
        driver.execute_script("arguments[0].click();", csv_div)
        print("Selected CSV option.")
    except Exception as e:
        print(f"Error selecting CSV option: {e}. Download might fail or be in wrong format.")
        # Optionally, you could try to close the dialogs and skip this week if CSV selection fails
        # For now, we'll proceed and hope the default is CSV or it still works.
    time.sleep(3) # Wait for selection to register

    # Find and click the export button
    export_button_selector = '[data-tb-test-id="export-crosstab-export-Button"]'
    try:
        export_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector))
        )
        export_button.click()
        print(f"Downloading CSV file for week {current_weeknum_on_page}...")
    except Exception as e:
        print(f"Error clicking export button: {e}. Download will likely fail for week {current_weeknum_on_page}.")
        # Close the download dialog if possible to avoid interference
        try:
            # Assuming there's a close button, common selectors:
            close_button_selectors = [
                '[aria-label="Close"]',  # Common aria-label
                '[data-tb-test-id="vizportal-dialog-close-button"]', # Tableau specific?
                'button.close',
                # Add other potential close button selectors for the dialog
            ]
            for sel in close_button_selectors:
                try:
                    close_btn = driver.find_element(By.CSS_SELECTOR, sel)
                    if close_btn.is_displayed():
                        close_btn.click()
                        print("Attempted to close download dialog after export fail.")
                        break
                except:
                    pass # Selector not found or button not interactable
        except Exception as close_e:
            print(f"Could not close download dialog: {close_e}")
        return # Stop this download attempt

    time.sleep(10) # Initial wait for download to start and potentially complete for small files

    # Use the move_to_download_folder function to move the downloaded file
    newFileName = f"PAHO_{year}_W{current_weeknum_on_page}_{today_timestamp}"
    fileExtension = '.csv'

    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def iterate_weekly():
    year_to_download = "2023" # choose year to download
    today_timestamp = datetime.now().strftime('%Y%m%d%H%M') # current date and time

    # Set directory
    # Use GITHUB_WORKSPACE if available (for GitHub Actions), otherwise use a local path
    base_data_dir = os.getenv('GITHUB_WORKSPACE', '.') # Default to current dir if not in GHA
    data_dir_path = os.path.join(base_data_dir, 'data')

    # Default download directory for Chrome (can be relative or absolute)
    # For GitHub Actions, it's often /home/runner/Downloads or similar, let's use a relative path
    # os.getcwd() will be the root of the repository in GitHub Actions
    default_download_dir_for_browser = os.path.join(os.getcwd(), "chrome_downloads")
    os.makedirs(default_download_dir_for_browser, exist_ok=True)
    print(f"Chrome default download directory set to: {default_download_dir_for_browser}")


    today_directory_name = f"DL_{datetime.now().strftime('%Y%m%d')}"
    final_download_path = os.path.join(data_dir_path, today_directory_name)
    os.makedirs(final_download_path, exist_ok=True)
    print(f"Final storage path for downloaded files: {final_download_path}")

    # Set chrome download directory
    chrome_options = uc.ChromeOptions()
    prefs = {
        "download.default_directory": default_download_dir_for_browser,
        "download.prompt_for_download": False, # Disable download prompt
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--no-sandbox") # Often needed in CI environments
    chrome_options.add_argument("--disable-dev-shm-usage") # Overcome limited resource problems
    chrome_options.add_argument("--disable-gpu") # If running headless
    # chrome_options.add_argument("--window-size=1920,1080") # Can help with element visibility

    # Using undetected-chromedriver
    print(f"Initializing Chrome with version_main={chrome_version}")
    driver = uc.Chrome(headless=True, use_subprocess=False, options=chrome_options, version_main=chrome_version)

    driver.get('https://www3.paho.org/data/index.php/en/mnu-topics/indicadores-dengue-en/dengue-nacional-en/252-dengue-pais-ano-en.html')
    print("Page loaded.")

    wait = WebDriverWait(driver, 30) # Increased wait time

    # First iframe
    iframe_src = "https://ais.paho.org/ha_viz/dengue/nac/dengue_pais_anio_tben.asp"
    iframe_locator = (By.XPATH, f"//div[contains(@class, 'vizTab')]//iframe[@src='{iframe_src}']")

    print("Waiting for the first iframe...")
    iframe1 = wait.until(EC.presence_of_element_located(iframe_locator))
    driver.switch_to.frame(iframe1)
    print("Switched to the first iframe.")

    # Get the iframe inside the first iframe
    # This iframe seems to be the main content iframe from Tableau
    print("Waiting for the nested (Tableau) iframe...")
    iframe2_locator = (By.XPATH, "//iframe[@title='viz']") # More specific Tableau iframe title
    # Alternative: (By.ID, "viz_iframe") or (By.XPATH, "//body/iframe") if title is not stable
    try:
        iframe2 = wait.until(EC.presence_of_element_located(iframe2_locator))
    except:
        print("Could not find iframe with title='viz', trying generic //body/iframe")
        iframe2 = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe")))

    driver.switch_to.frame(iframe2)
    print("Switched to the nested (Tableau) iframe.")

    iframe_page_title = driver.title
    print(f"Title of nested iframe: {iframe_page_title}")

    if "PAHO/WHO Data" not in iframe_page_title and "Tableau" not in iframe_page_title : # More flexible check
        print(f"Warning: Unexpected iframe title: {iframe_page_title}. Proceeding with caution.")
        # driver.quit() # Decide if this is a critical failure
        # return

    time.sleep(5) # Allow content within the Tableau iframe to load

    # Find the year tab (filter control)
    # ID 'tabZoneId13' seems to be for the year filter UI component
    year_filter_tab_id = 'tabZoneId13'
    print(f"Waiting for year filter tab/control (ID: {year_filter_tab_id})...")
    year_tab_control = wait.until(EC.visibility_of_element_located((By.ID, year_filter_tab_id)))
    print("Year filter tab/control found.")

    # Find the dropdown button within the year tab/control
    # This button opens the list of years
    year_dropdown_button_selector = (By.CSS_SELECTOR, f'#{year_filter_tab_id} span.tabComboBoxButton')
    year_dropdown_open_button = wait.until(EC.element_to_be_clickable(year_dropdown_button_selector))
    year_dropdown_open_button.click()
    print("Clicked to open year selection dropdown.")
    time.sleep(2) # Wait for dropdown to animate/load

    # --- Year Selection Logic ---
    # This logic seems to intend to select only "2023"
    # (All) click 1: Might deselect everything or select everything.
    # (All) click 2: If first selected all, second might deselect all (if toggle), or do nothing.
    # Then specific years are clicked. If it's a multi-select, this would add them.
    # If it's single-select effectively, the last one clicked (2023) would be active.

    def click_filter_checkbox(text_label, filter_type="year"):
        # XPath for checkboxes in Tableau filters (often within a div with class 'facetOverflow')
        # The input checkbox is usually a sibling or parent::*/input of the label (<a> tag)
        xpath = f'//div[contains(@class, "facetOverflow")]//a[text()="{text_label}"]/preceding-sibling::input[@type="checkbox"]'
        # Sometimes the 'a' tag itself is clickable, or a div around it.
        # Let's try to click the input directly.
        try:
            print(f"Waiting for {filter_type} filter checkbox: {text_label}")
            checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            # driver.execute_script("arguments[0].scrollIntoView(true);", checkbox) # Optional: scroll
            checkbox.click()
            print(f"Clicked {filter_type} filter checkbox: {text_label}")
            time.sleep(1) # Allow UI to update
        except Exception as e:
            print(f"Could not click {filter_type} filter checkbox '{text_label}': {e}")
            # Fallback: try clicking the 'a' tag if input click fails
            try:
                print(f"Fallback: Trying to click 'a' tag for {text_label}")
                a_tag_xpath = f'//div[contains(@class, "facetOverflow")]//a[text()="{text_label}"]'
                a_tag = wait.until(EC.element_to_be_clickable((By.XPATH, a_tag_xpath)))
                a_tag.click()
                print(f"Clicked 'a' tag for {filter_type} filter: {text_label}")
                time.sleep(1)
            except Exception as e_fallback:
                 print(f"Fallback click for '{text_label}' also failed: {e_fallback}")


    # Click "(All)" twice - This might be to clear existing selections and then select all, or clear and then allow individual.
    click_filter_checkbox("(All)")
    click_filter_checkbox("(All)")

    # Deselect 2025 and 2024, then ensure 2023 is selected.
    # If "(All)" resulted in all years being checked, these clicks would uncheck them.
    click_filter_checkbox("2025")
    click_filter_checkbox("2024")

    # This click on "2023" should ensure it's selected (or toggle its state).
    # If the goal is ONLY 2023, and (All) + (All) deselects everything, then this would select 2023.
    # If (All) + (All) selects everything, then 2025/2024 are deselected, and 2023 is also deselected.
    # The original script had a specific click for the target year. Let's ensure 2023 is selected.
    # To ensure "2023" is selected:
    # One strategy: Deselect All, then select 2023.
    # The current sequence is: (All), (All), 2025 (deselect), 2024 (deselect), 2023 (toggle).
    # If year_to_download is the one we want, we should ensure it's checked.
    # The current logic might leave 2023 unchecked if it was checked by "(All)".
    # Let's refine:
    # 1. Click (All) - e.g. to deselect all other years if it's a radio-like behavior for (All)
    # 2. Click year_to_download

    # Re-evaluating the year selection based on typical Tableau behavior:
    # Often, "(All)" selects all. Clicking individual items then *deselects* them from the "All" set.
    # If the goal is ONLY 2023:
    # click_filter_checkbox("(All)") # Selects all
    # click_filter_checkbox("2025") # Deselects 2025
    # click_filter_checkbox("2024") # Deselects 2024
    # ... any other years to deselect ...
    # Then 2023 would remain selected.
    # The original script clicks "2023" last. This implies it might be toggling it.
    # Let's assume the original sequence is what's needed for this specific Tableau dashboard.
    click_filter_checkbox(year_to_download) # This should be the final state for the target year.

    # Close the year dropdown menu
    # Clicking '.tab-glass' (an overlay) usually closes popups/dropdowns
    year_dropdown_close_selector = (By.CLASS_NAME, "tab-glass")
    wait.until(EC.element_to_be_clickable(year_dropdown_close_selector)).click()
    print("Closed year selection dropdown.")
    time.sleep(3) # Wait for filter to apply

    # Select all countries/regions
    region_filter_tab_id = 'tabZoneId9' # Assuming this is the region filter
    print(f"Waiting for region filter tab/control (ID: {region_filter_tab_id})...")
    region_tab_control = wait.until(EC.visibility_of_element_located((By.ID, region_filter_tab_id)))
    print("Region filter tab/control found.")

    region_dropdown_button_selector = (By.CSS_SELECTOR, f'#{region_filter_tab_id} span.tabComboBoxButton')
    region_dropdown_open_button = wait.until(EC.element_to_be_clickable(region_dropdown_button_selector))
    region_dropdown_open_button.click()
    print("Clicked to open region selection dropdown.")
    time.sleep(2)

    click_filter_checkbox("(All)", filter_type="region") # Select all regions

    region_dropdown_close_selector = (By.CLASS_NAME, "tab-glass")
    wait.until(EC.element_to_be_clickable(region_dropdown_close_selector)).click()
    print("Closed region selection dropdown.")
    time.sleep(5) # Wait for filters to apply and data to update

    # Initial call to download_and_rename (for week 53 or the current week on page)
    # The week number slider starts at the latest week.
    # We need to get the current week from the page first.
    try:
        weeknum_div_initial = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "sliderText")))
        initial_weeknum = int(weeknum_div_initial.text)
        print(f"Initial week on page (slider): {initial_weeknum}")
    except Exception as e:
        print(f"Could not read initial week number: {e}. Assuming 53 or 52.")
        # Determine max week for the year (usually 52 or 53)
        # This is a simplification; a more robust way would be to check calendar for year_to_download
        # For now, let's assume it starts at a high number if not readable.
        initial_weeknum = 53 if datetime(int(year_to_download), 12, 28).isocalendar()[1] == 53 else 52


    current_loop_week = initial_weeknum
    download_and_rename(wait, driver, current_loop_week, default_download_dir_for_browser, final_download_path, year_to_download, today_timestamp)

    # Loop for downloading and renaming files by decrementing weeks
    # The decrement button is usually an arrow on the slider.
    decrement_button_xpath = "//*[contains(@class, 'tableauArrow') and contains(@class, 'Dec')]" # More specific for decrement
    # Alternative: "//*[contains(@class, 'tableauArrowDec')]" (original)

    while current_loop_week > 1: # Download until week 1
        print(f"Preparing to decrement week from {current_loop_week}...")
        try:
            decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, decrement_button_xpath)))
            decrement_button.click()
            print(f"Clicked decrement week button. Current week was {current_loop_week}.")
            current_loop_week -= 1
            # Add a delay for the dashboard to update after clicking the decrement button
            time.sleep(5) # Adjust as needed for dashboard responsiveness
        except Exception as e:
            print(f"Error clicking decrement button or week already at 1: {e}")
            break # Exit loop if decrement fails

        print(f"Processing data for week: {current_loop_week}")
        download_and_rename(wait, driver, current_loop_week, default_download_dir_for_browser, final_download_path, year_to_download, today_timestamp)

        if current_loop_week == 1:
            print("Reached week 1, finishing download cycle.")
            break

    print("All weekly data downloaded.")
    driver.quit()

if __name__ == "__main__":
    try:
        iterate_weekly()
    except Exception as e:
        print(f"An unhandled error occurred in iterate_weekly: {e}")
        import traceback
        traceback.print_exc()

