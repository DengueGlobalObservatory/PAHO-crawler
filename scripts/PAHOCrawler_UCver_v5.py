import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
import time
import os
from datetime import datetime
import subprocess
import re
import sys
import traceback
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
            # Try different commands to retrieve Chrome version on Linux
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
        logger.error(f"Failed to get Chrome version: {e}")
        # Fallback to a common version
        return 129  # Use a recent stable version as fallback

# Get the major version of Chrome installed
chrome_version = get_chrome_version()
logger.info(f"Chrome version detected: {chrome_version}")

def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    attempts = 0
    max_attempts = 30  # 5 minutes max wait

    while not got_file and attempts < max_attempts:
        try:
            # Use glob to get the current file name
            files = [default_dir + "/" + f for f in os.listdir(default_dir)]
            if not files:
                raise FileNotFoundError("No files in download directory")

            currentFile = max(files, key=os.path.getctime)

            # Ensure the file exists and is not a .crdownload file
            if os.path.exists(currentFile) and not currentFile.endswith('.crdownload'):
                got_file = True
            else:
                raise FileNotFoundError("File not found or still downloading")

        except Exception as e:
            attempts += 1
            logger.info(f"Attempt {attempts}: File has not finished downloading - {e}")
            time.sleep(10)

    if not got_file:
        raise TimeoutError("File download timed out after 5 minutes")

    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)

    # Move the file
    os.rename(currentFile, fileDestination)
    logger.info(f"Moved file to {fileDestination}")


def download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today):
    """Download and rename the file for the given week number."""
    try:
        # Wait for the week number to update
        weeknum_div = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
        )
        current_week = int(weeknum_div.text)
        logger.info(f"Current Week Number on page: {current_week}")

        # Find and click the download button at the bottom of the dashboard
        download_button = wait.until(
            EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
        )
        download_button.click()
        logger.info("Clicked download button")

        time.sleep(5)

        # Find and click the crosstab button (in a pop up window)
        crosstab_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
        )
        crosstab_button.click()
        logger.info("Clicked crosstab button")

        time.sleep(5)

        # Find and select the CSV option
        csv_div = shadow_doc2.find_element(By.CSS_SELECTOR, "input[type='radio'][value='csv']")
        driver.execute_script("arguments[0].click();", csv_div)
        logger.info("Selected CSV option")
        time.sleep(5)

        # Find and click the export button
        export_button = shadow_doc2.find_element(By.CSS_SELECTOR, '[data-tb-test-id="export-crosstab-export-Button"]')
        export_button.click()
        logger.info("Clicked export button - starting download")
        time.sleep(5)

        # Use the move_to_download_folder function to move the downloaded file
        newFileName = f"PAHO_{year}_W{current_week}_{today}"
        fileExtension = '.csv'

        move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)

        return current_week

    except Exception as e:
        logger.error(f"Error in download_and_rename: {e}")
        logger.error(traceback.format_exc())
        raise


def create_robust_driver():
    """Create a Chrome driver with robust settings for GitHub Actions"""
    try:
        # More robust way to get the base data path
        if os.getenv('GITHUB_WORKSPACE'):
            # Running in GitHub Actions
            base_data_path = os.path.join(os.getenv('GITHUB_WORKSPACE'), 'data')
        else:
            # Running locally
            base_data_path = os.path.join(os.getcwd(), 'data')

        # Ensure the base data directory exists
        os.makedirs(base_data_path, exist_ok=True)

        # Create temp downloads directory
        default_dir = os.path.join(os.getcwd(), "temp_downloads")
        os.makedirs(default_dir, exist_ok=True)

        # Create the download path within the data directory
        downloadPath = os.path.join(base_data_path, f"DL_{datetime.now().strftime('%Y%m%d')}")
        os.makedirs(downloadPath, exist_ok=True)

        chrome_options = uc.ChromeOptions()

        # Download preferences
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": default_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        })

        # Essential headless options for GitHub Actions
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")

        # Additional stability options for CI environments
        chrome_options.add_argument("--disable-background-timer-throttling")
        chrome_options.add_argument("--disable-backgrounding-occluded-windows")
        chrome_options.add_argument("--disable-renderer-backgrounding")
        chrome_options.add_argument("--disable-features=TranslateUI")
        chrome_options.add_argument("--disable-ipc-flooding-protection")
        chrome_options.add_argument("--memory-pressure-off")
        chrome_options.add_argument("--max_old_space_size=4096")
        chrome_options.add_argument("--single-process")  # This can help with stability in CI

        logger.info("Creating Chrome driver...")

        # Try to create driver with error handling
        try:
            driver = uc.Chrome(
                headless=True,
                use_subprocess=False,
                options=chrome_options,
                version_main=chrome_version
            )
            logger.info("Chrome driver created successfully")
            return driver, default_dir, downloadPath

        except Exception as e:
            logger.error(f"Failed to create Chrome driver: {e}")
            # Try without version_main specification
            logger.info("Retrying without version_main...")
            driver = uc.Chrome(
                headless=True,
                use_subprocess=False,
                options=chrome_options
            )
            logger.info("Chrome driver created successfully (without version_main)")
            return driver, default_dir, downloadPath

    except Exception as e:
        logger.error(f"Fatal error creating driver: {e}")
        raise


def iterate_weekly():
    year = "2023_2025"
    today = datetime.now().strftime('%Y%m%d%H%M')
    driver = None

    try:
        # Create driver with robust error handling
        driver, default_dir, downloadPath = create_robust_driver()

        logger.info("Navigating to PAHO website...")
        driver.get('https://www.paho.org/en/arbo-portal/dengue-data-and-analysis/dengue-analysis-country')

        # Define wait outside the loop
        wait = WebDriverWait(driver, 30)

        # Wait for page to load completely
        logger.info("Waiting for page to load...")
        time.sleep(10)

        # First, ensure we're on the correct tab (Cases tab should be active by default)
        try:
            cases_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Cases")))
            cases_tab.click()
            logger.info("Clicked Cases tab")
            time.sleep(2)
        except Exception as e:
            logger.warning(f"Cases tab not found or already active: {e}")

        # Updated iframe locator - try multiple approaches
        iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008_NAC_ES_Indicadores_reporte_semanal.asp"

        logger.info("Looking for iframe...")
        # Approach 1: Direct iframe search
        try:
            iframe = wait.until(EC.presence_of_element_located((By.XPATH, f"//iframe[@src='{iframe_src}']")))
            logger.info("Found iframe using direct search")
        except Exception as e:
            logger.warning(f"Direct iframe search failed: {e}")
            # Approach 2: Look within active tab pane
            try:
                iframe = wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//div[contains(@class, 'tab-pane') and contains(@class, 'active')]//iframe[@src='{iframe_src}']")))
                logger.info("Found iframe within active tab pane")
            except Exception as e2:
                logger.warning(f"Tab pane iframe search failed: {e2}")
                # Approach 3: Look within paragraph structure
                iframe = wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//div[contains(@class, 'paragraph')]//iframe[@src='{iframe_src}']")))
                logger.info("Found iframe within paragraph structure")

        logger.info(f"Switching to iframe: {iframe}")
        driver.switch_to.frame(iframe)

        # Wait a bit for iframe to load
        time.sleep(5)

        # Get the nested iframe inside the first iframe (if it exists)
        try:
            iframe2 = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe")))
            driver.switch_to.frame(iframe2)
            logger.info("Switched to nested iframe")
            shadow_doc2 = driver.execute_script('return document')
        except Exception as e:
            logger.info(f"No nested iframe found, using current frame: {e}")
            # If no nested iframe, use current frame
            shadow_doc2 = driver.execute_script('return document')

        iframe_page_title = driver.title
        logger.info(f"Page title: {iframe_page_title}")

        # Rest of your iframe setup code...
        time.sleep(3)

        # Continue with the rest of your setup and download logic
        logger.info("Setting up year filter...")
        setup_year_filter(wait, shadow_doc2)

        logger.info("Setting up country filter...")
        setup_country_filter(wait, shadow_doc2)

        logger.info("Starting download process...")
        download_all_weeks(wait, shadow_doc2, default_dir, downloadPath, driver, year, today)

    except Exception as e:
        logger.error(f"An error occurred in iterate_weekly: {e}")
        logger.error(traceback.format_exc())

        if driver:
            try:
                logger.info("Taking screenshot for debugging...")
                screenshot_path = f"debug_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                driver.save_screenshot(screenshot_path)
                logger.info(f"Debug screenshot saved: {screenshot_path}")
            except Exception as screenshot_error:
                logger.error(f"Failed to take screenshot: {screenshot_error}")
        raise

    finally:
        if driver:
            try:
                driver.quit()
                logger.info("Driver closed successfully")
            except Exception as e:
                logger.error(f"Error closing driver: {e}")


def setup_year_filter(wait, shadow_doc2):
    """Set up year filter with better error handling"""
    try:
        # Find the year tab with more robust selectors
        try:
            year_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId16')))
        except:
            # Try alternative selectors if the ID changed
            try:
                year_tab = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//div[contains(@class, 'tabZone') and contains(text(), 'Year')]")))
            except:
                year_tab = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//div[contains(@id, 'tabZone') and contains(@id, '16')]")))

        # Find the dropdown button within the year tab
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = year_tab.find_element(*dd_locator)
        dd_open.click()
        logger.info("Opened year dropdown")

        # Clear all selections first
        if not click_tableau_element(shadow_doc2, "(All)", "year (uncheck All)"):
            logger.warning("Failed to uncheck 'All' for years - continuing anyway")
        time.sleep(2)

        # Select specific years using helper function
        years_to_select = ['2025', '2024', '2023']
        for year_select in years_to_select:
            if not click_tableau_element(shadow_doc2, year_select, "year"):
                logger.warning(f"Failed to select year {year_select}")
            time.sleep(0.5)

        # Close the dropdown menu
        try:
            dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
            dd_close.click()
        except:
            shadow_doc2.find_element(By.TAG_NAME, "body").click()

        logger.info("Year filter setup completed")
        time.sleep(3)

    except Exception as e:
        logger.error(f"Error setting up year filter: {e}")
        raise


def setup_country_filter(wait, shadow_doc2):
    """Set up country filter with better error handling"""
    try:
        # Select all countries (similar updates for region selection)
        try:
            region_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId13')))
        except:
            # Try alternative selectors
            region_tab = wait.until(EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@id, 'tabZone') and contains(@id, '13')]")))

        # Find the dropdown button within the region tab
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = region_tab.find_element(*dd_locator)
        dd_open.click()
        logger.info("Opened country dropdown")

        if not click_tableau_element(shadow_doc2, "(All)", "country"):
            logger.warning("Failed to select 'All' countries")
        time.sleep(3)

        try:
            dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
            dd_close.click()
        except:
            shadow_doc2.find_element(By.TAG_NAME, "body").click()

        logger.info("Country filter setup completed")

    except Exception as e:
        logger.error(f"Error setting up country filter: {e}")
        raise


def download_all_weeks(wait, shadow_doc2, default_dir, downloadPath, driver, year, today):
    """Iterate from week 53 to 1, downloading weekly data."""
    weeknum = 53
    consecutive_failures = 0
    max_consecutive_failures = 3

    while weeknum > 0 and consecutive_failures < max_consecutive_failures:
        logger.info(f"Processing Week Number: {weeknum}")

        try:
            actual_week = download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today)
            logger.info(f"Successfully downloaded week {actual_week}")
            consecutive_failures = 0  # Reset failure counter on success

        except Exception as e:
            consecutive_failures += 1
            logger.error(f"Error in download_and_rename for week {weeknum} (failure {consecutive_failures}): {e}")

            if consecutive_failures >= max_consecutive_failures:
                logger.error(f"Too many consecutive failures ({max_consecutive_failures}), stopping")
                break

        weeknum -= 1
        if weeknum == 0:
            logger.info("Reached week 0, exiting loop.")
            break

        # Navigate to previous week
        if not navigate_to_previous_week(driver, wait, weeknum):
            logger.error(f"Failed to navigate to week {weeknum}, stopping")
            break


def navigate_to_previous_week(driver, wait, target_week):
    """Navigate to the previous week with robust error handling"""
    try:
        # Step 1: Switch back to the main document (default content)
        driver.switch_to.default_content()
        logger.info("Switched to default content (main page).")
        time.sleep(1)

        # Step 2: Scroll the main page to get the iframe in view
        iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008_NAC_ES_Indicadores_reporte_semanal.asp"
        main_iframe_element = wait.until(EC.presence_of_element_located((By.XPATH, f"//iframe[@src='{iframe_src}']")))

        # Scroll the main iframe element into view
        driver.execute_script("arguments[0].scrollIntoView(false);", main_iframe_element)
        logger.info("Scrolled main iframe element into view.")
        time.sleep(1)

        # Step 3: Switch back into the nested iframe
        driver.switch_to.frame(main_iframe_element)
        iframe2_element = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe")))
        driver.switch_to.frame(iframe2_element)
        logger.info("Switched back to iframe2 for decrement button interaction.")
        time.sleep(1)

        # Step 4: Scroll to the top of the iframe
        driver.execute_script("window.scrollTo(0, 0);")
        driver.execute_script("document.documentElement.scrollTop = 0; document.body.scrollTop = 0;")
        time.sleep(2)

        # Step 5: Find and click the decrement button
        decrement_selectors = [
            "//div[contains(@class, 'dijitSliderDecrementIconH') and contains(@class, 'cpLeftArrowBlack')]",
            "//div[contains(@class, 'dijitSliderDecrementIconH')]",
            "//div[contains(@class, 'cpLeftArrowBlack')]",
        ]

        decrement_button = None
        for selector in decrement_selectors:
            try:
                decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                logger.info(f"Found decrement button with selector: {selector}")
                break
            except Exception as e:
                logger.warning(f"Selector '{selector}' failed: {e}")
                continue

        if decrement_button:
            try:
                # Scroll the element into view and click
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", decrement_button)
                time.sleep(1)
                decrement_button.click()
                logger.info(f"Successfully clicked decrement button for week {target_week}")
            except Exception as e:
                logger.warning(f"Direct click failed, trying JavaScript click: {e}")
                driver.execute_script("arguments[0].click();", decrement_button)
                logger.info(f"Clicked decrement button via JavaScript for week {target_week}")

            time.sleep(3)  # Allow Tableau to update
            return True
        else:
            logger.error(f"Could not find decrement button for week {target_week}")
            return False

    except Exception as e:
        logger.error(f"Error navigating to week {target_week}: {e}")
        logger.error(traceback.format_exc())

        # Take screenshot for debugging
        try:
            screenshot_path = f"debug_decrement_error_week_{target_week}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            driver.save_screenshot(screenshot_path)
            logger.info(f"Debug screenshot saved: {screenshot_path}")
        except:
            pass

        return False


def click_tableau_element(shadow_doc, option_text, element_type="year"):
    """
    Helper function to click Tableau filter elements with multiple fallback approaches
    """
    success = False

    # Approach 1: Click the input checkbox
    try:
        xpath = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]/preceding-sibling::input[@class="FICheckRadio"]'
        element = shadow_doc.find_element(By.XPATH, xpath)
        element.click()
        logger.info(f"Successfully clicked {element_type} '{option_text}' via checkbox")
        success = True
    except Exception as e:
        logger.warning(f"Checkbox click failed for {element_type} '{option_text}': {e}")

        # Approach 2: Click the text link
        try:
            xpath_alt = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]'
            element = shadow_doc.find_element(By.XPATH, xpath_alt)
            element.click()
            logger.info(f"Successfully clicked {element_type} '{option_text}' via text link")
            success = True
        except Exception as e2:
            logger.warning(f"Text link click failed for {element_type} '{option_text}': {e2}")

            # Approach 3: Click the fake checkbox div
            try:
                xpath_fake = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]/preceding-sibling::div[@class="fakeCheckBox"]'
                element = shadow_doc.find_element(By.XPATH, xpath_fake)
                element.click()
                logger.info(f"Successfully clicked {element_type} '{option_text}' via fake checkbox")
                success = True
            except Exception as e3:
                logger.warning(f"Fake checkbox click failed for {element_type} '{option_text}': {e3}")

    return success

# Run the main function
if __name__ == "__main__":
    try:
        logger.info("Starting PAHO crawler...")
        iterate_weekly()
        logger.info("PAHO crawler completed successfully")
    except Exception as e:
        logger.error(f"PAHO crawler failed: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)
