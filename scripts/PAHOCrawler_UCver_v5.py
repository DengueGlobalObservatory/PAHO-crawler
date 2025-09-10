import undetected_chromedriver as uc
#from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

#from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome.service import Service
import time
import os
from datetime import datetime
import subprocess
import re
import sys
import os
import time

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
        raise RuntimeError("Failed to get Chrome version") from e

# Get the major version of Chrome installed
chrome_version = get_chrome_version()

def move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension):
    got_file = False
    while not got_file:
        try:
            # Use glob to get the current file name
            currentFile = max([default_dir + "/" + f for f in os.listdir(default_dir)], key=os.path.getctime)

            # Ensure the file exists before proceeding
            if os.path.exists(currentFile):
                got_file = True
            else:
                raise FileNotFoundError("File not found. Retrying...")

        except Exception as e:
            print("File has not finished downloading")
            time.sleep(10)

    # Create new file name
    fileDestination = os.path.join(downloadPath, newFileName + fileExtension)

    # Move the file
    os.rename(currentFile, fileDestination)
    print(f"Moved file to {fileDestination}")


def download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today):
    """Download and rename the file for the given week number."""

    # Wait for the week number to update
    weeknum_div = wait.until(
        EC.presence_of_element_located((By.CLASS_NAME, "sliderText"))
    )
    weeknum = int(weeknum_div.text)

    # print(f"Week Number: {weeknum}")

    # Find and click the download button at the bottom of the dashboard
    download_button = wait.until(
        EC.element_to_be_clickable((By.ID, "download-ToolbarButton"))
    )
    download_button.click()

    time.sleep(5)

    # Find and click the crosstab button (in a pop up window)
    crosstab_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]'))
    )
    crosstab_button.click()

    time.sleep(5)

    # Find and select the CSV option
    #csv_div = wait.until(
    #    EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='radio'][value='csv']"))
    #)
    csv_div = shadow_doc2.find_element(By.CSS_SELECTOR, "input[type='radio'][value='csv']")
    #driver.execute_script("arguments[0].scrollIntoView();", csv_div)
    driver.execute_script("arguments[0].click();", csv_div)
    time.sleep(5)

    # Find and click the export button
    export_button = shadow_doc2.find_element(By.CSS_SELECTOR, '[data-tb-test-id="export-crosstab-export-Button"]')
    export_button.click()
    print("Downloading CSV file")
    time.sleep(5)

    # Use the move_to_download_folder function to move the downloaded file
    downloadPath = downloadPath
    default_dir = default_dir
    newFileName = f"PAHO_{year}_W{weeknum}_{today}"  # Base filename
    fileExtension = '.csv'  # File extension

    move_to_download_folder(default_dir, downloadPath, newFileName, fileExtension)


def iterate_weekly():
    year = "2023_2025"
    today = datetime.now().strftime('%Y%m%d%H%M')

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
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": default_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    })
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")

    # using undetected-chromedriver
    driver = uc.Chrome(headless=True, use_subprocess=False, options = chrome_options, version_main=chrome_version)
    driver.get('https://www.paho.org/en/arbo-portal/dengue-data-and-analysis/dengue-analysis-country')

    # Define wait outside the loop
    wait = WebDriverWait(driver, 30)  # Increased timeout
    try:
        # Wait for page to load completely
        time.sleep(5)

        # First, ensure we're on the correct tab (Cases tab should be active by default)
        try:
            cases_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Cases")))
            cases_tab.click()
            time.sleep(2)
        except:
            print("Cases tab not found or already active")

        # Updated iframe locator - try multiple approaches
        iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008en_NAC_Indicadores_reporte_semanal.asp"

        # Approach 1: Direct iframe search
        try:
            iframe = wait.until(EC.presence_of_element_located((By.XPATH, f"//iframe[@src='{iframe_src}']")))
        except:
            # Approach 2: Look within active tab pane
            try:
                iframe = wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//div[contains(@class, 'tab-pane') and contains(@class, 'active')]//iframe[@src='{iframe_src}']")))
            except:
                # Approach 3: Look within paragraph structure
                iframe = wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//div[contains(@class, 'paragraph')]//iframe[@src='{iframe_src}']")))

        print(f"Found iframe: {iframe}")
        driver.switch_to.frame(iframe)

        # Wait a bit for iframe to load
        time.sleep(3)

        # Get the nested iframe inside the first iframe (if it exists)
        try:
            iframe2 = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe")))
            driver.switch_to.frame(iframe2)
            shadow_doc2 = driver.execute_script('return document')
        except:
            # If no nested iframe, use current frame
            shadow_doc2 = driver.execute_script('return document')

        iframe_page_title = driver.title
        print(f"Page title: {iframe_page_title}")

        # Updated title check (title might have changed)
        expected_titles = [
            "PAHO/WHO Data - National Dengue fever cases",
            "Dengue: analysis by country - PAHO/WHO | Pan American Health Organization",
            "National Dengue fever cases"
        ]

        if not any(title in iframe_page_title for title in expected_titles):
            print(f"Unexpected page title: {iframe_page_title}")
            print("Continuing anyway...")

        time.sleep(3)

        # Find the year tab with more robust selectors
        try:
            year_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId13')))
        except:
            # Try alternative selectors if the ID changed
            try:
                year_tab = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//div[contains(@class, 'tabZone') and contains(text(), 'Year')]")))
            except:
                year_tab = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, "//div[contains(@id, 'tabZone') and contains(@id, '13')]")))

        # Find the dropdown button within the year tab


        # Find the dropdown button within the year tab
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = year_tab.find_element(*dd_locator)
        dd_open.click()
        time.sleep(3)  # Give more time for dropdown to open

        # Clear all selections first - keep your intentional double click
        print("Attempting to uncheck 'All' for years (first attempt)")
        if not click_tableau_element(shadow_doc2, "(All)", "year (uncheck All)"):
            print("Failed to uncheck 'All' for years - continuing anyway")
        time.sleep(5)

        print("Attempting to uncheck 'All' for years (second attempt)")
        if not click_tableau_element(shadow_doc2, "(All)", "year (uncheck All)"):
            print("Failed to uncheck 'All' for years - continuing anyway")
        time.sleep(5)

        # Select specific years using helper function
        years_to_select = ['2025', '2024', '2023']
        for year_select in years_to_select:
            print(f"Attempting to select year: {year_select}")
            success = click_tableau_element(shadow_doc2, year_select, "year")
            if not success:
                print(f"Failed to select year {year_select}")
                # Try alternative approach - click by visible text
                try:
                    alt_element = shadow_doc2.find_element(By.XPATH, f'//a[contains(text(), "{year_select}")]')
                    driver.execute_script("arguments[0].click();", alt_element)
                    print(f"Successfully selected year {year_select} via alternative method")
                except Exception as e:
                    print(f"Alternative selection also failed for {year_select}: {e}")
            else:
                print(f"Successfully selected year: {year_select}")
            time.sleep(2)

        # Wait longer before closing dropdown to ensure selections are registered
        time.sleep(5)

        # Close the dropdown menu
        try:
            dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
            dd_close.click()
            print("Closed dropdown via tab-glass")
        except:
            driver.find_element(By.TAG_NAME, "body").click()
            print("Closed dropdown via body click")

        time.sleep(5)  # Give time for the filter to apply

        # Debug: Check if years are actually selected by looking at the interface
        try:
            # Re-open dropdown to verify selections
            dd_open = year_tab.find_element(*dd_locator)
            dd_open.click()
            time.sleep(2)

            # Check which years appear selected
            selected_years = shadow_doc2.find_elements(By.XPATH, '//input[@class="FICheckRadio" and @checked]/../a')
            selected_year_texts = [elem.text for elem in selected_years]
            print(f"Currently selected years: {selected_year_texts}")

            # Close dropdown again
            try:
                dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
                dd_close.click()
            except:
                driver.find_element(By.TAG_NAME, "body").click()

            time.sleep(3)
        except Exception as e:
            print(f"Debug check failed: {e}")

        # Select all countries (similar updates for region selection)
        try:
            region_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId9')))
        except:
            # Try alternative selectors
            region_tab = wait.until(EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@id, 'tabZone') and contains(@id, '9')]")))

        # Find the dropdown button within the region tab
        dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
        dd_open = region_tab.find_element(*dd_locator)
        dd_open.click()

        if not click_tableau_element(shadow_doc2, "(All)", "country"):
            print("Failed to select 'All' countries")
        time.sleep(3)

        if not click_tableau_element(shadow_doc2, "(All)", "country"):
            print("Failed to select 'All' countries")
        time.sleep(3)

        try:
            dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
            dd_close.click()
        except:
            driver.find_element(By.TAG_NAME, "body").click()

        # Begin downloading all weeks
        download_all_weeks(wait, shadow_doc2, default_dir, downloadPath, driver, year, today)

    except Exception as e:
        print(f"An error occurred: {e}")
        print("Taking screenshot for debugging...")
        screenshot_path = f"debug_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        driver.save_screenshot(screenshot_path)
        print(f"Debug screenshot saved: {screenshot_path}")

    finally:
        driver.quit()


def download_all_weeks(wait, shadow_doc2, default_dir, downloadPath, driver, year, today):
    """Iterate from week 53 to 1, downloading weekly data."""
    weeknum = 53

    while weeknum > 0:
        print(f"Processing Week Number: {weeknum}")

        try:
            download_and_rename(wait, shadow_doc2, weeknum, default_dir, downloadPath, driver, year, today)
        except Exception as e:
            print(f"Error in download_and_rename for week {weeknum}: {e}")

        weeknum -= 1
        if weeknum == 0:
            print("Reached week 0, exiting loop.")
            break

        try:
            # Step 1: Switch back to the main document (default content)
            driver.switch_to.default_content()
            print("Switched to default content (main page).")
            time.sleep(1)

            # Step 2: Scroll the main page down to get the fixed header out of the way
            iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008en_NAC_Indicadores_reporte_semanal.asp"
            main_iframe_element = wait.until(EC.presence_of_element_located((By.XPATH, f"//iframe[@src='{iframe_src}']")))

            # Scroll the main iframe element into view
            driver.execute_script("arguments[0].scrollIntoView(false);", main_iframe_element)
            print("Scrolled main iframe element into view (to clear header).")
            time.sleep(1)

            # Step 3: Switch back into the nested iframe (iframe2)
            driver.switch_to.frame(main_iframe_element)
            iframe2_element = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe")))
            driver.switch_to.frame(iframe2_element)
            print("Switched back to iframe2 for decrement button interaction.")
            time.sleep(1)

            # NEW: Step 4: Scroll to the top of the iframe to ensure decrement button is visible
            print("Scrolling to top of iframe to bring decrement button into view...")
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)

            # Alternative: You can also try scrolling the iframe content specifically
            # This ensures we're at the very top of the iframe content
            driver.execute_script("document.documentElement.scrollTop = 0; document.body.scrollTop = 0;")
            time.sleep(1)

            # Step 4: Scroll to the top of the iframe to ensure decrement button is visible
            print("Scrolling to top of iframe to bring decrement button into view...")
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)

            # Alternative: You can also try scrolling the iframe content specifically
            # This ensures we're at the very top of the iframe content
            driver.execute_script("document.documentElement.scrollTop = 0; document.body.scrollTop = 0;")
            time.sleep(1)

            # Step 5: Find and click the decrement button within iframe2
            decrement_selectors = [
                "//div[contains(@class, 'dijitSliderDecrementIconH') and contains(@class, 'cpLeftArrowBlack')]",
                "//div[contains(@class, 'dijitSliderDecrementIconH')]",
                "//div[contains(@class, 'cpLeftArrowBlack')]",
                "//div[contains(@class, 'dijitSliderDecrementIconH') and contains(@class, 'cpLeftArrowBlack')]/span[contains(@class, 'dijitSliderButtonInner')]"
            ]

            decrement_button = None
            for selector in decrement_selectors:
                try:
                    decrement_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    print(f"Found decrement button in iframe2 with selector: {selector}")
                    break
                except Exception as e:
                    print(f"Iframe2 selector '{selector}' failed to locate decrement button: {e}")
                    continue

            if decrement_button:
                try:
                    # Scroll the specific element into view before clicking
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", decrement_button)
                    time.sleep(1)

                    # Attempt direct click first
                    decrement_button.click()
                    print(f"Successfully clicked decrement button directly for week {weeknum}.")
                except Exception as e:
                    print(f"Direct click failed for decrement button for week {weeknum}, attempting JavaScript click: {e}")
                    # Use JavaScript click as fallback
                    driver.execute_script("arguments[0].click();", decrement_button)
                    print(f"Clicked decrement button via JavaScript for week {weeknum}.")
                time.sleep(3) # Allow Tableau to update view
            else:
                print(f"Could not find decrement button in iframe2 for week {weeknum} – exiting loop.")
                break

        except Exception as e:
            print(f"An error occurred while trying to decrement week {weeknum}: {e}")
            screenshot_path = f"debug_decrement_error_week_{weeknum}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            driver.save_screenshot(screenshot_path)
            print(f"Debug screenshot saved: {screenshot_path}")
            break

# Additional helper function for debugging
def debug_page_structure(driver, wait):
    """Helper function to debug the current page structure"""
    try:
        # Print all iframes on the page
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"Found {len(iframes)} iframes:")
        for i, iframe in enumerate(iframes):
            src = iframe.get_attribute("src")
            print(f"  Iframe {i}: {src}")

        # Print tab structure
        tabs = driver.find_elements(By.XPATH, "//div[contains(@class, 'tab-pane')]")
        print(f"Found {len(tabs)} tab panes:")
        for i, tab in enumerate(tabs):
            tab_id = tab.get_attribute("id")
            is_active = "active" in tab.get_attribute("class")
            print(f"  Tab {i}: {tab_id} (active: {is_active})")

    except Exception as e:
        print(f"Debug error: {e}")

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
        print(f"Successfully clicked {element_type} '{option_text}' via checkbox")
        success = True
    except Exception as e:
        print(f"Checkbox click failed for {element_type} '{option_text}': {e}")

        # Approach 2: Click the text link
        try:
            xpath_alt = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]'
            element = shadow_doc.find_element(By.XPATH, xpath_alt)
            element.click()
            print(f"Successfully clicked {element_type} '{option_text}' via text link")
            success = True
        except Exception as e2:
            print(f"Text link click failed for {element_type} '{option_text}': {e2}")

            # Approach 3: Click the fake checkbox div
            try:
                xpath_fake = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]/preceding-sibling::div[@class="fakeCheckBox"]'
                element = shadow_doc.find_element(By.XPATH, xpath_fake)
                element.click()
                print(f"Successfully clicked {element_type} '{option_text}' via fake checkbox")
            except Exception as e3:
                print(f"Fake checkbox click failed for {element_type} '{option_text}': {e3}")

    return success

def run_diagnostics():
    """
    Standalone diagnostics function to verify iframe sources and tabZone IDs
    Returns True if all diagnostics pass, False otherwise
    """
    print("\n" + "="*60)
    print("RUNNING STANDALONE DIAGNOSTICS")
    print("="*60)

    # Set up Chrome driver for diagnostics
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")

    driver = uc.Chrome(headless=True, use_subprocess=False, options=chrome_options, version_main=chrome_version)
    wait = WebDriverWait(driver, 30)

    diagnostics_passed = True
    issues_found = []

    try:
        print("Loading PAHO website...")
        driver.get('https://www.paho.org/en/arbo-portal/dengue-data-and-analysis/dengue-analysis-country')
        time.sleep(5)

        # Click Cases tab
        try:
            cases_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Cases")))
            cases_tab.click()
            time.sleep(2)
            print("✓ Successfully clicked Cases tab")
        except:
            print("⚠ Cases tab not found or already active")

        # 1. IFRAME DIAGNOSTICS
        print("\n1. IFRAME DIAGNOSTICS:")
        print("-" * 30)

        expected_iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008en_NAC_Indicadores_reporte_semanal.asp"
        print(f"Expected iframe src: {expected_iframe_src}")

        try:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            print(f"Total iframes found: {len(iframes)}")

            iframe_found = False
            for i, iframe in enumerate(iframes):
                src = iframe.get_attribute("src") or "No src attribute"
                name = iframe.get_attribute("name") or "No name"
                id_attr = iframe.get_attribute("id") or "No id"
                print(f"  Iframe {i+1}:")
                print(f"    src: {src}")
                print(f"    name: {name}")
                print(f"    id: {id_attr}")

                if src == expected_iframe_src:
                    print(f"    ✓ MATCHES EXPECTED SRC")
                    iframe_found = True
                else:
                    print(f"    ✗ Does not match expected")

            if not iframe_found:
                diagnostics_passed = False
                issues_found.append("Expected iframe src not found")
                print(f"\n✗ CRITICAL: Expected iframe not found!")
            else:
                print(f"\n✓ Expected iframe found successfully")

        except Exception as e:
            diagnostics_passed = False
            issues_found.append(f"Error checking iframes: {e}")
            print(f"✗ Error checking iframes: {e}")

        # 2. SWITCH TO IFRAME AND CHECK TABZONES
        print("\n2. TABZONE DIAGNOSTICS:")
        print("-" * 30)

        try:
            # Switch to the main iframe
            iframe = driver.find_element(By.XPATH, f"//iframe[@src='{expected_iframe_src}']")
            driver.switch_to.frame(iframe)
            print("✓ Switched to main iframe")

            # Check for nested iframe
            nested_iframe_found = False
            try:
                iframe2 = driver.find_element(By.XPATH, "//body/iframe")
                driver.switch_to.frame(iframe2)
                print("✓ Switched to nested iframe")
                nested_iframe_found = True
            except:
                print("⚠ No nested iframe found, staying in current frame")

            # Check page title
            iframe_page_title = driver.title
            print(f"Iframe page title: {iframe_page_title}")

            expected_titles = [
                "PAHO/WHO Data - National Dengue fever cases",
                "Dengue: analysis by country - PAHO/WHO | Pan American Health Organization",
                "National Dengue fever cases"
            ]

            title_match = any(title in iframe_page_title for title in expected_titles)
            if title_match:
                print("✓ Page title matches expected patterns")
            else:
                print("⚠ Page title doesn't match expected patterns")
                issues_found.append(f"Unexpected page title: {iframe_page_title}")

            # Check for expected tabZone IDs
            expected_tabzones = {
                'tabZoneId13': 'Year tab',
                'tabZoneId9': 'Region/Country tab'
            }

            tabzones_found = 0
            for tabzone_id, description in expected_tabzones.items():
                print(f"\nChecking {description} (ID: {tabzone_id}):")
                try:
                    element = driver.find_element(By.ID, tabzone_id)
                    print(f"  ✓ Found: {tabzone_id}")
                    print(f"  Element tag: {element.tag_name}")
                    print(f"  Element classes: {element.get_attribute('class')}")

                    # Check for dropdown button within this tabzone
                    try:
                        dropdown_btn = element.find_element(By.CSS_SELECTOR, 'span.tabComboBoxButton')
                        print(f"  ✓ Dropdown button found in {tabzone_id}")
                        tabzones_found += 1
                    except:
                        print(f"  ✗ No dropdown button found in {tabzone_id}")
                        issues_found.append(f"No dropdown button in {tabzone_id}")

                except Exception as e:
                    print(f"  ✗ Not found: {tabzone_id} - {e}")
                    issues_found.append(f"TabZone not found: {tabzone_id}")

                    # Try to find similar elements
                    print(f"  Searching for similar tabZone elements...")
                    try:
                        similar_elements = driver.find_elements(By.XPATH, "//div[contains(@id, 'tabZone')]")
                        print(f"  Found {len(similar_elements)} elements with 'tabZone' in ID:")
                        for elem in similar_elements[:5]:  # Show first 5
                            elem_id = elem.get_attribute('id')
                            elem_text = elem.text[:30] if elem.text else "No text"
                            print(f"    - ID: {elem_id}, Text: {elem_text}...")
                    except Exception as e2:
                        print(f"  Error searching for similar elements: {e2}")

            if tabzones_found < 2:
                diagnostics_passed = False
                print(f"\n✗ CRITICAL: Only {tabzones_found}/2 required tabZones found with dropdown buttons")
            else:
                print(f"\n✓ All required tabZones found successfully")

            # 3. CHECK OTHER CRITICAL ELEMENTS
            print(f"\n3. OTHER CRITICAL ELEMENTS:")
            print("-" * 30)

            # Check for slider elements (week selector)
            try:
                slider_text = driver.find_element(By.CLASS_NAME, "sliderText")
                print(f"✓ Week slider found - Current value: {slider_text.text}")
            except Exception as e:
                print(f"✗ Week slider not found: {e}")
                issues_found.append("Week slider not found")
                diagnostics_passed = False

            # Check for decrement button
            decrement_found = False
            decrement_selectors = [
                "//div[contains(@class, 'dijitSliderDecrementIconH') and contains(@class, 'cpLeftArrowBlack')]",
                "//div[contains(@class, 'dijitSliderDecrementIconH')]",
                "//div[contains(@class, 'cpLeftArrowBlack')]"
            ]

            for i, selector in enumerate(decrement_selectors):
                try:
                    element = driver.find_element(By.XPATH, selector)
                    print(f"✓ Decrement button found with selector {i+1}")
                    decrement_found = True
                    break
                except:
                    continue

            if not decrement_found:
                print(f"✗ Decrement button not found with any selector")
                issues_found.append("Decrement button not found")
                diagnostics_passed = False

            # Check for download button
            try:
                download_btn = driver.find_element(By.ID, "download-ToolbarButton")
                print(f"✓ Download button found")
            except Exception as e:
                print(f"✗ Download button not found: {e}")
                issues_found.append("Download button not found")
                diagnostics_passed = False

        except Exception as e:
            diagnostics_passed = False
            issues_found.append(f"Error during iframe/tabZone diagnostics: {e}")
            print(f"✗ Error during iframe/tabZone diagnostics: {e}")

    except Exception as e:
        diagnostics_passed = False
        issues_found.append(f"General error: {e}")
        print(f"✗ General error: {e}")

    finally:
        driver.quit()

    # SUMMARY
    print("\n" + "="*60)
    print("DIAGNOSTICS SUMMARY")
    print("="*60)

    if diagnostics_passed:
        print("🎉 ALL DIAGNOSTICS PASSED!")
        print("✓ The script should work correctly")
        print("✓ All required elements found")
        print("✓ Ready to run iterate_weekly()")
    else:
        print("❌ DIAGNOSTICS FAILED!")
        print("✗ Issues found that may prevent the script from working:")
        for issue in issues_found:
            print(f"  - {issue}")
        print("\n⚠ Fix these issues before running the main script")

    print("="*60 + "\n")
    return diagnostics_passed


def main():
    """
    Main function that runs diagnostics first, then the main script if diagnostics pass
    """
    print("Starting PAHO Dengue Data Scraper...")
    print("Step 1: Running diagnostics...")

    if run_diagnostics():
        print("Step 2: Diagnostics passed! Starting data download automatically...")
        iterate_weekly()
    else:
        print("Step 2: Diagnostics failed! Please check the issues above.")
        print("Script will NOT run until diagnostics pass.")


# Run the main function
if __name__ == "__main__":
    main()
