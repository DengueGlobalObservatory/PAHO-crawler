"""
PAHO Dengue Data Scraper - Visual Debug Script
==============================================

This script runs the same logic as the main script but:
1. Shows the browser window (not headless)
2. Pauses for visual confirmation at each step
3. Highlights elements being interacted with
4. Provides clear feedback about what's working/failing

Use this script to debug issues before running the main automated script.
"""

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import os
from datetime import datetime
import subprocess
import re
import sys

# Function to get the installed Chrome version (same as main script)
def get_chrome_version():
    try:
        if sys.platform == "win32":
            output = subprocess.check_output(
                r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
                shell=True,
                text=True
            )
            version = re.search(r'\s+version\s+REG_SZ\s+(\d+)\.', output)
        else:
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

def click_tableau_element(shadow_doc, option_text, element_type="year"):
    """Helper function to click Tableau filter elements (same as main script)"""
    success = False

    # Approach 1: Click the input checkbox
    try:
        xpath = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]/preceding-sibling::input[@class="FICheckRadio"]'
        element = shadow_doc.find_element(By.XPATH, xpath)
        element.click()
        print(f"✅ Successfully clicked {element_type} '{option_text}' via checkbox")
        success = True
    except Exception as e:
        print(f"❌ Checkbox click failed for {element_type} '{option_text}': {e}")

        # Approach 2: Click the text link
        try:
            xpath_alt = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]'
            element = shadow_doc.find_element(By.XPATH, xpath_alt)
            element.click()
            print(f"✅ Successfully clicked {element_type} '{option_text}' via text link")
            success = True
        except Exception as e2:
            print(f"❌ Text link click failed for {element_type} '{option_text}': {e2}")

            # Approach 3: Click the fake checkbox div
            try:
                xpath_fake = f'//div[@class="facetOverflow"]/a[@title="{option_text}" and text()="{option_text}"]/preceding-sibling::div[@class="fakeCheckBox"]'
                element = shadow_doc.find_element(By.XPATH, xpath_fake)
                element.click()
                print(f"✅ Successfully clicked {element_type} '{option_text}' via fake checkbox")
                success = True
            except Exception as e3:
                print(f"❌ Fake checkbox click failed for {element_type} '{option_text}': {e3}")

    return success

def visual_test_main_script_flow():
    """Visual test that follows the exact same flow as the main script"""

    print("=" * 60)
    print("PAHO DENGUE DATA SCRAPER - VISUAL DEBUG MODE")
    print("=" * 60)
    print("This script follows the exact same steps as the main script")
    print("but runs in visual mode for debugging purposes.\n")

    chrome_version = get_chrome_version()
    print(f"🚀 Chrome version: {chrome_version}")

    # Create temp downloads directory (same as main script)
    default_dir = os.path.join(os.getcwd(), "temp_downloads")
    os.makedirs(default_dir, exist_ok=True)
    print(f"📁 Created temp downloads directory: {default_dir}")

    # Chrome options setup (same as main script but WITHOUT headless)
    chrome_options = uc.ChromeOptions()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": default_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    })
    # chrome_options.add_argument("--headless=new")  # DISABLED FOR VISUAL MODE
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")

    driver = None
    try:
        print("🔧 Initializing Chrome in VISUAL mode...")
        driver = uc.Chrome(headless=False, use_subprocess=False, options=chrome_options, version_main=chrome_version)
        print("✅ Chrome initialized successfully")

        print("🌐 Loading PAHO website...")
        driver.get('https://www.paho.org/en/arbo-portal/dengue-data-and-analysis/dengue-analysis-country')
        print("✅ Page loaded successfully")

        wait = WebDriverWait(driver, 30)

        # Wait for page to load completely (same as main script)
        print("⏳ Waiting 5 seconds for page to fully load...")
        time.sleep(5)

        print(f"🔍 Current URL: {driver.current_url}")
        print(f"📄 Page title: {driver.title}")

        # First, ensure we're on the correct tab (same as main script)
        try:
            print("🔍 Looking for Cases tab...")
            cases_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Cases")))

            # Highlight the Cases tab
            driver.execute_script("arguments[0].style.border='3px solid red'", cases_tab)
            input("👁️ VISUAL CHECK: Is the 'Cases' tab highlighted in red? Press Enter to click it...")

            cases_tab.click()
            time.sleep(2)
            print("✅ Successfully clicked Cases tab")
        except:
            print("ℹ️ Cases tab not found or already active")

        # Updated iframe locator (same as main script)
        iframe_src = "https://ais.paho.org/ArboPortal/DENG/1008en_NAC_Indicadores_reporte_semanal.asp"
        print(f"🔍 Looking for iframe with src: {iframe_src}")

        # Approach 1: Direct iframe search (same as main script)
        iframe = None
        try:
            iframe = wait.until(EC.presence_of_element_located((By.XPATH, f"//iframe[@src='{iframe_src}']")))
            print("✅ Found iframe using direct search")
        except:
            try:
                iframe = wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//div[contains(@class, 'tab-pane') and contains(@class, 'active')]//iframe[@src='{iframe_src}']")))
                print("✅ Found iframe using tab-pane search")
            except:
                iframe = wait.until(EC.presence_of_element_located(
                    (By.XPATH, f"//div[contains(@class, 'paragraph')]//iframe[@src='{iframe_src}']")))
                print("✅ Found iframe using paragraph search")

        if iframe:
            # Highlight the iframe
            driver.execute_script("arguments[0].style.border='5px solid green'", iframe)
            input("👁️ VISUAL CHECK: Is the iframe highlighted in green? Press Enter to switch to it...")

            print(f"Found iframe: {iframe}")
            driver.switch_to.frame(iframe)
            print("✅ Switched to main iframe")

        # Wait a bit for iframe to load (same as main script)
        time.sleep(3)

        # Get the nested iframe inside the first iframe (same as main script)
        try:
            print("🔍 Looking for nested iframe...")
            iframe2 = wait.until(EC.presence_of_element_located((By.XPATH, "//body/iframe")))

            # Highlight nested iframe
            driver.execute_script("arguments[0].style.border='5px solid blue'", iframe2)
            input("👁️ VISUAL CHECK: Is the nested iframe highlighted in blue? Press Enter to switch to it...")

            driver.switch_to.frame(iframe2)
            shadow_doc2 = driver.execute_script('return document')
            print("✅ Switched to nested iframe")
        except:
            print("ℹ️ No nested iframe found, using current frame")
            shadow_doc2 = driver.execute_script('return document')

        iframe_page_title = driver.title
        print(f"📄 Final page title: {iframe_page_title}")

        # Title check (same as main script)
        expected_titles = [
            "PAHO/WHO Data - National Dengue fever cases",
            "Dengue: analysis by country - PAHO/WHO | Pan American Health Organization",
            "National Dengue fever cases"
        ]

        if not any(title in iframe_page_title for title in expected_titles):
            print(f"⚠️ Unexpected page title: {iframe_page_title}")
            print("Continuing anyway...")

        time.sleep(3)

        input("👁️ VISUAL CHECK: Can you see the Tableau dashboard with filters and charts? Press Enter to continue...")

        # YEAR FILTER TESTING (same as main script)
        print("\n" + "="*50)
        print("TESTING YEAR FILTER (tabZoneId13)")
        print("="*50)

        try:
            year_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId13')))
            print("✅ Found year filter tab (tabZoneId13)")

            # Highlight the year tab
            driver.execute_script("arguments[0].style.border='3px solid red'", year_tab)
            input("👁️ VISUAL CHECK: Is the year filter highlighted in red? Press Enter to continue...")

            # Find the dropdown button within the year tab (same as main script)
            dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
            dd_open = year_tab.find_element(*dd_locator)

            # Highlight dropdown button
            driver.execute_script("arguments[0].style.border='2px solid orange'", dd_open)
            input("👁️ VISUAL CHECK: Is the year dropdown button highlighted in orange? Press Enter to click it...")

            dd_open.click()
            time.sleep(3)
            print("✅ Opened year dropdown")

            input("👁️ VISUAL CHECK: Did the year dropdown menu open? You should see year options. Press Enter to continue...")

            # Clear all selections first (same as main script - intentional double click)
            print("Attempting to uncheck 'All' for years (first attempt)")
            if not click_tableau_element(shadow_doc2, "(All)", "year (uncheck All)"):
                print("Failed to uncheck 'All' for years - continuing anyway")
            time.sleep(5)

            print("Attempting to uncheck 'All' for years (second attempt)")
            if not click_tableau_element(shadow_doc2, "(All)", "year (uncheck All)"):
                print("Failed to uncheck 'All' for years - continuing anyway")
            time.sleep(5)

            input("👁️ VISUAL CHECK: Are all years now unchecked? Press Enter to select specific years...")

            # Select specific years (same as main script)
            years_to_select = ['2025', '2024', '2023']
            for year_select in years_to_select:
                print(f"Attempting to select year: {year_select}")
                success = click_tableau_element(shadow_doc2, year_select, "year")
                if not success:
                    print(f"Failed to select year {year_select}")
                else:
                    print(f"Successfully selected year: {year_select}")
                time.sleep(2)

            input("👁️ VISUAL CHECK: Are years 2025, 2024, and 2023 now checked? Press Enter to close dropdown...")

            # Close the dropdown menu (same as main script)
            try:
                dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
                dd_close.click()
                print("✅ Closed dropdown via tab-glass")
            except:
                driver.find_element(By.TAG_NAME, "body").click()
                print("✅ Closed dropdown via body click")

            time.sleep(5)

        except Exception as e:
            print(f"❌ Error with year filter: {e}")
            input("Press Enter to continue to next test...")

        # COUNTRY FILTER TESTING (same as main script)
        print("\n" + "="*50)
        print("TESTING COUNTRY FILTER (tabZoneId9)")
        print("="*50)

        try:
            region_tab = wait.until(EC.visibility_of_element_located((By.ID, 'tabZoneId9')))
            print("✅ Found country filter tab (tabZoneId9)")

            # Highlight the region tab
            driver.execute_script("arguments[0].style.border='3px solid blue'", region_tab)
            input("👁️ VISUAL CHECK: Is the country filter highlighted in blue? Press Enter to continue...")

            # Find the dropdown button within the region tab (same as main script)
            dd_locator = (By.CSS_SELECTOR, 'span.tabComboBoxButton')
            dd_open = region_tab.find_element(*dd_locator)

            # Highlight dropdown button
            driver.execute_script("arguments[0].style.border='2px solid orange'", dd_open)
            input("👁️ VISUAL CHECK: Is the country dropdown button highlighted in orange? Press Enter to click it...")

            dd_open.click()
            time.sleep(3)
            print("✅ Opened country dropdown")

            input("👁️ VISUAL CHECK: Did the country dropdown menu open? You should see country options. Press Enter to continue...")

            # Select all countries (same as main script)
            if not click_tableau_element(shadow_doc2, "(All)", "country"):
                print("Failed to select 'All' countries")
            time.sleep(3)

            if not click_tableau_element(shadow_doc2, "(All)", "country"):
                print("Failed to select 'All' countries")
            time.sleep(3)

            input("👁️ VISUAL CHECK: Is 'All' countries selected? Press Enter to close dropdown...")

            # Close dropdown (same as main script)
            try:
                dd_close = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tab-glass")))
                dd_close.click()
                print("✅ Closed country dropdown")
            except:
                driver.find_element(By.TAG_NAME, "body").click()
                print("✅ Closed country dropdown (fallback)")

        except Exception as e:
            print(f"❌ Error with country filter: {e}")
            input("Press Enter to continue to next test...")

        # WEEK NAVIGATION TESTING (same as main script)
        print("\n" + "="*50)
        print("TESTING WEEK NAVIGATION")
        print("="*50)

        try:
            # Wait for the week number to update (same as main script)
            weeknum_div = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "sliderText")))
            current_weeknum = int(weeknum_div.text)
            print(f"✅ Found current week: {current_weeknum}")

            # Highlight week display
            driver.execute_script("arguments[0].style.border='3px solid green'", weeknum_div)
            input(f"👁️ VISUAL CHECK: Is the week number ({current_weeknum}) highlighted in green? Press Enter to test decrement...")

            # Test decrement button (same selectors as main script)
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
                    print(f"✅ Found decrement button with selector: {selector}")
                    break
                except Exception as e:
                    print(f"❌ Selector '{selector}' failed: {e}")
                    continue

            if decrement_button:
                # Highlight decrement button
                driver.execute_script("arguments[0].style.border='3px solid orange'", decrement_button)
                input("👁️ VISUAL CHECK: Is the decrement button (left arrow) highlighted in orange? Press Enter to click it...")

                try:
                    # Same click logic as main script
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", decrement_button)
                    time.sleep(1)
                    decrement_button.click()
                    print("✅ Successfully clicked decrement button")
                except Exception as e:
                    print(f"Direct click failed, trying JavaScript click: {e}")
                    driver.execute_script("arguments[0].click();", decrement_button)
                    print("✅ Clicked decrement button via JavaScript")

                time.sleep(3)

                # Check if week changed
                new_weeknum = int(weeknum_div.text)
                print(f"📊 Week changed from {current_weeknum} to {new_weeknum}")

                if new_weeknum == current_weeknum - 1:
                    print("✅ Week decremented correctly!")
                else:
                    print("⚠️ Week didn't decrement as expected")
            else:
                print("❌ Could not find decrement button")

        except Exception as e:
            print(f"❌ Error testing week navigation: {e}")

        # DOWNLOAD BUTTON TESTING (same as main script)
        print("\n" + "="*50)
        print("TESTING DOWNLOAD FUNCTIONALITY")
        print("="*50)

        try:
            # Find and highlight download button (same as main script)
            download_button = wait.until(EC.element_to_be_clickable((By.ID, "download-ToolbarButton")))
            print("✅ Found download button")

            # Highlight download button
            driver.execute_script("arguments[0].style.border='3px solid purple'", download_button)
            input("👁️ VISUAL CHECK: Is the download button highlighted in purple? Press Enter to test click...")

            download_button.click()
            print("✅ Clicked download button")
            time.sleep(5)

            input("👁️ VISUAL CHECK: Did a download dialog appear? Press Enter to continue...")

            # Look for crosstab button (same as main script)
            try:
                crosstab_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-tb-test-id="DownloadCrosstab-Button"]')))

                # Highlight crosstab button
                driver.execute_script("arguments[0].style.border='3px solid cyan'", crosstab_button)
                input("👁️ VISUAL CHECK: Is the crosstab button highlighted in cyan? Press Enter to click...")

                crosstab_button.click()
                print("✅ Clicked crosstab button")
                time.sleep(5)

                input("👁️ VISUAL CHECK: Did the export options appear? Press Enter to continue...")

                # Test CSV selection (same as main script)
                try:
                    csv_div = shadow_doc2.find_element(By.CSS_SELECTOR, "input[type='radio'][value='csv']")
                    driver.execute_script("arguments[0].style.border='3px solid yellow'", csv_div)
                    input("👁️ VISUAL CHECK: Is the CSV radio button highlighted in yellow? Press Enter to select...")

                    driver.execute_script("arguments[0].click();", csv_div)
                    print("✅ Selected CSV option")
                    time.sleep(5)

                    input("👁️ VISUAL CHECK: Is CSV option now selected? You should see it's checked. Press Enter to finish...")

                except Exception as e:
                    print(f"❌ Error selecting CSV: {e}")

            except Exception as e:
                print(f"❌ Error with crosstab button: {e}")

        except Exception as e:
            print(f"❌ Error testing download: {e}")

        print("\n" + "="*60)
        print("🎉 VISUAL TESTING COMPLETE!")
        print("="*60)
        print("Summary of what was tested:")
        print("✓ Page loading and iframe navigation")
        print("✓ Year filter (tabZoneId13) interaction")
        print("✓ Country filter (tabZoneId9) interaction")
        print("✓ Week navigation (decrement button)")
        print("✓ Download functionality")
        print("\nIf all steps worked visually, the main script should work in headless mode.")
        print("If any step failed, you now know which part needs debugging.")

        input("👁️ FINAL CHECK: Review the current state of the dashboard. Press Enter to close browser...")

    except Exception as e:
        print(f"❌ Test error: {e}")
        print("Taking screenshot for debugging...")
        screenshot_path = f"debug_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved: {screenshot_path}")
        input("Press Enter to close browser...")

    finally:
        if driver:
            try:
                driver.quit()
                print("✅ Browser closed successfully")
            except:
                pass

if __name__ == "__main__":
    print("Starting visual debug mode...")
    print("This will show you exactly what the main script does, step by step.")
    print("Use this to identify any issues before running the automated version.\n")

    visual_test_main_script_flow()
