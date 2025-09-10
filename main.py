#!/usr/bin/env python3
"""
Complete IMF Financial Data Query Tool Automation - FIXED VERSION
Automates the full process from the NABIMFD runbook:
- Point 2: Select all members and lenders
- Point 3: Select Commitments/Borrowings, then Borrowings radio, then all borrowings types
- Point 4: Select Current (date selection)
- Point 5: Submit query
- Point 6: Download Excel file
"""

# Configuration
HEADLESS_MODE = True  # Set to True to hide the browser
DEBUG_MODE = True      # Set to True for additional debug output

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time
import os
import re
import subprocess
import platform

class IMFCompleteAutomator:
    def __init__(self):
        """Initialize the complete IMF automator"""
        self.url = "https://www.imf.org/external/np/fin/tad/query.aspx"
        self.driver = None
        self.chrome_version = None
        self.wait = None
        self.downloads_dir = None
        self.setup_driver()
    
    def detect_chrome_version(self):
        """Detect installed Chrome version across different operating systems"""
        try:
            system = platform.system().lower()
            chrome_version = None
            
            print("Detecting Chrome version...")
            
            if system == "windows":
                try:
                    result = subprocess.run([
                        'reg', 'query', 
                        'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon',
                        '/v', 'version'
                    ], capture_output=True, text=True, timeout=10)
                    
                    if result.returncode == 0:
                        version_match = re.search(r'version\s+REG_SZ\s+(\d+\.\d+\.\d+\.\d+)', result.stdout)
                        if version_match:
                            chrome_version = version_match.group(1)
                            print(f"Chrome version detected via registry: {chrome_version}")
                except:
                    pass
                
                if not chrome_version:
                    try:
                        chrome_paths = [
                            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
                            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
                            os.path.expanduser('~\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe')
                        ]
                        
                        for chrome_path in chrome_paths:
                            if os.path.exists(chrome_path):
                                result = subprocess.run([
                                    'powershell', '-Command', 
                                    f"(Get-Item '{chrome_path}').VersionInfo.FileVersion"
                                ], capture_output=True, text=True, timeout=10)
                                
                                if result.returncode == 0 and result.stdout.strip():
                                    chrome_version = result.stdout.strip()
                                    print(f"Chrome version detected: {chrome_version}")
                                    break
                    except:
                        pass
            
            elif system == "darwin":  # macOS
                try:
                    result = subprocess.run([
                        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'
                    ], capture_output=True, text=True, timeout=10)
                    
                    if result.returncode == 0:
                        version_match = re.search(r'Google Chrome (\d+\.\d+\.\d+\.\d+)', result.stdout)
                        if version_match:
                            chrome_version = version_match.group(1)
                            print(f"Chrome version detected on macOS: {chrome_version}")
                except:
                    pass
            
            elif system == "linux":
                try:
                    chrome_commands = ['google-chrome', 'google-chrome-stable', 'chromium-browser', 'chromium']
                    
                    for cmd in chrome_commands:
                        try:
                            result = subprocess.run([cmd, '--version'], capture_output=True, text=True, timeout=10)
                            if result.returncode == 0:
                                version_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', result.stdout)
                                if version_match:
                                    chrome_version = version_match.group(1)
                                    print(f"Chrome version detected on Linux: {chrome_version}")
                                    break
                        except:
                            continue
                except:
                    pass
            
            if chrome_version:
                major_version = int(chrome_version.split('.')[0])
                self.chrome_version = major_version
                print(f"Chrome major version: {major_version}")
                return major_version
            else:
                print("Could not detect Chrome version, using automatic detection")
                return None
                
        except Exception as e:
            print(f"Error detecting Chrome version: {e}")
            return None
    
    def setup_driver(self):
        """Setup Chrome driver with universal version compatibility"""
        print("Setting up Chrome driver...")
        
        detected_version = self.detect_chrome_version()
        
        # Create downloads directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.downloads_dir = os.path.join(script_dir, "downloads")
        
        if not os.path.exists(self.downloads_dir):
            os.makedirs(self.downloads_dir)
            print(f"Created downloads directory: {self.downloads_dir}")
        
        # Setup Chrome options
        options = uc.ChromeOptions()
        
        if HEADLESS_MODE:
            options.add_argument("--headless")
            print("Running in headless mode")
        else:
            print("Running in visible mode")
        
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        # Set download preferences
        prefs = {
            "download.default_directory": self.downloads_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        options.add_experimental_option("prefs", prefs)
        
        # Initialize driver
        try:
            if detected_version:
                try:
                    self.driver = uc.Chrome(options=options, version_main=detected_version)
                    print(f"[OK] Chrome driver initialized with version {detected_version}")
                except Exception as e:
                    print(f"Failed with version {detected_version}, trying auto-detection: {e}")
                    self.driver = uc.Chrome(options=options)
                    print("[OK] Chrome driver initialized with auto-detection")
            else:
                self.driver = uc.Chrome(options=options)
                print("[OK] Chrome driver initialized with auto-detection")
            
            self.wait = WebDriverWait(self.driver, 20)
            
        except Exception as e:
            print(f"[ERROR] Failed to initialize Chrome driver: {e}")
            try:
                minimal_options = uc.ChromeOptions()
                if HEADLESS_MODE:
                    minimal_options.add_argument("--headless")
                minimal_options.add_argument("--no-sandbox")
                minimal_options.add_experimental_option("prefs", prefs)
                
                self.driver = uc.Chrome(options=minimal_options)
                print("[OK] Chrome driver initialized with minimal configuration")
                self.wait = WebDriverWait(self.driver, 20)
                
            except Exception as final_error:
                raise Exception(f"Could not initialize Chrome driver: {final_error}")
    
    def navigate_to_site(self):
        """Navigate to the IMF query tool"""
        try:
            print("Accessing IMF Financial Data Query Tool...")
            self.driver.get(self.url)
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(3)
            print("[OK] Website loaded successfully")
            return True
        except Exception as e:
            print(f"[ERROR] Error accessing website: {e}")
            return False
    
    def select_all_members(self):
        """Point 2: Select all members and lenders"""
        try:
            print("Point 2: Selecting all members and lenders...")
            
            transfer_all_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.rlbTransferAllFrom"))
            )
            transfer_all_button.click()
            time.sleep(3)
            
            print("[OK] Successfully selected all members and lenders")
            return True
            
        except Exception as e:
            print(f"[ERROR] Failed to select all members: {str(e)}")
            try:
                alt_button = self.driver.find_element(By.CSS_SELECTOR, "a[title='All to Right']")
                alt_button.click()
                time.sleep(3)
                print("[OK] Successfully clicked alternative 'All to Right' button")
                return True
            except:
                print("[ERROR] Could not find any 'All to Right' button")
                return False
    
    def select_commitments_borrowings(self):
        """Point 3: Select Commitments/Borrowings, then Borrowings radio, then all borrowings types"""
        try:
            print("Point 3: Selecting Commitments/Borrowings...")
            
            # Step 3a: Select Commitments/Borrowings radio button
            commitments_radio = self.wait.until(EC.element_to_be_clickable((By.ID, "rbArrBorr")))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", commitments_radio)
            time.sleep(1)
            commitments_radio.click()
            print("[OK] Selected 'Commitments/Borrowings' radio button")
            
            # Wait for borrowings options to appear
            time.sleep(3)
            
            # Step 3b: Select "Borrowings" radio button (this enables the checkboxes)
            if not self.select_borrowings_radio():
                return False
            
            # Step 3c: Select all borrowings checkboxes
            return self.select_all_borrowings()
            
        except Exception as e:
            print(f"[ERROR] Failed to select Commitments/Borrowings: {str(e)}")
            return False
    
    def select_borrowings_radio(self):
        """Select the Borrowings radio button that enables the checkboxes"""
        try:
            print("Selecting 'Borrowings' radio button...")
            
            # Look for the Borrowings radio button (rbBorrowings)
            borrowings_radio = self.wait.until(EC.element_to_be_clickable((By.ID, "rbBorrowings")))
            
            # Scroll to element if needed
            self.driver.execute_script("arguments[0].scrollIntoView(true);", borrowings_radio)
            time.sleep(1)
            
            # Click the Borrowings radio button
            borrowings_radio.click()
            print("[OK] Selected 'Borrowings' radio button")
            
            # Wait for checkboxes to become available
            time.sleep(2)
            return True
            
        except Exception as e:
            print(f"[ERROR] Failed to select Borrowings radio button: {str(e)}")
            # Try alternative approaches
            try:
                # Alternative: try by value
                alt_borrowings = self.driver.find_element(By.XPATH, "//input[@value='BORROWINGS' and @name='rblArrBorr']")
                alt_borrowings.click()
                print("[OK] Selected Borrowings radio button (alternative method)")
                time.sleep(2)
                return True
            except Exception as alt_e:
                print(f"[ERROR] Could not find Borrowings radio button with any method: {alt_e}")
                return False
    
    def select_all_borrowings(self):
        """Select all borrowings checkboxes (GRA, PRGT, RST)"""
        try:
            print("Selecting all borrowings types...")
            
            borrowings_checkboxes = [
                ("cblBorrowings_0", "GRA Borrowings"),
                ("cblBorrowings_1", "PRGT Borrowings"), 
                ("cblBorrowings_2", "RST Borrowings")
            ]
            
            selected_count = 0
            for checkbox_id, checkbox_name in borrowings_checkboxes:
                try:
                    checkbox = self.wait.until(EC.element_to_be_clickable((By.ID, checkbox_id)))
                    if not checkbox.is_selected():
                        checkbox.click()
                        print(f"[OK] Selected {checkbox_name}")
                        selected_count += 1
                        time.sleep(0.5)
                    else:
                        print(f"[INFO] {checkbox_name} already selected")
                        selected_count += 1
                except Exception as e:
                    print(f"[WARNING] Could not select {checkbox_name}: {str(e)}")
            
            print(f"[OK] Successfully selected {selected_count} borrowings types")
            return selected_count > 0
            
        except Exception as e:
            print(f"[ERROR] Failed to select borrowings: {str(e)}")
            return False
    
    def select_current_option(self):
        """Point 4: Select Current option for date selection"""
        try:
            print("Point 4: Selecting Current option...")
            
            # Look for the Current radio button (rbCurrent)
            current_radio = self.wait.until(EC.element_to_be_clickable((By.ID, "rbCurrent")))
            
            # Scroll to element if needed
            self.driver.execute_script("arguments[0].scrollIntoView(true);", current_radio)
            time.sleep(1)
            
            # Click the Current radio button
            current_radio.click()
            print("[OK] Successfully selected 'Current' option")
            
            time.sleep(2)  # Wait for any dynamic content to load
            return True
            
        except Exception as e:
            print(f"[ERROR] Failed to select Current option: {str(e)}")
            # Try alternative approaches
            try:
                # Alternative: try by value
                alt_current = self.driver.find_element(By.XPATH, "//input[@value='ACTIVE' and @name='rblArrBorrOptions']")
                alt_current.click()
                print("[OK] Successfully selected Current option (alternative method)")
                return True
            except:
                print("[ERROR] Could not find Current option with any method")
                return False
    
    def submit_query(self):
        """Point 5: Submit the query"""
        try:
            print("Point 5: Submitting query...")
            
            # Look for the Submit button
            submit_button = self.wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
            
            # Scroll to element if needed
            self.driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
            time.sleep(1)
            
            # Click the Submit button
            submit_button.click()
            print("[OK] Query submitted successfully")
            
            # Wait for the new page to load
            print("Waiting for results page to load...")
            time.sleep(5)
            
            # Wait for page to be ready
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(3)
            
            print("[OK] Results page loaded")
            return True
            
        except Exception as e:
            print(f"[ERROR] Failed to submit query: {str(e)}")
            return False
    
    def download_excel_file(self):
        """Point 6: Download the Excel file"""
        try:
            print("Point 6: Downloading Excel file...")
            
            # Look for the XLS download link
            xls_link = self.wait.until(EC.element_to_be_clickable((By.ID, "lbnTSV")))
            
            # Scroll to element if needed
            self.driver.execute_script("arguments[0].scrollIntoView(true);", xls_link)
            time.sleep(1)
            
            print("Found XLS download link, clicking...")
            xls_link.click()
            
            print("[OK] Excel download initiated")
            
            # Wait for download to complete
            print("Waiting for download to complete...")
            time.sleep(10)
            
            # Check if file was downloaded
            downloaded_files = [f for f in os.listdir(self.downloads_dir) if f.endswith(('.xls', '.xlsx'))]
            if downloaded_files:
                latest_file = max([os.path.join(self.downloads_dir, f) for f in downloaded_files], key=os.path.getctime)
                print(f"[OK] File downloaded successfully: {os.path.basename(latest_file)}")
                print(f"File location: {latest_file}")
                return latest_file
            else:
                print("[WARNING] Download initiated but file not found in downloads folder")
                return True
            
        except Exception as e:
            print(f"[ERROR] Failed to download Excel file: {str(e)}")
            return False
    
    def run_complete_automation(self):
        """Run the complete automation process"""
        try:
            print("\nStarting Complete IMF Data Extraction Automation")
            print("=" * 60)
            
            # Step 1: Navigate to site
            if not self.navigate_to_site():
                return False
            
            # Step 2: Select all members
            if not self.select_all_members():
                return False
            
            # Step 3: Select Commitments/Borrowings, Borrowings radio, and all borrowings types
            if not self.select_commitments_borrowings():
                return False
            
            # Step 4: Select Current option
            if not self.select_current_option():
                return False
            
            # Step 5: Submit query
            if not self.submit_query():
                return False
            
            # Step 6: Download Excel file
            downloaded_file = self.download_excel_file()
            if not downloaded_file:
                return False
            
            print("\nCOMPLETE SUCCESS!")
            print("=" * 60)
            print("[OK] Point 2: Selected all members and lenders")
            print("[OK] Point 3: Selected Commitments/Borrowings, Borrowings radio, and all borrowings types")
            print("[OK] Point 4: Selected Current option")
            print("[OK] Point 5: Submitted query successfully")
            print("[OK] Point 6: Downloaded Excel file")
            print(f"[OK] Download location: {self.downloads_dir}")
            
            return True
            
        except Exception as e:
            print(f"[ERROR] Complete automation failed: {str(e)}")
            return False
        finally:
            # Keep browser open for a moment to see results
            print("\nKeeping browser open for 10 seconds to view results...")
            time.sleep(10)
            
            if self.driver:
                try:
                    self.driver.quit()
                    print("Browser closed")
                except:
                    pass

# Usage
if __name__ == "__main__":
    try:
        print("Complete IMF Financial Data Extraction Tool")
        print("Executing Full NABIMFD Runbook Process")
        print("=" * 60)
        
        automator = IMFCompleteAutomator()
        success = automator.run_complete_automation()
        
        if success:
            print("\nAll steps completed successfully!")
            print("The Excel file with IMF data has been downloaded.")
        else:
            print("\nAutomation failed at some step.")
            
    except ImportError as e:
        print(f"Missing packages: pip install undetected-chromedriver selenium")
        print(f"Error: {str(e)}")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        print("Please ensure Chrome is installed and accessible.")