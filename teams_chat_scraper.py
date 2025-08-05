"""
Microsoft Teams Chat Scraper with Virtual Scrolling Message Accumulation
Handles Teams' virtual DOM unloading, deduplicates messages, and downloads images/attachments.
"""

import os
import time
import json
import csv
import requests
import hashlib
import re
import zipfile
import io
import subprocess
import platform
import sys
import glob
from datetime import datetime
from urllib.parse import urlparse, urljoin
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

class TeamsCollector:
    def load_chat_hashes(self, chat_name):
        """
        Loads known message_hashes for a chat from a file.
        """
        hash_file = os.path.join(self.output_dir, f"{chat_name}_hashes.txt")
        if not os.path.exists(hash_file):
            return set()
        try:
            with open(hash_file, 'r', encoding='utf-8') as f:
                return set(line.strip() for line in f if line.strip())
        except Exception as e:
            print(f"Error loading hash file for chat {chat_name}: {e}")
            return set()

    def save_chat_hashes(self, chat_name, new_hashes):
        """
        Saves new message_hashes for a chat to the file.
        """
        hash_file = os.path.join(self.output_dir, f"{chat_name}_hashes.txt")
        try:
            with open(hash_file, 'a', encoding='utf-8') as f:
                for h in new_hashes:
                    f.write(h + '\n')
        except Exception as e:
            print(f"Error saving hash file for chat {chat_name}: {e}")

    def extract_and_accumulate_only_new_messages(self, chat_container, chat_name="Unknown"):
        """
        Extracts and saves only new messages for a chat.
        """
        known_hashes = self.load_chat_hashes(chat_name)
        print(f"Known messages for chat '{chat_name}': {len(known_hashes)}")
        self.accumulated_messages.clear()
        self.message_hashes.clear()
        current_messages = self.get_current_messages()
        new_hashes = set()
        newly_found = 0
        for i, msg_elem in enumerate(current_messages):
            try:
                text_content = ""
                try:
                    text_content = msg_elem.get_attribute('innerText') or msg_elem.text
                except:
                    try:
                        p_tag = msg_elem.find_element(By.TAG_NAME, 'p')
                        text_content = p_tag.get_attribute('innerText') or p_tag.text
                    except:
                        continue
                if not text_content or len(text_content.strip()) == 0:
                    continue
                timestamp = "Unknown"
                author = "Unknown"
                try:
                    parent = msg_elem.find_element(By.XPATH, '..')
                    time_elem = parent.find_element(By.CSS_SELECTOR, '[data-tid*="timestamp"], .message-timestamp, time')
                    timestamp = time_elem.get_attribute('innerText') or time_elem.text
                except:
                    pass
                try:
                    parent = msg_elem.find_element(By.XPATH, '..')
                    author_elem = parent.find_element(By.CSS_SELECTOR, '[data-tid*="author"], .message-author')
                    author = author_elem.get_attribute('innerText') or author_elem.text
                except:
                    pass
                msg_hash = self.create_message_hash(text_content.strip(), author, timestamp)
                if msg_hash in known_hashes:
                    continue
                images = []
                attachments = []
                if self.download_images:
                    images = self.extract_images_from_message(msg_elem)
                    attachments = self.extract_attachments_from_message(msg_elem)
                message_data = {
                    'chat_name': chat_name,
                    'message_hash': msg_hash,
                    'author': author,
                    'timestamp': timestamp,
                    'content': text_content.strip(),
                    'images': images,
                    'attachments': attachments,
                    'extracted_at': datetime.now().isoformat()
                }
                self.accumulated_messages[msg_hash] = message_data
                new_hashes.add(msg_hash)
                newly_found += 1
            except Exception as e:
                print(f"Error processing message element: {e}")
                continue
        self.save_chat_hashes(chat_name, new_hashes)
        print(f"New messages for chat '{chat_name}' found and saved: {newly_found}")
        return newly_found
    def download_edge_driver_direct(self, version, driver_dir=None):
        """
        Downloads msedgedriver directly from the Microsoft URL with version number and unzips it.
        Uses the system proxy (requests uses environment variables by default).
        Args:
            version (str): e.g. '138.0.3351.95'
            driver_dir (str): Target directory for the driver
        Returns:
            str: Path to the extracted msedgedriver.exe or None on error
        """
        import zipfile, io
        if not driver_dir:
            driver_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "edgedriver")
        os.makedirs(driver_dir, exist_ok=True)
        url = f"https://msedgedriver.microsoft.com/{version}/edgedriver_win64.zip"
        print(f"Trying direct download from: {url}")
        try:
            response = requests.get(url, timeout=60, allow_redirects=True)
            if response.status_code != 200:
                print(f"Error downloading: HTTP {response.status_code}")
                return None
            with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
                zip_file.extractall(driver_dir)
            # Search for msedgedriver.exe in the target directory
            for file in os.listdir(driver_dir):
                if file.lower().startswith("msedgedriver") and file.lower().endswith(".exe"):
                    driver_path = os.path.join(driver_dir, file)
                    print(f"✓ msedgedriver downloaded and extracted: {driver_path}")
                    return driver_path
            print("No msedgedriver.exe found after extraction.")
            return None
        except Exception as e:
            print(f"Error during direct download/extraction: {e}")
            return None
    def __init__(self, output_dir="teams_export", headless=False, download_images=True, auto_select_all=False):
        self.output_dir = output_dir
        self.headless = headless
        self.download_images = download_images
        self.auto_select_all = auto_select_all  # Option to automatically select all chats
        self.driver = None
        self.wait = None
        self.chat_data = []
        self.downloaded_images = set()
        self.images_dir = os.path.join(output_dir, "images")
        self.accumulated_messages = {}
        self.message_hashes = set()
        self.driver_path = None  # Path to the msedgedriver
        self.current_chat_name = "Unknown"  # Current chat name for image naming
        
        # Constants for scrolling and loading
        self.SCROLL_SPEED = 5  # Number of scroll steps
        self.SLEEP_TIME_BETWEEN_SCROLLS = 2  # Wait time between scroll actions
        self.SLEEP_AFTER_LOAD_MORE = 1  # Wait time after clicking "load more" button
        self.MAX_LOAD_MORE_ATTEMPTS = 50  # Maximum attempts to avoid infinite loops

        self.selectors = {
            'chat_list': 'div[data-tid="chat-pane-list"]',
            'chat_item': 'li[data-tid*="chat-item"]',
            'message_list': 'div[data-tid="chat-messages-list"]',
            'message_body': '[data-tid="message-body"]',
            'message_time': '[data-tid="message-timestamp"]',
            'message_author': '[data-tid="message-author-name"]',
            'chat_title': '[data-tid="chat-header-title"]',
            'scroll_container': '[data-tid="chat-pane-runway"]',
            'chat_container': '[data-tid="chat-messages-container"]',
            'message_images': 'img, [data-tid="message-image"]',
            'message_attachments': '[data-tid="message-attachment"], .attachment-item'
        }

        os.makedirs(self.output_dir, exist_ok=True)
        if download_images:
            os.makedirs(self.images_dir, exist_ok=True)
            
    def ensure_edge_driver(self):
        """
        Ensures that msedgedriver is available. First checks for local driver, then tries package installation.
        Returns the path to the msedgedriver executable.
        """
        # First, check if we already have a local driver
        driver_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "edgedriver")
        if os.path.exists(driver_dir):
            print(f"Checking for existing driver in: {driver_dir}")
            # Look for existing msedgedriver.exe
            for file in os.listdir(driver_dir):
                if file.lower().startswith("msedgedriver") and file.lower().endswith(".exe"):
                    driver_path = os.path.join(driver_dir, file)
                    print(f"✓ Found existing msedgedriver at: {driver_path}")
                    # Verify the file exists and is executable
                    if os.path.exists(driver_path):
                        return driver_path
        
        # Try to use the msedgedriver Python package only if no local driver found
        try:
            import msedgedriver
            print("Checking for msedgedriver using msedgedriver Python package...")
            driver_path = msedgedriver.install()
            print(f"✓ msedgedriver installed at: {driver_path}")
            return driver_path
        except Exception as e:
            print(f"⚠️ Error installing msedgedriver: {e}")
            print("Trying fallback: Download via Microsoft URL with Edge version...")
            # Try to determine Edge version
            edge_version = self.get_edge_version()
            if edge_version:
                major_minor_patch = '.'.join(edge_version.split('.')[:4])
                driver_path = self.download_edge_driver_direct(major_minor_patch)
                if driver_path:
                    return driver_path
            print("Will attempt to use the default driver provided by Selenium.")
            return None
    
    def get_edge_version(self):
        """
        Gets the installed Microsoft Edge version.
        
        Returns:
            str: Edge version string (e.g., "115.0.1901.183") or None if not found
        """
        try:
            if sys.platform == "win32":
                # Windows: Use reg query
                try:
                    cmd = r'reg query "HKEY_CURRENT_USER\Software\Microsoft\Edge\BLBeacon" /v version'
                    result = subprocess.check_output(cmd, shell=True, text=True)
                    match = re.search(r'version\s+REG_SZ\s+(\d+\.\d+\.\d+\.\d+)', result)
                    if match:
                        return match.group(1)
                except:
                    # Try alternative registry path
                    try:
                        cmd = r'reg query "HKEY_CURRENT_USER\Software\Microsoft\EdgeUpdate\Clients\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}" /v pv'
                        result = subprocess.check_output(cmd, shell=True, text=True)
                        match = re.search(r'pv\s+REG_SZ\s+(\d+\.\d+\.\d+\.\d+)', result)
                        if match:
                            return match.group(1)
                    except:
                        pass
            
            elif sys.platform == "darwin":
                # macOS
                try:
                    cmd = r'/Applications/Microsoft\ Edge.app/Contents/MacOS/Microsoft\ Edge --version'
                    result = subprocess.check_output(cmd, shell=True, text=True)
                    match = re.search(r'Microsoft Edge (\d+\.\d+\.\d+\.\d+)', result)
                    if match:
                        return match.group(1)
                except:
                    pass
            
            else:
                # Linux
                try:
                    cmd = 'microsoft-edge --version'
                    result = subprocess.check_output(cmd, shell=True, text=True)
                    match = re.search(r'Microsoft Edge (\d+\.\d+\.\d+\.\d+)', result)
                    if match:
                        return match.group(1)
                except:
                    pass
            
            # If we get here, try to run Edge with --version flag
            try:
                possible_paths = []
                if sys.platform == "win32":
                    possible_paths = [
                        r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
                        r'C:\Program Files\Microsoft\Edge\Application\msedge.exe'
                    ]
                
                for path in possible_paths:
                    if os.path.exists(path):
                        cmd = f'"{path}" --version'
                        result = subprocess.check_output(cmd, shell=True, text=True)
                        match = re.search(r'Microsoft Edge (\d+\.\d+\.\d+\.\d+)', result)
                        if match:
                            return match.group(1)
            except:
                pass
                
            return None
        except Exception as e:
            print(f"Error getting Edge version: {e}")
            return None
    
    def download_edge_driver(self, driver_dir, version=None):
        """
        Downloads the msedgedriver for the specified version or the latest version.
        For Edge versions >= 120, always use LATEST_STABLE, as LATEST_RELEASE_xxx no longer exists.
        """
        try:
            # Determine the platform
            if sys.platform == "win32":
                platform_name = "win64"  # Using 64-bit version
                driver_name = "msedgedriver.exe"
            elif sys.platform == "darwin":
                platform_name = "mac64"
                driver_name = "msedgedriver"
            else:
                platform_name = "linux64"
                driver_name = "msedgedriver"

            # For Edge versions >= 120, there is no LATEST_RELEASE_xxx URL anymore
            use_latest_stable = False
            if version:
                try:
                    major = int(str(version).split(".")[0])
                    if major >= 120:
                        use_latest_stable = True
                except Exception:
                    use_latest_stable = True
            if use_latest_stable or not version:
                driver_version_url = "https://msedgedriver.azureedge.net/LATEST_STABLE"
            else:
                driver_version_url = f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{version}"

            print(f"Fetching driver version from: {driver_version_url}")
            response = requests.get(driver_version_url, timeout=30)
            if response.status_code != 200:
                print(f"⚠️ Failed to get driver version: HTTP {response.status_code}")
                # Try alternative source
                if version and not use_latest_stable:
                    print("Trying alternative source for driver version...")
                    alt_url = f"https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/LATEST_RELEASE_{version}"
                    response = requests.get(alt_url, timeout=30)
                    if response.status_code != 200:
                        print(f"⚠️ Alternative source also failed: HTTP {response.status_code}")
                        return None
                else:
                    return None

            driver_version = response.text.strip()
            print(f"✓ Driver version to download: {driver_version}")

            # Download the driver
            download_url = f"https://msedgedriver.azureedge.net/{driver_version}/edgedriver_{platform_name}.zip"
            print(f"Downloading from: {download_url}")

            response = requests.get(download_url, timeout=60)
            if response.status_code != 200:
                print(f"⚠️ Failed to download driver: HTTP {response.status_code}")
                # Try alternative source
                print("Trying alternative source for driver download...")
                alt_url = f"https://msedgewebdriverstorage.blob.core.windows.net/edgewebdriver/{driver_version}/edgedriver_{platform_name}.zip"
                response = requests.get(alt_url, timeout=60)
                if response.status_code != 200:
                    print(f"⚠️ Alternative download source also failed: HTTP {response.status_code}")
                    return None

            # Extract the driver
            with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
                zip_file.extractall(driver_dir)

            # Rename the driver to include the version
            driver_path = os.path.join(driver_dir, driver_name)
            versioned_driver_path = os.path.join(driver_dir, f"msedgedriver_{version}.exe" if sys.platform == "win32" else f"msedgedriver_{version}")

            if os.path.exists(driver_path):
                # Make sure the old driver is removed if it exists
                if os.path.exists(versioned_driver_path):
                    os.remove(versioned_driver_path)
                os.rename(driver_path, versioned_driver_path)

                # Make the driver executable on Unix-like systems
                if sys.platform != "win32":
                    os.chmod(versioned_driver_path, 0o755)

                print(f"✓ Driver downloaded and extracted to: {versioned_driver_path}")
                return versioned_driver_path
            else:
                print(f"⚠️ Driver not found in extracted files. Looking for any driver in the directory...")
                # Look for any driver in the extracted directory
                for file in os.listdir(driver_dir):
                    if file.startswith("msedgedriver") or file == "msedgedriver" or file == "msedgedriver.exe":
                        found_path = os.path.join(driver_dir, file)
                        print(f"✓ Found driver at: {found_path}")
                        return found_path

                print("⚠️ No driver found in extracted files.")
                return None

        except Exception as e:
            print(f"⚠️ Error downloading Edge driver: {e}")
            return None

    def setup_driver(self):
        edge_options = EdgeOptions()
        edge_options.add_argument("--disable-blink-features=AutomationControlled")
        edge_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        edge_options.add_experimental_option('useAutomationExtension', False)
        
        # Reduce error logs and suppress Task Manager errors
        edge_options.add_argument("--log-level=3")  # FATAL level
        edge_options.add_argument("--silent")
        edge_options.add_argument("--disable-logging")
        edge_options.add_argument("--disable-in-process-stack-traces")
        edge_options.add_argument("--disable-extensions")
        edge_options.add_argument("--disable-component-extensions-with-background-pages")
        edge_options.add_argument("--disable-default-apps")
        edge_options.add_argument("--disable-background-networking")
        edge_options.add_argument("--disable-background-timer-throttling")
        edge_options.add_argument("--disable-backgrounding-occluded-windows")
        edge_options.add_argument("--disable-breakpad")
        edge_options.add_argument("--disable-client-side-phishing-detection")
        edge_options.add_argument("--disable-features=TranslateUI,BlinkGenPropertyTrees,LazyFrameLoading,site-per-process")
        
        # Standard options
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-dev-shm-usage")
        edge_options.add_argument("--disable-gpu")
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument("--disable-features=VizDisplayCompositor")
        edge_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0")
        
        if self.headless:
            edge_options.add_argument("--headless")
        
        try:
            # Create a service object with suppressed output and the correct driver path
            service = EdgeService(
                executable_path=self.driver_path,
                log_output=os.devnull
            )
            
            # Use the service object when creating the webdriver
            self.driver = webdriver.Edge(options=edge_options, service=service)
            print("✓ Edge WebDriver initialized successfully")
        except Exception as e:
            print(f"⚠️ Edge WebDriver initialization failed: {e}")
            print("Trying with default Selenium WebDriver manager...")
            try:
                # Fall back to default Selenium WebDriver manager
                self.driver = webdriver.Edge(options=edge_options)
                print("✓ Edge WebDriver initialized with default manager")
            except Exception as e2:
                print(f"⚠️ Default WebDriver manager also failed: {e2}")
                return False
            
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 20)
        return True

    def navigate_to_teams(self):
        print("Navigating to Microsoft Teams...")
        self.driver.get("https://teams.microsoft.com/v2/")
        time.sleep(5)
        return True

    def wait_for_sso_login(self):
        time.sleep(5)
        return True

    def wait_for_loading_screen_to_disappear(self):
        """Wait for the loading screen to disappear."""
        try:
            loading_screen_selectors = [
                '#loading-screen',
                'div[role="progressbar"]',
                '.loading-screen',
                '[aria-valuetext="Loading..."]'
            ]
            
            for selector in loading_screen_selectors:
                try:
                    # Wait for loading screen to be invisible or not present
                    WebDriverWait(self.driver, 30).until_not(
                        lambda d: d.find_element(By.CSS_SELECTOR, selector).is_displayed()
                    )
                    print("✓ Loading screen disappeared")
                    return True
                except:
                    continue
            
            # If we couldn't find any loading screen, assume it's already gone
            return True
        except Exception as e:
            print(f"Warning: Could not determine loading screen status: {e}")
            # Continue anyway
            return True

    def navigate_to_chats(self):
        try:
            print("Navigating to the chat area...")
            
            # First wait for any loading screens to disappear
            self.wait_for_loading_screen_to_disappear()
            
            # Add a small delay to ensure the page is stable
            time.sleep(5)
            
            chat_selectors = [
                'div[aria-label="Chat"]',
                '[data-tid="app-bar-button-chat"]',
                'button[aria-label*="Chat"]',
                'button[title="Chat"]',
                'button[aria-label="Chats"]'
            ]
            
            chat_button = None
            max_attempts = 3
            
            for attempt in range(max_attempts):
                if attempt > 0:
                    print(f"Retry attempt {attempt} to click chat button...")
                    time.sleep(3)
                
                for selector in chat_selectors:
                    try:
                        # Use explicit wait for element to be clickable
                        from selenium.webdriver.support import expected_conditions as EC
                        chat_button = WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                        
                        if chat_button.is_displayed() and chat_button.is_enabled():
                            # Try JavaScript click if regular click might be intercepted
                            try:
                                self.driver.execute_script("arguments[0].click();", chat_button)
                                time.sleep(3)
                                print("✓ Chat area opened (via JavaScript click)")
                                return True
                            except Exception as js_e:
                                print(f"JavaScript click failed, trying regular click: {js_e}")
                                chat_button.click()
                                time.sleep(3)
                                print("✓ Chat area opened (via regular click)")
                                return True
                    except:
                        continue
            
            if chat_button:
                # One last attempt with ActionChains
                try:
                    ActionChains(self.driver).move_to_element(chat_button).click().perform()
                    time.sleep(3)
                    print("✓ Chat area opened (via ActionChains)")
                    return True
                except Exception as ac_e:
                    print(f"ActionChains click failed: {ac_e}")
            
            print("✗ Chat button not found or not clickable after multiple attempts")
            return False
        except Exception as e:
            print(f"Error navigating to chats: {e}")
            return False

    def get_chat_list(self):
        try:
            print("Gathering chat list...")
            time.sleep(3)
            print("Attempting to load more chats using fixed CSS selector ...")
            # Try to directly find and click the "show more" button by its data-testid
            for i in range(self.MAX_LOAD_MORE_ATTEMPTS):
                try:
                    load_more = None
                    try:
                        load_more = self.driver.find_element(By.CSS_SELECTOR, '[data-testid="load-next-page-button"]')
                    except Exception:
                        pass
                    if load_more and load_more.is_displayed():
                        print(f"  - Found 'show more' button (attempt {i+1}/{self.MAX_LOAD_MORE_ATTEMPTS})")
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", load_more)
                        time.sleep(0.5)
                        self.driver.execute_script("arguments[0].click();", load_more)
                        time.sleep(2)
                    else:
                        print(f"  - No more 'show more' button after {i+1} attempts.")
                        break
                except Exception as e:
                    print(f"  - Could not find or click 'show more' button on attempt {i+1}: {str(e)[:100]}")
                    break
            # Define selectors for chat items
            chat_selectors = [
                'div[data-item-type="chat"][data-testid="list-item"]',
                'div[data-item-type="chats"] div[role="group"] > *',
                'div[data-tid="chat-pane-list"] > *',
                'li[data-tid*="chat-item"]',
                '[role="listitem"]'
            ]
            # Collect chat items
            chat_items = []
            best_selector = None
            for selector in chat_selectors:
                try:
                    items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if items and len(items) > len(chat_items):
                        chat_items = items
                        best_selector = selector
                        print(f"✓ {len(chat_items)} chats found with selector: {selector}")
                except:
                    continue
            if chat_items:
                print(f"✓ Total chats found: {len(chat_items)}")
            else:
                print("✗ No chat items found")
                return []
            return chat_items
        except Exception as e:
            print(f"Error gathering chat list: {e}")
            return []

    def search_chats(self):
        """
        Allows the user to search for chats and returns the search results.
        Uses the chat filter search field in the left sidebar.
        
        Returns:
            A list of chat elements that match the search criteria
        """
        try:
            print("\nChat search activated")
            search_term = input("Enter a search term: ")
            
            if not search_term or search_term.strip() == "":
                print("No search term entered. Returning to normal chat list.")
                return self.get_chat_list()
                
            print(f"Searching for chats with the term: '{search_term}'")
            
            # Exact selectors based on the HTML code of the search field
            search_box_selectors = [
                # User-provided selector
                '#simple-collab-left-rail-sticky-filter-input-id',
                # Very specific selectors based on the HTML code
                'input[data-testid="simple-collab-left-rail-sticky-filter-input"]',
                'input#simple-collab-left-rail-sticky-filter-input-id',
                'input[placeholder="Filter by name or group name"]',
                'input[aria-label="Filter by name or group name"]',
                '.fui-Input__input',
                # German translations
                'input[placeholder*="Filter nach Name"]',
                'input[aria-label*="Filter nach Name"]',
                # More general fallback selectors
                'input[placeholder*="Filter"]',
                'input[aria-label*="Filter"]',
                'input.fui-Input__input'
            ]
            
            print("Searching for the chat filter search field...")
            
            # Debug output of all visible input fields
            try:
                all_inputs = self.driver.find_elements(By.TAG_NAME, 'input')
                visible_inputs = [inp for inp in all_inputs if inp.is_displayed()]
                print(f"Found visible input fields: {len(visible_inputs)}")
                for i, inp in enumerate(visible_inputs[:5]):  # Show only the first 5
                    print(f"  Input {i+1}: placeholder='{inp.get_attribute('placeholder')}', aria-label='{inp.get_attribute('aria-label')}', id='{inp.get_attribute('id')}', class='{inp.get_attribute('class')}'")
            except Exception as e:
                print(f"Error listing input fields: {e}")
            
            # Search for the search field
            search_box = None
            for selector in search_box_selectors:
                try:
                    print(f"Trying selector: {selector}")
                    search_box = WebDriverWait(self.driver, 5).until(
                        lambda d: d.find_element(By.CSS_SELECTOR, selector)
                    )
                    if search_box and search_box.is_displayed():
                        print(f"✓ Chat filter search field found with selector: {selector}")
                        break
                    else:
                        print(f"Search field found, but not visible: {selector}")
                except Exception as e:
                    print(f"Selector not found: {selector} ({str(e)[:50]}...)")
                    continue
            
            # If the search field is not found, try Ctrl+Shift+F
            if not search_box:
                print("Chat filter search field not found. Trying Ctrl+Shift+F...")
                actions = ActionChains(self.driver)
                actions.key_down(Keys.CONTROL)
                actions.key_down(Keys.SHIFT)
                actions.send_keys('f')
                actions.key_up(Keys.SHIFT)
                actions.key_up(Keys.CONTROL)
                actions.perform()
                
                # Wait briefly and try to find the search field again
                time.sleep(3)
                
                for selector in search_box_selectors:
                    try:
                        search_box = WebDriverWait(self.driver, 5).until(
                            lambda d: d.find_element(By.CSS_SELECTOR, selector)
                        )
                        if search_box and search_box.is_displayed():
                            print(f"✓ Chat filter search field found after Ctrl+Shift+F with selector: {selector}")
                            break
                    except:
                        continue
            
            # If still no search field found, try one last alternative
            if not search_box:
                print("Chat filter search field not found. Trying alternative method...")
                try:
                    # Try to search through all visible input fields
                    for inp in visible_inputs:
                        placeholder = inp.get_attribute('placeholder') or ''
                        aria_label = inp.get_attribute('aria-label') or ''
                        if ('filter' in placeholder.lower() or 'filter' in aria_label.lower() or
                            'suchen' in placeholder.lower() or 'suchen' in aria_label.lower()):
                            search_box = inp
                            print(f"✓ Matching input field found: placeholder='{placeholder}', aria-label='{aria_label}'")
                            break
                except Exception as e:
                    print(f"Error in alternative search: {e}")
            
            if search_box:
                # Click on the search field and enter the search term
                try:
                    print("Clicking on chat filter search field...")
                    search_box.click()
                    time.sleep(1)
                    print("Clearing existing text...")
                    search_box.clear()
                    print(f"Entering search term: {search_term}")
                    search_box.send_keys(search_term)
                    # Don't press Enter as the filter field filters in real-time
                    print("Waiting for filtered results...")
                    time.sleep(2)
                except Exception as e:
                    print(f"Error entering text in the chat filter search field: {e}")
                    # Try with JavaScript
                    try:
                        print("Trying input with JavaScript...")
                        self.driver.execute_script(f"arguments[0].value = '{search_term}';", search_box)
                        self.driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", search_box)
                    except Exception as js_e:
                        print(f"JavaScript input failed: {js_e}")
            else:
                print("No matching search field found. Returning to normal chat list.")
                return self.get_chat_list()
            
            # Wait for the filtered results
            time.sleep(3)
            
            # Get the filtered chat elements
            chat_item_selectors = [
                'div[data-item-type="chats"] div[role="group"] > *',
                'div[data-tid="chat-pane-list"] > *',
                'li[data-tid*="chat-item"]',
                '[role="listitem"]'
            ]
            
            chat_items = []
            for selector in chat_item_selectors:
                try:
                    print(f"Searching for filtered chat elements with selector: {selector}")
                    items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if items:
                        chat_items = items
                        print(f"✓ {len(chat_items)} filtered chat elements found with selector: {selector}")
                        break
                except Exception as e:
                    print(f"No chat elements with selector: {selector} ({str(e)[:50]}...)")
                    continue
                    
            if not chat_items:
                print("No filtered chat elements found. Returning to normal chat list.")
                return self.get_chat_list()
                
            return chat_items
            
        except Exception as e:
            print(f"Error in chat search: {e}")
            print("Returning to normal chat list.")
            return self.get_chat_list()
            
    def create_message_hash(self, content, author, timestamp):
        message_string = f"{author}:{timestamp}:{content[:100]}"
        return hashlib.md5(message_string.encode('utf-8')).hexdigest()

    def extract_and_accumulate_current_messages(self, chat_name="Unknown"):
        current_messages = self.get_current_messages()
        newly_found = 0
        for i, msg_elem in enumerate(current_messages):
            try:
                text_content = ""
                try:
                    text_content = msg_elem.get_attribute('innerText') or msg_elem.text
                except:
                    try:
                        p_tag = msg_elem.find_element(By.TAG_NAME, 'p')
                        text_content = p_tag.get_attribute('innerText') or p_tag.text
                    except:
                        continue
                if not text_content or len(text_content.strip()) == 0:
                    continue
                timestamp = "Unknown"
                author = "Unknown"
                try:
                    parent = msg_elem.find_element(By.XPATH, '..')
                    time_elem = parent.find_element(By.CSS_SELECTOR, '[data-tid*="timestamp"], .message-timestamp, time')
                    timestamp = time_elem.get_attribute('innerText') or time_elem.text
                except:
                    pass
                try:
                    parent = msg_elem.find_element(By.XPATH, '..')
                    author_elem = parent.find_element(By.CSS_SELECTOR, '[data-tid*="author"], .message-author')
                    author = author_elem.get_attribute('innerText') or author_elem.text
                except:
                    pass
                msg_hash = self.create_message_hash(text_content.strip(), author, timestamp)
                if msg_hash not in self.message_hashes:
                    images = []
                    attachments = []
                    if self.download_images:
                        images = self.extract_images_from_message(msg_elem)
                        attachments = self.extract_attachments_from_message(msg_elem)
                    message_data = {
                        'chat_name': chat_name,
                        'message_hash': msg_hash,
                        'author': author,
                        'timestamp': timestamp,
                        'content': text_content.strip(),
                        'images': images,
                        'attachments': attachments,
                        'extracted_at': datetime.now().isoformat()
                    }
                    self.accumulated_messages[msg_hash] = message_data
                    self.message_hashes.add(msg_hash)
                    newly_found += 1
            except Exception as e:
                print(f"Error processing message element: {e}")
                continue
        return newly_found

    def scroll_to_load_all_messages_with_accumulation(self, chat_container, chat_name="Unknown"):
        print("Starting enhanced infinite scroll with message accumulation...")
        self.accumulated_messages.clear()
        self.message_hashes.clear()
        scroll_attempts = 0
        max_scroll_attempts = 100
        consecutive_no_new_messages = 0
        max_consecutive_no_new = 3
        load_more_attempts = 0
        
        # Nur noch echtes Scrollen, kein "Mehr anzeigen" Button mehr
        while scroll_attempts < max_scroll_attempts:
            newly_found = self.extract_and_accumulate_current_messages(chat_name)
            total_accumulated = len(self.accumulated_messages)
            print(f"Scroll attempt {scroll_attempts + 1}: {newly_found} new messages found, {total_accumulated} total accumulated")
            if newly_found == 0:
                consecutive_no_new_messages += 1
                print(f"  → No new messages found ({consecutive_no_new_messages}/{max_consecutive_no_new})")
                if consecutive_no_new_messages >= max_consecutive_no_new:
                    print("No new messages found after multiple scroll attempts. All messages accumulated.")
                    break
            else:
                consecutive_no_new_messages = 0
            self.perform_enhanced_scroll_strategies(chat_container)
            time.sleep(1)
            scroll_attempts += 1
            #time.sleep(min(scroll_attempts * 0.1, 2))
        final_messages = list(self.accumulated_messages.values())
        for i, msg in enumerate(final_messages):
            msg['message_id'] = i
        print(f"Enhanced infinite scroll completed. Total messages accumulated: {len(final_messages)}")
        return final_messages

    def perform_enhanced_scroll_strategies(self, container):
        strategies = [
            # lambda: self.driver.execute_script("arguments[0].scrollTop = 0;", container),
            # lambda: self.driver.execute_script("arguments[0].scrollTop = Math.max(0, arguments[0].scrollTop - 2000);", container),
            # lambda: [container.send_keys(Keys.PAGE_UP) for _ in range(5)],
           # lambda: container.send_keys(Keys.HOME),
            # New scroll strategy: use scroll_up method
            lambda: self.scroll_up(),
           
            # lambda: self.driver.execute_script("""
            #     var element = arguments[0];
            #     var event = new WheelEvent('wheel', {
            #         deltaY: -1500,
            #         deltaMode: 0,
            #         bubbles: true
            #     });
            #     element.dispatchEvent(event);
            # """, container),
            # lambda: self.driver.execute_script("""
            #     var element = arguments[0];
            #     var currentScroll = element.scrollTop;
            #     var scrollStep = Math.max(500, currentScroll / 4);
            #     element.scrollTop = Math.max(0, currentScroll - scrollStep);
            # """, container)
        ]
        for i, strategy in enumerate(strategies):
            try:
                strategy()
                time.sleep(1.2)
            except Exception as e:
                print(f"Scroll strategy {i+1} failed: {e}")
                continue
                
    def scroll_up(self) -> bool:
        """
        Scrolls up to load older messages.
        Returns whether the viewport remained unchanged (stuck).
        """
        stuck = True
        try:
            start_viewport_text = self.driver.find_elements(By.CSS_SELECTOR, '[data-tid="chat-pane-message"]')[0].text
            for _ in range(self.SCROLL_SPEED):  # Adjust number of scroll steps
                viewport = self.driver.find_elements(By.CSS_SELECTOR, '[data-tid="chat-pane-message"]')[0]
                stuck = stuck and (viewport.text == start_viewport_text)
                viewport.send_keys(Keys.ARROW_UP)
                time.sleep(self.SLEEP_TIME_BETWEEN_SCROLLS)
            return stuck
        except Exception as e:
            print(f"Could not scroll up to load more messages: {e}")
            return stuck

    def get_current_messages(self):
        message_selectors = [
            '[data-tid="message-body"]',
            '[data-tid="chat-message"]',
            '.message-body',
            '.chat-message',
            '[role="listitem"] [data-tid*="message"]',
            'div[data-tid*="message-content"]',
            '[data-tid="chat-pane-item"]'
        ]
        for selector in message_selectors:
            try:
                messages = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if messages:
                    return messages
            except:
                continue
        return []

    def extract_images_from_message(self, message_element):
        images = []
        image_selectors = [
            'img',
            '[data-tid="message-image"]',
            '.message-image',
            '.attachment-image',
            'img[src*="teams.microsoft.com"]',
            'img[src*="sharepoint.com"]',
            'img[src*="onedrive.com"]'
        ]
        for selector in image_selectors:
            try:
                img_elements = message_element.find_elements(By.CSS_SELECTOR, selector)
                for img_elem in img_elements:
                    image_info = self.process_image_element(img_elem)
                    if image_info:
                        images.append(image_info)
            except Exception as e:
                print(f"Error extracting images with selector {selector}: {e}")
                continue
        # KNOWN BUG: Screenshots are not extracted, only emojis and inline images
        return images

    def process_image_element(self, img_element):
        try:
            img_src = img_element.get_attribute('src')
            if not img_src:
                return None
            if img_src.startswith('data:') and len(img_src) < 1000:
                return None
            image_info = {
                'src': img_src,
                'alt': img_element.get_attribute('alt') or '',
                'title': img_element.get_attribute('title') or '',
                'width': img_element.get_attribute('width') or '',
                'height': img_element.get_attribute('height') or '',
                'local_path': None,
                'download_status': 'pending'
            }
            if self.download_images:
                # Get and sanitize the current chat name
                raw_chat_name = self.current_chat_name if hasattr(self, 'current_chat_name') else "Unknown"
                chat_name = self.sanitize_filename(raw_chat_name)
                local_path = self.download_image(img_src, chat_name)
                if local_path:
                    image_info['local_path'] = local_path
                    image_info['download_status'] = 'success'
                    print(f"✓ Image downloaded: {os.path.basename(local_path)}")
                else:
                    image_info['download_status'] = 'failed'
            return image_info
        except Exception as e:
            print(f"Error processing image element: {e}")
            return None

    def download_image(self, img_url, chat_name="Unknown"):
        try:
            url_hash = hashlib.md5(img_url.encode()).hexdigest()
            if url_hash in self.downloaded_images:
                return None
            
            # Sanitize chat name for filename (double-check)
            safe_chat_name = self.sanitize_filename(chat_name)
            
            if img_url.startswith('data:'):
                return self.save_base64_image(img_url, url_hash, safe_chat_name)
            elif img_url.startswith('blob:'):
                return self.download_blob_image(img_url, url_hash, safe_chat_name)
            elif img_url.startswith('http'):
                return self.download_http_image(img_url, url_hash, safe_chat_name)
            else:
                absolute_url = urljoin("https://teams.microsoft.com", img_url)
                return self.download_http_image(absolute_url, url_hash, safe_chat_name)
        except Exception as e:
            print(f"Error downloading image {img_url}: {e}")
            return None

    def save_base64_image(self, data_url, filename_hash, chat_name="Unknown"):
        try:
            import base64
            header, data = data_url.split(',', 1)
            image_data = base64.b64decode(data)
            if 'jpeg' in header or 'jpg' in header:
                ext = 'jpg'
            elif 'png' in header:
                ext = 'png'
            elif 'gif' in header:
                ext = 'gif'
            else:
                ext = 'png'
            
            # Sanitize chat name and create chat-specific directory
            safe_chat_name = self.sanitize_filename(chat_name)
            chat_images_dir = os.path.join(self.images_dir, safe_chat_name)
            os.makedirs(chat_images_dir, exist_ok=True)
            
            filename = f"{filename_hash}.{ext}"
            filepath = os.path.join(chat_images_dir, filename)
            with open(filepath, 'wb') as f:
                f.write(image_data)
            self.downloaded_images.add(filename_hash)
            return filepath
        except Exception as e:
            print(f"Error saving base64 image: {e}")
            return None

    def download_http_image(self, img_url, filename_hash, chat_name="Unknown"):
        try:
            cookies = self.driver.get_cookies()
            session_cookies = {cookie['name']: cookie['value'] for cookie in cookies}
            response = requests.get(
                img_url,
                cookies=session_cookies,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
                    'Referer': 'https://teams.microsoft.com/',
                    'Accept': 'image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'
                },
                timeout=30
            )
            if response.status_code == 200:
                content_type = response.headers.get('content-type', '')
                if 'jpeg' in content_type or 'jpg' in content_type:
                    ext = 'jpg'
                elif 'png' in content_type:
                    ext = 'png'
                elif 'gif' in content_type:
                    ext = 'gif'
                else:
                    parsed = urlparse(img_url)
                    ext = os.path.splitext(parsed.path)[1].lstrip('.') or 'jpg'
                
                # Sanitize chat name and create chat-specific directory
                safe_chat_name = self.sanitize_filename(chat_name)
                chat_images_dir = os.path.join(self.images_dir, safe_chat_name)
                os.makedirs(chat_images_dir, exist_ok=True)
                
                filename = f"{filename_hash}.{ext}"
                filepath = os.path.join(chat_images_dir, filename)
                with open(filepath, 'wb') as f:
                    f.write(response.content)
                self.downloaded_images.add(filename_hash)
                return filepath
        except Exception as e:
            print(f"Error downloading HTTP image {img_url}: {e}")
            return None

    def download_blob_image(self, blob_url, filename_hash, chat_name="Unknown"):
        """
        Downloads an image from a blob URL by converting it to a data URL.
        
        Args:
            blob_url: The blob URL of the image
            filename_hash: A hash value for the filename
            chat_name: The name of the chat for the filename
            
        Returns:
            The path to the saved file or None on error
        """
        try:
            print(f"Attempting to convert blob URL: {blob_url}")
            
            # Simplified JavaScript without arrow functions and with asynchronous XHR
            script = """
            function fetchBlob(url, callback) {
                var xhr = new XMLHttpRequest();
                xhr.open('GET', url, true);
                xhr.responseType = 'blob';
                
                xhr.onload = function() {
                    if (xhr.status === 200) {
                        var reader = new FileReader();
                        reader.onloadend = function() {
                            callback(reader.result);
                        };
                        reader.onerror = function() {
                            callback('ERROR: FileReader error');
                        };
                        reader.readAsDataURL(xhr.response);
                    } else {
                        callback('ERROR: XHR status ' + xhr.status);
                    }
                };
                
                xhr.onerror = function() {
                    callback('ERROR: XHR error');
                };
                
                xhr.send();
            }
            
            var done = arguments[arguments.length - 1];
            fetchBlob(arguments[0], done);
            """
            
            # Execute JavaScript and wait for the result (Data URL)
            data_url = self.driver.execute_async_script(script, blob_url)
            
            # Check if the conversion was successful
            if data_url and isinstance(data_url, str) and data_url.startswith('data:'):
                print(f"✓ Blob URL successfully converted to data URL")
                # Use the existing method to save base64 images
                return self.save_base64_image(data_url, filename_hash, chat_name)
            else:
                print(f"✗ Blob URL conversion failed: {data_url}")
                return None
                
        except Exception as e:
            print(f"Error downloading blob URL {blob_url}: {e}")
            return None
            
    def extract_attachments_from_message(self, message_element):
        attachments = []
        attachment_selectors = [
            '[data-tid="message-attachment"]',
            '.message-attachment',
            '.attachment-item',
            'a[href*="sharepoint"]',
            'a[href*="onedrive"]',
            '[data-tid*="attachment"]'
        ]
        for selector in attachment_selectors:
            try:
                attachment_elements = message_element.find_elements(By.CSS_SELECTOR, selector)
                for attachment_elem in attachment_elements:
                    attachment_info = {
                        'name': attachment_elem.text or attachment_elem.get_attribute('title') or 'Unknown',
                        'url': attachment_elem.get_attribute('href') or attachment_elem.get_attribute('src') or '',
                        'type': 'attachment',
                        'size': attachment_elem.get_attribute('data-size') or ''
                    }
                    attachments.append(attachment_info)
            except Exception as e:
                print(f"Error extracting attachments: {e}")
                continue
        return attachments

    def extract_messages_from_chat(self, chat_name="Unknown"):
        try:
            print(f"Extracting messages from chat: {chat_name}")
            # Set the current chat name for image naming
            self.current_chat_name = chat_name
            time.sleep(2)
            chat_container_selectors = [
                '[data-tid="chat-messages-container"]',
                '[data-tid="message-list"]',
                '.chat-messages',
                '.message-list-container',
                '[data-tid="chat-pane-runway"]'
            ]
            chat_container = None
            for selector in chat_container_selectors:
                try:
                    chat_container = self.driver.find_element(By.CSS_SELECTOR, selector)
                    print(f"✓ Found chat container with selector: {selector}")
                    break
                except:
                    continue
            if not chat_container:
                print("⚠ Chat container not found, using document body")
                chat_container = self.driver.find_element(By.TAG_NAME, "body")
            all_messages = self.scroll_to_load_all_messages_with_accumulation(chat_container, chat_name)
            print(f"✓ {len(all_messages)} unique messages successfully accumulated")
            if self.download_images:
                total_images = sum(len(msg.get('images', [])) for msg in all_messages)
                print(f"✓ {total_images} images found, {len(self.downloaded_images)} downloaded")
            return all_messages
        except Exception as e:
            print(f"Error extracting messages: {e}")
            return []

    def get_chat_names(self, chat_items):
        """Extracts the names of chats from the chat elements."""
        chat_names = []
        for i, chat_item in enumerate(chat_items, 1):
            chat_name = f"Chat_{i}"
            try:
                name_elem = chat_item.find_element(By.CSS_SELECTOR, 'div, span')
                potential_name = name_elem.get_attribute('innerText') or name_elem.text
                if potential_name and len(potential_name.strip()) > 0:
                    chat_name = potential_name.strip()[:50]
            except:
                pass
            chat_names.append(chat_name)
        return chat_names

    def display_chat_selection(self, chat_names):
        """Shows the available chats to the user and lets them select."""
        # If auto_select_all is enabled, automatically select all chats
        if self.auto_select_all:
            print("\nAutomatic selection of all chats enabled.")
            return list(range(len(chat_names)))
            
        print("\n" + "=" * 60)
        print("AVAILABLE CHATS")
        print("=" * 60)
        for i, name in enumerate(chat_names, 1):
            print(f"{i}. {name}")
        
        print("\nSelect an option:")
        print("1. Select chats from the list")
        print("2. Search for chats")
        option = input("Option (1/2): ").strip()
        
        if option == "2":
            print("Starting chat search...")
            return "search"
            
        print("\nSelect chats to scrape:")
        print("- Individual chats: '1,3,5'")
        print("- Range: '5-7'")
        print("- Combined: '1,3,5-7'")
        print("- All chats: 'all'")
        selection = input("> ").strip().lower()
        
        if selection == 'alle' or selection == 'all':
            return list(range(len(chat_names)))
        
        try:
            # Process input like "1,3,5-7"
            selected_indices = []
            parts = selection.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    # Range selection like "5-7"
                    start, end = map(int, part.split('-'))
                    selected_indices.extend(range(start-1, end))
                else:
                    # Single number
                    selected_indices.append(int(part) - 1)
            
            # Filter invalid indices
            valid_indices = [i for i in selected_indices if 0 <= i < len(chat_names)]
            if not valid_indices:
                print("No valid chats selected. All chats will be scraped.")
                return list(range(len(chat_names)))
            return valid_indices
        except Exception as e:
            print(f"Error processing input: {e}")
            print("Invalid input. All chats will be scraped.")
            return list(range(len(chat_names)))

    def process_all_chats(self):
        try:
            if not self.navigate_to_chats():
                return False
            chat_items = self.get_chat_list()
            if not chat_items:
                print("No chats found to process")
                return False
            
            # Extract chat names and show selection options
            chat_names = self.get_chat_names(chat_items)
            
            # If auto_select_all is enabled, filter out already exported chats
            if self.auto_select_all:
                print("\nAutomatic selection of all chats enabled.")
                print("Checking for already exported chats...")
                
                remaining_chats = []
                remaining_indices = []
                
                for i, chat_name in enumerate(chat_names):
                    if self.is_chat_already_exported(chat_name):
                        print(f"⏭ Skipping already exported chat: {chat_name}")
                    else:
                        remaining_chats.append(chat_name)
                        remaining_indices.append(i)
                
                if not remaining_chats:
                    print("✓ All chats have already been exported!")
                    return True
                
                print(f"📋 Found {len(remaining_chats)} new chats to export (skipped {len(chat_names) - len(remaining_chats)} already exported)")
                selected_indices = remaining_indices
            else:
                selected_indices = self.display_chat_selection(chat_names)
            
            # Check if the user selected the search option
            if selected_indices == "search":
                print("\nStarting chat search...")
                chat_items = self.search_chats()
                if not chat_items:
                    print("No chats found in search")
                    return False
                
                # Extract chat names from search results
                chat_names = self.get_chat_names(chat_items)
                
                # Show the found chats and let the user select
                print("\n" + "=" * 60)
                print("FOUND CHATS")
                print("=" * 60)
                for i, name in enumerate(chat_names, 1):
                    print(f"{i}. {name}")
                
                print("\nSelect chats to scrape:")
                print("- Individual chats: '1,3,5'")
                print("- Range: '5-7'")
                print("- Combined: '1,3,5-7'")
                print("- All chats: 'all'")
                selection = input("> ").strip().lower()
                
                if selection == 'alle' or selection == 'all':
                    selected_indices = list(range(len(chat_names)))
                else:
                    try:
                        # Process input like "1,3,5-7"
                        selected_indices = []
                        parts = selection.split(',')
                        for part in parts:
                            part = part.strip()
                            if '-' in part:
                                # Range selection like "5-7"
                                start, end = map(int, part.split('-'))
                                selected_indices.extend(range(start-1, end))
                            else:
                                # Single number
                                selected_indices.append(int(part) - 1)
                        
                        # Filter invalid indices
                        selected_indices = [i for i in selected_indices if 0 <= i < len(chat_names)]
                        if not selected_indices:
                            print("No valid chats selected. All chats will be scraped.")
                            selected_indices = list(range(len(chat_names)))
                    except Exception as e:
                        print(f"Error processing input: {e}")
                        print("Invalid input. All chats will be scraped.")
                        selected_indices = list(range(len(chat_names)))
            
            # Filter the selected chat elements
            selected_chats = [chat_items[i] for i in selected_indices]
            
            print(f"\nStarting enhanced processing of {len(selected_chats)} selected chats...")
            for i, chat_item in enumerate(selected_chats, 1):
                try:
                    chat_index = selected_indices[i-1]
                    chat_name = chat_names[chat_index]
                    print(f"\n--- Chat {i}/{len(selected_chats)} ({chat_name}) ---")
                    
                    # KNOWN BUG: Scraping may stop in longer chats for unknown reason
                    ActionChains(self.driver).move_to_element(chat_item).click().perform()
                    time.sleep(3)
                    messages = self.extract_messages_from_chat(chat_name)
                    
                    # Add messages to the total list
                    self.chat_data.extend(messages)
                    
                    # Immediately save the data for this chat
                    self.save_chat_data(chat_name, messages)
                    
                    print(f"✓ Chat '{chat_name}' processed: {len(messages)} messages")
                except Exception as e:
                    print(f"✗ Error with chat {i}: {e}")
                    continue
            return True
        except Exception as e:
            print(f"Error processing chats: {e}")
            return False

    def sanitize_filename(self, filename):
        """Sanitize a string to be used as a filename."""
        # Replace invalid filename characters with underscores
        invalid_chars = '<>:"/\\|?*\n\r\t'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        # Remove any remaining whitespace characters that could cause issues
        filename = ' '.join(filename.split())
        # Ensure the filename is not too long
        if len(filename) > 100:
            filename = filename[:97] + '...'
        # Ensure the filename is not empty
        if not filename or filename.isspace():
            filename = "Unnamed_Chat"
        return filename.strip()
        
    def save_chat_data(self, chat_name, messages):
        """Saves the data of a single chat immediately after processing."""
        if not messages:
            print(f"No data to save for chat '{chat_name}'")
            return
            
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = self.sanitize_filename(chat_name)
        
        # Save chat as JSON file
        json_file = os.path.join(self.output_dir, f"{safe_name}_{timestamp}.json")
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(messages, f, indent=2, ensure_ascii=False)
        
        # Save chat as CSV file
        csv_file = os.path.join(self.output_dir, f"{safe_name}_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            fieldnames = ['message_id', 'author', 'timestamp', 'content',
                         'images_count', 'attachments_count', 'extracted_at']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for msg in messages:
                writer.writerow({
                    'message_id': msg.get('message_id', ''),
                    'author': msg.get('author', ''),
                    'timestamp': msg.get('timestamp', ''),
                    'content': msg.get('content', ''),
                    'images_count': len(msg.get('images', [])),
                    'attachments_count': len(msg.get('attachments', [])),
                    'extracted_at': msg.get('extracted_at', '')
                })
                
        print(f"✓ Chat '{chat_name}' with {len(messages)} messages saved:")
        print(f"  - JSON: {json_file}")
        print(f"  - CSV: {csv_file}")
        
        if self.download_images:
            image_count = sum(len(msg.get('images', [])) for msg in messages)
            print(f"  - Images found: {image_count}")

    def save_data(self):
        """Saves summary files for all processed chats."""
        if not self.chat_data:
            print("No data to save")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save combined CSV file for all chats
        csv_file = os.path.join(self.output_dir, f"teams_export_{timestamp}.csv")
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            if self.chat_data:
                fieldnames = ['chat_name', 'message_id', 'author', 'timestamp', 'content',
                             'images_count', 'attachments_count', 'extracted_at']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for msg in self.chat_data:
                    writer.writerow({
                        'chat_name': msg.get('chat_name', ''),
                        'message_id': msg.get('message_id', ''),
                        'author': msg.get('author', ''),
                        'timestamp': msg.get('timestamp', ''),
                        'content': msg.get('content', ''),
                        'images_count': len(msg.get('images', [])),
                        'attachments_count': len(msg.get('attachments', [])),
                        'extracted_at': msg.get('extracted_at', '')
                    })
        
        # Save combined JSON file for compatibility
        combined_json_file = os.path.join(self.output_dir, f"teams_export_{timestamp}.json")
        with open(combined_json_file, 'w', encoding='utf-8') as f:
            json.dump(self.chat_data, f, indent=2, ensure_ascii=False)
        
        # Save image summary if images were downloaded
        if self.download_images:
            image_summary = {
                'total_messages': len(self.chat_data),
                'total_images_found': sum(len(msg.get('images', [])) for msg in self.chat_data),
                'total_images_downloaded': len(self.downloaded_images),
                'images_directory': self.images_dir,
                'export_timestamp': timestamp
            }
            summary_file = os.path.join(self.output_dir, f"image_summary_{timestamp}.json")
            with open(summary_file, 'w', encoding='utf-8') as f:
                json.dump(image_summary, f, indent=2, ensure_ascii=False)
        
        print(f"\n✓ Summary data saved:")
       # print(f"  - Combined JSON: {combined_json_file}")
        print(f"  - Combined CSV: {csv_file}")
        print(f"  - Total messages: {len(self.chat_data)}")
        if self.download_images:
            print(f"  - Downloaded images: {len(self.downloaded_images)}")
            print(f"  - Images directory: {self.images_dir}")
            print(f"  - Image summary: {summary_file}")

    def is_chat_already_exported(self, chat_name):
        """Check if a chat has already been exported by looking for existing JSON files."""
        # Extract only the actual chat name (first line before any newlines)
        actual_chat_name = chat_name.split('\n')[0].strip()
        safe_name = self.sanitize_filename(actual_chat_name)
        
        # Look for existing JSON files
        all_json_files = glob.glob(os.path.join(self.output_dir, "*.json"))
        
        for json_file in all_json_files:
            filename = os.path.basename(json_file)
            
            # Remove the download timestamp suffix (YYYYMMDD_HHMMSS.json)
            filename_without_download_timestamp = re.sub(r'_\d{8}_\d{6}\.json$', '', filename)
            
            # Remove the message time and content part (everything after the chat name)
            # Look for patterns like: _HH_MM_ or _HH.MM._ or _DD.MM._ etc.
            parts = filename_without_download_timestamp.split('_')
            
            # Find where a time/date pattern starts
            chat_name_parts = []
            for i, part in enumerate(parts):
                # Check for various time/date patterns:
                # 1. HH_MM format (e.g., "15_38")
                # 2. HH.MM. format (e.g., "16.05.")
                # 3. DD.MM. format (e.g., "02.01.")
                is_time_pattern = False
                
                if i < len(parts) - 1:
                    # Pattern 1: HH_MM (two consecutive numeric parts)
                    if (part.isdigit() and len(part) <= 2 and 
                        parts[i + 1].isdigit() and len(parts[i + 1]) <= 2):
                        is_time_pattern = True
                
                # Pattern 2: Contains dots and numbers (like "16.05." or "02.01.")
                if re.match(r'^\d{1,2}\.\d{1,2}\.?$', part):
                    is_time_pattern = True
                
                if is_time_pattern:
                    break
                
                chat_name_parts.append(part)
            
            # Reconstruct the chat name from the filename
            extracted_chat_name = '_'.join(chat_name_parts)
            
            # Compare with our current chat name (only the actual name part)
            if extracted_chat_name == safe_name:
                print(f"  Found existing export for '{actual_chat_name}': {filename}")
                return True
        
        return False

    def run(self):
        try:
            print("=" * 60)
            print("ENHANCED MICROSOFT TEAMS CHAT SCRAPER")
            print("WITH COMPLETE SCROLLING & IMAGE EXTRACTION")
            print("=" * 60)
            print("FEATURE: Chat selection enabled - You can choose which chats to scrape")
            print("FEATURE: Automatic msedgedriver management")
            print("FEATURE: Chat serach functionality. To backup specific chats, use the search option.")
            print("=" * 60)
            
            # Ensure the correct msedgedriver is available
            self.driver_path = self.ensure_edge_driver()
            
            if not self.setup_driver():
                return False
            if not self.navigate_to_teams():
                return False
            if not self.wait_for_sso_login():
                return False
            if not self.process_all_chats():
                return False
            self.save_data()
            print("\n" + "=" * 60)
            print("ENHANCED SCRAPING SUCCESSFULLY COMPLETED!")
            print("=" * 60)
            return True
        except Exception as e:
            print(f"\nError during execution: {e}")
            return False
        finally:
            if self.driver:
                print("\nClosing browser...")
                self.driver.quit()

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Microsoft Teams Chat Scraper")
    parser.add_argument("--output-dir", default="teams_complete_export", help="Output directory for scraped data")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--no-images", action="store_true", help="Don't download images")
    parser.add_argument("--auto-select-all", action="store_true", help="Automatically select all chats (no user prompt)")
    
    args = parser.parse_args()
    
    scraper = TeamsCollector(
        output_dir=args.output_dir,
        headless=args.headless,
        download_images=not args.no_images,
        auto_select_all=args.auto_select_all
    )
    scraper.run()










