import pandas as pd
import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import json
from bs4 import BeautifulSoup, Tag, NavigableString
from urllib.parse import unquote, urlparse, parse_qs
import os

class ZameenScraper:
    def __init__(self, headless=False):
        """Initialize the Zameen scraper with Chrome driver"""
        self.driver = None
        self.setup_driver(headless)
        
    def setup_driver(self, headless=False):
        """Setup Chrome driver with appropriate options"""
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            print("Chrome driver initialized successfully")
        except Exception as e:
            print(f"Error initializing Chrome driver: {e}")
            print("Please make sure ChromeDriver is installed and in your PATH")
            raise
    
    def _wait_for_page_load(self, timeout=10):
        """Wait for page to load completely"""
        try:
            WebDriverWait(self.driver, timeout).until(
                lambda driver: driver.execute_script("return document.readyState") == "complete"
            )
            time.sleep(1)  # Additional wait for dynamic content
        except TimeoutException:
            print("Page load timeout, continuing anyway...")
    
    def _find_search_inputs(self):
        """Find search input fields on the page"""
        try:
            # Wait for page to load
            self._wait_for_page_load()
            
            # Common selectors for Zameen search inputs
            input_selectors = [
                "input[placeholder*='society']",
                "input[placeholder*='area']",
                "input[placeholder*='search']",
                "input[placeholder*='location']",
                "input[type='text']",
                "input[type='search']",
                ".search-input input",
                ".form-control",
                "[data-testid*='search'] input",
                "[class*='search'] input"
            ]
            
            found_inputs = []
            
            for selector in input_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        if element.is_displayed() and element not in found_inputs:
                            found_inputs.append(element)
                            print(f"Found input: {selector} - placeholder: '{element.get_attribute('placeholder')}'")
                except Exception as e:
                    continue
            
            # Remove duplicates while preserving order
            unique_inputs = []
            for inp in found_inputs:
                if inp not in unique_inputs:
                    unique_inputs.append(inp)
            
            print(f"Total unique inputs found: {len(unique_inputs)}")
            return unique_inputs[:2] if len(unique_inputs) >= 2 else unique_inputs
            
        except Exception as e:
            print(f"Error finding search inputs: {e}")
            return []

    def _type_and_select_suggestion(self, input_element, text_to_type, wait_time=5):
        """Type text and select the FIRST suggestion that appears"""
        try:
            print(f"Typing '{text_to_type}' in input field...")
            
            # Clear and type with more deliberate actions
            input_element.clear()
            time.sleep(0.1)
            
            # Type character by character for better detection
            for char in text_to_type:
                input_element.send_keys(char)
                time.sleep(0.1)
            
            # Wait longer for suggestions to appear
            print(f"Waiting {wait_time} seconds for suggestions...")
            time.sleep(wait_time)
            
            # Try multiple suggestion selectors
            suggestion_selectors = [
                ".suggestion-list li",
                ".dropdown-menu li", 
                ".autocomplete-suggestion",
                "[class*='suggestion']",
                "[class*='dropdown'] li",
                "[class*='autocomplete'] div",
                ".list-group-item",
                "ul li",
                "[role='option']",
                "div[class*='_3b7a06ea']",  # Zameen specific class
                "div[class*='_885847b8']",
                ".dropdown-item",
                "[data-testid*='suggestion']"
            ]
            
            suggestion_clicked = False  # Flag to track if we've successfully clicked
            
            
            for selector in suggestion_selectors:
                if suggestion_clicked:  # If already clicked, break out
                    break
                    
                try:
                    suggestions = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    visible_suggestions = [s for s in suggestions if s.is_displayed()]
                    
                    if visible_suggestions:
                        print(f"Found {len(visible_suggestions)} suggestions using selector: {selector}")
                        
                        # Always select the FIRST suggestion
                        first_suggestion = visible_suggestions[0]
                        suggestion_text = first_suggestion.text.strip()
                        print(f"First suggestion text: '{suggestion_text}'")
                        print(f"Will click the first suggestion regardless of content")
                        
                        # # Enhanced mouse click methods - more aggressive strategies
                        # click_methods = [
                        #     # Method 1: JavaScript click (most reliable for dynamic content)
                        #     {
                        #         'name': 'JavaScript click',
                        #         'action': lambda: (self.driver.execute_script("arguments[0].click();", first_suggestion),print("method1"))
                        #     }
                        #     # # Method 2: Action chains with explicit move and click
                        #     # {
                        #     #     'name': 'ActionChains move+click',
                        #     #     'action': lambda: ActionChains(self.driver).move_to_element(first_suggestion).pause(0.5).click().perform()
                        #     # },
                        #     # # Method 3: Scroll into view then JavaScript click
                        #     # {
                        #     #     'name': 'Scroll + JS click',
                        #     #     'action': lambda: (
                        #     #         self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first_suggestion),
                        #     #         time.sleep(0.3),
                        #     #         self.driver.execute_script("arguments[0].click();", first_suggestion)
                        #     #     )[-1]  # Return last value
                        #     # },
                        #     # # Method 4: Force click with mouse event dispatch
                        #     # {
                        #     #     'name': 'Dispatch mouse event',
                        #     #     'action': lambda: self.driver.execute_script("""
                        #     #         var element = arguments[0];
                        #     #         element.scrollIntoView({block: 'center'});
                        #     #         var rect = element.getBoundingClientRect();
                        #     #         var clickEvent = new MouseEvent('click', {
                        #     #             view: window,
                        #     #             bubbles: true,
                        #     #             cancelable: true,
                        #     #             clientX: rect.left + rect.width/2,
                        #     #             clientY: rect.top + rect.height/2
                        #     #         });
                        #     #         element.dispatchEvent(clickEvent);
                        #     #     """, first_suggestion)
                        #     # },
                        #     # # Method 5: Regular click as last resort
                        #     # {
                        #     #     'name': 'Regular click',
                        #     #     'action': lambda: first_suggestion.click()
                        #     # }
                        # ]
                        
                        # for method in click_methods:
                        try:
                                self.driver.execute_script("arguments[0].click();", first_suggestion)
                                print("Successfully clicked first suggestion using Java Click")
                                suggestion_clicked = True 
                                print(suggestion_clicked)
                                time.sleep(0.1)
                                # print(f"✓ Successfully clicked first suggestion using Java Click")
                                # suggestion_clicked = True  # Mark as successfully clicked
                                
                                    # print(f"Attempting {method['name']}...")
                                    # method['action']()
                                    # # time.sleep(4)
                                    # print(f"✓ Successfully clicked first suggestion using {method['name']}")
                                    # suggestion_clicked = True  # Mark as successfully clicked
                                    # time.sleep(1.5)  # Wait for page to respond
                                #     break  # Break out of click methods loop
                            
                        except Exception as e:
                                print(f"failed: {e}")
                                continue
                        
                        if suggestion_clicked:
                            break  # Break out of selector loop if clicked successfully
                                
                except Exception as e:
                    continue
            
            if suggestion_clicked:
                return True
            else:
                # No suggestions found at all
                print(f"No suggestions found or clickable for '{text_to_type}'")
                return False
            
        except Exception as e:
            print(f"Error in _type_and_select_suggestion: {e}")
            return False

    def _find_and_click_search_result(self, timeout=15):
        """Find and click on the first search result"""
        try:
            print("Looking for search results...")
            time.sleep(0.1)  # Wait for results to load
            
            # Selectors for search result items
            result_selectors = [
                "[class*='listing']",
                "[class*='property']",
                "[class*='result']",
                "[class*='card']",
                ".listing-item",
                ".property-card",
                ".search-result",
                "a[href*='/property/']",
                "div[data-testid*='listing']",
                "article",
                ".tile"
            ]
            
            for selector in result_selectors:
                try:
                    results = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    visible_results = [r for r in results if r.is_displayed()]
                    
                    if visible_results:
                        print(f"Found {len(visible_results)} results using selector: {selector}")
                        
                        # Try to click the first result
                        first_result = visible_results[0]
                        
                        # Try multiple click approaches
                        click_methods = [
                            lambda: first_result.click(),
                            lambda: self.driver.execute_script("arguments[0].click();", first_result),
                            lambda: ActionChains(self.driver).move_to_element(first_result).click().perform(),
                        ]
                        
                        for i, click_method in enumerate(click_methods):
                            try:
                                click_method()
                                print(f"Clicked search result using method {i+1}")
                                time.sleep(0.1)
                                return True
                            except Exception as e:
                                print(f"Click method {i+1} failed: {e}")
                                continue
                                
                except Exception as e:
                    continue
            
            print("No clickable search results found")
            return False
            
        except Exception as e:
            print(f"Error finding search result: {e}")
            return False

    def _find_and_click_location_button(self, timeout=10):
        """Find and click location/navigate button"""
        try:
            print("Looking for location/navigate button...")
            time.sleep(0.1)
            
            # Selectors for location/navigate buttons
            location_selectors = [
                "a[href*='maps.google']",
                "button:contains('Navigate')",
                "a:contains('Navigate')",
                "button:contains('Location')",
                "a:contains('Location')",
                "button:contains('Directions')",
                "a:contains('Directions')",
                "[class*='navigate']",
                "[class*='location']",
                "[class*='direction']",
                "[class*='maps']",
                "a[href*='google.com/maps']"
            ]
            
            # Also try to find by text content
            text_patterns = ['navigate', 'location', 'directions', 'view on map', 'get directions']
            
            # Try CSS selectors first
            for selector in location_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        if element.is_displayed():
                            try:
                                element.click()
                                print(f"✓ Clicked location button using selector: {selector}")
                                return True
                            except Exception as e:
                                print(f"Failed to click element: {e}")
                                continue
                except Exception as e:
                    continue
            
            # Try finding by text content using XPath
            for pattern in text_patterns:
                try:
                    xpath_selectors = [
                        f"//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{pattern}')]",
                        f"//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{pattern}')]",
                        f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{pattern}')]"
                    ]
                    
                    for xpath in xpath_selectors:
                        try:
                            elements = self.driver.find_elements(By.XPATH, xpath)
                            for element in elements:
                                if element.is_displayed():
                                    try:
                                        element.click()
                                        print(f"Clicked location button using text: {pattern}")
                                        return True
                                    except Exception as e:
                                        continue
                        except Exception as e:
                            continue
                            
                except Exception as e:
                    continue
            
            print("Location/navigate button not found")
            return False
            
        except Exception as e:
            print(f"Error finding location button: {e}")
            return False

    def _extract_coordinates_from_url(self, url):
        """Extract latitude and longitude from Google Maps URL"""
        try:
            url = unquote(url)
            print(f"Extracting coordinates from URL: {url[:100]}...")
            
            # Multiple regex patterns for different URL formats
            patterns = [
                r'@(-?\d+\.\d+),(-?\d+\.\d+)',  # @lat,lng
                r'!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)',  # !3dlat!4dlng
                r'destination=(-?\d+\.\d+)(?:,|%2C|%2c)(-?\d+\.\d+)',  # destination=lat,lng
                r'q=(-?\d+\.\d+)(?:,|%2C|%2c)(-?\d+\.\d+)',  # q=lat,lng
                r'place/(-?\d+\.\d+),(-?\d+\.\d+)',  # place/lat,lng
                r'll=(-?\d+\.\d+),(-?\d+\.\d+)',  # ll=lat,lng
            ]
            
            for pattern in patterns:
                match = re.search(pattern, url)
                if match:
                    lat, lng = float(match.group(1)), float(match.group(2))
                    if -90 <= lat <= 90 and -180 <= lng <= 180:
                        print(f"Extracted coordinates: {lat}, {lng}")
                        return lat, lng
            
            # Try parsing query parameters
            parsed = urlparse(url)
            params = parse_qs(parsed.query)
            
            for key in ['q', 'destination', 'll']:
                if key in params:
                    value = params[key][0]
                    parts = value.split(',')
                    if len(parts) >= 2:
                        try:
                            lat, lng = float(parts[0]), float(parts[1])
                            if -90 <= lat <= 90 and -180 <= lng <= 180:
                                print(f"Extracted coordinates from {key}: {lat}, {lng}")
                                return lat, lng
                        except ValueError:
                            continue
            
            print("Could not extract coordinates from URL")
            return None, None
            
        except Exception as e:
            print(f"Error extracting coordinates: {e}")
            return None, None

    def scrape_single_location(self, column_b_value, column_a_value):
        """Scrape coordinates for a single location
        
        Args:
            column_b_value: Value from Excel column B (typed in FIRST search bar)
            column_a_value: Value from Excel column A (typed in SECOND search bar)
        """
        try:
            print(f"\n=== Scraping: Column B (1st search)='{column_b_value}', Column A (2nd search)='{column_a_value}' ===")
            
            # Navigate to Zameen plot finder
            print("Opening Zameen plot finder...")
            self.driver.get("https://www.zameen.com/plotfinder/Karachi-30/")
            self._wait_for_page_load()
            
            # Take a screenshot for debugging
            try:
                self.driver.save_screenshot(f"step1_initial_page.png")
                print("Screenshot saved: step1_initial_page.png")
            except:
                pass
            
            # STEP 1: Find the FIRST (and initially ONLY) search input
            inputs = self._find_search_inputs()
            if len(inputs) < 1:
                raise Exception("No search input found on the page")
            
            first_input = inputs[0]
            print(f"Found first search input with placeholder: '{first_input.get_attribute('placeholder')}'")
            
            # STEP 2: Type Column B value in the FIRST search bar and select suggestion
            print(f"Step 1: Typing Column B value ('{column_b_value}') in the FIRST search bar...")
            if not self._type_and_select_suggestion(first_input, column_b_value):
                raise Exception(f"Failed to select suggestion for Column B: {column_b_value}")
            
            # STEP 3: Wait for the SECOND search bar to appear after clicking first suggestion
            print("Step 2: Waiting for SECOND search bar to appear after first selection...")
            time.sleep(0.1)  # Wait for second input to appear
            
            # STEP 4: Find the NEW second search bar that should have appeared
            print("Step 3: Looking for the NEW second search bar...")
            
            second_input = None
            max_attempts = 3
            
            for attempt in range(max_attempts):
                print(f"Attempt {attempt + 1} to find second search bar...")
                
                # Get all current inputs
                current_inputs = self._find_search_inputs()
                
                # Look for inputs that are different from the first one
                for inp in current_inputs:
                    if inp != first_input and inp.is_displayed():
                        second_input = inp
                        print(f"Found second input with placeholder: '{inp.get_attribute('placeholder')}'")
                        break
                
                # If not found, try broader search
                if not second_input:
                    all_inputs = self.driver.find_elements(By.CSS_SELECTOR, "input")
                    for inp in all_inputs:
                        if inp != first_input and inp.is_displayed():
                            input_type = inp.get_attribute('type') or ""
                            placeholder = inp.get_attribute('placeholder') or ""
                            # Check if it looks like a search input
                            if (input_type in ['text', 'search', ''] and 
                                ('search' in placeholder.lower() or 
                                 'location' in placeholder.lower() or 
                                 'area' in placeholder.lower() or 
                                 placeholder == "")):
                                second_input = inp
                                print(f"Found second input (broader search) with placeholder: '{placeholder}'")
                                break
                
                if second_input:
                    break
                    
                if attempt < max_attempts - 1:
                    print(f"Second input not found, waiting 2 more seconds...")
                    time.sleep(0.1)
            
            if not second_input:
                raise Exception("Second search bar did not appear after selecting first suggestion")
            
            # STEP 5: Type Column A value in the SECOND search bar and select suggestion
            print(f"Step 4: Typing Column A value ('{column_a_value}') in the SECOND search bar...")
            if not self._type_and_select_suggestion(second_input, column_a_value):
                raise Exception(f"Failed to select suggestion for Column A: {column_a_value}")
            
            time.sleep(0.1)  # Wait for search results
            
            # Take screenshot after search
            try:
                self.driver.save_screenshot("step2_after_search.png")
                print("Screenshot saved: step2_after_search.png")
            except:
                pass
            
            # Find and click location/navigate button
            print("Step 4: Looking for location button...")
            original_windows = self.driver.window_handles
            
            if not self._find_and_click_location_button():
                raise Exception("Location/navigate button not found")
            
            time.sleep(0.1)  # Wait for navigation
            
            # Handle new tab if opened
            new_windows = self.driver.window_handles
            if len(new_windows) > len(original_windows):
                print("New tab opened, switching to it...")
                self.driver.switch_to.window(new_windows[-1])
            
            # Wait for Google Maps to load
            try:
                WebDriverWait(self.driver, 15).until(
                    lambda d: "google.com/maps" in d.current_url.lower() or "maps.google" in d.current_url.lower()
                )
            except TimeoutException:
                print("Timeout waiting for Google Maps, checking current URL anyway...")
            
            # Extract coordinates
            current_url = self.driver.current_url
            print(f"Current URL: {current_url[:150]}...")
            
            lat, lng = self._extract_coordinates_from_url(current_url)
            
            if lat is None or lng is None:
                raise Exception("Could not extract coordinates from Google Maps URL")
            
            print(f"SUCCESS: Latitude={lat}, Longitude={lng}")
            return {
                'success': True,
                'latitude': lat,
                'longitude': lng,
                'maps_url': current_url
            }
            
        except Exception as e:
            error_msg = str(e)
            print(f"ERROR: {error_msg}")
            
            # Take error screenshot
            try:
                self.driver.save_screenshot("error_screenshot.png")
                print("Error screenshot saved: error_screenshot.png")
            except:
                pass
                
            return {
                'success': False,
                'error': error_msg,
                'latitude': None,
                'longitude': None,
                'maps_url': None
            }

    def process_excel_file(self, file_path, area_col="B", location_col="A", 
                          lat_col="C", lng_col="D", url_col="E", 
                          output_file=None, has_header=False):
        """Process Excel file with locations
        
        Args:
            area_col: Column with area/society names (typed in FIRST search bar)
            location_col: Column with specific locations (typed in SECOND search bar)
        """
        try:
            print(f"Processing Excel file: {file_path}")
            
            # Read Excel file
            df = pd.read_excel(file_path, header=0 if has_header else None)
            print(f"Loaded {len(df)} rows from Excel file")
            
            # Convert column letters to indices
            def col_to_index(col_letter):
                return ord(col_letter.upper()) - ord('A')
            
            area_idx = col_to_index(area_col)
            location_idx = col_to_index(location_col)
            lat_idx = col_to_index(lat_col)
            lng_idx = col_to_index(lng_col)
            url_idx = col_to_index(url_col)
            
            # Ensure dataframe has enough columns
            max_col_idx = max(area_idx, location_idx, lat_idx, lng_idx, url_idx)
            while len(df.columns) <= max_col_idx:
                df[len(df.columns)] = None
            
            results = []
            successful = 0
            failed = 0
            
            for idx, row in df.iterrows():
                print(f"\n{'='*50}")
                print(f"Processing row {idx + 1} of {len(df)}")
                
                # Get area and location values (CORRECTED ORDER)
                area_val = row.iloc[area_idx] if pd.notna(row.iloc[area_idx]) else ""        # Column B - First search
                location_val = row.iloc[location_idx] if pd.notna(row.iloc[location_idx]) else ""  # Column A - Second search
                
                print(f"Column B (First search): '{area_val}'")
                print(f"Column A (Second search): '{location_val}'")
                
                # Skip empty rows
                if not area_val and not location_val:
                    print("Skipping empty row")
                    df.iloc[idx, lat_idx] = "Empty row"
                    df.iloc[idx, lng_idx] = "Empty row"
                    df.iloc[idx, url_idx] = "N/A"
                    continue
                
                # Scrape this location (B first, then A)
                result = self.scrape_single_location(str(area_val), str(location_val))
                
                # Update dataframe
                if result['success']:
                    df.iloc[idx, lat_idx] = result['latitude']
                    df.iloc[idx, lng_idx] = result['longitude']
                    df.iloc[idx, url_idx] = result['maps_url']
                    successful += 1
                else:
                    df.iloc[idx, lat_idx] = "Location not found"
                    df.iloc[idx, lng_idx] = "Location not found"
                    df.iloc[idx, url_idx] = "N/A"
                    failed += 1
                
                results.append({
                    'row': idx + 1,
                    'area': area_val,
                    'location': location_val,
                    **result
                })
                
                # Save progress every 5 rows
                if (idx + 1) % 5 == 0:
                    temp_output = output_file or file_path.replace('.xlsx', '_progress.xlsx')
                    df.to_excel(temp_output, index=False, header=has_header)
                    print(f"Progress saved to: {temp_output}")
                
                # Polite delay between requests
                time.sleep(0.1)
            
            # Save final results
            final_output = output_file or file_path.replace('.xlsx', '_with_coordinates.xlsx')
            df.to_excel(final_output, index=False, header=has_header)
            
            print(f"\n{'='*50}")
            print("SCRAPING COMPLETED!")
            print(f"Final results saved to: {final_output}")
            print(f"Total processed: {len(results)}")
            print(f"Successful: {successful}")
            print(f"Failed: {failed}")
            print(f"Success rate: {(successful / len(results) * 100):.1f}%" if results else "0%")
            
            return pd.DataFrame(results)
            
        except Exception as e:
            print(f"Error processing Excel file: {e}")
            raise

    def close(self):
        """Close the browser"""
        if self.driver:
            try:
                self.driver.quit()
                print("Browser closed successfully")
            except Exception as e:
                print(f"Error closing browser: {e}")

def main():
    """Main function"""
    # Configuration - UPDATE THESE PATHS AND SETTINGS
    EXCEL_FILE_PATH = "zameen.xlsx"  # Your Excel file path
    AREA_COLUMN = "B"        # Column B: Area/society names (FIRST search bar)
    LOCATION_COLUMN = "A"    # Column A: Specific locations (SECOND search bar)
    LAT_OUTPUT_COLUMN = "C"  # Where to write latitude
    LNG_OUTPUT_COLUMN = "D"  # Where to write longitude
    URL_OUTPUT_COLUMN = "E"  # Where to write Google Maps URL
    OUTPUT_FILE = "addresses_with_coordinates.xlsx"       # None = auto-generate name, or specify custom path
    HAS_HEADER = False       # True if first row contains headers
    HEADLESS_MODE = False    # True to run without showing browser window
    
    # Check if file exists
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"ERROR: Excel file not found at: {EXCEL_FILE_PATH}")
        print("Please update the EXCEL_FILE_PATH variable with the correct path.")
        return
    
    scraper = None
    try:
        print("Starting Zameen Property Scraper...")
        print(f"Excel file: {EXCEL_FILE_PATH}")
        print(f"Column B (First search): {AREA_COLUMN}, Column A (Second search): {LOCATION_COLUMN}")
        print(f"Output columns: Lat={LAT_OUTPUT_COLUMN}, Lng={LNG_OUTPUT_COLUMN}, URL={URL_OUTPUT_COLUMN}")
        print(f"Search sequence: Column B -> First search bar -> Column A -> Second search bar")
        print("NOTE: The script will ALWAYS click the FIRST suggestion that appears, regardless of its content")
        
        # Initialize scraper
        scraper = ZameenScraper(headless=HEADLESS_MODE)
        
        # Process the Excel file
        results = scraper.process_excel_file(
            file_path=EXCEL_FILE_PATH,
            area_col=AREA_COLUMN,
            location_col=LOCATION_COLUMN,
            lat_col=LAT_OUTPUT_COLUMN,
            lng_col=LNG_OUTPUT_COLUMN,
            url_col=URL_OUTPUT_COLUMN,
            output_file=OUTPUT_FILE,
            has_header=HAS_HEADER
        )
        
        print("\nScraping completed successfully!")
        
    except KeyboardInterrupt:
        print("\nScraping interrupted by user")
    except Exception as e:
        print(f"\nError occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if scraper:
            scraper.close()

if __name__ == "__main__":
    main()

"""
INSTALLATION REQUIREMENTS:

1. Install required packages:
   pip install selenium pandas openpyxl beautifulsoup4

2. Install ChromeDriver:
   - Download from: https://chromedriver.chromium.org/
   - Make sure it matches your Chrome version
   - Add to system PATH or place in script directory

3. Update the configuration in main() function:
   - EXCEL_FILE_PATH: Path to your Excel file
   - Column letters for input and output
   - Set HAS_HEADER=True if first row contains headers

FEATURES:
[+] SIMPLIFIED: Always clicks the FIRST suggestion that appears
[+] No complex matching logic - just selects the first option
[+] Improved element finding with multiple strategies
[+] Better error handling and debugging
[+] Progress saving every 5 rows
[+] Screenshots for debugging
[+] More robust coordinate extraction
[+] Detailed logging and status messages
[+] Fallback strategies for different scenarios
[+] Support for various Zameen page layouts

BEHAVIOR CHANGE:
- The script will now ALWAYS click the first suggestion that appears
- It doesn't try to match the text content anymore
- This should work for cases like "Plot# B157" where B157 was your input
- Much simpler and more reliable approach
"""
