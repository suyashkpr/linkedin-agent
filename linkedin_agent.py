#!/usr/bin/env python3

import json
import time
import random
import pickle
import os
from typing import List, Dict, Optional
from dataclasses import dataclass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import logging

from gsheet import CompanyTracker

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@dataclass
class Person:
    name: str
    title: str
    company: str
    profile_url: str

@dataclass
class Config:
    max_connections_per_company: int = 3
    connection_message_template: str = ""
    delay_between_requests: tuple = (5, 10)
    headless: bool = False
    waiting_time: int = 3

class LinkedInAgent:
    def __init__(self, config: Config):
        self.config = config
        self.driver = None
        self.cookies_file = "linkedin_cookies.pkl"
        self.setup_driver()
        self.tracker = CompanyTracker(
            credentials_file="gcloud.json",
            spreadsheet_name="Shruti Company Reachout",
            worksheet_name="Sheet1"  # Optional, uses first sheet if not provided
        )
    
    def setup_driver(self):
        chrome_options = Options()
        if self.config.headless:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(self.config.waiting_time)
    
    def save_cookies(self):
        try:
            with open(self.cookies_file, 'wb') as f:
                pickle.dump(self.driver.get_cookies(), f)
            logger.info("Cookies saved successfully")
        except Exception as e:
            logger.error(f"Failed to save cookies: {str(e)}")
    
    def load_cookies(self):
        try:
            if os.path.exists(self.cookies_file):
                self.driver.get("https://www.linkedin.com")
                with open(self.cookies_file, 'rb') as f:
                    cookies = pickle.load(f)
                for cookie in cookies:
                    self.driver.add_cookie(cookie)
                logger.info("Cookies loaded successfully")
                return True
            return False
        except Exception as e:

            logger.error(f"Failed to load cookies: {str(e)}")
            return False
    
    def is_logged_in(self) -> bool:
        try:
            self.driver.get("https://www.linkedin.com/feed")
            time.sleep(self.config.waiting_time)
            return "linkedin.com/feed" in self.driver.current_url
        except Exception as e:
            logger.error(f"Error checking login status: {str(e)}")
            return False
    
    def login_to_linkedin(self, username: str, password: str) -> bool:
        if self.load_cookies() and self.is_logged_in():
            logger.info("Already logged in using saved cookies")
            return True
        
        try:
            self.driver.get("https://www.linkedin.com/login")
            
            username_field = self.driver.find_element(By.ID, "username")
            password_field = self.driver.find_element(By.ID, "password")
            
            username_field.send_keys(username)
            password_field.send_keys(password)
            
            login_button = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            login_button.click()
            
            WebDriverWait(self.driver, 10).until(
                EC.url_contains("linkedin.com/feed")
            )
            
            self.save_cookies()
            logger.info("Successfully logged in to LinkedIn")
            return True
            
        except Exception as e:
            logger.error(f"Failed to login: {str(e)}")
            return False
    
    def search_company_people(self, company_name: str) -> List[Person]:
        search_query = f"{company_name} product"
        people = []
        
        try:
            search_url = f"https://www.linkedin.com/search/results/people/?keywords={search_query.replace(' ', '%20')}"
            self.driver.get(search_url)
            
            time.sleep(self.config.waiting_time)
            
            selectors_to_try = [
                # New LinkedIn layout selectors
                '[data-view-name="search-entity-result-universal-template"]',
                '.entity-result__content',
                '[data-test-id="search-result"]',
                # Fallback selectors
                '.search-result__wrapper',
                '.reusable-search__result-container',
                '[data-anonymize="person-name"]'
            ]
            
            search_results = []
            for selector in selectors_to_try:
                search_results = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if search_results:
                    logger.info(f"Found results using selector: {selector}")
                    break
            
            if not search_results:
                logger.warning("No search results found with any selector")
                return []
            
            for result in search_results[:10]:
                try:
                    # Try multiple methods to find name and profile URL
                    name = None
                    profile_url = None
                    
                    # Method 1: Look for links with aria-label containing name
                    links = result.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        href = link.get_attribute("href")
                        if href and "/in/" in href and "linkedin.com" in href:
                            profile_url = href
                            name = link.get_attribute("aria-label") or link.text.strip()
                            if name and not name.startswith("View"):
                                # Split by newline and get first part (clean name)
                                name = name.split('\n')[0].strip()
                                # If name has multiple words, just take the first name
                                name = name.split(' ')[0].strip()
                                break
                    
                    title = ""
                    
                    # Look for any element that might contain job title
                    potential_title_elements = []
                    
                    # Try common patterns for job titles
                    # potential_title_elements.extend(result.find_elements(By.TAG_NAME, "p"))
                    potential_title_elements.extend(result.find_elements(By.TAG_NAME, "div"))
                    potential_title_elements.extend(result.find_elements(By.TAG_NAME, "span"))
                    
                    for element in potential_title_elements:
                        text = element.text.strip()
                        if (text and 
                            len(text) > 5 and len(text) < 100 and
                            any(keyword in text.lower() for keyword in ['product', 'manager', 'director', 'lead', 'engineer', 'pm', 'senior', 'staff']) and
                            not any(skip_word in text.lower() for skip_word in ['view', 'connect', 'message', 'follow', 'linkedin'])):
                            title = text
                            break
                    
                    if name and profile_url and any(keyword in title.lower() for keyword in ['product', 'pm', 'manager', 'lead']):
                        person = Person(
                            name=name,
                            title=title,
                            company=company_name,
                            profile_url=profile_url
                        )
                        people.append(person)
                        logger.info(f"Found person: {name} - {title}")
                        
                        if len(people) >= self.config.max_connections_per_company:
                            break
                            
                except Exception as e:
                    logger.warning(f"Error parsing search result: {str(e)}")
                    continue
            
            logger.info(f"Found {len(people)} people for {company_name}")
            return people
            
        except Exception as e:
            logger.error(f"Error searching for people at {company_name}: {str(e)}")
            return []
    
    def extract_latest_company(self) -> Optional[str]:
        """Extract the latest company from the person's profile experience section"""
        try:
            # Look for "a" tags with data-field="experience_company_logo"
            all_links = self.driver.find_elements(By.TAG_NAME, "a")
            
            for link in all_links:
                data_field = link.get_attribute("data-field")
                if data_field == "experience_company_logo":
                    href = link.get_attribute("href")
                    if href:
                        # Store current URL to navigate back
                        current_url = self.driver.current_url
                        
                        # Navigate to company page
                        self.driver.get(href)
                        time.sleep(2)
                        
                        # Find h1 element and get company name
                        h1_elements = self.driver.find_elements(By.TAG_NAME, "h1")
                        for h1 in h1_elements:
                            company_text = h1.text.strip()
                            if (company_text and
                                len(company_text) > 2 and len(company_text) < 80 and
                                any(char.isupper() for char in company_text)):
                                logger.info(f"Found company from experience_company_logo: {company_text}")
                                
                                # Navigate back to original page
                                self.driver.get(current_url)
                                time.sleep(2)
                                
                                return company_text
                        
                        # Navigate back even if no h1 found
                        self.driver.get(current_url)
                        time.sleep(2)
                    
        except Exception as e:
            logger.warning(f"Error extracting company: {str(e)}")
        
        return None
    
    def _find_button_by_text(self, buttons, search_terms, button_type="button"):
        """Helper method to find a button by searching for text in aria-label, text, or innerHTML"""
        for button in buttons:
            aria_label = (button.get_attribute("aria-label") or "").lower()
            button_text = (button.text or "").lower()
            inner_html = (button.get_attribute("innerHTML") or "").lower()
            
            for term in search_terms:
                if (term in aria_label or term in button_text or term in inner_html) and \
                   button.is_enabled() and button.is_displayed():
                    logger.info(f"Found {button_type} button: aria-label='{button.get_attribute('aria-label')}', text='{button.text}'")
                    return button
        return None
    
    def _find_connect_button_in_more_menu(self):
        """Try to find Connect button by clicking More button first"""
        all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
        
        # Find More button
        more_button = self._find_button_by_text(all_buttons, ['more'], "More")
        
        if more_button:
            try:
                more_button.click()
                time.sleep(1)  # Wait for dropdown/menu to appear
                
                # Look for Connect button in the expanded menu (check both button and div elements)
                menu_elements = self.driver.find_elements(By.TAG_NAME, "button") + self.driver.find_elements(By.TAG_NAME, "div")
                connect_button = self._find_button_by_text(menu_elements, ['connect'], "Connect (in More menu)")
                
                return connect_button
                
            except Exception as e:
                logger.warning(f"Failed to click More button or find Connect in menu: {str(e)}")
                return None
        
        return None
    
    def _try_click_button(self, button, button_type="button"):
        """Try to click a button with fallback for click intercepted exceptions"""
        try:
            button.click()
            return True
        except ElementClickInterceptedException:
            logger.warning(f"{button_type} button click intercepted, trying JavaScript click")
            try:
                self.driver.execute_script("arguments[0].click();", button)
                return True
            except Exception as e:
                logger.error(f"Failed to click {button_type} button with JavaScript: {str(e)}")
                return False
        except Exception as e:
            logger.error(f"Failed to click {button_type} button: {str(e)}")
            return False
    
    def send_connection_request(self, person: Person, company_id: str = None) -> bool:
        try:
            # Check if person already exists in tracker for this company
            if company_id:
                try:
                    company_data = self.tracker.get_company_row(company_id)
                    if company_data:
                        # Check if person's URL already exists in the people list
                        for existing_person in company_data['people']:
                            # Extract URL from HYPERLINK formula if present
                            person_data = existing_person['data']
                            if person.profile_url in person_data:
                                logger.info(f"Person {person.name} ({person.profile_url}) already exists for company {company_id}, skipping")
                                return False
                except Exception as e:
                    logger.warning(f"Could not check existing people for company {company_id}: {str(e)}")
            
            self.driver.get(person.profile_url)
            time.sleep(2)
            
            # Extract actual company from latest experience on profile page
            actual_company = self.extract_latest_company()
            if actual_company:
                # Update person's company with actual company from profile
                person.company = actual_company
                logger.info(f"Updated company for {person.name}: {actual_company}")
            else:
                logger.error(f"Could not extract company for {person.name}, skipping connection request")
                return False

            
            # Debug: Let's see what buttons are actually on the page
            logger.info("=== DEBUG: Looking for Connect button ===")
            all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
            logger.info(f"Found {len(all_buttons)} buttons on the page")
            
            for i, button in enumerate(all_buttons[:10]):  # Log first 10 buttons
                aria_label = button.get_attribute("aria-label") or "No aria-label"
                button_text = button.text.strip() or "No text"
                button_class = button.get_attribute("class") or "No class"
                logger.info(f"Button {i}: aria-label='{aria_label}', text='{button_text}', class='{button_class[:50]}...'")
            
            # Try to find Connect button using multiple methods
            connect_button = None

            # Method 1: Direct search for Connect button
            logger.info("Trying direct search for Connect button...")
            connect_button = self._find_button_by_text(all_buttons, ['connect'], "Connect")
            
            # Method 2: Try clicking Connect button, fallback to More menu if click intercepted
            if connect_button:
                if self._try_click_button(connect_button, "Connect"):
                    # Connect button clicked successfully
                    pass
                else:
                    # Connect button click failed, try More button approach
                    logger.info("Connect button click failed, trying More button approach...")
                    connect_button = self._find_connect_button_in_more_menu()
                    if connect_button:
                        if not self._try_click_button(connect_button, "Connect (from More menu)"):
                            connect_button = None
            else:
                # No direct Connect button found, try More button approach
                logger.info("Direct Connect button not found, trying More button approach...")
                connect_button = self._find_connect_button_in_more_menu()
                if connect_button:
                    if not self._try_click_button(connect_button, "Connect (from More menu)"):
                        connect_button = None
            
            # If still no Connect button found or clicked, check for Message button
            if not connect_button:
                logger.info("Connect button still not found or clickable, checking for Message button...")
                all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
                message_button = self._find_button_by_text(all_buttons, ['message'], "Message")
                
                if message_button:
                    logger.info(f"Person {person.name} appears to already be connected (Message button found), skipping")
                    return False
                
                # Neither Connect nor Message button found - this is an error
                with open('debug_page.html', 'w', encoding='utf-8') as f:
                    f.write(self.driver.page_source)
                logger.error("Could not find or click Connect button, and no Message button found. Page source saved to debug_page.html")
                raise Exception("Could not find or click Connect button")
            
            try:
                # Try to find "Add a note" button using multiple methods
                add_note_button = None
                
                # Method 2: Manual search through all buttons
                if not add_note_button:
                    all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    for button in all_buttons:
                        aria_label = (button.get_attribute("aria-label") or "").lower()
                        button_text = (button.text or "").lower()
                        inner_html = (button.get_attribute("innerHTML") or "").lower()
                        
                        if (('add a note' in aria_label or 'add a note' in button_text or 'add a note' in inner_html) and 
                            button.is_enabled() and button.is_displayed()):
                            add_note_button = button
                            logger.info(f"Found Add a note button manually: aria-label='{button.get_attribute('aria-label')}', text='{button.text}'")
                            break
                
                if add_note_button:
                    add_note_button.click()
                else:
                    logger.warning("Could not find 'Add a note' button, proceeding without note")
                
                message = self.config.connection_message_template.format(
                    name=person.name,
                    company=person.company
                )
                
                note_field = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.TAG_NAME, "textarea"))
                )
                time.sleep(2)
                # Scroll into view and click to focus
                self.driver.execute_script("arguments[0].scrollIntoView();", note_field)
                note_field.click()
                time.sleep(1)
                note_field.send_keys(message)
                
                # Find Send button using the robust method
                send_button = None
                all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
                for button in all_buttons:
                    aria_label = (button.get_attribute("aria-label") or "").lower()
                    button_text = (button.text or "").lower()
                    inner_html = (button.get_attribute("innerHTML") or "").lower()
                    
                    if (('send' in aria_label or 'send' in button_text or 'send' in inner_html) and 
                        button.is_enabled() and button.is_displayed()):
                        send_button = button
                        logger.info(f"Found Send button manually: aria-label='{button.get_attribute('aria-label')}', text='{button.text}'")
                        break
                
                if send_button:
                    send_button.click()
                else:
                    logger.error("Could not find Send button")
                    raise Exception("Could not find Send button")
                
            except Exception as e:
                logger.error("Could not find Send now button")
                raise Exception("Could not find Send now button")
        
            logger.info(f"Connection request sent to {person.name}")
            
            # Add person to tracker if company_id is provided
            if company_id:
                try:
                    self.tracker.add_person_to_company(company_id, person.name, person.profile_url)
                    logger.info(f"Added {person.name} to company {company_id} in tracker")
                except Exception as e:
                    logger.error(f"Failed to add {person.name} to tracker: {str(e)}")
            
            delay = random.uniform(*self.config.delay_between_requests)
            time.sleep(delay)
            
            return True
            
        except Exception as e:
            error_message = f"Failed to send connection request to {person.name}: {str(e)}"
            logger.error(error_message)
            
            # Update company status with error message if company_id is provided
            if company_id:
                try:
                    self.tracker.update_company_status(company_id, error_message)
                    logger.info(f"Updated status for company {company_id} with error message")
                except Exception as status_error:
                    logger.error(f"Failed to update status for company {company_id}: {str(status_error)}")
            
            return False
    
    def process_companies(self, companies: List[Dict[str, any]]) -> Dict[str, int]:
        results = {}
        
        for company in companies:
            logger.info(f"Processing company: {company}")

            if company['status'] != '':
                logger.info(f"Skipping company: {company} because status is not empty")
                continue

            people = self.search_company_people(company['company_name'])
            sent_count = 0
            
            for person in people:
                if self.send_connection_request(person, company['company_id']):
                    sent_count += 1

                if sent_count >= self.config.max_connections_per_company:
                    break
            
            # Mark status as successful only if status is currently empty and we sent at least one request
            if sent_count > 0:
                try:
                    # Check current status first
                    company_data = self.tracker.get_company_row(company['company_id'])
                    if company_data and company_data['status'] == '':
                        self.tracker.update_company_status(
                            company['company_id'], 
                            "Completed", 
                            f"Successfully sent {sent_count} connection requests"
                        )
                        logger.info(f"Marked company {company['company_id']} as completed")
                except Exception as e:
                    logger.error(f"Failed to update completion status for company {company['company_id']}: {str(e)}")
            
            time.sleep(random.uniform(10, 20))
        
        return results
    
    def close(self):
        if self.driver:
            self.driver.quit()

def main():
    # Initialize tracker to get companies from Google Sheets
    tracker = CompanyTracker(
        credentials_file="gcloud.json",
        spreadsheet_name="Shruti Company Reachout",
        worksheet_name="Sheet1"
    )
    
    # Get companies from Google Sheets
    companies = tracker.get_all_companies()
    
    with open('config.json', 'r') as f:
        config_data = json.load(f)
    
    config = Config(**config_data)
    agent = LinkedInAgent(config)
    
    try:
        with open('credentials.json', 'r') as f:
            credentials = json.load(f)
        
        username = credentials['linkedin_username']
        password = credentials['linkedin_password']
        
        if agent.login_to_linkedin(username, password):
            results = agent.process_companies(companies)
            
            print("\n=== Results ===")
            for company, count in results.items():
                print(f"{company}: {count} connection requests sent")
        
    finally:
        agent.close()

if __name__ == "__main__":
    main()