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
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import logging

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
            
            # Alternative selectors to bypass encrypted classes
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
                                break
                    
                    # Method 2: Look for span with dir="ltr" (common pattern)
                    if not name:
                        name_spans = result.find_elements(By.CSS_SELECTOR, 'span[dir="ltr"]')
                        for span in name_spans:
                            text = span.text.strip()
                            if text and len(text) > 3 and not any(word in text.lower() for word in ['view', 'connect', 'message']):
                                name = text
                                break
                    
                    # Method 3: Find title/subtitle using flexible approach
                    title = ""
                    
                    # Look for any element that might contain job title
                    potential_title_elements = []
                    
                    # Try common patterns for job titles
                    potential_title_elements.extend(result.find_elements(By.TAG_NAME, "p"))
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
    
    def send_connection_request(self, person: Person) -> bool:
        try:
            self.driver.get(person.profile_url)
            time.sleep(2)
            
            # Debug: Let's see what buttons are actually on the page
            logger.info("=== DEBUG: Looking for Connect button ===")
            all_buttons = self.driver.find_elements(By.TAG_NAME, "button")
            logger.info(f"Found {len(all_buttons)} buttons on the page")
            
            for i, button in enumerate(all_buttons[:10]):  # Log first 10 buttons
                aria_label = button.get_attribute("aria-label") or "No aria-label"
                button_text = button.text.strip() or "No text"
                button_class = button.get_attribute("class") or "No class"
                logger.info(f"Button {i}: aria-label='{aria_label}', text='{button_text}', class='{button_class[:50]}...'")
            
            # Try multiple methods to find the Connect button
            connect_button = None
            
            # # Method 1: Look for button with aria-label containing "Invite" and "connect"
            # try:
            #     connect_button = WebDriverWait(self.driver, 3).until(
            #         EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'Invite') and contains(@aria-label, 'connect')]"))
            #     )
            #     logger.info("Found Connect button using Method 1")
            # except TimeoutException:
            #     pass
            #
            # # Method 2: Look for button with aria-label just containing "connect"
            # if not connect_button:
            #     try:
            #         connect_button = WebDriverWait(self.driver, 3).until(
            #             EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'connect')]"))
            #         )
            #         logger.info("Found Connect button using Method 2")
            #     except TimeoutException:
            #         pass
            #
            # # Method 3: Look for button with span text "Connect"
            # if not connect_button:
            #     try:
            #         connect_button = WebDriverWait(self.driver, 3).until(
            #             EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Connect']]"))
            #         )
            #         logger.info("Found Connect button using Method 3")
            #     except TimeoutException:
            #         pass
            
            # Method 4: Manual search through all buttons
            if not connect_button:
                logger.info("Trying manual search through all buttons...")
                for button in all_buttons:
                    aria_label = (button.get_attribute("aria-label") or "").lower()
                    button_text = (button.text or "").lower()
                    # Also check inner HTML for "Connect" text
                    inner_html = (button.get_attribute("innerHTML") or "").lower()
                    
                    if (('connect' in aria_label or 'connect' in button_text or 'connect' in inner_html) and 
                        button.is_enabled() and button.is_displayed()):
                        connect_button = button
                        logger.info(f"Found Connect button manually: aria-label='{button.get_attribute('aria-label')}', text='{button.text}', innerHTML contains connect")
                        break
            
            if not connect_button:
                # Save page source for debugging
                with open('debug_page.html', 'w', encoding='utf-8') as f:
                    f.write(self.driver.page_source)
                logger.error("Could not find Connect button. Page source saved to debug_page.html")
                raise Exception("Could not find Connect button")
            
            connect_button.click()
            
            try:
                # Try to find "Add a note" button using multiple methods
                add_note_button = None
                
                # Method 1: XPath with text
                # try:
                #     add_note_button = WebDriverWait(self.driver, 3).until(
                #         EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add a note')]"))
                #     )
                # except TimeoutException:
                #     pass
                
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
                    EC.presence_of_element_located((By.TAG_NAME, "textarea"))
                )
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
                
            except TimeoutException:
                logger.error("Could not find Send now button")
                raise Exception("Could not find Send now button")
        
            logger.info(f"Connection request sent to {person.name}")
            
            delay = random.uniform(*self.config.delay_between_requests)
            time.sleep(delay)
            
            return True
            
        except Exception as e:
            logger.error(f"Failed to send connection request to {person.name}: {str(e)}")
            return False
    
    def process_companies(self, companies: List[str]) -> Dict[str, int]:
        results = {}
        
        for company in companies:
            logger.info(f"Processing company: {company}")
            
            people = self.search_company_people(company)
            sent_count = 0
            
            for person in people:
                if self.send_connection_request(person):
                    sent_count += 1
                
                if sent_count >= self.config.max_connections_per_company:
                    break
            
            results[company] = sent_count
            
            time.sleep(random.uniform(10, 20))
        
        return results
    
    def close(self):
        if self.driver:
            self.driver.quit()

def main():
    with open('companies.json', 'r') as f:
        companies = json.load(f)
    
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