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
        self.driver.implicitly_wait(10)
    
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
            time.sleep(10)
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
            
            time.sleep(10)
            
            search_results = self.driver.find_elements(
                By.CSS_SELECTOR, 
                ".search-result__wrapper .search-result__info"
            )
            
            for result in search_results[:10]:
                try:
                    name_element = result.find_element(By.CSS_SELECTOR, ".search-result__result-link")
                    name = name_element.text.strip()
                    profile_url = name_element.get_attribute("href")
                    
                    title_element = result.find_element(By.CSS_SELECTOR, ".subline-level-1")
                    title = title_element.text.strip()
                    
                    if any(keyword in title.lower() for keyword in ['product', 'pm', 'manager', 'lead']):
                        person = Person(
                            name=name,
                            title=title,
                            company=company_name,
                            profile_url=profile_url
                        )
                        people.append(person)
                        
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
            
            connect_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'Connect')]"))
            )
            connect_button.click()
            
            try:
                add_note_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Add a note')]"))
                )
                add_note_button.click()
                
                message = self.config.connection_message_template.format(
                    name=person.name,
                    company=person.company
                )
                
                note_field = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.TAG_NAME, "textarea"))
                )
                note_field.send_keys(message)
                
                send_button = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Send')]")
                send_button.click()
                
            except TimeoutException:
                send_button = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Send now')]")
                send_button.click()
            
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