#!/usr/bin/env python3

import gspread
from google.oauth2.service_account import Credentials
from typing import List, Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)

class GoogleSheetsManager:
    def __init__(self, credentials_file: str, spreadsheet_name: str):
        """
        Initialize Google Sheets manager
        
        Args:
            credentials_file: Path to Google service account JSON credentials file
            spreadsheet_name: Name of the Google Spreadsheet to work with
        """
        self.credentials_file = credentials_file
        self.spreadsheet_name = spreadsheet_name
        self.gc = None
        self.spreadsheet = None
        self._authenticate()
    
    def _authenticate(self):
        """Authenticate with Google Sheets API"""
        try:
            # Define the scope
            scope = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # Load credentials
            creds = Credentials.from_service_account_file(
                self.credentials_file, 
                scopes=scope
            )
            
            # Authorize the client
            self.gc = gspread.authorize(creds)
            
            # Open the spreadsheet
            self.spreadsheet = self.gc.open(self.spreadsheet_name)
            
            logger.info(f"Successfully authenticated and opened spreadsheet: {self.spreadsheet_name}")
            
        except Exception as e:
            logger.error(f"Failed to authenticate with Google Sheets: {str(e)}")
            raise
    
    def get_worksheet(self, worksheet_name: str = None, index: int = 0):
        """
        Get a specific worksheet
        
        Args:
            worksheet_name: Name of the worksheet (optional)
            index: Index of the worksheet if name not provided (default: 0)
        
        Returns:
            Worksheet object
        """
        try:
            if worksheet_name:
                worksheet = self.spreadsheet.worksheet(worksheet_name)
            else:
                worksheet = self.spreadsheet.get_worksheet(index)
            
            logger.info(f"Successfully accessed worksheet: {worksheet.title}")
            return worksheet
            
        except Exception as e:
            logger.error(f"Failed to get worksheet: {str(e)}")
            raise
    
    def create_worksheet(self, title: str, rows: int = 1000, cols: int = 26) -> Any:
        """
        Create a new worksheet
        
        Args:
            title: Name of the new worksheet
            rows: Number of rows (default: 1000)
            cols: Number of columns (default: 26)
        
        Returns:
            New worksheet object
        """
        try:
            worksheet = self.spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
            logger.info(f"Created new worksheet: {title}")
            return worksheet
        except Exception as e:
            logger.error(f"Failed to create worksheet {title}: {str(e)}")
            raise
    
    def write_data(self, data: List[List[Any]], worksheet_name: str = None, 
                   start_cell: str = "A1", clear_first: bool = False):
        """
        Write data to a worksheet
        
        Args:
            data: 2D list of data to write
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
            start_cell: Starting cell (default: "A1")
            clear_first: Whether to clear the worksheet before writing (default: False)
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            
            if clear_first:
                worksheet.clear()
                logger.info(f"Cleared worksheet: {worksheet.title}")
            
            # Write data
            worksheet.update(start_cell, data)
            logger.info(f"Successfully wrote {len(data)} rows to {worksheet.title}")
            
        except Exception as e:
            logger.error(f"Failed to write data: {str(e)}")
            raise
    
    def read_data(self, worksheet_name: str = None, range_name: str = None) -> List[List[str]]:
        """
        Read data from a worksheet
        
        Args:
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
            range_name: Specific range to read (e.g., "A1:C10") (optional, reads all if not provided)
        
        Returns:
            2D list of cell values
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            
            if range_name:
                data = worksheet.get(range_name)
            else:
                data = worksheet.get_all_values()
            
            logger.info(f"Successfully read {len(data)} rows from {worksheet.title}")
            return data
            
        except Exception as e:
            logger.error(f"Failed to read data: {str(e)}")
            raise
    
    def append_row(self, row_data: List[Any], worksheet_name: str = None):
        """
        Append a single row to the end of the worksheet
        
        Args:
            row_data: List of values to append as a new row
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            worksheet.append_row(row_data)
            logger.info(f"Successfully appended row to {worksheet.title}")
            
        except Exception as e:
            logger.error(f"Failed to append row: {str(e)}")
            raise
    
    def append_rows(self, rows_data: List[List[Any]], worksheet_name: str = None):
        """
        Append multiple rows to the end of the worksheet
        
        Args:
            rows_data: List of lists, each inner list is a row to append
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            worksheet.append_rows(rows_data)
            logger.info(f"Successfully appended {len(rows_data)} rows to {worksheet.title}")
            
        except Exception as e:
            logger.error(f"Failed to append rows: {str(e)}")
            raise
    
    def update_cell(self, row: int, col: int, value: Any, worksheet_name: str = None):
        """
        Update a specific cell
        
        Args:
            row: Row number (1-indexed)
            col: Column number (1-indexed)
            value: Value to set
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            worksheet.update_cell(row, col, value)
            logger.info(f"Successfully updated cell ({row}, {col}) in {worksheet.title}")
            
        except Exception as e:
            logger.error(f"Failed to update cell: {str(e)}")
            raise
    
    def get_cell_value(self, row: int, col: int, worksheet_name: str = None) -> str:
        """
        Get value from a specific cell
        
        Args:
            row: Row number (1-indexed)
            col: Column number (1-indexed)
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
        
        Returns:
            Cell value as string
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            value = worksheet.cell(row, col).value
            logger.info(f"Successfully read cell ({row}, {col}) from {worksheet.title}")
            return value
            
        except Exception as e:
            logger.error(f"Failed to read cell: {str(e)}")
            raise
    
    def find_cell(self, search_value: str, worksheet_name: str = None) -> Optional[Any]:
        """
        Find a cell containing specific value
        
        Args:
            search_value: Value to search for
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
        
        Returns:
            Cell object if found, None otherwise
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            cell = worksheet.find(search_value)
            if cell:
                logger.info(f"Found '{search_value}' at row {cell.row}, col {cell.col}")
                return cell
            else:
                logger.info(f"'{search_value}' not found in {worksheet.title}")
                return None
                
        except Exception as e:
            logger.error(f"Failed to find cell: {str(e)}")
            raise
    
    def clear_worksheet(self, worksheet_name: str = None):
        """
        Clear all data from a worksheet
        
        Args:
            worksheet_name: Name of the worksheet (optional, uses first sheet if not provided)
        """
        try:
            worksheet = self.get_worksheet(worksheet_name)
            worksheet.clear()
            logger.info(f"Successfully cleared worksheet: {worksheet.title}")
            
        except Exception as e:
            logger.error(f"Failed to clear worksheet: {str(e)}")
            raise


class CompanyTracker(GoogleSheetsManager):
    """
    Specialized class for tracking companies and people reached out to
    Expected columns: Company ID, Company Name, Status, Comments, Person 1, Person 2, etc.
    """
    
    def __init__(self, credentials_file: str, spreadsheet_name: str, worksheet_name: str = None):
        super().__init__(credentials_file, spreadsheet_name)
        self.worksheet_name = worksheet_name
        self.base_columns = ["Company ID", "Company Name", "Status", "Comments"]
        self.base_column_count = len(self.base_columns)
    
    def get_company_row(self, company_id: str) -> Optional[Dict[str, Any]]:
        """
        Get company details by Company ID
        
        Args:
            company_id: The Company ID to search for
            
        Returns:
            Dictionary with company details or None if not found
        """
        try:
            worksheet = self.get_worksheet(self.worksheet_name)
            all_data = worksheet.get_all_values()
            
            if not all_data:
                return None
                
            headers = all_data[0]
            
            # Find the row with matching Company ID
            for row_idx, row in enumerate(all_data[1:], start=2):  # Start from row 2 (1-indexed)
                if row and row[0] == company_id:  # Company ID is in first column
                    company_data = {
                        "row_number": row_idx,
                        "company_id": row[0] if len(row) > 0 else "",
                        "company_name": row[1] if len(row) > 1 else "",
                        "status": row[2] if len(row) > 2 else "",
                        "comments": row[3] if len(row) > 3 else "",
                        "people": []
                    }
                    
                    # Extract person data from remaining columns
                    for col_idx in range(self.base_column_count, len(row)):
                        if col_idx < len(headers) and row[col_idx]:
                            person_header = headers[col_idx]
                            company_data["people"].append({
                                "column": person_header,
                                "data": row[col_idx]
                            })
                    
                    logger.info(f"Found company {company_id} at row {row_idx}")
                    return company_data
                    
            logger.info(f"Company {company_id} not found")
            return None
            
        except Exception as e:
            logger.error(f"Failed to get company row: {str(e)}")
            raise
    
    def get_all_companies(self) -> List[Dict[str, Any]]:
        """
        Get all companies with their details
        
        Returns:
            List of dictionaries with company details
        """
        try:
            worksheet = self.get_worksheet(self.worksheet_name)
            all_data = worksheet.get_all_values()
            
            if not all_data or len(all_data) < 2:
                return []
                
            headers = all_data[0]
            companies = []
            
            for row_idx, row in enumerate(all_data[1:], start=2):
                if row and row[0]:  # Skip empty rows
                    company_data = {
                        "row_number": row_idx,
                        "company_id": row[0] if len(row) > 0 else "",
                        "company_name": row[1] if len(row) > 1 else "",
                        "status": row[2] if len(row) > 2 else "",
                        "comments": row[3] if len(row) > 3 else "",
                        "people": []
                    }
                    
                    # Extract person data
                    for col_idx in range(self.base_column_count, len(row)):
                        if col_idx < len(headers) and row[col_idx]:
                            person_header = headers[col_idx]
                            company_data["people"].append({
                                "column": person_header,
                                "data": row[col_idx]
                            })
                    
                    companies.append(company_data)
            
            logger.info(f"Retrieved {len(companies)} companies")
            return companies
            
        except Exception as e:
            logger.error(f"Failed to get all companies: {str(e)}")
            raise
    
    def update_company_status(self, company_id: str, status: str, comments: str = None):
        """
        Update company status and optionally comments
        
        Args:
            company_id: Company ID to update
            status: New status
            comments: New comments (optional)
        """
        try:
            company_data = self.get_company_row(company_id)
            if not company_data:
                raise ValueError(f"Company {company_id} not found")
            
            row_number = company_data["row_number"]
            worksheet = self.get_worksheet(self.worksheet_name)
            
            # Update status (column 3)
            worksheet.update_cell(row_number, 3, status)
            
            # Update comments if provided (column 4)
            if comments is not None:
                worksheet.update_cell(row_number, 4, comments)
            
            logger.info(f"Updated company {company_id} status to {status}")
            
        except Exception as e:
            logger.error(f"Failed to update company status: {str(e)}")
            raise
    
    def _get_next_person_column(self) -> str:
        """
        Find the next available Person column number
        
        Returns:
            Next person column name (e.g., "Person 1", "Person 2", etc.)
        """
        try:
            worksheet = self.get_worksheet(self.worksheet_name)
            headers = worksheet.row_values(1)  # Get first row (headers)
            
            person_columns = []
            for header in headers:
                if header.startswith("Person "):
                    try:
                        person_num = int(header.split("Person ")[1])
                        person_columns.append(person_num)
                    except (IndexError, ValueError):
                        continue
            
            if person_columns:
                next_person_num = max(person_columns) + 1
            else:
                next_person_num = 1
                
            return f"Person {next_person_num}"
            
        except Exception as e:
            logger.error(f"Failed to get next person column: {str(e)}")
            raise
    
    def add_person_to_company(self, company_id: str, person_name: str, linkedin_url: str):
        """
        Add a new person to a company
        
        Args:
            company_id: Company ID to add person to
            person_name: Name of the person
            linkedin_url: LinkedIn profile URL
        """
        try:
            company_data = self.get_company_row(company_id)
            if not company_data:
                raise ValueError(f"Company {company_id} not found")
            
            worksheet = self.get_worksheet(self.worksheet_name)
            row_number = company_data["row_number"]
            
            # Get next person column
            next_person_column = self._get_next_person_column()
            
            # Find the column number for the new person column
            headers = worksheet.row_values(1)
            
            # Check if this person column already exists in headers
            if next_person_column in headers:
                col_number = headers.index(next_person_column) + 1
            else:
                # Add new column header
                col_number = len(headers) + 1
                worksheet.update_cell(1, col_number, next_person_column)
            
            # Format person data as clickable hyperlink
            person_data = f'=HYPERLINK("{linkedin_url}","{person_name}")'
            
            # Add person data to the company row
            worksheet.update_cell(row_number, col_number, person_data)
            
            logger.info(f"Added {person_name} to company {company_id} in column {next_person_column}")
            
        except Exception as e:
            logger.error(f"Failed to add person to company: {str(e)}")
            raise
    
    def initialize_spreadsheet(self):
        """
        Initialize the spreadsheet with base headers if it's empty
        """
        try:
            worksheet = self.get_worksheet(self.worksheet_name)
            
            # Check if spreadsheet is empty
            try:
                first_row = worksheet.row_values(1)
                if first_row:
                    logger.info("Spreadsheet already has headers")
                    return
            except:
                pass
            
            # Initialize with base headers
            worksheet.update('A1:D1', [self.base_columns])
            logger.info("Initialized spreadsheet with base headers")
            
        except Exception as e:
            logger.error(f"Failed to initialize spreadsheet: {str(e)}")
            raise


def example_company_tracker():
    """Example of how to use the CompanyTracker class"""
    
    try:
        # Initialize company tracker
        tracker = CompanyTracker(
            credentials_file="gcloud.json",
            spreadsheet_name="Shruti Company Reachout",
            worksheet_name="Sheet1"  # Optional, uses first sheet if not provided
        )
        
        # Initialize spreadsheet with headers if empty
        tracker.initialize_spreadsheet()
        
        # Example: Get details for a specific company
        company_data = tracker.get_company_row("1")
        if company_data:
            print(f"Company: {company_data['company_name']}")
            print(f"Status: {company_data['status']}")
            print(f"People contacted: {len(company_data['people'])}")
        
        # Example: Get all companies
        all_companies = tracker.get_all_companies()
        for company in all_companies:
            print(f"{company['company_id']}: {company['company_name']} - {company['status']}")
        
        # Example: Update company status
        tracker.update_company_status(
            company_id="1",
            status="In Progress",
            comments="Started outreach campaign"
        )
        
        # Example: Add a person to a company
        tracker.add_person_to_company(
            company_id="1",
            person_name="John Smith",
            linkedin_url="https://linkedin.com/in/johnsmith"
        )
        
        # Example: Add another person to the same company
        tracker.add_person_to_company(
            company_id="1",
            person_name="Jane Doe", 
            linkedin_url="https://linkedin.com/in/janedoe"
        )
        
        print("Company tracker operations completed successfully!")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    # Run the company tracker example
    example_company_tracker()