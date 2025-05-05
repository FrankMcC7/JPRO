#!/usr/bin/env python3
"""
Bloomberg LEI Search Script
--------------------------
This script searches lei.bloomberg.com for Legal Entity Identifiers (LEIs) of funds.
It reads fund names from a hard-coded Excel file path and updates the file with LEI information.

Requirements:
- Python 3.6+
- requests
- beautifulsoup4
- pandas
- openpyxl (for Excel file handling)

Install dependencies with:
pip install requests beautifulsoup4 pandas openpyxl
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import argparse
import time
import sys
import re
import os
from datetime import datetime

# ================================================================
# CUSTOMIZE THESE SETTINGS FOR YOUR EXCEL FILE
# ================================================================
# Path to your Excel file containing fund names
EXCEL_FILE_PATH = "fund_names.xlsx"  # Replace with your file path

# Name of the sheet in the Excel file (0 for first sheet, or sheet name as string)
SHEET_NAME = 0  

# Column name that contains fund names
FUND_NAME_COLUMN = "Fund Name"

# Column name where LEI information will be stored
LEI_COLUMN = "LEI"
# ================================================================

class BloombergLEISearcher:
    def __init__(self, max_results=5, delay=1, verbose=True):
        self.base_url = "https://lei.bloomberg.com"
        self.search_url = f"{self.base_url}/search"
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Referer': self.base_url,
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }
        self.session = requests.Session()
        self.max_results = max_results
        self.delay = delay
        self.verbose = verbose
        
    def search_lei(self, fund_name):
        """Search for fund's LEI by name"""
        if self.verbose:
            print(f"Searching for: {fund_name}")
        
        params = {
            'q': fund_name,
            'page': 1
        }
        
        try:
            response = self.session.get(self.search_url, headers=self.headers, params=params)
            response.raise_for_status()
            
            results = self._parse_search_results(response.text)
            
            if not results:
                if self.verbose:
                    print(f"No results found for '{fund_name}'")
                return None
                
            if self.verbose:
                print(f"Found {len(results)} potential matches.")
            
            # For each result, get the details page
            detailed_results = []
            for idx, result in enumerate(results[:self.max_results]):  # Limit to max_results
                if self.verbose:
                    print(f"Processing result {idx+1}/{min(len(results), self.max_results)}: {result['name']}")
                details = self._get_entity_details(result['url'])
                if details:
                    detailed_results.append({**result, **details})
                time.sleep(self.delay)  # Be nice to the server
                
            return self._format_results(detailed_results)
            
        except requests.exceptions.RequestException as e:
            if self.verbose:
                print(f"Error during search: {e}")
            return None
    
    def batch_search(self, fund_names):
        """Search for multiple fund names and combine results"""
        all_results = []
        
        for idx, fund_name in enumerate(fund_names):
            if self.verbose:
                print(f"\nProcessing {idx+1}/{len(fund_names)}: {fund_name}")
            
            # Skip empty fund names
            if pd.isna(fund_name) or fund_name.strip() == '':
                if self.verbose:
                    print("Skipping empty fund name")
                continue
                
            results = self.search_lei(fund_name)
            
            if results is not None and not results.empty:
                # Add the search query to the results for reference
                results['search_query'] = fund_name
                all_results.append(results)
                
            # Be extra nice to the server between different fund searches
            if idx < len(fund_names) - 1:  # Don't sleep after the last one
                time.sleep(self.delay * 2)
                
        if all_results:
            # Combine all results into a single DataFrame
            combined_results = pd.concat(all_results, ignore_index=True)
            return combined_results
        else:
            return pd.DataFrame()
    
    def _parse_search_results(self, html_content):
        """Parse the search results page"""
        soup = BeautifulSoup(html_content, 'html.parser')
        results = []
        
        result_items = soup.select('.search-result-item')
        
        for item in result_items:
            name_elem = item.select_one('.lei-link')
            if not name_elem:
                continue
                
            name = name_elem.text.strip()
            url = name_elem.get('href')
            
            # Basic information in the search results
            lei_number = ''
            status = ''
            
            lei_elem = item.select_one('.lei-code')
            if lei_elem:
                lei_number = lei_elem.text.strip()
                
            status_elem = item.select_one('.registration-status')
            if status_elem:
                status = status_elem.text.strip()
                
            results.append({
                'name': name,
                'lei': lei_number,
                'status': status,
                'url': self.base_url + url if url.startswith('/') else url
            })
            
        return results
    
    def _get_entity_details(self, url):
        """Get detailed information from the entity's page"""
        try:
            response = self.session.get(url, headers=self.headers)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            details = {}
            
            # Extract more detailed information
            detail_items = soup.select('.profile-data-item')
            for item in detail_items:
                label_elem = item.select_one('.profile-data-label')
                value_elem = item.select_one('.profile-data-value')
                
                if label_elem and value_elem:
                    label = label_elem.text.strip().rstrip(':')
                    value = value_elem.text.strip()
                    
                    # Clean up the key for use as a dictionary key
                    key = re.sub(r'[^a-zA-Z0-9]', '_', label).lower()
                    details[key] = value
            
            # Handle Bloomberg LEI vs Additional LEI sections
            # Check if there's an "Additional LEI" section
            additional_lei_section = soup.select_one('.additional-lei-section')
            if additional_lei_section:
                # Extract the LEIs from the additional section
                additional_leis = []
                lei_items = additional_lei_section.select('.lei-item')
                
                for lei_item in lei_items:
                    lei_code = lei_item.select_one('.lei-code')
                    lei_source = lei_item.select_one('.lei-source')
                    
                    if lei_code and lei_source:
                        additional_leis.append({
                            'code': lei_code.text.strip(),
                            'source': lei_source.text.strip()
                        })
                
                details['additional_leis'] = additional_leis
            
            # The primary LEI section (Bloomberg LEI)
            bloomberg_lei_section = soup.select_one('.bloomberg-lei-section, .primary-lei-section')
            if bloomberg_lei_section:
                lei_code = bloomberg_lei_section.select_one('.lei-code')
                if lei_code:
                    details['bloomberg_lei'] = lei_code.text.strip()
            
            return details
            
        except requests.exceptions.RequestException as e:
            print(f"Error getting details: {e}")
            return {}
    
    def _format_results(self, results):
        """Format the results for output"""
        if not results:
            return pd.DataFrame()
            
        # Process additional LEIs if present
        for result in results:
            if 'additional_leis' in result and isinstance(result['additional_leis'], list):
                # Create a formatted string of additional LEIs for display
                additional_leis_str = "; ".join([
                    f"{lei_info['code']} ({lei_info['source']})" 
                    for lei_info in result['additional_leis']
                ])
                result['additional_leis_formatted'] = additional_leis_str
                
                # Add the first additional LEI as a separate column for convenience
                if result['additional_leis']:
                    result['first_additional_lei'] = result['additional_leis'][0]['code']
                    result['first_additional_lei_source'] = result['additional_leis'][0]['source']
        
        df = pd.DataFrame(results)
        
        # Reorder columns to put important information first
        priority_cols = ['name', 'lei', 'bloomberg_lei', 'first_additional_lei', 
                         'first_additional_lei_source', 'status', 'url', 
                         'additional_leis_formatted']
        
        # Only include columns that actually exist
        priority_cols = [col for col in priority_cols if col in df.columns]
        other_cols = [col for col in df.columns if col not in priority_cols and 
                      col != 'additional_leis']  # Exclude the raw additional_leis list
        
        df = df[priority_cols + other_cols]
        
        return df
    
def main():
    parser = argparse.ArgumentParser(description='Search for LEIs on Bloomberg')
    
    # Processing options
    parser.add_argument('--detailed', '-d', action='store_true', 
                      help='Show all detailed information in console output')
    parser.add_argument('--max-results', '-m', type=int, default=5,
                      help='Maximum number of results to process per fund (default: 5)')
    parser.add_argument('--delay', type=float, default=1.0,
                      help='Delay between requests in seconds (default: 1.0)')
    parser.add_argument('--quiet', '-q', action='store_true',
                      help='Suppress verbose output')
    parser.add_argument('--output', '-o',
                      help='Save detailed results to a separate Excel file')
    
    args = parser.parse_args()
    
    # Create the searcher with configured options
    searcher = BloombergLEISearcher(
        max_results=args.max_results,
        delay=args.delay,
        verbose=not args.quiet
    )
    
    try:
        # Check if the Excel file exists
        if not os.path.exists(EXCEL_FILE_PATH):
            print(f"Error: Excel file '{EXCEL_FILE_PATH}' not found.")
            print("Please update the EXCEL_FILE_PATH variable in the script.")
            sys.exit(1)
        
        # Read the Excel file
        if not args.quiet:
            print(f"Reading fund names from {EXCEL_FILE_PATH}")
        
        input_df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
        
        if FUND_NAME_COLUMN not in input_df.columns:
            print(f"Error: Column '{FUND_NAME_COLUMN}' not found in the Excel file.")
            print(f"Available columns: {', '.join(input_df.columns)}")
            print("Please update the FUND_NAME_COLUMN variable in the script.")
            sys.exit(1)
        
        # Add LEI column if it doesn't exist
        if LEI_COLUMN not in input_df.columns:
            input_df[LEI_COLUMN] = None
        
        # Extract fund names
        fund_names = input_df[FUND_NAME_COLUMN].dropna().astype(str).tolist()
        
        if not args.quiet:
            print(f"Found {len(fund_names)} fund names to process")
        
        # Perform batch search
        results = searcher.batch_search(fund_names)
        
        if results is not None and not results.empty:
            # Display results
            if not args.quiet:
                print("\nSearch Results:")
                
                # Create a more readable display version
                display_cols = ['search_query', 'name', 'lei', 'bloomberg_lei', 
                               'first_additional_lei', 'status']
                
                # Only include columns that actually exist
                display_cols = [col for col in display_cols if col in results.columns]
                
                if args.detailed:
                    # Show all columns if detailed flag is provided
                    print(results.to_string())
                else:
                    # Otherwise show a simplified view
                    display_df = results[display_cols].copy() if display_cols else results
                    print(display_df.to_string())
                    print("\nNote: Use --detailed or -d flag to see all information")
            
            # Create a mapping from search query to first LEI found
            lei_mapping = {}
            for _, row in results.iterrows():
                search_query = row['search_query']
                if search_query not in lei_mapping:
                    # Use bloomberg_lei if available, otherwise use lei
                    if 'bloomberg_lei' in row and pd.notna(row['bloomberg_lei']):
                        lei_mapping[search_query] = row['bloomberg_lei']
                    elif 'lei' in row and pd.notna(row['lei']):
                        lei_mapping[search_query] = row['lei']
            
            # Update the Excel file with LEI information
            updated_count = 0
            for i, fund_name in enumerate(input_df[FUND_NAME_COLUMN]):
                if pd.notna(fund_name) and str(fund_name) in lei_mapping:
                    input_df.at[i, LEI_COLUMN] = lei_mapping[str(fund_name)]
                    updated_count += 1
            
            # Save updated Excel file
            input_df.to_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME, index=False)
            
            if not args.quiet:
                print(f"\nUpdated {updated_count} LEIs in {EXCEL_FILE_PATH}")
            
            # Also save detailed results to a separate output file if specified
            if args.output:
                results.to_excel(args.output, index=False)
                if not args.quiet:
                    print(f"Detailed results saved to {args.output}")
        else:
            print("No results found or an error occurred.")
    
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()