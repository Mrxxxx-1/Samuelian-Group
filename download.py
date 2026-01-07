"""
Download Excel files from CTCAC applications page
"""

import os
import requests
from bs4 import BeautifulSoup
from typing import List


def download_excel_files(url: str, output_dir: str = "applications", limit: int = None) -> List[str]:
    """
    Download Excel files from the CTCAC applications page
    
    Args:
        url: URL to the applications page
        output_dir: Directory to save downloaded files
        limit: Maximum number of files to download (None for all)
    
    Returns:
        List of file paths to downloaded Excel files
    """
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find all links to Excel files
        excel_files = []
        base_url = "https://www.treasurer.ca.gov"
        seen_files = set()  # Track downloaded files to avoid duplicates
        
        for link in soup.find_all('a', href=True):
            if limit and len(excel_files) >= limit:
                break
                
            href = link['href']
            link_text = link.get_text(strip=True).lower()
            
            # Check if it's an Excel file link
            is_excel = (href.endswith(('.xlsx', '.xls')) or 
                       '.xlsx' in href.lower() or 
                       '.xls' in href.lower() or
                       'excel' in link_text or
                       'download' in link_text)
            
            if is_excel:
                # Handle relative URLs
                if href.startswith('/'):
                    file_url = f"{base_url}{href}"
                elif href.startswith('http'):
                    file_url = href
                else:
                    # Relative to current page
                    file_url = f"{url.rsplit('/', 1)[0]}/{href}"
                
                # Extract filename
                filename = os.path.basename(href.split('?')[0])  # Remove query params
                if not filename.endswith(('.xlsx', '.xls')):
                    # Try to get filename from URL or use a generated name
                    if '/' in href:
                        filename = href.split('/')[-1].split('?')[0]
                    if not filename.endswith(('.xlsx', '.xls')):
                        filename += '.xlsx'  # Assume xlsx if extension missing
                
                # Skip if already downloaded
                if filename in seen_files:
                    continue
                seen_files.add(filename)
                
                file_path = os.path.join(output_dir, filename)
                
                # Skip if file already exists
                if os.path.exists(file_path):
                    print(f"Skipping (already exists): {filename}")
                    excel_files.append(file_path)
                    continue
                
                # Download file
                try:
                    file_response = requests.get(file_url, timeout=60)
                    file_response.raise_for_status()
                    
                    # Verify it's actually an Excel file by checking content type or magic bytes
                    content_type = file_response.headers.get('content-type', '').lower()
                    if 'excel' in content_type or 'spreadsheet' in content_type or file_response.content[:4] == b'PK\x03\x04':
                        with open(file_path, 'wb') as f:
                            f.write(file_response.content)
                        excel_files.append(file_path)
                        print(f"Downloaded: {filename}")
                    else:
                        print(f"Skipping (not Excel): {filename}")
                except Exception as e:
                    print(f"Error downloading {file_url}: {str(e)}")
        
        return excel_files
    except Exception as e:
        print(f"Error accessing URL: {str(e)}")
        return []


def find_excel_files_in_directory(directory: str = "applications") -> List[str]:
    """Find all Excel files in a directory"""
    excel_files = []
    if os.path.exists(directory):
        for file in os.listdir(directory):
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(directory, file))
    return excel_files

