import pandas as pd
import os
import re
from PyPDF2 import PdfReader
import pathlib
from fuzzywuzzy import fuzz  # For fuzzy matching

def extract_text_from_pdf(pdf_path):
    """Extract all text from a PDF file."""
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def normalize_text(text):
    """Normalize text by converting 'Ñ' to 'N' and handling other special characters."""
    # Replace Ñ with N as per requirements
    text = text.replace('Ñ', 'N')
    # Convert to lowercase for case-insensitive matching
    text = text.lower()
    return text

def exact_match(name, text):
    """Perform exact matching between the name and text."""
    normalized_name = normalize_text(name.strip())
    normalized_text = normalize_text(text)
    return normalized_name in normalized_text

def fuzzy_match(name, text, threshold=70):
    """
    Perform fuzzy matching between the name and text.
    Returns True if the match score is above the threshold.
    """
    normalized_name = normalize_text(name.strip())
    normalized_text = normalize_text(text)
    match_score = fuzz.partial_ratio(normalized_name, normalized_text)
    return match_score >= threshold, match_score

def find_leftmost_number(text, name, search_range=100):
    """
    Find the leftmost number in the row where the name is found.
    """
    normalized_name = normalize_text(name.strip())
    normalized_text = normalize_text(text)
    
    # Find the position of the name in the text
    name_pos = normalized_text.find(normalized_name)
    if name_pos == -1:
        return None
    
    # Find the start of the line containing the name
    line_start = normalized_text.rfind('\n', 0, name_pos) + 1
    # Find the end of the line containing the name
    line_end = normalized_text.find('\n', name_pos)
    if line_end == -1:
        line_end = len(normalized_text)
    
    # Extract the line containing the name
    line = normalized_text[line_start:line_end]
    
    # Look for the leftmost number in the line
    number_match = re.search(r'^\s*(\d+)\s*', line)
    if number_match:
        return int(number_match.group(1))
    
    return None

def find_matches_in_pdfs(names, directory="."):
    """
    Search for each name in all PDF files in the directory using exact and fuzzy matching.
    Returns a dictionary mapping names to their corresponding numbers.
    """
    name_to_number = {}
    pdf_files = list(pathlib.Path(directory).glob("*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in {directory}")
        return name_to_number
    
    print(f"Searching through {len(pdf_files)} PDF files...")
    
    # Extract and cache text from all PDFs
    pdf_texts = {}
    for pdf_file in pdf_files:
        pdf_path = str(pdf_file)
        try:
            pdf_texts[pdf_path] = extract_text_from_pdf(pdf_path)
            print(f"Successfully extracted text from {pdf_path}")
        except Exception as e:
            print(f"Error extracting text from {pdf_path}: {e}")
    
    # For each name, perform exact and fuzzy search in the PDFs
    for name in names:
        if not name or pd.isna(name):
            continue
        
        print(f"Searching for '{name}'...")
        
        for pdf_path, text in pdf_texts.items():
            # Perform exact matching first
            if exact_match(name, text):
                print(f"Found exact match for '{name}' in {pdf_path}")
                number = find_leftmost_number(text, name)
                if number is not None:
                    name_to_number[name] = str(number)
                    print(f"  Associated number: {number}")
                    break
            else:
                # Fall back to fuzzy matching if exact match fails
                found, match_score = fuzzy_match(name, text)
                if found:
                    print(f"Found fuzzy match for '{name}' in {pdf_path} (match score: {match_score})")
                    number = find_leftmost_number(text, name)
                    if number is not None:
                        name_to_number[name] = str(number)
                        print(f"  Associated number: {number}")
                        break
        
        if name not in name_to_number:
            print(f"No match or number found for '{name}'")
    
    return name_to_number

def process_excel_file(excel_path):
    """Process the Excel file, find matches, and update SQN column."""
    try:
        # Load the Excel file
        df = pd.read_excel(excel_path)
        
        # Check if 'Name' column exists
        if 'Name' not in df.columns:
            print("Column 'Name' not found in the Excel file")
            return False
            
        # Check if 'SQN' column exists
        if 'SQN' not in df.columns:
            print("Column 'SQN' not found in the Excel file")
            return False
            
        # Get the list of names from 'Name' column
        names = df['Name'].tolist()
            
        # Find matches in PDFs
        name_to_number = find_matches_in_pdfs(names)
        
        # Update 'SQN' column starting from row 2
        for i, name in enumerate(names, start=0):
            if i < 1:  # Skip the first row (header)
                continue
                
            if name in name_to_number:
                df.loc[i, 'SQN'] = name_to_number[name]
            # Else leave it blank as per requirements
        
        # Save the updated Excel file
        output_path = excel_path.replace('.xlsx', '_updated.xlsx')
        if output_path == excel_path:
            output_path = excel_path.replace('.xls', '_updated.xls')
        if output_path == excel_path:
            output_path = "updated_" + excel_path
            
        df.to_excel(output_path, index=False)
        print(f"Updated Excel file saved as {output_path}")
        return True
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        return False

if __name__ == "__main__":
    # Get the Excel file path from user
    excel_path = input("\nEnter the path to the Excel file (or press Enter to use the first Excel file in the current directory): ")
    
    if not excel_path:
        # Find the first Excel file in the current directory
        excel_files = list(pathlib.Path('.').glob('*.xlsx')) + list(pathlib.Path('.').glob('*.xls'))
        if not excel_files:
            print("No Excel files found in the current directory.")
            exit(1)
        excel_path = str(excel_files[0])
        print(f"Using Excel file: {excel_path}")
    
    # Process the Excel file
    success = process_excel_file(excel_path)
    
    if success:
        print("Processing completed successfully.")
    else:
        print("Processing failed.")