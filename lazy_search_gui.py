import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import os
import threading
import pandas as pd
import re
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
import pytesseract
from pathlib import Path  # Import Path from pathlib

class LazySearchGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Lazy Search GUI")
        self.root.geometry("600x400")

        # Variables to store file paths
        self.excel_file_path = tk.StringVar()
        self.pdf_directory_path = tk.StringVar()
        self.output_file_path = tk.StringVar()

        # Labels for instructions
        self.instruction_label = tk.Label(root, text="Drag and drop the Excel file and PDF directory below:", font=("Arial", 12))
        self.instruction_label.pack(pady=10)

        # Drag & Drop for Excel file
        self.excel_frame = tk.Frame(root)
        self.excel_frame.pack(pady=10)
        self.excel_drop_label = tk.Label(self.excel_frame, text="Drop Excel File Here", bg="lightgray", width=40, height=2)
        self.excel_drop_label.pack(side=tk.LEFT, padx=5)
        self.excel_drop_label.drop_target_register(DND_FILES)
        self.excel_drop_label.dnd_bind('<<Drop>>', self.on_excel_drop)
        self.clear_excel_button = tk.Button(self.excel_frame, text="Clear", command=self.clear_excel)
        self.clear_excel_button.pack(side=tk.LEFT)

        # Drag & Drop for PDF directory
        self.pdf_frame = tk.Frame(root)
        self.pdf_frame.pack(pady=10)
        self.pdf_drop_label = tk.Label(self.pdf_frame, text="Drop PDF Directory Here", bg="lightgray", width=40, height=2)
        self.pdf_drop_label.pack(side=tk.LEFT, padx=5)
        self.pdf_drop_label.drop_target_register(DND_FILES)
        self.pdf_drop_label.dnd_bind('<<Drop>>', self.on_pdf_drop)
        self.clear_pdf_button = tk.Button(self.pdf_frame, text="Clear", command=self.clear_pdf)
        self.clear_pdf_button.pack(side=tk.LEFT)

        # Status label
        self.status_label = tk.Label(root, text="Status: Waiting for files...", font=("Arial", 10))
        self.status_label.pack(pady=10)

        # Progress label
        self.progress_label = tk.Label(root, text="Progress: 0%", font=("Arial", 10))
        self.progress_label.pack(pady=10)

        # Button to open the updated Excel file (initially disabled)
        self.open_excel_button = tk.Button(root, text="Open Updated Excel File", state=tk.DISABLED, command=self.open_updated_excel)
        self.open_excel_button.pack(pady=10)

        # Button to run the lazy search algorithm
        self.run_button = tk.Button(root, text="Run Lazy Search", command=self.run_lazy_search)
        self.run_button.pack(pady=10)

    def on_excel_drop(self, event):
        """Handle Excel file drop event."""
        self.excel_file_path.set(event.data.strip('{}'))
        self.excel_drop_label.config(text=f"Excel File: {os.path.basename(self.excel_file_path.get())}")
        self.update_status()

    def on_pdf_drop(self, event):
        """Handle PDF directory drop event."""
        self.pdf_directory_path.set(event.data.strip('{}'))
        self.pdf_drop_label.config(text=f"PDF Directory: {os.path.basename(self.pdf_directory_path.get())}")
        self.update_status()

    def clear_excel(self):
        """Clear the selected Excel file."""
        self.excel_file_path.set("")
        self.excel_drop_label.config(text="Drop Excel File Here")
        self.update_status()

    def clear_pdf(self):
        """Clear the selected PDF directory."""
        self.pdf_directory_path.set("")
        self.pdf_drop_label.config(text="Drop PDF Directory Here")
        self.update_status()

    def update_status(self):
        """Update the status label based on file inputs."""
        if self.excel_file_path.get() and self.pdf_directory_path.get():
            self.status_label.config(text="Status: Ready to run Lazy Search.")
        else:
            self.status_label.config(text="Status: Waiting for files...")

    def run_lazy_search(self):
        """Run the lazy search algorithm in a separate thread."""
        if not self.excel_file_path.get() or not self.pdf_directory_path.get():
            messagebox.showerror("Error", "Please provide both Excel and PDF files.")
            return

        # Disable the run button while processing
        self.run_button.config(state=tk.DISABLED)
        self.status_label.config(text="Status: Running Lazy Search...")
        self.progress_label.config(text="Progress: 0%")

        # Run the lazy search algorithm in a separate thread
        def run_script():
            try:
                # Call the process_excel_file function directly
                success = process_excel_file(self.excel_file_path.get(), self.pdf_directory_path.get())

                if success:
                    self.status_label.config(text="Status: Lazy Search completed successfully!")
                    self.progress_label.config(text="Progress: 100%")
                    self.output_file_path.set(self.excel_file_path.get().replace('.xlsx', '_updated.xlsx'))
                    self.open_excel_button.config(state=tk.NORMAL)
                else:
                    self.status_label.config(text="Status: Lazy Search failed.")
                    self.progress_label.config(text="Progress: 0%")
                    messagebox.showerror("Error", "Lazy Search failed.")

            except Exception as e:
                self.status_label.config(text="Status: Lazy Search failed.")
                self.progress_label.config(text="Progress: 0%")
                messagebox.showerror("Error", str(e))
            finally:
                self.run_button.config(state=tk.NORMAL)

        # Start the thread
        threading.Thread(target=run_script).start()

    def open_updated_excel(self):
        """Open the updated Excel file."""
        if self.output_file_path.get():
            try:
                os.startfile(self.output_file_path.get())
            except Exception as e:
                messagebox.showerror("Error", f"Could not open the updated Excel file: {e}")
        else:
            messagebox.showerror("Error", "No updated Excel file found.")

# Functions from lazy_search.py
def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file, including scanned PDFs using OCR."""
    try:
        # Try extracting text using PyPDF2
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        
        # If no text is extracted, assume it's a scanned PDF and use OCR
        if not text.strip():
            print(f"No text extracted from {pdf_path}. Attempting OCR...")
            images = convert_from_path(pdf_path)
            text = ""
            for image in images:
                text += pytesseract.image_to_string(image)
        
        if not text.strip():
            raise Exception(f"No text could be extracted from {pdf_path}.")
        
        return text
    except Exception as e:
        raise Exception(f"Error reading PDF file '{pdf_path}': {e}")

def normalize_text(text):
    """Normalize text by converting 'ñ' to 'n' and handling other special characters."""
    # Replace ñ with n as per requirements
    text = text.replace('ñ', 'n')
    # Convert to lowercase for case-insensitive matching
    text = text.lower()
    return text

def character_by_character_search(name, text):
    """
    Search for a name in text character by character.
    Returns True if the exact name is found, False otherwise.
    """
    normalized_name = normalize_text(name.strip())
    normalized_text = normalize_text(text)
    
    # Ensure we're looking for complete words/names
    pattern = r'\b' + re.escape(normalized_name) + r'\b'
    matches = re.finditer(pattern, normalized_text)
    
    # Check if we have any matches
    match_positions = [m.start() for m in matches]
    return len(match_positions) > 0, match_positions

def find_matches_in_pdfs(names, directory="."):
    """
    Search for each name in all PDF files in the directory using character-by-character matching.
    Returns a dictionary mapping names to their corresponding numbers.
    """
    name_to_number = {}
    pdf_files = list(Path(directory).glob("*.pdf"))  # Use Path from pathlib
    
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in {directory}")
    
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
            raise
    
    # For each name, perform character-by-character search in the PDFs
    for name in names:
        if not name or pd.isna(name):
            continue
        
        print(f"Searching for '{name}'...")
        
        for pdf_path, text in pdf_texts.items():
            # Perform character-by-character search
            found, positions = character_by_character_search(str(name), text)
            
            if found:
                print(f"Found exact match for '{name}' in {pdf_path}")
                
                # For each position where the name was found, look for a number after it
                for pos in positions:
                    # Get the text after the name match
                    text_after_match = text[pos + len(normalize_text(str(name))):pos + len(normalize_text(str(name))) + 50]
                    
                    # Look for a number following the name
                    number_match = re.search(r'\s*(\d+)', text_after_match)
                    
                    if number_match:
                        # Fix the +1 issue by subtracting 1 from the found number
                        number = int(number_match.group(1)) - 1
                        name_to_number[name] = str(number)
                        print(f"  Associated number: {number}")
                        break
                
                # If we found a match with a number for this name, we can stop searching
                if name in name_to_number:
                    break
        
        if name not in name_to_number:
            print(f"No match or number found for '{name}'")
    
    return name_to_number

def process_excel_file(excel_path, pdf_directory):
    """Process the Excel file, find matches, and update SQN column."""
    try:
        # Load the Excel file
        df = pd.read_excel(excel_path)
        
        # Check if 'Name' column exists
        if 'Name' not in df.columns:
            raise ValueError("Column 'Name' not found in the Excel file")
            
        # Check if 'SQN' column exists
        if 'SQN' not in df.columns:
            raise ValueError("Column 'SQN' not found in the Excel file")
            
        # Get the list of names from 'Name' column
        names = df['Name'].tolist()
            
        # Find matches in PDFs
        name_to_number = find_matches_in_pdfs(names, pdf_directory)
        
        # Update 'SQN' column starting from row 2
        for i, name in enumerate(names, start=0):
            if i < 1:  # Skip the first row (header)
                continue
                
            if name in name_to_number:
                df.loc[i, 'SQN'] = name_to_number[name]
                print(f"Updated '{name}' with number: {name_to_number[name]}")
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
        raise Exception(f"Error processing Excel file: {e}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = LazySearchGUI(root)
    root.mainloop()