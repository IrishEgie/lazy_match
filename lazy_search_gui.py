import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox, filedialog
import subprocess
import os
import threading

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
        self.excel_drop_label = tk.Label(root, text="Drop Excel File Here", bg="lightgray", width=50, height=2)
        self.excel_drop_label.pack(pady=10)
        self.excel_drop_label.drop_target_register(DND_FILES)
        self.excel_drop_label.dnd_bind('<<Drop>>', self.on_excel_drop)

        # Drag & Drop for PDF directory
        self.pdf_drop_label = tk.Label(root, text="Drop PDF Directory Here", bg="lightgray", width=50, height=2)
        self.pdf_drop_label.pack(pady=10)
        self.pdf_drop_label.drop_target_register(DND_FILES)
        self.pdf_drop_label.dnd_bind('<<Drop>>', self.on_pdf_drop)

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

        # Run the lazy_search.py script in a separate thread
        def run_script():
            try:
                # Run the lazy_search.py script with the provided file paths
                result = subprocess.run(
                    ["python", "lazy_search.py", self.excel_file_path.get(), self.pdf_directory_path.get()],
                    capture_output=True, text=True
                )

                if result.returncode == 0:
                    self.status_label.config(text="Status: Lazy Search completed successfully!")
                    self.progress_label.config(text="Progress: 100%")
                    self.output_file_path.set(self.excel_file_path.get().replace('.xlsx', '_updated.xlsx'))
                    self.open_excel_button.config(state=tk.NORMAL)
                else:
                    self.status_label.config(text="Status: Lazy Search failed.")
                    self.progress_label.config(text="Progress: 0%")
                    messagebox.showerror("Error", result.stderr)

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
            os.startfile(self.output_file_path.get())
        else:
            messagebox.showerror("Error", "No updated Excel file found.")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = LazySearchGUI(root)
    root.mainloop()