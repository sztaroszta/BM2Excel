"""
BM2Excel

Description:
    A tool for processing and exporting bookmarks from an HTML file to an Excel file.
    The program executes the following numbered steps:
      1. Open a file dialog for the user to select the HTML bookmarks file.
      2. Parse the selected HTML file with BeautifulSoup to extract bookmark and folder data.
      3. Present a GUI to allow the user to select specific bookmark folders for processing.
      4. Process and extract bookmark details (title, URL, folder, date added) recursively from the HTML.
      5. Filter the bookmarks based on the user-selected folders.
      6. Split the folder paths into hierarchical levels (separate columns) for easier reporting.
      7. Optionally include the input file name as a column in the resulting output.
      8. Save the organized bookmark data to an Excel workbook (using xlsxwriter) with proper formatting.
      9. Display a final summary dialog with details about the processed bookmarks.

Usage:
    - Ensure all required libraries are installed:
          pip install beautifulsoup4 pandas xlsxwriter
    - Run the script using:
          python bookmark2excel.py
    - Adjust constants (e.g., DEFAULT_FOLDER, MAX_FOLDER_LEVELS, DATE_FORMAT) as needed.

Author: Vitalii Starosta
GitHub: https://github.com/sztaroszta
License: GNU Affero General Public License v3 (AGPLv3)
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from bs4 import BeautifulSoup
import pandas as pd
import os
import sys
import xlsxwriter
from datetime import datetime
from typing import List, Dict, Tuple, Optional

# Global configuration: constants used across the application.
DEFAULT_FOLDER = "All Bookmarks"
MAX_FOLDER_LEVELS = 10
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

class CancelException(Exception):
    """Custom exception to signal that processing has been canceled by the user."""
    pass

class BookmarkProcessor:
    """
    Processes bookmarks from an HTML file and exports them to an Excel workbook.
    
    The processing workflow includes:
      - Selecting an HTML file containing bookmarks.
      - Allowing the user to select target folders via a GUI.
      - Extracting bookmark details from the HTML.
      - Filtering and organizing bookmarks.
      - Optionally appending the input file name.
      - Exporting the final data to a formatted Excel file.
      - Displaying a summary upon completion.
    """
    
    def __init__(self):
        """Initializes the BookmarkProcessor and sets the cancellation flag."""
        self.cancelled = False
    
    @staticmethod
    def get_file_path() -> str:
        """
        Opens a file dialog for selection of the HTML bookmarks file.
        
        Returns:
            str: The full file path of the selected HTML file.
        
        Exits:
            Exits the program if no file is selected.
        """
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="Select HTML Bookmarks File",
            filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")]
        )
        root.destroy()
        if not file_path:
            print("No file selected. Exiting.")
            sys.exit(1)
        return file_path

    @staticmethod
    def get_save_file_path(default_name: str = "bookmarks.xlsx") -> str:
        """
        Opens a file dialog to allow the user to specify where to save the output Excel file.
        
        Args:
            default_name (str): The default file name for the output Excel file.
        
        Returns:
            str: The save file path for the Excel workbook.
        
        Exits:
            Exits the program if no save location is selected.
        """
        root = tk.Tk()
        root.withdraw()
        save_file_path = filedialog.asksaveasfilename(
            title="Save Bookmarks As",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx"), ("All Files", "*.*")]
        )
        root.destroy()
        if not save_file_path:
            print("No save file chosen. Exiting.")
            sys.exit(1)
        return save_file_path

    def select_folders_with_confirm(self, soup: BeautifulSoup) -> List[str]:
        """
        Displays a GUI to let the user select the bookmark folders to process.
        
        The window is divided into three panels:
          - Left Panel: Displays available folders from the HTML.
          - Middle Panel: Contains buttons to add or remove folders.
          - Right Panel: Displays folders selected by the user.
        
        Args:
            soup (BeautifulSoup): The BeautifulSoup object of the parsed HTML bookmarks file.
        
        Returns:
            List[str]: A list of folder names selected by the user.
        """
        # Extract unique folder names from <h3> tags.
        folder_names = {tag.text.strip() for tag in soup.find_all("h3") if tag.text.strip()}
        available = sorted(folder_names)
        if DEFAULT_FOLDER not in available:
            available.insert(0, DEFAULT_FOLDER)
        else:
            # Place DEFAULT_FOLDER at the top.
            available.sort(key=lambda x: (x != DEFAULT_FOLDER, x))
        
        # Create the selection window.
        sel_root = tk.Tk()
        sel_root.title("Select Folders to Process")
        sel_root.geometry("750x500")
        sel_root.protocol("WM_DELETE_WINDOW", lambda: self.on_window_close(sel_root))
        
        # Configure grid layout.
        sel_root.grid_columnconfigure(0, weight=1, uniform="panels")
        sel_root.grid_columnconfigure(2, weight=1, uniform="panels")
        sel_root.grid_rowconfigure(0, weight=1)
        sel_root.grid_rowconfigure(1, weight=0)
        
        # Left Panel: Show available folders.
        left_frame = tk.Frame(sel_root, padx=10, pady=10)
        left_frame.grid(row=0, column=0, sticky="nsew")
        left_frame.grid_propagate(False)
        left_frame.config(width=300, height=300)
        tk.Label(left_frame, text="Available Folders:", font=('Arial', 12, 'bold')).pack(anchor="w")
        search_var = tk.StringVar()
        search_entry = tk.Entry(left_frame, textvariable=search_var, font=('Arial', 11), width=50)
        search_entry.pack(fill=tk.X, pady=5)
        search_entry.focus_set()
        avail_listbox = tk.Listbox(
            left_frame,
            selectmode=tk.EXTENDED,
            exportselection=False,
            width=60,
            height=25,
            font=('Arial', 13),
            activestyle='dotbox'
        )
        avail_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(avail_listbox)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        avail_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=avail_listbox.yview)
        
        # Middle Panel: Buttons to add or remove folder selections.
        mid_frame = tk.Frame(sel_root, padx=10)
        mid_frame.grid(row=0, column=1, sticky="ns")
        def add_selected():
            for i in avail_listbox.curselection():
                item = avail_listbox.get(i)
                if item not in selected_listbox.get(0, tk.END):
                    selected_listbox.insert(tk.END, item)
        def remove_selected():
            for i in reversed(selected_listbox.curselection()):
                selected_listbox.delete(i)
        tk.Button(mid_frame, text="Add >>", command=add_selected, width=12, font=('Arial', 10)).pack(pady=10)
        tk.Button(mid_frame, text="<< Remove", command=remove_selected, width=12, font=('Arial', 10)).pack(pady=10)
        
        # Right Panel: List the folders selected by the user.
        right_frame = tk.Frame(sel_root, padx=10, pady=10)
        right_frame.grid(row=0, column=2, sticky="nsew")
        right_frame.grid_propagate(False)
        right_frame.config(width=300, height=300)
        tk.Label(right_frame, text="Selected Folders:", font=('Arial', 12, 'bold')).pack(anchor="w")
        selected_listbox = tk.Listbox(
            right_frame,
            selectmode=tk.EXTENDED,
            exportselection=False,
            width=60,
            height=25,
            font=('Arial', 13),
            activestyle='dotbox'
        )
        selected_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(selected_listbox)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        selected_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=selected_listbox.yview)
        
        # Bottom Panel: Confirm or cancel folder selection.
        bottom_frame = tk.Frame(sel_root, pady=15)
        bottom_frame.grid(row=1, column=0, columnspan=3)
        selected_result = []
        def on_confirm():
            nonlocal selected_result
            selected_result = list(selected_listbox.get(0, tk.END))
            if not selected_result:
                selected_result = [DEFAULT_FOLDER]
            sel_root.destroy()
        def on_cancel():
            self.cancelled = True
            sel_root.destroy()
            sys.exit(0)
        confirm_btn = tk.Button(bottom_frame, text="Confirm", width=15,
                                  command=on_confirm, default=tk.ACTIVE, font=('Arial', 10, 'bold'))
        confirm_btn.pack(side=tk.LEFT, padx=20)
        cancel_btn = tk.Button(bottom_frame, text="Cancel", width=15,
                                 command=on_cancel, font=('Arial', 10))
        cancel_btn.pack(side=tk.RIGHT, padx=20)
        
        # Update available folders based on search.
        def update_avail_listbox(*args):
            query = search_var.get().lower()
            avail_listbox.delete(0, tk.END)
            for folder in available:
                if query in folder.lower():
                    avail_listbox.insert(tk.END, folder)
        search_var.trace("w", update_avail_listbox)
        update_avail_listbox()
        sel_root.bind("<Return>", lambda e: on_confirm())
        sel_root.mainloop()
        
        return selected_result

    def on_window_close(self, window):
        """
        Handles the closing of a window by setting the cancellation flag and exiting.
        
        Args:
            window (tk.Tk): The Tkinter window being closed.
        """
        self.cancelled = True
        window.destroy()
        sys.exit(0)

    def process_bookmarks(
        self, elements: List, parent_folder: str = "", bookmarks: Optional[Dict] = None
    ) -> List[Tuple]:
        """
        Recursively processes HTML elements to extract bookmark data.
        
        The function visits each HTML element in the provided list, extracting:
          - The bookmark title from anchor (<a>) elements.
          - The bookmark URL.
          - The folder structure (built recursively from <h3> tags).
          - The date the bookmark was added (formatted according to DATE_FORMAT).
        
        Args:
            elements (List): A list of BeautifulSoup elements to be processed.
            parent_folder (str): The current folder path (used for recursive calls).
            bookmarks (Optional[Dict]): Internal dictionary used during recursion to avoid duplicates.
        
        Returns:
            List[Tuple]: A list of tuples in the format (Bookmark, URL, Folder, Date_Added).
        """
        if bookmarks is None:
            bookmarks = {}
        for item in elements:
            if self.cancelled:
                raise CancelException("Processing cancelled by user.")
            if item.name == 'a':
                href = item.get('href', '')
                if href and href not in bookmarks:
                    date_added = item.get('add_date', '')
                    try:
                        if date_added:
                            timestamp = int(date_added)
                            if timestamp > 0:
                                date_str = datetime.fromtimestamp(timestamp).strftime(DATE_FORMAT)
                            else:
                                date_str = ''
                        else:
                            date_str = ''
                    except (ValueError, TypeError):
                        date_str = ''
                    bookmarks[href] = (
                        item.text.strip(),
                        href,
                        parent_folder,
                        date_str
                    )
            elif item.name == 'h3':
                folder_name = item.text.strip().replace("/", "___")
                dl_tag = item.find_next_sibling("dl")
                if dl_tag:
                    new_parent = f"{parent_folder}/{folder_name}" if parent_folder else folder_name
                    self.process_bookmarks(dl_tag.find_all(["a", "h3"]), new_parent, bookmarks)
        return list(bookmarks.values())

    def filter_bookmarks(self, df: pd.DataFrame, selected_folders: List[str]) -> pd.DataFrame:
        """
        Filters the bookmarks DataFrame based on user-selected folders.
        
        If the DEFAULT_FOLDER is selected, all bookmarks are returned. Otherwise, only bookmarks whose 
        folder path contains any of the selected folder names (case-insensitive) are kept.
        
        Args:
            df (pd.DataFrame): DataFrame containing bookmark data.
            selected_folders (List[str]): List of folders selected by the user.
        
        Returns:
            pd.DataFrame: A filtered DataFrame containing only bookmarks in the selected folders.
        """
        if DEFAULT_FOLDER in [s.strip() for s in selected_folders]:
            return df
        selected_set = {s.strip().lower() for s in selected_folders}
        def is_in_selected(folder_path: str) -> bool:
            if not folder_path:
                return False
            return any(part.lower() in selected_set for part in folder_path.split("/"))
        return df[df["Folder"].apply(is_in_selected)]

    def split_folder_levels(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Splits the folder path into separate hierarchical levels in the DataFrame.
        
        Each level is output to a new column (Level_1, Level_2, etc.), up to the maximum defined by MAX_FOLDER_LEVELS.
        
        Args:
            df (pd.DataFrame): The DataFrame containing bookmark data.
        
        Returns:
            pd.DataFrame: The DataFrame updated to include separate columns for each folder level.
        """
        folders_split = df['Folder'].str.split('/', expand=True)
        num_columns = min(folders_split.shape[1], MAX_FOLDER_LEVELS)
        for i in range(num_columns):
            df[f"Level_{i+1}"] = folders_split[i].str.replace("___", "/")
        return df

    def ask_store_input_filename(self, file_path: str) -> bool:
        """
        Asks the user if they wish to include the input file's name as the first column in the output Excel file.
        
        Args:
            file_path (str): The path to the input HTML file.
        
        Returns:
            bool: True if the user wants to include the file name, otherwise False.
        """
        root = tk.Tk()
        root.withdraw()
        response = messagebox.askyesno(
            "Store Input File Name",
            "Do you want to store the input file's name as the first column in the output?"
        )
        root.destroy()
        return response

    def save_to_excel(self, df: pd.DataFrame, excel_path: str):
        """
        Saves the processed bookmark DataFrame to an Excel file with proper formatting.
        
        The method sets column widths and applies specific formats (for URLs and dates) to the Excel worksheet.
        
        Args:
            df (pd.DataFrame): The DataFrame containing bookmark data.
            excel_path (str): The destination file path for the Excel workbook.
        """
        # Convert the Date_Added column to datetime objects.
        df['Date_Added'] = pd.to_datetime(df['Date_Added'], errors='coerce', format=DATE_FORMAT)
        writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Bookmarks')
        workbook = writer.book
        worksheet = writer.sheets['Bookmarks']

        # Define Excel cell formats for URLs and dates.
        url_format = workbook.add_format({'font_color': 'blue', 'underline': 1})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

        # Apply formatting based on column names.
        column_settings = [
            ("URL", 40, url_format),
            ("Date_Added", 20, date_format),
            ("Folder", 40, None),
            ("Bookmark", 60, None)
        ]
        for idx, col in enumerate(df.columns):
            for name, width, fmt in column_settings:
                if col == name:
                    worksheet.set_column(idx, idx, width, fmt)
                    break
            else:
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)
        writer.close()

    def show_summary(self, num_imported: int, output_path: str):
        """
        Displays a GUI summary dialog showing the outcome of the bookmark processing.
        
        Args:
            num_imported (int): The total number of bookmarks processed.
            output_path (str): The file path where the Excel file was saved.
        """
        summary_root = tk.Tk()
        summary_root.title("Processing Complete")
        message = (
            f"Processing complete!\n\n"
            f"Total bookmarks imported: {num_imported}\n"
            f"Saved to: {output_path}"
        )
        tk.Label(summary_root, text=message, padx=20, pady=20,
                 font=('Arial', 11), justify=tk.LEFT).pack()
        tk.Button(summary_root, text="OK", command=summary_root.destroy,
                  width=15, font=('Arial', 10)).pack(pady=10)
        summary_root.mainloop()

    def run(self):
        """
        Orchestrates the entire bookmark processing workflow.
        
        The method follows these steps:
          1. Prompt the user to select the input HTML file.
          2. Parse the HTML file.
          3. Display a GUI to let the user choose bookmark folders.
          4. Process the bookmarks from the parsed HTML.
          5. Filter and reorganize the bookmark data.
          6. Optionally add the input file name as a column.
          7. Save the processed data to an Excel file.
          8. Display a summary of the processing results.
        """
        try:
            # Step 1: Get the HTML file from the user.
            file_path = self.get_file_path()

            # Step 2: Parse the HTML file using BeautifulSoup.
            with open(file_path, "r", encoding="utf-8") as f:
                soup = BeautifulSoup(f.read(), 'html.parser')

            # Step 3: Ask the user to select which bookmark folders to process.
            selected_folders = self.select_folders_with_confirm(soup)
            print("Selected folders:", selected_folders)

            # Step 4: Process the bookmarks (recursive extraction).
            all_bookmarks = self.process_bookmarks(soup.find_all(["a", "h3"]))
            df = pd.DataFrame(all_bookmarks, columns=["Bookmark", "URL", "Folder", "Date_Added"])

            # Step 5: Filter the bookmarks based on the selected folders.
            df = self.filter_bookmarks(df, selected_folders)
            df = self.split_folder_levels(df)

            # Step 6: Optionally include the input file name.
            if self.ask_store_input_filename(file_path):
                df.insert(0, "Input File", os.path.basename(file_path))

            # Step 7: Save the bookmarks to an Excel file.
            default_name = f"bookmarks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            excel_path = self.get_save_file_path(default_name)
            self.save_to_excel(df, excel_path)
            print("Saved to:", excel_path)

            # Step 8: Show a summary of the process.
            self.show_summary(len(df), excel_path)
        except CancelException:
            print("Processing cancelled by user.")
            sys.exit(0)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            sys.exit(1)

# Entry point: This is where the script is executed.
if __name__ == "__main__":
    processor = BookmarkProcessor()
    processor.run()
