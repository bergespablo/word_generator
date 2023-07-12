# Import tkinter and other modules
import tkinter as tk
from tkinter import filedialog, scrolledtext
from tkinter.ttk import Combobox, Progressbar
from typing import List
import pandas as pd
from docxtpl import DocxTemplate
import threading
import configparser
import os
import re
import pythoncom
import webbrowser
from scripts.create_pdf_from_docx import create_pdf_from_docx


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        # Define a file name to store the data
        self.data_file = "config.ini"

        # Create a configparser object
        self.config = configparser.ConfigParser()

        # Create variables to store the file and folder paths
        self.word_file = tk.StringVar()
        self.excel_file = tk.StringVar()
        self.output_folder = tk.StringVar()

        # Stop generation
        self.stop_execution = tk.BooleanVar(value=False)

        # Create labels and entries for the file and folder fields
        self.word_label = tk.Label(self, text="Word template file:")
        self.word_label.grid(row=0, column=0, sticky=tk.W)
        self.word_entry = tk.Entry(self, textvariable=self.word_file,
                                   width=80, state="readonly")
        self.word_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.word_button = tk.Button(
            self, text="Browse", command=self.browse_word)
        self.word_button.grid(row=0, column=2)
        self.word_button_open = tk.Button(
            self, text="Open", command=self.open_word_file)
        self.word_button_open.grid(row=0, column=3)
        self.word_button_open.grid_remove()  # Initially invisible

        self.excel_label = tk.Label(self, text="Excel information file:")
        self.excel_label.grid(row=1, column=0, sticky=tk.W)
        self.excel_entry = tk.Entry(self, textvariable=self.excel_file,
                                    width=80, state="readonly")
        self.excel_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.excel_button = tk.Button(
            self, text="Browse", command=self.browse_excel)
        self.excel_button.grid(row=1, column=2)
        self.excel_button_open = tk.Button(
            self, text="Open", command=self.open_excel_file)
        self.excel_button_open.grid(row=1, column=3)
        self.excel_button_open.grid_remove()  # Initially invisible

        self.folder_label = tk.Label(self, text="Output folder:")
        self.folder_label.grid(row=2, column=0, sticky=tk.W)
        self.folder_entry = tk.Entry(self, textvariable=self.output_folder,
                                     width=80, state="readonly")
        self.folder_entry.grid(row=2, column=1, sticky=tk.W, padx=5)
        self.folder_button = tk.Button(
            self, text="Browse", command=self.browse_folder)
        self.folder_button.grid(row=2, column=2)
        self.folder_button_open = tk.Button(
            self, text="Open", command=self.open_folder)
        self.folder_button_open.grid(row=2, column=3)
        self.folder_button_open.grid_remove()  # Initially invisible

        # Create the Excel values combobox
        self.combobox_label = tk.Label(
            self, text="Column to use as file name:")
        self.combobox_label.grid(row=3, column=0, sticky=tk.W)
        self.combobox_entry = Combobox(self, state="readonly", width=37)
        self.combobox_entry.grid(row=3, column=1, sticky=tk.W, padx=5)
        self.combobox_label.grid_remove()  # Initially invisible
        self.combobox_entry.grid_remove()  # Initially invisible

        self.checkbutton_label = tk.Label(
            self, text="Type of files to generate:")
        self.checkbutton_label.grid(row=4, column=0, sticky=tk.W)
        self.checkbuttonframe = tk.Frame(self, border=2)
        self.checkbuttonframe.grid(
            row=4, column=1, columnspan=3, sticky=tk.W, padx=5)
        self.word_selected = tk.BooleanVar()
        self.pdf_selected = tk.BooleanVar()
        self.cb_word = tk.Checkbutton(self.checkbuttonframe, text="word",
                                      variable=self.word_selected, state="disabled")
        self.cb_pdf = tk.Checkbutton(self.checkbuttonframe, text="pdf",
                                     variable=self.pdf_selected)
        self.cb_word.grid(row=0, column=0)
        self.cb_word.select()
        self.cb_pdf.grid(row=0, column=1)

        # Create a button to generate files
        self.generate_button = tk.Button(self, text="Generate files",
                                         command=lambda: threading.Thread(target=self.generate_word_files).start())
        self.generate_button.grid(row=5, columnspan=4, pady=10)
        self.generate_button.configure(state='disabled')

        # Create a scrolledtext widget to display the logs
        self.log_area = scrolledtext.ScrolledText(self, width=100, height=10)
        self.log_area.grid(row=6, columnspan=4)
        self.log_area.configure(state='disabled')

        # Create a progressbar
        self.pb = Progressbar(self, orient='horizontal',
                              mode='determinate', length="800")
        self.pb.grid(row=7, columnspan=4)

        # Create a button to clear the log area
        self.clear_button = tk.Button(
            self, text="Clear", command=self.clear_screen)
        self.clear_button.grid(row=8, columnspan=4)

        # Create a button to clear the log area
        self.stop_button = tk.Button(self, text="Stop generation",
                                     command=self.stop_generation)
        self.stop_button.grid(row=9, columnspan=4)
        self.stop_button.grid_remove()  # Initially invisible

        # Call the load_data function at the beginning of the program
        self.load_data()

    def save_data(self):
        """Save the word file path, excel file path and output folder path in a config file
        """
        self.stop_execution.set(True)

        # Get the values of the file and folder paths
        word_file_path = self.word_file.get()
        excel_file_path = self.excel_file.get()
        output_folder_path = self.output_folder.get()

        # Create a section for the data
        self.config["DATA"] = {}

        # Set the values of the file and folder paths in the section
        self.config["DATA"]["word_path"] = word_file_path
        self.config["DATA"]["excel_path"] = excel_file_path
        self.config["DATA"]["folder_path"] = output_folder_path

        # Open the data file in write mode
        with open(self.data_file, "w") as f:
            # Use configparser to write the data into the file
            self.config.write(f)

        # Destroy the root window
        root.destroy()

    def load_data(self):
        """Load the word file path, excel file path and output folder path from the config file
        """

        # Try to read the data file
        try:
            self.config.read(self.data_file)
            # Get the values of the file and folder paths from the section
            word_path = self.config["DATA"]["word_path"]
            excel_path = self.config["DATA"]["excel_path"]
            folder_path = self.config["DATA"]["folder_path"]
            # Update the variables with the data
            if os.path.exists(word_path):
                self.word_file.set(word_path)
                self.word_button_open.grid()
            if os.path.exists(excel_path):
                self.excel_file.set(excel_path)
                self.excel_button_open.grid()
                self.load_combobox(excel_path)
            if os.path.exists(folder_path):
                self.output_folder.set(folder_path)
                self.folder_button_open.grid()
            self.checkPaths()
        # If the file does not exist or is corrupted, do nothing
        except:
            pass

    # Define a function to browse for a word file

    def browse_word(self):
        # Use filedialog to ask for a word file
        word_path = filedialog.askopenfilename(
            filetypes=[("Word files", "*.docx")])
        if word_path:
            # Update the word_file variable with the selected path
            self.word_file.set(word_path)
            self.word_button_open.grid()
            self.checkPaths()

    # Define a function to browse for an excel file

    def browse_excel(self):
        # Use filedialog to ask for an excel file
        excel_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")])
        if excel_path:
            # Update the excel_file variable with the selected path
            self.excel_file.set(excel_path)
            # Populate the combobox
            self.load_combobox(excel_path)
            self.excel_button_open.grid()
            self.checkPaths()

    def load_combobox(self, excel_path):
        # Populate the combobox with column names from the selected Excel file
        column_names = self.get_excel_column_names(excel_path)
        self.combobox_entry['values'] = column_names
        self.combobox_entry.current(0)
        # Show the combobox
        self.combobox_label.grid()
        self.combobox_entry.grid()

    # Define a function to browse for an output folder

    def browse_folder(self):
        """Browse for an output folder
        """

        # Use filedialog to ask for a folder
        folder_path = filedialog.askdirectory()
        if folder_path:
            # Update the output_folder variable with the selected path
            self.output_folder.set(folder_path)
            self.folder_button_open.grid()
            self.checkPaths()

    def checkPaths(self):
        """Enable the generation button if the word file path, excel file path, and output folder path are filled 
        """
        word_file_path = self.word_file.get()
        excel_file_path = self.excel_file.get()
        output_folder_path = self.output_folder.get()
        if word_file_path and excel_file_path and output_folder_path:
            self.generate_button.configure(state='normal')
        else:
            self.generate_button.configure(state='disabled')

    def generate_word_files(self):
        """Generate word and pdf files
        """
        pythoncom.CoInitialize()
        self.clear_screen()

        # Show stop button and hide clear button
        self.stop_button.grid()
        self.clear_button.grid_remove()

        # Get the values of the file and folder paths
        word_file_path = self.word_file.get()
        excel_file_path = self.excel_file.get()
        folder_path = self.output_folder.get()

        word_folder_path = folder_path + "/word"
        pdf_folder_path = folder_path + "/pdf"
        column_name = self.combobox_entry.get().strip().replace(' ', '_')

        # Check if the paths are valid
        if not word_file_path or not excel_file_path or not folder_path:
            # Display an error message in the log area
            self.writeLog("Please select valid files and folder.")
            return

        # Create a word directory if it does not exist
        if not os.path.exists(word_folder_path):
            os.makedirs(word_folder_path)
        if not os.path.exists(pdf_folder_path):
            os.makedirs(pdf_folder_path)

        # Display a message in the log area that the generation is starting
        self.writeLog(
            f"Generating word files from {word_file_path} and {excel_file_path} into {folder_path}...")

        # Write your logic to generate word files using the word template and the excel information
        doc = DocxTemplate(word_file_path)
        df = pd.read_excel(excel_file_path)
        all_columns = list(df)  # Creates list of all column headers
        df[all_columns] = df[all_columns].fillna('').astype(str)

        df.columns = df.columns.str.replace(' ', '_')
        num_word_files = 0
        num_word_errors = 0
        num_pdf_files = 0
        num_pdf_errors = 0
        self.writeLog(f"\n------- Generation of files ----------")
        for index, row in df.iterrows():
            if self.stop_execution.get():
                break
            context = {**row}
            doc.render(context)
            filename = f"row_index_{index+2}"
            column_value = row[column_name].strip()
            column_value = re.sub('[/\\@#/{/}]', '_', column_value, flags=re.I)

            if column_value != "":
                filename = column_value
            self.pb['value'] = 100*(index+1)/len(df)
            try:
                doc.save(f"{word_folder_path}\{filename}.docx")
                self.writeLog(f"File '{filename}.docx' correcty generated.")
                num_word_files = num_word_files+1
            except:
                self.writeLog(f"Error generating file '{filename}'.")
                num_word_errors = num_word_errors+1
            if (self.pdf_selected.get()):
                try:
                    create_pdf_from_docx(
                        f"{word_folder_path}/{filename}.docx", f"{pdf_folder_path}/{filename}.pdf")
                    self.writeLog(f"File '{filename}.pdf' correcty generated.")
                    num_pdf_files = num_pdf_files+1
                except:
                    self.writeLog(f"Error generating file '{filename}'.")
                    num_pdf_errors = num_pdf_errors+1

        # Display a message in the log area that the generation is done
        self.writeLog(f"\n------- Generation summary ----------")
        if (num_word_errors == 0 and num_pdf_errors == 0):
            self.folder_button_open.config(background='#a3f590')
        else:
            self.folder_button_open.config(background='yellow')

        if (num_word_files > 0):
            self.writeLog(f"Number of word files generated: {num_word_files}")
        if (num_word_errors > 0):
            self.writeLog(
                f"Number of word files not generated because of errors: {num_word_files}")
        if (num_pdf_files > 0):
            self.writeLog(f"Number of pdf files generated: {num_word_files}")
        if (num_pdf_errors > 0):
            self.writeLog(
                f"Number of pdf files not generated because of errors: {num_word_files}")

        # Hide stop button and show clear button
        self.stop_button.grid_remove()
        self.clear_button.grid()

    def writeLog(self, text: str):
        """Write to log area

        Args:
            text (str): Text to show in log area
        """
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, text+"\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def get_excel_column_names(self, file_path: str) -> List[str]:
        """Get column names from excel

        Args:
            file_path (str): Path to excel file

        Returns:
            List[str]: List of column names
        """
        try:
            df = pd.read_excel(file_path)
            return df.columns.tolist()
        except Exception as e:
            self.writeLog(f"Error reading Excel file: {str(e)}")
            return []

    def clear_screen(self):
        """Clear screen: log area and progress bar
        """

        # Delete all the text in the log area
        self.log_area.configure(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.configure(state='disabled')
        self.pb['value'] = 0
        self.folder_button_open.config(background='SystemButtonFace')
        self.stop_execution.set(False)
        self.checkPaths()

    def open_folder(self):
        folder = self.output_folder.get()
        if os.path.exists(folder):
            webbrowser.open(folder)
        else:
            self.writeLog(f'Error: The folder "{folder}" does not exist.')

    def open_excel_file(self):
        file = self.excel_file.get()
        if os.path.exists(file):
            webbrowser.open(file)
        else:
            self.writeLog(f'Error: The excel file "{file}" does not exist.')

    def open_word_file(self):
        file = self.word_file.get()
        if os.path.exists(file):
            webbrowser.open(file)
        else:
            self.writeLog(f'Error: The word file "{file}" does not exist.')

    def stop_generation(self):
        self.stop_execution.set(True)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Word File Generator")
    app = MainApplication(root)
    app.pack(side="top", fill="both", expand=True, padx=10, pady=10)
    # Call the save_data function before exiting the program
    root.protocol("WM_DELETE_WINDOW", app.save_data)
    # Start the main loop of the GUI
    root.mainloop()
