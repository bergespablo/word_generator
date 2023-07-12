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


def save_data():
    """Save the word file path, excel file path and output folder path in a config file
    """
    stop_execution.set(True)

    # Get the values of the file and folder paths
    word_file_path = word_file.get()
    excel_file_path = excel_file.get()
    output_folder_path = output_folder.get()

    # Create a section for the data
    config["DATA"] = {}

    # Set the values of the file and folder paths in the section
    config["DATA"]["word_path"] = word_file_path
    config["DATA"]["excel_path"] = excel_file_path
    config["DATA"]["folder_path"] = output_folder_path

    # Open the data file in write mode
    with open(data_file, "w") as f:
        # Use configparser to write the data into the file
        config.write(f)

    # Destroy the root window
    root.destroy()


def load_data():
    """Load the word file path, excel file path and output folder path from the config file
    """

    # Try to read the data file
    try:
        config.read(data_file)
        # Get the values of the file and folder paths from the section
        word_path = config["DATA"]["word_path"]
        excel_path = config["DATA"]["excel_path"]
        folder_path = config["DATA"]["folder_path"]
        # Update the variables with the data
        if os.path.exists(word_path):
            word_file.set(word_path)
            word_button_open.grid()
        if os.path.exists(excel_path):
            excel_file.set(excel_path)
            excel_button_open.grid()
            load_combobox(excel_path)
        if os.path.exists(folder_path):
            output_folder.set(folder_path)
            folder_button_open.grid()
        checkPaths()
    # If the file does not exist or is corrupted, do nothing
    except:
        pass


# Define a function to browse for a word file
def browse_word():
    # Use filedialog to ask for a word file
    word_path = filedialog.askopenfilename(
        filetypes=[("Word files", "*.docx")])
    if word_path:
        # Update the word_file variable with the selected path
        word_file.set(word_path)
        word_button_open.grid()
        checkPaths()

# Define a function to browse for an excel file


def browse_excel():
    # Use filedialog to ask for an excel file
    excel_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])
    if excel_path:
        # Update the excel_file variable with the selected path
        excel_file.set(excel_path)
        # Populate the combobox
        load_combobox(excel_path)
        excel_button_open.grid()
        checkPaths()


def load_combobox(excel_path):
    # Populate the combobox with column names from the selected Excel file
    column_names = get_excel_column_names(excel_path)
    combobox_entry['values'] = column_names
    combobox_entry.current(0)
    # Show the combobox
    combobox_label.grid()
    combobox_entry.grid()

# Define a function to browse for an output folder


def browse_folder():
    """Browse for an output folder
    """

    # Use filedialog to ask for a folder
    folder_path = filedialog.askdirectory()
    if folder_path:
        # Update the output_folder variable with the selected path
        output_folder.set(folder_path)
        folder_button_open.grid()
        checkPaths()


def checkPaths():
    """Enable the generation button if the word file path, excel file path, and output folder path are filled 
    """
    word_file_path = word_file.get()
    excel_file_path = excel_file.get()
    output_folder_path = output_folder.get()
    if word_file_path and excel_file_path and output_folder_path:
        generate_button.configure(state='normal')
    else:
        generate_button.configure(state='disabled')


def generate_word_files():
    """Generate word and pdf files
    """
    pythoncom.CoInitialize()
    clear_screen()

    # Show stop button and hide clear button
    stop_button.grid()
    clear_button.grid_remove()

    # Get the values of the file and folder paths
    word_file_path = word_file.get()
    excel_file_path = excel_file.get()
    folder_path = output_folder.get()

    word_folder_path = folder_path + "/word"
    pdf_folder_path = folder_path + "/pdf"
    column_name = combobox_entry.get().strip().replace(' ', '_')

    # Check if the paths are valid
    if not word_file_path or not excel_file_path or not folder_path:
        # Display an error message in the log area
        writeLog("Please select valid files and folder.")
        return

    # Create a word directory if it does not exist
    if not os.path.exists(word_folder_path):
        os.makedirs(word_folder_path)
    if not os.path.exists(pdf_folder_path):
        os.makedirs(pdf_folder_path)

    # Display a message in the log area that the generation is starting
    writeLog(
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
    writeLog(f"\n------- Generation of files ----------")
    for index, row in df.iterrows():
        if stop_execution.get():
            break
        context = {**row}
        doc.render(context)
        filename = f"row_index_{index+2}"
        column_value = row[column_name].strip()
        column_value = re.sub('[/\\@#/{/}]', '_', column_value, flags=re.I)

        if column_value != "":
            filename = column_value
        pb['value'] = 100*(index+1)/len(df)
        try:
            doc.save(f"{word_folder_path}\{filename}.docx")
            writeLog(f"File '{filename}.docx' correcty generated.")
            num_word_files = num_word_files+1
        except:
            writeLog(f"Error generating file '{filename}'.")
            num_word_errors = num_word_errors+1
        if (pdf_selected.get()):
            try:
                create_pdf_from_docx(
                    f"{word_folder_path}/{filename}.docx", f"{pdf_folder_path}/{filename}.pdf")
                writeLog(f"File '{filename}.pdf' correcty generated.")
                num_pdf_files = num_pdf_files+1
            except:
                writeLog(f"Error generating file '{filename}'.")
                num_pdf_errors = num_pdf_errors+1

    # Display a message in the log area that the generation is done
    writeLog(f"\n------- Generation summary ----------")
    if (num_word_errors == 0 and num_pdf_errors == 0):
        folder_button_open.config(background='#a3f590')
    else:
        folder_button_open.config(background='yellow')

    if (num_word_files > 0):
        writeLog(f"Number of word files generated: {num_word_files}")
    if (num_word_errors > 0):
        writeLog(
            f"Number of word files not generated because of errors: {num_word_files}")
    if (num_pdf_files > 0):
        writeLog(f"Number of pdf files generated: {num_word_files}")
    if (num_pdf_errors > 0):
        writeLog(
            f"Number of pdf files not generated because of errors: {num_word_files}")

    # Hide stop button and show clear button
    stop_button.grid_remove()
    clear_button.grid()


def writeLog(text: str):
    """Write to log area

    Args:
        text (str): Text to show in log area
    """
    log_area.configure(state='normal')
    log_area.insert(tk.END, text+"\n")
    log_area.see(tk.END)
    log_area.configure(state='disabled')


def get_excel_column_names(file_path: str) -> List[str]:
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
        print(f"Error reading Excel file: {str(e)}")
        return []


def clear_screen():
    """Clear screen: log area and progress bar
    """

    # Delete all the text in the log area
    log_area.configure(state='normal')
    log_area.delete(1.0, tk.END)
    log_area.configure(state='disabled')
    pb['value'] = 0
    folder_button_open.config(background='SystemButtonFace')
    stop_execution.set(False)
    checkPaths()


def open_folder():
    webbrowser.open(output_folder.get())


def open_excel_file():
    webbrowser.open(excel_file.get())


def open_word_file():
    webbrowser.open(word_file.get())


def stop_generation():
    stop_execution.set(True)


# Define a file name to store the data
data_file = "config.ini"

# Create a configparser object
config = configparser.ConfigParser()

# Create a root window
root = tk.Tk()
root.title("Word File Generator")

# Create a frame to hold the widgets
frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

# Create variables to store the file and folder paths
word_file = tk.StringVar()
excel_file = tk.StringVar()
output_folder = tk.StringVar()

# Stop generation
stop_execution = tk.BooleanVar(value=False)

# Create labels and entries for the file and folder fields
word_label = tk.Label(frame, text="Word template file:")
word_label.grid(row=0, column=0, sticky=tk.W)
word_entry = tk.Entry(frame, textvariable=word_file,
                      width=80, state="readonly")
word_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
word_button = tk.Button(frame, text="Browse", command=browse_word)
word_button.grid(row=0, column=2)
word_button_open = tk.Button(
    frame, text="Open", command=open_word_file)
word_button_open.grid(row=0, column=3)
word_button_open.grid_remove()  # Initially invisible

excel_label = tk.Label(frame, text="Excel information file:")
excel_label.grid(row=1, column=0, sticky=tk.W)
excel_entry = tk.Entry(frame, textvariable=excel_file,
                       width=80, state="readonly")
excel_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
excel_button = tk.Button(frame, text="Browse", command=browse_excel)
excel_button.grid(row=1, column=2)
excel_button_open = tk.Button(
    frame, text="Open", command=open_excel_file)
excel_button_open.grid(row=1, column=3)
excel_button_open.grid_remove()  # Initially invisible

folder_label = tk.Label(frame, text="Output folder:")
folder_label.grid(row=2, column=0, sticky=tk.W)
folder_entry = tk.Entry(frame, textvariable=output_folder,
                        width=80, state="readonly")
folder_entry.grid(row=2, column=1, sticky=tk.W, padx=5)
folder_button = tk.Button(frame, text="Browse", command=browse_folder)
folder_button.grid(row=2, column=2)
folder_button_open = tk.Button(
    frame, text="Open", command=open_folder)
folder_button_open.grid(row=2, column=3)
folder_button_open.grid_remove()  # Initially invisible


# Create the Excel values combobox
combobox_label = tk.Label(frame, text="Column to use as file name:")
combobox_label.grid(row=3, column=0, sticky=tk.W)
combobox_entry = Combobox(frame, state="readonly", width=37)
combobox_entry.grid(row=3, column=1, sticky=tk.W, padx=5)
combobox_label.grid_remove()  # Initially invisible
combobox_entry.grid_remove()  # Initially invisible

checkbutton_label = tk.Label(
    frame, text="Type of files to generate:")
checkbutton_label.grid(row=4, column=0, sticky=tk.W)
checkbuttonframe = tk.Frame(frame, border=2)
checkbuttonframe.grid(row=4, column=1, columnspan=3, sticky=tk.W, padx=5)
word_selected = tk.BooleanVar()
pdf_selected = tk.BooleanVar()
cb_word = tk.Checkbutton(checkbuttonframe, text="word",
                         variable=word_selected, state="disabled")
cb_pdf = tk.Checkbutton(checkbuttonframe, text="pdf",
                        variable=pdf_selected)
cb_word.grid(row=0, column=0)
cb_word.select()
cb_pdf.grid(row=0, column=1)


# Create a button to generate files
generate_button = tk.Button(frame, text="Generate files",
                            command=lambda: threading.Thread(target=generate_word_files).start())
generate_button.grid(row=5, columnspan=4, pady=10)
generate_button.configure(state='disabled')

# Create a scrolledtext widget to display the logs
log_area = scrolledtext.ScrolledText(frame, width=100, height=10)
log_area.grid(row=6, columnspan=4)
log_area.configure(state='disabled')

# Create a progressbar
pb = Progressbar(frame, orient='horizontal', mode='determinate', length="800")
pb.grid(row=7, columnspan=4)

# Create a button to clear the log area
clear_button = tk.Button(frame, text="Clear", command=clear_screen)
clear_button.grid(row=8, columnspan=4)

# Create a button to clear the log area
stop_button = tk.Button(frame, text="Stop generation",
                        command=stop_generation)
stop_button.grid(row=9, columnspan=4)
stop_button.grid_remove()  # Initially invisible

# Call the load_data function at the beginning of the program
load_data()


# Call the save_data function before exiting the program
root.protocol("WM_DELETE_WINDOW", save_data)

# Start the main loop of the GUI
root.mainloop()
