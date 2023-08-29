import os
import re
from pathlib import Path
import threading
import configparser
import pythoncom
import webbrowser
import win32com.client
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from tkinter.ttk import Combobox, Progressbar
import pandas as pd
from docxtpl import DocxTemplate

class UI(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.word_file = tk.StringVar()
        self.excel_file = tk.StringVar()
        self.output_folder = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        self.word_label = tk.Label(self, text="Word template file:")
        self.word_label.grid(row=0, column=0, sticky=tk.W)
        self.word_entry = tk.Entry(
            self, textvariable=self.word_file, width=80, state="readonly")
        self.word_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.word_button = tk.Button(
            self, text="Browse", command=self.parent.browse_word)
        self.word_button.grid(row=0, column=2)
        self.word_button_open = tk.Button(
            self, text="Open", command=self.parent.open_word_file)
        self.word_button_open.grid(row=0, column=3)
        self.word_button_open.grid_remove()

        self.excel_label = tk.Label(self, text="Excel information file:")
        self.excel_label.grid(row=1, column=0, sticky=tk.W)
        self.excel_entry = tk.Entry(
            self, textvariable=self.excel_file, width=80, state="readonly")
        self.excel_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.excel_button = tk.Button(
            self, text="Browse", command=self.parent.browse_excel)
        self.excel_button.grid(row=1, column=2)
        self.excel_button_open = tk.Button(
            self, text="Open", command=self.parent.open_excel_file)
        self.excel_button_open.grid(row=1, column=3)
        self.excel_button_open.grid_remove()

        self.folder_label = tk.Label(self, text="Output folder:")
        self.folder_label.grid(row=2, column=0, sticky=tk.W)
        self.folder_entry = tk.Entry(
            self, textvariable=self.output_folder, width=80, state="readonly")
        self.folder_entry.grid(row=2, column=1, sticky=tk.W, padx=5)
        self.folder_button = tk.Button(
            self, text="Browse", command=self.parent.browse_folder)
        self.folder_button.grid(row=2, column=2)
        self.folder_button_open = tk.Button(
            self, text="Open", command=self.parent.open_folder)
        self.folder_button_open.grid(row=2, column=3)
        self.folder_button_open.grid_remove()

        self.combobox_label = tk.Label(
            self, text="Column to use as file name:")
        self.combobox_label.grid(row=3, column=0, sticky=tk.W)
        self.combobox_entry = Combobox(self, state="readonly", width=37)
        self.combobox_entry.grid(row=3, column=1, sticky=tk.W, padx=5)
        self.combobox_label.grid_remove()
        self.combobox_entry.grid_remove()

        self.generate_word_button = tk.Button(
            self, text="Generate word files", command=self.parent.generate_word_files_in_thread)
        self.generate_word_button.grid(row=4, columnspan=2, column=0, pady=10)
        self.generate_word_button.configure(state='disabled')

        self.generate_pdf_button = tk.Button(
            self, text="Generate pdf files", command=self.parent.generate_pdf_files_in_thread)
        self.generate_pdf_button.grid(row=4, columnspan=2, column=1, pady=10)
        self.generate_pdf_button.configure(state='disabled')

        self.log_area = scrolledtext.ScrolledText(self, width=100, height=10)
        self.log_area.grid(row=6, columnspan=4)
        self.log_area.configure(state='disabled')

        self.pb = Progressbar(self, orient='horizontal',
                              mode='determinate', length="800")
        self.pb.grid(row=7, columnspan=4)

        self.clear_button = tk.Button(
            self, text="Clear", command=self.parent.clear_screen)
        self.clear_button.grid(row=8, columnspan=4)

        self.stop_button = tk.Button(
            self, text="Stop generation", command=self.parent.stop_generation)
        self.stop_button.grid(row=8, columnspan=4)
        self.stop_button.grid_remove()

    def write_log(self, text: str):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, text + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.data_file = "config.ini"
        self.config = configparser.ConfigParser()

        self.ui = UI(self)
        self.ui.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        self.stop_execution = tk.BooleanVar(value=False)

        # self.load_data()

    def save_data(self):
        self.stop_execution.set(True)

        word_file_path = self.ui.word_file.get()
        excel_file_path = self.ui.excel_file.get()
        output_folder_path = self.ui.output_folder.get()

        self.config["DATA"] = {
            "word_path": word_file_path,
            "excel_path": excel_file_path,
            "folder_path": output_folder_path,
        }

        with open(self.data_file, "w") as f:
            self.config.write(f)

        root.destroy()

    def load_data(self):
        try:
            self.config.read(self.data_file)
            word_path = self.config["DATA"]["word_path"]
            excel_path = self.config["DATA"]["excel_path"]
            folder_path = self.config["DATA"]["folder_path"]

            if os.path.exists(word_path):
                self.ui.word_file.set(word_path)
                self.ui.word_button_open.grid()
            if os.path.exists(excel_path):
                self.ui.excel_file.set(excel_path)
                self.ui.excel_button_open.grid()
                self.load_combobox(excel_path)
            if os.path.exists(folder_path):
                self.ui.output_folder.set(folder_path)
                self.ui.folder_button_open.grid()
            self.checkPaths()
        except Exception:
            print(Exception)
        root.unbind('<Visibility>')

    def browse_word(self):
        if word_path := filedialog.askopenfilename(filetypes=[("Word files", "*.docx")]):
            self.ui.word_file.set(word_path)
            self.ui.word_button_open.grid()
            self.checkPaths()

    def browse_excel(self):
        if excel_path := filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]):
            self.ui.excel_file.set(excel_path)
            self.load_combobox(excel_path)
            self.ui.excel_button_open.grid()
            self.checkPaths()

    def browse_folder(self):
        if folder_path := filedialog.askdirectory():
            self.ui.output_folder.set(folder_path)
            self.ui.folder_button_open.grid()
            self.checkPaths()

    def load_combobox(self, excel_path):
        column_names = self.get_excel_column_names(excel_path)
        self.ui.combobox_entry['values'] = column_names
        self.ui.combobox_entry.current(0)
        self.ui.combobox_label.grid()
        self.ui.combobox_entry.grid()

    def checkPaths(self):
        word_file_path = self.ui.word_file.get()
        excel_file_path = self.ui.excel_file.get()
        output_folder_path = self.ui.output_folder.get()
        if os.path.exists(word_file_path) and os.path.exists(excel_file_path) and os.path.exists(output_folder_path):
            self.ui.generate_word_button.configure(state='normal')
        else:
            self.ui.generate_word_button.configure(state='disabled')

        doc_files = Path(output_folder_path).glob("[!~]*.doc*")
        if len(list(doc_files)) > 0:
            self.ui.generate_pdf_button.configure(state='normal')
        else:
            self.ui.generate_pdf_button.configure(state='disabled')

    def generate_word_files_in_thread(self):
        excel_file_path = self.ui.excel_file.get()
        df = pd.read_excel(excel_file_path)

        answer = messagebox.askokcancel(
            "Question", f"There are {len(df)} rows in the excel file '{excel_file_path}'.\n\nDo you want to generate these  {len(df)} word files?")
        if answer:
            # Disable the Generate button to prevent multiple generations at once
            self.clear_screen()
            self.ui.stop_button.grid()
            self.ui.clear_button.grid_remove()
            self.ui.generate_word_button.configure(state='disabled')
            # Start the file generation in a separate thread
            threading.Thread(target=self.generate_word_files).start()

    def generate_pdf_files_in_thread(self):
        output_folder = self.ui.output_folder.get()
        doc_files = Path(output_folder).glob("[!~]*.doc*")
        answer = messagebox.askokcancel(
            "Question", f"There are {len(list(doc_files))} word files in folder '{output_folder}/'.\n\nDo you want to generate pdf files from these word files?")
        if answer:
            # Disable the Generate button to prevent multiple generations at once
            self.clear_screen()
            self.ui.stop_button.grid()
            self.ui.clear_button.grid_remove()
            self.ui.generate_pdf_button.configure(state='disabled')
            # Start the file generation in a separate thread
            threading.Thread(target=self.generate_pdf_files).start()

    def generate_word_files(self):
        pythoncom.CoInitialize()

        word_file_path = self.ui.word_file.get()
        excel_file_path = self.ui.excel_file.get()
        folder_path = self.ui.output_folder.get()

        if not word_file_path or not excel_file_path or not folder_path:
            self.ui.write_log("Please select valid files and folder.")
            return

        self.ui.write_log(
            f"Generating word files from {word_file_path} and {excel_file_path} into {folder_path}...")

        doc = DocxTemplate(word_file_path)
        df = pd.read_excel(excel_file_path)
        df = self.clean_data_frame(df)
        self.generate_word_files_from_datafame(doc, df, folder_path)
        self.checkPaths()

    def clean_data_frame(self, df):
        all_columns = list(df)
        df[all_columns] = df[all_columns].fillna('').astype(str)
        df.columns = df.columns.str.replace(' ', '_')
        return df

    def generate_word_files_from_datafame(self, doc, df, word_folder_path):
        column_name = self.ui.combobox_entry.get().strip().replace(' ', '_')

        num_word_files = 0
        num_word_errors = 0

        self.ui.write_log(f"\n------- Generation of word files ----------")

        for index, row in df.iterrows():
            if self.stop_execution.get():
                break

            context = {**row}
            filename = self.generate_file_name(row, column_name)

            self.ui.pb['value'] = 100 * (index + 1) / len(df)

            word_file_path = os.path.join(word_folder_path, f"{filename}.docx")

            if self.generate_word_file(doc, context, word_file_path):
                num_word_files += 1
                self.ui.write_log(
                    f"File '{filename}.docx' correctly generated.")

            else:
                num_word_errors += 1
                self.ui.write_log(f"Error generating file '{filename}.docx'.")

        # Update the GUI elements after the file generation is completed
        self.after(0, self.update_log_with_generation_summary,
                   num_word_errors, num_word_files)
        self.after(0, self.ui.stop_button.grid_remove)
        self.after(0, self.ui.clear_button.grid)

        # Enable the Generate button again
        self.after(
            0, lambda: self.ui.generate_word_button.configure(state="normal"))

    def generate_pdf_files(self):

        output_folder_path = self.ui.output_folder.get()
        word = win32com.client.Dispatch(
            "Word.Application", pythoncom.CoInitialize())
        wdFormatPDF = 17

        num_pdf_files = 0
        num_pdf_errors = 0

        self.ui.write_log(
            f"Generating pdf files from word files in folder '{output_folder_path}'...")

        self.ui.write_log(f"\n------- Generation of pdf files ----------")
        doc_files = sorted(Path(output_folder_path).glob("[!~]*.doc*"))

        for index, docx_filepath in enumerate(doc_files):
            if self.stop_execution.get():
                break
            filename = str(docx_filepath.stem)
            pdf_filepath = Path(output_folder_path) / \
                (filename + ".pdf")
            try:
                doc = word.Documents.Open(str(docx_filepath))
                doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
                doc.Close(0)
                self.ui.write_log(
                    f"File '{filename}.pdf' correctly generated.")
                num_pdf_files += 1
            except Exception:
                num_pdf_errors += 1
                self.ui.write_log(f"Error generating file '{filename}.pdf'.")

            self.ui.pb['value'] = 100 * (index + 1) / len(doc_files)

        # Update the GUI elements after the file generation is completed
        self.after(0, self.update_log_with_generation_summary,
                   num_pdf_errors, num_pdf_files)
        self.after(0, self.ui.stop_button.grid_remove)
        self.after(0, self.ui.clear_button.grid)

        # Enable the Generate button again
        self.after(
            0, lambda: self.ui.generate_pdf_button.configure(state="normal"))

    def update_log_with_generation_summary(self, num_errors, num_files):
        self.ui.write_log(f"\n------- Generation summary ----------")

        if num_errors == 0:
            self.ui.folder_button_open.config(background='#a3f590')
        else:
            self.ui.folder_button_open.config(background='yellow')

        if num_files > 0:
            self.ui.write_log(
                f"Number of files generated: {num_files}")

        if num_errors > 0:
            self.ui.write_log(
                f"Number of files not generated because of errors: {num_errors}")

    def generate_file_name(self, row, column_name):
        filename = f"row_index_{row.name + 2}"
        column_value = row[column_name].strip()
        column_value = re.sub('[/\\@#/{/}]', '_', column_value, flags=re.I)

        if column_value != "":
            filename = column_value

        return filename

    def generate_word_file(self, doc, context, word_file_path):
        try:
            with open(word_file_path, "wb") as f:
                doc.render(context)
                doc.save(f)
            return True
        except Exception:
            return False

    def get_excel_column_names(self, file_path):
        try:
            df = pd.read_excel(file_path)
            return df.columns.tolist()
        except Exception as e:
            self.ui.write_log(f"Error reading Excel file: {str(e)}")
            return []

    def clear_screen(self):
        self.ui.log_area.configure(state='normal')
        self.ui.log_area.delete(1.0, tk.END)
        self.ui.log_area.configure(state='disabled')
        self.ui.pb['value'] = 0
        self.ui.folder_button_open.config(background='SystemButtonFace')
        self.stop_execution.set(False)
        self.checkPaths()

    def open_folder(self):
        folder = self.ui.output_folder.get()
        if os.path.exists(folder):
            webbrowser.open(folder)
        else:
            self.ui.write_log(f'Error: The folder "{folder}" does not exist.')

    def open_excel_file(self):
        file = self.ui.excel_file.get()
        if os.path.exists(file):
            webbrowser.open(file)
        else:
            self.ui.write_log(
                f'Error: The excel file "{file}" does not exist.')

    def open_word_file(self):
        file = self.ui.word_file.get()
        if os.path.exists(file):
            webbrowser.open(file)
        else:
            self.ui.write_log(f'Error: The word file "{file}" does not exist.')

    def stop_generation(self):
        self.stop_execution.set(True)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Word File Generator v0.0.4")
    root.resizable(False, False)
    app = MainApplication(root)
    app.pack(side="top", fill="both", expand=True, padx=10, pady=10)
    root.protocol("WM_DELETE_WINDOW", app.save_data)
    root.bind('<Visibility>', lambda e: app.load_data())
    root.mainloop()
