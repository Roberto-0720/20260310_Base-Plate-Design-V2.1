"""
data/file_manager_mixin.py
===========================
Project file I/O — new, open, save, save-as for the .json project file.

App state used: self.current_file, self.reaction_csv_file,
                self.coordinate_csv_file, self.bpl_folder,
                self.status_label
"""

import os
import json
from tkinter import filedialog, messagebox


class FileManagerMixin:

    def new_file(self):
        """Create new file"""
        self.current_file = None
        self.reaction_csv_file = None
        self.coordinate_csv_file = None

        self.status_label.config(text="● New file created", fg='#90ee90')

    def save_file(self):
        """Save current data"""
        if not self.current_file:
            self.save_as_file()
        else:
            self.save_data_to_file(self.current_file)

    def save_as_file(self):
        """Save data to new file"""
        initial_dir = None
        initial_file = "Bpl.json"
        if self.bpl_folder and os.path.exists(self.bpl_folder):
            initial_dir = self.bpl_folder

        file_path = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Save Base Plate Design Data"
        )
        if file_path:
            self.current_file = file_path
            self.save_data_to_file(file_path)

    def save_data_to_file(self, file_path):
        """Save all input data to JSON file"""
        try:
            data = {}

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)

            self.status_label.config(text=f"● File saved: {os.path.basename(file_path)}", fg='#90ee90')
            messagebox.showinfo("Success", f"Data saved successfully to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")
            self.status_label.config(text="● Save failed", fg='#ff4d4d')

    def open_file(self):
        """Open and load data from JSON file"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Open Base Plate Design Data"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                self.current_file = file_path
                self.status_label.config(text=f"● File loaded: {os.path.basename(file_path)}", fg='#90ee90')
                messagebox.showinfo("Success", f"Project loaded from:\n{file_path}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file:\n{str(e)}")
                self.status_label.config(text="● Load failed", fg='#ff4d4d')
