"""
ui/tab_column_base_mixin.py
============================
Column Base tab UI + data-loading helpers for manual CSV workflow
(coordinate and reaction data from file).

App state used: self.notebook, self.current_file,
                self.coordinate_software_var, self.reaction_software_var,
                self.include_rc_pier_var,
                self.coord_get_model_btn, self.coord_load_auto_btn,
                self.coord_disconnect_btn, self.coordinate_sap_path_label,
                self.coordinate_status_label, self.status_label
"""

import os
import csv
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox


class ColumnBaseTabMixin:

    def create_column_base_tab(self):
        """Create Column Base tab"""
        arrangement_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(arrangement_frame, text='📍 Column Base')

        # Center container
        center_container = tk.Frame(arrangement_frame, bg='white')
        center_container.pack(expand=True)

        # Icon and title
        title_label = tk.Label(
            center_container,
            text="📍 COLUMN BASE DATA",
            font=('Arial', 20, 'bold'),
            bg='white',
            fg='#1a472a'
        )
        title_label.pack(pady=(50, 20))

        # Description
        desc_label = tk.Label(
            center_container,
            text="Load column base data from your structural analysis model",
            font=('Arial', 11),
            bg='white',
            fg='#666666'
        )
        desc_label.pack(pady=(0, 20))

        # Software Selection Dropdown
        software_frame = tk.Frame(center_container, bg='white')
        software_frame.pack(pady=(0, 10))

        software_label = tk.Label(
            software_frame,
            text="Software:",
            font=('Arial', 11, 'bold'),
            bg='white',
            fg='#1a472a'
        )
        software_label.pack(side='left', padx=(0, 10))

        self.coordinate_software_var = tk.StringVar(value="SAP2000")
        software_menu = ttk.OptionMenu(
            software_frame,
            self.coordinate_software_var,
            "SAP2000",
            "SAP2000",
            "STAAD PRO"
        )
        software_menu.config(width=12)
        software_menu.pack(side='left')

        # Include RC Pier Selection
        pier_frame = tk.Frame(center_container, bg='white')
        pier_frame.pack(pady=(0, 15))

        pier_label = tk.Label(
            pier_frame,
            text="Include RC Pier?",
            font=('Arial', 11, 'bold'),
            bg='white',
            fg='#1a472a'
        )
        pier_label.pack(side='left', padx=(0, 10))

        self.include_rc_pier_var = tk.StringVar(value="No")
        pier_menu = ttk.OptionMenu(
            pier_frame,
            self.include_rc_pier_var,
            "No",
            "No",
            "Yes"
        )
        pier_menu.config(width=12)
        pier_menu.pack(side='left')

        # Buttons frame
        buttons_frame = tk.Frame(center_container, bg='white')
        buttons_frame.pack(pady=15)

        # Get Model button (Blue)
        self.coord_get_model_btn = tk.Button(
            buttons_frame,
            text="🏗️ GET MODEL",
            command=self.get_sap_model_coordinates,
            bg='#2196F3',
            fg='white',
            font=('Arial', 11, 'bold'),
            relief='flat',
            padx=30,
            pady=12,
            cursor='hand2',
            borderwidth=0
        )
        self.coord_get_model_btn.pack(side='left', padx=5)
        self.coord_get_model_btn.bind('<Enter>', lambda e: e.widget.config(bg='#1976D2') if e.widget['state'] != 'disabled' else None)
        self.coord_get_model_btn.bind('<Leave>', lambda e: e.widget.config(bg='#2196F3') if e.widget['state'] != 'disabled' else None)

        # Disconnect button (Orange)
        self.coord_disconnect_btn = tk.Button(
            buttons_frame,
            text="🔌 DISCONNECT",
            command=self.disconnect_sap_model_coord,
            bg='#FF9800',
            fg='white',
            font=('Arial', 11, 'bold'),
            relief='flat',
            padx=30,
            pady=12,
            cursor='hand2',
            borderwidth=0,
            state='disabled'
        )
        self.coord_disconnect_btn.pack(side='right', padx=5)
        self.coord_disconnect_btn.bind('<Enter>', lambda e: e.widget.config(bg='#E68900') if e.widget['state'] != 'disabled' else None)
        self.coord_disconnect_btn.bind('<Leave>', lambda e: e.widget.config(bg='#FF9800') if e.widget['state'] != 'disabled' else None)

        # Load Auto button (Green)
        self.coord_load_auto_btn = tk.Button(
            buttons_frame,
            text="📂 LOAD (AUTO)",
            command=self.load_coordinates_auto,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 11, 'bold'),
            relief='flat',
            padx=30,
            pady=12,
            cursor='hand2',
            borderwidth=0,
            state='disabled'
        )
        self.coord_load_auto_btn.pack(side='left', padx=5)
        self.coord_load_auto_btn.bind('<Enter>', lambda e: e.widget.config(bg='#45a049') if e.widget['state'] != 'disabled' else None)
        self.coord_load_auto_btn.bind('<Leave>', lambda e: e.widget.config(bg='#4CAF50') if e.widget['state'] != 'disabled' else None)

        # Status info - SAP Model Path
        self.coordinate_sap_path_label = tk.Label(
            center_container,
            text="",
            font=('Arial', 9, 'italic'),
            bg='white',
            fg='#2196F3'
        )
        self.coordinate_sap_path_label.pack(pady=(5, 0))

        # Status info - Data status
        self.coordinate_status_label = tk.Label(
            center_container,
            text="No coordinate data loaded",
            font=('Arial', 10, 'italic'),
            bg='white',
            fg='#999999'
        )
        self.coordinate_status_label.pack(pady=(10, 0))

        # Instructions
        inst_frame = tk.LabelFrame(
            center_container,
            text=" Instructions ",
            font=('Arial', 10, 'bold'),
            bg='white',
            fg='#1a472a',
            relief='solid',
            bd=1
        )
        inst_frame.pack(pady=(30, 0), padx=50, fill='x')

        instructions = """
1. Save your project first (File → Save)
2. Select software type (SAP2000 or STAAD PRO)
3. Click "GET MODEL": Connect to open SAP2000 model
4. Click "LOAD (AUTO)": Automatically extract Coordinates & Loading
  • Include RC Pier: select RC Pier + Steel Column only
  • Exclude RC Pier: select Steel Column only
  • Select Element only, Don't select Joint
5. Reaction file will be generated and opened automatically
        """

        inst_text = tk.Label(
            inst_frame,
            text=instructions,
            font=('Arial', 9),
            bg='white',
            fg='#424242',
            justify='left',
            anchor='w'
        )
        inst_text.pack(padx=15, pady=10, anchor='w')

    # ------------------------------------------------------------------
    # Manual CSV data loading (non-auto workflow)
    # ------------------------------------------------------------------

    def load_coordinate_data(self):
        """Load coordinate data from model analysis"""
        if not self.current_file:
            messagebox.showwarning("Warning",
                "Please save your project first!\n\n"
                "The CSV file will be created in the same folder as your project file.")
            return

        project_dir = os.path.dirname(self.current_file)
        software = self.coordinate_software_var.get()

        if software == "STAAD PRO":
            coord_file = os.path.join(project_dir, "ColumnBaseCoordinate_staad.csv")
        else:
            coord_file = os.path.join(project_dir, "ColumnBaseCoordinate_sap.csv")

        instruction = (
            "COLUMN BASE COORDINATES LOADING INSTRUCTIONS\n\n"
            "1. A csv file has been created at:\n"
            f"   {coord_file}\n\n"
            "2. The file will open automatically\n\n"
            "3. Paste your column base coordinates (X, Y, Z) starting from cell A1\n\n"
            "4. Save the csv file and close it\n\n"
            "5. Click OK below to confirm\n\n"
            "6. System will generate column_base_data.csv\n"
        )

        try:
            if platform.system() == 'Windows':
                os.startfile(coord_file)
            elif platform.system() == 'Darwin':
                subprocess.call(['open', coord_file])
            else:
                subprocess.call(['xdg-open', coord_file])
        except:
            pass

        result = messagebox.showinfo("Instructions", instruction)

        if result:
            self.coordinate_csv_file = coord_file
            self.generate_column_base_data_file(coord_file, project_dir)

            self.coordinate_status_label.config(
                text=f"✓ Coordinates loaded: {os.path.basename(coord_file)}",
                fg='#4CAF50'
            )
            self.status_label.config(text=f"● Column base coordinates ready", fg='#90ee90')

    def generate_column_base_data_file(self, coord_file, project_dir):
        """Generate column_base_data.csv from coordinate file"""
        try:
            data_file = os.path.join(project_dir, "column_base_data.csv")
            software = self.coordinate_software_var.get()

            base_coords = []
            with open(coord_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)

                if software == "SAP2000":
                    for _ in range(3):
                        next(reader, None)

                    for row in reader:
                        if len(row) >= 5:
                            node = row[0].strip()
                            x = row[3].strip()
                            y = row[4].strip()
                            z = row[5].strip() if len(row) > 5 else "0.0"
                            base_coords.append([node, x, y, z])
                else:
                    next(reader, None)
                    for row in reader:
                        if len(row) >= 4:
                            node = row[0].strip()
                            x = row[1].strip()
                            y_staad = row[2].strip()
                            z_staad = row[3].strip()

                            y_adjusted = z_staad
                            if z_staad.strip():
                                try:
                                    z_float = float(z_staad.strip())
                                    if z_float != 0:
                                        y_adjusted = str(-z_float)
                                except:
                                    pass

                            base_coords.append([node, x, y_adjusted, y_staad])

            with open(data_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Column Base Node', 'X (m)', 'Y (m)', 'Z (m)', 'Section Type'])
                for coord in base_coords:
                    writer.writerow(coord + [''])

            messagebox.showinfo("Success",
                f"column_base_data.csv created successfully!\n\n"
                f"Location: {data_file}\n\n"
                f"Total column bases: {len(base_coords)}\n\n"
                f"Please fill in 'Section Type' column manually")

            if platform.system() == 'Windows':
                os.startfile(data_file)
            elif platform.system() == 'Darwin':
                subprocess.call(['open', data_file])
            else:
                subprocess.call(['xdg-open', data_file])

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate column_base_data.csv:\n{str(e)}")

    def load_reaction_data(self):
        """Load reaction data from CSV file"""
        if not self.current_file:
            messagebox.showwarning("Warning",
                "Please save your project first!\n\n"
                "The CSV file will be created in the same folder as your project file.")
            return

        project_dir = os.path.dirname(self.current_file)
        software = self.reaction_software_var.get()

        if software == "STAAD PRO":
            csv_file = os.path.join(project_dir, "reaction_data_staadpro.csv")
            header_skip = 2
        else:
            csv_file = os.path.join(project_dir, "reaction_data_sap2000.csv")
            header_skip = 3

        instruction = (
            "REACTION DATA LOADING INSTRUCTIONS\n\n"
            "1. A csv file has been created at:\n"
            f"   {csv_file}\n\n"
            "2. The file will open automatically\n\n"
            "3. Paste your reaction loads starting from cell A1\n\n"
            "4. Save the csv file and close it\n\n"
            "5. Click OK below to confirm\n"
        )

        try:
            if platform.system() == 'Windows':
                os.startfile(csv_file)
            elif platform.system() == 'Darwin':
                subprocess.call(['open', csv_file])
            else:
                subprocess.call(['xdg-open', csv_file])
        except:
            pass

        result = messagebox.showinfo("Instructions", instruction)

        if result:
            self.reaction_csv_file = csv_file

            try:
                if platform.system() == 'Windows':
                    os.startfile(csv_file)
                elif platform.system() == 'Darwin':
                    subprocess.call(['open', csv_file])
                else:
                    subprocess.call(['xdg-open', csv_file])
            except:
                pass

            self.status_label.config(text=f"● Reaction CSV ready", fg='#90ee90')
