"""
ui/main_window_mixin.py
========================
Top-level window chrome: header bar, menu bar, main content (notebook +
tab wiring), and status bar.

App state created here:
  self.notebook, self.status_label
"""

import tkinter as tk
from tkinter import ttk


class MainWindowMixin:

    def create_header(self):
        """Create header section"""
        header_frame = tk.Frame(self.root, bg='#1a472a', height=70)
        header_frame.pack(fill='x', padx=0, pady=0)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="BASE PLATE DESIGN",
            font=('Arial', 20, 'bold'),
            bg='#1a472a',
            fg='white'
        )
        title_label.pack(pady=(15, 0))

    def create_menu_bar(self):
        """Create menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Open", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Save", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As...", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        # Tools Menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Export to Excel", command=self.export_to_excel)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Guide", command=self.show_guide)
        help_menu.add_command(label="About", command=self.show_about)

        # Keyboard shortcuts
        self.root.bind('<Control-n>', lambda e: self.new_file())
        self.root.bind('<Control-o>', lambda e: self.open_file())
        self.root.bind('<Control-s>', lambda e: self.save_file())

    def create_main_content(self):
        """Create main content area with tabs"""
        main_container = tk.Frame(self.root, bg='#f5f5f5')
        main_container.pack(fill='both', expand=True, padx=12, pady=12)

        # Create notebook for tabs
        style = ttk.Style()
        style.configure('Custom.TNotebook', background='#f5f5f5')
        style.configure('Custom.TNotebook.Tab',
                       font=('Arial', 11, 'bold'),
                       padding=[20, 10])

        self.notebook = ttk.Notebook(main_container, style='Custom.TNotebook')
        self.notebook.pack(fill='both', expand=True)

        # Create tabs
        self.create_column_base_tab()
        self.create_material_definition_tab()
        self.create_base_plate_detail_tab()

        # Status bar at bottom
        self.create_status_bar()

        # Auto reload material data khi chuyển sang tab Material Definition
        def on_tab_changed(event):
            selected_tab = self.notebook.select()
            tab_text = self.notebook.tab(selected_tab, "text")
            if tab_text == '📋 Material Definition':
                self.reload_all_material_data()

        self.notebook.bind("<<NotebookTabChanged>>", on_tab_changed)

    def create_status_bar(self):
        """Create status bar at bottom"""
        status_frame = tk.Frame(self.root, bg='#2d5a3d', height=35)
        status_frame.pack(fill='x', side='bottom')
        status_frame.pack_propagate(False)

        self.status_label = tk.Label(
            status_frame,
            text="● Ready",
            font=('Arial', 10, 'bold'),
            bg='#2d5a3d',
            fg='#90ee90',
            anchor='w',
            padx=15
        )
        self.status_label.pack(side='left', fill='both', expand=True)
