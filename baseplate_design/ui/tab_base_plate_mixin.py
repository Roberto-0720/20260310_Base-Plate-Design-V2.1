"""
ui/tab_base_plate_mixin.py
===========================
"Base Plate Detail" tab: left control panel (node selector, dropdowns,
Apply/Edit/Copy buttons, Run Design Check) + right matplotlib canvas.

App state created here:
  self.fig, self.ax, self.canvas
  self.bpl_status_label, self.node_info_label, self.dropdown_frame
  self.apply_btn, self.edit_detail_btn, self.copy_multi_btn
  self.run_design_btn, self.label_display_var
"""

import tkinter as tk
from tkinter import ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure


class TabBasePlateMixin:

    def create_base_plate_detail_tab(self):
        """Create Base Plate Detail tab with interactive plan view"""
        diagram_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(diagram_frame, text='📊 Base Plate Detail')

        # Main container with 2 panels
        main_container = tk.Frame(diagram_frame, bg='white')
        main_container.pack(fill='both', expand=True, padx=10, pady=10)

        # LEFT PANEL (30%) - Control panel with scrollbar
        left_panel_outer = tk.Frame(main_container, bg='#f8f9fa', relief='groove', bd=2)
        left_panel_outer.pack(side='left', fill='both', padx=(0, 10), pady=0)
        left_panel_outer.pack_propagate(False)
        left_panel_outer.config(width=350)

        # Create Canvas and Scrollbar for left panel
        left_canvas = tk.Canvas(left_panel_outer, bg='#f8f9fa', highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_panel_outer, orient='vertical', command=left_canvas.yview)
        left_panel = tk.Frame(left_canvas, bg='#f8f9fa', width=335)

        left_panel.bind(
            '<Configure>',
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox('all'))
        )

        left_canvas.create_window((0, 0), window=left_panel, anchor='nw', width=335)
        left_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side='right', fill='y')
        left_canvas.pack(side='left', fill='both', expand=True)

        # RIGHT PANEL (70%) - Matplotlib canvas
        right_panel = tk.Frame(main_container, bg='white', relief='groove', bd=2)
        right_panel.pack(side='right', fill='both', expand=True)

        # ==================== LEFT PANEL CONTENT ====================

        # Title
        title_label = tk.Label(
            left_panel,
            text="📊 BASE PLATE DETAIL",
            font=('Arial', 14, 'bold'),
            bg='#f8f9fa',
            fg='#1a472a'
        )
        title_label.pack(pady=(15, 10))

        # Status section
        status_frame = tk.LabelFrame(
            left_panel,
            text=" Plan Status ",
            font=('Arial', 10, 'bold'),
            bg='#f8f9fa',
            fg='#1a472a'
        )
        status_frame.pack(fill='x', padx=15, pady=(0, 10))

        self.bpl_status_label = tk.Label(
            status_frame,
            text="No data loaded",
            font=('Arial', 10),
            bg='#f8f9fa',
            fg='#999999'
        )
        self.bpl_status_label.pack(padx=10, pady=8)

        reload_btn = tk.Button(
            status_frame,
            text="🔄 Reload Data",
            command=self.load_base_plate_plan,
            bg='#FF9800',
            fg='white',
            font=('Arial', 9, 'bold'),
            cursor='hand2',
            relief='flat',
            padx=15,
            pady=5
        )
        reload_btn.pack(pady=(0, 8))

        # Selected Node section
        node_frame = tk.LabelFrame(
            left_panel,
            text=" Selected Node ",
            font=('Arial', 10, 'bold'),
            bg='#f8f9fa',
            fg='#1a472a'
        )
        node_frame.pack(fill='both', expand=True, padx=15, pady=(0, 10))

        # Node info
        self.node_info_label = tk.Label(
            node_frame,
            text="Click a node to select",
            font=('Arial', 10, 'italic'),
            bg='#f8f9fa',
            fg='#999999',
            justify='left'
        )
        self.node_info_label.pack(padx=10, pady=10, anchor='w')

        # Dropdown frame (will be populated when node selected)
        self.dropdown_frame = tk.Frame(node_frame, bg='#f8f9fa')
        self.dropdown_frame.pack(fill='x', padx=10, pady=(0, 10))

        # Apply button
        self.apply_btn = tk.Button(
            node_frame,
            text="💾 Apply to Node",
            command=self.apply_node_settings,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10, 'bold'),
            cursor='hand2',
            relief='flat',
            padx=20,
            pady=8,
            state='disabled'
        )
        self.apply_btn.pack(pady=(5, 10))

        # Edit Node Detail button
        self.edit_detail_btn = tk.Button(
            node_frame,
            text="✏️ Edit Node Detail",
            command=self.edit_node_detail,
            bg='#FF9800',
            fg='white',
            font=('Arial', 10, 'bold'),
            cursor='hand2',
            relief='flat',
            padx=20,
            pady=8,
            state='disabled'
        )
        self.edit_detail_btn.pack(pady=(0, 10))

        # Copy to multiple button
        self.copy_multi_btn = tk.Button(
            node_frame,
            text="📋 Copy to Multiple Nodes",
            command=self.copy_to_multiple,
            bg='#2196F3',
            fg='white',
            font=('Arial', 9, 'bold'),
            cursor='hand2',
            relief='flat',
            padx=15,
            pady=6,
            state='disabled'
        )
        self.copy_multi_btn.pack(pady=(0, 10))

        # Run Design Check button (at bottom)
        run_frame = tk.Frame(left_panel, bg='#f8f9fa')
        run_frame.pack(side='bottom', fill='x', padx=15, pady=15)

        self.run_design_btn = tk.Button(
            run_frame,
            text="▶️ RUN DESIGN CHECK",
            command=self.run_design_check,
            bg='#1a472a',
            fg='white',
            font=('Arial', 11, 'bold'),
            cursor='hand2',
            relief='flat',
            padx=20,
            pady=12,
            state='disabled'
        )
        self.run_design_btn.pack(fill='x')

        # ==================== RIGHT PANEL - MATPLOTLIB ====================

        # Create matplotlib figure
        self.fig = Figure(figsize=(8, 6), dpi=100, facecolor='white')
        self.ax = self.fig.add_subplot(111)

        # Create canvas
        self.canvas = FigureCanvasTkAgg(self.fig, master=right_panel)
        self.canvas.draw()

        # Toolbar
        toolbar_frame = tk.Frame(right_panel, bg='white')
        toolbar_frame.pack(side='top', fill='x')
        toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        toolbar.update()

        # Pack canvas
        self.canvas.get_tk_widget().pack(side='top', fill='both', expand=True)

        # Legend & Display Label frame
        bottom_frame = tk.Frame(right_panel, bg='white')
        bottom_frame.pack(side='bottom', fill='x', pady=5)

        # Legend
        legend_frame = tk.Frame(bottom_frame, bg='white')
        legend_frame.pack(side='left', fill='x')

        tk.Label(legend_frame, text="Legend:", font=('Arial', 9, 'bold'), bg='white').pack(side='left', padx=10)

        legends = [
            ("⬤", "#CCCCCC", "Not Assigned"),
            ("⬤", "#2196F3", "Assigned"),
            ("⬤", "#4CAF50", "Design OK"),
            ("⬤", "#f44336", "Design NG")
        ]

        for symbol, color, label in legends:
            tk.Label(legend_frame, text=symbol, fg=color, font=('Arial', 12, 'bold'), bg='white').pack(side='left', padx=(10, 2))
            tk.Label(legend_frame, text=label, font=('Arial', 9), bg='white').pack(side='left', padx=(0, 10))

        # Display Label options frame
        label_frame = tk.LabelFrame(
            bottom_frame,
            text=" Display Label ",
            font=('Arial', 10, 'bold'),
            bg='#f0f0f0',
            fg='#1a472a',
            relief='solid',
            bd=1
        )
        label_frame.pack(side='left', padx=20)

        # Radio buttons for node label/ratio display
        tk.Radiobutton(
            label_frame,
            text="Label",
            variable=self.label_display_var,
            value="label",
            font=('Arial', 9),
            bg='#f0f0f0',
            command=self.update_plot_display
        ).pack(side='left', padx=8, pady=3)

        tk.Radiobutton(
            label_frame,
            text="Ratio",
            variable=self.label_display_var,
            value="ratio",
            font=('Arial', 9),
            bg='#f0f0f0',
            command=self.update_plot_display
        ).pack(side='left', padx=8, pady=3)

        tk.Radiobutton(
            label_frame,
            text="Both",
            variable=self.label_display_var,
            value="both",
            font=('Arial', 9),
            bg='#f0f0f0',
            command=self.update_plot_display
        ).pack(side='left', padx=8, pady=3)

        # Connect click event
        self.canvas.mpl_connect('button_press_event', self.on_node_click)

        # Initial plot
        self.plot_empty_state()

        # Auto-load data when tab is created
        self.root.after(500, self.load_base_plate_plan)
