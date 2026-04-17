"""
ui/tab_material_mixin.py
=========================
Material Definition tab UI — main tab + 4 sub-tabs:
  - Material Strength  (2x2 grid of Anchor Bolt / Concrete / Rebar / Base Plate)
  - Anchor Bolt Table
  - Hinge Type
  - Rebar Dev Length

App state used: self.notebook, self.material_trees, self.material_data,
                self.material_status_label,
                self.anchor_bolt_material_tree, self.concrete_tree,
                self.rebar_tree, self.base_plate_material_tree
"""

import tkinter as tk
from tkinter import ttk


class MaterialTabMixin:

    def create_material_definition_tab(self):
        """Create Material Definition tab with 3 sub-tabs"""
        material_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(material_frame, text='📋 Material Definition')

        # Sub-notebook cho 3 tab
        sub_style = ttk.Style()
        sub_style.configure('Material.TNotebook', background='white')
        sub_style.configure('Material.TNotebook.Tab', font=('Arial', 10, 'bold'), padding=[15, 8])

        material_notebook = ttk.Notebook(material_frame, style='Material.TNotebook')
        material_notebook.pack(fill='both', expand=True, padx=20, pady=(20, 0))

        # Tạo 4 sub-tab
        self.create_material_strength_tab(material_notebook)
        self.create_anchor_bolt_tab(material_notebook)
        self.create_hinge_type_tab(material_notebook)
        self.create_rebar_dev_length_tab(material_notebook)

        # Bottom frame để đặt nút Refresh ở dưới cùng, align right
        bottom_frame = tk.Frame(material_frame, bg='white')
        bottom_frame.pack(fill='x', pady=(0, 20))

        tk.Button(
            bottom_frame,
            text="🔄 Refresh",
            command=self.reload_all_material_data,
            bg='#FF9800',
            fg='white',
            font=('Arial', 11, 'bold'),
            cursor='hand2',
            relief='flat',
            padx=20,
            pady=10
        ).pack(side='right', padx=30)

    def create_material_strength_tab(self, parent_notebook):
            """Create Material Strength sub-tab with 4 compact tables: Anchor Bolt, Concrete, Rebar, Base Plate"""
            frame = tk.Frame(parent_notebook, bg='white')
            parent_notebook.add(frame, text='⚙️ Material Strength')

            # Main container với 4 cột đều nhau (2 hàng x 2 cột)
            container = tk.Frame(frame, bg='white')
            container.pack(fill='both', expand=True, padx=30, pady=30)

            # Configure grid
            container.grid_rowconfigure(0, weight=1)
            container.grid_rowconfigure(1, weight=1)
            container.grid_columnconfigure(0, weight=1)
            container.grid_columnconfigure(1, weight=1)

            # Hàm helper để tạo bảng chung
            def create_material_table(parent, row, col, title_text, title_icon, columns, tree_name):
                table_frame = tk.Frame(parent, bg='white', relief='groove', bd=2)
                table_frame.grid(row=row, column=col, sticky='nsew', padx=15, pady=15)

                # Title căn giữa
                title_frame = tk.Frame(table_frame, bg='#f0f0f0')
                title_frame.pack(fill='x')
                tk.Label(
                    title_frame,
                    text=f'{title_icon} {title_text}',
                    font=('Arial', 14, 'bold'),
                    bg='#f0f0f0',
                    fg='#1a472a'
                ).pack(pady=10)

                # Treeview
                tree = ttk.Treeview(table_frame, columns=columns, height=6, show='headings')
                tree.column(columns[0], anchor=tk.W, width=250)
                tree.column(columns[1], anchor=tk.CENTER, width=120)
                tree.heading(columns[0], text=columns[0])
                tree.heading(columns[1], text=columns[1])

                scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
                tree.configure(yscroll=scrollbar.set)

                tree.pack(side='top', fill='both', expand=True, padx=10, pady=(0,10))
                scrollbar.pack(side='right', fill='y', padx=(0,10), pady=(0,10))

                # Buttons dưới bảng
                btn_frame = tk.Frame(table_frame, bg='white')
                btn_frame.pack(fill='x', pady=(0,10))
                tk.Button(btn_frame, text='➕ Add', command=lambda t=tree: self.add_row_to_material_tree(t, tree_name),
                          bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'), width=10).pack(side='left', padx=5)
                tk.Button(btn_frame, text='❌ Delete', command=lambda t=tree: self.delete_row_from_treeview(t),
                          bg='#f44336', fg='white', font=('Arial', 9, 'bold'), width=10).pack(side='left', padx=5)
                tk.Button(btn_frame, text='💾 Save', command=self.save_material_strength_data,
                          bg='#2196F3', fg='white', font=('Arial', 9, 'bold'), width=10).pack(side='left', padx=5)

                tree.bind('<Double-1>', lambda e, t=tree: self.edit_treeview_cell(e, t))

                # Lưu tree
                self.material_trees[tree_name] = tree
                return tree

            # Tạo 4 bảng (2x2 grid)
            self.anchor_bolt_material_tree = create_material_table(
                container, 0, 0, 'ANCHOR BOLT', '🔩',
                ('Material Type', 'futa (MPa)'), 'anchor_bolt'
            )

            self.concrete_tree = create_material_table(
                container, 0, 1, 'CONCRETE', '🏗️',
                ('Material Type', "f'c (MPa)"), 'concrete'
            )

            self.rebar_tree = create_material_table(
                container, 1, 0, 'REBAR', '⚙️',
                ('Material Type', 'fy (MPa)'), 'rebar'
            )

            self.base_plate_material_tree = create_material_table(
                container, 1, 1, 'BASE PLATE', '📐',
                ('Material Type', 'fy (MPa)'), 'base_plate'
            )

            # Load data khi mở tab
            self.reload_all_material_data()

    def create_anchor_bolt_tab(self, parent_notebook):
        """Create Anchor Bolt Table sub-tab"""
        frame = tk.Frame(parent_notebook, bg='white')
        parent_notebook.add(frame, text='🔩 Anchor Bolt Table')

        # Treeview
        columns = ('db', 'Rmin', 'a', 'W', 'T', 'S', 'NutW', 'nt', 'Nut Allowance', 'Edge Min', 'Leng A1', 'Leng A2')
        tree = ttk.Treeview(frame, columns=columns, height=15)
        tree.column('#0', width=0, stretch=tk.NO)

        for i, col in enumerate(columns):
            width = 80 if i < 2 else 70
            tree.column(col, anchor=tk.CENTER, width=width)
            tree.heading(col, text=col, anchor=tk.CENTER)

        # Scrollbars
        scrollbar_y = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        scrollbar_x = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        tree.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        scrollbar_y.grid(row=0, column=1, sticky='ns', pady=10)
        scrollbar_x.grid(row=1, column=0, sticky='ew', padx=10)

        # Buttons frame
        btn_frame = tk.Frame(frame, bg='white')
        btn_frame.grid(row=2, column=0, columnspan=2, sticky='ew', padx=10, pady=10)

        tk.Button(btn_frame, text='➕ Add Row', command=lambda: self.add_row_to_treeview(tree, 'Anchor Bolt Table', columns),
                 bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)
        tk.Button(btn_frame, text='❌ Delete Row', command=lambda: self.delete_row_from_treeview(tree),
                 bg='#f44336', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)
        tk.Button(btn_frame, text='💾 Save', command=lambda: self.save_material_data(tree, 'Anchor Bolt Table'),
                 bg='#2196F3', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)

        # Configure grid
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        # Load data from Excel
        self.load_anchor_bolt_data(tree)

        # Bind double-click to edit
        tree.bind('<Double-1>', lambda e: self.edit_treeview_cell(e, tree))

        self.material_data['Anchor Bolt Table'] = tree
        self.material_trees['Anchor Bolt Table'] = tree

    def create_hinge_type_tab(self, parent_notebook):
            """Create Hinge Type sub-tab"""
            frame = tk.Frame(parent_notebook, bg='white')
            parent_notebook.add(frame, text='🔩 Hinge Type')

            # Treeview - 22 cột (BỎ AB size, AB Hole, THÊM 10 cột mới)
            columns = ('Column size', 'Type', 'No.AB', 'P1', 'N', 'A', 'B', 'C', 'E', 'F', 'P2', 'Y',
                      'Np', 'Bp', 'c', 'nrb', 'drb', 'dtb', 'X-leg', 'Y-leg', 'Layer 1', 'Layer 2')
            tree = ttk.Treeview(frame, columns=columns, height=15)
            tree.column('#0', width=0, stretch=tk.NO)

            # Set column widths
            for i, col in enumerate(columns):
                if i == 0:  # Column size
                    width = 120
                else:
                    width = 70
                tree.column(col, anchor=tk.CENTER, width=width)
                tree.heading(col, text=col, anchor=tk.CENTER)

            # Scrollbars
            scrollbar_y = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
            scrollbar_x = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
            tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

            tree.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
            scrollbar_y.grid(row=0, column=1, sticky='ns', pady=10)
            scrollbar_x.grid(row=1, column=0, sticky='ew', padx=10)

            # Buttons frame
            btn_frame = tk.Frame(frame, bg='white')
            btn_frame.grid(row=2, column=0, columnspan=2, sticky='ew', padx=10, pady=10)

            tk.Button(btn_frame, text='➕ Add Row', command=lambda: self.add_row_to_treeview(tree, 'Hinge Type', columns),
                     bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)
            tk.Button(btn_frame, text='❌ Delete Row', command=lambda: self.delete_row_from_treeview(tree),
                     bg='#f44336', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)
            tk.Button(btn_frame, text='💾 Save', command=lambda: self.save_material_data(tree, 'Hinge Type'),
                     bg='#2196F3', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)

            # Configure grid
            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)

            # Load data from Excel
            self.load_hinge_type_data(tree)

            # Bind double-click to edit
            tree.bind('<Double-1>', lambda e: self.edit_treeview_cell(e, tree))

            self.material_data['Hinge Type'] = tree
            self.material_trees['Hinge Type'] = tree

    def create_rebar_dev_length_tab(self, parent_notebook):
            """Create Rebar Development Length sub-tab"""
            frame = tk.Frame(parent_notebook, bg='white')
            parent_notebook.add(frame, text='📏 Rebar Dev Length')

            # Treeview - 3 cột
            columns = ('Bars', 'Ld', 'Ldh')
            tree = ttk.Treeview(frame, columns=columns, height=15)
            tree.column('#0', width=0, stretch=tk.NO)

            for col in columns:
                tree.column(col, anchor=tk.CENTER, width=150)
                tree.heading(col, text=col, anchor=tk.CENTER)

            # Scrollbars
            scrollbar_y = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
            tree.configure(yscroll=scrollbar_y.set)

            tree.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
            scrollbar_y.grid(row=0, column=1, sticky='ns', pady=10)

            # Buttons frame
            btn_frame = tk.Frame(frame, bg='white')
            btn_frame.grid(row=1, column=0, columnspan=2, sticky='ew', padx=10, pady=10)

            tk.Button(btn_frame, text='➕ Add Row', command=lambda: self.add_row_to_treeview(tree, 'Rebar Development Length', columns),
                     bg='#4CAF50', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)
            tk.Button(btn_frame, text='❌ Delete Row', command=lambda: self.delete_row_from_treeview(tree),
                     bg='#f44336', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)
            tk.Button(btn_frame, text='💾 Save', command=lambda: self.save_material_data(tree, 'Rebar Development Length'),
                     bg='#2196F3', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack(side='left', padx=5)

            # Configure grid
            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)

            # Load data from Excel
            self.load_rebar_dev_length_data(tree)

            # Bind double-click to edit
            tree.bind('<Double-1>', lambda e: self.edit_treeview_cell(e, tree))

            self.material_data['Rebar Development Length'] = tree
            self.material_trees['Rebar Development Length'] = tree
