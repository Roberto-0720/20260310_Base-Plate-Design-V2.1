"""
logic/plot_manager_mixin.py
============================
Matplotlib visualization for the Base Plate plan view.
Handles plotting, node highlighting, and label display.

App state used: self.ax, self.canvas, self.scatter, self.label_texts,
                self.base_plate_nodes, self.selected_node,
                self.label_display_var, self.bpl_folder,
                self.bpl_status_label, self.run_design_btn, self.status_label
"""

import os
import csv
import math
import matplotlib.pyplot as plt
from tkinter import messagebox


class PlotManagerMixin:

    def plot_empty_state(self):
        """Plot empty state message"""
        self.ax.clear()
        self.ax.text(0.5, 0.5, 'No base plate data loaded\n\nClick "Reload Data" to load from bpl_coordinate.csv',
                    ha='center', va='center', fontsize=12, color='#999999',
                    transform=self.ax.transAxes)
        self.ax.set_xlim(0, 1)
        self.ax.set_ylim(0, 1)
        self.ax.axis('off')
        self.canvas.draw()

    def load_base_plate_plan(self):
        """Load base plate coordinates from CSV and plot"""
        if not self.bpl_folder or not os.path.exists(self.bpl_folder):
            messagebox.showwarning("Warning",
                "Base Plate Design folder not found!\n\n"
                "Please connect to SAP2000 model first (Column Base tab)")
            return

        csv_file = os.path.join(self.bpl_folder, "bpl_coordinate.csv")
        if not os.path.exists(csv_file):
            messagebox.showwarning("Warning",
                f"bpl_coordinate.csv not found!\n\n"
                f"Please load column base data first (Column Base tab)")
            return

        try:
            # Read CSV
            self.base_plate_nodes = {}
            with open(csv_file, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    node_name = row['Column Base'].strip()
                    x = float(row['X (m)'].strip())
                    y = float(row['Y (m)'].strip())
                    section = row.get('Section', '').strip()
                    beta = int(float(row.get('Beta', '0').strip() or '0'))

                    self.base_plate_nodes[node_name] = {
                        'x': x,
                        'y': y,
                        'section': section,
                        'beta': beta,
                        'material_anchor_bolt': None,
                        'material_base_plate': None,
                        'material_concrete': None,
                        'material_rebar': None,
                        'bolt_size': None,
                        'hinge_fixed_type': None,
                        'detail_type': None,
                        'design_status': 'Not Checked'
                    }

            if not self.base_plate_nodes:
                messagebox.showwarning("Warning", "No valid data found in bpl_coordinate.csv!")
                return

            # Update status
            self.bpl_status_label.config(
                text=f"✓ {len(self.base_plate_nodes)} nodes loaded",
                fg='#4CAF50'
            )

            # Enable run button
            self.run_design_btn.config(state='normal')

            # Plot
            self.plot_base_plate_plan()

            self.status_label.config(text=f"● Base plate plan loaded: {len(self.base_plate_nodes)} nodes", fg='#90ee90')

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load base plate data:\n{str(e)}")

    def plot_base_plate_plan(self):
        """Plot all base plate nodes"""
        self.ax.clear()

        if not self.base_plate_nodes:
            self.plot_empty_state()
            return

        # Extract coordinates and colors
        x_coords = []
        y_coords = []
        colors = []

        for node_name, data in self.base_plate_nodes.items():
            x_coords.append(data['x'])
            y_coords.append(data['y'])

            # Color based on status
            if data['design_status'] == 'OK':
                colors.append('#4CAF50')  # Green
            elif data['design_status'] == 'NG':
                colors.append('#f44336')  # Red
            elif data['bolt_size'] is not None:  # Has assignment
                colors.append('#2196F3')  # Blue
            else:
                colors.append('#CCCCCC')  # Gray

        # Plot scatter
        self.scatter = self.ax.scatter(x_coords, y_coords, c=colors, s=150,
                                      edgecolors='black', linewidths=1.5,
                                      picker=True, pickradius=5)

        # Clear previous labels
        self.label_texts = []

        # Add labels based on display option
        display_mode = self.label_display_var.get()

        if display_mode != "none":  # If not "none", display something
            for node_name, data in self.base_plate_nodes.items():
                text_label = ""

                # Build label based on display mode
                if display_mode == "label":
                    text_label = node_name
                elif display_mode == "ratio":
                    if data.get('max_ratio') is not None:
                        text_label = f"{data['max_ratio']:.2f}"
                    else:
                        text_label = "N/A"
                elif display_mode == "both":
                    text_label = node_name
                    if data.get('max_ratio') is not None:
                        text_label += f"\n{data['max_ratio']:.2f}"

                # Use annotate for better font styling (no bbox border)
                if text_label:
                    txt = self.ax.annotate(text_label,
                                           (data['x'], data['y']),
                                           xytext=(5, 5),
                                           textcoords='offset points',
                                           fontsize=8,
                                           fontfamily='Aptos',
                                           alpha=0.85)
                    self.label_texts.append(txt)

        # Styling
        self.ax.set_xlabel('X (m)', fontsize=10, fontweight='bold')
        self.ax.set_ylabel('Y (m)', fontsize=10, fontweight='bold')
        self.ax.set_title('Base Plate Plan View', fontsize=12, fontweight='bold', pad=15)
        self.ax.grid(True, alpha=0.3, linestyle='--')
        self.ax.set_aspect('equal', adjustable='datalim')

        # Auto-adjust limits with margin
        if x_coords and y_coords:
            x_margin = (max(x_coords) - min(x_coords)) * 0.1 or 1
            y_margin = (max(y_coords) - min(y_coords)) * 0.1 or 1
            self.ax.set_xlim(min(x_coords) - x_margin, max(x_coords) + x_margin)
            self.ax.set_ylim(min(y_coords) - y_margin, max(y_coords) + y_margin)

        self.canvas.draw()

    def on_node_click(self, event):
        """Handle node click event"""
        if event.inaxes != self.ax or not self.base_plate_nodes:
            return

        # Find closest node
        min_dist = float('inf')
        closest_node = None

        for node_name, data in self.base_plate_nodes.items():
            dist = math.sqrt((event.xdata - data['x'])**2 + (event.ydata - data['y'])**2)
            if dist < min_dist:
                min_dist = dist
                closest_node = node_name

        # Select if close enough (within 2 units)
        if min_dist < 2.0:
            self.select_node(closest_node)

    def highlight_selected_node(self):
        """Highlight the selected node on plot"""
        if not self.selected_node:
            return

        # Redraw plot with highlight
        self.plot_base_plate_plan()

        # Add highlight circle
        node_data = self.base_plate_nodes[self.selected_node]
        circle = plt.Circle((node_data['x'], node_data['y']), 1.5,
                           color='yellow', fill=False, linewidth=3,
                           linestyle='--', alpha=0.8)
        self.ax.add_patch(circle)

        self.canvas.draw()

    def update_plot_display(self):
        """Update plot when display mode changes (Label, Ratio, or Both)"""
        # Refresh plot
        self.plot_base_plate_plan()
        if self.selected_node:
            self.highlight_selected_node()

        display_mode = self.label_display_var.get()
        self.status_label.config(
            text=f"● Display: {display_mode.title()}",
            fg='#FF9800'
        )
