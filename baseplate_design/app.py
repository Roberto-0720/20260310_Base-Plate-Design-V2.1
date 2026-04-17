"""
baseplate_design/app.py
========================
Final assembly: BasePlateApp inherits from every mixin so that all
self.xxx references resolve correctly — no logic changes anywhere.

Mixin inheritance order (MRO left-to-right):
  UI                  → MainWindowMixin, TabBasePlateMixin,
                         MaterialTabMixin, ColumnBaseTabMixin,
                         WidgetsMixin, DialogsMixin
  Logic               → PlotManagerMixin, NodeManagerMixin,
                         CalculationsMixin, DesignCheckMixin,
                         HingeXlsxMixin
  Data                → MaterialDataMixin, FileManagerMixin,
                         Sap2000Mixin, ExcelExportMixin
"""

import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# ── UI mixins ────────────────────────────────────────────────────────────────
from baseplate_design.ui.main_window_mixin import MainWindowMixin
from baseplate_design.ui.tab_base_plate_mixin import TabBasePlateMixin
from baseplate_design.ui.tab_material_mixin import MaterialTabMixin
from baseplate_design.ui.tab_column_base_mixin import ColumnBaseTabMixin
from baseplate_design.ui.widgets_mixin import WidgetsMixin
from baseplate_design.ui.dialogs_mixin import DialogsMixin

# ── Logic mixins ─────────────────────────────────────────────────────────────
from baseplate_design.logic.plot_manager_mixin import PlotManagerMixin
from baseplate_design.logic.node_manager_mixin import NodeManagerMixin
from baseplate_design.logic.calculations_mixin import CalculationsMixin
from baseplate_design.logic.design_check_mixin import DesignCheckMixin
from baseplate_design.logic.hinge_xlsx_mixin import HingeXlsxMixin

# ── Data mixins ──────────────────────────────────────────────────────────────
from baseplate_design.data.material_data_mixin import MaterialDataMixin
from baseplate_design.data.file_manager_mixin import FileManagerMixin
from baseplate_design.data.sap2000_mixin import Sap2000Mixin
from baseplate_design.data.excel_export_mixin import ExcelExportMixin


class BasePlateApp(
    # UI
    MainWindowMixin,
    TabBasePlateMixin,
    MaterialTabMixin,
    ColumnBaseTabMixin,
    WidgetsMixin,
    DialogsMixin,
    # Logic
    PlotManagerMixin,
    NodeManagerMixin,
    CalculationsMixin,
    DesignCheckMixin,
    HingeXlsxMixin,
    # Data
    MaterialDataMixin,
    FileManagerMixin,
    Sap2000Mixin,
    ExcelExportMixin,
):
    """
    Main application class.

    All business logic lives in the mixin classes; this class only owns
    __init__ (state initialisation + UI bootstrap) and check_license.
    """

    def __init__(self, root):
        self.root = root

        # Check license first
        if not self.check_license():
            return

        self.root.title("🏗️ BASE PLATE DESIGN")
        self.root.geometry("1500x800")
        self.root.configure(bg='#f5f5f5')

        # Initialize current file
        self.current_file = None
        self.reaction_csv_file = None
        self.coordinate_csv_file = None

        # Initialize SAP2000 connection
        self.SapModel = None
        self.is_sap_connected = False
        self.sap_model_file = None
        self.bpl_folder = None  # Path to "Base Plate Design" folder

        # Initialize material data structures
        self.material_trees = {}  # Để lưu reference các tree (sẽ set ở các hàm create sub-tab)
        self.material_status_label = None  # Sẽ tạo ở create_material_define_tab
        self.material_data = {}  # Để lưu dữ liệu material từ Data.xlsx

        # Initialize Base Plate Detail data
        self.base_plate_nodes = {}  # Store node data
        self.selected_node = None   # Currently selected node
        self.fig = None            # Matplotlib figure
        self.ax = None             # Matplotlib axes
        self.canvas = None         # Matplotlib canvas
        self.scatter = None        # Scatter plot object
        self.label_display_var = tk.StringVar(value="both")  # Display: "label", "ratio", or "both"
        self.label_texts = []      # Store text objects for labels

        # Create UI components
        self.create_header()
        self.create_menu_bar()
        self.create_main_content()

    def check_license(self):
        """Check if application license is valid"""
        from datetime import datetime

        expiry_date = datetime(2026, 8, 31)  # Hết tháng 8, 2026
        today = datetime.now()

        if today > expiry_date:
            messagebox.showerror(
                "⛔ License Expired",
                "This application has expired.\n\n"
                "Please contact Roberto to unlock the application.\n\n"
                "Application will now close."
            )
            self.root.quit()
            return False
        return True
