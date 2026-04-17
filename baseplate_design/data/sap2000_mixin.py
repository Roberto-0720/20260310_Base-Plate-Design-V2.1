"""
data/sap2000_mixin.py
======================
SAP2000 integration — connect, disconnect, auto-load coordinates,
reaction data, element joint forces, RC Pier mode, and
load case selection dialog.

App state used: self.SapModel, self.is_sap_connected, self.sap_model_file,
                self.bpl_folder, self.current_file, self.reaction_csv_file,
                self.include_rc_pier_var,
                self.coord_get_model_btn, self.coord_load_auto_btn,
                self.coord_disconnect_btn, self.coordinate_sap_path_label,
                self.coordinate_status_label, self.status_label, self.root
"""

import os
import csv
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import comtypes.client


class Sap2000Mixin:

    # ------------------------------------------------------------------
    # Public entry points
    # ------------------------------------------------------------------

    def get_sap_model_coordinates(self):
        """Connect to SAP2000 and enable auto load for coordinates"""
        if self.is_sap_connected:
            messagebox.showinfo("Info", "Already connected to SAP2000 model:\n" + self.sap_model_file)
            return

        self._connect_to_sap_model("coordinates")

    def get_sap_model_reaction(self):
        """Connect to SAP2000 for reaction data"""
        if self.is_sap_connected:
            messagebox.showinfo("Info", "Already connected to SAP2000 model:\n" + self.sap_model_file)
            return

        self._connect_to_sap_model("reaction")

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def _connect_to_sap_model(self, source_tab):
        """Internal: Connect to SAP2000 model"""
        self.coord_get_model_btn.config(state='disabled')
        self.root.update()

        connection_errors = []

        # Method 1: Try SAP2000v1.Helper (works for v21+)
        try:
            myHelper = comtypes.client.CreateObject("SAP2000v1.Helper")
            SapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject")
            self.SapModel = SapObject.SapModel
            self._connection_success_sap()
            return
        except Exception as e:
            connection_errors.append(f"Method 1 (SAP2000v1.Helper): {str(e)}")

        # Method 2: Try GetActiveObject for older versions (v17-v20)
        try:
            SapObject = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
            self.SapModel = SapObject.SapModel
            self._connection_success_sap()
            return
        except Exception as e:
            connection_errors.append(f"Method 2 (GetActiveObject): {str(e)}")

        # Method 3: Try with version-specific ProgID for SAP2000 v20
        try:
            SapObject = comtypes.client.GetActiveObject("SAP2000.cOAPI")
            self.SapModel = SapObject.SapModel
            self._connection_success_sap()
            return
        except Exception as e:
            connection_errors.append(f"Method 3 (SAP2000.cOAPI): {str(e)}")

        # All methods failed
        self.is_sap_connected = False

        error_details = "\n".join(connection_errors)
        messagebox.showerror("❌ Connection Error",
                            f"Cannot connect to SAP2000.\n"
                            f"Please ensure:\n"
                            f"• SAP2000 is running (v17.2.0+)\n"
                            f"• Model is open\n"
                            f"• Python and SAP2000 are same architecture (64-bit)\n\n"
                            f"Error details:\n{error_details}")

        self.coord_get_model_btn.config(state='normal')
        self.status_label.config(text="● SAP2000 connection failed", fg='#ff4d4d')

    def _connection_success_sap(self):
        """Handle successful SAP2000 connection"""
        try:
            model_file = self.SapModel.GetModelFilename()

            if not model_file:
                messagebox.showwarning("⚠️ Warning", "Please open a SAP2000 model file first!")
                self.coord_get_model_btn.config(state='normal')
                return

            self.is_sap_connected = True
            self.sap_model_file = model_file

            # Create "Base Plate Design" folder in the same directory as SAP model
            sap_dir = os.path.dirname(model_file)
            self.bpl_folder = os.path.join(sap_dir, "Base Plate Design")
            try:
                folder_exists = os.path.exists(self.bpl_folder)

                if not folder_exists:
                    os.makedirs(self.bpl_folder)
                    # Folder is newly created, so create Data.xlsx
                    self.create_or_check_data_file()
                else:
                    # Folder already exists, check if Data.xlsx exists
                    self.create_or_check_data_file()

                # Copy Template.xlsx from script directory if not exists
                self._copy_template_to_bpl_folder()

            except Exception as e:
                messagebox.showwarning("⚠️ Warning", f"Could not create Base Plate Design folder:\n{str(e)}")

            # Update UI
            sap_path_text = f"📁 SAP Model: {os.path.basename(model_file)}"
            self.coordinate_sap_path_label.config(text=sap_path_text)

            self.coord_load_auto_btn.config(state='normal')
            self.coord_get_model_btn.config(state='disabled')
            self.coord_disconnect_btn.config(state='normal')

            messagebox.showinfo("✅ Connected to SAP2000",
                              f"Successfully connected to SAP2000!\n\n"
                              f"Model: {os.path.basename(model_file)}")

            self.status_label.config(text=f"● Connected to SAP2000: {os.path.basename(model_file)}", fg='#90ee90')
            self.reload_all_material_data()

        finally:
            self.root.update()

    def disconnect_sap_model(self):
        """Disconnect from SAP2000 and reset state"""
        self.SapModel = None
        self.is_sap_connected = False
        self.sap_model_file = None
        self.bpl_folder = None

        self.coordinate_sap_path_label.config(text="")
        self.coord_load_auto_btn.config(state='disabled')
        self.coord_get_model_btn.config(state='normal')
        self.coord_disconnect_btn.config(state='disabled')

        self.status_label.config(text="● Disconnected from SAP2000", fg='#90ee90')
        messagebox.showinfo("Disconnected", "Disconnected from SAP2000 model")

    def disconnect_sap_model_coord(self):
        """Wrapper for disconnect from Column Base tab"""
        self.disconnect_sap_model()

    # ------------------------------------------------------------------
    # Auto Load
    # ------------------------------------------------------------------

    def load_coordinates_auto(self):
        """Load column base coordinates and reaction data from selected beam elements"""
        if not self.is_sap_connected or not self.SapModel:
            messagebox.showerror("Error", "Not connected to SAP2000 model!")
            return

        if not self.current_file:
            messagebox.showwarning("Warning", "Please save your project first!")
            return

        self.load_coordinates_and_reaction_auto()

    def load_coordinates_and_reaction_auto(self):
        """Load column base coordinates from SELECTED BEAM ELEMENTS and create reaction data files from start point nodes"""
        if not self.is_sap_connected or not self.SapModel:
            messagebox.showerror("Error", "Not connected to SAP2000 model!")
            return

        if not self.current_file:
            messagebox.showwarning("Warning", "Please save your project first!")
            return

        try:
            project_dir = os.path.dirname(self.current_file)

            # Get SELECTED FRAME ELEMENTS
            selected_frame_elements = []
            selected_joints = []

            try:
                ret = self.SapModel.SelectObj.GetSelected()
                num_sel = ret[0]

                if num_sel > 0:
                    obj_types = ret[1]
                    obj_names = ret[2]

                    for i in range(num_sel):
                        obj_type = obj_types[i]
                        obj_name = obj_names[i]

                        print(f"Debug: Selected object: {obj_name}, Type: {obj_type}")

                        try:
                            ret_prop = self.SapModel.FrameObj.GetPoints(obj_name)
                            print(f"Debug: GetPoints result for {obj_name}: {ret_prop}")
                            if ret_prop[2] == 0:
                                start_node = ret_prop[0]

                                section_name = "Unknown"
                                try:
                                    ret_section = self.SapModel.FrameObj.GetSection(obj_name)
                                    print(f"Debug: GetSection result for {obj_name}: {ret_section}")
                                    if len(ret_section) >= 3 and ret_section[2] == 0:
                                        section_name = ret_section[0] if ret_section[0] else "Unknown"
                                        print(f"Debug: Got section name from GetSection: {section_name}")
                                    elif len(ret_section) >= 1 and ret_section[0]:
                                        section_name = ret_section[0]
                                        print(f"Debug: Using section name from index 0: {section_name}")
                                except Exception as e:
                                    print(f"Debug: GetSection failed: {e}")
                                    section_name = "Unknown"

                                print(f"Debug: Final section name for {obj_name}: {section_name}")

                                selected_joints.append(start_node)
                                print(f"Debug: Added frame element {obj_name} with start node {start_node}, section {section_name}")
                            else:
                                print(f"Debug: GetPoints failed for {obj_name}, return code: {ret_prop[2]}")
                        except Exception as e:
                            print(f"Debug: {obj_name} error: {e}")
                            continue
            except Exception as e:
                print(f"SelectObj.GetSelected failed: {e}")

            # Cấu trúc đúng — toàn bộ nằm trong for loop:
            for i in range(num_sel):
                obj_type = obj_types[i]
                obj_name = obj_names[i]

                try:
                    ret_prop = self.SapModel.FrameObj.GetPoints(obj_name)
                    if ret_prop[2] == 0:
                        start_node = ret_prop[0]

                        # Get section name
                        section_name = "Unknown"
                        try:
                            ret_section = self.SapModel.FrameObj.GetSection(obj_name)
                            if len(ret_section) >= 3 and ret_section[2] == 0:
                                section_name = ret_section[0] if ret_section[0] else "Unknown"
                            elif len(ret_section) >= 1 and ret_section[0]:
                                section_name = ret_section[0]
                        except Exception as e:
                            print(f"Debug: GetSection failed: {e}")
                            section_name = "Unknown"

                        # ✅ Get beta — TRONG loop, SAU section_name, TRƯỚC append
                        beta = 0
                        try:
                            ret_axes = self.SapModel.FrameObj.GetLocalAxes(obj_name)
                            if ret_axes[2] == 0:
                                raw_angle = ret_axes[0]
                                beta = 0 if abs(raw_angle % 180) < 45 else 90
                                print(f"Debug: {obj_name} beta = {beta} (raw={raw_angle})")
                        except Exception as e:
                            print(f"Debug: GetLocalAxes failed for {obj_name}: {e}")
                            beta = 0

                        # ✅ Append với beta — TRONG loop
                        selected_frame_elements.append((obj_name, start_node, section_name, beta))
                        selected_joints.append(start_node)
                        print(f"Debug: Added {obj_name}, node={start_node}, section={section_name}, beta={beta}")
                    else:
                        print(f"Debug: GetPoints failed for {obj_name}")
                except Exception as e:
                    print(f"Debug: {obj_name} error: {e}")
                    continue

            if not selected_frame_elements:
                messagebox.showwarning("⚠️ No Beam Elements Selected",
                    "No beam/column elements detected in selection!\n\n"
                    "Please:\n"
                    "1. Select beam/column elements in SAP2000\n"
                    "2. Do NOT select nodes directly\n"
                    "3. Ensure selection is visible (highlighted)\n"
                    "4. Click 'Load (Auto)' again")
                return

            # Branch for Include RC Pier? = Yes
            if self.include_rc_pier_var.get() == "Yes":
                self._load_with_rc_pier(selected_frame_elements, project_dir)
                return

            selected_joints = list(dict.fromkeys(selected_joints))

            print(f"Debug: Found {len(selected_frame_elements)} selected frame elements")
            print(f"Debug: Extracted {len(selected_joints)} unique start point nodes")

            # Get coordinates for SELECTED joints
            node_to_section = {}
            node_to_beta = {}
            for elem_name, start_node, section_name, beta in selected_frame_elements:
                if start_node not in node_to_section:
                    node_to_section[start_node] = section_name
                    node_to_beta[start_node] = beta

            base_coords = []
            for joint_name in selected_joints:
                ret = self.SapModel.PointObj.GetCoordCartesian(joint_name)
                x, y, z = ret[0], ret[1], ret[2]
                section_name = node_to_section.get(joint_name, "Unknown")
                beta = node_to_beta.get(joint_name, 0)
                base_coords.append([joint_name, str(x), str(y), str(z), section_name, str(beta)])

            # Sort by name
            def sort_key(base_name):
                import re
                parts = re.split(r'(\d+)', base_name[0])
                result = []
                for i, part in enumerate(parts):
                    if i % 2 == 0:
                        result.append((0, part))
                    else:
                        result.append((1, int(part)))
                return result

            base_coords.sort(key=sort_key)

            # Create bpl_coordinate.csv in Base Plate Design folder
            if not self.bpl_folder:
                self.bpl_folder = os.path.join(project_dir, "Base Plate Design")
                if not os.path.exists(self.bpl_folder):
                    os.makedirs(self.bpl_folder)

            coord_file = os.path.join(self.bpl_folder, "bpl_coordinate.csv")
            with open(coord_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Column Base', 'X (m)', 'Y (m)', 'Z (m)', 'Section', 'Beta'])
                for coord in base_coords:
                    writer.writerow(coord)

            messagebox.showinfo("✅ Success",
                f"Column base coordinates loaded from SAP2000!\n\n"
                f"Selected joints: {len(base_coords)}")

            self.coordinate_status_label.config(
                text=f"✓ Auto-loaded: {len(base_coords)} SELECTED joints",
                fg='#4CAF50'
            )
            self.status_label.config(text=f"● {len(base_coords)} selected column bases loaded", fg='#90ee90')

            # Load reaction data for same selected joints
            self._load_reaction_data_for_selected_joints(selected_joints, project_dir)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load column base coordinates:\n{str(e)}")
            self.status_label.config(text="● Auto-load failed", fg='#ff4d4d')

    def _load_reaction_data_for_selected_joints(self, selected_joints, project_dir):
        """Load reaction data for the selected joints"""
        try:
            # Show Load Case Selection Dialog
            selected_lcs = self.show_load_case_selection_dialog()
            if not selected_lcs:
                self.status_label.config(text="● Load case selection cancelled", fg='#ff9800')
                return

            # Confirm re-run analysis
            confirm = messagebox.askyesno("Confirm Analysis",
                f"To retrieve results for the {len(selected_lcs)} selected cases/combos,\n"
                "the output selection needs to be updated and the analysis re-run in SAP2000.\n\n"
                "This may take some time depending on the model size.\n"
                "Continue?")
            if not confirm:
                self.status_label.config(text="● Cancelled by user", fg='#ff9800')
                return

            # Create Progress Dialog
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Loading Reaction Data")
            progress_window.geometry("500x150")
            progress_window.transient(self.root)
            progress_window.grab_set()
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
            y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
            progress_window.geometry(f"+{x}+{y}")
            tk.Label(progress_window, text="Extracting Reaction Data from SAP2000...", font=('Arial', 12, 'bold')).pack(pady=(20, 10))
            progress_label = tk.Label(progress_window, text="Initializing...", font=('Arial', 10))
            progress_label.pack(pady=5)
            progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=400, mode='determinate')
            progress_bar.pack(pady=10)
            progress_percent = tk.Label(progress_window, text="0%", font=('Arial', 10, 'bold'), fg='#2196F3')
            progress_percent.pack(pady=5)

            def update_progress(value, label_text):
                progress_bar['value'] = value
                progress_label.config(text=label_text)
                progress_percent.config(text=f"{int(value)}%")
                progress_window.update()

            update_progress(5, "Getting load case/combo lists...")

            ret_cases = self.SapModel.LoadCases.GetNameList()
            num_cases = ret_cases[0]
            case_names = ret_cases[1] if num_cases > 0 else []

            ret_combos = self.SapModel.RespCombo.GetNameList()
            num_combos = ret_combos[0]
            combo_names = ret_combos[1] if num_combos > 0 else []

            ret_cases = self.SapModel.LoadCases.GetNameList()
            case_set = set(ret_cases[1]) if ret_cases[0] > 0 else set()
            ret_combos = self.SapModel.RespCombo.GetNameList()
            combo_set = set(ret_combos[1]) if ret_combos[0] > 0 else set()

            update_progress(10, "Deselecting all cases/combos...")
            self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()

            update_progress(15, "Setting selected cases/combos...")
            for lc_name in selected_lcs:
                if lc_name in case_set:
                    self.SapModel.Results.Setup.SetCaseSelectedForOutput(lc_name)
                elif lc_name in combo_set:
                    self.SapModel.Results.Setup.SetComboSelectedForOutput(lc_name)

            update_progress(20, "Running analysis (may take time)...")
            ret_analyze = self.SapModel.Analyze.RunAnalysis()
            if ret_analyze != 0:
                progress_window.after(500, progress_window.destroy)
                messagebox.showerror("Error", "Analysis failed in SAP2000! Check model.")
                self.status_label.config(text="● Analysis failed", fg='#ff4d4d')
                return

            update_progress(40, "Extracting reaction data...")

            reaction_data = []
            reaction_data.append(['TABLE: Joint Reactions'])
            reaction_data.append(['Joint', 'LoadCase', 'CaseType', 'F1', 'F2', 'F3', 'M1', 'M2', 'M3'])
            reaction_data.append(['Text', 'Text', 'Text', 'KN', 'KN', 'KN', 'KN-m', 'KN-m', 'KN-m'])

            for node in selected_joints:
                try:
                    ret = self.SapModel.Results.JointReact(node, 0)
                    num_results = ret[0]
                    if num_results > 0:
                        load_cases = ret[3]
                        f1_vals = ret[6]
                        f2_vals = ret[7]
                        f3_vals = ret[8]
                        m1_vals = ret[9]
                        m2_vals = ret[10]
                        m3_vals = ret[11]

                        for i in range(num_results):
                            lc_name = load_cases[i].strip()
                            if lc_name in selected_lcs:
                                case_type_str = "Combination"
                                reaction_data.append([
                                    node,
                                    lc_name,
                                    case_type_str,
                                    f"{f1_vals[i]:.6g}",
                                    f"{f2_vals[i]:.6g}",
                                    f"{f3_vals[i]:.6g}",
                                    f"{m1_vals[i]:.6g}",
                                    f"{m2_vals[i]:.6g}",
                                    f"{m3_vals[i]:.6g}"
                                ])
                                print(f"Debug: Got reaction for {node}, {lc_name}")
                except Exception as e:
                    print(f"Error for joint {node}: {e}")

            if len(reaction_data) <= 3:
                progress_window.after(500, progress_window.destroy)
                messagebox.showwarning("No Data",
                    f"No reaction data extracted!\n\n"
                    f"Check:\n- Model analyzed successfully\n- Selected cases have results")
                return

            update_progress(70, "Writing CSV file...")

            if not self.bpl_folder:
                self.bpl_folder = os.path.join(project_dir, "Base Plate Design")
                if not os.path.exists(self.bpl_folder):
                    os.makedirs(self.bpl_folder)

            reaction_csv_file = os.path.join(self.bpl_folder, "reaction_data_sap2000.csv")
            with open(reaction_csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for row in reaction_data:
                    writer.writerow(row)

            self.reaction_csv_file = reaction_csv_file

            update_progress(100, "Done!")
            progress_window.after(500, progress_window.destroy)

            messagebox.showinfo("✅ Reaction Data Loaded",
                f"Reaction data extracted successfully!\n\n"
                f"Nodes: {len(selected_joints)}\n"
                f"Load Cases/Combos: {len(selected_lcs)}\n"
                f"Data points: {len(reaction_data)}")

            try:
                if platform.system() == 'Windows':
                    os.startfile(reaction_csv_file)
                elif platform.system() == 'Darwin':
                    subprocess.call(['open', reaction_csv_file])
                else:
                    subprocess.call(['xdg-open', reaction_csv_file])
            except:
                pass

            self.status_label.config(text=f"● Reaction data loaded: {len(reaction_data)} data points", fg='#90ee90')

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load reaction data:\n{str(e)}")
            self.status_label.config(text="● Reaction load failed", fg='#ff4d4d')

    def _load_with_rc_pier(self, selected_frame_elements, project_dir):
        """Load coordinates and element joint forces when Include RC Pier = Yes.

        Algorithm:
        1. Get Start/End node coordinates for all selected elements
        2. Pair elements by shared node (Base Plate Node)
        3. Distinguish: element with lower-Z other node = RC Pier, higher-Z = Steel Column
        4. Write bpl_coordinate.csv using shared node coords + Steel Column section
        5. Extract Element Joint Forces at End node of each RC Pier
        """
        try:
            import re

            # Step 1: Get full node info for all selected elements
            element_info = []  # [(elem_name, start_node, end_node, start_z, end_z, section_name)]

            for elem_name, start_node, section_name, beta in selected_frame_elements:
                try:
                    ret_points = self.SapModel.FrameObj.GetPoints(elem_name)
                    if ret_points[2] == 0:
                        s_node = ret_points[0]
                        e_node = ret_points[1]

                        ret_s = self.SapModel.PointObj.GetCoordCartesian(s_node)
                        s_x, s_y, s_z = ret_s[0], ret_s[1], ret_s[2]

                        ret_e = self.SapModel.PointObj.GetCoordCartesian(e_node)
                        e_x, e_y, e_z = ret_e[0], ret_e[1], ret_e[2]

                        ret_axes = self.SapModel.FrameObj.GetLocalAxes(elem_name)
                        raw_angle = ret_axes[0] if ret_axes[2] == 0 else 0
                        beta = 0 if abs(raw_angle % 180) < 45 else 90

                        element_info.append({
                            'name': elem_name,
                            'start_node': s_node,
                            'end_node': e_node,
                            'start_xyz': (s_x, s_y, s_z),
                            'end_xyz': (e_x, e_y, e_z),
                            'section': section_name,
                            'beta': beta
                        })
                        print(f"Debug RC Pier: Element {elem_name}: Start={s_node}(z={s_z:.3f}), End={e_node}(z={e_z:.3f}), Section={section_name}")
                except Exception as e:
                    print(f"Debug RC Pier: Error getting info for {elem_name}: {e}")
                    continue

            if len(element_info) < 2:
                messagebox.showwarning("⚠️ Insufficient Elements",
                    "Need at least 2 frame elements (RC Pier + Steel Column pair)!\n\n"
                    "Please select pairs of elements:\n"
                    "- RC Pier (below base plate)\n"
                    "- Steel Column (above base plate)")
                return

            # Step 2: Build node-to-elements map for pairing
            node_to_elements = {}  # node_name -> [element_info, ...]
            for elem in element_info:
                for node in [elem['start_node'], elem['end_node']]:
                    if node not in node_to_elements:
                        node_to_elements[node] = []
                    node_to_elements[node].append(elem)

            # Step 3: Find pairs sharing a common node (Base Plate Node)
            pairs = []  # [(pier_elem, column_elem, shared_node)]
            used_elements = set()

            for node, elems in node_to_elements.items():
                if len(elems) == 2 and elems[0]['name'] != elems[1]['name']:
                    e1, e2 = elems[0], elems[1]

                    if e1['name'] in used_elements or e2['name'] in used_elements:
                        continue

                    # Find the "other" node (non-shared) Z for each element
                    def get_other_z(elem, shared_node):
                        if elem['start_node'] == shared_node:
                            return elem['end_xyz'][2]  # Z of end node
                        else:
                            return elem['start_xyz'][2]  # Z of start node

                    z1_other = get_other_z(e1, node)
                    z2_other = get_other_z(e2, node)

                    # Element with lower-Z other node = RC Pier
                    if z1_other < z2_other:
                        pier_elem, column_elem = e1, e2
                    else:
                        pier_elem, column_elem = e2, e1

                    pairs.append((pier_elem, column_elem, node))
                    used_elements.add(e1['name'])
                    used_elements.add(e2['name'])

                    print(f"Debug RC Pier: Pair found at node {node}: "
                          f"Pier={pier_elem['name']}(section={pier_elem['section']}), "
                          f"Column={column_elem['name']}(section={column_elem['section']})")

            if not pairs:
                messagebox.showwarning("⚠️ No Pairs Found",
                    "Could not find RC Pier + Steel Column pairs!\n\n"
                    "Make sure:\n"
                    "1. Each pair shares a common node\n"
                    "2. RC Pier is below, Steel Column is above\n"
                    "3. Select both elements of each pair\n"
                    "4. Don't select node, select element only")
                return

            print(f"Debug RC Pier: Found {len(pairs)} pier-column pairs")

            # Step 4: Create bpl_coordinate.csv using shared node coords + Steel Column section
            base_coords = []
            pier_elements_for_force = []  # [(pier_name, shared_node)]

            for pier_elem, column_elem, shared_node in pairs:
                # Get shared node coordinates
                ret_coord = self.SapModel.PointObj.GetCoordCartesian(shared_node)
                x, y, z = ret_coord[0], ret_coord[1], ret_coord[2]

                # Section name from Steel Column
                section_name = column_elem['section']
                beta = column_elem.get('beta', 0)

                base_coords.append([shared_node, str(x), str(y), str(z), section_name, str(beta)])
                pier_elements_for_force.append((pier_elem['name'], shared_node))

                print(f"Debug RC Pier: Base plate at node {shared_node}: "
                      f"({x:.3f}, {y:.3f}, {z:.3f}), Section={section_name}")

            # Sort by name
            def sort_key(base_name):
                parts = re.split(r'(\d+)', base_name[0])
                result = []
                for i, part in enumerate(parts):
                    if i % 2 == 0:
                        result.append((0, part))
                    else:
                        result.append((1, int(part)))
                return result

            base_coords.sort(key=sort_key)

            # Write bpl_coordinate.csv
            if not self.bpl_folder:
                self.bpl_folder = os.path.join(project_dir, "Base Plate Design")
                if not os.path.exists(self.bpl_folder):
                    os.makedirs(self.bpl_folder)

            coord_file = os.path.join(self.bpl_folder, "bpl_coordinate.csv")
            with open(coord_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Column Base', 'X (m)', 'Y (m)', 'Z (m)', 'Section', 'Beta'])
                for coord in base_coords:
                    writer.writerow(coord)

            messagebox.showinfo("✅ Success",
                f"Column base coordinates loaded (RC Pier mode)!\n\n"
                f"Pairs found: {len(pairs)}\n"
                f"File: bpl_coordinate.csv")

            self.coordinate_status_label.config(
                text=f"✓ RC Pier mode: {len(pairs)} pairs loaded",
                fg='#4CAF50'
            )
            self.status_label.config(text=f"● {len(pairs)} pier-column pairs loaded", fg='#90ee90')

            # Step 5: Load Element Joint Forces for RC Pier end nodes
            self._load_elejointforce_for_piers(pier_elements_for_force, project_dir)

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to load with RC Pier:\n{str(e)}")
            self.status_label.config(text="● RC Pier load failed", fg='#ff4d4d')

    def _load_elejointforce_for_piers(self, pier_elements, project_dir):
        """Load Element Joint Forces at End node of each RC Pier and write to elejointforce.csv.

        Args:
            pier_elements: list of (pier_name, shared_node) tuples
            project_dir: project directory path
        """
        try:
            # Show Load Case Selection Dialog
            selected_lcs = self.show_load_case_selection_dialog()
            if not selected_lcs:
                self.status_label.config(text="● Load case selection cancelled", fg='#ff9800')
                return

            # Confirm re-run analysis
            confirm = messagebox.askyesno("Confirm Analysis",
                f"To retrieve Element Joint Forces for {len(pier_elements)} RC Pier elements\n"
                f"and {len(selected_lcs)} selected cases/combos,\n"
                "the analysis needs to be re-run in SAP2000.\n\n"
                "Continue?")
            if not confirm:
                self.status_label.config(text="● Cancelled by user", fg='#ff9800')
                return

            # Create Progress Dialog
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Loading Element Joint Forces")
            progress_window.geometry("500x150")
            progress_window.transient(self.root)
            progress_window.grab_set()
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
            y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
            progress_window.geometry(f"+{x}+{y}")
            tk.Label(progress_window, text="Extracting Element Joint Forces from SAP2000...", font=('Arial', 12, 'bold')).pack(pady=(20, 10))
            progress_label = tk.Label(progress_window, text="Initializing...", font=('Arial', 10))
            progress_label.pack(pady=5)
            progress_bar = ttk.Progressbar(progress_window, orient='horizontal', length=400, mode='determinate')
            progress_bar.pack(pady=10)
            progress_percent = tk.Label(progress_window, text="0%", font=('Arial', 10, 'bold'), fg='#2196F3')
            progress_percent.pack(pady=5)

            def update_progress(value, label_text):
                progress_bar['value'] = value
                progress_label.config(text=label_text)
                progress_percent.config(text=f"{int(value)}%")
                progress_window.update()

            update_progress(5, "Getting load case/combo lists...")

            ret_cases = self.SapModel.LoadCases.GetNameList()
            case_set = set(ret_cases[1]) if ret_cases[0] > 0 else set()
            ret_combos = self.SapModel.RespCombo.GetNameList()
            combo_set = set(ret_combos[1]) if ret_combos[0] > 0 else set()

            update_progress(10, "Deselecting all cases/combos...")
            self.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()

            update_progress(15, "Setting selected cases/combos...")
            for lc_name in selected_lcs:
                if lc_name in case_set:
                    self.SapModel.Results.Setup.SetCaseSelectedForOutput(lc_name)
                elif lc_name in combo_set:
                    self.SapModel.Results.Setup.SetComboSelectedForOutput(lc_name)

            update_progress(20, "Running analysis (may take time)...")
            ret_analyze = self.SapModel.Analyze.RunAnalysis()
            if ret_analyze != 0:
                progress_window.after(500, progress_window.destroy)
                messagebox.showerror("Error", "Analysis failed in SAP2000! Check model.")
                self.status_label.config(text="● Analysis failed", fg='#ff4d4d')
                return

            update_progress(40, "Extracting element joint forces...")

            # Build CSV data
            force_data = []
            force_data.append(['TABLE: Element Joint Forces - Pier End'])
            force_data.append(['Joint', 'PierElement', 'LoadCase', 'CaseType', 'F1', 'F2', 'F3', 'M1', 'M2', 'M3'])
            force_data.append(['Text', 'Text', 'Text', 'Text', 'KN', 'KN', 'KN', 'KN-m', 'KN-m', 'KN-m'])

            total_piers = len(pier_elements)
            for idx, (pier_name, shared_node) in enumerate(pier_elements):
                progress_pct = 40 + (idx / total_piers) * 50
                update_progress(progress_pct, f"Processing pier {pier_name} ({idx+1}/{total_piers})...")

                try:
                    # FrameJointForce returns forces at element joints (2 per element: start & end)
                    # Parameters: (Name, ItemTypeElm) where 0 = ObjectElm
                    ret = self.SapModel.Results.FrameJointForce(pier_name, 0)
                    num_results = ret[0]

                    print(f"Debug RC Pier: FrameJointForce for {pier_name}: {num_results} results")

                    if num_results > 0:
                        obj_names = ret[1]   # Object names
                        elm_names = ret[2]   # Element names
                        point_elms = ret[3]  # Point element names (joint names)
                        load_cases = ret[4]  # Load case names
                        step_types = ret[5]  # Step types
                        step_nums = ret[6]   # Step numbers
                        f1_vals = ret[7]     # F1 values
                        f2_vals = ret[8]     # F2 values
                        f3_vals = ret[9]     # F3 values
                        m1_vals = ret[10]    # M1 values
                        m2_vals = ret[11]    # M2 values
                        m3_vals = ret[12]    # M3 values

                        for i in range(num_results):
                            lc_name = load_cases[i].strip()
                            joint_name = point_elms[i].strip() if point_elms[i] else ""

                            # Only keep results at END node (shared node = base plate node)
                            if joint_name == shared_node and lc_name in selected_lcs:
                                case_type_str = "Combination" if lc_name in combo_set else "LinStatic"
                                force_data.append([
                                    shared_node,
                                    pier_name,
                                    lc_name,
                                    case_type_str,
                                    f"{f1_vals[i]:.6g}",
                                    f"{f2_vals[i]:.6g}",
                                    f"{f3_vals[i]:.6g}",
                                    f"{m1_vals[i]:.6g}",
                                    f"{m2_vals[i]:.6g}",
                                    f"{m3_vals[i]:.6g}"
                                ])
                                print(f"Debug RC Pier: Got force for {pier_name} at {shared_node}, LC={lc_name}")

                except Exception as e:
                    print(f"Error for pier {pier_name}: {e}")
                    import traceback
                    traceback.print_exc()

            if len(force_data) <= 3:
                progress_window.after(500, progress_window.destroy)
                messagebox.showwarning("No Data",
                    f"No element joint force data extracted!\n\n"
                    f"Check:\n- Model analyzed successfully\n- Selected cases have results\n"
                    f"- RC Pier elements are correctly identified")
                return

            update_progress(95, "Writing CSV file...")

            if not self.bpl_folder:
                self.bpl_folder = os.path.join(project_dir, "Base Plate Design")
                if not os.path.exists(self.bpl_folder):
                    os.makedirs(self.bpl_folder)

            force_csv_file = os.path.join(self.bpl_folder, "elejointforce.csv")
            with open(force_csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for row in force_data:
                    writer.writerow(row)

            # Map elejointforce → reaction_data_sap2000.csv (drop PierElement, negate forces)
            reaction_data = []
            reaction_data.append(['TABLE: Joint Reactions'])
            reaction_data.append(['Joint', 'LoadCase', 'CaseType', 'F1', 'F2', 'F3', 'M1', 'M2', 'M3'])
            reaction_data.append(['Text', 'Text', 'Text', 'KN', 'KN', 'KN', 'KN-m', 'KN-m', 'KN-m'])

            for row in force_data[3:]:  # Skip 3 header rows
                # row = [Joint, PierElement, LoadCase, CaseType, F1, F2, F3, M1, M2, M3]
                joint = row[0]
                load_case = row[2]   # Col C → Col B
                case_type = row[3]   # Col D → Col C
                # Negate forces: Col E-J → Col D-I
                f1 = f"{-float(row[4]):.6g}"
                f2 = f"{-float(row[5]):.6g}"
                f3 = f"{-float(row[6]):.6g}"
                m1 = f"{-float(row[7]):.6g}"
                m2 = f"{-float(row[8]):.6g}"
                m3 = f"{-float(row[9]):.6g}"
                reaction_data.append([joint, load_case, case_type, f1, f2, f3, m1, m2, m3])

            reaction_csv_file = os.path.join(self.bpl_folder, "reaction_data_sap2000.csv")
            with open(reaction_csv_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                for row in reaction_data:
                    writer.writerow(row)

            self.reaction_csv_file = reaction_csv_file

            update_progress(100, "Done!")
            progress_window.after(500, progress_window.destroy)

            data_point_count = len(force_data) - 3  # Exclude header rows
            messagebox.showinfo("✅ Element Joint Forces Loaded",
                f"Element Joint Forces extracted successfully!\n\n"
                f"RC Pier elements: {len(pier_elements)}\n"
                f"Load Cases/Combos: {len(selected_lcs)}\n"
                f"Data points: {data_point_count}")

            try:
                if platform.system() == 'Windows':
                    os.startfile(reaction_csv_file)
                elif platform.system() == 'Darwin':
                    subprocess.call(['open', reaction_csv_file])
                else:
                    subprocess.call(['xdg-open', reaction_csv_file])
            except:
                pass

            self.status_label.config(text=f"● Reaction data loaded (RC Pier): {data_point_count} data points", fg='#90ee90')

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to load element joint forces:\n{str(e)}")
            self.status_label.config(text="● Element joint force load failed", fg='#ff4d4d')

    # ------------------------------------------------------------------
    # Load Case Selection Dialog
    # ------------------------------------------------------------------

    def show_load_case_selection_dialog(self):
        """Show dialog to select Load Cases/Combinations from SAP2000 model"""
        if not self.SapModel:
            return None
        try:
            ret_cases = self.SapModel.LoadCases.GetNameList()
            num_cases = ret_cases[0]
            case_names = ret_cases[1] if num_cases > 0 else []

            ret_combos = self.SapModel.RespCombo.GetNameList()
            num_combos = ret_combos[0]
            combo_names = ret_combos[1] if num_combos > 0 else []

            if num_cases == 0 and num_combos == 0:
                messagebox.showwarning("No Load Cases", "No Load Cases or Combinations found in model!")
                return None

            import re
            def sort_key(name):
                num_part = re.search(r'\d+', name)
                return (int(num_part.group()) if num_part else float('inf'), name)

            case_names = sorted(case_names, key=sort_key)
            combo_names = sorted(combo_names, key=sort_key)

            dialog = tk.Toplevel(self.root)
            dialog.title("Select Load Cases/Combinations")
            dialog.geometry("900x550")
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
            y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
            dialog.geometry(f"+{x}+{y}")

            tk.Label(dialog, text="Select Load Cases/Combinations for Analysis", font=('Arial', 12, 'bold'), fg='#1a472a').pack(pady=(10, 3))

            case_vars = {}
            combo_vars = {}

            # Top: Deselect All
            btn_top_frame = tk.Frame(dialog)
            btn_top_frame.pack(pady=3)
            def deselect_all():
                for var in case_vars.values():
                    var.set(False)
                for var in combo_vars.values():
                    var.set(False)
            tk.Button(btn_top_frame, text="✗ Deselect All", command=deselect_all, bg='#f44336', fg='white', font=('Arial', 9, 'bold'), cursor='hand2').pack()

            # Helper functions
            def get_prefix_digit(name):
                num_part = re.search(r'\d+', name)
                if num_part and num_part.group():
                    return num_part.group()[0]
                return None

            def parse_prefix_input(prefix_input):
                if not prefix_input.strip():
                    return set()
                ranges = []
                for part in prefix_input.strip().split(','):
                    part = part.strip()
                    if '-' in part:
                        try:
                            start, end = map(int, part.split('-'))
                            ranges.extend(range(start, end + 1))
                        except:
                            pass
                    else:
                        try:
                            ranges.append(int(part))
                        except:
                            pass
                return set(str(d) for d in ranges)

            selection_state = {'last_clicked_index': None, 'last_type': None}

            # ==================== TWO-COLUMN LAYOUT ====================
            columns_frame = tk.Frame(dialog)
            columns_frame.pack(fill='both', expand=True, padx=10, pady=5)
            columns_frame.columnconfigure(0, weight=3)
            columns_frame.columnconfigure(1, weight=2)

            def create_lc_column(parent, col_idx, title, title_bg, title_fg, names, var_dict, item_type):
                col_frame = tk.Frame(parent, bg='white', bd=1, relief='groove')
                col_frame.grid(row=0, column=col_idx, sticky='nsew', padx=(0 if col_idx == 0 else 5, 0))

                header = tk.Frame(col_frame, bg=title_bg)
                header.pack(fill='x')
                tk.Label(header, text=title, font=('Arial', 10, 'bold'),
                        bg=title_bg, fg=title_fg).pack(side='left', padx=5, pady=4)

                ctrl = tk.Frame(col_frame, bg='#f5f5f5')
                ctrl.pack(fill='x', padx=3, pady=3)

                tk.Label(ctrl, text="Prefix:", font=('Arial', 8), bg='#f5f5f5').pack(side='left', padx=2)
                prefix_entry = tk.Entry(ctrl, width=8, font=('Arial', 9))
                prefix_entry.pack(side='left', padx=2)

                def apply_prefix():
                    prefix_set = parse_prefix_input(prefix_entry.get())
                    if prefix_set:
                        for name, var in var_dict.items():
                            if get_prefix_digit(name) in prefix_set:
                                var.set(True)
                def select_all():
                    for var in var_dict.values():
                        var.set(True)
                def select_none():
                    for var in var_dict.values():
                        var.set(False)

                tk.Button(ctrl, text="Apply", command=apply_prefix, bg='#2196F3', fg='white',
                         font=('Arial', 8, 'bold'), cursor='hand2').pack(side='left', padx=2)
                tk.Button(ctrl, text="✓ All", command=select_all, bg='#4CAF50', fg='white',
                         font=('Arial', 8, 'bold'), cursor='hand2').pack(side='left', padx=2)
                tk.Button(ctrl, text="✗ None", command=select_none, bg='#FF9800', fg='white',
                         font=('Arial', 8, 'bold'), cursor='hand2').pack(side='left', padx=2)
                tk.Label(ctrl, text=f"({len(names)})", font=('Arial', 8), bg='#f5f5f5', fg='#999').pack(side='left', padx=3)

                scroll_container = tk.Frame(col_frame, bg='white')
                scroll_container.pack(fill='both', expand=True, padx=2, pady=2)

                canvas = tk.Canvas(scroll_container, bg='white', highlightthickness=0)
                scrollbar = ttk.Scrollbar(scroll_container, orient='vertical', command=canvas.yview)
                scrollable = tk.Frame(canvas, bg='white')

                scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
                canvas.create_window((0, 0), window=scrollable, anchor='nw')
                canvas.configure(yscrollcommand=scrollbar.set)

                def _on_mw(event):
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                def _bind_mw(event):
                    canvas.bind_all("<MouseWheel>", _on_mw)
                def _unbind_mw(event):
                    canvas.unbind_all("<MouseWheel>")
                canvas.bind("<Enter>", _bind_mw)
                canvas.bind("<Leave>", _unbind_mw)

                canvas.pack(side='left', fill='both', expand=True)
                scrollbar.pack(side='right', fill='y')

                for idx, name in enumerate(names):
                    var = tk.BooleanVar(value=False)
                    var_dict[name] = var
                    cb = tk.Checkbutton(scrollable, text=f" {name}", variable=var,
                                       font=('Arial', 10), bg='white', anchor='w')

                    def on_click(event, index=idx, item_name=name, itype=item_type):
                        if event.state & 0x1:
                            last_idx = selection_state['last_clicked_index']
                            if last_idx is not None and selection_state['last_type'] == itype:
                                start, end = min(last_idx, index), max(last_idx, index)
                                for i in range(start, end + 1):
                                    var_dict[names[i]].set(True)
                            else:
                                var_dict[item_name].set(True)
                        else:
                            var_dict[item_name].set(not var_dict[item_name].get())
                        selection_state['last_clicked_index'] = index
                        selection_state['last_type'] = itype
                        return 'break'

                    cb.pack(fill='x', padx=5, pady=1)
                    cb.bind('<Button-1>', on_click)

                if not names:
                    tk.Label(scrollable, text="(None)", font=('Arial', 9, 'italic'),
                            bg='white', fg='#ccc').pack(pady=20)

            # LEFT: Load Combinations (more commonly used)
            create_lc_column(columns_frame, 0,
                           "🔗 Load Combinations", '#FFF3E0', '#E65100',
                           combo_names, combo_vars, 'combo')

            # RIGHT: Load Cases
            create_lc_column(columns_frame, 1,
                           "📋 Load Cases", '#E8F5E9', '#2E7D32',
                           case_names, case_vars, 'case')

            tk.Label(dialog, text=f"Total: {num_cases} Load Cases, {num_combos} Combinations | Shift+Click for range",
                    font=('Arial', 9, 'italic'), fg='#666666').pack(pady=3)

            result = {'selected': None}
            def on_ok():
                selected = []
                for name, var in case_vars.items():
                    if var.get():
                        selected.append(name)
                for name, var in combo_vars.items():
                    if var.get():
                        selected.append(name)
                if not selected:
                    messagebox.showwarning("No Selection", "Please select at least one Load Case or Combination!")
                    return
                result['selected'] = selected
                dialog.destroy()

            def on_cancel():
                dialog.destroy()

            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=10)
            tk.Button(btn_frame, text="OK", command=on_ok, bg='#4CAF50', fg='white', font=('Arial', 10, 'bold'), width=12, cursor='hand2').pack(side='left', padx=10)
            tk.Button(btn_frame, text="Cancel", command=on_cancel, bg='#f44336', fg='white', font=('Arial', 10, 'bold'), width=12, cursor='hand2').pack(side='left', padx=10)

            dialog.wait_window()
            return result['selected']
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get Load Cases:\n{str(e)}")
            return None
