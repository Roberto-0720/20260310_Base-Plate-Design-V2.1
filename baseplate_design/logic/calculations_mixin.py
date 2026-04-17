"""
logic/calculations_mixin.py
============================
Low-level calculation helpers used by the hinge/fixed XLSX generator
and the design-check engine.

App state used: self.material_trees
"""

import math
import re


class CalculationsMixin:

    def parse_section_name(self, section):
        """Parse section name like 'H340X250X9X14' to get d, bf, tw, tf"""
        import re
        # Remove 'H' or 'BH' prefix
        section_clean = section.upper().replace('BH', '').replace('H', '')
        # Split by 'X'
        parts = section_clean.split('X')
        if len(parts) >= 4:
            try:
                return {
                    'd': float(parts[0]),
                    'bf': float(parts[1]),
                    'tw': float(parts[2]),
                    'tf': float(parts[3])
                }
            except:
                pass
        return {'d': '', 'bf': '', 'tw': '', 'tf': ''}

    def get_hinge_type_row_data(self, section):
        """Get row data from Hinge Type sheet matching section"""
        tree = self.material_trees.get('Hinge Type')
        if not tree:
            return None

        for item in tree.get_children():
            values = tree.item(item, 'values')
            if values and values[0]:
                if section.upper() in values[0].upper():
                    # Return dict with all column values
                    cols = ['Column size', 'Type', 'No.AB', 'P1', 'N', 'A', 'B', 'C', 'E', 'F', 'P2', 'Y',
                           'Np', 'Bp', 'c', 'nrb', 'drb', 'dtb', 'X-leg', 'Y-leg', 'Layer 1', 'Layer 2']
                    return {cols[i]: values[i] if i < len(values) else '' for i in range(len(cols))}
        return None

    def get_material_strength(self, material_name, material_type):
        """Get strength value from Material Strength sheet"""
        tree_map = {
            'anchor_bolt': 'anchor_bolt',
            'base_plate': 'base_plate',
            'concrete': 'concrete',
            'rebar': 'rebar'
        }

        tree = self.material_trees.get(tree_map.get(material_type))
        if not tree or not material_name:
            return ''

        for item in tree.get_children():
            values = tree.item(item, 'values')
            if values and len(values) >= 2 and values[0] == material_name:
                return values[1]
        return ''

    def get_anchor_bolt_data(self, db_value):
        """Get anchor bolt data from Anchor Bolt Table"""
        tree = self.material_trees.get('Anchor Bolt Table')
        if not tree:
            return None

        # Parse db value (e.g., "M30" -> 30)
        try:
            db_num = int(db_value.replace('M', '').strip())
        except:
            return None

        for item in tree.get_children():
            values = tree.item(item, 'values')
            if values and values[0]:
                try:
                    if int(float(values[0])) == db_num:
                        # Return dict: db, Rmin, a, W, T, S, NutW, nt, Nut Allowance, Edge Min, Leng A1, Leng A2
                        return {
                            'db': values[0] if len(values) > 0 else '',
                            'Rmin': values[1] if len(values) > 1 else '',
                            'a': values[2] if len(values) > 2 else '',
                            'W': values[3] if len(values) > 3 else '',
                            'T': values[4] if len(values) > 4 else '',
                            'S': values[5] if len(values) > 5 else '',
                            'NutW': values[6] if len(values) > 6 else '',
                            'nt': values[7] if len(values) > 7 else '',
                            'Nut_Allowance': values[8] if len(values) > 8 else '',
                            'Edge_Min': values[9] if len(values) > 9 else '',
                            'Leng_A1': values[10] if len(values) > 10 else '',
                            'Leng_A2': values[11] if len(values) > 11 else ''
                        }
                except:
                    continue
        return None

    def calculate_anchor_bolt_ase(self, db_value):
        """Calculate Ase = PI()/4*(db-25.4*0.9743/nt)^2"""
        import math

        bolt_data = self.get_anchor_bolt_data(db_value)
        if not bolt_data:
            return ''

        try:
            db = float(bolt_data['db'])
            nt = float(bolt_data['nt'])

            ase = (math.pi / 4) * ((db - 25.4 * 0.9743 / nt) ** 2)
            return round(ase, 1)
        except:
            return ''

    def calculate_anchor_bolt_values(self, db_value, bolt_data, futa):
            """Calculate A1, Proj, heff for anchor bolt"""
            try:
                # A1 = Leng A1 (từ cột K trong Anchor Bolt Table)
                a1 = float(bolt_data.get('Leng_A1', 0)) if bolt_data.get('Leng_A1') else ''

                # Proj = Nut Allowance (từ cột I trong Anchor Bolt Table)
                proj = float(bolt_data.get('Nut_Allowance', 0)) if bolt_data.get('Nut_Allowance') else ''

                # heff = A1 - Proj - 30 - S
                s_val = float(bolt_data.get('S', 0)) if bolt_data.get('S') else 0
                if a1 != '' and proj != '':
                    heff = a1 - proj - 30 - s_val
                else:
                    heff = ''

                return a1, proj, heff
            except Exception as e:
                print(f"Error calculating anchor bolt values: {e}")
                return '', '', ''

    def calculate_main_bar_ase(self, qty, size):
        """Calculate Main Bar Ase based on qty and rebar size"""
        import math
        try:
            qty_val = int(qty) if qty else 0
            # Parse size: "D19" -> 19
            size_num = int(size.replace('D', '').strip()) if size else 0

            # Ase = qty * π * d² / 4
            ase = qty_val * math.pi * (size_num ** 2) / 4
            return round(ase, 1)
        except:
            return ''
