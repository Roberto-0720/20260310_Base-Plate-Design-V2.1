"""
ui/widgets_mixin.py
====================
Shared Treeview helper methods (add row, delete row, inline cell edit).
No app state dependencies — pure widget interaction.
"""

import tkinter as tk


class WidgetsMixin:

    def add_row_to_material_tree(self, tree, material_type):
        if material_type in ['steel', 'rebar']:
            new_row = ('', '')  # Material Type, strength
        else:  # concrete
            new_row = ('', '')
        tree.insert('', 'end', values=new_row)

    def add_row_to_treeview(self, tree, sheet_name, columns):
        """Add empty row to treeview for tables like Anchor Bolt, Hinge Type, Rebar Dev Length"""
        empty_values = tuple('' for _ in columns)
        tree.insert('', 'end', values=empty_values)

    def delete_row_from_treeview(self, tree):
        """Delete selected row from treeview"""
        selected = tree.selection()
        if selected:
            for item in selected:
                tree.delete(item)

    def edit_treeview_cell(self, event, tree):
        """Edit cell on double-click"""
        item = tree.identify('item', event.x, event.y)
        column = tree.identify('column', event.x, event.y)

        if not item or not column or column == '#0':
            return

        # Get cell position
        x, y, w, h = tree.bbox(item, column)

        # Create entry widget
        entry = tk.Entry(tree, width=15)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, tree.item(item, 'values')[int(column[1:]) - 1])
        entry.focus()

        def save_edit():
            values = list(tree.item(item, 'values'))
            values[int(column[1:]) - 1] = entry.get()
            tree.item(item, values=values)
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind('<Return>', lambda e: save_edit())
        entry.bind('<Escape>', cancel_edit)
        entry.bind('<FocusOut>', lambda e: save_edit())
