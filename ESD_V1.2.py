import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
import pandas as pd
import sqlite3
import os
import re
from tkinter.filedialog import asksaveasfilename
from pandas.io.sql import DatabaseError
import datetime
from datetime import datetime

class ExcelSQLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel SQL Developer")
        self.root.geometry("1100x800")

        # Configuration
        self.max_sample_rows = 1000  # For previews
        self.result_limit = 100000  # Safety limit for exports

        # Define a light color scheme for better visibility
        self.bg_color = "#f0f0f0"  # Light gray background for root and main frames
        self.frame_bg_color = "#ffffff"  # White for inner frames/labels
        self.text_color = "#333333"  # Dark gray for general text

        self.button_bg_color = "#007bff"  # Blue for primary buttons
        self.button_fg_color = "#ffffff"  # White for button text (ensures visibility on blue)
        self.button_active_bg_color = "#0056b3"  # Darker blue on hover/active

        self.entry_bg_color = "#ffffff"  # White for entry fields
        self.entry_fg_color = "#333333"  # Dark gray for entry text

        self.treeview_bg_color = "#ffffff"  # White for treeview background
        self.treeview_fg_color = "#333333"  # Dark gray for treeview text
        self.treeview_heading_bg = "#e0e0e0"  # Light gray for treeview headings
        self.treeview_selected_bg = "#cce0ff"  # Light blue for selected items

        # Initialize variables
        self.file_path = ""
        self.conn = None
        self.table_mapping = {}
        self.current_results = None  # This will hold the DataFrame for export
        self.query_history = []
        self.query_executed = ""  # This will hold the processed query for full export

        # UI Setup
        self.configure_styles()
        self.create_widgets()

    def configure_styles(self):
        """Configure ttk styles for the application"""
        style = ttk.Style()

        # General styles for frames and labels
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.frame_bg_color, foreground=self.text_color)
        style.configure("TLabelframe", background=self.bg_color, foreground=self.text_color)
        style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.text_color)

        # Treeview styles
        style.configure("Treeview",
                        background=self.treeview_bg_color,
                        foreground=self.treeview_fg_color,
                        fieldbackground=self.treeview_bg_color,
                        font=('Helvetica', 9))
        style.configure("Treeview.Heading",
                        background=self.treeview_heading_bg,
                        foreground=self.text_color,
                        font=('Helvetica', 9, 'bold'))
        style.map("Treeview",
                  background=[('selected', self.treeview_selected_bg)])

        # Scrollbar styles
        style.configure("Vertical.TScrollbar", background=self.bg_color, troughcolor=self.bg_color)
        style.configure("Horizontal.TScrollbar", background=self.bg_color, troughcolor=self.bg_color)

    def create_widgets(self):
        """Create all UI widgets"""
        self.root.configure(bg=self.bg_color)  # Set root window background

        # Configure main grid
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        # Left panel - Tables tree and info
        left_frame = tk.Frame(self.root, width=300, bg=self.frame_bg_color, bd=2, relief=tk.RIDGE)
        left_frame.grid(row=0, column=0, rowspan=2, sticky="nswe", padx=5, pady=5)
        left_frame.grid_propagate(False)  # Prevent frame from resizing to content

        # Right top - Query input
        query_frame = ttk.LabelFrame(self.root, text=" SQL Query ")  # Use ttk.LabelFrame for styling
        query_frame.grid(row=0, column=1, sticky="we", padx=5, pady=5)

        # Right bottom - Results
        result_frame = ttk.LabelFrame(self.root, text=" Query Results ")  # Use ttk.LabelFrame for styling
        result_frame.grid(row=1, column=1, sticky="nswe", padx=5, pady=5)

        # Configure panels
        self.setup_left_panel(left_frame)
        self.setup_query_panel(query_frame)
        self.setup_results_panel(result_frame)

    def setup_left_panel(self, frame):
        """Configure the tables explorer panel"""
        frame.grid_columnconfigure(0, weight=1)

        # Browse button (using tk.Button for direct color control)
        browse_btn = tk.Button(frame, text="📂 Browse Excel Files", command=self.browse_files,
                               bg=self.button_bg_color, fg=self.button_fg_color,
                               activebackground=self.button_active_bg_color, activeforeground=self.button_fg_color,
                               relief=tk.RAISED, font=('Helvetica', 10, 'bold'))
        browse_btn.grid(row=0, column=0, pady=10, sticky="ew")

        # Search box
        search_frame = tk.Frame(frame, bg=self.frame_bg_color)
        search_frame.grid(row=1, column=0, sticky="ew", pady=5)

        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.search_var,
                                bg=self.entry_bg_color, fg=self.entry_fg_color,
                                insertbackground=self.entry_fg_color)  # Cursor color
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Search button (using tk.Button for direct color control)
        search_btn = tk.Button(search_frame, text="🔍", command=self.filter_tables, width=3,
                               bg=self.button_bg_color, fg=self.button_fg_color,
                               activebackground=self.button_active_bg_color, activeforeground=self.button_fg_color,
                               relief=tk.RAISED)
        search_btn.pack(side=tk.RIGHT)

        # Tables tree
        tree_frame = tk.Frame(frame, bg=self.frame_bg_color)
        tree_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tables_tree = ttk.Treeview(tree_frame)
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tables_tree.yview)
        self.tables_tree.configure(yscrollcommand=scroll_y.set)

        self.tables_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")

        tk.Label(frame, text="Additional Details:", bg=self.frame_bg_color, fg=self.text_color).grid(row=3, column=0,
                                                                                                   sticky="w",
                                                                                                   pady=(5, 0))
        self.configure_tables_tree()

        # Bind right-click for context menu
        self.tables_tree.bind("<Button-3>", self.show_tables_tree_context_menu)
        self.tables_tree_context_menu = tk.Menu(self.root, tearoff=0)
        self.tables_tree_context_menu.add_command(label="Show Columns", command=self.show_columns_for_selected_table)
        self.tables_tree_context_menu.add_command(label="Copy Table Name", command=self.copy_table_name_to_clipboard)

        # --- New Warning Display Area ---
        self.warning_label = tk.Label(frame, text="Warnings:", bg=self.frame_bg_color, fg="red")
        self.warning_label.grid(row=4, column=0, sticky="w", padx=2, pady=(5, 0))

        self.warning_text = tk.Text(frame, height=5, wrap=tk.WORD,
                                    bg="#fff3cd", relief=tk.FLAT, state=tk.DISABLED, font=('Helvetica', 8))
        self.warning_text.grid(row=5, column=0, sticky="ew", padx=2, pady=2)
        # Define tags for different warning types
        self.warning_text.tag_configure("info", foreground="#664d03")  # Dark yellow
        self.warning_text.tag_configure("error", foreground="#dc3545")  # Red
        self.warning_text.tag_configure("default", foreground="#888888")  # Lighter gray for no warnings

        # Status bar (moved to row 6)
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        tk.Label(frame, textvariable=self.status_var, bd=1, relief=tk.SUNKEN,
                 anchor='w', bg=self.frame_bg_color, fg=self.text_color).grid(row=6, column=0, sticky="ew", pady=5)

    def setup_query_panel(self, frame):
        """Configure the SQL query input panel with keyword highlighting"""
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        # Query text area with syntax highlighting
        self.query_text = tk.Text(frame, height=15, width=100, wrap=tk.NONE,
                                  font=('Consolas', 10),
                                  bg=self.entry_bg_color, fg=self.entry_fg_color,
                                  insertbackground=self.entry_fg_color, undo=True)

        # Configure text tags for highlighting
        self.query_text.tag_configure("keyword", foreground="#0000FF")  # Blue
        self.query_text.tag_configure("string", foreground="#008000")  # Green
        self.query_text.tag_configure("comment", foreground="#808080")  # Gray

        # Basic SQL keywords for highlighting
        self.sql_keywords = [
            "SELECT", "FROM", "WHERE", "GROUP BY", "ORDER BY", "JOIN",
            "INNER JOIN", "LEFT JOIN", "RIGHT JOIN", "AND", "OR", "NOT",
            "IN", "LIKE", "IS NULL", "IS NOT NULL", "INSERT", "UPDATE",
            "DELETE", "CREATE", "ALTER", "DROP", "TABLE", "VIEW", "INDEX",
            "AS", "DISTINCT", "COUNT", "SUM", "AVG", "MIN", "MAX",
            "HAVING", "LIMIT", "OFFSET"
        ]

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.query_text.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=self.query_text.xview)
        self.query_text.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        # Layout
        self.query_text.grid(row=0, column=0, sticky="nsew", pady=5)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="we")

        # Bind highlight update on key release
        self.query_text.bind("<KeyRelease>", self.highlight_syntax)

        # Bind undo/redo shortcuts
        self.query_text.bind("<Control-z>", self._undo_text)
        self.query_text.bind("<Control-y>", self._redo_text)
        self.query_text.bind("<Control-Z>", self._undo_text)
        self.query_text.bind("<Control-Shift-Z>", self._redo_text)

        # Query buttons
        btn_frame = tk.Frame(frame, bg=self.bg_color)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=5)

        buttons = [
            ("▶ Execute", self.execute_query_handler),
            ("📝 Show Tables", self.show_tables_info),
            ("📖 Sample Data", self.show_sample_data),
            ("📋 Clear", self.clear_query),
            ("⏱ History", self.show_query_history),
        ]

        for i, (text, cmd) in enumerate(buttons):
            btn = tk.Button(btn_frame, text=text, command=cmd,
                            bg=self.button_bg_color, fg=self.button_fg_color,
                            activebackground=self.button_active_bg_color,
                            activeforeground=self.button_fg_color,
                            relief=tk.RAISED, font=('Helvetica', 10, 'bold'))
            btn.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)

        # Add to your btn_frame in setup_query_panel
        spool_frame = tk.Frame(frame, bg=self.bg_color)
        spool_frame.grid(row=3, column=0, columnspan=2, sticky="we", pady=5)

        # Spool toggle button
        self.spool_btn = tk.Button(spool_frame, text="🔴 Start Spooling",
                                   command=self.toggle_spooling,
                                   bg="#880000", fg="white",
                                   relief=tk.RAISED, font=('Helvetica', 10))
        self.spool_btn.pack(side=tk.LEFT, padx=2)

        # Spool file label
        self.spool_label = tk.Label(spool_frame, text="No active spool file",
                                    bg=self.bg_color, fg="grey")
        self.spool_label.pack(side=tk.LEFT, padx=5)

    def toggle_spooling(self):
            """Toggle spooling on/off"""
            if getattr(self, 'spooling_active', False):
                self.disable_spool()
                self.spool_btn.config(text="🔴 Start Spooling", bg="#880000")
                self.spool_label.config(text="No active spool file", fg="grey")
            else:
                file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                         filetypes=[("CSV Files", "*.csv"),
                                                                    ("Text Files", "*.txt")])
                if file_path:
                    if self.enable_spool(file_path):
                        self.spool_btn.config(text="✅ Stop Spooling", bg="#008800")
                        self.spool_label.config(text=f"Spooling to: {file_path}", fg="black")

    def highlight_syntax(self, event=None):
        """Basic SQL syntax highlighting"""
        # Remove previous highlighting
        self.query_text.tag_remove("keyword", "1.0", tk.END)
        self.query_text.tag_remove("string", "1.0", tk.END)
        self.query_text.tag_remove("comment", "1.0", tk.END)

        text = self.query_text.get("1.0", tk.END)

        # Highlight keywords
        for word in self.sql_keywords:
            start = "1.0"
            while True:
                start = self.query_text.search(rf'\m{word}\M', start,
                                               stopindex=tk.END,
                                               regexp=True,
                                               nocase=True)
                if not start:
                    break
                end = f"{start}+{len(word)}c"
                self.query_text.tag_add("keyword", start, end)
                start = end

        # Highlight strings (simple single/double quoted)
        start = "1.0"
        while True:
            start = self.query_text.search(r'["\'].*?["\']', start, stopindex=tk.END)
            if not start:
                break
            end = self.query_text.index(f"{start} + {len(self.query_text.get(start, start + '1c'))}c")
            self.query_text.tag_add("string", start, end)
            start = end

        # Highlight comments (both -- and /* */ styles)
        start = "1.0"
        while True:
            start = self.query_text.search(r'--.*?$', start, stopindex=tk.END)
            if not start:
                break
            end = self.query_text.index(f"{start} lineend")
            self.query_text.tag_add("comment", start, end)
            start = end

        start = "1.0"
        while True:
            start = self.query_text.search(r'/\*.*?\*/', start, stopindex=tk.END, regexp=True)
            if not start:
                break
            end = self.query_text.index(f"{start} + {len(self.query_text.get(start, start + '1c'))}c")
            self.query_text.tag_add("comment", start, end)
            start = end

    def _undo_text(self, event=None):
        try:
            self.query_text.edit_undo()
        except tk.TclError:
            pass  # Nothing to undo
        return "break"  # Prevent default binding

    def _redo_text(self, event=None):
        try:
            self.query_text.edit_redo()
        except tk.TclError:
            pass  # Nothing to redo
        return "break"  # Prevent default binding

    def setup_results_panel(self, frame):
        """Configure the results display panel with proper scrolling"""
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        # Create container frame for treeview and scrollbars
        container = tk.Frame(frame, bg=self.bg_color)
        container.grid(row=0, column=0, sticky="nsew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Create horizontal scrollbar first
        hscroll = ttk.Scrollbar(container, orient="horizontal")
        hscroll.grid(row=1, column=0, sticky="ew")

        # Create vertical scrollbar
        vscroll = ttk.Scrollbar(container, orient="vertical")
        vscroll.grid(row=0, column=1, sticky="ns")

        # Create the treeview
        self.result_tree = ttk.Treeview(
            container,
            xscrollcommand=hscroll.set,
            yscrollcommand=vscroll.set,
            show="headings",
            selectmode="extended"
        )
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # Configure scrollbars
        hscroll.config(command=self.result_tree.xview)
        vscroll.config(command=self.result_tree.yview)

        # Initialize context menu for result tree
        self.result_tree_context_menu = tk.Menu(self.root, tearoff=0)
        self.result_tree_context_menu.add_command(label="Copy Cell Value", command=self.copy_cell_value)
        self.result_tree_context_menu.add_command(label="Copy Column Name", command=self.copy_column_name)
        self.result_tree.bind("<Button-3>", self.show_result_tree_context_menu)

        # Result buttons
        btn_frame = tk.Frame(frame, bg=self.bg_color)
        btn_frame.grid(row=2, column=0, sticky="e", pady=5)

        buttons = [
            ("💾 Export to Excel", self.export_to_excel),
            ("🧹 Clear Results", self.clear_results),
        ]

        for text, cmd in reversed(buttons):
            btn = tk.Button(btn_frame, text=text, command=cmd,
                            bg=self.button_bg_color, fg=self.button_fg_color,
                            activebackground=self.button_active_bg_color,
                            activeforeground=self.button_fg_color,
                            relief=tk.RAISED, font=('Helvetica', 10, 'bold'))
            btn.pack(side="right", padx=2)

        # Status bar
        self.result_status_var = tk.StringVar()
        self.result_status_var.set("No results")
        tk.Label(frame, textvariable=self.result_status_var, bd=1, relief=tk.SUNKEN,
                 anchor='w', bg=self.frame_bg_color, fg=self.text_color).grid(row=3, column=0,
                                                                              columnspan=2, sticky="ew")

    def show_result_tree_context_menu(self, event):
        """Display context menu for result treeview"""
        try:
            # Identify the item and column that was right-clicked
            item_id = self.result_tree.identify_row(event.y)
            column_id = self.result_tree.identify_column(event.x)

            # Store the clicked column ID for use in copy_cell_value and copy_column_name
            self._clicked_column_id = column_id

            if item_id and column_id:
                # Select the item and focus on it
                self.result_tree.selection_set(item_id)
                self.result_tree.focus(item_id)

                # Enable both options
                self.result_tree_context_menu.entryconfig("Copy Cell Value", state="normal")
                self.result_tree_context_menu.entryconfig("Copy Column Name", state="normal")
                self.result_tree_context_menu.tk_popup(event.x_root, event.y_root)
            elif column_id:  # Only column header was clicked
                # No item selected, but a column header was clicked
                self.result_tree_context_menu.entryconfig("Copy Cell Value", state="disabled")
                self.result_tree_context_menu.entryconfig("Copy Column Name", state="normal")
                self.result_tree_context_menu.tk_popup(event.x_root, event.y_root)
            else:
                # If neither item nor column is clicked, disable both options
                self.result_tree_context_menu.entryconfig("Copy Cell Value", state="disabled")
                self.result_tree_context_menu.entryconfig("Copy Column Name", state="disabled")
        finally:
            self.result_tree_context_menu.grab_release()

    def copy_cell_value(self):
        """Copy the value of the selected cell to the clipboard."""
        selected_item = self.result_tree.focus()
        if not selected_item:
            self.status_var.set("No cell selected to copy.")
            return

        # Use the stored clicked column ID
        column_id = getattr(self, '_clicked_column_id', None)
        if not column_id:
            self.status_var.set("Could not determine which column was clicked.")
            return

        column_index = int(column_id.replace("#", "")) - 1  # Convert to 0-based index

        # Get the values of the identified row
        row_values = self.result_tree.item(selected_item, 'values')

        if 0 <= column_index < len(row_values):
            cell_value = row_values[column_index]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(cell_value))  # Ensure it's a string
            self.status_var.set(f"Copied value: {cell_value}")  # Update status instead of message box
        else:
            self.status_var.set("Invalid column selected.")

    def copy_column_name(self):
        """Copy the name of the selected column to the clipboard."""
        # Use the stored clicked column ID
        column_id = getattr(self, '_clicked_column_id', None)
        if not column_id:
            self.status_var.set("Could not determine which column was clicked.")
            return

        # Get the column name from the treeview
        column_name = self.result_tree.heading(column_id)['text']

        if column_name:
            self.root.clipboard_clear()
            self.root.clipboard_append(column_name)
            self.status_var.set(f"Copied column name: {column_name}")  # Update status instead of message box
        else:
            self.status_var.set("No column name found.")

    def configure_tables_tree(self):
        """Configure the tables explorer treeview"""
        self.tables_tree["columns"] = ("type", "rows")
        self.tables_tree.column("#0", width=200, minwidth=200)
        self.tables_tree.column("type", width=80, minwidth=80)
        self.tables_tree.column("rows", width=70, minwidth=70)

        self.tables_tree.heading("#0", text="Table")
        self.tables_tree.heading("type", text="Type")
        self.tables_tree.heading("rows", text="Rows")

    def show_tables_tree_context_menu(self, event):
        """Display context menu for tables tree"""
        try:
            # Select the item that was right-clicked
            item_id = self.tables_tree.identify_row(event.y)
            if item_id:
                self.tables_tree.selection_set(item_id)
                self.tables_tree.focus(item_id)

                item = self.tables_tree.item(item_id)
                # Enable "Show Columns" and "Copy Table Name" only for sheet nodes
                if item['values'] and item['values'][0] == "Sheet":
                    self.tables_tree_context_menu.entryconfig("Show Columns", state="normal")
                    self.tables_tree_context_menu.entryconfig("Copy Table Name", state="normal")
                else:
                    self.tables_tree_context_menu.entryconfig("Show Columns", state="disabled")
                    self.tables_tree_context_menu.entryconfig("Copy Table Name", state="disabled")

                self.tables_tree_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tables_tree_context_menu.grab_release()

    def _update_warning_display(self, warnings_list):
        """
        Updates the dedicated warning text area.
        warnings_list is a list of tuples: (message, type) where type is "info" or "error".
        """
        self.warning_text.config(state=tk.NORMAL)
        self.warning_text.delete("1.0", tk.END)
        if warnings_list:
            for msg, msg_type in warnings_list:
                # Insert with a tag for coloring
                if msg_type == "error":
                    self.warning_text.insert(tk.END, msg + "\n", "error")  # Use "error" tag
                else:  # Default to info/yellow
                    self.warning_text.insert(tk.END, msg + "\n", "info")  # Use "info" tag
        else:
            self.warning_text.insert(tk.END, "No warnings.", "default")  # Use "default" tag
        self.warning_text.config(state=tk.DISABLED)

    def browse_files(self):
        """Browse for Excel files and load them into SQLite"""
        path = filedialog.askdirectory()
        if not path:
            return

        self.file_path = path
        self.status_var.set("Loading Excel files...")
        self.current_results = None
        self.clear_ui()
        self.root.update_idletasks()

        collected_warnings = []  # Collect warnings here as (message, type) tuples

        try:
            self.conn = sqlite3.connect(':memory:')
            self.conn.text_factory = str
            self.table_mapping = {}

            excel_files = [f for f in os.listdir(self.file_path)
                           if f.lower().endswith(('.xlsx', '.xls'))]

            if not excel_files:
                self.status_var.set("No Excel files found in selected directory.")
                self._update_warning_display([("No Excel files found in selected directory.", "info")])
                return

            # Process files with progress
            total_files = len(excel_files)
            for i, filename in enumerate(excel_files, 1):
                self.status_var.set(f"Loading files ({i}/{total_files}): {filename[:20]}...")
                self.root.update_idletasks()
                # Pass the warnings list to load_excel_file
                self.load_excel_file(filename, collected_warnings)

            self.populate_tables_tree()
            final_status_message = f"Loaded {len(self.table_mapping)} tables from {len(excel_files)} files"
            if collected_warnings:
                final_status_message += f" ({len(collected_warnings)} warnings)"
            self.status_var.set(final_status_message)
            self._update_warning_display(collected_warnings)  # Display collected warnings

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel files:\n{str(e)}")
            self.status_var.set("Error loading files")
            self._update_warning_display([
                (f"Critical error loading files: {str(e)}", "error")
            ])
            self.conn = None

    def load_excel_file(self, filename, collected_warnings):
        """Load all sheets from an Excel file into SQLite"""
        file_path = os.path.join(self.file_path, filename)
        file_base = os.path.splitext(filename)[0]

        # Sanitize file_key and check for changes
        original_file_key = file_base.lower().replace(' ', '_')
        file_key = re.sub(r'[^a-z0-9_]', '', original_file_key)
        if original_file_key != file_key:
            collected_warnings.append(
                (f"File name '{file_base}' sanitized to '{file_key}' for internal use due to special characters/spaces.",
                 "info")
            )

        try:
            xls = pd.ExcelFile(file_path)

            for sheet_name in xls.sheet_names:
                full_sheet_name_display = f"{file_base}.{sheet_name}"  # For display in warnings

                # Sanitize sheet_name for SQL table name and check for changes
                original_sheet_name_for_sql = sheet_name.lower().replace(' ', '_')
                sanitized_sheet_name_for_sql = re.sub(r'[^a-z0-9_]', '', original_sheet_name_for_sql)

                if original_sheet_name_for_sql != sanitized_sheet_name_for_sql:
                    collected_warnings.append(
                        (f"Sheet name '{sheet_name}' in file '{file_base}' sanitized to '{sanitized_sheet_name_for_sql}' for internal use due to special characters/spaces.",
                         "info")
                    )

                try:
                    # Read the Excel sheet without a header first
                    df_raw = pd.read_excel(xls, sheet_name, header=None, na_values=['', 'NA', 'NULL'])

                    if df_raw.empty:
                        collected_warnings.append(
                            (f"Sheet '{full_sheet_name_display}' is empty and will not be loaded.", "info"))
                        continue  # Skip empty sheets

                    # Extract the first row as potential headers
                    original_headers = df_raw.iloc[0].tolist()
                    # Drop the header row from the DataFrame
                    df = df_raw[1:].copy()
                    df.reset_index(drop=True, inplace=True)  # Reset index after dropping row

                    # Drop columns that are entirely NaN after header extraction
                    df = df.dropna(axis=1, how='all')

                    # --- Custom Column Name Processing ---
                    processed_columns = []
                    seen_final_names = set()
                    for i, col_name_raw in enumerate(original_headers):
                        # If column was dropped due to all NaNs, its header won't be used
                        if i not in df.columns:
                            continue

                        original_col_str = str(col_name_raw).strip() if pd.notna(
                            col_name_raw) else f"Unnamed_Column_{i}"

                        # Step 1: Handle original duplicates
                        base_name_for_dup_check = original_col_str
                        count = 1
                        while base_name_for_dup_check in seen_final_names:
                            base_name_for_dup_check = f"{original_col_str}_{count}"
                            count += 1

                        if base_name_for_dup_check != original_col_str:
                            collected_warnings.append(
                                (f"'{full_sheet_name_display}': Original column '{original_col_str}' is a duplicate. Renamed to '{base_name_for_dup_check}'.",
                                 "info")
                            )

                        # Step 2: Sanitize for special characters/spaces
                        sanitized_name = re.sub(r'[^a-zA-Z0-9_]', '', base_name_for_dup_check.replace(' ', '_'))

                        if sanitized_name != base_name_for_dup_check:
                            collected_warnings.append(
                                (f"'{full_sheet_name_display}': Column '{base_name_for_dup_check}' renamed to '{sanitized_name}' due to special characters or spaces.",
                                 "info")
                            )

                        # Final check for uniqueness after full sanitization (should be rare if logic is correct)
                        final_col_name = sanitized_name
                        counter_final = 1
                        while final_col_name in seen_final_names:
                            final_col_name = f"{sanitized_name}_{counter_final}"
                            counter_final += 1
                            # This case should ideally not happen if previous duplicate handling is robust
                            # but acts as a safeguard.
                            if counter_final == 2:  # Only warn once for the first append
                                collected_warnings.append(
                                    (f"'{full_sheet_name_display}': Column '{sanitized_name}' became a duplicate after sanitization. Renamed to '{final_col_name}'.",
                                     "info")
                                )

                        processed_columns.append(final_col_name)
                        seen_final_names.add(final_col_name)

                    # Assign the new column names to the DataFrame
                    df.columns = processed_columns

                    sql_table_name = f"{file_key}_{sanitized_sheet_name_for_sql}"
                    sql_table_name = re.sub(r'[^a-z0-9_]', '', sql_table_name)  # Final check for table name

                    self.table_mapping[full_sheet_name_display] = sql_table_name
                    df.to_sql(sql_table_name, self.conn, index=False, if_exists='replace')

                except Exception as e:
                    collected_warnings.append((f"Error loading sheet '{full_sheet_name_display}': {str(e)}", "error"))
                    print(f"Error loading {filename} sheet {sheet_name}: {str(e)}")  # Keep for console debug

        except Exception as e:
            collected_warnings.append((f"Error loading file '{filename}': {str(e)}", "error"))
            print(f"Error loading {filename}: {str(e)}")  # Keep for console debug

    def populate_tables_tree(self):
        """Display all tables in a hierarchical view"""
        for item in self.tables_tree.get_children():
            self.tables_tree.delete(item)

        # Group by file
        files = {}
        for dot_name, sql_name in self.table_mapping.items():
            file, sheet = dot_name.split('.', 1)
            if file not in files:
                files[file] = []
            row_count = self.get_row_count(sql_name)
            files[file].append((sheet, sql_name, row_count))

        # Add to tree
        for file, sheets in sorted(files.items()):
            # Count total rows for the file
            file_rows = sum(row_count for _, _, row_count in sheets)

            file_node = self.tables_tree.insert("", "end", text=file,
                                                values=("Excel", f"{file_rows:,}"))

            for sheet, sql_name, row_count in sorted(sheets):
                self.tables_tree.insert(file_node, "end", text=sheet,
                                        values=("Sheet", f"{row_count:,}"))

    def get_row_count(self, table_name):
        """Get row count for a table"""
        try:
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM \"{table_name}\"")
            return cursor.fetchone()[0]
        except sqlite3.Error as e:
            print(f"Error getting row count for {table_name}: {e}")
            return 0

    def execute_query_handler(self):
        """Handles query execution while preserving selection functionality"""
        query_text = self._get_query_to_execute()
        if not query_text:
            messagebox.showwarning("Input Error", "Please enter or select a SQL query")
            return

        # Support both single query (original behavior) and multiple queries separated by semicolons
        queries = [q.strip() for q in query_text.split(';') if q.strip()]

        try:
            for i, query in enumerate(queries):
                if hasattr(self, 'spooling_active') and self.spooling_active:
                    self._write_query_header(query, i == 0)

                processed_query = self.process_query(query)
                result_df = pd.read_sql_query(processed_query, self.conn)
                self._handle_query_results(result_df, i, len(queries))

        except Exception as e:
            self.handle_sql_error(str(e))

    def _write_query_header(self, query, is_first_query):
        """Write query header to spool file"""
        self.spool_file.write(f"\n--- Query executed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        self.spool_file.write(f"{query}\n")
        self.spool_file.write("-" * 80 + "\n")

    def _handle_query_results(self, result_df, query_index, total_queries):
        """Process and display query results"""
        if hasattr(self, 'spooling_active') and self.spooling_active:
            result_df.to_csv(
                self.spool_file,
                mode='a',
                index=False,
                header=(query_index == 0)  # Only write headers for first query
            )

        # For single query, show immediately
        # For multiple queries, show last query's results but keep all in current_results
        if query_index == total_queries - 1:
            self.current_results = result_df
            self.show_results(result_df)

    def _execute_core_query(self, query_text_to_execute):  # Renamed to be an internal helper
        """Core logic for executing a SQL query and displaying results."""
        query = query_text_to_execute.strip()
        if not query:
            messagebox.showwarning("Input Error", "Please enter a SQL query")
            self.current_results = None
            return

        self.result_status_var.set("Executing query...")
        self.root.update_idletasks()

        try:
            # Validate the original query string (before processing)
            self.validate_query(query)

            # Process the query for table name mapping
            processed_query = self.process_query(query)

            # Store the processed query for full export later
            self.query_executed = processed_query

            limited_query = f"{processed_query} LIMIT {self.max_sample_rows}"
            # Check if the original query already contained a LIMIT clause
            if "LIMIT" not in query.upper():
                limited_query += " -- Original query automatically limited"

            result_df = pd.read_sql_query(limited_query, self.conn)

            self.current_results = result_df

            # Only add to history if it's not a sample data query (which adds itself)
            # and if it's not already the last query in history (to avoid duplicates from re-execution)
            if not query.startswith("-- Sample data from") and (
                    not self.query_history or self.query_history[-1] != query):
                self.query_history.append(query)

            self.show_results(result_df)

            row_count = len(result_df.index)
            limited_note = " (limited)" if "LIMIT" not in query.upper() else ""
            self.result_status_var.set(f"Showing {row_count:,} rows{limited_note}")

        except pd.io.sql.DatabaseError as e:
            self.handle_sql_error(str(e))
            self.current_results = None
        except Exception as e:
            self.show_error("Error", str(e))
            self.result_status_var.set("Query failed")
            self.current_results = None

    def process_query(self, query):
        """
        Converts file.sheet notation to SQL table names (e.g., "file_sheet")
        while preserving aliases and not misinterpreting alias.column_name.
        """
        processed_query = query

        # Sort table mappings by the length of the dot_name in descending order.
        # This ensures that if "file.sheet_a" and "file.sheet_a_b" exist,
        # "file.sheet_a_b" is replaced first, preventing partial replacements.
        sorted_table_mappings = sorted(self.table_mapping.items(), key=lambda item: len(item[0]), reverse=True)

        # Iterate through the sorted table mappings and perform replacements.
        # The key is to use a regex that specifically targets the table name
        # and avoids matching alias.column_name patterns.
        for dot_name, sql_name in sorted_table_mappings:
            # Construct a regex pattern for the current dot_name.
            # re.escape() handles special characters in dot_name (like the dot itself).
            #
            # IMPORTANT CHANGE HERE:
            # The original pattern `\b` + re.escape(dot_name) + `\b(?!\.\w+)` was too restrictive.
            # The final `\b` (word boundary) would fail if dot_name ended with a non-word char (like comma).
            #
            # New pattern:
            # `\b` at the start ensures we match a whole word.
            # `re.escape(dot_name)` matches the literal dot_name.
            # `(?=\W|$)` is a positive lookahead that asserts the match is followed by
            # a non-word character (`\W`) OR the end of the string (`$`).
            # This allows dot_names ending in punctuation to be matched correctly.
            # `(?!\.\w+)` is still needed to prevent matching `alias.column_name` where `alias` is a dot_name.
            pattern = r'\b' + re.escape(dot_name) + r'(?=\W|$)(?!\.\w+)'

            # Replace the found 'dot_name' with the quoted SQL table name.
            # re.IGNORECASE ensures case-insensitive matching for the dot_name.
            processed_query = re.sub(pattern, f'"{sql_name}"', processed_query, flags=re.IGNORECASE)

        return processed_query

    def validate_query(self, query):
        """Basic query validation to prevent harmful operations, ignoring comments"""
        # Remove single-line comments
        clean_query = re.sub(r'--.*?$', '', query, flags=re.MULTILINE)
        # Remove multi-line comments
        clean_query = re.sub(r'/\*.*?\*/', '', clean_query, flags=re.DOTALL)

        # Check for blocked keywords in the cleaned query
        blocked_keywords = [
            "DROP ", "DELETE ", "UPDATE ", "INSERT ", "ALTER ",
            "CREATE ", "VACUUM ", "ATTACH ", "DETACH ", "PRAGMA ",
            "TRANSACTION ", "ROLLBACK", "COMMIT", "REINDEX"
        ]

        # Convert the cleaned query to uppercase for case-insensitive comparison
        clean_query_upper = clean_query.upper()

        # Check if any blocked keyword is present in the cleaned query
        if any(kw in clean_query_upper for kw in blocked_keywords):
            raise DatabaseError("Modification queries are not allowed")

        # Check for multiple statements by looking for more than one semicolon
        # after stripping comments and leading/trailing whitespace
        no_comments = clean_query.strip()
        if no_comments.count(';') > 1 or (no_comments.count(';') == 1 and not no_comments.endswith(';')):
            raise DatabaseError("Multiple statements not allowed")

    def export_to_excel(self):
        """Export current results to Excel file"""
        if not hasattr(self, 'current_results') or self.current_results.empty:
            messagebox.showwarning("No Data", "No query results to export. Please run a query first.")
            return

        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Results As"
            )

            if filename:
                self.current_results.to_excel(filename, index=False)
                self.status_var.set(f"Exported {len(self.current_results)} rows to {filename}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data:\n{str(e)}")

    def show_results(self, df):
        """Display pandas dataframe in the result tree"""
        self.clear_results()  # This will clear the treeview display and reset status, but not self.current_results

        if df.empty:
            self.result_status_var.set("Query returned empty results")
            self.current_results = df  # Set current_results even if empty for consistency
            return

        # Configure columns
        self.result_tree["columns"] = list(df.columns)
        for col in df.columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=100)

        # Insert data with progress for large datasets
        self.result_status_var.set(f"Loading {len(df):,} rows...")
        self.root.update_idletasks()

        step = max(1, len(df) // 50)  # Update progress every 2%

        for i, row in df.iterrows():
            self.result_tree.insert("", "end", values=list(row))
            if i % step == 0 or i == len(df) - 1:
                self.result_status_var.set(f"Loading rows {i + 1}/{len(df)}")
                self.root.update_idletasks()

        # Auto-resize columns
        self.auto_resize_columns(df)
        self.result_status_var.set(f"Showing {len(df):,} rows")  # Final status update

    def auto_resize_columns(self, df):
        """Automatically resize columns based on content"""
        for col in df.columns:
            max_len = max(df[col].astype(str).apply(len).max(), len(col))
            width = min(300, max(50, max_len * 8))
            self.result_tree.column(col, width=width)

    def show_tables_info(self):
        """Show metadata about all tables"""
        if not self.conn:
            messagebox.showwarning("No Database", "Please load Excel files first")
            return

        try:
            # Query to get table info from sqlite_master and pragma_table_info
            # Note: pragma_table_info needs to be executed for each table, so we collect it
            # and then build a DataFrame.
            cursor = self.conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
            table_names = [row[0] for row in cursor.fetchall()]

            tables_info = []
            for sql_name in table_names:
                # Find the original dot_name from table_mapping
                original_name = next((k for k, v in self.table_mapping.items() if v == sql_name), sql_name)

                # Get column info
                col_cursor = self.conn.cursor()
                col_cursor.execute(f"PRAGMA table_info(\"{sql_name}\");")
                columns_data = col_cursor.fetchall()
                columns_count = len(columns_data)
                columns_list = ", ".join([col[1] for col in columns_data])  # col[1] is the column name

                # Get row count
                row_count = self.get_row_count(sql_name)

                tables_info.append({
                    'original_name': original_name,
                    'sql_name': sql_name,
                    'columns_count': columns_count,
                    'columns': columns_list,
                    'rows': row_count
                })

            # Create and show dataframe
            info_df = pd.DataFrame(tables_info)
            info_df = info_df[['original_name', 'sql_name', 'columns_count', 'rows', 'columns']]

            self.current_results = info_df  # Set current_results for export
            # A generic query for export that reflects the data shown
            self.query_executed = "SELECT original_name, sql_name, columns_count, rows, columns FROM tables_metadata_view"
            self.show_results(info_df)

        except Exception as e:
            self.show_error("Error", f"Failed to get tables info:\n{str(e)}")
            self.current_results = None  # Clear results on error

    def show_columns_for_selected_table(self):
        """Show all columns and their types for the selected table in a new window."""
        if not self.conn:
            messagebox.showwarning("No Database", "Please load Excel files first.")
            return

        selected = self.tables_tree.focus()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a table from the left panel first.")
            return

        item = self.tables_tree.item(selected)
        if not item['values'] or item['values'][0] != "Sheet":
            messagebox.showwarning("Invalid Selection", "Please select a specific sheet (table) to view columns.")
            return

        file_name = self.tables_tree.item(self.tables_tree.parent(selected))['text']
        sheet_name = item['text']
        dot_name = f"{file_name}.{sheet_name}"

        if dot_name not in self.table_mapping:
            messagebox.showerror("Error", "Table mapping not found for selected item.")
            return

        sql_name = self.table_mapping[dot_name]

        try:
            cursor = self.conn.cursor()
            cursor.execute(f"PRAGMA table_info(\"{sql_name}\");")
            columns_info = cursor.fetchall()

            if not columns_info:
                messagebox.showinfo("Columns Info", f"No columns found for table: {dot_name}")
                return

            # Create a new Toplevel window for column details
            columns_window = tk.Toplevel(self.root)
            columns_window.title(f"Columns for {dot_name}")
            columns_window.geometry("600x400")
            columns_window.configure(bg=self.bg_color)
            columns_window.transient(self.root)  # Make it appear on top of the main window
            columns_window.grab_set()  # Make it modal

            # Frame for treeview
            tree_frame = tk.Frame(columns_window, bg=self.frame_bg_color)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            tree_frame.grid_rowconfigure(0, weight=1)
            tree_frame.grid_columnconfigure(0, weight=1)

            columns_tree = ttk.Treeview(tree_frame, show="headings")
            columns_tree["columns"] = ("Name", "Type", "Not Null", "Primary Key")
            columns_tree.heading("Name", text="Name")
            columns_tree.heading("Type", text="Type")
            columns_tree.heading("Not Null", text="Not Null")
            columns_tree.heading("Primary Key", text="Primary Key")

            columns_tree.column("Name", width=150, anchor="w")
            columns_tree.column("Type", width=100, anchor="w")
            columns_tree.column("Not Null", width=80, anchor="center")
            columns_tree.column("Primary Key", width=100, anchor="center")

            scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=columns_tree.yview)
            scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=columns_tree.xview)
            columns_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

            columns_tree.grid(row=0, column=0, sticky="nsew")
            scroll_y.grid(row=0, column=1, sticky="ns")
            scroll_x.grid(row=1, column=0, sticky="we")

            # Populate the treeview
            for col_id, name, col_type, notnull, default_val, pk in columns_info:
                columns_tree.insert("", "end",
                                    values=(name, col_type, "Yes" if notnull else "No", "Yes" if pk else "No"))

            # Button to copy columns
            copy_btn = tk.Button(columns_window, text="📋 Copy All Columns to Clipboard",
                                 command=lambda: self.copy_columns_to_clipboard(columns_tree),
                                 bg=self.button_bg_color, fg=self.button_fg_color,
                                 activebackground=self.button_active_bg_color, activeforeground=self.button_fg_color,
                                 relief=tk.RAISED, font=('Helvetica', 10, 'bold'))
            copy_btn.pack(pady=10)

            columns_window.wait_window()  # Wait for the window to close

        except Exception as e:
            self.show_error("Error", f"Failed to retrieve column info:\n{str(e)}")

    def copy_columns_to_clipboard(self, treeview):
        """Copies all column names from the given treeview to the clipboard."""
        column_names = []
        for item_id in treeview.get_children():
            values = treeview.item(item_id, 'values')
            if values:
                column_names.append(values[0])  # Assuming column name is the first value

        if column_names:
            columns_string = ", ".join(column_names)
            self.root.clipboard_clear()
            self.root.clipboard_append(columns_string)
            messagebox.showinfo("Copied", "Column names copied to clipboard!")
        else:
            messagebox.showwarning("No Columns", "No column names to copy.")

    def copy_table_name_to_clipboard(self):
        """Copies the selected table's file.sheet name to the clipboard."""
        selected = self.tables_tree.focus()
        if not selected:
            self.status_var.set("No table selected to copy.")
            return

        item = self.tables_tree.item(selected)
        if not item['values'] or item['values'][0] != "Sheet":
            self.status_var.set("Please select a specific sheet (table) to copy its name.")
            return

        file_name = self.tables_tree.item(self.tables_tree.parent(selected))['text']
        sheet_name = item['text']
        dot_name = f"{file_name}.{sheet_name}"

        self.root.clipboard_clear()
        self.root.clipboard_append(dot_name)
        self.status_var.set(f"Copied table name: {dot_name}")  # Update status instead of message box

    def show_sample_data(self):
        """Show sample data for selected table"""
        if not self.conn:  # Check if database is loaded
            messagebox.showwarning("No Database", "Please load Excel files first.")
            return

        selected = self.tables_tree.focus()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a table from the left panel first.")
            return

        item = self.tables_tree.item(selected)
        # Ensure it's a sheet node, not a file node
        if not item['values'] or item['values'][0] != "Sheet":
            messagebox.showwarning("Invalid Selection", "Please select a specific sheet (table) to view sample data.")
            return

        # Get file.sheet notation
        file_name = self.tables_tree.item(self.tables_tree.parent(selected))['text']
        sheet_name = item['text']
        dot_name = f"{file_name}.{sheet_name}"

        if dot_name not in self.table_mapping:
            messagebox.showerror("Error", "Table mapping not found for selected item.")
            return

        sql_name = self.table_mapping[dot_name]
        query = f'SELECT * FROM "{sql_name}" LIMIT {self.max_sample_rows}'

        try:
            result_df = pd.read_sql_query(query, self.conn)
            self.current_results = result_df  # Set current_results for export
            self.query_executed = query  # Store the query for full export
            self.query_history.append(f"-- Sample data from {dot_name}\n{query}")

            self.show_results(result_df)  # Display results in the treeview

            row_count = len(result_df.index)
            self.result_status_var.set(
                f"Showing {row_count:,} sample rows from {dot_name}")  # Update status specifically for sample data

        except Exception as e:
            self.show_error("Error", f"Failed to get sample data:\n{str(e)}")
            self.current_results = None  # Clear results on error

    def filter_tables(self):
        """Filter tables tree based on search text"""
        search_text = self.search_var.get().lower()
        self.populate_tables_tree_filtered(search_text)

    def populate_tables_tree_filtered(self, search_text=""):
        """Display all tables in a hierarchical view, filtered by search_text"""
        for item in self.tables_tree.get_children():
            self.tables_tree.delete(item)

        files = {}
        for dot_name, sql_name in self.table_mapping.items():
            file, sheet = dot_name.split('.', 1)
            row_count = self.get_row_count(sql_name)

            # Check if file or sheet name matches the search text
            if search_text.lower() in file.lower() or search_text.lower() in sheet.lower():
                if file not in files:
                    files[file] = []
                files[file].append((sheet, sql_name, row_count))

        for file, sheets in sorted(files.items()):
            file_rows = sum(row_count for _, _, row_count in sheets)
            file_node = self.tables_tree.insert("", "end", text=file,
                                                values=("Excel", f"{file_rows:,}"))
            self.tables_tree.item(file_node, open=True)  # Keep parent open if it has matching children

            for sheet, sql_name, row_count in sorted(sheets):
                self.tables_tree.insert(file_node, "end", text=sheet,
                                        values=("Sheet", f"{row_count:,}"))

    def show_query_history(self):
        """Display previously executed queries"""
        if not self.query_history:
            messagebox.showinfo("History", "No queries in history yet")
            return

        history_window = tk.Toplevel(self.root)
        history_window.title("Query History")
        history_window.geometry("800x600")
        history_window.configure(bg=self.bg_color)
        history_window.transient(self.root)  # Make it appear on top of the main window
        history_window.grab_set()  # Make it modal

        text = tk.Text(history_window, wrap=tk.WORD,
                       bg=self.entry_bg_color, fg=self.entry_fg_color,
                       font=('Consolas', 10))
        scroll = ttk.Scrollbar(history_window, command=text.yview)
        text.configure(yscrollcommand=scroll.set)

        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Display history in reverse chronological order
        for i, q in enumerate(reversed(self.query_history)):
            text.insert(tk.END, f"--- Query {len(self.query_history) - i} ---\n")
            text.insert(tk.END, q)
            text.insert(tk.END, "\n\n")

        text.configure(state='disabled')

        def load_query():
            try:
                # Get selected text, remove history header
                selected_text = text.get(tk.SEL_FIRST, tk.SEL_LAST)
                # Remove the "--- Query X ---" line if present
                lines = selected_text.splitlines()
                if lines and lines[0].startswith("--- Query"):
                    selected_text = "\n".join(lines[1:])

                self.query_text.delete("1.0", tk.END)
                self.query_text.insert("1.0", selected_text.strip())
                history_window.destroy()
            except tk.TclError:  # No text selected
                messagebox.showwarning("Selection Error", "Please select a query to load.")

        # This button and wait_window must be inside the method, after history_window is defined.
        load_btn = tk.Button(history_window, text="Load Selected Query", command=load_query,
                             bg=self.button_bg_color, fg=self.button_fg_color,
                             activebackground=self.button_active_bg_color, activeforeground=self.button_fg_color,
                             relief=tk.RAISED, font=('Helvetica', 10, 'bold'))
        load_btn.pack(pady=5)

        history_window.wait_window()  # Wait for the window to close

    def clear_query(self):
        """Clear the query text area"""
        self.query_text.delete("1.0", tk.END)

    def clear_results(self):
        """Clear the results tree"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.result_tree["columns"] = []
        self.result_status_var.set("Results cleared")
        # IMPORTANT: Do NOT set self.current_results = None here.
        # It should only be set to None if a query fails or returns truly empty results.
        # clear_results is just for clearing the display.

    def clear_ui(self):
        """Clear all UI elements"""
        self.clear_query()
        self.clear_results()
        self.tables_tree.delete(*self.tables_tree.get_children())
        self.status_var.set("Ready")  # Reset main status
        self._update_warning_display([])  # Clear warnings
        self.current_results = None  # Clear current_results when loading new files or clearing all UI
        self.query_executed = ""  # Clear stored query
        self.populate_tables_tree()  # Re-populate the tree without filter

    def handle_sql_error(self, error_msg):
        """Handle SQL errors with helpful suggestions"""
        # Remove the automatically added LIMIT clause and comment from the error message
        # This makes the error message cleaner and less confusing for the user.
        clean_error_msg = re.sub(r" LIMIT \d+\s*--.*$", "", error_msg, flags=re.IGNORECASE)
        clean_error_msg = re.sub(r" LIMIT \d+", "", clean_error_msg, flags=re.IGNORECASE) # Catch cases without the comment

        # Specific handling for "no such table"
        if "no such table" in clean_error_msg.lower():
            match = re.search(r"no such table: (.+)", clean_error_msg)
            if match:
                table_name = match.group(1).strip('"')  # Remove quotes if present
                suggestion = self.suggest_table_name(table_name)
                if suggestion:
                    clean_error_msg += f"\n\nDid you mean:\n{suggestion}"
                else:
                    clean_error_msg += "\n\nNo similar table names found."

        # Specific handling for syntax errors (can be more detailed if needed)
        elif "syntax error" in clean_error_msg.lower():
            clean_error_msg += "\n\nPlease check your SQL syntax."

        self.show_error("SQL Error", clean_error_msg) # Use the cleaned message
        self.result_status_var.set("Query failed")
        self.current_results = None  # Clear results on error


    def suggest_table_name(self, wrong_name):
        """Suggest similar table names based on loaded tables"""
        all_tables = list(self.table_mapping.keys())

        # Prioritize exact case-insensitive matches
        suggestions = [t for t in all_tables if t.lower() == wrong_name.lower()]
        if suggestions:
            return "\n- " + "\n- ".join(sorted(suggestions))

        # Then, partial case-insensitive matches
        suggestions = [t for t in all_tables if wrong_name.lower() in t.lower()]
        if suggestions:
            return "\n- " + "\n- ".join(sorted(suggestions))

        # Finally, fuzzy matching on parts of the name
        wrong_parts = re.split(r'[._\s]', wrong_name.lower())
        fuzzy_suggestions = set()
        for table in all_tables:
            table_parts = re.split(r'[._\s]', table.lower())
            if any(part in table_parts for part in wrong_parts if len(part) > 2):  # Only consider parts > 2 chars
                fuzzy_suggestions.add(table)

        if fuzzy_suggestions:
            return "\n- " + "\n- ".join(sorted(list(fuzzy_suggestions)))

        return None  # No suggestions found

    def show_error(self, title, message):
        """Show formatted error message in a messagebox"""
        messagebox.showerror(title, message)
        self.root.lift()  # Bring the main window to the front after showing error

    def enable_spool(self, file_path):
        """Enable spooling to a file"""
        try:
            self.spool_file = open(file_path, 'w', encoding='utf-8')
            self.spooling_active = True
            return True
        except Exception as e:
            messagebox.showerror("Spool Error", f"Cannot open file {file_path}:\n{str(e)}")
            return False

    def disable_spool(self):
        """Disable spooling"""
        if hasattr(self, 'spool_file') and self.spool_file:
            self.spool_file.close()
        self.spooling_active = False

    def write_to_spool(self, content):
        """Write content to spool file if active"""
        if getattr(self, 'spooling_active', False):
            self.spool_file.write(content)
            self.spool_file.flush()

    def _get_query_to_execute(self):
        """Get either selected text or full query text"""
        try:
            if self.query_text.tag_ranges(tk.SEL):
                return self.query_text.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
        except tk.TclError:
            pass
        return self.query_text.get("1.0", tk.END).strip()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSQLApp(root)
    root.mainloop()
