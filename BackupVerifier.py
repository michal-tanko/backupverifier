import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from functools import partial
import os
from datetime import datetime, timedelta
from dateutil import parser
from dateutil.parser import ParserError

__version__ = "2.2"

class CSVViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Backup Verifier [v{__version__}]")
        self.root.state('zoomed')
        self.df = None
        self.filter_values = {}
        self.original_data = []
        self.filtered_data = []
        self.file_path = ""
        self.highlighted_rows_count = 0
        self.sort_directions = {}

        # Top frame for the "Open CSV" button and search input
        self.top_frame = tk.Frame(self.root)
        self.top_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.open_button = tk.Button(self.top_frame, text="Open CSV", command=self.load_csv)
        self.open_button.pack(side="left")

        # Search label and input field next to Open CSV button
        self.search_label = tk.Label(self.top_frame, text="Search in all fields")
        self.search_label.pack(side="left", padx=(15, 5))

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self.top_frame, textvariable=self.search_var, width=40)  # width ~300px visually
        self.search_entry.pack(side="left")

        self.search_button = tk.Button(self.top_frame, text="Search", command=self.search_all_fields)
        self.search_button.pack(side="left", padx=(5, 0))
        self.search_entry.bind("<Return>", lambda e: self.search_all_fields())

        # Filter frame (just below the button)
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 5))

        # Treeview for the data grid
        self.tree = ttk.Treeview(self.root, columns=(), show="headings")
        self.tree.grid(row=2, column=0, sticky="nsew")

        # Scrollbars
        self.scroll_y = tk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.scroll_y.grid(row=2, column=1, sticky="ns")
        self.scroll_x = tk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.scroll_x.grid(row=3, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)

        # Make Treeview expandable
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        self.filter_comboboxes = {}
        self.tree.tag_configure('highlight_row', background="red", foreground="white")
        self.tree.tag_configure('warning_row', background="yellow", foreground="black")
        self.tree.tag_configure('amber_row', background="#FFBF00", foreground="black")  # Amber

        self.column_indexes = {}
        self.date_of_report = None
        self.set_app_icon()
        self.tree.bind("<Button-1>", self.record_cell_click)  # Left-click to record cell
        self.tree.bind("<Control-c>", self.copy_selected_cell_to_clipboard)
        self.tree.bind("<Control-C>", self.copy_selected_cell_to_clipboard)

        self.last_clicked_column = None
        self.cell_highlight = None

        # Bind scroll events to remove highlight
        self.tree.bind("<MouseWheel>", self.remove_cell_highlight)  # Vertical scroll
        self.root.bind("<Button-1>", self.on_global_click, add="+")

        self.tree.bind("<Up>", self.remove_cell_highlight)
        self.tree.bind("<Down>", self.remove_cell_highlight)
        self.tree.bind("<Left>", self.remove_cell_highlight)
        self.tree.bind("<Right>", self.remove_cell_highlight)
        self.tree.bind("<Prior>", self.remove_cell_highlight)  # Page Up
        self.tree.bind("<Next>", self.remove_cell_highlight)   # Page Down

    def on_global_click(self, event):
        # Identify region in treeview for the click coordinates relative to the treeview
        x, y = event.x_root, event.y_root
        # Translate root coords to tree coords
        tree_x = self.tree.winfo_rootx()
        tree_y = self.tree.winfo_rooty()
        rel_x = x - tree_x
        rel_y = y - tree_y

        region = self.tree.identify("region", rel_x, rel_y)
        if region != "cell":
            # Click was outside treeview cells, remove highlight
            self.remove_cell_highlight()

    def resource_path(self, relative_path): # Function to handle resource paths (icon)
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def set_app_icon(self):
        try:
            # Construct the path to the icon file using the resource_path function
            icon_path = self.resource_path("icon.ico") #Use this all the time

            # Set the icon
            self.root.iconbitmap(icon_path)  # Use iconbitmap on Windows
        except Exception as e:
            print(f"Error setting icon: {e}")

    def exit_app(self):
       self.root.destroy()

    def load_csv(self):
        self.file_path = filedialog.askopenfilename(
            title="Select CSV File", filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if self.file_path:
            try:
                self.root.title(f"Backup Verifier [v{__version__}] - {os.path.basename(self.file_path)}")
                self.df = self.read_csv_file(self.file_path)
                self.original_data = self.df.values.tolist()
                self.populate_table(self.df)
            except Exception as e:
                tk.messagebox.showerror("Error", f"Error loading CSV: {e}")

    def read_csv_file(self, file_path):
        required_columns = ["Object", "Object Type", "Object State",
                             "Protection Status", "SLA Domain", "Last Successful Backup",
                             "Latest Archival Snapshot", "Latest Replication Snapshot",
                             "Cluster", "Location"]
        try:
            df = pd.read_csv(file_path)
            missing_columns = set(required_columns) - set(df.columns)
            if missing_columns:
                raise ValueError("Missing columns: {}".format(missing_columns))
            
            # Get report date from filename (format _YYYY-MM-DD_)
            try:
                filename = os.path.basename(file_path)
                match = re.search(r"_(\d{4}-\d{2}-\d{2})_", filename)
                if not match:
                    raise ValueError(f"Filename does not contain date in format _YYYY-MM-DD_: {filename}")
                self.date_of_report = datetime.strptime(match.group(1), "%Y-%m-%d")
            except ValueError as e:
                tk.messagebox.showerror("Error", str(e))
                raise
            except Exception as e:
                tk.messagebox.showerror("Error", "Could not parse date from filename. Please contact administrator.")
                raise e
                

            # Store column indexes
            self.column_indexes["object"] = df.columns.get_loc("Object")
            self.column_indexes["object_type"] = df.columns.get_loc("Object Type")
            self.column_indexes["object_state"] = df.columns.get_loc("Object State")
            self.column_indexes["protection_status"] = df.columns.get_loc("Protection Status")
            self.column_indexes["sla_domain"] = df.columns.get_loc("SLA Domain")
            self.column_indexes["last_successful_backup"] = df.columns.get_loc("Last Successful Backup")
            self.column_indexes["latest_archival_snapshot"] = df.columns.get_loc("Latest Archival Snapshot")
            self.column_indexes["latest_replication_snapshot"] = df.columns.get_loc("Latest Replication Snapshot")
            self.column_indexes["cluster"] = df.columns.get_loc("Cluster")
            self.column_indexes["location"] = df.columns.get_loc("Location")

        except Exception as e:
            tk.messagebox.showerror("Error", "Invalid CSV file, contact administrator, please!")
            raise e

        filename = os.path.basename(file_path)
        self.is_m365 = filename.find("M365") != -1
        if filename.startswith("Daily-Backups-M365"):
            df = df.iloc[1:, :]
        return df

    # New method to handle search button click
    def search_all_fields(self):
        if self.df is None or len(self.original_data) == 0:
            messagebox.showinfo("Info", "Please load a CSV file first.")
            return

        search_text = self.search_var.get().strip().lower()
        if not search_text:
            # If search box is empty, reset filters and show all data
            self.filtered_data = self.original_data
            self.populate_table(self.df)  # Re-populate table with original dataframe
            return

        # Filter rows where any cell contains the search text (case-insensitive)
        filtered_rows = []
        for row in self.original_data:
            if any(search_text in str(cell).lower() for cell in row):
                filtered_rows.append(row)

        self.filtered_data = filtered_rows
        self.populate_filtered_data_with_tags(filtered_rows)

    # Helper function to populate treeview with filtered data with tags and highlighting
    def populate_filtered_data_with_tags(self, filtered_rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        self.highlighted_rows_count = 0
        highlighted_rows = []
        normal_rows = []
        nan_rows = []
        warning_rows = []

        for row in filtered_rows:
            if row[self.column_indexes["object_state"]] == "Active" and row[self.column_indexes["protection_status"]] == "Protected":
                backup_value = str(row[self.column_indexes["last_successful_backup"]])
                if self.is_m365 and (backup_value.lower() == "nan" or backup_value.strip() == ""):
                    nan_rows.append(row)
                else:
                    try:
                        date_of_last_backup = parser.parse(str(row[self.column_indexes["last_successful_backup"]]))
                        time_delta = self.date_of_report - date_of_last_backup
                        if self.is_m365:
                            if time_delta > timedelta(hours=24) and time_delta <= timedelta(hours=48):
                                warning_rows.append(row)
                            elif time_delta > timedelta(hours=48):
                                highlighted_rows.append(row)
                                self.highlighted_rows_count += 1
                            else:
                                normal_rows.append(row)
                        elif time_delta > timedelta(hours=24):
                            highlighted_rows.append(row)
                            self.highlighted_rows_count += 1
                        else:
                            normal_rows.append(row)
                    except ValueError:
                        highlighted_rows.append(row)
                        self.highlighted_rows_count += 1
            else:
                normal_rows.append(row)

        for row in highlighted_rows:
            self.tree.insert("", tk.END, values=row, tags=("highlight_row",))
        for row in nan_rows:
            self.tree.insert("", tk.END, values=row, tags=("amber_row",))
        for row in warning_rows:
            self.tree.insert("", tk.END, values=row, tags=("warning_row",))
        for row in normal_rows:
            self.tree.insert("", tk.END, values=row)

        if self.highlighted_rows_count > 0:
            messagebox.showwarning("Result", "Contains possible failures. Please check details.")
        else:
            messagebox.showinfo("Result", "Search completed. No issues found.")

    # Modify populate_table to set filtered_data to original_data
    def populate_table(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.clear_filters()
        self.tree["columns"] = list(df.columns)
        for index, col in enumerate(df.columns):
            self.tree.heading(col, text=col, command=partial(self.sort_column, col))
            self.tree.column(col, stretch=True)
            unique_values = [""] + sorted(list(df[col].astype(str).unique()), key=str)
            combobox = ttk.Combobox(self.filter_frame, values=unique_values, state="readonly")
            combobox.pack(side="left", fill="x", expand=True)
            self.filter_comboboxes[col] = combobox
            self.filter_values[col] = ""
            combobox.bind("<<ComboboxSelected>>", partial(self.apply_filter, col))
            if index == self.column_indexes.get("object_state", -1) and "Active" in unique_values:
                combobox.set("Active")
                self.filter_values[col] = "Active"
            if index == self.column_indexes.get("protection_status", -1) and "Protected" in unique_values:
                combobox.set("Protected")
                self.filter_values[col] = "Protected"

        self.filtered_data = self.original_data  # Reset filtered_data on loading new CSV
        self.apply_filter(None)

    def apply_filter(self, column, event=None):
        if column is not None:
            filter_text = self.filter_comboboxes[column].get()
            self.filter_values[column] = filter_text
        self.update_combobox_filters(column)
        self.filtered_data = self.filter_data()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.highlighted_rows_count = 0
        highlighted_rows = []
        normal_rows = []
        nan_rows = []
        warning_rows = []
        for row in self.filtered_data:
            if row[self.column_indexes["object_state"]] == "Active" and row[self.column_indexes["protection_status"]] == "Protected":
                backup_value = str(row[self.column_indexes["last_successful_backup"]])
                if self.is_m365 and (backup_value.lower() == "nan" or backup_value.strip() == ""):
                    tags = ("amber_row",)
                    nan_rows.append(row)
                else:
                    try:
                        date_of_last_backup = parser.parse(str(row[self.column_indexes["last_successful_backup"]]))
                        time_delta = self.date_of_report - date_of_last_backup
                        if self.is_m365:
                            if time_delta > timedelta(hours=24) and time_delta <= timedelta(hours=48):
                                tags = ("warning_row",)
                                warning_rows.append(row)
                            elif time_delta > timedelta(hours=48):
                                tags = ("highlight_row",)
                                self.highlighted_rows_count += 1
                                highlighted_rows.append(row)
                            else:
                                normal_rows.append(row)
                        elif time_delta > timedelta(hours=24):
                            tags = ("highlight_row",)
                            self.highlighted_rows_count += 1
                            highlighted_rows.append(row)
                        else:
                            normal_rows.append(row)

                    except ValueError:
                        tags = ("highlight_row",)
                        self.highlighted_rows_count += 1
                        highlighted_rows.append(row)
            else:
               normal_rows.append(row)
        
        for row in highlighted_rows:
          self.tree.insert("", tk.END, values=row, tags=("highlight_row",))
        for row in nan_rows:
          self.tree.insert("", tk.END, values=row, tags=("amber_row",))
        for row in warning_rows:
          self.tree.insert("", tk.END, values=row, tags=("warning_row",))
        for row in normal_rows:
          self.tree.insert("", tk.END, values=row)

        if event is None:
          if self.highlighted_rows_count > 0:
              messagebox.showwarning("Result", "Contains possible failures. Please check details.")
          else:
              messagebox.showinfo("Result", "Looks allright")

    def filter_data(self):
        filtered_data = []
        for row in self.original_data:
            valid = True
            for col, filter_value in self.filter_values.items():
                if filter_value != "":
                    if str(row[self.df.columns.get_loc(col)]) != filter_value:
                        valid = False
                        break
            if valid:
                filtered_data.append(row)
        return filtered_data

    def update_combobox_filters(self, changed_column):
        filtered_data = self.filter_data()
        filtered_df = pd.DataFrame(filtered_data, columns=self.df.columns)

        for col, combobox in self.filter_comboboxes.items():
            unique_values = [""]
            if not filtered_df.empty:
                 unique_values += sorted(list(filtered_df[col].astype(str).unique()), key=str)
            combobox['values'] = unique_values

    def clear_filters(self):
        for combobox in self.filter_comboboxes.values():
            combobox.destroy()
        self.filter_comboboxes = {}
        self.filter_values = {}

    def sort_column(self, col):
        if self.df is None:
            return
        
        filtered_df = pd.DataFrame(self.filtered_data, columns=self.df.columns)
        if filtered_df.empty:
            return

        # Get current sort direction for the column (True=ascending, False=descending)
        ascending = self.sort_directions.get(col, True)

        # Sort by string representation to avoid str/float comparison errors
        sort_col = filtered_df[col].astype(str)
        filtered_df = filtered_df.assign(_sort_key=sort_col).sort_values(by="_sort_key", ascending=ascending).drop(columns=["_sort_key"])
                                                        
                
                                                        
        self.filtered_data = filtered_df.values.tolist()

        # Reset all column headers text to remove arrows
        for c in self.df.columns:
            self.tree.heading(c, text=c)

        # Update the clicked column header with arrow
        arrow = "↑" if ascending else "↓"
        self.tree.heading(col, text=f"{col} {arrow}")

        # Toggle the sort direction for next time
        self.sort_directions[col] = not ascending

        # Clear existing rows in the tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Re-insert rows with proper tags (highlighting etc.)
            
        highlighted_rows = []
        normal_rows = []
        nan_rows = []
        warning_rows = []
        for row in self.filtered_data:
            if row[self.column_indexes["object_state"]] == "Active" and row[self.column_indexes["protection_status"]] == "Protected":
                backup_value = str(row[self.column_indexes["last_successful_backup"]])
                if self.is_m365 and (backup_value.lower() == "nan" or backup_value.strip() == ""):
                    tags = ("amber_row",)
                    nan_rows.append(row)
                else:
                    try:
                        date_of_last_backup = parser.parse(str(row[self.column_indexes["last_successful_backup"]]))
                        time_delta = self.date_of_report - date_of_last_backup
                        if self.is_m365:
                            if time_delta > timedelta(hours=24) and time_delta <= timedelta(hours=48):
                                tags = ("warning_row",)
                                warning_rows.append(row)
                            elif time_delta > timedelta(hours=48):
                                tags = ("highlight_row",)
                                highlighted_rows.append(row)
                            else:
                                normal_rows.append(row)
                        elif time_delta > timedelta(hours=24):
                            tags = ("highlight_row",)
                            highlighted_rows.append(row)
                        else:
                            normal_rows.append(row)
                    except ValueError:
                        tags = ("highlight_row",)
                        highlighted_rows.append(row)
            else:
                normal_rows.append(row)
        for row in highlighted_rows:
          self.tree.insert("", tk.END, values=row, tags=("highlight_row",))
        for row in nan_rows:
          self.tree.insert("", tk.END, values=row, tags=("amber_row",))
        for row in warning_rows:
          self.tree.insert("", tk.END, values=row, tags=("warning_row",))
        for row in normal_rows:
          self.tree.insert("", tk.END, values=row)

    def record_cell_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            self.last_clicked_column = self.tree.identify_column(event.x)
            row_id = self.tree.identify_row(event.y)
            col_id = self.last_clicked_column

            # Clean up previous highlight
            self.remove_cell_highlight()  # Remove previous highlight

            if row_id and col_id:
                bbox = self.tree.bbox(row_id, col_id)
                if bbox:
                    x, y, width, height = bbox
                    cell_value = self.tree.item(row_id, "values")[int(col_id.replace("#", "")) - 1]

                    self.cell_highlight = tk.Label(self.tree, text=cell_value,
                                                background="#e1f5fe",  # light blue
                                                borderwidth=2, relief="solid")
                    self.cell_highlight.place(x=x, y=y, width=width, height=height)


    def copy_selected_cell_to_clipboard(self, event=None):
        selected_item = self.tree.focus()
        if selected_item and self.last_clicked_column:
            column_index = int(self.last_clicked_column.replace("#", "")) - 1
            values = self.tree.item(selected_item, "values")
            if 0 <= column_index < len(values):
                cell_value = values[column_index]
                self.root.clipboard_clear()
                self.root.clipboard_append(cell_value)
                self.root.update()

    def remove_cell_highlight(self, event=None):
        if self.cell_highlight:
            self.cell_highlight.destroy()
            self.cell_highlight = None

if __name__ == "__main__":
    root = tk.Tk()
    app = CSVViewerApp(root)
    root.mainloop()
