import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from functools import partial
import os
from datetime import datetime, timedelta

class CSVViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Backup Verifier")
        self.root.state('zoomed')
        self.df = None
        self.filter_values = {}
        self.original_data = []
        self.filtered_data = []
        self.file_path = ""
        self.highlighted_rows_count = 0
        self.menu_bar = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Open CSV", command=self.load_csv)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.exit_app)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.root.config(menu=self.menu_bar)
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.grid(row=0, column=0, sticky="ew")
        self.tree = ttk.Treeview(self.root, columns=(), show="headings")
        self.tree.grid(row=1, column=0, sticky="nsew")
        self.scroll_y = tk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.scroll_y.grid(row=1, column=1, sticky="ns")
        self.scroll_x = tk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.scroll_x.grid(row=2, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.filter_comboboxes = {}
        self.tree.tag_configure('highlight_row', background="red", foreground="white")
        self.tree.tag_configure('warning_row', background="yellow", foreground="black")
        self.column_indexes = {}
        self.date_of_report = None
        self.set_app_icon() # adding here to run at the start

    def resource_path(self, relative_path): # Function to handle resource paths (icon)
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
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
            
            # Get file creation time
            try:
                timestamp = os.path.getctime(file_path)
                datetime_object = datetime.fromtimestamp(timestamp)
                self.date_of_report = datetime_object.strftime("%m/%d/%Y %I:%M:%S%p")
                self.date_of_report = datetime.strptime(self.date_of_report, "%m/%d/%Y %I:%M:%S%p")
            except Exception as e:
                tk.messagebox.showerror("Error", "Could not retrieve or parse file creation date. Please contact administrator.")
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

    def populate_table(self, df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.clear_filters()
        self.tree["columns"] = list(df.columns)
        for index, col in enumerate(df.columns):
            self.tree.heading(col, text=col, command=partial(self.sort_column, col))
            unique_values = [""] + sorted(list(df[col].astype(str).unique()))
            combobox = ttk.Combobox(self.filter_frame, values=unique_values, state="readonly")
            combobox.pack(side="left", fill="x", expand=True)
            self.filter_comboboxes[col] = combobox
            self.filter_values[col] = ""
            combobox.bind("<<ComboboxSelected>>", partial(self.apply_filter, col))
            self.tree.column(col, width=150, stretch=True)
            self.tree.bind("<Configure>", self._adjust_column_widths)
            self.tree.heading(col, text=col)
            if index == self.column_indexes["object_state"] and "Active" in unique_values:
                combobox.set("Active")
                self.filter_values[col] = "Active"
            if index == self.column_indexes["protection_status"] and "Protected" in unique_values:
                combobox.set("Protected")
                self.filter_values[col] = "Protected"
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
        warning_rows = []
        for row in self.filtered_data:
           tags = ()
           if row[self.column_indexes["object_state"]] == "Active" and row[self.column_indexes["protection_status"]] == "Protected":
                try:
                    date_of_last_backup = datetime.strptime(str(row[self.column_indexes["last_successful_backup"]]), "%m/%d/%Y %I:%M:%S%p")
                    
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
                 unique_values += sorted(list(filtered_df[col].astype(str).unique()))
            combobox['values'] = unique_values

    def _adjust_column_widths(self, event):
        for col in self.tree["columns"]:
            self.tree.column(col, width=max(tk.font.Font().measure(self.tree.heading(col, text=col)), self.tree.column(col, width=None)), stretch=True)

    def clear_filters(self):
        for combobox in self.filter_comboboxes.values():
            combobox.destroy()
        self.filter_comboboxes = {}
        self.filter_values = {}

    def sort_column(self, col):
        if self.df is None:
            return
        
        filtered_df = pd.DataFrame(self.filtered_data, columns=self.df.columns)

        if not filtered_df.empty:
          filtered_df = filtered_df.sort_values(by=col)
          if self.tree.heading(col, text=f"{col} ↑"):
            filtered_df = filtered_df.sort_values(by=col, ascending=False)
            self.tree.heading(col, text=f"{col} ↓")
          else:
            self.tree.heading(col, text=f"{col} ↑")
          self.filtered_data = filtered_df.values.tolist()
          
        for item in self.tree.get_children():
          self.tree.delete(item)
          
        highlighted_rows = []
        normal_rows = []
        warning_rows = []
        for row in self.filtered_data:
           tags = ()
           if row[self.column_indexes["object_state"]] == "Active" and row[self.column_indexes["protection_status"]] == "Protected":
                try:
                    date_of_last_backup = datetime.strptime(str(row[self.column_indexes["last_successful_backup"]]), "%m/%d/%Y %I:%M:%S%p")
                    
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
        for row in warning_rows:
          self.tree.insert("", tk.END, values=row, tags=("warning_row",))
        for row in normal_rows:
          self.tree.insert("", tk.END, values=row)


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVViewerApp(root)
    root.mainloop()