import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from functools import partial
import os
from datetime import datetime, timedelta

class CSVViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Backup Verifier")  # Changed Window Title
        self.root.state('zoomed')  # Start maximized

        self.df = None  # to store pandas DataFrame
        self.filter_values = {}  # Track filter values for each column
        self.original_data = []  # keeps original data
        self.filtered_data = []
        self.file_path = ""
        self.highlighted_rows_count = 0 # Track highlighted rows

        # Menu
        self.menu_bar = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Open CSV", command=self.load_csv)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.exit_app)  # Correct Exit command
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)
        self.root.config(menu=self.menu_bar)

        # Table and Filter Frame
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.grid(row=0, column=0, sticky="ew")

        self.tree = ttk.Treeview(self.root, columns=(), show="headings")
        self.tree.grid(row=1, column=0, sticky="nsew")

        self.scroll_y = tk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.scroll_y.grid(row=1, column=1, sticky="ns")

        self.scroll_x = tk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.scroll_x.grid(row=2, column=0, sticky="ew")

        self.tree.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)

        # configure rows and columns to resize
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        self.filter_comboboxes = {}  # dict of filter combobox widgets
        self.tree.tag_configure('highlight_row', background="red", foreground="white")
    
    def exit_app(self):
       self.root.destroy() # properly close the application

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
        filename = os.path.basename(file_path)
        df = pd.read_csv(file_path, usecols=range(9))  # load only first 9 columns
        if filename.startswith("Daily-Backups-M365"):
            df = df.iloc[1:, :]  # skip the first data row
        return df

    def populate_table(self, df):
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Clear filters
        self.clear_filters()

        # Set columns
        self.tree["columns"] = list(df.columns)

        for index, col in enumerate(df.columns):
            self.tree.heading(col, text=col, command=partial(self.sort_column, col))

            # Filter Input in filter_frame
            unique_values = [""] + sorted(list(df[col].astype(str).unique()))

            combobox = ttk.Combobox(self.filter_frame, values=unique_values, state="readonly")
            combobox.pack(side="left", fill="x", expand=True)

            self.filter_comboboxes[col] = combobox
            self.filter_values[col] = ""
            combobox.bind("<<ComboboxSelected>>", partial(self.apply_filter, col))
            self.tree.column(col, width=150, stretch=True)
            self.tree.bind("<Configure>", self._adjust_column_widths)
            self.tree.heading(col, text=col)
            
            # Set default value for 3rd combobox
            if index == 2 and "Active" in unique_values:
                combobox.set("Active")
                self.filter_values[col] = "Active"
            # Set default value for 5th combobox
            if index == 4 and "Protected" in unique_values:
                combobox.set("Protected")
                self.filter_values[col] = "Protected"
        
        # Trigger filter
        self.apply_filter(None)

    def insert_row_with_validation(self, row):
        tags = ()
        if len(row) > 6 and row[2] == "Active" and row[4] == "Protected":
            try:
                date_4th = datetime.strptime(str(row[3]), "%m/%d/%Y %I:%M:%S%p").date() # Cast to string before parsing
                date_7th = datetime.strptime(str(row[6]), "%m/%d/%Y %I:%M:%S%p").date()
                if str(row[5]) == "Once a week -weekend":
                  if date_7th < date_4th - timedelta(days=8): # compare with 4th - 8 days
                    tags = ("highlight_row",)
                elif date_7th < date_4th - timedelta(days=1): # compare with 4th - 1 day if it is not 'Once a week -weekend'
                  tags = ("highlight_row",)
            except ValueError:
                tags = ("highlight_row",)
        self.tree.insert("", tk.END, values=row, tags=tags)

    def apply_filter(self, column, event=None):
        if column is not None:
            filter_text = self.filter_comboboxes[column].get()
            self.filter_values[column] = filter_text  # store filter
        
        # Update filters
        self.update_combobox_filters(column)
    
        self.filtered_data = self.filter_data()
        
        # Clear data in table
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert filtered data
        self.highlighted_rows_count = 0
        highlighted_rows = []
        normal_rows = []
        for row in self.filtered_data:
           tags = ()
           if len(row) > 6 and row[2] == "Active" and row[4] == "Protected":
                try:
                    date_4th = datetime.strptime(str(row[3]), "%m/%d/%Y %I:%M:%S%p").date()
                    date_7th = datetime.strptime(str(row[6]), "%m/%d/%Y %I:%M:%S%p").date()
                    if str(row[5]) == "Once a week -weekend":
                      if date_7th < date_4th - timedelta(days=8):
                        tags = ("highlight_row",)
                        self.highlighted_rows_count += 1 # Increment highlighted row counter
                        highlighted_rows.append(row)
                      else:
                          normal_rows.append(row)
                    elif date_7th < date_4th - timedelta(days=1):
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
        
        for row in highlighted_rows: # insert highlighted rows first
          self.tree.insert("", tk.END, values=row, tags=("highlight_row",))
        for row in normal_rows:
          self.tree.insert("", tk.END, values=row)
        # Display popup message after loading
        if event is None: # check if apply_filter was not triggered by the user
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
            combobox['values'] = unique_values # assign the data to the combobox

    def _adjust_column_widths(self, event):
        for col in self.tree["columns"]:
            self.tree.column(col, width=max(tk.font.Font().measure(self.tree.heading(col, text=col)), self.tree.column(col, width=None)), stretch=True)

    def clear_filters(self):
        # destroy any widgets associated with filter
        for combobox in self.filter_comboboxes.values():
            combobox.destroy()
        self.filter_comboboxes = {}
        self.filter_values = {}

    def sort_column(self, col):
        if self.df is None:
            return
        
        filtered_df = pd.DataFrame(self.filtered_data, columns=self.df.columns)

        if not filtered_df.empty:
          # Sort the data
          filtered_df = filtered_df.sort_values(by=col)
          if self.tree.heading(col, text=f"{col} ↑"):
            filtered_df = filtered_df.sort_values(by=col, ascending=False)
            self.tree.heading(col, text=f"{col} ↓")
          else:
            self.tree.heading(col, text=f"{col} ↑")
          self.filtered_data = filtered_df.values.tolist()
          
        # Clear data in table
        for item in self.tree.get_children():
          self.tree.delete(item)
          
        # Insert filtered data
        highlighted_rows = []
        normal_rows = []
        for row in self.filtered_data:
           tags = ()
           if len(row) > 6 and row[2] == "Active" and row[4] == "Protected":
                try:
                    date_4th = datetime.strptime(str(row[3]), "%m/%d/%Y %I:%M:%S%p").date()
                    date_7th = datetime.strptime(str(row[6]), "%m/%d/%Y %I:%M:%S%p").date()
                    if str(row[5]) == "Once a week -weekend":
                      if date_7th < date_4th - timedelta(days=8):
                        tags = ("highlight_row",)
                        highlighted_rows.append(row)
                      else:
                        normal_rows.append(row)
                    elif date_7th < date_4th - timedelta(days=1):
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
        for row in normal_rows:
          self.tree.insert("", tk.END, values=row)


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVViewerApp(root)
    root.mainloop()
