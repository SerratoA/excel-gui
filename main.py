import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import datetime
import csv
import openpyxl
import pandas as pd


#Need to add
#Flag for delete
#Multiple row selection for delete and edit
#undo and redo?
#Real time changes 


#Create log file
log_file = "history_log.txt" 
username = ""  # Global variable to store the username
path = "data.xlsx"

def convertCSVtoExcel(csv_path, excel_path):
    # Read the CSV file into a pandas DataFrame
    df = pd.read_csv(csv_path)

    # Write the DataFrame to an Excel file
    df.to_excel(excel_path, index=False)

csv_path = "data.csv"
excel_path = "data.xlsx"


def convertExceltoCSV(excel_path, csv_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(excel_path)

    # Write the DataFrame to a CSV file
    df.to_csv(csv_path, index=False)

def getUsername():
    global username
    username = simpledialog.askstring("Username", "Please enter your username:")
    if username:
        root.title("Data Management App - User: " + username)
    else:
        root.destroy()


#Load data from Excel Sheet into Treeview
def loadData():
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    treeview.delete(*treeview.get_children())
    
    list_values = list(sheet.values)
    print(list_values)
        
    # Insert the data rows into the Treeview
    for row in list_values[1:]:
        treeview.insert("", "end", values=row)
        
def showSearchResults(results):
    search_window = tk.Toplevel(root)
    search_window.title("Search Results")

    # Create a Treeview in the search window to display the results
    search_treeview = ttk.Treeview(search_window)
    search_treeview.pack()

    # Configure the Treeview columns
    columns = ["Name", "Age", "Subscription", "Employment"]
    search_treeview["columns"] = columns
    search_treeview["show"] = "headings"
    for col in columns:
        search_treeview.heading(col, text=col)
        search_treeview.column(col, width=100)

    # Insert the search results into the Treeview
    for result in results:
        search_treeview.insert("", "end", values=result)

    def copySelectedRow():
        selected_item = search_treeview.focus()
        if selected_item:
            item_values = search_treeview.item(selected_item, "values")
            if item_values:
                name_entry.delete(0, "end")
                name_entry.insert(0, item_values[0])
                age_entry.delete(0, "end")
                age_entry.insert(0, item_values[1])
                status_combobox.set(item_values[2])
                if item_values[3] == "Employed":
                    checkbutton.state(["selected"])
                    a.set(1)
                else:
                    checkbutton.state(["!selected"])
                    a.set(0)


    # Add a "Copy" button
    copy_button = ttk.Button(search_window, text="Copy", command=copySelectedRow)
    copy_button.pack()

def performSearch():
    search_window = tk.Toplevel(root)
    search_window.title("Search")
    
    search_label = ttk.Label(search_window, text="Enter search query:")
    search_label.pack()
    
    search_entry = ttk.Entry(search_window)
    search_entry.pack()
    
    search_button = ttk.Button(search_window, text="Search", command=lambda: searchData(search_entry.get()))
    search_button.pack()

def searchData(x):
    search_text = x  # Get the search text from an entry widget

    # Collect the search results
    results = []
    for item_id in treeview.get_children():
        item_values = treeview.item(item_id)['values']
        if search_text.lower() in [str(value).lower() for value in item_values]:
            results.append(item_values)

    # Show the search results in a pop-up window
    if results:
        showSearchResults(results)
    else:
        messagebox.showinfo("No Results", "No matching results found.")




#Insert row into Excel Sheet and Treeview and clear entry widgets
def insertRow():
    #Get values from entry widgets
    name = name_entry.get()
    age = age_entry.get()
    status = status_combobox.get()
    employment = "Employed" if a.get() else "Unemployed"

# Check if name or age is empty
    if name == "":
        messagebox.showerror("Error", "Name field cannot be empty.")
        return
    if age == "":
        messagebox.showerror("Error", "Age field cannot be empty.")
        return

    # Check if age is a valid integer
    try:
        age = int(age)
    except ValueError:
        messagebox.showerror("Error", "Age must be a valid integer.")
        return

    # Insert row into Excel Sheet
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        row_values = [name, age, status, employment]
        sheet.append(row_values)
        workbook.save(path)

    except Exception as e:
        messagebox.showerror("Error", str(e))
        return


    #Insert row into Treeview
    treeview.insert("", "end", values=row_values)

    #Clear entry widgets
    name_entry.delete(0, 'end')
    name_entry.insert(0, 'Name')
    age_entry.delete(0, 'end')
    age_entry.insert(0, 'Age')
    status_combobox.current(0)
    checkbutton.state(['!selected'])

    addHistoryEntry("Inserted row: " + str(row_values))

    # Function to delete selected row from Excel Sheet and Treeview
def deleteRow():
    selected_item = treeview.focus()
    if selected_item:
        confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this row from the data?")
        if confirm:
            item_values = treeview.item(selected_item, "values")
            if item_values:
            # Delete row from Excel Sheet

                workbook = openpyxl.load_workbook(path)
                sheet = workbook.active
                row_index = int(treeview.index(selected_item))
                addHistoryEntry("Deleted row: " + str(treeview.item(selected_item)['values']))
                sheet.delete_rows(row_index + 2)  # Adding 2 to compensate for header row and 0-based indexing
                workbook.save(path)

                # Delete row from Treeview
                treeview.delete(selected_item)
                addHistoryEntry("Deleted row: " + str(treeview.item(selected_item)['values']))

    else:
        messagebox.showinfo("No Row Selected", "Please select a row to delete.")


# Edit selected row in Excel Sheet and Treeview
def editRow():
    selected_item = treeview.focus()
    if selected_item:
        item_values = treeview.item(selected_item, "values")
            # Open a new window for editing the row
        edit_window = tk.Toplevel(root)
        edit_window.title("Edit Row")
        row_index = int(treeview.index(selected_item))
            # Create labels and entry widgets for editing the row
        labels = cols
        entries = []
        for i, label in enumerate(labels):
            ttk.Label(edit_window, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            entry = ttk.Entry(edit_window, width=20)
            entry.insert(0, item_values[i])
            entry.grid(row=i, column=1, padx=5, pady=5)
            entries.append(entry)
            
            # Save button to update the row in Excel Sheet and Treeview
            def saveChanges():
                new_values = [entry.get() for entry in entries]
                try:
                    workbook = openpyxl.load_workbook(path)
                    sheet = workbook.active

                    # Delete the old row
                    sheet.delete_rows(row_index + 2)

                    # Insert the updated row at the same position
                    sheet.insert_rows(row_index + 2)
                    for col_index, value in enumerate(new_values):
                        #print(col_index)
                        sheet.cell(row=row_index + 2, column=col_index+1).value = value

                    workbook.save(path)

                    # Update the row in Treeview
                    treeview.item(selected_item, values=new_values)

                    edit_window.destroy()  # Close the edit window
                    # Add edit entry to history log
                    addHistoryEntry("Edited row: " + str(item_values) + " -> " + str(new_values))
                except Exception as e:
                    print(e)
            
            ttk.Button(edit_window, text="Save", command=saveChanges).grid(row=len(labels), column=0, columnspan=2, padx=5, pady=10)
            
    else:
        messagebox.showinfo("No Row Selected", "Please select a row to edit.")
    
# Function to copy the contents of a row
def copyRow():
    selected_item = treeview.focus()
    if selected_item:
        item_values = treeview.item(selected_item, "values")
        if item_values:
            name_entry.delete(0, "end")
            name_entry.insert(0, item_values[0])
            age_entry.delete(0, "end")
            age_entry.insert(0, item_values[1])
            status_combobox.set(item_values[2])
            if item_values[3] == "Employed":
                checkbutton.state(["selected"])
                a.set(1)
            else:       
                checkbutton.state(["!selected"])
                a.set(0)
        else:
            messagebox.showinfo("No Row Selected", "Please select a row to copy.")
    else:
        messagebox.showinfo("No Row Selected", "Please select a row to copy.")

def addHistoryEntry(entry):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"{timestamp} - {username}: {entry}\n"

    with open(log_file, "a") as file:
        file.write(log_entry)
        
# Function to open the log window
#maybe write to a file instead of a text box??
def openLogWindow():
    log_window = tk.Toplevel(root)
    log_window.title("History Logs")
    log_window.geometry("600x400")

    # Create a Text widget in the log window to display the history logs
    log_text = tk.Text(log_window, height=35, width=90)
    log_text.pack()

    # Open the log file and populate the Text widget with its contents
    with open(log_file, "r") as file:
        logs = file.read()
        log_text.insert(tk.END, logs)

def exitApp():
    root.quit()



#GUI Setup
root = tk.Tk()
root.geometry("1600x800")
log_window = None
getUsername()


menubar = tk.Menu(root)
root.config(menu=menubar)

gui_menu = tk.Menu(menubar, tearoff=False)
history_menu = tk.Menu(menubar, tearoff=False)

menubar.add_cascade(label="Menu", menu=gui_menu)
menubar.add_cascade(label = "History", menu = history_menu)

# Add search option to the search menu
gui_menu.add_command(label="Search", command=performSearch)
#separator
gui_menu.add_separator()

gui_menu.add_command(label="Exit", command=exitApp)

history_menu.add_command(label="View History", command=openLogWindow)


#Style for Tkinter
root.tk.call('source', 'forest-dark.tcl')
ttk.Style().theme_use('forest-dark')

#Title and Frame
frame = ttk.Frame(root, cursor = 'arrow')
frame.pack()

#Widgets on left side of GUI
widgets_entry = ttk.LabelFrame(frame, text='Insert Data Row')
widgets_entry.grid(column=0, row=0, sticky='nsew', padx=10, pady=10)

def create_entry_widget(parent, row, column, width, default_text):
    entry = ttk.Entry(parent, width=width)
    entry.insert(0, default_text)
    entry.bind('<FocusIn>', lambda event: entry.delete(0, 'end'))
    entry.grid(column=column, row=row, padx=5, pady=5, sticky='ew')
    return entry

# Function to create and configure a checkbutton widget
def create_checkbutton_widget(parent, row, column, text, variable):
    checkbutton = ttk.Checkbutton(parent, text=text, variable=variable, onvalue=True, offvalue=False)
    checkbutton.grid(column=column, row=row, padx=5, pady=5, sticky='ew')
    return checkbutton

# Function to create and configure a spinbox widget
def create_spinbox_widget(parent, row, column, from_, to, width, default_text):
    spinbox = ttk.Spinbox(parent, from_=from_, to=to, width=width)
    spinbox.insert(0, default_text)
    spinbox.bind('<FocusIn>', lambda event: spinbox.delete(0, 'end'))
    spinbox.grid(column=column, row=row, padx=5, pady=5, sticky='ew')
    return spinbox

# Function to create and configure a combobox widget
def create_combobox_widget(parent, row, column, values, width, default_index):
    combobox = ttk.Combobox(parent, values=values, state='readonly', width=width)
    combobox.current(default_index)
    combobox.grid(column=column, row=row, padx=5, pady=5, sticky='ew')
    return combobox

# Create entry widgets
custodian_entry = create_entry_widget(widgets_entry, 0, 0, 5, 'Enter Custodian')
source_dest_entry = create_entry_widget(widgets_entry, 1, 0, 5, 'Enter Source Destination')
source_file_entry = create_entry_widget(widgets_entry, 2, 0, 5, 'Enter Source File')
date_format_entry = create_entry_widget(widgets_entry, 3, 0, 5, 'Enter Date Format')
header_delimiter_entry = create_entry_widget(widgets_entry, 4, 0, 5, 'Enter Header Delimiter')
additional_delimiter_entry = create_entry_widget(widgets_entry, 6, 0, 5, 'Enter Additional Header Delimiter')
xls_sheet_name_entry = create_entry_widget(widgets_entry, 9, 0, 5, 'Enter XLS Sheet Name')
zip_file_name_entry = create_entry_widget(widgets_entry, 12, 0, 5, 'Enter Zip File Name')
filter_file_name_entry = create_entry_widget(widgets_entry, 13, 0, 5, 'Enter Filter File Name')
start_string_entry = create_entry_widget(widgets_entry, 1, 1, 5, 'Enter Start String')
first_record_identifier_entry = create_entry_widget(widgets_entry, 5, 1, 5, 'Enter First Record Identifier')
strip_leading_characters_entry = create_entry_widget(widgets_entry, 9, 1, 5, 'Enter Strip Leading Characters')
sequence_entry = create_entry_widget(widgets_entry, 10, 1, 5, 'Enter Sequence')
newColumnDate_entry = create_entry_widget(widgets_entry, 4, 2, 5, 'New Column Date')
newColumnIndex_entry = create_entry_widget(widgets_entry, 5, 2, 5, 'New Column Index')
newColumnCount_entry = create_entry_widget(widgets_entry, 6, 2, 5, 'New Column Count')
fileLabel_entry = create_entry_widget(widgets_entry, 9, 2, 5, 'File Label')
server_entry = create_entry_widget(widgets_entry, 10, 2, 5, 'Server')
config_file_entry = create_entry_widget(widgets_entry, 12, 1, 5, 'Enter Config File')
delimiter_entry = create_entry_widget(widgets_entry, 13, 1, 5, 'Enter Delimiter')
filter_value_entry = create_entry_widget(widgets_entry, 0, 2, 5, 'Enter Filter Value')
json_key_name_entry = create_entry_widget(widgets_entry, 2, 2, 5, 'Enter JSON KeyName')
snowflakeAccount_entry = create_entry_widget(widgets_entry, 12, 2, 5, 'Snowflake Account')
snowflakeAuthenticator_entry = create_entry_widget(widgets_entry, 13, 2, 5, 'Snowflake Authenticator')
snowflakeWarehouse_entry = create_entry_widget(widgets_entry, 14, 2, 5, 'Snowflake Warehouse')
snowflakeDatabase_entry = create_entry_widget(widgets_entry, 15, 2, 5, 'Snowflake Database')
snowflakeSchema_entry = create_entry_widget(widgets_entry, 16, 2, 5, 'Snowflake Schema')
snowflakeFileFormat_entry = create_entry_widget(widgets_entry, 0, 3, 5, 'Snowflake FileFormat')
storedProcedure_entry = create_entry_widget(widgets_entry, 1, 3, 5, 'Stored Procedure')
priority_entry = create_entry_widget(widgets_entry, 3, 3, 5, 'Priority')
notes_entry = create_entry_widget(widgets_entry, 4, 3, 5, 'Notes')

# Create checkbutton widgets
xls_to_csv_var = tk.BooleanVar()
xls_to_csv_checkbutton = create_checkbutton_widget(widgets_entry, 8, 0, 'Convert XLS to CSV', xls_to_csv_var)
unzip_file_var = tk.BooleanVar()
unzip_file_checkbutton = create_checkbutton_widget(widgets_entry, 11, 0, 'Unzip File', unzip_file_var)
combine_files_var = tk.BooleanVar()
combine_files_checkbutton = create_checkbutton_widget(widgets_entry, 14, 0, 'Combine Files', combine_files_var)
remove_header_trailer_var = tk.BooleanVar()
remove_header_trailer_checkbutton = create_checkbutton_widget(widgets_entry, 15, 0, 'Remove Header and Trailer', remove_header_trailer_var)
additional_eol_var = tk.BooleanVar()
additional_eol_checkbutton = create_checkbutton_widget(widgets_entry, 2, 1, 'Additional EOL', additional_eol_var)
remove_additional_eol_var = tk.BooleanVar()
remove_additional_eol_checkbutton = create_checkbutton_widget(widgets_entry, 3, 1, 'Remove Additional EOL', remove_additional_eol_var)
add_record_id_var = tk.BooleanVar()
add_record_id_checkbutton = create_checkbutton_widget(widgets_entry, 4, 1, 'Add Record ID', add_record_id_var)
flatten_file_var = tk.BooleanVar()
flatten_file_checkbutton = create_checkbutton_widget(widgets_entry, 6, 1, 'Flatten File', flatten_file_var)
add_sequence_var = tk.BooleanVar()
add_sequence_checkbutton = create_checkbutton_widget(widgets_entry, 7, 1, 'Add Sequence', add_sequence_var)
fill_with_blank_lines_var = tk.BooleanVar()
fill_with_blank_lines_checkbutton = create_checkbutton_widget(widgets_entry, 8, 1, 'Fill With Blank Lines', fill_with_blank_lines_var)
delimit_fixed_width_var = tk.BooleanVar()
delimit_fixed_width_checkbutton = create_checkbutton_widget(widgets_entry, 11, 1, 'Delimit Fixed Width', delimit_fixed_width_var)
filter_records_var = tk.BooleanVar()
filter_records_checkbutton = create_checkbutton_widget(widgets_entry, 14, 1, 'Filter Records', filter_records_var)
inverted_filter_var = tk.BooleanVar()
inverted_filter_checkbutton = create_checkbutton_widget(widgets_entry, 15, 1, 'Inverted Filter', inverted_filter_var)
escapeQuotes_var = tk.BooleanVar()
escapeQuotes_checkbutton = create_checkbutton_widget(widgets_entry, 7, 2, 'Escape Quotes', escapeQuotes_var)
insertFileControl_var = tk.BooleanVar()
insertFileControl_checkbutton = create_checkbutton_widget(widgets_entry, 8, 2, 'Insert File Control', insertFileControl_var)
deleteSourceFile_var = tk.BooleanVar()
deleteSourceFile_checkbutton = create_checkbutton_widget(widgets_entry, 11, 2, 'Delete Source File', deleteSourceFile_var)
json_scrapping_needed_var = tk.BooleanVar()
json_scrapping_needed_checkbutton = create_checkbutton_widget(widgets_entry, 1, 2, 'JSON Scrapping Needed', json_scrapping_needed_var)
add_column_delimiter_var = tk.BooleanVar()
add_column_delimiter_checkbutton = create_checkbutton_widget(widgets_entry, 3, 2, 'Add Column Delimiter', add_column_delimiter_var)
flag_entry_var = tk.BooleanVar()
flag_checkbutton = create_checkbutton_widget(widgets_entry, 5, 3, 'Deleteable File', flag_entry_var)

# Create spinbox and combobox widgets
date_position_entry = create_spinbox_widget(widgets_entry, 5, 0, 1, 100, 5, 'Enter Date Position')
num_quarters_entry = create_spinbox_widget(widgets_entry, 10, 0, 1, 100, 5, 'Enter Number of Quarters')
num_header_lines_entry = create_spinbox_widget(widgets_entry, 16, 0, 1, 100, 5, 'Enter Number of Header Lines')
num_trailer_lines_entry = create_spinbox_widget(widgets_entry, 0, 1, 1, 100, 5, 'Enter Number of Trailer Lines')
column_number_entry = create_spinbox_widget(widgets_entry, 16, 1, 1, 100, 5, 'Enter Column Number')

file_type_list = ['Fixed_Width', 'CSV', 'Pipe_Delimited', 'JSON', 'SQL']
file_type_combobox = create_combobox_widget(widgets_entry, 7, 0, file_type_list, 5, 0)
complexity_list = ['Low', 'Medium', 'High']
complexity_combobox = create_combobox_widget(widgets_entry, 2, 3, complexity_list, 5, 0)


#Insert button widget
insert_button = ttk.Button(frame, text='Insert', command = insertRow)
insert_button.grid(column=0, row=1,padx = 5, pady = 5, sticky='ew')

# Create the delete button widget
delete_button = ttk.Button(frame, text='Delete', command=deleteRow)
delete_button.grid(column=0, row=2, padx=5, pady=5, sticky='ew')

# Create the edit button widget
edit_button = ttk.Button(frame, text="Edit", command=editRow)
edit_button.grid(column=0, row=3, padx=5, pady=5, sticky="ew")

# Copy button widget
copy_button = ttk.Button(frame, text="Copy", command=copyRow)
copy_button.grid(column=0, row=4, padx=5, pady=5, sticky="ew")

#Frame for Treeview on right side of GUI
tree_frame = ttk.LabelFrame(frame, text='Data Tree')
tree_frame.grid(column=1, row=0, padx=10, pady=10, sticky='nsew')

# Scrollbar
tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
tree_scroll_y = ttk.Scrollbar(tree_frame)

tree_scroll_x.pack(side="bottom", fill="x")
tree_scroll_y.pack(side="right", fill="y")



# Treeview
cols = (
    "Custodian",
    "src_Destination_Table",
    "Source_File",
    "Date_Format",
    "Header_Delimiter",
    "Date_Position_OR_Column",
    "Additional_Delimiter",
    "File_Type",
    "XLS_to_CSV",
    "XLS_Sheet_Name",
    "Number_Of_Quarters",
    "Unzip_File",
    "Zip_File_Name",
    "Filter_File_Name",
    "Combine_Files",
    "Remove_Header_Trailer",
    "Num_Header_Lines",
    "Num_Trailer_Lines",
    "Start_String",
    "Additional_EOL",
    "Remove_Additional_EOL",
    "Add_Record_ID",
    "First_Record_Identifier",
    "Flatten_File",
    "Add_Sequence",
    "Fill_With_Blank_Lines",
    "Strip_Leading_Characters",
    "Sequence",
    "Delimit_Fixed_Width",
    "Config_File",
    "Delimiter",
    "Filter_Records",
    "Inverted_Filter",
    "Column_Number",
    "Filter_Value",
    "JSON_Scrapping_Needed",
    "JSON_KeyName",
    "Add_Column_Delimiter",
    "New_Column_Date",
    "New_Column_Index",
    "New_Column_Count",
    "Escape_Quotes",
    "Insert_File_Control",
    "File_Label",
    "Server",
    "Delete_Source_File",
    "Snowflake_Account",
    "Snowflake_Authenticator",
    "Snowflake_Warehouse",
    "Snowflake_Database",
    "Snowflake_Schema",
    "Snowflake_FileFormat",
    "Stored_Procedure",
    "Complexity",
    "Priority",
    "Notes",
)
treeview = ttk.Treeview(tree_frame, columns=cols, show="headings", height=20, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
for col in cols:
    treeview.heading(col, text=col)
    treeview.column(col, width=15, anchor='center')

treeview.pack()


tree_scroll_y.config(command=treeview.yview)
tree_scroll_x.config(command=treeview.xview)

# Load data into Treeview
loadData()

# Run GUI
root.mainloop()
