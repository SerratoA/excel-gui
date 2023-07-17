import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import openpyxl
import datetime
import csv


#Need to add
#Flag for delete
#Multiple row selection for delete and edit
#undo and redo?
#Real time changes 


#Create log file
log_file = "history_log.txt" 
username = ""  # Global variable to store the username


def getUsername():
    global username
    username = simpledialog.askstring("Username", "Please enter your username:")
    if username:
        root.title("Data Management App - User: " + username)
    else:
        root.destroy()


#Load data from Excel Sheet into Treeview
def loadData():
    path = "data.csv"
    with open(path, "r") as file:
        reader = csv.reader(file)
        header = next(reader)  # Read the header row
        
        treeview.delete(*treeview.get_children())
        
        # Configure the Treeview columns
        treeview["columns"] = header
        for col in header:
            treeview.heading(col, text=col)
            treeview.column(col, width=100)
        
        # Insert the data rows into the Treeview
        for row in reader:
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


def searchData():
    search_text = search_entry.get()  # Get the search text from an entry widget

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
        path = "people.xlsx"
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
                path = "people.xlsx"
                workbook = openpyxl.load_workbook(path)
                sheet = workbook.active
                row_index = int(treeview.index(selected_item))
                addHistoryEntry("Deleted row: " + str(treeview.item(selected_item)['values']))
                sheet.delete_rows(row_index + 2)  # Adding 2 to compensate for header row and 0-based indexing
                workbook.save(path)

                # Delete row from Treeview
                treeview.delete(selected_item)

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
        labels = ["Name", "Age", "Subscription", "Employment"]
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
                    path = "people.xlsx"
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
    log_text = tk.Text(log_window, height=35, width=50)
    log_text.pack()

    # Open the log file and populate the Text widget with its contents
    with open(log_file, "r") as file:
        logs = file.read()
        log_text.insert(tk.END, logs)



#GUI Setup
root = tk.Tk()
log_window = None
getUsername()



#Style for Tkinter
root.tk.call('source', 'forest-dark.tcl')
ttk.Style().theme_use('forest-dark')

#Title and Frame
frame = ttk.Frame(root, cursor = 'arrow')
frame.pack()

#Widgets on left side of GUI
widgets_entry = ttk.LabelFrame(frame, text='Insert Data Row')
widgets_entry.grid(column=0, row=0, sticky='nsew', padx=20, pady=10)

#Name entry widget
name_entry = ttk.Entry(widgets_entry, width=20)
name_entry.insert(0, 'Name')
name_entry.bind('<FocusIn>', lambda event: name_entry.delete(0, 'end'))
name_entry.grid(column=0, row=0, padx = 5, pady = 5, sticky='ew')

#Age entry widget
age_entry = ttk.Spinbox(widgets_entry, from_=1, to=100, width=5)
age_entry.insert(0, 'Age')
age_entry.bind('<FocusIn>', lambda event: age_entry.delete(0, 'end'))
age_entry.grid(column=0, row=1,padx = 5, pady = 5, sticky='ew')

#Status combobox widget
status_list = ['Active', 'Inactive']
status_combobox = ttk.Combobox(widgets_entry, values = status_list, state='readonly')
status_combobox.current(0)
status_combobox.grid(column=0, row=2, padx = 5, pady = 5, sticky='ew')

#Employment checkbutton widget
a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_entry, text='Employment', variable=a)
checkbutton.grid(column=0, row=3, padx = 5, pady = 5,  sticky='ew')

#Insert button widget
insert_button = ttk.Button(widgets_entry, text='Insert', command = insertRow)
insert_button.grid(column=0, row=4,padx = 5, pady = 5, sticky='ew')

# Create the delete button widget
delete_button = ttk.Button(widgets_entry, text='Delete', command=deleteRow)
delete_button.grid(column=0, row=5, padx=5, pady=5, sticky='ew')

# Create the edit button widget
edit_button = ttk.Button(widgets_entry, text="Edit", command=editRow)
edit_button.grid(column=0, row=6, padx=5, pady=5, sticky="ew")

# Copy button widget
copy_button = ttk.Button(widgets_entry, text="Copy", command=copyRow)
copy_button.grid(column=0, row=7, padx=5, pady=5, sticky="ew")

# Create the check logs button widget
check_logs_button = ttk.Button(frame, text="Check Logs", command=openLogWindow)
check_logs_button.grid(column=1, row=2, padx=5, pady=5)

# Create a Text widget for history log
history_log = tk.Text(frame, height=10, width=50)
history_log.grid(column=0, row=6, columnspan=2, padx=20, pady=10, sticky="nsew")

#Separator
separator = ttk.Separator(widgets_entry, orient='horizontal')
separator.grid(column=0, row=8, padx=10, pady=15, sticky='nsew')

# Search button widget
search_button = ttk.Button(widgets_entry, text='Search', command=searchData)
search_button.grid(column=0, row=10, padx=5, pady=5, sticky='ew')

# Search entry widget
search_entry = ttk.Entry(widgets_entry, width=20)
search_entry.grid(column=0, row=9, padx=5, pady=5, sticky='ew')


#Frame for Treeview on right side of GUI
tree_frame = ttk.LabelFrame(frame, text='Data Tree')
tree_frame.grid(column=1, row=0, padx= 10, pady=10, sticky='nsew')

# Scrollbar
tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
tree_scroll_y = ttk.Scrollbar(tree_frame)

tree_scroll_x.pack(side="bottom", fill="x")
tree_scroll_y.pack(side="right", fill="y")

# Canvas for horizontal scrollbar
tree_canvas = tk.Canvas(tree_frame, bd=0, xscrollcommand=tree_scroll_x.set)
tree_canvas.pack(side="left", fill="both", expand=True)

tree_scroll_x.config(command=tree_canvas.xview)
tree_scroll_y.config(command=tree_canvas.yview)

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
treeview = ttk.Treeview(tree_canvas, columns=cols, show="headings", height=35, yscrollcommand=tree_scroll_y.set)

treeview.pack()

# Configure column widths
for col in cols:
    treeview.column(col, width=100)

tree_scroll_y.config(command=treeview.yview)

# Load data into Treeview
loadData()

# Run GUI
root.mainloop()
