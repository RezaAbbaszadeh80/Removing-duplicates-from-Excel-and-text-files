# install packages :

# pip install keyboard
# pip install pandas
# pip install pyinstaller
# pip install xlrd

# ============================================= import packages ======================================================================

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import keyboard
import os
import platform
import re 
import numpy as np 
#================================================= Make_root ==============================================================================

root = tk.Tk()
root.geometry('1200x600')

root.title('Removing duplicates from Excel and text files')

# Create a widget to display data from Excel and duplicate values
text_widget_excel = tk.Text(root, width=40)
text_widget_duplicates = tk.Text(root, width=40, undo=True)

text_widget_excel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
text_widget_duplicates.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

text_widget_excel.config(bg='#E0FFFF')
text_widget_duplicates.config(bg='#E0FFFF')

#=================================== Open_file_fuction_and_found_same_values =================================================================

import re

def clean_text(text):
    
    if isinstance(text, str):
        # Remove parentheses and similar symbols
        text = re.sub(r'[\[\]{}()]', '', text)
        
       # Remove special characters
        text = re.sub(r'[@!%$#*&^/|,\\]', '', text)
        
        # Remove duplicate words
        text = ' '.join(dict.fromkeys(text.split()))
    
    else:
       # If the text is not a string, convert it to a string
        text = str(text)
    
    return text


def open_file_dialog():
    global list_of_dicts
    
    # Clear the list_of_dicts before reading the new file
    list_of_dicts = []
    
    # Open a file dialog to select an Excel file with the .xlsx, .xls, or .xlsm extension
    file_path = filedialog.askopenfilename(filetypes=[
        ("Excel file", "*.xlsx"),
        ("Excel file", "*.xls"),
        ("Excel file", "*.xlsm"),
        ("Text Files", "*.txt")
    ])
    
    # The rest of your code to read and display the file remains unchanged

    # Check if a file was selected
    if file_path:
        
        global name_of_file 
        name_of_file = os.path.basename(file_path).split('.')

        if file_path.endswith((".xlsx", ".xls", ".xlsm")):
                
            # Read the selected Excel file into a pandas DataFrame
            try:
                data = pd.read_excel(file_path)
                
            except :
                pass

            # Convert the DataFrame to a list of dictionaries
            list_of_dicts = data.to_dict(orient='records') # It doesn't include the row numbers, making it cleaner

            # Clear the content of the text widgets and display the data and duplicate values
            text_widget_excel.delete(1.0, tk.END)
            text_widget_duplicates.delete(1.0, tk.END)

            text_widget_excel.insert(tk.END, "Data from the Excel file:\n")
            for item in list_of_dicts:
                for key, value in item.items():
                    text_widget_excel.insert(tk.END, f"{key}: {value}\n")

                # Clean the text columns using the clean_text function
                text_columns = data.select_dtypes(include=['object']).columns
                data[text_columns] = data[text_columns].applymap(clean_text)

            # Apply the cleanup function to all data
            data = data.applymap(lambda x: clean_text(x) if isinstance(x, str) else x)
            
            # ایجاد لیست unique_dict_list
            unique_dict_list = [dict(t) for t in {tuple(d.items()) for d in data.to_dict(orient='records')}]
            
            unique_data_df = pd.DataFrame(unique_dict_list)
            unique_data_df.dropna(inplace=True)
            unique_data_df.drop_duplicates()
                 
            # Display the data inside the DataFrame in the text widget
            text_widget_duplicates.delete(1.0, tk.END)
            text_widget_duplicates.insert(tk.END, "Data without duplicates from the Excel file:\n")

            # Create a dictionary to count the occurrences of each value
            value_counts = {}
            for item in unique_dict_list:
                for value in item.values():
                    if value in value_counts:
                        value_counts[value] += 1
                    else:
                        value_counts[value] = 1

            # Display only one value of each repeated value
            for key, value in value_counts.items():
                text_widget_duplicates.insert(tk.END, f"{key}: {value}\n")
                

            column_names = unique_data_df.columns.tolist()
            
            for index, row in unique_data_df.iterrows():
                for col_name, value in row.items():
                    # unique_data_df[value] = unique_data_df[value].drop_duplicates()
                    
                    if value == 'nan':
                        unique_data_df.at[index, col_name] = np.nan
                        unique_data_df.dropna(how='all', inplace=True)    
            
            # Detect duplicate rows and replace duplicate values ​​with NaN
            for col_name in column_names:
                unique_data_df[col_name] = unique_data_df[col_name].drop_duplicates()

            # Save on the original file itself
            # with pd.ExcelWriter(f"{file_path}", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            #     unique_data_df.to_excel(writer, index=False)

            # Save cleaned data to a new file
            with pd.ExcelWriter(f"clean_{name_of_file[0]}.{name_of_file[1]}", engine='openpyxl') as writer:
                unique_data_df.to_excel(writer, index=False)
            
            messagebox.showinfo('Delete duplicates Excel file ','Deleting duplicates of the Excel file was done successfully')
            

        elif file_path.endswith(".txt"):
            try:
                # Clear the content of the text widgets and display the data and duplicate values
                text_widget_excel.delete(1.0, tk.END)
                text_widget_duplicates.delete(1.0, tk.END)
                
                with open(file_path, 'r', encoding='utf-8') as file:
                    text = file.readlines()  # We receive the text lines as a list
                    
                    text_widget_excel.insert(tk.END, "Data from the Excel file:\n")
                    text_widget_excel.insert(tk.END, ''.join(text)) # We convert the text into a string using join and display it in the widget
                    
                    text_widget_duplicates.insert(tk.END,'Data without duplicates from the Text file:\n')
                    # Clean and update each line of text
                    cleaned_lines = []
                    for line in text:
                        cleaned_line = clean_text(line.strip())  # Clear each line and remove extra spaces at the beginning and end of the line
                        cleaned_lines.append(cleaned_line + '\n')  # Add a space to the end of each cleared line
                        text_widget_duplicates.insert(tk.END, cleaned_line + '\n')  # Display the cleared text in the widget
                        
                    # Save cleaned text back to the same file
                    with open(f"clean_{name_of_file[0]}.{name_of_file[1]}", 'w', encoding='utf-8') as file:
                        file.writelines(cleaned_lines)  # Write the cleaned text to the file
                    
                    messagebox.showinfo('Delete duplicates Text file ','Deleting duplicates of the Text file was done successfully')
                        
            except Exception as e:
                print("Error:", e)
                
            
btn_upload = tk.Button(root, text='Upload Excel or Text file', command=open_file_dialog)
btn_upload.pack(pady=5)


#==================================================== Entry_for_filter =================================================================

def apply_custom_filter():

    global list_of_dicts  # Using the global variable

   # Getting the search pattern from the entry
    filter_text = entry_widget.get()

    # Delete the content of the previous text
    text_widget_duplicates.delete(1.0, tk.END)
    text_widget_duplicates.insert(tk.END, "Filtered values:\n")

    # Search and display similar
    for item in list_of_dicts:
        for key, value in item.items():
            if filter_text in str(value):
                text_widget_duplicates.insert(tk.END, f"{key}: {value}\n")

# Create an Entry widget for filtering values
entry_widget = tk.Entry(root, exportselection=False)


#==================================================== Function_process =================================================================

# Function to copy selected text in the result box

def copy_text():
    selected_text = text_widget_duplicates.tag_ranges(tk.SEL)
    if selected_text:
        copied_text = text_widget_duplicates.get(tk.SEL_FIRST, tk.SEL_LAST)
        root.clipboard_clear()
        root.clipboard_append(copied_text)

# Function to cut selected text in the result box
def cut_text():
    selected_text = text_widget_duplicates.tag_ranges(tk.SEL)
    if selected_text:
        global cutted_text
        copied_text = text_widget_duplicates.get(tk.SEL_FIRST, tk.SEL_LAST)
        root.clipboard_clear()
        root.clipboard_append(copied_text)
        cutted_text = copied_text
        text_widget_duplicates.delete(tk.SEL_FIRST, tk.SEL_LAST)
        text_widget_duplicates.edit_separator()

# Function to paste text into the result box
def paste_text():
    clipboard_text = root.clipboard_get()
    if clipboard_text:
        selected_text = text_widget_duplicates.tag_ranges(tk.SEL)
        if selected_text:
            text_widget_duplicates.delete(tk.SEL_FIRST, tk.SEL_LAST)  # Delete the selected text
        text_widget_duplicates.insert(tk.INSERT, clipboard_text)

# Function to undo changes in the result box
def undo_text():
    try:
        text_widget_duplicates.edit_undo()
    except tk.TclError:
        pass
    
# Function to select all text in the result box
def select_all_text(event=None):
    text_widget_duplicates.tag_add(tk.SEL, "1.0", tk.END)
    text_widget_duplicates.mark_set(tk.SEL_FIRST, "1.0")
    text_widget_duplicates.mark_set(tk.SEL_LAST, tk.END)


# Create a context menu for the Entry widget for filtering
context_menu_entry = tk.Menu(root, tearoff=0)

# Function to cut selected text in the Entry widget
def cut_entry():
    try:
        selected_text = entry_widget.selection_get(selection="CLIPBOARD")
        if selected_text:
            root.clipboard_clear()
            root.clipboard_append(selected_text)
            entry_widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
    except tk.TclError:
        pass


# Function to copy selected text in the Entry widget
def copy_entry():
    selected_text = entry_widget.selection_get()
    if selected_text:
        root.clipboard_clear()
        root.clipboard_append(selected_text)

# Function to paste text into the Entry widget
def paste_entry():
    clipboard_text = root.clipboard_get()
    if clipboard_text:
        entry_widget.insert(tk.INSERT, clipboard_text)


#==================================================== Menu_key_functions =================================================================

# Create a context menu for the result box
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Cut", command=cut_text)
context_menu.add_command(label="Copy", command=copy_text)
context_menu.add_command(label="Paste", command=paste_text)
context_menu.add_command(label="Undo", command=undo_text)
context_menu.add_command(label="Select all", command=select_all_text)

# Bind the context menu to the result box's right-click event
text_widget_duplicates.bind("<Button-3>", lambda e: context_menu.post(e.x_root, e.y_root))

# Create menu for right click filter input
context_menu_entry.add_command(label="Cut", command=cut_entry)
context_menu_entry.add_command(label="Copy", command=copy_entry)
context_menu_entry.add_command(label="Paste", command=paste_entry)

# Bind the context menu to the Entry widget's right-click event
entry_widget.bind("<Button-3>", lambda event: context_menu_entry.post(event.x_root, event.y_root))

#============================================== Add_key_functions_to_keyboard ===============================================================

# Add keyboard shortcuts for undo, and select all

system = platform.system()

if system == "Windows":
    keyboard.add_hotkey('ctrl+z', undo_text)
    keyboard.add_hotkey('ctrl+a', select_all_text)

#================================================= Highlight_line ====================================================================

# Variable to save the highlight state of the lines
highlighted_lines = set()

def highlight_line(event):
    # Get the clicked location
    clicked_index = text_widget_duplicates.index(tk.CURRENT)
    
    # Parsing the clicked location into rows and columns
    line, column = map(int, clicked_index.split('.'))
    
    # Detection of the clicked line number
    clicked_line = line
    
    if clicked_line in highlighted_lines:
        # If the line was already highlighted, un-highlight that line
        text_widget_duplicates.tag_remove("highlighted", f"{clicked_line}.0", f"{clicked_line + 1}.0")
        highlighted_lines.remove(clicked_line)
    else:
        # Otherwise, apply line highlighting
        text_widget_duplicates.tag_add("highlighted", f"{clicked_line}.0", f"{clicked_line + 1}.0")
        highlighted_lines.add(clicked_line)
        text_widget_duplicates.tag_config("highlighted", background="cyan")

#================================================= main_loop ====================================================================

root.mainloop()

