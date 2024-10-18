import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from AI import vectorizer, model
from datetime import datetime

# Define global variable for the DataFrame and checkbox states
df = None
checkbox_states = {}  # To store checkbox states for each row
current_view = "full"  # To keep track of current view (full or filtered)
# Global variables to store selected month and year
selected_month = None
selected_year = None

# Define categories
categories = [
    "ไฟฟ้า", "ประปา", "Air", "ปลวก", "สถาบัน", "Fur", "อุปกรณ์",
    "รายรับอื่นๆ", "ขาย", "วัตถุดิบ", "หมู", "ไก่ไข่", "เนื้อ",
    "วัชระ+สุรพล", "ลูกชิ้น", "กุ้ง", "ผัก", "เครื่องดื่ม", "ไอติม",
    "ขนมจีบ", "ของหวาน", "ผลไม้", "อุปกรณ์ใช้สอย", "วัสดุสิ้นเปลือง",
    "ค่าใช้สอย", "PR", "ค่าไฟ", "ค่าแรง", "Ptime", "สวัสดิการ", "Manage"
]

def import_file():
    global df, checkbox_states, file_path
    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ Excel เพื่อโหลด",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if file_path:
        try:
            df = pd.read_excel(file_path)

            # Handle missing columns
            if 'Selected' not in df.columns:
                df['Selected'] = ["❌" for _ in range(len(df))]
            if 'Category' not in df.columns:
                df['Category'] = [""] * len(df)
            
            if 'วันที่' in df.columns:
                # Convert to datetime and handle errors
                df['วันที่'] = pd.to_datetime(df['วันที่'], errors='coerce')

                # Fill NaT values with the previous valid date
                df['วันที่'] = df['วันที่'].fillna(method='ffill')

                # Create 'month' and 'year' columns
                df['month'] = df['วันที่'].dt.month
                df['year'] = df['วันที่'].dt.year

                # Apply formatting for valid dates to m/d/yyyy format
                def format_date(x):
                    if pd.notna(x):
                        return f"{x.month}/{x.day}/{x.year}"  # Format to m/d/yyyy without leading zeros
                    return ""  # Handle NaT or invalid dates

                df['วันที่'] = df['วันที่'].apply(format_date)

            # Ensure 'Selected' and 'Category' are in the correct positions
            df.insert(3, 'Selected', df.pop('Selected'))  # Move 'Selected' to column D (index 3)
            df.insert(4, 'Category', df.pop('Category'))  # Move 'Category' to column E (index 4)

            # Handle NaT, NaN values and format other columns
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].apply(lambda x: "" if pd.isna(x) else x.strftime('%Y-%m-%d'))
                elif pd.api.types.is_float_dtype(df[col]):
                    df[col] = df[col].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "")
                else:
                    df[col] = df[col].fillna("")

            checkbox_states = {i: False for i in range(len(df))}

            # Display the full data in the Treeview
            update_treeview()  # Ensure this is called here

            messagebox.showinfo("ไฟล์นำเข้า", f"ไฟล์นำเข้า: {file_path}")

        except pd.errors.EmptyDataError:
            messagebox.showwarning("ไฟล์ว่าง", "ไฟล์ที่เลือกว่างเปล่าหรือไม่สามารถอ่านได้.")
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถอ่านไฟล์ได้: {str(e)}")

def load_csv():
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if file_path:
        try:
            df = pd.read_csv(file_path)
            update_treeview(df)  # Update the Treeview with the loaded data
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

def fill_na_dates():
    last_valid_date = None
    for i in range(len(df)):
        if pd.isna(df['วันที่'].iloc[i]):
            df['วันที่'].iloc[i] = last_valid_date
        else:
            last_valid_date = df['วันที่'].iloc[i]
    # Update 'month' and 'year' after filling NaT
    df['month'] = df['วันที่'].dt.month
    df['year'] = df['วันที่'].dt.year


def update_treeview(filtered_df=None):
    # Clear the Treeview before displaying new data
    tree.delete(*tree.get_children())

    # Determine DataFrame to use
    data_to_display = filtered_df if filtered_df is not None else df
    if data_to_display is None:
        return
    # Remove 'month' and 'year' columns from the displayed data
    data_to_display = data_to_display.drop(columns=['month', 'year'], errors='ignore')
    # Set columns in the Treeview
    columns = list(data_to_display.columns)
    tree["columns"] = columns
    tree["show"] = "headings"

    # Configure the column headers
    for col in columns:
        if col == 'Selected':
            tree.heading(col, text="Selected", anchor="w")
            tree.column(col, width=70, anchor="center", stretch=tk.NO)
        elif col == 'Category':
            tree.heading(col, text="Category", anchor="w")
            tree.column(col, width=150, anchor="center", stretch=tk.NO)
        else:
            tree.heading(col, text=col, anchor="w")
            tree.column(col, width=150, anchor="center", stretch=tk.NO)

    # Insert rows into the Treeview
    for i, row in data_to_display.iterrows():
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        tree.insert("", "end", values=list(row), tags=(tag,))

def filter_by_month_year(month, year):
    filtered_df = df[(df['month'] == month) & (df['year'] == year)]
    
    print(f"Number of rows after filtering for month {month} and year {year}: {len(filtered_df)}")
    if filtered_df.empty:
        print("No data found for the specified month and year.")
    else:
        print("Filtered DataFrame:\n", filtered_df)
        update_treeview(filtered_df)  # แสดงข้อมูลที่กรองใน Treeview

def show_selected_month_year():
    global selected_month, selected_year
    month = int(month_combobox.get())
    year = int(year_combobox.get())
    selected_month = month
    selected_year = year

    # Now call your filter function
    filter_by_month_year(month, year)

def toggle_checkbox(event):
    item = tree.identify_row(event.y)
    if item:
        row_id = tree.index(item)
        checkbox_states[row_id] = not checkbox_states[row_id]
        new_value = "✅" if checkbox_states[row_id] else "❌"
        tree.set(item, 'Selected', new_value)
        
        # Update the DataFrame directly
        df.at[row_id, 'Selected'] = new_value

        # Assign category from the combobox if selected, else assign default
        if new_value == "✅":
            selected_category = category_combobox.get() or "ไฟฟ้า"  # Default to "ไฟฟ้า"
            tree.set(item, 'Category', selected_category)
            df.at[row_id, 'Category'] = selected_category
        else:
            tree.set(item, 'Category', "")
            df.at[row_id, 'Category'] = ""

def export_selected_rows():
    global df
    
    if df is None:
        messagebox.showwarning("ไม่มีข้อมูล", "ไม่มีข้อมูลสำหรับการส่งออก กรุณานำเข้าไฟล์ก่อน.")
        return
    
    # Filter rows in the DataFrame where 'Selected' column is '✅'
    selected_rows = df[df['Selected'] == '✅']
    
    if selected_rows.empty:
        messagebox.showwarning("ไม่มีการเลือก", "ไม่มีแถวที่เลือกสำหรับการส่งออก.")
        return
    
    selected_df = df.loc[selected_rows.index]

    # Prepare data for export
    selected_export_data = selected_df.drop(columns=['Selected'])

    # Define all categories
    all_categories = [
    "ปรับปรุง","ไฟฟ้า", "ประปา", "Air", "ปลวก", "สถาบัน", "Fur", "อุปกรณ์",
    "รายรับอื่นๆ", "ขาย", "วัตถุดิบ", "หมู", "ไก่ไข่", "เนื้อ",
    "วัชระ+สุรพล", "ลูกชิ้น", "กุ้ง", "ผัก", "เครื่องดื่ม", "ไอติม",
    "ขนมจีบ", "ของหวาน", "ผลไม้", "อุปกรณ์ใช้สอย", "วัสดุสิ้นเปลือง",
    "ค่าใช้สอย", "PR", "ค่าไฟ", "ค่าแรง", "Ptime", "สวัสดิการ", "Manage"
]
    sums = [0] * len(all_categories)  # Initialize sums with zeros

    # Calculate sums for each category
    for i, category in enumerate(all_categories):
        if category in selected_df['Category'].values:  # Check if category exists
            if category in ["ขาย", "รายรับอื่นๆ"]:
                sums[i] = pd.to_numeric(selected_df[selected_df['Category'] == category].iloc[:, 5], errors='coerce').sum()
            else:
                sums[i] = pd.to_numeric(selected_df[selected_df['Category'] == category].iloc[:, 6], errors='coerce').sum()

    # Create DataFrame for sum by category
    sum_by_category_df = pd.DataFrame([sums], columns=all_categories)

    # Save to Excel
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="บันทึกเป็น"
    )

    if file_path:
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                selected_export_data.to_excel(writer, sheet_name='Merged Data', index=False)
                
                # Write the sum by category DataFrame
                sum_by_category_df.to_excel(
                    writer,
                    sheet_name='Merged Data',
                    index=False,
                    startcol=selected_export_data.shape[1] + 2  # Leave one column space
                )

                # Access workbook and worksheet
                workbook  = writer.book
                worksheet = writer.sheets['Merged Data']

                # Write headers for sum columns
                for col_num, value in enumerate(sum_by_category_df.columns):
                    worksheet.write(0, selected_export_data.shape[1] + 2 + col_num, value)

                # Adjust column widths for better readability
                for i, col in enumerate(sum_by_category_df.columns):
                    worksheet.set_column(selected_export_data.shape[1] + 2 + i, selected_export_data.shape[1] + 2 + i, 20)

            messagebox.showinfo("การส่งออกสำเร็จ", f"ข้อมูลที่เลือกและยอดรวมตามหมวดหมู่ได้ถูกบันทึกไว้ที่: {file_path}")
        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถส่งออกไฟล์ได้: {str(e)}")

def submit_selection():
    # Filter the DataFrame based on selected rows
    filtered_df = df[df['Selected'] == '✅']  # Select only rows that have been marked as selected

    if filtered_df.empty:
        messagebox.showinfo("ไม่มีข้อมูล", "ไม่มีข้อมูลที่ถูกเลือก.")
        return

    # Display the filtered data in the Treeview on the "Filtered Data" page
    update_filtered_treeview(filtered_df)

    # Automatically calculate and update the sum by category
    sum_by_category()

    # Switch to the "Filtered Data" tab
    notebook.select(filtered_data_frame)

def update_filtered_treeview(filtered_df):
    # Clear the Treeview before displaying new data
    filtered_tree.delete(*filtered_tree.get_children())

    # Set columns with correct headings and widths
    columns = list(filtered_df.columns)
    filtered_tree["columns"] = columns
    filtered_tree["show"] = "headings"

    for col in columns:
        filtered_tree.heading(col, text=col)
        filtered_tree.column(col, width=150, anchor="center", stretch=tk.NO)

    # Insert rows into the filtered Treeview
    for i, row in filtered_df.iterrows():
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        filtered_tree.insert("", "end", values=list(row), tags=(tag,))

    # Switch to the "Filtered Data" tab
    notebook.select(filtered_data_frame)

def calculate_sales_sum():
    global df
    if df is None:
        return

    # Filter for rows where Category is 'ขาย'
    sales_df = df[df['Category'] == 'ขาย']

    # Ensure column index 5 contains numeric values (coerce non-numeric to NaN)
    sales_df.iloc[:, 5] = pd.to_numeric(sales_df.iloc[:, 5], errors='coerce')

    # Sum the values in column 5 (Amount column)
    total_sales = sales_df.iloc[:, 5].sum()

    # Update the sum result in the Filtered Data Treeview
    update_sum_treeview(total_sales)

def sum_by_category():
    global df
    if df is None:
        return  # No data available to sum

    # Filter selected rows (where 'Selected' is "✅")
    selected_df = df[df['Selected'] == '✅']

    if selected_df.empty:
        return  # No selected rows, so nothing to sum

    try:
        # Convert relevant columns to numeric, handling errors if necessary
        selected_df['Sum Column 5'] = pd.to_numeric(selected_df.iloc[:, 5], errors='coerce')
        selected_df['Sum Column 6'] = pd.to_numeric(selected_df.iloc[:, 6], errors='coerce')

        # Define custom summing logic for "ขาย" and "รายรับอื่นๆ"
        def custom_sum(group):
            if group.name in ["ขาย", "รายรับอื่นๆ"]:
                return group['Sum Column 5'].sum()
            else:
                return group['Sum Column 6'].sum()

        # Group by category and apply the custom summing logic
        result = selected_df.groupby('Category').apply(custom_sum)

        # Update the Sum Treeview on the Filtered Data page
        update_sum_treeview(result)
    
    except Exception as e:
        messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาดในการคำนวณ: {str(e)}")

def update_sum_treeview(result):
    # Clear the Sum Treeview before displaying new data
    sum_tree.delete(*sum_tree.get_children())

    # Insert the summed values into the Sum Treeview
    for category, total in result.items():
        sum_tree.insert("", "end", values=(category, f"{total:.2f}"))


def display_sum_by_category(result):
    # Create a new window to display the sum results
    result_window = tk.Toplevel()
    result_window.title("Sum by Category")

    frame = tk.Frame(result_window)
    frame.pack(fill=tk.BOTH, expand=True)

    # Use a Treeview to display the results
    tree = ttk.Treeview(frame, show="headings")
    tree["columns"] = ("Category", "Sum")
    tree.heading("Category", text="หมวดหมู่")
    tree.heading("Sum", text="ยอดรวม")

    tree.column("Category", width=150, anchor="center")
    tree.column("Sum", width=100, anchor="center")

    # Insert the summed values into the Treeview
    for category, total in result.items():
        tree.insert("", "end", values=(category, f"{total:.2f}"))

    # Add vertical and horizontal scrollbars to the Treeview
    scrollbar_y = tk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scrollbar_x = tk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x.pack(side="bottom", fill="x")
    tree.pack(fill=tk.BOTH, expand=True)

    # Adjust the size of the window based on the content
    result_window.geometry(f"{result_window.winfo_reqwidth()}x{result_window.winfo_reqheight()}")

def predict_category(text):
    global vectorizer
    text_vec = vectorizer.transform([text])
    return model.predict(text_vec)[0]

def apply_predictions():
    global df
    # Iterate through each row in the Treeview
    for row_id in tree.get_children():
        values = tree.item(row_id, 'values')
        row_index = tree.index(row_id)  # Get the index of the row in the Treeview
        col2_value = values[2]  # Value from index 2 (text to predict)
        col5_value = values[5]  # Value from index 5 (amount)
        col6_value = values[6]
        category = None  # Default empty category

        # Check if there's a value in column 5
        if col5_value:
            # Set 'Category ขาย' unless specific terms from column 2 are found
            if col2_value in ['ขายสด', 'ขายโอน']:
                category = 'ขาย'
            else:
                category = 'รายรับอื่นๆ'

        else:
            # If column 5 is empty, use the AI model for prediction
            predicted_category = predict_category(col2_value)

            # Apply the mapping to change the predicted category accordingly
            if predicted_category == 'เนื้อวัว':
                category = 'เนื้อ'
            elif predicted_category == 'เนื้อหมู':
                category = 'หมู'
            elif predicted_category == ['เนื้อไก่','ไข่']:
                category = 'ไก่ไข่'
            elif predicted_category == 'ของทานเล่น':
                category = ''  # No category, leave it empty
            elif predicted_category == 'ลูกชิ้น':
                category = 'ลูกชิ้น'
            elif predicted_category in ['อื่นๆ', 'อาหารทะเล']:
                category = 'วัตถุดิบ'
            elif predicted_category == 'ผัก':
                category = ''
            elif predicted_category == 'เครื่องดื่ม':
                category = 'เครื่องดื่ม'
            elif predicted_category == 'ของหวาน':
                category = 'ของหวาน'

        if col6_value:  # Check if column index 6 has a value
            if col2_value == 'กุ้ง':
                category = 'กุ้ง'
            elif col2_value == 'ไข่':
                category = 'ไก่ไข่'
            elif col2_value == 'แมคโคร':
                category = 'วัตถุดิบ'
            elif col2_value == 'ผัก':
                category = 'ผัก'
        # Only update the Treeview and DataFrame if the category is not empty
        if category:
            tree.set(row_id, 'Category', category)
            df.at[row_index, 'Category'] = category
            # Mark the row as selected (✅) since a category was assigned
            tree.set(row_id, 'Selected', '✅')
            df.at[row_index, 'Selected'] = '✅'
        else:
            # If no category, keep it unselected (❌)
            tree.set(row_id, 'Selected', '❌')
            df.at[row_index, 'Selected'] = '❌'
            
# Create main window
root = tk.Tk()
root.title("Laohu")

# Create a Notebook for tabbed interface
notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True)

# Create frames for each tab
full_data_frame = tk.Frame(notebook)
filtered_data_frame = tk.Frame(notebook)

# Add frames to notebook
notebook.add(full_data_frame, text="Full Data")  #1
notebook.add(filtered_data_frame, text="Filtered Data") #2

# Create combobox for selecting month (1-12)
month_label = tk.Label(root, text="เดือน (1-12):")
month_label.pack()
month_combobox = ttk.Combobox(root, values=[str(i) for i in range(1, 13)], state="readonly")  # 1-12
month_combobox.pack()

# Create combobox for selecting year (e.g., 2022-2025)
year_label = tk.Label(root, text="ปี (YYYY):")
year_label.pack()
year_combobox = ttk.Combobox(root, values=[str(i) for i in range(2020, 2026)], state="readonly")  # 2020-2025
year_combobox.pack()

filter_button = tk.Button(root, text="แสดงข้อมูลตามเดือน/ปี", command=show_selected_month_year)
filter_button.pack()

# Create "Import File" button
import_button = tk.Button(full_data_frame, text="นำเข้าไฟล์ Excel", command=import_file)
# import_button.pack(side="top", padx=10, pady=5)
import_button.pack(anchor="nw", pady=10, padx=10)

# Create "Export Selected" button
export_button = tk.Button(filtered_data_frame, text="บันทึกไฟล์", command=export_selected_rows)
export_button.pack()

# Create combobox for category selection
category_combobox = ttk.Combobox(full_data_frame, values=categories, state="readonly")
category_combobox.set("เลือกหมวดหมู่")
category_combobox.pack(anchor="ne", pady=10, padx=10)

# Create "Submit" button
submit_button = tk.Button(full_data_frame, text="ยืนยัน", command=submit_selection)
submit_button.pack(side="top", padx=10, pady=5)

# Create the Predict Categories button
predict_button = tk.Button(full_data_frame, text="Predict Categories", command=apply_predictions)
predict_button.pack()

# Create Treeview widget for Full Data
tree = ttk.Treeview(full_data_frame, show="headings", selectmode="browse")

# Add vertical and horizontal scrollbars to the Full Data Treeview
scrollbar_y = tk.Scrollbar(full_data_frame, orient="vertical", command=tree.yview)
scrollbar_x = tk.Scrollbar(full_data_frame, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

# Pack scrollbars and Treeview into the full_data_frame
scrollbar_y.pack(side="right", fill="y")
scrollbar_x.pack(side="bottom", fill="x")
tree.pack(fill=tk.BOTH, expand=True)

# Configure alternating row colors for the Full Data Treeview
tree.tag_configure("evenrow", background="#F2F2F2")  # Light gray background
tree.tag_configure("oddrow", background="#FFFFFF")   # White background

# Ensure the Treeview expands with the window
full_data_frame.grid_rowconfigure(0, weight=1)
full_data_frame.grid_columnconfigure(0, weight=1)


# Create a Treeview widget for displaying sums in the Filtered Data page
sum_tree = ttk.Treeview(filtered_data_frame, show="headings", columns=("Category", "Total Amount"))
sum_tree.heading("Category", text="Category")
sum_tree.heading("Total Amount", text="Total Amount")
sum_tree.column("Category", width=150, anchor="center")
sum_tree.column("Total Amount", width=150, anchor="center")

# Pack the sum Treeview into the filtered data frame
sum_tree.pack(fill=tk.BOTH, expand=True)


# Add vertical and horizontal scrollbars to the Sum Treeview
sum_scrollbar_y = tk.Scrollbar(filtered_data_frame, orient="vertical", command=sum_tree.yview)
sum_scrollbar_x = tk.Scrollbar(filtered_data_frame, orient="horizontal", command=sum_tree.xview)
sum_tree.configure(yscrollcommand=sum_scrollbar_y.set, xscrollcommand=sum_scrollbar_x.set)

# Pack Sum Treeview and scrollbars
sum_scrollbar_y.pack(side="right", fill="y")
sum_scrollbar_x.pack(side="bottom", fill="x")
sum_tree.pack(fill=tk.BOTH, expand=True)

# Configure alternating row colors (gray and white)
tree.tag_configure("evenrow", background="#F2F2F2")  # Light gray background
tree.tag_configure("oddrow", background="#FFFFFF")  # White background

# Configure column headers to have bold font and darker background
style = ttk.Style()
style.configure("Treeview.Heading", font=('Calibri', 12, 'bold'), background="#D3D3D3", foreground="black")

# Add grid lines effect
style.configure("Treeview", highlightthickness=0, borderwidth=0)
style.map("Treeview", background=[('selected', '#C4E1FF')])

# Adjust grid weights to make the Treeview expand with the frame
full_data_frame.grid_rowconfigure(0, weight=1)
full_data_frame.grid_columnconfigure(0, weight=1)

# Create Treeview widget for Filtered Data
filtered_tree = ttk.Treeview(filtered_data_frame, show="headings", selectmode="browse")

# Add vertical and horizontal scrollbars to the Filtered Treeview
filtered_scrollbar_y = tk.Scrollbar(filtered_data_frame, orient="vertical", command=filtered_tree.yview)
filtered_scrollbar_x = tk.Scrollbar(filtered_data_frame, orient="horizontal", command=filtered_tree.xview)
filtered_tree.configure(yscrollcommand=filtered_scrollbar_y.set, xscrollcommand=filtered_scrollbar_x.set)

# Use pack to place Treeview and scrollbars
filtered_scrollbar_y.pack(side="right", fill="y")
filtered_scrollbar_x.pack(side="bottom", fill="x")
filtered_tree.pack(fill=tk.BOTH, expand=True)


# Configure alternating row colors for Filtered Data Treeview
filtered_tree.tag_configure("evenrow", background="#F2F2F2")  # Light gray background
filtered_tree.tag_configure("oddrow", background="#FFFFFF")  # White background

# Bind the toggle_checkbox function to mouse click events
tree.bind("<Button-1>", toggle_checkbox)


# Set the display size based on screen resolution
def set_display_size():
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    full_data_frame.config(width=int(screen_width * 0.8), height=int(screen_height * 0.6))

set_display_size()

print(df)

# Run the main loop
root.mainloop()