#%%
# original

import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
import numpy as np



# %%
# OPEN THE DATABASE AS CSV THEN PYTHON WILL CONVERTED FROM CSV TO XLSX
import pandas as pd

# Define file paths
csv_file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess\FC1DataBase.csv" # HIBLOW PHILIPPINES INC
xlsx_file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess\Converted_FC1DataBase.xlsx" # HIBLOW PHILIPPINES INC
# csv_file_path = r"C:\Users\hp\Desktop\PROG_SERVER_DATABASE\PROGRAMMING LANGUAGE\PYTHON (VScode)\HIBLOW PROJECT\PYTHON FILES\JHUN'S HIBLOW PROJECT\1ST PROJECT CORRELATION\DATABASE CSV\FC1DataBase.csv" # HOME CAVITE
# xlsx_file_path = r"C:\Users\hp\Desktop\PROG_SERVER_DATABASE\PROGRAMMING LANGUAGE\PYTHON (VScode)\HIBLOW PROJECT\PYTHON FILES\JHUN'S HIBLOW PROJECT\1ST PROJECT CORRELATION\DATABASE CSV\Converted_FC1DataBase.xlsx" # HOME CAVITE

# Load CSV file
df = pd.read_csv(csv_file_path)

# Save as an Excel file
df.to_excel(xlsx_file_path, index=False, engine='openpyxl')

print(f"CSV file successfully converted and saved as {xlsx_file_path}")



#%%
# Select specific (10x39) columns needeed for correlation
import pandas as pd
from openpyxl import load_workbook

# Define file path
# file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess\Converted_FC1DataBase.xlsx" # HIBLOW PHILIPPINES INC
file_path = xlsx_file_path

# Step 1: Load the Excel file
df = pd.read_excel(file_path, sheet_name="Sheet1", engine="openpyxl")  # Load first sheet

# Step 2: Select specific columns (10X39)
columns_to_transfer = ['VOLTAGE MAX (V)', 'WATTAGE MAX (W)', "MODEL CODE", 'CLOSED PRESSURE_MAX (kPa)', 'VOLTAGE Middle (V)', 'WATTAGE Middle (W)', 'AMPERAGE Middle (A)', 'CLOSED PRESSURE Middle (kPa)', 'VOLTAGE MIN (V)', 'WATTAGE MIN (W)', 'CLOSED PRESSURE MIN (kPa)','Process 1 Em2p Inspection 5 Average Data', 'Process 1 Em2p Inspection 10 Average Data', 'Process 1 Em2p Inspection 3 Minimum Data', 'Process 1 Em2p Inspection 4 Minimum Data', 'Process 1 Em2p Inspection 5 Minimum Data', 'Process 1 Em2p Inspection 3 Maximum Data', 'Process 1 Em2p Inspection 4 Maximum Data', 'Process 1 Em2p Inspection 5 Maximum Data', 'Process 1 Em3p Inspection 3 Average Data', 'Process 1 Em3p Inspection 4 Average Data', 'Process 1 Em3p Inspection 5 Average Data', 'Process 1 Em3p Inspection 10 Average Data', 'Process 1 Em3p Inspection 3 Minimum Data', 'Process 1 Em3p Inspection 4 Minimum Data', 'Process 1 Em3p Inspection 5 Minimum Data', 'Process 1 Em3p Inspection 3 Maximum Data', 'Process 1 Em3p Inspection 4 Maximum Data', 'Process 1 Em3p Inspection 5 Maximum Data', 'Process 1 Frame Inspection 1 Average Data', 'Process 1 Frame Inspection 2 Average Data', 'Process 1 Frame Inspection 3 Average Data', 'Process 1 Frame Inspection 4 Average Data', 'Process 1 Frame Inspection 5 Average Data', 'Process 1 Frame Inspection 6 Average Data', 'Process 1 Frame Inspection 7 Average Data', 'Process 1 Frame Inspection 1 Minimum Data', 'Process 1 Frame Inspection 2 Minimum Data', 'Process 1 Frame Inspection 3 Minimum Data', 'Process 1 Frame Inspection 4 Minimum Data', 'Process 1 Frame Inspection 5 Minimum Data', 'Process 1 Frame Inspection 6 Minimum Data', 'Process 1 Frame Inspection 7 Minimum Data', 'Process 1 Frame Inspection 1 Maximum Data', 'Process 1 Frame Inspection 2 Maximum Data', 'Process 1 Frame Inspection 3 Maximum Data', 'Process 1 Frame Inspection 4 Maximum Data', 'Process 1 Frame Inspection 5 Maximum Data', 'Process 1 Frame Inspection 6 Maximum Data', 'Process 1 Frame Inspection 7 Maximum Data'
]
df_selected = df[columns_to_transfer]  # Extract the required columns

# Step 3: Load the workbook and write to "Sheet2"
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a") as writer:
    df_selected.to_excel(writer, sheet_name="Sheet2", index=False)

print("Selected columns successfully transferred to 'Sheet2' in the same Excel file.")






import pandas as pd
#CLEANING OF DATA AND CREATE A NEW FILE NAMED CLEANED
# Load Excel File
# file_path = r"\\192.168.2.19\ai_team\AI Program\Outputs\CompiledProcess\Converted_FC1DataBase.xlsx"  # HIBLOW PHILIPPINES INC
file_path = xlsx_file_path
sheet_name = 'Sheet2'  # Specify the sheet name
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')  # Ensure compatibility

# Define values to remove
values_to_remove = ["NG PRESSURE", "NG AT PROCESS4", "MASTER PUMP", "NG PRESSURE AT PROCESS5", "NG AT PROCESS3"]

# Remove rows containing these values in any column
df_filtered = df[~df.isin(values_to_remove).any(axis=1)]

# Save to a new Excel file with an explicit writer engine
new_file_path = 'FC1DataBase_Cleaned.xlsx'  # Change to .xlsx format for better compatibility
with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
    df_filtered.to_excel(writer, index=False, sheet_name='Sheet9')

print(f"Values removed and data saved to {new_file_path}")




#%%
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np  # Added for regression line calculation

# Load Excel File
file_path = "FC1DataBase_Cleaned.xlsx"  # Ensure correct path
sheet_name = "Sheet9"  # Specify sheet name
df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# Predefine dropdown options
model_codes = df.iloc[:, 2].unique().tolist()

y_axis_options = ['VOLTAGE MAX (V)', 'WATTAGE MAX (W)', 'CLOSED PRESSURE_MAX (kPa)', 
                  'VOLTAGE Middle (V)', 'WATTAGE Middle (W)', 'AMPERAGE Middle (A)', 
                  'CLOSED PRESSURE Middle (kPa)', 'VOLTAGE MIN (V)', 'WATTAGE MIN (W)', 
                  'CLOSED PRESSURE MIN (kPa)']

x_axis_options = ['Process 1 Em2p Inspection 5 Average Data', 'Process 1 Em2p Inspection 10 Average Data', 'Process 1 Em2p Inspection 3 Minimum Data', 'Process 1 Em2p Inspection 4 Minimum Data', 'Process 1 Em2p Inspection 5 Minimum Data', 'Process 1 Em2p Inspection 3 Maximum Data', 'Process 1 Em2p Inspection 4 Maximum Data', 'Process 1 Em2p Inspection 5 Maximum Data', 'Process 1 Em3p Inspection 3 Average Data', 'Process 1 Em3p Inspection 4 Average Data', 'Process 1 Em3p Inspection 5 Average Data', 'Process 1 Em3p Inspection 10 Average Data', 'Process 1 Em3p Inspection 3 Minimum Data', 'Process 1 Em3p Inspection 4 Minimum Data', 'Process 1 Em3p Inspection 5 Minimum Data', 'Process 1 Em3p Inspection 3 Maximum Data', 'Process 1 Em3p Inspection 4 Maximum Data', 'Process 1 Em3p Inspection 5 Maximum Data', 'Process 1 Frame Inspection 1 Average Data', 'Process 1 Frame Inspection 2 Average Data', 'Process 1 Frame Inspection 3 Average Data', 'Process 1 Frame Inspection 4 Average Data', 'Process 1 Frame Inspection 5 Average Data', 'Process 1 Frame Inspection 6 Average Data', 'Process 1 Frame Inspection 7 Average Data', 'Process 1 Frame Inspection 1 Minimum Data', 'Process 1 Frame Inspection 2 Minimum Data', 'Process 1 Frame Inspection 3 Minimum Data', 'Process 1 Frame Inspection 4 Minimum Data', 'Process 1 Frame Inspection 5 Minimum Data', 'Process 1 Frame Inspection 6 Minimum Data', 'Process 1 Frame Inspection 7 Minimum Data', 'Process 1 Frame Inspection 1 Maximum Data', 'Process 1 Frame Inspection 2 Maximum Data', 'Process 1 Frame Inspection 3 Maximum Data', 'Process 1 Frame Inspection 4 Maximum Data', 'Process 1 Frame Inspection 5 Maximum Data', 'Process 1 Frame Inspection 6 Maximum Data', 'Process 1 Frame Inspection 7 Maximum Data']

# Function to filter data and plot
def plot_data():
    user_input_model = combobox_model.get()
    y_axis_selection = combobox_y.get()
    x_axis_selection = combobox_x.get()

    # Ensure all inputs are selected
    if not user_input_model or not y_axis_selection or not x_axis_selection:
        messagebox.showerror("Error", "Please select all required inputs!")
        return

    # Filter data based on selected Model Code
    filtered_df = df[df.iloc[:, 2] == user_input_model]

    if filtered_df.empty:
        messagebox.showerror("Error", "No matching data found!")
        return

    try:
        # Group the data by the selected X-axis column and calculate the mean for the selected Y-axis column
        grouped_df = filtered_df.groupby(x_axis_selection)[y_axis_selection].mean().reset_index()

        if grouped_df.empty:
            messagebox.showerror("Error", "No grouped data found for plotting!")
            return

        # Extract X and Y values for plotting
        x = grouped_df.iloc[:, 0]  # X-axis data
        y = grouped_df.iloc[:, 1]  # Y-axis data

        # Plot the average values (scatter plot)
        plt.figure(figsize=(14, 6))
        plt.scatter(x, y, marker='o', color='blue', label='Data Points')

        # Add a regression line
        z = np.polyfit(x, y, 1)  # Fit a linear regression model (degree = 1)
        p = np.poly1d(z)  # Create polynomial equation
        plt.plot(x, p(x), "r--", label='Regression Line')  # Plot regression line

        # Customize the plot
        plt.title(f"Model {user_input_model}: {y_axis_selection} vs {x_axis_selection}", fontsize=16)
        plt.xlabel(x_axis_selection, fontsize=12)  # X-axis label font size
        plt.ylabel(y_axis_selection, fontsize=18)  # Y-axis label font size
        plt.xticks(fontsize=12, rotation=45, ha='right')  # Rotate X-axis tick labels and align them
        plt.yticks(fontsize=12)  # Y-axis tick font size
        plt.grid(True)  # Optional grid lines for better readability
        plt.legend()  # Display the legend

        # Display the plot
        plt.tight_layout()  # Automatically adjusts layout
        plt.show()

    except Exception as e:
        # Handle any unexpected errors
        messagebox.showerror("Error", f"An error occurred while plotting: {e}")

# Create Tkinter Window
root = tk.Tk()
root.title("CORRELATION DATA ANALYSIS")
root.geometry("550x350")
root.resizable(False, False)

# Function to calculate the width for the dropdowns based on the longest item
def calculate_width(options):
    return max(len(option) for option in options) + 5

# Calculate the width for each dropdown
model_width = calculate_width(model_codes)
y_axis_width = calculate_width(y_axis_options)
x_axis_width = calculate_width(x_axis_options)

# Labels and Dropdowns
tk.Label(root, text="Model Code:", font=("Arial", 20, "bold")).pack(pady=5)
combobox_model = ttk.Combobox(root, values=model_codes, font=("Arial", 15), width=model_width)
combobox_model.pack(pady=5)
combobox_model.set("Choose Model Code")

tk.Label(root, text="Y-Axis:", font=("Arial", 20, "bold")).pack(pady=5)
combobox_y = ttk.Combobox(root, values=y_axis_options, font=("Arial", 15), width=y_axis_width)
combobox_y.pack(pady=5)
combobox_y.set("Choose Y-Axis")

tk.Label(root, text="X-Axis:", font=("Arial", 20, "bold")).pack(pady=5)
combobox_x = ttk.Combobox(root, values=x_axis_options, font=("Arial", 15), width=x_axis_width)
combobox_x.pack(pady=5)
combobox_x.set("Choose X-Axis")

# Button to trigger plotting
tk.Button(root, text="Plot Data", font=("Arial", 20, "bold"), fg="blue", command=plot_data).pack(pady=25)

# Run Tkinter
root.mainloop()
