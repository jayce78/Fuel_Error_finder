import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime

def process_file(file_path, bunker_file):
    try:
        #   READ DATA  
        # Detect file type for the main data
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)

        # Read bunker data
        df_bunker = pd.read_csv(bunker_file)

        # Create a proper datetime column from Bunker_Delivery_Date + Bunker_Delivery_Time
        df_bunker['Bunker_Delivery_Timestamp'] = pd.to_datetime(
            df_bunker['Bunker_Delivery_Date'] + ' ' + df_bunker['Bunker_Delivery_Time'],
            errors='coerce',
            dayfirst=True
        )

        # Identify date/time columns explicitly
        date_col = "Date_UTC"
        time_col = "Time_UTC"

        if date_col not in df.columns or time_col not in df.columns:
            messagebox.showerror("Error", "Could not find Date_UTC or Time_UTC column.")
            return

        # Ensure date/time format is correct
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce', dayfirst=True)

        # Convert Time_UTC column properly (handling HH:MM format)
        df[time_col] = pd.to_datetime(df[time_col], errors='coerce', format="%H:%M").dt.time

        # Merge Date + Time into Timestamp
        df['Timestamp'] = df.apply(
            lambda row: datetime.combine(row[date_col].date(), row[time_col])
                        if pd.notnull(row[date_col]) and pd.notnull(row[time_col])
                        else pd.NaT,
            axis=1
        )

        # DETECT ALL FUEL TYPES  
        possible_fuels = ['HFO','LFO','MGO','MDO','LNG','LPGP','LPGB','M','E']
        fuel_types = set()
        for col in df.columns:
            parts = col.split('_')
            if len(parts) > 2 and parts[-1] in possible_fuels:
                fuel_types.add(parts[-1])

        print("Detected Fuel Types:", fuel_types)  # Debugging output

        # PREPARE WRITER  
        output_file = os.path.splitext(file_path)[0] + "_FUEL.xlsx"
        with pd.ExcelWriter(output_file, engine='xlsxwriter', mode='w') as writer:
            workbook = writer.book

            # LOOP OVER FUEL TYPES  
            for fuel in fuel_types:
                rob_column = f"{fuel}_ROB"

                # Skip if no ROB column for this fuel
                if rob_column not in df.columns:
                    continue

                # Gather columns for this fuel
                fuel_columns = [col for col in df.columns if col.endswith(fuel)]
                selected_cols = [date_col, time_col, 'Timestamp'] + fuel_columns

                # Create a working DataFrame for the current fuel
                fuel_df = df.loc[:, selected_cols].copy()
                fuel_df[rob_column] = df.get(rob_column, 0)

                # Create a Bunkers column
                fuel_df["Bunkers"] = 0

#-----------------------------------------------------------------#
# MATCH BUNKERS: "Closest" logic
#----------------------------------------------------------------#

                for i, row in fuel_df.iterrows():
                    if pd.isnull(row['Timestamp']):
                        # If no valid Timestamp, skip matching
                        continue

                    fuel_bunkers = df_bunker[df_bunker['Fuel_Type'] == fuel].copy()
                    if fuel_bunkers.empty:
                        continue  # No bunkers of this fuel left

                    fuel_bunkers['TimeDiff'] = (fuel_bunkers['Bunker_Delivery_Timestamp'] - row['Timestamp']).abs()

                    # 3) Find the single bunker row with the minimal time difference
                    closest_index = fuel_bunkers['TimeDiff'].idxmin()
                    closest_diff = fuel_bunkers.at[closest_index, 'TimeDiff']

                    if closest_diff > pd.Timedelta(hours=12):
                        continue

                    fuel_df.at[i, 'Bunkers'] = df_bunker.at[closest_index, 'Mass']
                    df_bunker.drop(index=closest_index, inplace=True)

#--------------------------------------------------------------#
# END BUNKER MATCHING
#--------------------------------------------------------------#

                # Format Date_UTC as dd.mm.yyyy for output
                fuel_df[date_col] = fuel_df[date_col].dt.strftime('%d.%m.%Y')

                # Drop Timestamp column before saving
                if 'Timestamp' in fuel_df.columns:
                    fuel_df.drop(columns=['Timestamp'], inplace=True)

                # Write to Excel (new sheet for each fuel)
                fuel_df.to_excel(writer, sheet_name=fuel, index=False)
                worksheet = writer.sheets[fuel]

                # Auto-adjust column widths
                for col_idx, col_name in enumerate(fuel_df.columns):
                    max_len = max(fuel_df[col_name].astype(str).map(len).max(), len(col_name)) + 2
                    worksheet.set_column(col_idx, col_idx, max_len)

                # BOLD HEADERS
                bold_format = workbook.add_format({'bold': True})
                for col_idx, col_name in enumerate(fuel_df.columns):
                    worksheet.write(0, col_idx, col_name, bold_format)

                # Prepare TOTALS, FORMULAS, ETC.
                total_rows = len(fuel_df) + 1  # 1-based for Excel, includes header
                col_letters = {col_name: chr(65 + idx) for idx, col_name in enumerate(fuel_df.columns)}

                # The index for the ROB_Difference column:
                rob_diff_col_idx = len(fuel_df.columns)
                rob_diff_col_letter = chr(65 + rob_diff_col_idx)
                worksheet.write(0, rob_diff_col_idx, "ROB_Difference", bold_format)

                # Move TWO columns over to place "TOTALS" (one blank column in between).
                totals_col_idx = rob_diff_col_idx + 2
                totals_col_letter = chr(65 + totals_col_idx)
                worksheet.write(0, totals_col_idx, "TOTALS", bold_format)

                # Insert the formula for each row (row 2 onwards, since row 1 is header)
                rob_col_letter = col_letters[rob_column]
                bunker_col_letter = col_letters["Bunkers"]

                for row_num in range(2, total_rows + 1):
                    # e.g.: =IFERROR(C2 - C3 + D3, 0)
                    worksheet.write_formula(
                        row_num,
                        rob_diff_col_idx,
                        f"=IFERROR({rob_col_letter}{row_num}-{rob_col_letter}{row_num+1}+{bunker_col_letter}{row_num+1}, 0)"
                    )

                # Highlight negative ROB_Difference in yellow
                yellow_format = workbook.add_format({'bg_color': 'yellow'})
                worksheet.conditional_format(
                    f"{rob_diff_col_letter}2:{rob_diff_col_letter}{total_rows}",
                    {
                        'type': 'cell',
                        'criteria': '<',
                        'value': 0,
                        'format': yellow_format,
                    }
                )

                # Collect consumption columns (if any)
                consumption_columns = [c for c in fuel_df.columns if "_Consumption_" in c]

                # TOTALS column (reuse totals_col_letter)
                worksheet.write(0, totals_col_idx, "TOTALS", bold_format)

                # Write total consumption formulas
                row_index = 2
                for c in consumption_columns:
                    c_letter = col_letters[c]
                    # Label
                    worksheet.write(f"{totals_col_letter}{row_index}", c)
                    # SUM formula
                    worksheet.write(
                        f"{chr(65 + totals_col_idx + 1)}{row_index}",
                        f"=SUM({c_letter}2:{c_letter}{total_rows})"
                    )
                    row_index += 1

                # TOTAL CONSUMPTION (sum of the above sums)
                worksheet.write(f"{totals_col_letter}{row_index}", "TOTAL CONSUMPTION")
                worksheet.write(
                    f"{chr(65 + totals_col_idx + 1)}{row_index}",
                    f"=SUM({chr(65 + totals_col_idx + 1)}2:{chr(65 + totals_col_idx + 1)}{row_index-1})"
                )

                # Bunker totals
                bunker_total_row = row_index + 2
                worksheet.write(f"{totals_col_letter}{bunker_total_row}", "TOTAL BUNKERED")
                worksheet.write(
                    f"{chr(65 + totals_col_idx + 1)}{bunker_total_row}",
                    f"=SUM({bunker_col_letter}2:{bunker_col_letter}{total_rows})"
                )

                # ROB totals (sum from the ROB_Difference column)
                rob_total_row = bunker_total_row + 2
                worksheet.write(f"{totals_col_letter}{rob_total_row}", "TOTAL CONSUMED (ROB)")
                worksheet.write(
                    f"{chr(65 + totals_col_idx + 1)}{rob_total_row}",
                    f"=SUM({rob_diff_col_letter}2:{rob_diff_col_letter}{total_rows})"
                )

                # MISSING section
                missing_row = rob_total_row + 2
                red_bold = workbook.add_format({'bold': True, 'font_color': 'red'})
                worksheet.write(f"{totals_col_letter}{missing_row}", "MISSING", red_bold)

                # MISSING = TOTAL CONSUMPTION - TOTAL CONSUMED (ROB)
                worksheet.write(
                    f"{chr(65 + totals_col_idx + 1)}{missing_row}",
                    f"=({chr(65 + totals_col_idx + 1)}{row_index}"
                    f"-{chr(65 + totals_col_idx + 1)}{rob_total_row})",
                    red_bold)

#--------------------------------------------------------------------------#
# CREATE BDN SHEET (once, outside the fuel loop)  
#--------------------------------------------------------------------------#

            # Required columns for BDN. If not present, set them to zero.
            bdn_columns = [
                "ME_Consumption", "AE_Consumption", "Boiler_Consumption",
                "IGG_Consumption", "DPP_Consumption", "Incinerator_Consumption",
                "BDN_ROB",  # we will insert "Bunkers" right after this column
                "ROB_Fuel_BDN"
            ]

#------------------------------------------------------------------------#
# ADD HVO BUNKERS TO df FOR THE BDN TAB  
#------------------------------------------------------------------------#

            # Ensure the Bunkers column exists
            if "Bunkers" not in df.columns:
                df["Bunkers"] = 0

            allowed_fuel_types = {"HVO", "FAME", "Bio"}  # Add more as needed

            # Loop over each row in the main df
            for i, row in df.iterrows():
                if pd.isnull(row["Timestamp"]):
                    continue  # skip rows without a valid Timestamp

                # Find all relevant bunkers near this row’s timestamp
                valid_bunkers = df_bunker[
                    (df_bunker["Fuel_Type"].isin(allowed_fuel_types))
                    & (abs(df_bunker["Bunker_Delivery_Timestamp"] - row["Timestamp"]) <= pd.Timedelta(hours=12))
                ]
                
                if not valid_bunkers.empty:
                    # Compute which bunker timestamp is closest
                    valid_bunkers["TimeDiff"] = (valid_bunkers["Bunker_Delivery_Timestamp"] - row["Timestamp"]).abs()
                    closest_index = valid_bunkers["TimeDiff"].idxmin()

                    # Add that bunker's mass to df["Bunkers"]
                    df.at[i, "Bunkers"] += df_bunker.at[closest_index, "Mass"]

                    # Drop that used bunker row so it's not reused
                    df_bunker.drop(index=closest_index, inplace=True)

            # Extend bdn_columns to include "Bunkers" in the correct position
            bdn_columns = [
                "ME_Consumption", "AE_Consumption", "Boiler_Consumption",
                "IGG_Consumption", "DPP_Consumption", "Incinerator_Consumption",
                "BDN_ROB", "Bunkers", "ROB_Fuel_BDN"
            ]

            for col in bdn_columns:
                if col not in df.columns:
                    df[col] = 0

            # Build a new DataFrame for BDN
            bdn_df = df.loc[:, [date_col, time_col, 'Timestamp'] + bdn_columns].copy()

            # Format date as dd.mm.yyyy
            bdn_df[date_col] = bdn_df[date_col].dt.strftime('%d.%m.%Y')

            # Drop Timestamp if present
            if 'Timestamp' in bdn_df.columns:
                bdn_df.drop(columns=['Timestamp'], inplace=True)

            # Write the BDN DataFrame to Excel
            bdn_df.to_excel(writer, sheet_name="BDN", index=False)
            worksheet_bdn = writer.sheets["BDN"]

            # Auto-size columns
            for i, col_name in enumerate(bdn_df.columns):
                max_len = max(bdn_df[col_name].astype(str).map(len).max(), len(col_name)) + 2
                worksheet_bdn.set_column(i, i, max_len)

            # Bold headers
            bold_format = workbook.add_format({'bold': True})
            for col_idx, col_name in enumerate(bdn_df.columns):
                worksheet_bdn.write(0, col_idx, col_name, bold_format)

            total_rows = len(bdn_df) + 1
            col_letters = {col_name: chr(65 + idx) for idx, col_name in enumerate(bdn_df.columns)}

            # Index for the ROB_Difference column
            rob_diff_col_idx = len(bdn_df.columns)
            rob_diff_col_letter = chr(65 + rob_diff_col_idx)
            worksheet_bdn.write(0, rob_diff_col_idx, "ROB_Difference", bold_format)

            # Skip one column, place TOTALS two columns after ROB_Difference
            totals_col_idx = rob_diff_col_idx + 2
            totals_col_letter = chr(65 + totals_col_idx)
            worksheet_bdn.write(0, totals_col_idx, "TOTALS", bold_format)

            # Insert the formula for each row (using BDN_ROB)
            bdn_rob_col_letter = col_letters["BDN_ROB"]
            bunker_col_letter = col_letters["Bunkers"]
            for row_num in range(2, total_rows + 1):
                worksheet_bdn.write_formula(
                    row_num,
                    rob_diff_col_idx,
                    f"=IFERROR({bdn_rob_col_letter}{row_num}-{bdn_rob_col_letter}{row_num+1}+{bunker_col_letter}{row_num+1}, 0)"
                    )

            # Highlight negative ROB_Difference
            yellow_format = workbook.add_format({'bg_color': 'yellow'})
            worksheet_bdn.conditional_format(
                f"{rob_diff_col_letter}2:{rob_diff_col_letter}{total_rows}",
                {
                    'type': 'cell',
                    'criteria': '<',
                    'value': 0,
                    'format': yellow_format,
                }
            )

            # Consumption columns for totals
            consumption_columns = [c for c in bdn_columns if "Consumption" in c]

            worksheet_bdn.write(0, totals_col_idx, "TOTALS", bold_format)
            row_index = 2
            for c in consumption_columns:
                c_letter = col_letters[c]
                # Label in the totals column
                worksheet_bdn.write(f"{totals_col_letter}{row_index}", c)
                # Sum in the next column
                worksheet_bdn.write(
                    f"{chr(65 + totals_col_idx + 1)}{row_index}",
                    f"=SUM({c_letter}2:{c_letter}{total_rows})"
                )
                row_index += 1

            # TOTAL CONSUMPTION
            worksheet_bdn.write(f"{totals_col_letter}{row_index}", "TOTAL CONSUMPTION")
            worksheet_bdn.write(
                f"{chr(65 + totals_col_idx + 1)}{row_index}",
                f"=SUM({chr(65 + totals_col_idx + 1)}2:{chr(65 + totals_col_idx + 1)}{row_index-1})"
            )

            # TOTAL CONSUMED (ROB)
            rob_total_row = row_index + 2
            worksheet_bdn.write(f"{totals_col_letter}{rob_total_row}", "TOTAL CONSUMED (ROB)")
            worksheet_bdn.write(
                f"{chr(65 + totals_col_idx + 1)}{rob_total_row}",
                f"=SUM({rob_diff_col_letter}2:{rob_diff_col_letter}{total_rows})"
            )

            # MISSING
            missing_row = rob_total_row + 2
            red_bold = workbook.add_format({'bold': True, 'font_color': 'red'})
            worksheet_bdn.write(f"{totals_col_letter}{missing_row}", "MISSING", red_bold)

            # MISSING = TOTAL CONSUMPTION - TOTAL CONSUMED (ROB)
            worksheet_bdn.write(
                f"{chr(65 + totals_col_idx + 1)}{missing_row}",
                f"=({chr(65 + totals_col_idx + 1)}{row_index}"
                f"-{chr(65 + totals_col_idx + 1)}{rob_total_row})",
                red_bold
            )

            # Bunkers column letter
            bunker_col_letter = col_letters["Bunkers"]

            # Place the 'TOTAL BUNKERED' label and sum formula
            bunker_total_row = row_index + 1
            worksheet_bdn.write(f"{totals_col_letter}{bunker_total_row}", "TOTAL BUNKERED", bold_format)
            worksheet_bdn.write(
                f"{chr(65 + totals_col_idx + 1)}{bunker_total_row}",
                f"=SUM({bunker_col_letter}2:{bunker_col_letter}{total_rows})"
            )

        # Show success message
        messagebox.showinfo("Success", f"Processed successfully! Output saved as {output_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# File selection
def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")]
    )
    if not file_path:
        messagebox.showwarning("Warning", "Please select an OVD Excel file.")
        return

    bunker_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not bunker_file:
        messagebox.showwarning("Warning", "Please select a CSV Bunker report.")
        return

    process_file(file_path, bunker_file)

# Setup and layout
def create_gui():
    root = tk.Tk()
    root.title("Fuel Consumption Finder")
    root.geometry("400x250")

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True)

    title = tk.Label(frame, text="Select an OVD Excel file and a CSV Bunker report", 
                     font=("Arial", 12), justify="center")
    title.pack(pady=5)

    select_btn = tk.Button(frame, text="Select Files", command=select_file,
                           font=("Arial", 10), bg="lightblue")
    select_btn.pack(pady=10)

    exit_btn = tk.Button(frame, text="Exit", command=root.quit,
                         font=("Arial", 10), bg="lightcoral")
    exit_btn.pack()

    name_label = tk.Label(root, text="Devolped by Jason Leeworthy", font=("Arial", 8))
    name_label.pack(side="bottom", pady=5)

    root.mainloop()

# Run the GUI
create_gui()

