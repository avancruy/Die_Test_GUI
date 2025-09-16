import os
import tkinter as tk
import re
import pandas as pd
import math
from datetime import datetime
from openpyxl.styles import Font
import numpy as np
from pulsed_classes import Base

def linear_regression(x, y):
    """
    Custom linear regression function using numpy to replace scipy.stats.linregress
    Returns slope, intercept, r_value, p_value, std_err
    """
    n = len(x)
    if n < 2:
        return np.nan, np.nan, np.nan, np.nan, np.nan

    x = np.array(x)
    y = np.array(y)

    # Calculate means
    x_mean = np.mean(x)
    y_mean = np.mean(y)

    # Calculate slope and intercept
    numerator = np.sum((x - x_mean) * (y - y_mean))
    denominator = np.sum((x - x_mean) ** 2)

    if denominator == 0:
        return np.nan, np.nan, np.nan, np.nan, np.nan

    slope = numerator / denominator
    intercept = y_mean - slope * int(x_mean)

    # Calculate correlation coefficient (r_value)
    y_pred = slope * x + intercept
    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - y_mean) ** 2)

    if ss_tot == 0:
        r_value = np.nan
    else:
        r_squared = 1 - (ss_res / ss_tot)
        r_value = np.sqrt(r_squared) if r_squared >= 0 else -np.sqrt(-r_squared)
        if slope < 0:
            r_value = -r_value

    # Calculate standard error of slope
    if n > 2:
        s_y = np.sqrt(ss_res / (n - 2))
        std_err = s_y / np.sqrt(denominator)
    else:
        std_err = np.nan

    # p_value calculation is complex, so we'll return NaN for simplicity
    p_value = np.nan

    return slope, intercept, r_value, p_value, std_err

def extract_date_from_filename(filename):
    """
    Extract date from filename and format as MM/DD/YYYY
    Expected format: ...20250729T093320.xlsx
    """
    try:
        # Use regex to find the date pattern (8 digits followed by T and 6 digits)
        pattern = r'(\d{8})T\d{6}'
        match = re.search(pattern, filename)

        if match:
            date_str = match.group(1)  # Extract the 8-digit date
            # Parse the date string (YYYYMMDD format)
            year = date_str[:4]
            month = date_str[4:6]
            day = date_str[6:8]

            # Format as MM/DD/YYYY
            formatted_date = f"{month}/{day}/{year}"
            return formatted_date
        else:
            print(f"  Warning: Could not extract date from filename: {filename}")
            return np.nan
    except Exception as e:
        print(f"  Error extracting date from {filename}: {e}")
        return np.nan

class Extraction(Base):
    path = ""
    def setup_tab(self, parent):

        tk.Label(parent, text="Click \"Run\" to run all 4").grid(row=0, column=0, pady=2)

        liv_button = tk.Button(parent, text="LIV Data", command= lambda: self.get_LIV_data(self.path),
                                    bg='#2196F3', fg='white', font=('Arial', 8, 'bold'), relief='raised')
        liv_button.grid(row=1, column=0, pady=2, sticky="W")

        ext_button = tk.Button(parent, text="Extinction Data", command=lambda: self.get_extinction(self.path),
                               bg='#2196F3', fg='white', font=('Arial', 8, 'bold'), relief='raised')
        ext_button.grid(row=2, column=0, pady=2, sticky="W")

        spectrum_button = tk.Button(parent, text="Spectrum Data", command=lambda: self.get_spectrum_data(self.path),
                                    bg='#2196F3', fg='white', font=('Arial', 8, 'bold'), relief='raised')
        spectrum_button.grid(row=3, column=0, pady=2, sticky="W")

        org_data_button = tk.Button(parent, text="Organized Data", command=lambda: self.get_organized_data(self.path),
                                    bg='#2196F3', fg='white', font=('Arial', 8, 'bold'), relief='raised')
        org_data_button.grid(row=4, column=0, pady=2, sticky="W")

    def run_test(self, data_path="", device_id="", temperature="", timestamp=""):
        print(f"Running Data Extraction at {data_path}")
        self.get_LIV_data(str(data_path))
        self.get_extinction(str(data_path))
        self.get_spectrum_data(str(data_path))
        self.get_final_data(str(data_path))

    def get_LIV_data(self, input_dir):
        print("\nGetting LIV Data...")
        if not os.path.isdir(input_dir):
            print(f"Error: Folder {input_dir} not found")
            return
        results = []

        for filename in os.listdir(input_dir):
            if filename.endswith(".xlsx") and "_LIV_" in filename and not filename.startswith('~'):
                # Chip Name
                chip_name = None

                filepath = os.path.join(input_dir, filename)
                try:
                    match = re.search(r'[A-Za-z]{2}\d{4}', filename)
                    if match:
                        chip_name = match.group()
                except Exception as e:
                    print(f"Error extracting chip name from {filename}: {e}")
                    continue

                # Date
                file_date = extract_date_from_filename(filename)

                try:
                    df = pd.read_excel(filepath)
                    pd_current = [None, None, None, None, None]
                    current_intercept = None

                    # PD Current from Laser Current
                    idx = 0
                    for l_current in [0, 80, 100]:
                        row = df[df['SMU1_Ch2_Laser_Current_Set_mA'] == l_current]
                        if not row.empty:
                            pd_current[idx] = row['SMU1_Ch1_PD_Current_Meas_mA'].iloc[0]
                        else:
                            print(f"No Data found for Laser {l_current} mA in {filename}")
                        idx += 1

                    # PD Current from EAM Voltage
                    for eam_v in [0, -3]:
                        row = df[df['SMU2_Ch1_EAM_Voltage_Set_V'] == eam_v]
                        if not row.empty:
                            pd_current[idx] = row['SMU1_Ch1_PD_Current_Meas_mA'].iloc[0]
                        else:
                            if not (eam_v == -3):  # Added to ignore warning for -3 EAM
                                print(f"No Data found for EAM {eam_v} V in {filename}")
                        idx += 1

                    # Intercept
                    filtered_df = df[(df['SMU1_Ch2_Laser_Current_Set_mA'] >= 30) &
                                     (df['SMU1_Ch2_Laser_Current_Set_mA'] <= 50)]

                    if not filtered_df.empty and len(filtered_df) >= 2:  # Need at least 2 points for regression
                        x_data = filtered_df['SMU1_Ch2_Laser_Current_Set_mA']
                        y_data = filtered_df['SMU1_Ch1_PD_Current_Meas_mA']

                        # Perform linear regression using our custom function
                        slope, intercept, r_value, p_value, std_err = linear_regression(x_data, y_data)

                        if slope != 0:  # Avoid division by zero
                            current_intercept = -intercept / slope
                        else:
                            print(
                                f"  Warning: Slope is zero for laser current intercept calculation in {filename}. Intercept cannot be determined.")
                    else:
                        print(
                            f"  Warning: Not enough data points (30-50mA Laser Current) for intercept calculation in {filename}")

                    results.append({
                        "Chip Name": chip_name,
                        "Date": file_date,  # New: Add date column
                        "PD_Current_at_0mA_Laser": pd_current[0],
                        "PD_Current_at_80mA_Laser": pd_current[1],
                        "PD_Current_at_100mA_Laser": pd_current[2],
                        "Laser_Current_Intercept_mA": current_intercept,  # New: Add to results
                        "PD_Current_at_0V_EAM": pd_current[3],
                        "PD_Current_at_-3V_EAM": pd_current[4]
                    })

                except FileNotFoundError:
                    print(f"Error: File not found - {filepath}")
                except KeyError as e:
                    print(f"Error: Missing expected column '{e}' in file {filename}. Please check column names.")
                except Exception as e:
                    print(f"An unexpected error occurred while processing {filename}: {e}")

        results_df = pd.DataFrame(results)

        if not results_df.empty:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            output_path = os.path.join(input_dir, f"threshold_values_{timestamp}.xlsx")
            results_df.to_excel(output_path, index=False)
            print(f"File Saved to {output_path}")

    def get_extinction(self, input_dir):
        print("\nGetting Extinction Data...")
        if not os.path.isdir(input_dir):
            return

        chip_names = []
        ext_vals_80 = []
        ext_vals_100 = []
        bold_flags_80 = []
        bold_flags_100 = []

        for filename in os.listdir(input_dir):
            if filename.endswith(".xlsx") and "_EAM_" in filename and not filename.startswith('~'):
                chip_name = None
                bold = False
                ext_val = None

                # Extract Chip Name
                try:
                    match = re.search(r'[A-Za-z]{2}\d{4}', filename)
                    if match:
                        chip_name = match.group()
                except Exception as e:
                    print(f"Error extracting chip name from {filename}: {e}")
                    continue

                file_path = os.path.join(input_dir, filename)
                df = pd.read_excel(file_path)
                ext_i = df.iat[1, 5]
                ext_f = df.iat[31, 5]

                if pd.notna(ext_i) and pd.notna(ext_f):
                    numer = abs(ext_f)
                    denom = abs(ext_i)

                    # Determine if subtraction should occur
                    if numer > 0.111 and denom > 0.111:
                        numer -= 0.111
                        denom -= 0.111
                    else:
                        bold = True

                    # Attempt to calculate ext value
                    if denom != 0 and numer > 0:
                        ext_val = 10 * math.log10(numer / denom)

                    else:
                        print("Math Error")
                        ext_val = f"{ext_i},{ext_f}"

                else:
                    print("Error: Unsuccessful Extraction of PD Current Values")

                # Add values to output arrays
                chip_names.append(chip_name)

                if "LDBias(80)" in filename or "80mA" in filename:
                    ext_vals_80.append(ext_val)
                    bold_flags_80.append(bold)

                    ext_vals_100.append(None)
                    bold_flags_100.append(False)

                elif "LDBias(100)" in filename or "100mA" in filename:
                    ext_vals_100.append(ext_val)
                    bold_flags_100.append(bold)

                    ext_vals_80.append(None)
                    bold_flags_80.append(False)

        output_df = pd.DataFrame({
            "Chip Name": chip_names,
            "Extinction Ratio (80mA)": ext_vals_80,
            "Extinction Ratio (100mA)": ext_vals_100
        })
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_path = os.path.join(input_dir, f"extinction_values_{timestamp}.xlsx")

        writer = pd.ExcelWriter(output_path, engine='openpyxl')
        output_df.to_excel(writer, sheet_name="Results", index=False)

        workbook = writer.book
        sheet = workbook["Results"]

        for idx, (ext, is_bold) in enumerate(zip(ext_vals_80, bold_flags_80), start=2):
            if isinstance(ext, (float, int)) and is_bold:
                cell = sheet.cell(row=idx, column=2)  # Column 2 is "Extinction Ratio(80mA)"
                cell.font = Font(bold=True)

        for idx, (ext, is_bold) in enumerate(zip(ext_vals_100, bold_flags_100), start=2):
            if isinstance(ext, (float, int)) and is_bold:
                cell = sheet.cell(row=idx, column=3)  # Column 3 is "Extinction Ratio(100mA)"
                cell.font = Font(bold=True)

        workbook.save(output_path)
        print(f"File Saved to {output_path}")

    def get_spectrum_data(self, input_dir):
        print("\nGetting Spectrum Data...")
        if not os.path.isdir(input_dir):
            return

        required_string = "pkpow_pkwl_smsr_"
        chip_names = []

        # Get chip names
        for filename in os.listdir(input_dir):
            if filename.endswith(".xlsx") and "_EAM_" in filename and not filename.startswith('~'):
                try:
                    match = re.search(r'[A-Za-z]{2}\d{4}', filename)
                    if match:
                        chip_name = match.group()
                        chip_names.append(chip_name)
                except Exception as e:
                    print(f"Error extracting chip name from {filename}: {e}")
                    continue

        all_columns = set()
        chip_found = False
        spectrum_data = [[], [], [], [], [],
                         [], [], [], []]

        for chip_name in chip_names:
            spectrum_data[0].append(chip_name)

            for filename in os.listdir(input_dir):
                if filename.endswith(".csv") and required_string in filename and chip_name in filename:
                    if chip_found:
                        print("Error: duplicate chip name.")
                        break
                    chip_found = True
                    df = pd.read_csv(os.path.join(input_dir, filename))
                    for i in range(1, 9):
                        spectrum_data[i].append(df.iat[0, i - 1])

            if not chip_found:
                for i in range(1, 9):
                    spectrum_data[i].append(None)
            else:
                chip_found = False

        results_df = pd.DataFrame({
            "Chip Name": spectrum_data[0],
            "pkpow": spectrum_data[1],
            "pkwl": spectrum_data[2],
            "wl1": spectrum_data[3],
            "pow1": spectrum_data[4],
            "wl2": spectrum_data[5],
            "pow2": spectrum_data[6],
            "dwl": spectrum_data[7],
            "smsr": spectrum_data[8],
        })

        if not results_df.empty:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            output_path = os.path.join(input_dir, f"spectrum_values_{timestamp}.xlsx")
            results_df.to_excel(output_path, index=False)
            print(f"File Saved to {output_path}")

    def get_organized_data(self, input_dir):
        results = [[], [], [], [],
                   [], [], [], []]

        for filename in os.listdir(input_dir):
            name = None
            pd_curr = [None, None]
            thresh = None
            ext = None
            bold = False
            pkwl = None
            smsr = None
            date = None

            if filename.endswith(".xlsx") and "_LIV_" in filename and not filename.startswith('~'):
                # Chip Name
                filepath = os.path.join(input_dir, filename)
                try:
                    match = re.search(r'[A-Za-z]{2}\d{4}', filename)
                    if match:
                        name = match.group()
                except Exception as e:
                    print(f"Error extracting chip name from {filename}: {e}")
                    continue

                # Date
                date = extract_date_from_filename(filename)

                try:
                    df = pd.read_excel(filepath)

                    # PD Current from Laser Current
                    idx = 0
                    for l_current in [80, 100]:
                        row = df[df['SMU1_Ch2_Laser_Current_Set_mA'] == l_current]
                        if not row.empty:
                            pd_curr[idx] = row['SMU1_Ch1_PD_Current_Meas_mA'].iloc[0]
                        else:
                            print(f"No Data found for Laser {l_current} mA in {filename}")
                        idx += 1

                    # Intercept
                    filtered_df = df[(df['SMU1_Ch2_Laser_Current_Set_mA'] >= 30) &
                                     (df['SMU1_Ch2_Laser_Current_Set_mA'] <= 50)]

                    if not filtered_df.empty and len(filtered_df) >= 2:  # Need at least 2 points for regression
                        x_data = filtered_df['SMU1_Ch2_Laser_Current_Set_mA']
                        y_data = filtered_df['SMU1_Ch1_PD_Current_Meas_mA']

                        # Perform linear regression using our custom function
                        slope, intercept, r_value, p_value, std_err = linear_regression(x_data, y_data)

                        if slope != 0:  # Avoid division by zero
                            thresh = -intercept / slope
                        else:
                            print(
                                f"  Warning: Slope is zero for laser current intercept calculation in {filename}. Intercept cannot be determined.")
                    else:
                        print(
                            f"  Warning: Not enough data points (30-50mA Laser Current) for intercept calculation in {filename}")

                except FileNotFoundError:
                    print(f"Error: File not found - {filepath}")
                except KeyError as e:
                    print(f"Error: Missing expected column '{e}' in file {filename}. Please check column names.")
                except Exception as e:
                    print(f"An unexpected error occurred while processing {filename}: {e}")

                for filename in os.listdir(input_dir):

                    if filename.endswith(".xlsx") and "_EAM_" in filename and not filename.startswith(
                            '~') and name in filename:
                        ext_val = None
                        file_path = os.path.join(input_dir, filename)
                        df = pd.read_excel(file_path)
                        ext_i = df.iat[1, 5]
                        ext_f = df.iat[31, 5]

                        if pd.notna(ext_i) and pd.notna(ext_f):
                            numer = abs(ext_f)
                            denom = abs(ext_i)

                            # Determine if subtraction should occur
                            if numer > 0.111 and denom > 0.111:
                                numer -= 0.111
                                denom -= 0.111
                            else:
                                bold = True

                            # Attempt to calculate ext value
                            if denom != 0 and numer > 0:
                                ext_val = 10 * math.log10(numer / denom)

                            else:
                                print("Math Error")
                                ext_val = f"{ext_i},{ext_f}"

                        else:
                            print(f"Error: Unsuccessful Extraction of PD Current Values: {filename}")

                        if "LDBias(80)" in filename or "80mA" in filename:
                            ext = ext_val

                        break

                for filename in os.listdir(input_dir):
                    if filename.endswith(
                            ".csv") and "pkpow_pkwl_smsr_" in filename and name in filename and not filename.startswith(
                            '~'):
                        df = pd.read_csv(os.path.join(input_dir, filename))
                        pkwl = df.iat[0, 1]
                        smsr = df.iat[0, 7]

                        break

                results[0].append(name)
                results[1].append(thresh)
                results[2].append(pd_curr[0])
                results[3].append(pd_curr[1])
                results[4].append(ext)
                results[5].append(pkwl)
                results[6].append(smsr)
                results[7].append(date)

        results_df = pd.DataFrame({
            "Chip Name": results[0],
            "Threshold": results[1],
            "PD Current 80mA": results[2],
            "PD Current 100mA": results[3],
            "Extinction": results[4],
            "pkwl": results[5],
            "smsr": results[6],
            "Date": results[7]
        })

        if not results_df.empty:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            output_path = os.path.join(input_dir, f"final_values_{timestamp}.xlsx")
            results_df.to_excel(output_path, index=False)