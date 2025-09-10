import pandas as pd
import os
import tkinter.messagebox as messagebox


def add_site_and_product(
    file1_path: str,
    sheet1_name: str,
    file2_path: str,
    sheet2_name: str,
    status_callback,
):
    status_callback("--- Starting 'Add Site & Product' Process ---")

    ip_col_name = "IP Address"
    site_col_name = "Site"
    product_col_name = "Product"
    output_sheet_name = "Site&Product"

    try:
        status_callback(f"Loading file 1: '{file1_path}' (Sheet: '{sheet1_name}')")
        df1 = pd.read_excel(file1_path, sheet_name=sheet1_name)
        df1.columns = df1.columns.str.strip()

        status_callback(f"Loading file 2: '{file2_path}' (Sheet: '{sheet2_name}')")
        df2 = pd.read_excel(file2_path, sheet_name=sheet2_name)
        df2.columns = df2.columns.str.strip()
        status_callback("Excel files loaded successfully.")
    except FileNotFoundError as e:
        status_callback(f"Error: One of the files was not found. {e}")
        messagebox.showerror("File Error", f"File not found: {e}")
        return
    except ValueError as e:
        status_callback(f"Error: Sheet name not found. {e}")
        messagebox.showerror("Sheet Error", f"Sheet name not found: {e}")
        return
    except Exception as e:
        status_callback(f"Error loading Excel files: {e}")
        messagebox.showerror("Loading Error", str(e))
        return

    if ip_col_name not in df1.columns:
        msg = f"Error: '{ip_col_name}' column not found in File 1 sheet '{sheet1_name}'."
        status_callback(msg)
        messagebox.showerror("Column Error", msg)
        return

    if not set([ip_col_name, site_col_name, product_col_name]).issubset(set(df2.columns)):
        msg = f"Error: One or more required columns ('{ip_col_name}', '{site_col_name}', '{product_col_name}') not found in File 2 sheet '{sheet2_name}'."
        status_callback(msg)
        messagebox.showerror("Column Error", msg)
        return

    status_callback("Processing unique IP addresses from File 1...")
    unique_ips_df = pd.DataFrame(df1[ip_col_name].dropna().astype(str).str.strip().unique(), columns=[ip_col_name])

    if unique_ips_df.empty:
        status_callback("Warning: No unique IP addresses found in File 1. The output sheet will be empty.")
        final_df = unique_ips_df
    else:
        status_callback(f"Found {len(unique_ips_df)} unique IP addresses. Merging with data from File 2...")

        df2_subset = df2[[ip_col_name, site_col_name, product_col_name]].dropna(subset=[ip_col_name]).astype(str)
        df2_subset = df2_subset.drop_duplicates(subset=[ip_col_name]).set_index(ip_col_name)

        final_df = unique_ips_df.merge(df2_subset, left_on=ip_col_name, right_index=True, how='left')

        final_df[[site_col_name, product_col_name]] = final_df[[site_col_name, product_col_name]].fillna('Not Found')

    status_callback(f"Writing results to a new sheet named '{output_sheet_name}' in '{file1_path}'...")
    try:
        with pd.ExcelWriter(file1_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            final_df.to_excel(writer, sheet_name=output_sheet_name, index=False)

        full_path = os.path.abspath(file1_path)
        status_callback("✅ Process complete! 'Site' and 'Product' data has been added.")
        status_callback(f"Check the new sheet '{output_sheet_name}' in file: {full_path}")
        messagebox.showinfo("Success", f"Process complete!\nThe new sheet '{output_sheet_name}' has been created in '{os.path.basename(file1_path)}' with the data.")

    except Exception as e:
        status_callback(f"Error writing output to Excel file: {e}")
        messagebox.showerror("Write Error", str(e))
