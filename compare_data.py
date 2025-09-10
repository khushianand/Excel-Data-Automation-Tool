import pandas as pd
import os
import tkinter.messagebox as messagebox


def compare_excel_data(
    tracking_file_path: str,
    tracking_sheet_name: str,
    raw_file_path: str,
    raw_sheet_name: str,
    status_callback,
    output_file_path: str = None,
    area_filter: str = None,
    sites_filter: list = None,
):
    if not output_file_path:
        status_callback("Error: No output file path provided.")
        messagebox.showerror("Missing Output File", "Please select a file path to save the output Excel file.")
        return

    status_callback("--- Starting Excel File Comparison ---")

    new_data_sheet_name = "new data"
    old_data_sheet_name = "old data"

    # Columns we try to present in the output; we'll only include those available
    tracking_common_cols = [
        "Plugin ID",
        "Site",
        "Product",
        "Hostname",
        "Element",
        "OWNER",
        "First Found",
        "Name",
        "Risk",
        "CVE",
        "Description",
        "NESUS Solution",
        "See Also",
        "Host (IP address)",
        "Protocol",
        "Port",
        "Vulnerability State",
    ]

    # Map from raw file columns to the tracking schema
    raw_to_tracking_col_map = {
        "Plugin ID": "Plugin ID",
        "Site": "Site",
        "Product": "Product",
        "Hostname": "Hostname",
        "Element": "Element",
        "Owner": "OWNER",
        "First Found": "First Found",
        "Name": "Name",
        "Risk": "Risk",
        "CVE": "CVE",
        "Description": "Description",
        "Solution": "NESUS Solution",
        "See Also": "See Also",
        "IP Address": "Host (IP address)",
        "Protocol": "Protocol",
        "Port": "Port",
        "Vulnerability State": "Vulnerability State",
    }

    comparison_key_cols = ["Name", "CVE", "Host (IP address)", "Port"]

    try:
        status_callback(f"Loading Tracking file: '{tracking_file_path}' (Sheet: '{tracking_sheet_name}')")
        tracking_df = pd.read_excel(tracking_file_path, sheet_name=tracking_sheet_name)
        tracking_df.columns = tracking_df.columns.str.strip()

        status_callback(f"Loading Raw file: '{raw_file_path}' (Sheet: '{raw_sheet_name}')")
        raw_df = pd.read_excel(raw_file_path, sheet_name=raw_sheet_name)
        raw_df.columns = raw_df.columns.str.strip()
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

    # Ensure tracking contains the columns we will use
    missing_cols_tracking = [col for col in tracking_common_cols if col not in tracking_df.columns]
    if missing_cols_tracking:
        msg = f"Tracking file missing columns (these will be filled if possible): {', '.join(missing_cols_tracking)}"
        status_callback(msg)
        # Not fatal: we'll continue but notify the user. Only fatal if comparison keys are missing.

    # Apply filters if present
    if area_filter:
        if "Area" in tracking_df.columns:
            tracking_df = tracking_df[tracking_df['Area'] == area_filter].copy()
            status_callback(f"Tracking data filtered by Area '{area_filter}'.")
        else:
            msg = "Error: 'Area' column not found in Tracking file."
            status_callback(msg)
            messagebox.showerror("Column Error", msg)
            return

    if sites_filter:
        if "Site" in tracking_df.columns:
            tracking_df = tracking_df[tracking_df['Site'].isin(sites_filter)].copy()
            status_callback(f"Tracking data filtered by Sites: {', '.join(sites_filter)}.")
        else:
            msg = "Error: 'Site' column not found in Tracking file."
            status_callback(msg)
            messagebox.showerror("Column Error", msg)
            return

    if tracking_df.empty:
        status_callback("Warning: No data for the selected filters. Proceeding with empty tracking data.")

    # Check raw columns required for mapping (we only require the keys for comparison)
    required_raw_cols = [k for k in raw_to_tracking_col_map.keys()]
    missing_cols_raw = [col for col in required_raw_cols if col not in raw_df.columns]
    if missing_cols_raw:
        msg = f"Raw file missing columns: {', '.join(missing_cols_raw)}"
        status_callback(msg)
        messagebox.showerror("Column Error", msg)
        return

    # Normalize and map raw columns into tracking schema
    raw_df_processed = raw_df.rename(columns=raw_to_tracking_col_map)

    # Make sure comparison key columns exist after mapping
    for key in comparison_key_cols:
        if key not in raw_df_processed.columns and key not in tracking_df.columns:
            msg = f"Error: Comparison key column '{key}' not available in either file."
            status_callback(msg)
            messagebox.showerror("Column Error", msg)
            return

    raw_subset = raw_df_processed[comparison_key_cols].fillna('').astype(str).apply(lambda x: x.str.strip())
    tracking_subset = tracking_df[comparison_key_cols].fillna('').astype(str).apply(lambda x: x.str.strip()) if not tracking_df.empty else pd.DataFrame(columns=comparison_key_cols)

    status_callback("Comparing data...")
    tracking_keys = tracking_subset.agg('___'.join, axis=1).unique() if not tracking_subset.empty else []
    raw_keys = raw_subset.agg('___'.join, axis=1)
    is_old_data_mask = raw_keys.isin(tracking_keys)

    # Build output frames using columns available in raw_df_processed or tracking_common_cols
    available_output_cols = [c for c in tracking_common_cols if c in raw_df_processed.columns]

    new_data_df = raw_df_processed[~is_old_data_mask].copy()
    old_data_df = raw_df_processed[is_old_data_mask].copy()

    # If expected output columns are missing, add them with blank values so outputs are consistent
    for df in (new_data_df, old_data_df):
        for col in tracking_common_cols:
            if col not in df.columns:
                df[col] = ""
        # Keep only tracking_common_cols ordering
        df = df[tracking_common_cols]

    status_callback(f"Writing results to '{output_file_path}'...")
    try:
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            if not new_data_df.empty:
                new_data_df[tracking_common_cols].to_excel(writer, sheet_name=new_data_sheet_name, index=False)
                status_callback(f"{len(new_data_df)} new rows written to '{new_data_sheet_name}'.")
            else:
                status_callback("No new rows found.")

            if not old_data_df.empty:
                old_data_df[tracking_common_cols].to_excel(writer, sheet_name=old_data_sheet_name, index=False)
                status_callback(f"{len(old_data_df)} old rows written to '{old_data_sheet_name}'.")
            else:
                status_callback("No old rows found.")

        full_path = os.path.abspath(output_file_path)
        status_callback(f"Comparison done. Output saved to: {full_path}")
        status_callback("✅ Check 'new data' and 'old data' sheets in the Excel file.")
        messagebox.showinfo("Success", f"Comparison complete!\nResults saved to: {full_path}")
    except Exception as e:
        status_callback(f"Error writing output Excel file: {e}")
        messagebox.showerror("Write Error", str(e))