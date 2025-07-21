import pandas as pd
import numpy as np # For pd.notnull and potential numeric operations

def get_state(row1_j):
    """Convert fluid state to Calculation Sheet abbreviation."""
    if row1_j in ["VAPOR", "GAS"]:
        return "V"
    elif row1_j == "STEAM":
        return "S"
    elif row1_j == "LIQUID":
        return "L"
    else:
        return ""

def get_ratio(psv_type):
    """Return ratio by PSV type."""
    if psv_type == "C":
        return 0.1
    elif psv_type == "B":
        return 0.3
    elif psv_type == "P":
        return 1.0
    else:
        return ""

def get_rupture_disk(row1_AA):
    """Check if 'rupture disk' is in remark."""
    if isinstance(row1_AA, str) and "rupture disk" in row1_AA.lower():
        return "Y"
    else:
        return "N"

def get_sum_bp(row1_O):
    """Sum numbers in 'X / Y' string format."""
    try:
        return sum([float(x) for x in str(row1_O).split("/")])
    except:
        return ""

def get_left_bp(row1_O):
    """Get first number in 'X / Y' string format."""
    try:
        return float(str(row1_O).split("/")[0])
    except:
        return ""

def convert_data_to_calc_sheet(data_sheet_df, calc_sheet_template_df):
    """
    Convert Data Sheet to Calculation Sheet format.
    Args:
        data_sheet_df (pd.DataFrame): Data Sheet.
        calc_sheet_template_df (pd.DataFrame): Calculation Sheet template.
    Returns:
        pd.DataFrame: Converted Calculation Sheet.
    """
    # Mapping: Calculation Sheet col idx: (Data Sheet row offset, col idx, function)
    # row offset: 0 = first row, 1 = second row in each record pair
    mapping = {
        1:   (0, 0, None),
        2:   (0, 3, None),
        3:   (1, 0, None),
        4:   (None, None, lambda r1, r2: 1),
        5:   (1, 16, None),
        6:   (None, None, lambda r1, r2: get_ratio(r2[16])),
        7:   (None, None, lambda r1, r2: "CS"),
        8:   (None, None, lambda r1, r2: get_rupture_disk(r1[26] if len(r1) > 26 else "")),
        9:   (None, None, lambda r1, r2: "R"),
        11:  (1, 9, None),
        12:  (0, 9, lambda r1, r2: get_state(r1[9])),
        13:  (0, 10, None),
        14:  (1, 22, lambda r1, r2: r2[22]*1000 if get_state(r1[9])=="L" and pd.notnull(r2[22]) else ""),
        15:  (0, 21, None),
        16:  (0, 22, lambda r1, r2: r1[22] if get_state(r1[9]) in ["V","S"] else ""),
        17:  (0, 23, lambda r1, r2: r1[23] if get_state(r1[9]) in ["V","S"] else ""),
        18:  (1, 23, lambda r1, r2: r2[23] if get_state(r1[9]) in ["V","S"] else ""),
        20:  (0, 13, None),
        21:  (0, 17, None),
        22:  (0, 20, None),
        23:  (0, 14, lambda r1, r2: get_sum_bp(r1[14])),
        24:  (0, 14, lambda r1, r2: get_left_bp(r1[14]))
    }

    result_df = calc_sheet_template_df.copy()
    num_records = data_sheet_df.shape[0] // 2 # Each record is two rows

    for rec_idx in range(num_records):
        row1 = data_sheet_df.iloc[rec_idx*2]
        row2 = data_sheet_df.iloc[rec_idx*2+1]

        # Skip if Tag No. is empty
        if pd.isnull(row1[0]) or str(row1[0]).strip() == "":
            continue

        # Prepare new column for result_df
        new_col_data = pd.Series([""] * calc_sheet_template_df.shape[0], index=calc_sheet_template_df.index)

        for calc_col_idx, (ds_row_offset, ds_col_idx, func) in mapping.items():
            value_to_map = ""
            if func:
                value_to_map = func(row1, row2)
            elif ds_row_offset is not None and ds_col_idx is not None:
                src_row = row1 if ds_row_offset == 0 else row2
                if ds_col_idx < len(src_row):
                    value_to_map = src_row[ds_col_idx]
            if calc_col_idx < len(new_col_data):
                new_col_data[calc_col_idx] = value_to_map
        result_df[result_df.shape[1]] = new_col_data
    return result_df

if __name__ == '__main__':
    # Only runs when this file is executed directly (for testing)
    print("--- Data Sheet to Calculation Sheet conversion example ---")
    try:
        data_sheet_input_df = pd.read_excel("Data Sheet.xlsm", sheet_name="FORM", header=None)
        calc_sheet_template_input_df = pd.read_excel("Calculation Sheet.xlsm", sheet_name="PSV", header=None)
        converted_calc_df = convert_data_to_calc_sheet(data_sheet_input_df, calc_sheet_template_input_df)
        output_filename = "Calculation_Sheet_from_Data.xlsx"
        converted_calc_df.to_excel(output_filename, index=False, header=False)
        print(f"Conversion complete. Output: {output_filename}")
        print("\nFirst few rows of converted Calculation Sheet:")
        print(converted_calc_df.head())
    except FileNotFoundError:
        print("ERROR: 'Data Sheet.xlsm' and 'Calculation Sheet.xlsm' must exist in the same directory.")
    except Exception as e:
        print(f"ERROR: {e}")