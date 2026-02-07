import os
import pandas as pd
import gspread


def get_pii_columns_from_xlsform(key_path: str, xlsform_link: str, survey_sheet: str = "survey") -> list[str]:
    """Fetch XLSForm 'survey' sheet and return question names marked as dataset == 'pii'."""
    gc = gspread.service_account(filename=key_path)
    sh = gc.open_by_url(xlsform_link)

    try:
        ws = sh.worksheet(survey_sheet)
    except gspread.exceptions.WorksheetNotFound as e:
        raise ValueError(f"Worksheet '{survey_sheet}' not found in the XLSForm.") from e

    xlsfrm = pd.DataFrame(ws.get_all_records())

    # Validate expected columns
    required = {"dataset", "name"}
    missing = required - set(map(str.lower, xlsfrm.columns))
    # If columns might have different casing, normalize:
    cols_lower = {c.lower(): c for c in xlsfrm.columns}
    if "dataset" not in cols_lower or "name" not in cols_lower:
        raise ValueError("XLSForm 'survey' sheet must contain columns 'dataset' and 'name'.")

    dataset_col = cols_lower["dataset"]
    name_col = cols_lower["name"]

    piicol = xlsfrm[xlsfrm[dataset_col].astype(str).str.lower() == "pii"][name_col].tolist()
    return piicol


def mask_excel_file(
    input_file: str,
    pii_columns: list[str],
    replacement: str = "PII Field",
    output_ext: str = ".qa317",
) -> str:
    """
    Mask values in any column whose name matches pii_columns across all sheets of an Excel file.
    Writes an anonymized xlsx, then renames to output_ext.
    Returns the final output path.
    """
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    base, ext = os.path.splitext(input_file)
    if ext.lower() not in [".xlsx", ".xlsm", ".xls"]:
        raise ValueError(f"Expected an Excel file, got: {ext}")

    temp_output = f"{base}_anonymized.xlsx"
    final_output = f"{base}_anonymized{output_ext}"

    xls = pd.ExcelFile(input_file)
    updated_sheets = {}

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Only touch columns that actually exist
        cols_to_mask = [c for c in df.columns if c in pii_columns]
        for col in cols_to_mask:
            df.loc[df[col].notna(), col] = replacement

        updated_sheets[sheet_name] = df

    with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
        for sheet_name, df in updated_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Replace if already exists
    if os.path.exists(final_output):
        os.remove(final_output)
    os.rename(temp_output, final_output)

    return final_output


def main():

    print(
        """
Before you continue, make sure the following are ready:

1. A Google service account JSON key file
   - Ask the officers to provide this file.

2. The XLSForm Google Sheet link
   - The service account email
     qa-officers@dplt-472317.iam.gserviceaccount.com
     MUST have access to the XLSForm.

3. XLSForm structure requirements
   - The 'survey' sheet must exist.
   - It must contain a column named 'dataset'.
   - Fields marked as 'pii' in the 'dataset' column will be anonymized.

--------------------------------------------------
"""
    )
    key_path = input("Enter path to your service account key JSON file: ").strip()
    xlsform_link = input("Enter the Google Sheets (XLSForm) URL: ").strip()
    input_file = input("Enter path to your Raw dataset: ").strip()

    pii_cols = get_pii_columns_from_xlsform(key_path, xlsform_link)
    out_path = mask_excel_file(input_file, pii_cols)

    print(f"PII masking completed successfully. Output: {out_path}")


if __name__ == "__main__":
    main()
