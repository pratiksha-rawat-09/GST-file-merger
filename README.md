```python 
# GST-file-merger
import os
import re
import pandas as pd
from openpyxl import load_workbook
from tkinter import filedialog, messagebox, Tk, Label, Button

# Define how many header/info rows to preserve per sheet
sheet_headers = {
    'B2B': 6,
    'B2BA': 7,
    'CDNR': 6,
    'CDNRA': 7,
    'IMPG': 6,
    'IMPGSEZ': 6,
    'NIL': 6,
    'HSN': 6,
    'DOCS': 6
}

def extract_number(filename):
    """Extract numeric part from filename for proper sorting"""
    match = re.search(r'\d+', filename)
    return int(match.group()) if match else float('inf')

def extract_data(file_path, sheet_name, header_rows):
    """Read full sheet and extract rows below header"""
    try:
        df_full = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        df_data = df_full.iloc[header_rows:]  # Keep everything below header
        df_data.reset_index(drop=True, inplace=True)
        return df_data, len(df_data)
    except Exception:
        return pd.DataFrame(), 0

def merge_gstr2a_files():
    folder_path = filedialog.askdirectory(title="Select Folder with GSTR-2A Excel Files")
    if not folder_path:
        return

    # Sort files numerically based on filename
    files = sorted(
        [f for f in os.listdir(folder_path) if f.endswith('.xlsx')],
        key=extract_number
    )

    if not files:
        messagebox.showinfo("No Files", "No Excel files found in the selected folder.")
        return

    template_path = os.path.join(folder_path, files[0])
    wb = load_workbook(template_path)

    summary_data = []

    for sheet_name in wb.sheetnames:
        if sheet_name == 'Read Me':
            continue

        ws = wb[sheet_name]
        header_rows = sheet_headers.get(sheet_name, 6)

        # Clear existing data below headers
        if ws.max_row > header_rows:
            ws.delete_rows(header_rows + 1, ws.max_row - header_rows)

        total_rows = 0

        for file in files:
            file_path = os.path.join(folder_path, file)
            df, count = extract_data(file_path, sheet_name, header_rows)
            total_rows += count

            # Tag each row with source file name
            for row in df.itertuples(index=False):
                row_data = list(row)
                row_data.append(file)  # Add filename as last column
                ws.append(row_data)

        summary_data.append({
            'Sheet Name': sheet_name,
            'Total Rows Added': total_rows
        })

    # Add Summary Sheet
    if 'Summary' in wb.sheetnames:
        wb.remove(wb['Summary'])
    summary_sheet = wb.create_sheet(title='Summary')
    summary_sheet.append(['Sheet Name', 'Total Rows Added'])
    for entry in summary_data:
        summary_sheet.append([
            entry['Sheet Name'],
            entry['Total Rows Added']
        ])

    output_path = os.path.join(folder_path, 'GSTR2A_Consolidated.xlsx')
    wb.save(output_path)
    messagebox.showinfo("Success", f"Merged file saved at:\n{output_path}")
```

# GUI setup
root = Tk()
root.title("GSTR-2A Merger Utility")
root.geometry("400x180")

Label(root, text="Merge GSTR-2A Excel Files in Correct Order", pady=20).pack()
Button(root, text="Select Folder & Merge", command=merge_gstr2a_files).pack()

root.mainloop()
