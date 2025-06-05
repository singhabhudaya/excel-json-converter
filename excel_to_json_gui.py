import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def select_file():
    """Open a file dialog to pick an Excel file, then convert to JSON."""
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    if not file_path:
        return

    try:
        # Read the first sheet of the Excel file into a DataFrame.
        df = pd.read_excel(file_path, sheet_name=0)

        # To capture EVERY sheet instead, uncomment & use this:
        # all_sheets = pd.read_excel(file_path, sheet_name=None)
        # combined = {}
        # for sheet_name, sheet_df in all_sheets.items():
        #     combined[sheet_name] = sheet_df.to_dict(orient="records")
        # json_data = combined

        # For single‐sheet JSON:
        json_data = df.to_dict(orient="records")

        # Ask user where to save the JSON
        save_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile=os.path.splitext(os.path.basename(file_path))[0] + ".json"
        )
        if not save_path:
            return

        # Write JSON to disk (pretty‐print with indent=2)
        with open(save_path, "w", encoding="utf-8") as f:
            import json
            json.dump(json_data, f, indent=2, ensure_ascii=False)

        messagebox.showinfo("Success", f"JSON saved to:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert:\n{e}")

def main():
    root = tk.Tk()
    root.title("Excel → JSON Converter")
    root.geometry("350x120")
    root.resizable(False, False)

    lbl = tk.Label(root, text="Click below to select an Excel file:", pady=10)
    lbl.pack()

    btn = tk.Button(root, text="Select Excel and Convert", command=select_file, padx=10, pady=5)
    btn.pack()

    root.mainloop()

if __name__ == "__main__":
    main()
