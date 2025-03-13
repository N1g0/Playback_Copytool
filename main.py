import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import zipfile
import os
from tkinter import messagebox, ttk

# Don´t copy the cells in in row of 1_TIme, ec5_uuid, created_at, uploaded_at, title

# Simple Drag & Drop Window
def on_drop(event):
    file_path = event.data.strip('{"}')
    if file_path.endswith(".zip"):
        process_zip(file_path)
    elif file_path.endswith(".csv"):
        process_csv(file_path)
    else:
        messagebox.showerror("Invalid File", "Please drop a .zip or .csv file")


# Unzip and Process CSV
def process_zip(zip_path):
    extract_folder = os.path.splitext(zip_path)[0] + "_unzipped"
    os.makedirs(extract_folder, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)

    csv_file = [f for f in os.listdir(extract_folder) if f.endswith(".csv")][0]
    csv_path = os.path.join(extract_folder, csv_file)

    process_csv(csv_path)
    zip_result(zip_path, extract_folder)


# Process CSV File
def process_csv(csv_path):
    try:
        df = pd.read_csv(csv_path, sep=';')
    except pd.errors.ParserError:
        df = pd.read_csv(csv_path, sep=',')
    preview_csv(df)
    df = fill_yes_rows(df)
    if ';' in open(csv_path).read(1024):
        sep = ';'
    else:
        sep = ','
    df.to_csv(csv_path, sep=sep, index=False)
    messagebox.showinfo("Success", f"CSV has been modified: {csv_path}")


# Fill Rows with "Yes" in Column "2_NC"
def fill_yes_rows(df):
    protected_columns = {"1_TIme", "ec5_uuid", "created_at", "uploaded_at", "title"}

    if "2_NC" in df.columns:
        yes_rows = df["2_NC"] == "Yes"
        non_yes_rows = df[~yes_rows]

        for index in df[yes_rows].index:
            next_valid_row = non_yes_rows.loc[non_yes_rows.index > index].head(1)
            if not next_valid_row.empty:
                target_index = next_valid_row.index[0]

                # Copy values **except** for protected columns
                for col in df.columns:
                    if col not in protected_columns:
                        df.at[index, col] = df.at[target_index, col]

                # Restore the "Yes" value and mark highlights
                df.at[index, "2_NC"] = "Yes"
                df.at[target_index, "highlight"] = "red"
                df.at[index, "highlight"] = "green"

    return df


# Re-Zip Folder
def zip_result(zip_path, folder):
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(folder):
            for file in files:
                zipf.write(os.path.join(root, file), file)
    messagebox.showinfo("Success", "CSV has been modified and re-zipped")


# Preview CSV Window with Table, Scrollbars, and Highlighting
def preview_csv(df):
    preview = tk.Toplevel(root)
    preview.title("CSV Preview")
    preview.geometry("800x400")

    frame = tk.Frame(preview)
    frame.pack(fill=tk.BOTH, expand=True)

    tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
    scrollbar_y = tk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scrollbar_x = tk.Scrollbar(frame, orient="horizontal", command=tree.xview)

    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor=tk.W)

    for index, row in df.head(20).iterrows():
        values = list(row.drop("highlight", errors="ignore"))
        tags = []
        if "highlight" in row and row["highlight"] == "red":
            tags.append("highlight_red")
            tree.tag_configure("highlight_red", background="red")
        elif "highlight" in row and row["highlight"] == "green":
            tags.append("highlight_green")
            tree.tag_configure("highlight_green", background="green")

        tree.insert("", "end", values=values, tags=tags)

    ok_btn = tk.Button(preview, text="OK", command=preview.destroy)
    ok_btn.pack(pady=10)


# Main Window
root = TkinterDnD.Tk()
root.title("Drag & Drop CSV Tool")
root.geometry("400x200")
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)

label = tk.Label(root, text="Drag your .zip or .csv file here", font=("Arial", 16))
label.pack(expand=True, pady=50)

root.mainloop()
