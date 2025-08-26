import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Load the Excel
df = pd.read_excel("Datenbank Teilentladungsmessungen 2025.xlsx", header=None)
df.columns = [chr(65+i) for i in range(len(df.columns))]  # Excel style A, B, C...

# Create main window
root = tk.Tk()
root.title("Sorting GUI")
root.state("zoomed")  # Fullscreen

# ---- Dropdown frame (top) ----
top_frame = tk.Frame(root, pady=10)
top_frame.pack(side="top", fill="x")

tk.Label(top_frame, text="Achsh.mm:").pack(side="left", padx=5)
f_combo = ttk.Combobox(top_frame, state="readonly")
f_combo.pack(side="left", padx=10)

tk.Label(top_frame, text="Iso-Typ:").pack(side="left", padx=5)
h_combo = ttk.Combobox(top_frame, state="readonly", width=20)
h_combo.pack(side="left", padx=10)

tk.Label(top_frame, text="V:").pack(side="left", padx=5)
k_combo = ttk.Combobox(top_frame, state="readonly")
k_combo.pack(side="left", padx=10)

# ---- Output frame (center) ----
output_frame = tk.Frame(root, pady=20)
output_frame.pack(fill="both", expand=True)

tree = ttk.Treeview(output_frame, columns=("Typpr. T Stator \n K ", "Ph-Ph \n U B / 400V", "Ph-PE \n U B / 400V", "Typpr. T Stator \n K"), show="headings")
for col in ["Typpr. T Stator \n K ", "Ph-Ph \n U B / 400V", "Ph-PE \n U B / 400V", "Typpr. T Stator \n K"]:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)
tree.pack(fill="both", expand=True)

# ---- Stats label (bottom of output) ----
stats_label = tk.Label(output_frame, text="", font=("Arial", 12), pady=10, fg="blue")
stats_label.pack(side="bottom")

# Global filtered dataframe
filtered_df = pd.DataFrame()

# ---- Functions ----
def update_h(event):
    f_val = f_combo.get()
    if not f_val:
        return
    h_options = df[df["F"].astype(str) == f_val]["H"].dropna().astype(str).unique().tolist()
    h_combo["values"] = sorted(h_options)
    h_combo.set("")
    k_combo.set("")
    k_combo["values"] = []

def update_k(event):
    f_val = f_combo.get()
    h_val = h_combo.get()
    if not f_val or not h_val:
        return
    k_options = df[(df["F"].astype(str) == f_val) & (df["H"].astype(str) == h_val)]["K"].dropna().astype(str).unique().tolist()
    k_combo["values"] = sorted(k_options)
    k_combo.set("")

def show_output():
    global filtered_df
    f_val = f_combo.get()
    h_val = h_combo.get()
    k_val = k_combo.get()

    if not f_val or not h_val or not k_val:
        messagebox.showwarning("Missing Selection", "Please select values for F, H, and K before output.")
        return

    filtered_df = df.copy()
    filtered_df = filtered_df[filtered_df["F"].astype(str) == f_val]
    filtered_df = filtered_df[filtered_df["H"].astype(str) == h_val]
    filtered_df = filtered_df[filtered_df["K"].astype(str) == k_val]

    filtered_df = filtered_df[["O", "P", "Q", "R"]]

    # Clear table
    for row in tree.get_children():
        tree.delete(row)

    # Insert rows
    for _, row in filtered_df.iterrows():
        tree.insert("", "end", values=list(row))

    # ---- Calculate stats ----
    total = len(filtered_df)
    null_counts = filtered_df.isna().sum(axis=1)  # number of nulls per row
    all_null = (null_counts == 4).sum()
    one_null = (null_counts == 1).sum()
    more_null = (null_counts >= 2).sum()

    stats_text = f"Total outputs: {total} | Completely null rows: {all_null} | Rows with 1 null: {one_null} | Rows with 2+ nulls: {more_null}"
    stats_label.config(text=stats_text)

def save_excel():
    if filtered_df.empty:
        messagebox.showwarning("No Data", "No filtered data to save.")
        return

    # ---- File chooser for base name ----
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    # ---- First file: only O-P-Q-R renamed ----
    save_df = filtered_df.rename(columns={
        "O": "Typpr. T Stator ",
        "P": "Ph-Ph",
        "Q": "Ph-PE",
        "R": "Typpr. T Stator"
    })
    save_df.to_excel(file_path, index=False)

    # ---- Second file: full filtered output ----
    # keep all columns for the same filtered rows
    full_output = df.copy()
    full_output = full_output[
        (full_output["F"].astype(str) == f_combo.get()) &
        (full_output["H"].astype(str) == h_combo.get()) &
        (full_output["K"].astype(str) == k_combo.get())
    ]

    full_file_path = file_path.replace(".xlsx", "_full.xlsx")
    full_output.to_excel(full_file_path, index=False)

    messagebox.showinfo("Saved", f"Saved:\n{file_path}\n{full_file_path}")


# ---- Buttons ----
output_btn = tk.Button(top_frame, text="Output", command=show_output, bg="lightgreen")
output_btn.pack(side="left", padx=20)

bottom_frame = tk.Frame(root, pady=10)
bottom_frame.pack(side="bottom", fill="x")

save_btn = tk.Button(bottom_frame, text="Save Excel", command=save_excel, bg="lightblue")
save_btn.pack(side="right", padx=20, pady=10)

# ---- Bind events ----
f_combo.bind("<<ComboboxSelected>>", update_h)
h_combo.bind("<<ComboboxSelected>>", update_k)

# ---- Initialize F values ----
f_combo["values"] = sorted(df["F"].dropna().astype(str).unique().tolist())

root.mainloop()
