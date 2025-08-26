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

tree = ttk.Treeview(output_frame, columns=("Typpr. T Stator ", "Ph-Ph", "Ph-PE", "Typpr. T Stator"), show="headings")
for col in ["Typpr. T Stator ", "Ph-Ph", "Ph-PE", "Typpr. T Stator"]:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)
tree.pack(fill="both", expand=True)

# Global filtered dataframe
filtered_df = pd.DataFrame()

# ---- Functions ----
def update_h(event):
    """Update H dropdown when F is selected"""
    f_val = f_combo.get()
    if not f_val:
        return
    h_options = df[df["F"].astype(str) == f_val]["H"].dropna().astype(str).unique().tolist()
    h_combo["values"] = sorted(h_options)
    h_combo.set("")  # reset selection
    k_combo.set("")  # reset K too
    k_combo["values"] = []  # clear K options

def update_k(event):
    """Update K dropdown when H is selected"""
    f_val = f_combo.get()
    h_val = h_combo.get()
    if not f_val or not h_val:
        return
    k_options = df[(df["F"].astype(str) == f_val) & (df["H"].astype(str) == h_val)]["K"].dropna().astype(str).unique().tolist()
    k_combo["values"] = sorted(k_options)
    k_combo.set("")  # reset selection

def show_output():
    global filtered_df
    f_val = f_combo.get()
    h_val = h_combo.get()
    k_val = k_combo.get()

    if not f_val or not h_val or not k_val:
        messagebox.showwarning("Missing Selection", "Please select values for F, H, and K before output.")
        return

    # Filter dataframe
    filtered_df = df.copy()
    filtered_df = filtered_df[filtered_df["F"].astype(str) == f_val]
    filtered_df = filtered_df[filtered_df["H"].astype(str) == h_val]
    filtered_df = filtered_df[filtered_df["K"].astype(str) == k_val]

    # Only keep O-P-Q-R
    filtered_df = filtered_df[["O", "P", "Q", "R"]]

    # Clear table
    for row in tree.get_children():
        tree.delete(row)

    # Insert new rows
    for _, row in filtered_df.iterrows():
        tree.insert("", "end", values=list(row))

def save_excel():
    if filtered_df.empty:
        messagebox.showwarning("No Data", "No filtered data to save.")
        return

    # Rename columns before saving
    save_df = filtered_df.rename(columns={
        "O": "Typpr. T Stator ",
        "P": "Ph-Ph",
        "Q": "Ph-PE",
        "R": "Typpr. T Stator"
    })

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        save_df.to_excel(file_path, index=False)
        messagebox.showinfo("Saved", f"File saved as {file_path}")

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
