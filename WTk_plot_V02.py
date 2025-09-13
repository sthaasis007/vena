import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, Frame, Button, Text, Scrollbar, RIGHT, Y, BOTH, LEFT
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re

# ---------------- CSV LOADING & CLEANING ----------------
def load_csv(file_path):
    # Load with correct separator and decimal
    df = pd.read_csv(file_path, sep=";", decimal=",", dtype=str)

    # Normalize column names (strip spaces, BOM, lowercase)
    df.columns = [c.strip().replace("\ufeff", "") for c in df.columns]

    # Convert numeric columns
    for col in df.columns:
        if col.lower() in ["date", "time", "millisecond", "datano"]:
            continue
        # Clean numbers: remove spaces, replace commas with dot
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        # Convert to numeric (coerce errors to NaN)
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Combine Date + Time into single datetime
    if "Date" in df.columns and "Time" in df.columns:
        df["Timestamp"] = pd.to_datetime(df["Date"] + " " + df["Time"], errors="coerce")
    else:
        raise ValueError("CSV must contain 'Date' and 'Time' columns")

    return df

# ---------------- METRICS CALCULATION ----------------
def analyze_segments(df, t_no_load_start, t_no_load_end, t_full_load_start, t_full_load_end):
    results = {}

    # Select segments
    seg_no = df[(df["Timestamp"] >= t_no_load_start) & (df["Timestamp"] <= t_no_load_end)]
    seg_full = df[(df["Timestamp"] >= t_full_load_start) & (df["Timestamp"] <= t_full_load_end)]

    for name, seg in {"No-Load": seg_no, "Full-Load": seg_full}.items():
        res = {}
        # RMS Imbalance
        Irms_cols = ["Irms-1", "Irms-2", "Irms-3"]
        Urms_cols = ["Urms-1", "Urms-2", "Urms-3"]
        PF_cols = ["PF-1", "PF-2", "PF-3"]

        def imbalance(vals):
            return (np.max(vals) - np.min(vals)) / np.mean(vals) * 100 if len(vals) else np.nan

        res["Irms imbalance %"] = imbalance(seg[Irms_cols].mean())
        res["Urms imbalance %"] = imbalance(seg[Urms_cols].mean())
        res["PF imbalance"] = np.max(seg[PF_cols].mean()) - np.min(seg[PF_cols].mean())

        # Power components
        res["P_total"] = seg["P-SIGMA"].mean()
        res["Q_total"] = seg["Q-SIGMA"].mean()
        res["S_total"] = seg["S-SIGMA"].mean()
        res["PF_total"] = seg["PF-SIGMA"].mean()
        res["tan_phi"] = res["Q_total"] / res["P_total"] if res["P_total"] != 0 else np.nan

        # Distortion
        res["THD-U avg"] = seg[["Uthd-1", "Uthd-2", "Uthd-3"]].mean().mean()
        res["THD-I avg"] = seg[["Ithd-1", "Ithd-2", "Ithd-3"]].mean().mean()
        res["THD-U max"] = seg[["Uthd-1", "Uthd-2", "Uthd-3"]].max().max()
        res["THD-I max"] = seg[["Ithd-1", "Ithd-2", "Ithd-3"]].max().max()

        # Crest factors
        res["CfU avg"] = seg[["CfU-1", "CfU-2", "CfU-3"]].mean().mean()
        res["CfI avg"] = seg[["CfI-1", "CfI-2", "CfI-3"]].mean().mean()

        # DC components
        res["Idc avg"] = seg[["Idc-1", "Idc-2", "Idc-3"]].mean().mean()

        # Frequencies
        res["FreqU mean"] = seg["FreqU-1"].mean()
        res["FreqI mean"] = seg["FreqI-1"].mean()

        results[name] = res

    return results

# ---------------- GUI ----------------
class AnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WT5000 ASM Analysis")

        frame = Frame(root)
        frame.pack(fill=BOTH, expand=True)

        self.load_btn = Button(frame, text="Load CSV", command=self.load_file)
        self.load_btn.pack()

        self.text = Text(frame, height=15)
        self.text.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = Scrollbar(frame, command=self.text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.text.config(yscrollcommand=scrollbar.set)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        df = load_csv(file_path)

        # Example segmentation (hardcoded as in your description)
        t_no_start = pd.to_datetime("2025-08-20 14:00:58")
        t_no_end = pd.to_datetime("2025-08-20 14:01:32")
        t_full_start = pd.to_datetime("2025-08-20 14:01:32")
        t_full_end = pd.to_datetime("2025-08-20 14:01:45")

        results = analyze_segments(df, t_no_start, t_no_end, t_full_start, t_full_end)

        # Print results
        self.text.delete(1.0, "end")
        for seg, res in results.items():
            self.text.insert("end", f"--- {seg} ---\n")
            for k, v in res.items():
                self.text.insert("end", f"{k}: {v:.3f}\n")
            self.text.insert("end", "\n")

        # Plotting
        fig, ax = plt.subplots(figsize=(7, 4))
        ax.plot(df["Timestamp"], df["Irms-1"], label="Irms-1")
        ax.plot(df["Timestamp"], df["Irms-2"], label="Irms-2")
        ax.plot(df["Timestamp"], df["Irms-3"], label="Irms-3")
        ax.axvspan(t_no_start, t_no_end, color="cyan", alpha=0.3, label="No-Load")
        ax.axvspan(t_full_start, t_full_end, color="orange", alpha=0.3, label="Full-Load")
        ax.set_ylabel("Current (A)")
        ax.set_xlabel("Time")
        ax.legend()
        ax.set_title("Phase Currents with Load Segments")

        # Embed plot in Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.root)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=BOTH, expand=True)

# ---------------- MAIN ----------------
if __name__ == "__main__":
    root = Tk()
    app = AnalyzerApp(root)
    root.mainloop()
