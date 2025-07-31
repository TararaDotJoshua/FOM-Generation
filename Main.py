import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from scipy.optimize import curve_fit

def select_file():
    """Open a file picker to choose an Excel file."""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select FOM Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        exit(1)
    return file_path

def load_fom_data(filepath):
    """Load and validate the FOM data from the Excel file."""
    try:
        df = pd.read_excel(filepath, sheet_name="FOM_Data")
        df.dropna(subset=["X FOM", "Y FOM"], inplace=True)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")
        exit(1)

def get_chart_metadata(df):
    """Extract chart title and axis names from the sheet."""
    chart_title = df["Chart Title"].dropna().values[0] if "Chart Title" in df.columns and not df["Chart Title"].dropna().empty else "Figure of Merit (FOM) Scatter Plot"
    x_name = df["X FOM Name"].dropna().values[0] if "X FOM Name" in df.columns and not df["X FOM Name"].dropna().empty else "X FOM"
    y_name = df["Y FOM Name"].dropna().values[0] if "Y FOM Name" in df.columns and not df["Y FOM Name"].dropna().empty else "Y FOM"
    return chart_title, x_name, y_name

def parse_sota_override(df):
    """Parse the SOTA override formula if available."""
    formula = df["SOTA Override"].dropna()
    if not formula.empty:
        match = re.search(r"y\s*=\s*([\d\.]+)\s*\*\s*x\s*\+?\s*([\d\.-]*)", formula.values[0])
        if match:
            slope = float(match.group(1))
            intercept = float(match.group(2)) if match.group(2) else 0
            return lambda x: slope * x + intercept
    return None

def compute_sota_auto(df):
    """Fit SOTA line using only competitor data (ID starts with 'COMP')."""
    df_comp = df[df['ID'].str.startswith("COMP")]

    if df_comp.empty:
        print("⚠️ No competitor data. Using flat zero SOTA line.")
        return lambda x: np.full_like(x, fill_value=0)

    if len(df_comp) == 1:
        y_val = df_comp["Y FOM"].values[0]
        print(f"Only one competitor data point. Using horizontal SOTA line at y = {y_val}")
        return lambda x: np.full_like(x, fill_value=y_val)

    # Enough data to fit a line
    def linear(x, a, b):
        return a * x + b

    x = df_comp["X FOM"].values
    y = df_comp["Y FOM"].values
    params, _ = curve_fit(linear, x, y)
    return lambda x: params[0] * x + params[1]

def plot_fom(df, sota_fn):
    """Plot the FOM scatterplot with SOTA line and labels (log X-axis)."""
    chart_title, x_label, y_label = get_chart_metadata(df)
    plt.figure(figsize=(10, 7))

    # Separate by type
    products = df[df["ID"].str.startswith("PROD")]
    competitors = df[df["ID"].str.startswith("COMP")]

    # Plot competitors
    plt.scatter(competitors["X FOM"], competitors["Y FOM"], color='red', label="Competitor")
    for _, row in competitors.iterrows():
        plt.text(row["X FOM"], row["Y FOM"], str(row["Label"]), fontsize=8, ha='right', va='bottom')

    # Plot products
    plt.scatter(products["X FOM"], products["Y FOM"], color='blue', label="Product")
    for _, row in products.iterrows():
        plt.text(row["X FOM"], row["Y FOM"], str(row["Label"]), fontsize=8, ha='right', va='bottom')

    # SOTA Line
    x_vals = np.logspace(np.log10(df["X FOM"].min()), np.log10(df["X FOM"].max()), 100)
    plt.plot(x_vals, sota_fn(x_vals), '--', color='green', label='SOTA Line')

    plt.xscale("log")
    plt.xlabel(f"{x_label} (log scale)")
    plt.ylabel(y_label)
    plt.title(chart_title)
    plt.grid(True, which='both', linestyle='--', linewidth=0.5)
    plt.legend()
    plt.tight_layout()
    plt.show()

# --- Main Program ---
if __name__ == "__main__":
    file_path = select_file()
    df = load_fom_data(file_path)
    sota_fn = parse_sota_override(df) or compute_sota_auto(df)
    plot_fom(df, sota_fn)
