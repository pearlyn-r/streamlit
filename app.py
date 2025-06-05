# tkinter_model_config.py
import tkinter as tk
from tkinter import ttk, messagebox

# Model limitation items
model_limitation_items = [
    "GRID DENSITY",
    "PEO-CE DIFFERENCE",
    "DETERMINISTIC IR SPREAD",
    "SINGLE NAME TO INDEX BASIS",
    "COMM PARENT CHILD CURVE MAPPING",
    "CDS CURVE RESHAPING",
    "IMPLIED VOLATILITY"
]

def submit():
    cob_date = cob_entry.get()
    entity = entity_var.get()

    if not cob_date:
        messagebox.showerror("Missing Input", "Please enter a COB Date.")
        return

    results = f"COB Date: {cob_date}\nEntity: {entity}\n\nModel Limitation Selections:\n"
    for item in model_limitation_items:
        run_val = limitation_vars[item]['RUN'].get()
        agg_val = limitation_vars[item]['AGGREGATE'].get()
        results += f"{item}:\n  - Run: {run_val}\n  - Aggregate: {agg_val}\n"

    messagebox.showinfo("Form Submitted", results)

# Root window
root = tk.Tk()
root.title("Model Configuration Interface")
root.geometry("800x700")
root.resizable(True, True)

# --- COB DATE, ENTITY ---
title = ttk.Label(root, text="Model Configuration Interface", font=("Arial", 16, "bold"))
title.pack(pady=10)

frame_top = ttk.Frame(root)
frame_top.pack(pady=10, padx=10, fill='x')

# COB Date
ttk.Label(frame_top, text="COB DATE").grid(row=0, column=0, padx=10, sticky='w')
cob_entry = ttk.Entry(frame_top, width=20)
cob_entry.grid(row=1, column=0, padx=10, sticky='w')

# Entity
ttk.Label(frame_top, text="ENTITY").grid(row=0, column=1, padx=10, sticky='w')
entity_options = ["A", "B", "C", "D", "E"]
entity_var = tk.StringVar(value=entity_options[0])
entity_menu = ttk.Combobox(frame_top, textvariable=entity_var, values=entity_options, state="readonly")
entity_menu.grid(row=1, column=1, padx=10, sticky='w')

# Submit button
submit_btn = ttk.Button(frame_top, text="SUBMIT", command=submit)
submit_btn.grid(row=1, column=2, padx=20, sticky='w')

# --- Model Limitations ---
ttktitle = ttk.Label(root, text="MODEL LIMITATIONS", font=("Arial", 12, "bold"))
ttktitle.pack(pady=5)

header_frame = ttk.Frame(root)
header_frame.pack(padx=10, fill='x')
ttk.Label(header_frame, text="Limitation", width=40).grid(row=0, column=0)
ttk.Label(header_frame, text="RUN", width=10).grid(row=0, column=1)
ttk.Label(header_frame, text="AGGREGATE", width=10).grid(row=0, column=2)

# Store all select variables
limitation_vars = {}

# Each item row
for idx, item in enumerate(model_limitation_items):
    row_frame = ttk.Frame(root)
    row_frame.pack(fill='x', padx=10, pady=2)
    ttk.Label(row_frame, text=item, width=40).grid(row=0, column=0, sticky='w')

    run_var = tk.StringVar(value="No")
    run_box = ttk.Combobox(row_frame, textvariable=run_var, values=["Yes", "No"], state="readonly", width=8)
    run_box.grid(row=0, column=1, padx=5)

    agg_var = tk.StringVar(value="No")
    agg_box = ttk.Combobox(row_frame, textvariable=agg_var, values=["Yes", "No"], state="readonly", width=8)
    agg_box.grid(row=0, column=2, padx=5)

    limitation_vars[item] = {"RUN": run_var, "AGGREGATE": agg_var}

# Run main loop
root.mainloop()
