import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from steps.step1_data import step1_data
# from steps.step2_data import step2_data  # future
# from steps.step3_data import step3_data  # future

def launch_gui():
    root = tk.Tk()
    root.title("Excel Data Transformer")
    root.geometry("500x300")

    selected_file = tk.StringVar()
    selected_step = tk.StringVar()

    def browse_file():
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        selected_file.set(file_path)

    tk.Label(root, text="Selected file:").pack(pady=5)
    tk.Entry(root, textvariable=selected_file, width=60).pack(pady=5)
    tk.Button(root, text="Browse", command=browse_file).pack(pady=5)

    steps = ["Step 1: DATA"]  # add more steps as you implement them
    selected_step.set(steps[0])
    tk.Label(root, text="Select Operation:").pack(pady=5)
    ttk.Combobox(root, values=steps, textvariable=selected_step, state="readonly").pack(pady=5)

    def run():
        file_path = selected_file.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file!")
            return

        step = selected_step.get()
        if step == "Step 1: DATA":
            step1_data(file_path)
        # elif step == "Step 2: ..."  # future steps

    tk.Button(root, text="Run", command=run, bg="green", fg="white").pack(pady=20)

    root.mainloop()
