import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
from steps.step1_data import step1_data
# from steps.step2_data import step2_data
# from steps.step3_data import step3_data


def launch_gui():
    root = tk.Tk()
    root.title("McMaster Research Group for Stable Isotopologues - Data Transformation Tool")
    root.geometry("750x500")
    root.configure(bg="white")  # clean white background

    # ---------------- Style ----------------
    style = ttk.Style()
    style.theme_use("clam")  # cleaner ttk look

    style.configure("TButton", padding=6, relief="flat", font=("Segoe UI", 11))
    style.map("TButton",
              background=[("!active", "#4CAF50"), ("active", "#45a049")],
              foreground=[("!disabled", "white")])

    style.configure("TLabel", font=("Segoe UI", 11), background="white")
    style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background="white")

    # ---------------- Header ----------------
    header_frame = tk.Frame(root, bg="white")
    header_frame.pack(fill="x", pady=10, padx=15)

    # Title (centered)
    header = ttk.Label(
        header_frame,
        text="McMaster Research Group for Stable Isotopologues\nData Transformation Tool",
        style="Header.TLabel",
        anchor="center",
        justify="center"
    )
    header.pack(side="top", expand=True)

    # ---------------- Hamburger Menu ----------------
    def show_about():
        python_version = sys.version.split()[0]
        messagebox.showinfo(
            "About",
            f"""McMaster Research Group for Stable Isotopologues
Data Transformation Tool

Required Python Version: {python_version}

© 2025 McMaster University
Developer: Ibrahim Parvez"""
        )

    def show_menu(event):
        menu = tk.Menu(root, tearoff=0, bg="white", fg="black")
        menu.add_command(label="About", command=show_about)
        # future options:
        # menu.add_command(label="Settings", command=...)
        # menu.add_command(label="Help", command=...)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # Floating black hamburger (top-right)
    menu_btn = tk.Label(root, text="≡", font=("Segoe UI", 18, "bold"),
                        fg="black", bg="white", cursor="hand2")
    menu_btn.place(relx=1.0, x=-20, y=10, anchor="ne")  # top-right corner
    menu_btn.bind("<Button-1>", show_menu)

    # ---------------- File Picker ----------------
    file_frame = ttk.Frame(root)
    file_frame.pack(pady=10, padx=20, fill="x")

    selected_file = tk.StringVar()
    display_file = tk.StringVar(value="No file selected")

    def browse_file():
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            selected_file.set(file_path)
            display_file.set(os.path.basename(file_path))

    ttk.Label(file_frame, text="File:").pack(side="left", padx=(5, 10))
    file_entry = ttk.Entry(file_frame, textvariable=display_file, state="readonly", width=50)
    file_entry.pack(side="left", padx=(0, 10), fill="x", expand=True)

    browse_btn = ttk.Button(file_frame, text="📂", width=3, command=browse_file)
    browse_btn.pack(side="left")

    # ---------------- Step Selection ----------------
    steps_frame = ttk.LabelFrame(root, text="Select Steps to Run")
    steps_frame.pack(pady=15, padx=20, fill="x")

    step_vars = {
        "Step 1: DATA": tk.BooleanVar(value=True),
        "Step 2: Placeholder": tk.BooleanVar(value=False),
        "Step 3: Placeholder": tk.BooleanVar(value=False),
    }

    for step, var in step_vars.items():
        ttk.Checkbutton(steps_frame, text=step, variable=var).pack(anchor="w", padx=15, pady=3)

    # ---------------- Run Button ----------------
    def run():
        file_path = selected_file.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file!")
            return

        status_frame.pack(pady=15, padx=20, fill="both", expand=True)  # show status box only after run clicked
        log_message(f"Starting processing for: {os.path.basename(file_path)}", "white")

        if step_vars["Step 1: DATA"].get():
            log_message("Running Step 1: DATA...", "white")
            try:
                step1_data(file_path)
                log_message("✔ Step 1: DATA completed successfully.", "green")
            except Exception as e:
                log_message(f"✖ Step 1: DATA failed: {e}", "red")

        if step_vars["Step 2: Placeholder"].get():
            log_message("Running Step 2 (not yet implemented)...", "white")

        if step_vars["Step 3: Placeholder"].get():
            log_message("Running Step 3 (not yet implemented)...", "white")

        log_message("All selected steps finished.\n", "green")

    run_btn = ttk.Button(root, text="▶ Run Selected Steps", command=run)
    run_btn.pack(pady=(20, 10))

    # ---------------- Status Output ----------------
    status_frame = ttk.LabelFrame(root, text="Status")
    status_frame.pack_forget()  # hidden until Run clicked

    status_text = tk.Text(
        status_frame, height=10, wrap="word",
        state="disabled", bg="black", fg="white", insertbackground="white"
    )
    status_text.pack(fill="both", expand=True, padx=5, pady=5)

    def log_message(message, color="white"):
        status_text.config(state="normal")
        status_text.insert("end", message + "\n", color)
        status_text.tag_configure(color, foreground=color)
        status_text.see("end")
        status_text.config(state="disabled")

    root.mainloop()
