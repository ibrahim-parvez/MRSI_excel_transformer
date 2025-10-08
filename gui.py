import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import xlwings as xw
import time
import threading
import subprocess

# ---- Import Carbonate Steps ----
from steps.carbon.step1_data import step1_data
from steps.carbon.step2_tosort import step2_tosort
from steps.carbon.step3_last6 import step3_last6
from steps.carbon.step4_group import step4_group
from steps.carbon.step5_summary import step5_summary


def refresh_excel(file_path):
    """Open Excel, recalc, save, and close."""
    if not os.path.exists(file_path):
        print(f"File {file_path} not found for refresh.")
        return
    app = xw.App(visible=False)
    try:
        wb = app.books.open(os.path.abspath(file_path))
        wb.app.calculate()
        time.sleep(1)
        wb.save()
        wb.close()
    finally:
        app.quit()


def open_folder(file_path):
    if not file_path or not os.path.exists(file_path):
        messagebox.showwarning("Warning", "No valid file selected.")
        return
    folder = os.path.dirname(os.path.abspath(file_path))
    if sys.platform.startswith("darwin"):
        subprocess.run(["open", folder])
    elif os.name == "nt":
        os.startfile(folder)
    else:
        subprocess.run(["xdg-open", folder])


def open_file(file_path):
    if not file_path or not os.path.exists(file_path):
        messagebox.showwarning("Warning", "No valid file selected.")
        return
    abs_path = os.path.abspath(file_path)
    if sys.platform.startswith("darwin"):
        subprocess.run(["open", abs_path])
    elif os.name == "nt":
        os.startfile(abs_path)
    else:
        subprocess.run(["xdg-open", abs_path])


def launch_gui():
    root = tk.Tk()
    root.title("McMaster Research Group for Stable Isotopologues - Data Transformation Tool")
    root.geometry("780x580")
    root.configure(bg="#F8F9FA")
    root.minsize(780, 580)

    # ---------------- Modern Style ----------------
    style = ttk.Style()
    style.theme_use("clam")

    style.configure(
        "Modern.TEntry",
        fieldbackground="#f5f5f5",
        foreground="black",
        bordercolor="#cccccc",
        borderwidth=2,
        relief="flat",
        padding=6
    )
    style.map(
        "Modern.TEntry",
        fieldbackground=[("focus", "white"), ("!focus", "#f5f5f5")],
        bordercolor=[("focus", "#4CAF50"), ("!focus", "#cccccc")]
    )

    style.configure(
        "TButton",
        font=("Segoe UI", 11),
        padding=8,
        background="#4CAF50",
        foreground="white",
        borderwidth=0,
        focusthickness=0
    )
    style.map("TButton",
              background=[("active", "#45A049")],
              relief=[("pressed", "flat")])

    style.configure("TLabel", font=("Segoe UI", 11), background="#F8F9FA", foreground="#202124")
    style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background="#F8F9FA")

    # ---------------- Header ----------------
    header_frame = tk.Frame(root, bg="#F8F9FA")
    header_frame.pack(fill="x", pady=10)

    ttk.Label(
        header_frame,
        text="McMaster Research Group for Stable Isotopologues",
        style="Header.TLabel",
        anchor="center",
        justify="center"
    ).pack()
    ttk.Label(
        header_frame,
        text="Data Transformation Tool",
        font=("Segoe UI", 13),
        background="#F8F9FA"
    ).pack(pady=(0, 5))

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
        menu = tk.Menu(root, tearoff=0, bg="white", fg="black", activebackground="#E8E8E8")
        menu.add_command(label="About", command=show_about)
        menu.tk_popup(event.x_root, event.y_root)

    menu_btn = tk.Label(root, text="≡", font=("Segoe UI", 18, "bold"),
                        fg="#333", bg="#F8F9FA", cursor="hand2")
    menu_btn.place(relx=1.0, x=-20, y=10, anchor="ne")
    menu_btn.bind("<Button-1>", show_menu)

    # ---------------- File Picker ----------------
    card_frame = tk.Frame(root, bg="white", bd=1, relief="solid", highlightbackground="#E0E0E0")
    card_frame.pack(padx=20, pady=10, fill="x")

    selected_file = tk.StringVar()
    display_file = tk.StringVar(value="No file selected")

    file_action_frame = tk.Frame(root, bg="#F8F9FA")

    def browse_file():
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            selected_file.set(file_path)
            display_file.set(os.path.basename(file_path))
            file_action_frame.pack(pady=5)
        else:
            selected_file.set("")
            display_file.set("No file selected")
            file_action_frame.pack_forget()

    ttk.Label(card_frame, text="Excel File:").pack(side="left", padx=(10, 10), pady=10)
    file_entry = ttk.Entry(card_frame, textvariable=display_file, state="readonly", width=45, style="Modern.TEntry")
    file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=10)
    browse_btn = ttk.Button(card_frame, text="📂 Browse", command=browse_file)
    browse_btn.pack(side="left", padx=(0, 10), pady=10)

    open_folder_btn = ttk.Button(file_action_frame, text="📁 Open Folder",
                                 command=lambda: open_folder(selected_file.get()))
    open_file_btn = ttk.Button(file_action_frame, text="📄 Open File",
                               command=lambda: open_file(selected_file.get()))
    open_folder_btn.pack(side="left", padx=10)
    open_file_btn.pack(side="left", padx=10)
    file_action_frame.pack_forget()

    # ---------------- Notebook Tabs ----------------
    notebook = ttk.Notebook(root)
    notebook.pack(pady=15, padx=20, fill="both", expand=True)

    # ---- Carbonate Tab ----
    carbon_frame = tk.Frame(notebook, bg="white")
    notebook.add(carbon_frame, text="Carbonate")

    carbon_step_vars = {
        "Step 1: Data": tk.BooleanVar(value=True),
        "Step 2: To Sort": tk.BooleanVar(value=False),
        "Step 3: Last 6": tk.BooleanVar(value=False),
        "Step 4: Group": tk.BooleanVar(value=False),
        "Step 5: Summary": tk.BooleanVar(value=False),
    }

    ttk.Label(carbon_frame, text="Select Steps to Run", background="white").pack(anchor="w", padx=15, pady=(10, 5))

    # Step 1: Data
    step1_outer = tk.Frame(carbon_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    step1_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    step1_inner = tk.Frame(step1_outer, bg="#F5F5F5")
    step1_inner.pack(fill="x", padx=10, pady=8)

    ttk.Checkbutton(step1_inner, text="Step 1: Data", variable=carbon_step_vars["Step 1: Data"]).pack(side="left")
    ttk.Label(step1_inner, text="Initial Sheet:", background="#F5F5F5").pack(side="left", padx=(20, 5))
    sheet_name_var = tk.StringVar(value="Default_Gas_Bench.wke")
    sheet_name_entry = tk.Entry(
        step1_inner, textvariable=sheet_name_var,
        relief="flat", font=("Segoe UI", 10),
        insertbackground="black", highlightthickness=1,
        highlightcolor="#4CAF50", highlightbackground="#CFCFCF",
        bg="white", fg="black", width=25
    )
    sheet_name_entry.pack(side="left", ipady=3, padx=(0, 10))

    # Step 2: To Sort (with dropdown)
    step2_outer = tk.Frame(carbon_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    step2_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    step2_inner = tk.Frame(step2_outer, bg="#F5F5F5")
    step2_inner.pack(fill="x", padx=10, pady=8)

    ttk.Checkbutton(step2_inner, text="Step 2: To Sort", variable=carbon_step_vars["Step 2: To Sort"]).pack(side="left")
    ttk.Label(step2_inner, text="Filter:", background="#F5F5F5").pack(side="left", padx=(20, 5))
    filter_option = tk.StringVar(value="Last 6")
    ttk.Combobox(step2_inner, textvariable=filter_option,
                 values=["All", "Last 6", "Ref Avg", "Start", "End", "Delta"],
                 state="readonly", width=10).pack(side="left")

    # Step 3: Last 6 (boxed for consistency)
    step3_outer = tk.Frame(carbon_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    step3_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    step3_inner = tk.Frame(step3_outer, bg="#F5F5F5")
    step3_inner.pack(fill="x", padx=10, pady=8)
    ttk.Checkbutton(step3_inner, text="Step 3: Last 6", variable=carbon_step_vars["Step 3: Last 6"]).pack(side="left")

    # Step 4: Group (boxed for consistency)
    step3_outer = tk.Frame(carbon_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    step3_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    step3_inner = tk.Frame(step3_outer, bg="#F5F5F5")
    step3_inner.pack(fill="x", padx=10, pady=8)
    ttk.Checkbutton(step3_inner, text="Step 4: Group", variable=carbon_step_vars["Step 4: Group"]).pack(side="left")

    # Step 5: Summary (boxed for consistency)
    step3_outer = tk.Frame(carbon_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    step3_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    step3_inner = tk.Frame(step3_outer, bg="#F5F5F5")
    step3_inner.pack(fill="x", padx=10, pady=8)
    ttk.Checkbutton(step3_inner, text="Step 5: Summary", variable=carbon_step_vars["Step 5: Summary"]).pack(side="left")

    # ---- Water Tab ----
    water_frame = tk.Frame(notebook, bg="white")
    notebook.add(water_frame, text="Water")

    water_step_vars = {
        "Step 1: Data": tk.BooleanVar(value=True),
        "Step 2: Clean": tk.BooleanVar(value=False),
    }

    ttk.Label(water_frame, text="Select Steps to Run", background="white").pack(anchor="w", padx=15, pady=(10, 5))

    # Step 1 box
    water_step1_outer = tk.Frame(water_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    water_step1_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    water_step1_inner = tk.Frame(water_step1_outer, bg="#F5F5F5")
    water_step1_inner.pack(fill="x", padx=10, pady=8)
    ttk.Checkbutton(water_step1_inner, text="Step 1: Data", variable=water_step_vars["Step 1: Data"]).pack(side="left")

    # Step 2 box
    water_step2_outer = tk.Frame(water_frame, bg="#F5F5F5", highlightbackground="#E0E0E0", highlightthickness=1)
    water_step2_outer.pack(anchor="w", fill="x", padx=15, pady=5)
    water_step2_inner = tk.Frame(water_step2_outer, bg="#F5F5F5")
    water_step2_inner.pack(fill="x", padx=10, pady=8)
    ttk.Checkbutton(water_step2_inner, text="Step 2: Clean", variable=water_step_vars["Step 2: Clean"]).pack(side="left")

    # ---------------- Status Output ----------------
    status_frame = ttk.LabelFrame(root, text="Status", padding=(10, 5))
    status_frame.pack_forget()

    status_text = tk.Text(
        status_frame, height=10, wrap="word", state="disabled",
        bg="#202124", fg="#E8EAED", insertbackground="white",
        relief="flat", padx=8, pady=6
    )
    status_text.pack(fill="both", expand=True)

    def log_message(message, color="white"):
        def append():
            status_text.config(state="normal")
            status_text.insert("end", message + "\n", color)
            status_text.tag_configure(color, foreground=color)
            status_text.see("end")
            status_text.config(state="disabled")
        root.after(0, append)

    # ---------------- Background Run ----------------
    def run_steps(file_path, tab):
        if tab == "Carbonate":
            log_message(f"Starting Carbonate processing for: {os.path.basename(file_path)}", "white")

            if carbon_step_vars["Step 1: Data"].get():
                log_message("Running Step 1: DATA...", "white")
                try:
                    sheet_name = sheet_name_var.get().strip()
                    step1_data(file_path, sheet_name)
                    log_message(f"✔ Step 1: DATA completed successfully (Sheet: {sheet_name}).", "green")
                except Exception as e:
                    log_message(f"✖ Step 1: DATA failed: {e}", "red")

            if carbon_step_vars["Step 1: Data"].get() and carbon_step_vars["Step 2: To Sort"].get():
                log_message("Refreshing Excel...", "white")
                try:
                    refresh_excel(file_path)
                    log_message("✔ Excel refreshed successfully.", "green")
                except Exception as e:
                    log_message(f"✖ Excel refresh failed: {e}", "red")

            if carbon_step_vars["Step 2: To Sort"].get():
                log_message(f"Running Step 2: TO SORT (Filter: {filter_option.get()})...", "white")
                try:
                    step2_tosort(file_path, filter_option.get())
                    log_message(f"✔ Step 2: TO SORT ({filter_option.get()}) completed successfully.", "green")
                except Exception as e:
                    log_message(f"✖ Step 2: TO SORT failed: {e}", "red")

            if carbon_step_vars["Step 3: Last 6"].get():
                log_message("Running Step 3: LAST 6...", "white")
                try:
                    step3_last6(file_path)
                    log_message("✔ Step 3: LAST 6 completed successfully.", "green")
                except Exception as e:
                    log_message(f"✖ Step 3: LAST 6 failed: {e}", "red")
            
            if carbon_step_vars["Step 4: Group"].get():
                log_message("Running Step 4: GROUP...", "white")
                try:
                    from steps.carbon.step4_group import step4_group
                    step4_group(file_path)
                    log_message("✔ Step 4: GROUP completed successfully.", "green")
                except Exception as e:
                    log_message(f"✖ Step 4: GROUP failed: {e}", "red")

            if carbon_step_vars["Step 5: Summary"].get():
                log_message("Running Step 5: SUMMARY...", "white")
                try:
                    from steps.carbon.step5_summary import step5_summary
                    step5_summary(file_path)
                    log_message("✔ Step 5: SUMMARY completed successfully.", "green")
                except Exception as e:
                    log_message(f"✖ Step 5: SUMMARY failed: {e}", "red")


        elif tab == "Water":
            log_message(f"Starting Water processing for: {os.path.basename(file_path)}", "white")
            if water_step_vars["Step 1: Data"].get():
                log_message("Running Step 1: DATA...", "white")
                log_message("✔ Step 1: DATA completed successfully.", "green")
            if water_step_vars["Step 2: Clean"].get():
                log_message("Running Step 2: CLEAN...", "white")
                log_message("✔ Step 2: CLEAN completed successfully.", "green")

        log_message("All selected steps finished.\n", "green")

    def run():
        file_path = selected_file.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file!")
            return

        # Smooth window resize when showing status
        def expand_window():
            current_h = root.winfo_height()
            target_h = 760
            if current_h < target_h:
                root.geometry(f"780x{current_h + 10}")
                root.after(10, expand_window)

        if not status_frame.winfo_ismapped():
            status_frame.pack(pady=15, padx=20, fill="both", expand=True)
            expand_window()

        root.update_idletasks()

        current_tab = notebook.tab(notebook.select(), "text")
        thread = threading.Thread(target=run_steps, args=(file_path, current_tab), daemon=True)
        thread.start()

    run_btn = ttk.Button(root, text="▶ Run Selected Steps", command=run)
    run_btn.pack(pady=(20, 10))

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
