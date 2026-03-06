import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from PIL import Image, ImageTk
import os
import logging
import sys
# adjust path_to_project to the folder that contains ICOFR (absolute or computed)
# In your case ICOFR is at: .../Desktop/FS Utility/ICOFR
PROJECT_PARENT = Path(__file__).resolve().parents[1]  # moves two levels up to "Desktop/FS Utility"
sys.path.insert(0, str(PROJECT_PARENT))
# from Concurrent_Report import Concurrent_Report_Main
# from Concurrent_Dashboard import Concurrent_Dashboard_Main
# from KYC import KYC_main
# from FD_Focused_Sampling import FD_Main
# from CCIL import CCIL_main
from Concurrent_Audit_Report.Concurrent_Report import Concurrent_Report_Main
from Concurrent_Audit_Dashboard.Concurrent_Dashboard import Concurrent_Dashboard_Main
from IA_Report.IA_Report_Utility_v2 import IA_Report_Utility_Main
from KYC.KYC import KYC_main
from FD_Focused_Sampling.FD_Focused_Sampling import FD_Main
from CCIL.CCIL import CCIL_main
from FOBO.FOBO import main as fo_bo_main
from ICOFR import main as ICOFR_main
from IOA_Sampling.KPI1.KPI1_script import main as main1
from IOA_Sampling.KPI2.KPI2_script import main as main2
from IOA_Sampling.KPI3.KPI3_script import main as main3
from IOA_Sampling.KPI4.KPI4_script import main as main4
from Swift.Reconv3 import recon_main #recon_main(tracker_path, recon_path, output_folder_path)
from IA_Report_Template2.inspect_placeholders import IA_Report_Utility_MainT2

def main(excel_path, pptx_path, audit_type, utility_type, report_type, output_path, month, year):
    if audit_type == "FD - Sampling":
        output_path = FD_Main(excel_path, output_path)
    elif audit_type == "KYC":
        output_path = KYC_main(excel_path, output_path)
    elif audit_type == "Internal Audit" and utility_type == "Report":
        output_path = IA_Report_Utility_Main(excel_path, pptx_path, output_path)
    elif audit_type == "Concurrent Audit" and utility_type == "Report":
        output_path = Concurrent_Report_Main(excel_path, pptx_path, report_type, output_path, month, year)
    elif audit_type == "Concurrent Audit" and utility_type == "Dashboard":
        output_path = Concurrent_Dashboard_Main(excel_path, pptx_path, output_path)
    elif audit_type == "CCIL":
        output_path = CCIL_main(excel_path, output_path)
    elif audit_type == "ICOFR" and utility_type == "Dashboard":
        output_path = ICOFR_main(excel_path, pptx_path, output_path)
    elif audit_type == "FOBO":
        output_path = fo_bo_main(excel_path, output_path)
    elif audit_type == "IOA KPI1":
        output_path = main1(excel_path, output_path)
    elif audit_type == "IOA KPI2":
        output_path = main2(excel_path, output_path)
    elif audit_type == "IOA KPI3":
        output_path = main3(excel_path, output_path)
    elif audit_type == "IOA KPI4":
        output_path = main4(excel_path, output_path)
    elif audit_type == "Swift":
        output_path = recon_main(excel_path, output_path)
    elif audit_type == "Internal Audit-Zensar" and utility_type == "Report":
        output_path = IA_Report_Utility_MainT2(excel_path,pptx_path,output_path)
    else:
        output_path = "Output not generated. Try again"
    return output_path

# main(
#     r"C:\Users\aryansharma8\OneDrive - KPMG\Desktop\FS Utility Input\Concurrent Audit\Automation - Issue Tracker.xlsx",
#     r"C:\Users\aryansharma8\OneDrive - KPMG\Desktop\FS Utility Input\Concurrent Report\Report Template- ABC Bank Report.pptx",
#     "Concurrent Report",
#     "Report",
#     "Both",
#     r"C:\Users\aryansharma8\OneDrive - KPMG\Desktop\FS Utility Output",
#     "March",
#     "2025"
# )

# main(
#     r"C:\Users\satyambarnwal\OneDrive - KPMG\Sharma, Aryan's files - FS Automation\Report_Utility\IA_Report_Template2\Input\Input file for Report Automation.xlsx",
#     r"C:\Users\satyambarnwal\OneDrive - KPMG\Sharma, Aryan's files - FS Automation\Report_Utility\IA_Report_Template2\Input\Zensar_IA_Subcontracting Review_Q3 FY25-26 (Automation) - Template.pptx",
#     "Internal Audit-Zensar",
#     "Report",
#     "Both",
#     r"C:\Users\satyambarnwal\OneDrive - KPMG\Sharma, Aryan's files - FS Automation\Report_Utility\IA_Report_Template2\Output",
#     "March",
#     "2025"
# )
# --- Functions ---
def browse_excel_csv():
    file_path = filedialog.askopenfilename(
        title="Select Excel or CSV file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
    )
    if file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)
        excel_entry.config(fg='black')
        pptx_entry.config(state="normal")
        pptx_button.config(state="normal")

def browse_pptx():
    file_path = filedialog.askopenfilename(
        title="Select PowerPoint file",
        filetypes=[("PowerPoint files", "*.pptx")]
    )
    if file_path:
        pptx_entry.delete(0, tk.END)
        pptx_entry.insert(0, file_path)
        pptx_entry.config(fg='black')
        output_entry.config(state="normal")
        output_button.config(state="normal")

def browse_output_folder():
    folder_path = filedialog.askdirectory(title="Select Output Folder")
    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)
        output_entry.config(fg='black')

def submit_paths():
    excel_path = excel_entry.get().strip()
    pptx_path = pptx_entry.get().strip()
    output_folder = output_entry.get().strip()
    audit_type = audit_var.get()
    utility_type = utility_var.get()
    report_type = report_var.get()
    # Extract month and year
    selected_month = month_var.get()
    selected_year = year_var.get()

    # FOBO specific validation and call (handled first)
    if audit_type == "FOBO":
        # values from GUI
        excel_path = excel_entry.get().strip()
        output_folder = output_entry.get().strip()

        if not excel_path or excel_path.startswith("Select"):
            messagebox.showwarning("Missing Input", "Please select the FOBO input Excel file.")
            return
        if not output_folder or output_folder.startswith("Select"):
            messagebox.showwarning("Missing Input", "Please select the Output folder.")
            return

        try:
            # fo_bo_main(input_file: str = None, output_file: str = None) -> returns output path (str)
            out_path = fo_bo_main(excel_path, None)

            if out_path is None:
                messagebox.showinfo("Output", f"FOBO reconciliation saved next to input file or in default output location.")
            else:
                messagebox.showinfo("Output", f"FOBO reconciliation saved to:\n{out_path}")

            root.destroy()
            return

        except Exception as e:
            logging.exception("FOBO processing failed")
            messagebox.showerror("Error", f"FOBO processing failed:\n{e}")
            return
    
    # ICOFR specific validation and call
    if audit_type == "ICOFR":
        if not excel_path or excel_path.startswith("Select"):
            messagebox.showwarning("Missing Input", "Please select the Issue Tracker Excel file for ICOFR.")
            return
        if not pptx_path or pptx_path.startswith("Select"):
            messagebox.showwarning("Missing Input", "Please select the Template Presentation for ICOFR.")
            return
        if not output_folder or output_folder.startswith("Select"):
            messagebox.showwarning("Missing Input", "Please select the Output folder for ICOFR.")
            return

        try:
            # call ICOFR_main with values from GUI
            result = ICOFR_main(excel_path, pptx_path, output_folder)

            # handle return value defensively:
            if result is None:
                # ICOFR_main saved files in output_folder / "ICOFR Report"
                messagebox.showinfo("Output", f"ICOFR report saved under:\n{os.path.join(output_folder, 'ICOFR Report')}")
            else:
                # result may be filename or absolute path
                final_path = result if os.path.isabs(result) else os.path.join(output_folder, result)
                messagebox.showinfo("Output", f"Output generated at:\n{final_path}")

            root.destroy()
            return

        except Exception as e:
            logging.exception("ICOFR_main failed")
            messagebox.showerror("Error", f"ICOFR processing failed:\n{e}")
            return

    # Validation based on audit type
    if audit_type == "FD - Sampling" or audit_type == "KYC" or audit_type == "CCIL" or audit_type == "IOA KPI1" or audit_type == "IOA KPI2" or audit_type == "IOA KPI3" or audit_type == "IOA KPI4" or audit_type == "Swift":
        if not excel_path or not output_folder:
            messagebox.showwarning("Missing Input", "Please provide the Excel Issue Tracker file path and Output folder.")
            return

    elif audit_type == "Concurrent Audit" and utility_type == "Report":
        if not excel_path or not pptx_path or not output_folder:
            messagebox.showwarning("Missing Input", "Please provide all required paths.")
            return
        if selected_month == "Month" or selected_year == "Year":
            messagebox.showwarning("Missing Input", "Please select both Report Issuance Month and Year.")
            return

    else:
        if not excel_path or not pptx_path or not output_folder:
            messagebox.showwarning("Missing Input", "Please provide all required paths.")
            return

    try:
        # You can pass formatted_issuance_month to main() if needed
        output_filename = main(excel_path, pptx_path, audit_type, utility_type, report_type, output_folder, selected_month, selected_year)
        final_output_path = os.path.join(output_folder, output_filename)
        messagebox.showinfo("Output", f"Output generated at:\n{final_output_path}")
        root.destroy()

    except Exception as e:
        messagebox.showerror("Done")

#, f"An error occurred:\n{e}

def add_placeholder(entry, placeholder):
    entry.insert(0, placeholder)
    entry.config(fg='grey')

    def on_focus_in(event):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.config(fg='black')

    def on_focus_out(event):
        if entry.get() == '':
            entry.insert(0, placeholder)
            entry.config(fg='grey')

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

def toggle_report_type(*args):
    audit_type = audit_var.get()

    # Reset utility type and menu
    utility_var.set("Select Utility Type")
    utility_menu["menu"].delete(0, "end")

    # Hide all fields initially
    utility_menu.grid_remove()
    report_type_frame.grid_remove()
    pptx_frame.grid_remove()
    pptx_entry.config(state="disabled")
    pptx_button.config(state="disabled")

    excel_entry.config(state="disabled")
    excel_button.config(state="disabled")
    output_entry.config(state="disabled")
    output_button.config(state="disabled")

    if audit_type == "FD - Sampling" or audit_type == "KYC" or audit_type == "CCIL" or audit_type == "IOA KPI1" or audit_type == "IOA KPI2" or audit_type == "IOA KPI3" or audit_type == "IOA KPI4" or audit_type == "Swift":
        # Show only Excel and Output fields
        excel_frame.grid(row=3, column=1, sticky="e", padx=input_padx, pady=3)
        excel_entry.config(state="normal")
        excel_button.config(state="normal")

        output_frame.grid(row=5, column=1, sticky="e", padx=input_padx, pady=3)
        output_entry.config(state="normal")
        output_button.config(state="normal")
        return

    if audit_type == "FOBO":
        # FOBO needs only the input Excel and Output folder
        excel_frame.grid(row=3, column=1, sticky="e", padx=input_padx, pady=3)
        excel_entry.config(state="normal")
        excel_button.config(state="normal")

        output_frame.grid(row=5, column=1, sticky="e", padx=input_padx, pady=3)
        output_entry.config(state="normal")
        output_button.config(state="normal")
        return

    if audit_type == "ICOFR":
        excel_frame.grid(row=3, column=1, sticky="e", padx=input_padx, pady=3)
        excel_entry.config(state="normal")
        excel_button.config(state="normal")

        pptx_frame.grid(row=4, column=1, sticky="e", padx=input_padx, pady=3)
        pptx_entry.config(state="normal")
        pptx_button.config(state="normal")

        output_frame.grid(row=5, column=1, sticky="e", padx=input_padx, pady=3)
        output_entry.config(state="normal")
        output_button.config(state="normal")
        return

    # Show all fields for other audit types
    utility_menu.grid(row=1, column=1, sticky="e", padx=input_padx, pady=3)
    pptx_frame.grid(row=4, column=1, sticky="e", padx=input_padx, pady=3)

    excel_entry.config(state="disabled")
    excel_button.config(state="disabled")
    pptx_entry.config(state="disabled")
    pptx_button.config(state="disabled")
    output_entry.config(state="disabled")
    output_button.config(state="disabled")

    utility_menu.config(state="normal")
    if audit_type == "Internal Audit" or audit_type == "Internal Audit-Zensar":
        utility_menu["menu"].add_command(label="Report", command=lambda: utility_var.set("Report"))
    elif audit_type == "Concurrent Audit":
        for option in ["Dashboard", "Report"]:
            utility_menu["menu"].add_command(label=option, command=lambda opt=option: utility_var.set(opt))

# def enable_excel_entry(*args):
#     utility_type = utility_var.get()
#     audit_type = audit_var.get()
#     if utility_type != "Select Utility Type":
#         excel_entry.config(state="normal")
#         excel_button.config(state="normal")
#         pptx_entry.config(state="disabled")
#         pptx_button.config(state="disabled")
#         output_entry.config(state="disabled")
#         output_button.config(state="disabled")

#         if audit_type == "Concurrent Report" and utility_type == "Report":
#             report_type_frame.grid(row=2, column=1, sticky="e", padx=input_padx, pady=3)
#         else:
#             report_type_frame.grid_remove()

def enable_excel_entry(*args):
    utility_type = utility_var.get()
    audit_type = audit_var.get()

    if utility_type != "Select Utility Type":
        excel_entry.config(state="normal")
        excel_button.config(state="normal")
        pptx_entry.config(state="disabled")
        pptx_button.config(state="disabled")
        output_entry.config(state="disabled")
        output_button.config(state="disabled")


        if audit_type == "Concurrent Audit" and utility_type == "Report":
            report_type_frame.grid(row=2, column=1, sticky="e", padx=input_padx, pady=3)
            issuance_month_frame.grid(row=6, column=1, sticky="e", padx=input_padx, pady=3)
        else:
            report_type_frame.grid_remove()
            issuance_month_frame.grid_remove()


# --- GUI Setup ---
root = tk.Tk()
root.title("KPMG Report And Dashboard Automation Utility")
root.geometry("770x550")
root.resizable(False, False)

# Fonts and Padding
label_font = ("Arial", 11)
label_padx = (10, 5)
input_padx = (5, 10)

# Logo
try:
    logo_img = Image.open(r"Input_Data\KPMG_logo.png")
    logo_img = logo_img.resize((300, 120))
    logo_tk = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(root, image=logo_tk)
    logo_label.pack(pady=(10, 0))
except:
    logo_label = tk.Label(root, text="KPMG Logo", font=("Arial", 12))
    logo_label.pack(pady=(10, 0))

# Title
title_label = tk.Label(root, text="KPMG Report And Dashboard Automation Utility", font=("Arial", 16, "bold"))
title_label.pack(pady=10)

# Content Frame
content_frame = tk.Frame(root)
content_frame.pack(pady=10, fill="x")
content_frame.columnconfigure(1, weight=1)

# Audit Type
audit_var = tk.StringVar(value="Audit Type")
audit_var.trace_add("write", toggle_report_type)
tk.Label(content_frame, text="Audit Type:", anchor='w').grid(row=0, column=0, sticky="w", padx=label_padx, pady=3)
audit_menu = tk.OptionMenu(content_frame, audit_var, "Concurrent Audit", "Internal Audit","Internal Audit-Zensar", "FD - Sampling", "KYC", "CCIL" , "ICOFR" , "FOBO" , "IOA KPI1", "IOA KPI2", "IOA KPI3", "IOA KPI4" , "Swift")
audit_menu.config(width=25, bg="white")
audit_menu.grid(row=0, column=1, sticky="e", padx=input_padx, pady=3)

# Utility Type
utility_var = tk.StringVar(value="Select Utility Type")
utility_var.trace_add("write", enable_excel_entry)
tk.Label(content_frame, text="Select Utility Type:", anchor='w').grid(row=1, column=0, sticky="w", padx=label_padx, pady=3)
utility_menu = tk.OptionMenu(content_frame, utility_var, "")
utility_menu.config(width=25, bg="white", state="disabled")
utility_menu.grid(row=1, column=1, sticky="e", padx=input_padx, pady=3)

# Report Type (hidden by default)
report_var = tk.StringVar(value="Draft")
report_type_frame = tk.Frame(content_frame)
tk.Label(report_type_frame, text="Report Type:", anchor='w').pack(side="left")
tk.Radiobutton(report_type_frame, text="Both", variable=report_var, value="Both").pack(side="left")
tk.Radiobutton(report_type_frame, text="Draft", variable=report_var, value="Draft").pack(side="left")
tk.Radiobutton(report_type_frame, text="Final", variable=report_var, value="Final").pack(side="left")
report_type_frame.grid_remove()

# Report Issuance Month (hidden by default)
issuance_month_frame = tk.Frame(content_frame)

tk.Label(issuance_month_frame, text="Report Issuance Month:", anchor='w').pack(side="left")

# Month Dropdown
month_var = tk.StringVar(value="Month")
month_options = ["January", "February", "March", "April", "May", "June",
                 "July", "August", "September", "October", "November", "December"]
month_menu = tk.OptionMenu(issuance_month_frame, month_var, *month_options)
month_menu.config(width=10, bg="white")
month_menu.pack(side="left", padx=5)

# Year Dropdown
year_var = tk.StringVar(value="Year")
year_options = [str(year) for year in range(2020, 2030)]  # You can adjust the range
year_menu = tk.OptionMenu(issuance_month_frame, year_var, *year_options)
year_menu.config(width=6, bg="white")
year_menu.pack(side="left", padx=5)

issuance_month_frame.grid_remove()
# # KPIs selector (hidden by default)  <-- REPLACE existing kpi_frame / kpi_menu block with this
# kpi_var = tk.StringVar(value="select")   # default used by checks
# kpi_frame = tk.Frame(content_frame)
# tk.Label(kpi_frame, text="KPI:", anchor="w").pack(side="left")
# # Four explicit KPI options + All — labels the user will see
# kpi_menu = tk.OptionMenu(kpi_frame, kpi_var, "KPI1", "KPI2", "KPI3", "KPI4", "All")
# kpi_menu.config(width=12, bg="white")
# kpi_frame.grid(row=2, column=1, sticky="e", padx=input_padx, pady=3)
# kpi_frame.grid_remove()

# Excel File
tk.Label(content_frame, text="Select the Issue Tracker Excel File Path:", anchor='w').grid(row=3, column=0, sticky="w", padx=label_padx, pady=3)
excel_frame = tk.Frame(content_frame)
excel_entry = tk.Entry(excel_frame, width=40, bg="white", state="disabled")
excel_entry.pack(side="left", padx=5)
add_placeholder(excel_entry, "Select the Issue Tracker Excel File Path")
excel_button = tk.Button(excel_frame, text="Select", command=browse_excel_csv, state="disabled")
excel_button.pack(side="left")
excel_frame.grid(row=3, column=1, sticky="e", padx=input_padx, pady=3)

# PPTX File
tk.Label(content_frame, text="Select the Template Presentation:", anchor='w').grid(row=4, column=0, sticky="w", padx=label_padx, pady=3)
pptx_frame = tk.Frame(content_frame)
pptx_entry = tk.Entry(pptx_frame, width=40, bg="white", state="disabled")
pptx_entry.pack(side="left", padx=5)
add_placeholder(pptx_entry, "Select the Template Presentation")
pptx_button = tk.Button(pptx_frame, text="Select", command=browse_pptx, state="disabled")
pptx_button.pack(side="left")
pptx_frame.grid(row=4, column=1, sticky="e", padx=input_padx, pady=3)

# Output Folder
tk.Label(content_frame, text="Select Output Folder:", anchor='w').grid(row=5, column=0, sticky="w", padx=label_padx, pady=3)
output_frame = tk.Frame(content_frame)
output_entry = tk.Entry(output_frame, width=40, bg="white", state="disabled")
output_entry.pack(side="left", padx=5)
add_placeholder(output_entry, "Select Output Folder")
output_button = tk.Button(output_frame, text="Select", command=browse_output_folder, state="disabled")
output_button.pack(side="left")
output_frame.grid(row=5, column=1, sticky="e", padx=input_padx, pady=3)

# Submit Button
submit_btn = tk.Button(root, text="Run", command=submit_paths, bg="#00338D", fg="white", width=15)
submit_btn.pack(pady=20)

root.mainloop()
#