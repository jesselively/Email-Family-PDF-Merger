import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import shutil # For copying files for QC Docs
import pikepdf # New PDF library

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch
import tempfile
import threading
import sys # Import sys to help with path for bundled executable

# Register Verdana font if available, though placeholders will use Helvetica.
verdana_path = "C:\\Windows\\Fonts\\verdana.ttf" 
if os.path.exists(verdana_path):
    try:
        pdfmetrics.registerFont(TTFont("Verdana", verdana_path))
    except Exception as e:
        print(f"Could not register Verdana font: {e}")

# Function to get the correct path for bundled resources (like icons)
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def create_placeholder_pdf(control_number_with_suffix):
    """
    Creates a placeholder PDF using ReportLab.
    The "cleaning" step with PyPDF2 is removed; pikepdf will handle merging.
    'control_number_with_suffix' is the full base name, e.g., "CTRL001.1"
    """
    path_rl = None
    try:
        fd_rl, path_rl_temp = tempfile.mkstemp(suffix=f"_rl_placeholder_{control_number_with_suffix}.pdf", prefix="tmp_merger_")
        os.close(fd_rl)
        path_rl = path_rl_temp

        c = canvas.Canvas(path_rl, pagesize=letter)
        width, height = letter
        c.setFont("Helvetica-Bold", 24)
        main_message = "PRODUCED IN NATIVE FORMAT"
        message_y_main = height / 2 + 50
        c.drawCentredString(width / 2.0, message_y_main, main_message)

        c.setFont("Helvetica", 10)
        footer_text = control_number_with_suffix
        text_width_footer = pdfmetrics.stringWidth(footer_text, "Helvetica", 10)
        x_footer = width - text_width_footer - (0.75 * inch)
        y_footer = 0.5 * inch
        c.drawString(x_footer, y_footer, footer_text)
        c.showPage()
        c.save() # ReportLab closes the file here.
        
        # Check if the created PDF has pages (basic validation)
        try:
            with pikepdf.Pdf.open(path_rl) as temp_pdf_check:
                if not temp_pdf_check.pages:
                    raise ValueError(f"ReportLab PDF for {control_number_with_suffix} has no pages after creation.")
        except Exception as e_check: # Catch pikepdf.PdfError or other issues opening
            print(f"ERROR: Placeholder {control_number_with_suffix} created by ReportLab is invalid or unreadable by pikepdf: {e_check}")
            if os.path.exists(path_rl):
                try: os.unlink(path_rl)
                except OSError: pass
            return None

        return path_rl # Return the path to the ReportLab-generated PDF

    except Exception as e:
        print(f"ERROR: Failed to create placeholder for {control_number_with_suffix} with ReportLab: {e}")
        if path_rl and os.path.exists(path_rl): # Clean up if ReportLab failed mid-way
            try: os.unlink(path_rl)
            except OSError: pass
        return None # Indicate failure


def extract_family_key(filename):
    """Extracts the base control number (e.g., CTRL0000002291) from a filename."""
    name = os.path.splitext(filename)[0]
    return name.split(".")[0]

def extract_suffix_parts(filename):
    """
    Extracts suffix parts for sorting (e.g., from CTRL001.0001.0002.pdf -> [1, 2]).
    """
    name_no_ext = os.path.splitext(filename)[0]
    parts = name_no_ext.split(".")
    if len(parts) <= 1: 
        return [] 
    suffix_parts = []
    for part in parts[1:]: 
        try: suffix_parts.append(int(part))
        except ValueError: suffix_parts.append(part) 
    return suffix_parts


def merge_pdfs_worker(folder_path, progress_callback, completion_callback, log_callback_gui):
    """Worker function to handle PDF merging and extended output using pikepdf."""
    
    local_log_entries = []
    def log_message(message):
        log_callback_gui(message) 
        local_log_entries.append([message]) 

    base_output_folder = os.path.join(folder_path, "Merged Output")
    merged_pdfs_output_path = os.path.join(base_output_folder, "Merged PDFs")
    qc_docs_output_path = os.path.join(base_output_folder, "QC Docs")

    try:
        os.makedirs(base_output_folder, exist_ok=True)
        os.makedirs(merged_pdfs_output_path, exist_ok=True)
        os.makedirs(qc_docs_output_path, exist_ok=True)
        log_message(f"Output folders created/ensured: {base_output_folder}")
    except OSError as e:
        log_message(f"CRITICAL ERROR: Could not create output directories: {e}")
        completion_callback(False) 
        return

    all_files_in_folder = []
    try:
        all_files_in_folder = os.listdir(folder_path)
    except OSError as e:
        log_message(f"ERROR: Could not read source directory: {folder_path}. {e}")
        completion_callback(False)
        return

    families = {} 
    file_path_map = {} 
    placeholder_files_to_delete = set() 
    family_had_native_placeholder = {} 

    max_files_in_family_count = 0
    largest_family_merged_pdf_src_path = None 
    first_family_with_native_src_path = None  

    log_message("Starting Pass 1: Identifying families and creating placeholders...")
    for filename_from_dir in all_files_in_folder:
        full_path = os.path.join(folder_path, filename_from_dir)
        
        if os.path.isdir(full_path): 
            continue
        if filename_from_dir.lower().startswith("tmp_merger_"): 
            continue

        base_name_no_ext, ext = os.path.splitext(filename_from_dir)
        family_key = extract_family_key(filename_from_dir)
        
        families.setdefault(family_key, []).append(filename_from_dir)

        if ext.lower() == ".pdf":
            file_path_map[filename_from_dir] = full_path
        else: 
            log_message(f"Creating placeholder for native file: {filename_from_dir}")
            placeholder_path = create_placeholder_pdf(base_name_no_ext)
            if placeholder_path: 
                file_path_map[filename_from_dir] = placeholder_path
                placeholder_files_to_delete.add(placeholder_path) 
                family_had_native_placeholder[family_key] = True 
            else: 
                log_message(f"ERROR: Failed to create placeholder for {filename_from_dir}. It will be skipped.")
                if family_key in families and filename_from_dir in families[family_key]:
                    families[family_key].remove(filename_from_dir) 
    log_message("Pass 1 completed.")


    total_families_to_process = len(families)
    processed_families_count = 0

    log_message("Starting Pass 2: Merging PDF families using pikepdf...")
    for family_key, original_filenames_in_family in families.items():
        valid_files_in_family = [
            fname for fname in original_filenames_in_family if fname in file_path_map
        ]

        if not valid_files_in_family:
            log_message(f"No valid files to merge for family {family_key}. Skipping.")
            processed_families_count += 1
            progress_callback(processed_families_count, total_families_to_process)
            continue
            
        valid_files_in_family.sort(key=lambda fname: (
            len(extract_suffix_parts(fname)), 
            extract_suffix_parts(fname)      
        ))
        
        sorted_filenames_for_merge = valid_files_in_family

        # Use pikepdf for merging
        final_merged_pdf = pikepdf.Pdf.new() # Create a new empty PDF
        files_successfully_added_to_merger = 0
        current_family_contained_native = family_had_native_placeholder.get(family_key, False)
        
        log_message(f"Processing family: {family_key}. Files to merge (in order): {', '.join(sorted_filenames_for_merge)}")

        for filename_to_add in sorted_filenames_for_merge:
            path_to_add = file_path_map.get(filename_to_add) 
            
            if not path_to_add or not os.path.exists(path_to_add):
                log_message(f"WARNING: File path for {filename_to_add} not found or file does not exist. Skipping.")
                continue
            try:
                with pikepdf.Pdf.open(path_to_add) as source_pdf:
                    if not source_pdf.pages:
                        log_message(f"WARNING: File {filename_to_add} (Path: {path_to_add}) is a PDF with no pages (pikepdf). Skipping.")
                        continue
                    final_merged_pdf.pages.extend(source_pdf.pages)
                files_successfully_added_to_merger += 1
            except pikepdf.PdfError as e_pike: # Catch specific pikepdf errors
                 log_message(f"ERROR: pikepdf could not read/process {filename_to_add} (Path: {path_to_add}): {e_pike}. Skipping this file.")
            except Exception as e: 
                log_message(f"ERROR: Could not append {filename_to_add} (Path: {path_to_add}) using pikepdf: {e}. Skipping this file.")

        if files_successfully_added_to_merger > 0:
            output_filename_original = sorted_filenames_for_merge[0]
            base_output_name, _ = os.path.splitext(output_filename_original)
            final_output_filename = f"{base_output_name}.pdf" 
            
            current_merged_pdf_final_path = os.path.join(merged_pdfs_output_path, final_output_filename)
            try:
                final_merged_pdf.save(current_merged_pdf_final_path)
                log_message(f"Successfully merged {files_successfully_added_to_merger} file(s) into: {os.path.join('Merged PDFs', final_output_filename)}")

                if files_successfully_added_to_merger > max_files_in_family_count:
                    max_files_in_family_count = files_successfully_added_to_merger
                    largest_family_merged_pdf_src_path = current_merged_pdf_final_path
                
                if current_family_contained_native and first_family_with_native_src_path is None:
                    first_family_with_native_src_path = current_merged_pdf_final_path

            except Exception as e:
                log_message(f"ERROR: Could not write merged PDF {final_output_filename} using pikepdf: {e}")
        else:
            log_message(f"No files were successfully added to the merger for family {family_key}. No output generated.")

        processed_families_count += 1
        progress_callback(processed_families_count, total_families_to_process)
    log_message("Pass 2 completed.")

    log_message("Processing QC documents...")
    qc_files_copied_for_log = set() 

    if largest_family_merged_pdf_src_path and os.path.exists(largest_family_merged_pdf_src_path):
        try:
            dest_filename = os.path.basename(largest_family_merged_pdf_src_path)
            dest_path = os.path.join(qc_docs_output_path, dest_filename)
            shutil.copy2(largest_family_merged_pdf_src_path, dest_path) 
            log_message(f"Copied largest family PDF to QC Docs: {dest_filename}")
            qc_files_copied_for_log.add(largest_family_merged_pdf_src_path)
        except Exception as e:
            log_message(f"ERROR copying largest family PDF to QC Docs: {e}")

    if first_family_with_native_src_path and os.path.exists(first_family_with_native_src_path):
        if first_family_with_native_src_path != largest_family_merged_pdf_src_path:
            try:
                dest_filename = os.path.basename(first_family_with_native_src_path)
                dest_path = os.path.join(qc_docs_output_path, dest_filename)
                shutil.copy2(first_family_with_native_src_path, dest_path)
                log_message(f"Copied first PDF with native placeholder to QC Docs: {dest_filename}")
                qc_files_copied_for_log.add(first_family_with_native_src_path)
            except Exception as e:
                log_message(f"ERROR copying PDF with native to QC Docs: {e}")
        elif first_family_with_native_src_path not in qc_files_copied_for_log: 
             log_message(f"Largest family PDF ({os.path.basename(largest_family_merged_pdf_src_path)}) also contained a native; already copied for QC.")
    elif not any(family_had_native_placeholder.values()): 
        log_message("No families contained native placeholders; no specific QC doc for this criterion.")

    log_csv_path = os.path.join(base_output_folder, "Merge Log.csv")
    try:
        with open(log_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile)
            csv_writer.writerow(['Message']) 
            csv_writer.writerows(local_log_entries) 
        log_message(f"Merge log saved to: {os.path.join('Merged Output', 'Merge Log.csv')}")
    except IOError as e:
        log_message(f"ERROR: Could not write CSV log file: {e}")

    log_message("Cleaning up temporary files...")
    for temp_path in placeholder_files_to_delete: 
        try:
            if os.path.exists(temp_path): os.unlink(temp_path)
        except OSError as e:
            log_message(f"WARNING: Could not delete temporary placeholder {temp_path}: {e}")
    log_message("Temporary file cleanup completed.")

    completion_callback(True) 


class App:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("Email Family PDF Merger (v1.0.0)") 
        self.root.geometry("600x450") 

        # Set application icon for window and taskbar
        # Assumes 'Icon.ico' is in the same directory as the script,
        # or in the root of the bundled app if using PyInstaller.
        # The resource_path function helps find it when bundled.
        icon_path = resource_path("Icon.ico") 
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except tk.TclError:
                print(f"Warning: Could not set application icon. Ensure '{icon_path}' is a valid .ico file.")
        else:
            print(f"Warning: Icon file not found at '{icon_path}'.")


        self.folder_var = tk.StringVar()
        self.progress_var = tk.IntVar()
        self.merge_thread = None

        tk.Button(self.root, text="About", command=self.show_about).place(x=10, y=10)
        tk.Label(self.root, text="Select folder containing PDFs and native files:").pack(pady=(50, 5)) 
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(fill=tk.X, padx=10)
        tk.Entry(folder_frame, textvariable=self.folder_var, width=70).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(folder_frame, text="Browse...", command=self.browse_folder).pack(side=tk.LEFT, padx=(5,0))

        self.merge_button = tk.Button(self.root, text="Merge PDFs", command=self.start_merge_process, width=15, height=2)
        self.merge_button.pack(pady=(20, 5)) 
        
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, length=550, mode='determinate')
        self.progress_bar.pack(pady=(10, 5))
        self.progress_label = tk.Label(self.root, text="Select a folder and click 'Merge PDFs'")
        self.progress_label.pack()

        tk.Label(self.root, text="Log:").pack(pady=(10,0), anchor=tk.W, padx=10)
        log_frame = tk.Frame(self.root)
        log_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=10, state=tk.DISABLED)
        log_scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.add_log_message_gui("Application started. Ready to merge.")

    def add_log_message_gui(self, message):
        if not self.root.winfo_exists(): return 
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END) 
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks() 

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_var.set(folder_selected)
            self.add_log_message_gui(f"Selected folder: {folder_selected}")

    def update_progress_display(self, processed_count, total_count):
        if not self.root.winfo_exists(): return
        if total_count > 0:
            percent = int((processed_count / total_count) * 100)
            self.progress_var.set(percent)
            self.progress_label.config(text=f"Processing family {processed_count} of {total_count}...")
            self.progress_bar["maximum"] = 100 
        else:
            self.progress_label.config(text="No families found or all processed.") 
            self.progress_var.set(0) 
        self.root.update_idletasks()

    def on_merge_completion(self, success):
        if not self.root.winfo_exists(): return
        if success: 
            self.add_log_message_gui("Merge process finished. Please check 'Merged Output' folder and log for details.")
            messagebox.showinfo("Complete", "PDF merging process finished. Check the 'Merged Output' folder and log for details.")
        else: 
            self.add_log_message_gui("Merge process encountered critical errors or did not complete.")
            messagebox.showerror("Error", "PDF merging process failed or had critical issues. Check the log.")
        
        self.progress_label.config(text="Process finished. Ready for new task.")
        self.merge_button.config(state="normal")

    def start_merge_process(self):
        folder_path = self.folder_var.get()
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Error", "Please select a valid folder.")
            return

        self.log_text.config(state=tk.NORMAL) 
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.add_log_message_gui(f"Starting merge process for folder: {folder_path}")

        self.merge_button.config(state="disabled")
        self.progress_var.set(0)
        self.progress_label.config(text="Initializing merge...")

        self.merge_thread = threading.Thread(
            target=merge_pdfs_worker,
            args=(folder_path, self.update_progress_display, self.on_merge_completion, self.add_log_message_gui),
            daemon=True 
        )
        self.merge_thread.start()

    def show_about(self):
        messagebox.showinfo(
            "About Email Family PDF Merger",
            "Email Family PDF Merger\nVersion 1.0.0\n\n" 
            "This application merges email families (PDFs and native files)\n"
            "into single PDF documents based on control number structure.\n"
            "Native files are converted to placeholder PDFs.\n"
            "Outputs to 'Merged Output' folder with CSV log and QC Docs.\n\n"
            "Developed by Jesse Lively"
        )

if __name__ == "__main__":
    main_root = tk.Tk()
    app_instance = App(main_root)
    main_root.mainloop()
