import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog # Added simpledialog, filedialog
import tkinter.font as tkFont
import requests
import threading
import queue
import time
import pandas as pd # Import pandas

API_BASE_URL = "https://universitaly-backend.cineca.it/api/offerta-formativa/cerca-corsi"
TOTAL_PAGES = 575 # Set this to a smaller number (e.g., 10-20) for testing
FETCH_DELAY = 0.0
CONTACT_INFO = "For studying in Italy, contact Saurav Shriwastav: +39 350 971 9486"
EXPORT_PASSWORD = "Nepal123" # The required password

class CourseFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Italian University Course Explorer")
        self.root.geometry("1050x680") # Slightly taller for export button

        self.all_courses = []
        self.is_loading = False
        self.fetch_queue = queue.Queue()
        self.last_filter_time = 0
        self._after_id_filter = None

        self.footer_font = tkFont.Font(family="Helvetica", size=9, slant="italic")

        self.search_var = tk.StringVar()
        self.english_only_var = tk.BooleanVar()
        self.degree_type_var = tk.StringVar()
        self.degree_type_var.set("All Degree Types")

        self.create_filter_widgets() # Add export button here
        self.create_results_table()
        self.create_status_bar()
        self.create_footer()

        self.search_var.trace_add("write", self.schedule_filter_update)
        self.english_only_var.trace_add("write", self.schedule_filter_update)
        self.degree_type_combobox.bind("<<ComboboxSelected>>", self.schedule_filter_update)

        self.start_fetch_thread()
        self.check_fetch_queue()

    def create_filter_widgets(self):
        filter_frame = ttk.LabelFrame(self.root, text="Filters & Actions", padding="10")
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        # --- Filters ---
        ttk.Label(filter_frame, text="Search Course/Uni:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        english_checkbox = ttk.Checkbutton(filter_frame, text="English Only", variable=self.english_only_var, onvalue=True, offvalue=False)
        english_checkbox.grid(row=0, column=2, padx=10, pady=5, sticky="w")

        ttk.Label(filter_frame, text="Degree:").grid(row=0, column=3, padx=5, pady=5, sticky="w")
        degree_options = ["All Degree Types", "EN Triennale", "EN Magistrale"]
        self.degree_type_combobox = ttk.Combobox(filter_frame, textvariable=self.degree_type_var, values=degree_options, state="readonly", width=20)
        self.degree_type_combobox.grid(row=0, column=4, padx=5, pady=5, sticky="w")

        # --- Export Button ---
        export_button = ttk.Button(
            filter_frame,
            text="Export List to Excel...",
            command=self.export_to_excel # Link to the new export function
        )
        # Place it on the right side of the filter row
        export_button.grid(row=0, column=5, padx=(20, 5), pady=5, sticky="e")

        filter_frame.columnconfigure(1, weight=1) # Allow search entry to expand
        filter_frame.columnconfigure(5, weight=0) # Don't let export button expand unnecessarily

    def create_results_table(self):
        # (Identical - no changes needed here)
        table_frame = ttk.Frame(self.root, padding=(10, 0, 10, 5))
        table_frame.pack(expand=True, fill=tk.BOTH)
        columns = ("s_no", "name", "university", "degree", "language")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.tree.heading("s_no", text="S.No."); self.tree.heading("name", text="Course Name")
        self.tree.heading("university", text="University"); self.tree.heading("degree", text="Degree Type")
        self.tree.heading("language", text="Language")
        self.tree.column("s_no", width=50, anchor=tk.CENTER, stretch=tk.NO)
        self.tree.column("name", width=350); self.tree.column("university", width=300)
        self.tree.column("degree", width=150); self.tree.column("language", width=80, anchor=tk.CENTER)
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(expand=True, fill=tk.BOTH)

    def create_status_bar(self):
        # (Identical)
        status_frame = ttk.Frame(self.root, padding=(10, 5, 10, 0))
        status_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(status_frame, text="Initializing...")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", length=200, mode="determinate")
        self.progress_bar.pack(side=tk.RIGHT)

    def create_footer(self):
        # (Identical)
        footer_frame = ttk.Frame(self.root, padding=(10, 5, 10, 10))
        footer_frame.pack(fill=tk.X)
        separator = ttk.Separator(footer_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 5))
        contact_label = ttk.Label(footer_frame, text=CONTACT_INFO, font=self.footer_font, foreground="#333333")
        contact_label.pack(expand=True)

    def start_fetch_thread(self):
        # (Identical)
        if not self.is_loading:
            self.is_loading = True; self.all_courses = []; self.clear_treeview()
            self.status_label.config(text="Fetching courses...")
            self.progress_bar['value'] = 0; self.progress_bar['maximum'] = TOTAL_PAGES
            fetch_thread = threading.Thread(target=self.fetch_all_courses, daemon=True)
            fetch_thread.start()

    def fetch_all_courses(self):
        # (Identical)
        errors = []; session = requests.Session()
        for page in range(1, TOTAL_PAGES + 1):
            if FETCH_DELAY > 0: time.sleep(FETCH_DELAY)
            try:
                params = { "searchType": "u", "page": page, "tipoLaurea": "", "tipoClasse": "0", "durata": "", "lingua": "", "tipoAccesso": "", "modalitaErogazione": "", "searchText": "", "area": "", "order": "RND", "provincia": "", "provinciaSigla": "" }
                response = session.get(API_BASE_URL, params=params, timeout=25)
                response.raise_for_status(); data = response.json()
                page_courses = data.get("corsi", [])
                if page_courses: self.fetch_queue.put(("page_data", page_courses))
                self.fetch_queue.put(("progress", page))
            except requests.exceptions.Timeout: err="Timeout"; errors.append(f"{err} page {page}"); print(f"[Thr] {err} {page}"); self.fetch_queue.put(("error", f"{err} {page}"))
            except requests.exceptions.RequestException as e: err="ReqErr"; errors.append(f"{err} p{page}:{e}"); print(f"[Thr] {err} p{page}"); self.fetch_queue.put(("error", f"{err} p{page}"))
            except ValueError as e: err="JSONErr"; errors.append(f"{err} p{page}:{e}"); print(f"[Thr] {err} p{page}"); self.fetch_queue.put(("error", f"{err} p{page}"))
            except Exception as e: err="Exc"; errors.append(f"{err} p{page}:{e}"); print(f"[Thr] {err} p{page}"); self.fetch_queue.put(("error", f"{err} p{page}"))
        self.fetch_queue.put(("complete", errors)); print("[Thr] Fetch done.")

    def check_fetch_queue(self):
        # (Identical)
        processed_count = 0; max_to_process = 5
        try:
            while processed_count < max_to_process:
                msg_type, payload = self.fetch_queue.get_nowait()
                if msg_type == "progress": self.progress_bar['value'] = payload
                elif msg_type == "error": print(f"[GUI Err] {payload}")
                elif msg_type == "page_data":
                    if payload: self.all_courses.extend(payload); self.add_filtered_courses_to_tree(payload); processed_count += 1
                elif msg_type == "complete":
                    self.is_loading = False; self.progress_bar['value'] = TOTAL_PAGES
                    errs = payload; disp_cnt = len(self.tree.get_children())
                    if errs: fin_msg = f"Complete ({len(errs)} fetch errors). Total: {len(self.all_courses)}. Displayed: {disp_cnt}."
                    else: fin_msg = f"Complete. Total: {len(self.all_courses)}. Displayed: {disp_cnt}."
                    self.status_label.config(text=fin_msg); print("[GUI] Complete msg.")
                    break
                if self.is_loading:
                    disp_cnt = len(self.tree.get_children())
                    self.status_label.config(text=f"Fetching P{self.progress_bar['value']}/{TOTAL_PAGES} | Total:{len(self.all_courses)} | Show:{disp_cnt}")
        except queue.Empty: pass
        except Exception as e: print(f"[GUI Q Err] {e}")
        if self.is_loading or not self.fetch_queue.empty(): self.root.after(50, self.check_fetch_queue)

    def filter_course_list(self, course_list):
        # (Identical)
        search_term = self.search_var.get().lower().strip()
        english_only = self.english_only_var.get(); degree_type = self.degree_type_var.get()
        if not course_list: return []
        filtered = course_list
        if search_term:
            temp_search = []
            for c in filtered:
                if c:
                    cn = c.get("nomeCorsoEn", ""); un = c.get("nomeStruttura", "")
                    nm = isinstance(cn, str) and search_term in cn.lower()
                    um = isinstance(un, str) and search_term in un.lower()
                    if nm or um: temp_search.append(c)
            filtered = temp_search
        if english_only: filtered = [c for c in filtered if c and c.get("lingua") == "EN"]
        if degree_type != "All Degree Types": filtered = [c for c in filtered if c and isinstance(c.get("tipoLaurea"), dict) and c["tipoLaurea"].get("descrizioneEn") == degree_type]
        return filtered

    def add_filtered_courses_to_tree(self, courses_to_add):
        # (Identical)
        if not courses_to_add: return
        matching_courses = self.filter_course_list(courses_to_add)
        if not matching_courses: return
        start_s_no = len(self.tree.get_children()) + 1
        for i, course in enumerate(matching_courses):
             if course:
                s_no = start_s_no + i; name = course.get("nomeCorsoEn", "N/A")
                uni = course.get("nomeStruttura", "N/A"); deg_d = course.get("tipoLaurea")
                deg = deg_d.get("descrizioneEn", "N/A") if isinstance(deg_d, dict) else "N/A"
                lang = course.get("lingua", "N/A"); vals = (s_no, name, uni, deg, lang)
                if all(isinstance(v, (str, int, float)) for v in vals):
                     try: self.tree.insert("", tk.END, values=vals)
                     except tk.TclError as e: print(f"TclErr insert {course.get('id', '?')}: {e} Vals:{vals}")
                else: print(f"Skip insert bad type {course.get('id', '?')} Vals:{vals}")

    def clear_treeview(self):
        # (Identical)
        for item in self.tree.get_children():
            try: self.tree.delete(item)
            except tk.TclError as e: print(f"TclErr delete {item}: {e}")

    def schedule_filter_update(self, *args):
        # (Identical)
        if self._after_id_filter: self.root.after_cancel(self._after_id_filter)
        self._after_id_filter = self.root.after(400, self.apply_full_filter)

    def apply_full_filter(self):
        # (Identical)
        self._after_id_filter = None
        if self.is_loading and self.progress_bar['value'] < TOTAL_PAGES: return
        print("[Filter] Applying full filter...")
        start_time = time.time()
        self.clear_treeview()
        filtered_all = self.filter_course_list(self.all_courses)
        self.add_filtered_courses_to_tree(filtered_all)
        end_time = time.time()
        print(f"[Filter] Full filter done in {end_time - start_time:.3f}s.")
        disp_cnt = len(self.tree.get_children())
        stat_msg = f"Filtered: {disp_cnt} matching courses (of {len(self.all_courses)} total fetched)."
        if self.is_loading: stat_msg += " (Still loading...)"
        self.status_label.config(text=stat_msg)

    # --- NEW METHOD: Export to Excel ---
    def export_to_excel(self):
        """Exports the currently displayed data in the Treeview to an Excel file after password verification."""
        print("Export button clicked.")

        # 1. Get Password
        entered_password = simpledialog.askstring(
            "Password Required",
            "Enter the password to export:",
            show='*' # Mask the input
        )

        # Handle cancellation
        if entered_password is None:
            print("Export cancelled by user (password dialog).")
            return

        # Check password
        if entered_password != EXPORT_PASSWORD:
            messagebox.showerror("Export Error", "Incorrect password.")
            print("Incorrect password entered.")
            return

        print("Password correct.")

        # 2. Get Data from Treeview
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("Export Info", "There is no data currently displayed to export.")
            print("No data in Treeview to export.")
            return

        data_to_export = []
        for item_id in items:
            values = self.tree.item(item_id, 'values')
            # Convert the tuple from treeview (which might contain numbers as strings)
            # back to appropriate types if needed, but for direct export it's often fine.
            # Ensure S.No is integer if possible
            try:
                s_no = int(values[0])
            except (ValueError, IndexError):
                s_no = values[0] # Keep as is if conversion fails
            # Keep others as they are, pandas handles mixed types reasonably well
            data_to_export.append((s_no,) + values[1:]) # Reconstruct tuple

        print(f"Gathered {len(data_to_export)} rows from Treeview.")

        # 3. Choose Save Location
        filepath = filedialog.asksaveasfilename(
            title="Save Exported List As...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        # Handle cancellation
        if not filepath:
            print("Export cancelled by user (save file dialog).")
            return

        print(f"User chose to save to: {filepath}")

        # 4. Create DataFrame and Export
        try:
            # Define column names matching the Treeview
            column_names = ["S.No.", "Course Name", "University", "Degree Type", "Language"]
            df = pd.DataFrame(data_to_export, columns=column_names)

            # Export to Excel using openpyxl engine
            df.to_excel(filepath, index=False, engine='openpyxl')

            messagebox.showinfo("Export Successful", f"Data successfully exported to:\n{filepath}")
            print("Export successful.")

        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred during export:\n{e}")
            print(f"Export failed: {e}")


# --- Main Execution ---
if __name__ == "__main__":
    # Ensure pandas and openpyxl are installed, otherwise show error and exit
    try:
        import pandas
        import openpyxl # Just check if it can be imported
    except ImportError:
        root = tk.Tk()
        root.withdraw() # Hide the main empty window
        messagebox.showerror(
            "Missing Libraries",
            "Required libraries 'pandas' and 'openpyxl' are not installed.\nPlease install them using:\npip install pandas openpyxl"
        )
        root.destroy()
        exit() # Stop execution if libraries are missing

    root = tk.Tk()
    style = ttk.Style()
    available_themes = style.theme_names()
    if 'clam' in available_themes: style.theme_use('clam')
    elif 'alt' in available_themes: style.theme_use('alt')
    app = CourseFinderApp(root)
    root.mainloop()