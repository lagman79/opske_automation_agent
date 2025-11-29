#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
opske_final_v7.py
Αναβαθμισμένη έκδοση: αφαιρεμένα hardcoded credentials + modal login
Βασισμένο στο opske_final_v6.py που παρείχε ο χρήστης.
"""

import os
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
import re

# -------------------------- ΡΥΘΜΙΣΕΙΣ --------------------------
SCREEN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshots")
os.makedirs(SCREEN_DIR, exist_ok=True)

ALLOWED_EXT = {'.pdf', '.jpg', '.jpeg', '.png'}
MAX_SIZE_MB = 10

# Σχήματα και βάρη για έλεγχο ΑΦΜ
AFM_REGEX = r'^(.+) - (\d{9})$'
AFM_WEIGHTS = [256, 128, 64, 32, 16, 8, 4, 2]

# -------------------------- ΒΟΗΘΗΤΙΚΕΣ ΣΥΝΑΡΤΗΣΕΙΣ --------------------------
def pick_date_from_calendar(page, day, month, year):
    page.wait_for_selector(".p-datepicker:visible")
    calendar = page.locator(".p-datepicker:visible").last

    today = datetime.today()
    max_p = today.replace(year=today.year-100)
    max_f = today.replace(year=today.year+10)
    tgt = datetime(year=year, month=month, day=1)

    if tgt < max_p or tgt > max_f:
        raise RuntimeError(f"Ημερομηνία εκτός ορίων: {tgt.strftime('%d/%m/%Y')}")

    month_map = {"Ιανουάριος": 1, "Φεβρουάριος": 2, "Μάρτιος": 3, "Απρίλιος": 4,
                 "Μάιος": 5, "Ιούνιος": 6, "Ιούλιος": 7, "Αύγουστος": 8,
                 "Σεπτέμβριος": 9, "Οκτώβριος": 10, "Νοέμβριος": 11, "Δεκέμβριος": 12}

    # loop until calendar shows target month-year
    while True:
        hdr_text = calendar.locator(".p-datepicker-title").text_content()
        if not hdr_text:
            time.sleep(0.1)
            continue

        hdr = hdr_text.strip().split()
        if len(hdr) < 2:
            time.sleep(0.1)
            continue

        c_y, c_m = int(hdr[1]), month_map.get(hdr[0], None)
        if c_m is None:
            time.sleep(0.1)
            continue

        if c_y == year and c_m == month:
            break

        if tgt > datetime(year=c_y, month=c_m, day=1):
            calendar.locator(".p-datepicker-next").click()
        else:
            calendar.locator(".p-datepicker-prev").click()
        time.sleep(0.15)

    day_cell = calendar.locator(f"td:not(.p-disabled):not(.p-datepicker-other-month) >> text={int(day)}").first
    if day_cell.count() == 0:
        raise RuntimeError(f"Δεν βρέθηκε ενεργή μέρα {day}/{month}/{year}")

    day_cell.click()
    time.sleep(0.5)

def excel_date_to_parts(val):
    if pd.isna(val):
        today = datetime.today()
        return today.day, today.month, today.year
    if isinstance(val, (int, float)):
        d = datetime(1899, 12, 30) + timedelta(days=int(val))
    else:
        d = pd.to_datetime(val)
    return d.day, d.month, d.year

# -------------------------- LOGIN MODAL --------------------------
def prompt_credentials(root):
    """
    Εμφανίζει modal παράθυρο που ζητάει username/password.
    Επιστρέφει (username, password) ή None αν ο χρήστης ακυρώσει.
    """
    creds = {"username": None, "password": None}
    dlg = tk.Toplevel(root)
    dlg.title("Σύνδεση OPSKE")
    dlg.geometry("360x150")
    dlg.resizable(False, False)
    dlg.transient(root)
    dlg.grab_set()

    ttk.Label(dlg, text="Username:").pack(anchor='w', padx=10, pady=(10,0))
    user_ent = ttk.Entry(dlg)
    user_ent.pack(fill='x', padx=10)

    ttk.Label(dlg, text="Password:").pack(anchor='w', padx=10, pady=(8,0))
    pass_ent = ttk.Entry(dlg, show="*")
    pass_ent.pack(fill='x', padx=10)

    btn_frame = ttk.Frame(dlg)
    btn_frame.pack(fill='x', pady=10, padx=10)
    result = {"ok": False}

    def on_ok():
        u = user_ent.get().strip()
        p = pass_ent.get()
        if not u or not p:
            messagebox.showwarning("Σφάλμα", "Συμπλήρωσε username και password.")
            return
        creds["username"], creds["password"] = u, p
        result["ok"] = True
        dlg.destroy()

    def on_cancel():
        dlg.destroy()

    ttk.Button(btn_frame, text="OK", command=on_ok).pack(side='right', padx=5)
    ttk.Button(btn_frame, text="Άκυρο", command=on_cancel).pack(side='right')
    root.wait_window(dlg)

    return (creds["username"], creds["password"]) if result["ok"] else None

# -------------------------- GUI ΚΛΑΣΗ --------------------------
class AppGUI:
    def __init__(self, root, username, password):
        self.root = root
        self.root.title("OPSKE Agent - Final")
        self.root.geometry("1320x925")
        self.username = username
        self.password = password

        self.folder = None
        self.excel_path = None
        self.play_flag = threading.Event()
        self.play_flag.set()
        self.results = []
        self.submit_flag = False
        self.total_files = 0
        self.processed = 0
        self.check_performed = False
        self.check_passed = False

        # --- Buttons ---
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Επιλογή Excel", command=self.pick_excel).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Επιλογή Φακέλου", command=self.pick_folder).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Έλεγχος Excel & Αρχείων", command=self.check_files).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Έναρξη (Αποθήκευση)", command=lambda: self.start_thread(submit=False)).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Έναρξη (Υποβολή)", command=lambda: self.start_thread(submit=True)).pack(side='left', padx=5)
        self.pause_btn = ttk.Button(btn_frame, text="Pause", command=self.toggle_pause)
        self.pause_btn.pack(side='left', padx=5)

        # Progress bar
        self.progress_frame = tk.Frame(root)
        self.progress_frame.pack(fill='x', padx=10, pady=5)
        self.progress_label = tk.Label(self.progress_frame, text="Πρόοδος: 0/0 (0%)")
        self.progress_label.pack(side='left')
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient='horizontal', length=200, mode='determinate')
        self.progress_bar.pack(side='left', padx=5)

        # Current upload tree
        ttk.Label(root, text="Τρέχον upload:").pack(anchor='w', padx=10)
        self.current_tree = ttk.Treeview(root, columns=("Progress",), height=5, show='tree headings')
        self.current_tree.heading("#0", text="Όνομα Αρχείου")
        self.current_tree.heading("Progress", text="Πρόοδος")
        self.current_tree.pack(fill='x', padx=10)

        # Result tree
        ttk.Label(root, text="Αποτελέσματα:").pack(anchor='w', padx=10)
        result_container = ttk.Frame(root)
        result_container.pack(fill='both', expand=True, padx=10)
        self.result_tree = ttk.Treeview(result_container, columns=("Status",), height=10, show='tree headings')
        self.result_tree.heading("#0", text="Όνομα Αρχείου")
        self.result_tree.heading("Status", text="Κατάσταση")
        result_scroll = ttk.Scrollbar(result_container, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=result_scroll.set)
        self.result_tree.pack(side='left', fill='both', expand=True)
        result_scroll.pack(side='right', fill='y')

        self.result_tree.tag_configure("ok", foreground="green")
        self.result_tree.tag_configure("fail", foreground="red")
        self.result_tree.tag_configure("info", foreground="blue")

        # Summary
        ttk.Label(root, text="Σύνοψη:").pack(anchor='w', padx=10)
        summary_container = ttk.Frame(root)
        summary_container.pack(fill='both', expand=True, padx=10)
        self.summary_tree = ttk.Treeview(summary_container, columns=("Reason",), height=8, show='tree headings')
        self.summary_tree.heading("#0", text="Όνομα Αρχείου")
        self.summary_tree.heading("Reason", text="Λόγος Αποτυχίας")
        summary_scroll = ttk.Scrollbar(summary_container, orient="vertical", command=self.summary_tree.yview)
        self.summary_tree.configure(yscrollcommand=summary_scroll.set)
        self.summary_tree.pack(side='left', fill='both', expand=True)
        summary_scroll.pack(side='right', fill='y')

        self.summary_tree.tag_configure("fail", foreground="red")
        self.summary_tree.tag_configure("info", foreground="blue")

    # --- GUI helpers ---
    def pick_excel(self):
        f = filedialog.askopenfilename(title="Επίλεξε αρχείο Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if f:
            self.excel_path = f
            self.check_performed = False
            self.check_passed = False
            messagebox.showinfo("Excel", f"Επιλέχθηκε: {os.path.basename(f)}")

    def pick_folder(self):
        f = filedialog.askdirectory(title="Επιλέξε φάκελο με αρχεία")
        if f:
            self.folder = f
            self.check_performed = False
            self.check_passed = False
            messagebox.showinfo("Φάκελος", f"Επιλέχθηκε: {f}")

    def toggle_pause(self):
        if self.play_flag.is_set():
            self.play_flag.clear()
            self.pause_btn.config(text="Play")
        else:
            self.play_flag.set()
            self.pause_btn.config(text="Pause")

    def validate_file(self, fpath):
        ext = os.path.splitext(fpath)[1].lower().strip()
        if ext not in ALLOWED_EXT:
            return False, f"Μη επιτρεπτή επέκταση: '{ext}' (επιτρέπονται: {', '.join(sorted(ALLOWED_EXT))})"
        size = os.path.getsize(fpath) / (1024*1024)
        if size > MAX_SIZE_MB:
            return False, f"Υπέρβαση μεγέθους: {size:.2f} MB"
        return True, None

    def add_current(self, name):
        self.current_tree.insert("", "end", iid=name, text=name, values=("0 %",))

    def update_progress(self, name, val):
        self.current_tree.set(name, "Progress", f"{val} %")
        self.root.update_idletasks()

    def remove_current(self, name):
        try:
            self.current_tree.delete(name)
        except Exception:
            pass

    def add_result(self, name, ok, reason=""):
        if ok:
            self.result_tree.insert("", "end", text=name, values=("✓",), tags=("ok",))
        else:
            self.result_tree.insert("", "end", text=name, values=("✗",), tags=("fail",))
            self.results.append((name, reason))

    def update_total_progress(self, current, total):
        self.processed = current
        self.total_files = total
        percent = int((current / total) * 100) if total else 0
        self.progress_label.config(text=f"Πρόοδος: {current}/{total} ({percent}%)")
        self.progress_bar["value"] = percent
        self.root.update_idletasks()

    def run_agent(self):
        with sync_playwright() as pw:
            try:
                self.run(pw, self.submit_flag)
            except Exception as e:
                messagebox.showerror("Σφάλμα", str(e))
            finally:
                self.fill_summary()

    def run(self, playwright, submit=False):
        # Διαβάζει excel και κάνει login/υποβολές
        df = pd.read_excel(self.excel_path, dtype=str)
        if "Αποθήκευση" not in df.columns: df["Αποθήκευση"] = ""
        if "Υποβολή" not in df.columns: df["Υποβολή"] = ""
        if "Ημ/νία & ώρα υποβολής" not in df.columns:
            df["Ημ/νία & ώρα υποβολής"] = ""

        skip_indices = []
        for idx, row in df.iterrows():
            fname = str(row["Όνομα Αρχείου"]).strip()
            if not fname or fname == 'nan':
                skip_indices.append(idx)
                continue

            h_val = str(row.get("Αποθήκευση", "")).strip().upper()
            i_val = str(row.get("Υποβολή", "")).strip().upper()

            if i_val == "TRUE":
                self.results.append((fname, "Το αρχείο είχε υποβληθεί ήδη!"))
                skip_indices.append(idx)
                self.summary_tree.insert("", "end", text=fname, values=("Το αρχείο είχε υποβληθεί ήδη!",), tags=("info",))
            elif h_val == "TRUE" and not submit:
                self.results.append((fname, "Το αρχείο είχε αποθηκευτεί ήδη!"))
                skip_indices.append(idx)
                self.summary_tree.insert("", "end", text=fname, values=("Το αρχείο είχε αποθηκευτεί ήδη!",), tags=("info",))

        if submit:
            mask = df["Υποβολή"].astype(str).str.strip().str.upper() == "TRUE"
        else:
            h_mask = df["Αποθήκευση"].astype(str).str.strip().str.upper() == "TRUE"
            i_mask = df["Υποβολή"].astype(str).str.strip().str.upper() == "TRUE"
            mask = h_mask | i_mask

        if mask.all():
            msg = "Όλα τα δικαιολογητικά έχουν υποβληθεί" if submit else "Όλα τα δικαιολογητικά έχουν αποθηκευτεί/υποβληθεί"
            self.result_tree.insert("", "end", text=msg, values=("ℹ",), tags=("info",))
            return

        to_process = df[~mask].copy()
        total = len(to_process)
        self.update_total_progress(0, total)

        browser = playwright.chromium.launch(headless=False, slow_mo=300)
        page = browser.new_page()
        page.set_default_timeout(30_000)

        page.goto("https://app.opske.gr/ ")
        page.click("text=Σύνδεση ΑΑΔΕ")
        # ΧΡΗΣΗ credentials από self
        page.fill("#j_username", self.username)
        page.fill("#j_password", self.password)
        page.click("button:has-text('Σύνδεση')")
        page.wait_for_selector("#btn-submit")
        page.click("#btn-submit")
        page.click("text=Τα Δικαιολογητικά Δικαιούχου μου")

        for idx, (index, row) in enumerate(to_process.iterrows(), 1):
            self.play_flag.wait()
            base_name = str(row["Όνομα Αρχείου"]).strip()

            real_path = None
            for f in os.listdir(self.folder):
                if os.path.splitext(f)[0] == base_name:
                    candidate = os.path.join(self.folder, f)
                    if os.path.isfile(candidate):
                        real_path = candidate
                        break
            if real_path is None:
                self.add_result(base_name, False, f"Δεν βρέθηκε αρχείο με όνομα: {base_name}")
                self.update_total_progress(idx, total)
                continue

            ok, reason = self.validate_file(real_path)
            if not ok:
                self.add_result(base_name, False, reason)
                self.update_total_progress(idx, total)
                continue

            self.add_current(base_name)
            self.update_progress(base_name, 10)
            try:
                d_i, m_i, y_i = excel_date_to_parts(row["Ημερομηνία έκδοσης δικαιολογητικού"])
                d_e, m_e, y_e = excel_date_to_parts(row["Ημερομηνία λήξης δικαιολογητικού"])

                row_dict = {
                    "ben": row["Επωνυμία – ΑΦΜ"],
                    "app": row["Κωδικός έργου"],
                    "doc": row["Κωδικός Δικαιολογητικού"],
                    "note": row["Παρατηρήσεις ΟΠΣΚΕ"],
                    "d_i": d_i, "m_i": m_i, "y_i": y_i,
                    "d_e": d_e, "m_e": m_e, "y_e": y_e,
                    "fname": row["Όνομα αρχείου"]
                }

                self.upload_row(page, row_dict, real_path, submit)
                self.remove_current(base_name)
                self.add_result(base_name, True)

                # Ενημέρωση Excel
                df.loc[index, "Αποθήκευση"] = "TRUE"
                df.loc[index, "Υποβολή"] = "TRUE" if submit else "FALSE"
                now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                df.loc[index, "Ημ/νία & ώρα υποβολής"] = now_str

            except Exception as e:
                self.remove_current(base_name)
                self.add_result(base_name, False, str(e))
                print(f"Σφάλμα στο αρχείο {base_name}: {e}")

                try:
                    print("Προσπάθεια επαναφοράς στην αρχική σελίδα.")
                    page.goto("https://app.opske.gr/dashboard/invoices/supporting-document/my-supporting-document", timeout=10000)
                    page.wait_for_load_state("networkidle")
                except:
                    print("Η επαναφορά απέτυχε, συνεχίζουμε.")

            self.update_total_progress(idx, total)

        df.to_excel(self.excel_path, index=False)
        browser.close()

    def upload_row(self, page, row, fpath, submit=False):
        ben = str(row["ben"]).strip()
        app = str(row["app"]).strip()
        doc = str(row["doc"]).strip()
        note = str(row["note"]).strip()

        # 1. Κλικ Προσθήκη
        try:
            page.wait_for_selector("text=Προσθήκη", state="visible", timeout=10000)
            page.click("text=Προσθήκη")
        except:
            print("Δεν βρέθηκε το κουμπί Προσθήκη - Προσπάθεια επαναφοράς...")
            return

        time.sleep(1)

        # 2-5. Dropdowns & Comments
        page.click("(//span[@role='combobox'])[1]")
        flt = ben[:6]
        suf = ben[-3:]
        page.fill("//input[contains(@class,'p-dropdown-filter')]", flt)
        for li in page.locator("//li[@role='option']").all():
            if suf in li.text_content().strip():
                li.click()
                break

        page.click("div.p-multiselect-label-container")
        for li in page.locator("//li[@role='option']").all():
            if app.upper() in li.text_content().strip().upper():
                li.click()
                break
        page.click("body")

        page.click("(//span[@role='combobox'])[2]")
        page.fill("//input[contains(@class,'p-dropdown-filter')]", doc[:5])
        page.click("(//li[@role='option'])[1]")

        page.fill("#comments", note)

        # Dates
        page.click("#issueDate")
        pick_date_from_calendar(page, row['d_i'], row['m_i'], row['y_i'])
        page.click("#expirationDate")
        pick_date_from_calendar(page, row['d_e'], row['m_e'], row['y_e'])

        # Upload file
        with page.expect_file_chooser() as fc:
            page.click("text=Επιλογή Αρχείου")
        fc.value.set_files(fpath)

        # Save
        page.click("text=Αποθήκευση")
        print("① Κλικ Αποθήκευση")

        try:
            page.wait_for_selector('div.p-toast-detail:text("Η εγγραφή αποθηκεύτηκε επιτυχώς")', state='visible', timeout=15_000)
            print("② Toast εμφανίστηκε")
            time.sleep(2)
        except Exception as e:
            print("② Toast ΔΕΝ εμφανίστηκε:", e)

        # Submit if requested
        if submit:
            submit_btn = page.locator('//button[contains(., "Υποβολή")]').first
            submit_btn.scroll_into_view_if_needed()
            time.sleep(0.5)

            if submit_btn.is_enabled():
                submit_btn.click()
                print("③ Κλικ Υποβολής")
                try:
                    page.wait_for_selector('//button[contains(., "Υποβολή") and @disabled]', timeout=10_000)
                    print("④ Η Υποβολή καταχωρήθηκε")
                except: pass
                time.sleep(2)
                page.screenshot(path=os.path.join(SCREEN_DIR, f"submitted_{row['fname']}.png"))

        # Επιστροφή
        try:
            page.click('a[href="/dashboard/invoices/supporting-document/my-supporting-document"]', timeout=3_000)
        except:
            try:
                page.click('button:has-text("Επιστροφή")')
            except:
                pass
            time.sleep(1)

    def validate_afm(self, afm):
        if not afm.isdigit() or len(afm) != 9:
            return False
        digits = [int(d) for d in afm]
        total = sum(d * w for d, w in zip(digits[:8], AFM_WEIGHTS))
        remainder = total % 11
        check_digit = 0 if remainder == 10 else remainder
        return digits[8] == check_digit

    def check_excel_structure(self, df):
        errors = []
        GREEK_CHARS = 'α-ωά-ώΑ-ΩΆ-Ώ'

        for idx, row in df.iterrows():
            row_num = idx + 2
            file_name = str(row.get("Όνομα αρχείου", "")).strip()

            afm_field = str(row.get("Επωνυμία – ΑΦΜ", "")).strip()
            match = re.match(AFM_REGEX, afm_field)
            if not match:
                errors.append((file_name, f"Στήλη A (γραμμή {row_num}): Λάθος format '{afm_field}'"))
                continue

            company, afm = match.groups()
            if not self.validate_afm(afm):
                errors.append((file_name, f"Στήλη A (γραμμή {row_num}): Μη έγκυρο ΑΦΜ '{afm}'"))

            project_code = str(row.get("Κωδικός έργου", "")).strip()
            if not re.match(f'^[a-zA-Z{GREEK_CHARS}0-9]+-\\d+$', project_code):
                errors.append((file_name, f"Στήλη B (γραμμή {row_num}): Λάθος format '{project_code}'"))

            doc_code = str(row.get("Κωδικός Δικαιολογητικού", "")).strip()
            if not re.match(r'^\d{2}\.\d{2}$', doc_code):
                errors.append((file_name, f"Στήλη C (γραμμή {row_num}): Λάθος format '{doc_code}'"))

            try:
                issue_date = pd.to_datetime(row.get("Ημερομηνία έκδοσης δικαιολογητικού"), errors='coerce')
                if pd.isna(issue_date):
                    errors.append((file_name, f"Στήλη E (γραμμή {row_num}): Μη έγκυρη ημερομηνία"))
                elif issue_date > datetime.now():
                    errors.append((file_name, f"Στήλη E (γραμμή {row_num}): Ημερομηνία έκδοσης μεταγενέστερη από σήμερα"))
            except:
                errors.append((file_name, f"Στήλη E (γραμμή {row_num}): Μη έγκυρη ημερομηνία"))

            try:
                expiry_date = pd.to_datetime(row.get("Ημερομηνία λήξης δικαιολογητικού"), errors='coerce')
                if pd.isna(expiry_date):
                    errors.append((file_name, f"Στήλη F (γραμμή {row_num}): Μη έγκυρη ημερομηνία"))
                elif expiry_date < datetime.now():
                    errors.append((file_name, f"Στήλη F (γραμμή {row_num}): Ημερομηνία λήξης προγενέστερη από σήμερα"))
            except:
                errors.append((file_name, f"Στήλη F (γραμμή {row_num}): Μη έγκυρη ημερομηνία"))

        return errors

    def check_files(self):
        if not self.excel_path:
            messagebox.showwarning("Excel", "Παρακαλώ επιλέξτε αρχείο Excel!")
            return
        if not self.folder:
            messagebox.showwarning("Φάκελος", "Παρακαλώ επιλέξτε φάκελο!")
            return

        for i in (*self.result_tree.get_children(), *self.summary_tree.get_children()):
            self.result_tree.delete(i) if i in self.result_tree.get_children() else None
            self.summary_tree.delete(i) if i in self.summary_tree.get_children() else None

        self.check_performed = True
        self.ok_count = 0
        self.problem_count = 0

        try:
            df = pd.read_excel(self.excel_path, dtype=str)
            required_cols = ["Επωνυμία – ΑΦΜ", "Κωδικός έργου", "Κωδικός Δικαιολογητικού",
                           "Όνομα αρχείου", "Ημερομηνία έκδοσης δικαιολογητικού",
                           "Ημερομηνία λήξης δικαιολογητικού", "Παρατηρήσεις ΟΠΣΚΕ"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                messagebox.showerror("Σφάλμα Excel", f"Λείπουν οι στήλες: {', '.join(missing_cols)}")
                self.check_passed = False
                return
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Δεν μπόρεσε να διαβαστεί το Excel: {e}")
            self.check_passed = False
            return

        excel_errors = self.check_excel_structure(df)
        for file_name, error_msg in excel_errors:
            self.problem_count += 1
            self.result_tree.insert("", "end", text=file_name, values=("Σφάλμα Excel ✗",), tags=("fail"))
            self.result_tree.insert("", "end", text=f"  → {error_msg}", values=("",), tags=("fail"))

        for idx, row in df.iterrows():
            fname = str(row["Όνομα αρχείου"]).strip()
            if not fname or fname == 'nan':
                continue

            found = False
            for f in os.listdir(self.folder):
                if os.path.splitext(f)[0] == fname:
                    fpath = os.path.join(self.folder, f)
                    if os.path.isfile(fpath):
                        found = True
                        ok, reason = self.validate_file(fpath)
                        if ok:
                            self.ok_count += 1
                            self.result_tree.insert("", "end", text=fname, values=("Βρέθηκε ✓",), tags=("ok"))
                        else:
                            self.problem_count += 1
                            self.result_tree.insert("", "end", text=fname, values=("Σφάλμα ✗",), tags=("fail"))
                            self.result_tree.insert("", "end", text=f"  → {reason}", values=("",), tags=("fail"))
                        break

            if not found:
                self.problem_count += 1
                reason = f"Δεν βρέθηκε αρχείο με το όνομα: {fname}"
                self.result_tree.insert("", "end", text=fname, values=("Δεν βρέθηκε ✗",), tags=("fail"))
                self.result_tree.insert("", "end", text=f"  → {reason}", values=("",), tags=("fail"))

        self.check_passed = (self.ok_count > 0 and len(excel_errors) == 0)

        self.result_tree.insert("", "end", text="", values=("",))
        if self.check_passed:
            if self.problem_count == 0:
                msg = "Έλεγχος ολοκληρώθηκε με επιτυχία! Όλα είναι ΟΚ."
            else:
                msg = f"Έλεγχος ολοκληρώθηκε! Βρέθηκαν {self.ok_count} αρχεία ΟΚ και {self.problem_count} προβλήματα. Θα επεξεργαστούν μόνο τα ΟΚ."
            self.result_tree.insert("", "end", text=msg, values=("✓",), tags=("ok"))
        else:
            if excel_errors:
                msg = f"Ο έλεγχος βρήκε {len(excel_errors)} σφάλματα στο Excel και {self.problem_count} προβλήματα στα αρχεία. Διορθώστε όλα τα σφάλματα!"
            else:
                msg = "Ο έλεγχος βρήκε προβλήματα σε όλα τα αρχεία. Διορθώστε τα πριν συνεχίσετε!"
            self.result_tree.insert("", "end", text=msg, values=("✗",), tags=("fail"))

    def start_thread(self, submit=False):
        if not self.excel_path:
            messagebox.showwarning("Excel", "Παρακαλώ επιλέξτε αρχείο Excel!")
            return
        if not self.folder:
            messagebox.showwarning("Φάκελος", "Παρακαλώ επιλέξτε φάκελο!")
            return
        if not self.check_performed:
            messagebox.showwarning("Έλεγχος", "Παρακαλώ εκτελέστε πρώτα τον 'Έλεγχο Excel & Αρχείων'!")
            return
        if not self.check_passed:
            messagebox.showwarning("Έλεγχος", "Ο έλεγχος βρήκε σφάλματα. Διορθώστε τα πριν συνεχίσετε!")
            return

        for i in (*self.current_tree.get_children(), *self.result_tree.get_children(), *self.summary_tree.get_children()):
            self.current_tree.delete(i) if i in self.current_tree.get_children() else None
            self.result_tree.delete(i)   if i in self.result_tree.get_children()   else None
            self.summary_tree.delete(i)  if i in self.summary_tree.get_children()  else None
        self.results.clear()
        self.submit_flag = submit
        threading.Thread(target=self.run_agent, daemon=True).start()

    def fill_summary(self):
        for name, reason in self.results:
            if "ήδη" in reason:
                self.summary_tree.insert("", "end", text=name, values=(reason,), tags=("info",))
            else:
                self.summary_tree.insert("", "end", text=name, values=(reason,), tags=("fail",))

# -------------------------- MAIN --------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1x1+0+0")  # Ελάχιστο παράθυρο που φαίνεται μόνο στη γραμμή εργασιών
    root.update()  # ΚΛΙΔΙ: Ενημέρωση του GUI loop
    
    creds = prompt_credentials(root)
    if not creds:
        messagebox.showinfo("Ακύρωση", "Δεν εισήχθησαν credentials. Το πρόγραμμα τερματίζει.")
        root.destroy()
        raise SystemExit(0)

    username, password = creds
    root.deiconify()  # εμφανίζουμε το κύριο παράθυρο
    root.geometry("1320x925")  # Ορίστε το μέγεθος εδώ
    app = AppGUI(root, username, password)
    root.mainloop()