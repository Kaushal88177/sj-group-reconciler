import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import os
import platform

APP_NAME = "SJ Group - GST Reconciler Pro"
VERSION = "3.0 - Optional Column Mapping"

def normalize_voucher(v):
    if v is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(v):
            return ""
    except: pass
    return str(v).strip().upper().replace(" ", "")

def to_float(x):
    try:
        if x is None:
            return 0.0
        try:
            import pandas as pd
            if pd.isna(x):
                return 0.0
        except: pass
        return float(str(x).replace(",","").strip())
    except:
        return 0.0

def get_sheet_headers(file_path, header_row):
    wb = openpyxl.load_workbook(file_path, read_only=True)
    ws = wb.active
    headers = []
    for c in range(1, ws.max_column+1):
        v = ws.cell(row=header_row, column=c).value
        headers.append(str(v).strip() if v else f"Column {c}")
    wb.close()
    return headers

def find_header_row_guess(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True)
    ws = wb.active
    for r in range(1, 8):
        row_vals = [str(ws.cell(r,c).value).lower() if ws.cell(r,c).value else "" for c in range(1, 20)]
        if any("voucher" in v or "invoice" in v for v in row_vals):
            wb.close()
            return r
    wb.close()
    return 3

def auto_map_column(headers, keywords):
    headers_lower = [h.lower() for h in headers]
    for kw in keywords:
        kw = kw.lower()
        for i, h in enumerate(headers_lower):
            if kw in h:
                return headers[i]
    return headers[0] if headers else ""

def process_with_mapping(gstr_path, tally_path, gstr_map, tally_map, g_header, t_header, use_fields, output_dir, match_text="Match", not_match_text="Not Matching"):
    wb_g = openpyxl.load_workbook(gstr_path)
    ws_g = wb_g.active
    wb_t = openpyxl.load_workbook(tally_path)
    ws_t = wb_t.active

    def build_col_index(ws, header_row):
        m = {}
        for c in range(1, ws.max_column+1):
            v = ws.cell(row=header_row, column=c).value
            if v:
                m[str(v).strip()] = c
        return m

    g_headers = get_sheet_headers(gstr_path, g_header)
    t_headers = get_sheet_headers(tally_path, t_header)

    g_col_index = build_col_index(ws_g, g_header)
    t_col_index = build_col_index(ws_t, t_header)

    def get_col(col_map, col_index_map, key):
        selected_header_name = col_map.get(key)
        if not selected_header_name:
            return None
        if selected_header_name in col_index_map:
            return col_index_map[selected_header_name]
        for h, idx in col_index_map.items():
            if h.strip() == selected_header_name.strip():
                return idx
        if selected_header_name.startswith("Column"):
            try:
                num = int(selected_header_name.split(" ")[1])
                return num
            except:
                pass
        return None

    # Get col numbers
    cols_G = {}
    cols_T = {}
    for field_key in ['voucher','gross','taxable','igst','cgst','sgst','remark']:
        cols_G[field_key] = get_col(gstr_map, g_col_index, field_key)
        cols_T[field_key] = get_col(tally_map, t_col_index, field_key)

    if not cols_G['voucher'] or not cols_G['remark'] or not cols_T['voucher'] or not cols_T['remark']:
        raise ValueError("Voucher No. and Remark_1 mapping zaruri hai!")

    # Build keys only from selected use_fields
    # use_fields is dict like {'voucher': True, 'gross': False, ...}
    def build_key(ws, row_idx, cols, use_fields):
        parts = []
        # Always include voucher if enabled, else skip? But voucher is mandatory for identity
        # We'll build in order
        if use_fields.get('voucher'):
            v_raw = ws.cell(row=row_idx, column=cols['voucher']).value
            parts.append(normalize_voucher(v_raw))
        if use_fields.get('gross') and cols['gross']:
            parts.append(str(round(to_float(ws.cell(row=row_idx, column=cols['gross']).value),2)))
        if use_fields.get('taxable') and cols['taxable']:
            parts.append(str(round(to_float(ws.cell(row=row_idx, column=cols['taxable']).value),2)))
        if use_fields.get('igst') and cols['igst']:
            parts.append(str(round(to_float(ws.cell(row=row_idx, column=cols['igst']).value),2)))
        if use_fields.get('cgst') and cols['cgst']:
            parts.append(str(round(to_float(ws.cell(row=row_idx, column=cols['cgst']).value),2)))
        if use_fields.get('sgst') and cols['sgst']:
            parts.append(str(round(to_float(ws.cell(row=row_idx, column=cols['sgst']).value),2)))
        # If user selected only 2 fields, key will have only 2 parts
        return "|".join(parts)

    tally_dict = {}
    for i in range(t_header+1, ws_t.max_row+1):
        v_raw = ws_t.cell(row=i, column=cols_T['voucher']).value
        if v_raw is None or str(v_raw).strip()=="":
            continue
        key = build_key(ws_t, i, cols_T, use_fields)
        if key:
            tally_dict[key] = True

    gstr_dict = {}
    for i in range(g_header+1, ws_g.max_row+1):
        v_raw = ws_g.cell(row=i, column=cols_G['voucher']).value
        if v_raw is None or str(v_raw).strip()=="":
            continue
        key = build_key(ws_g, i, cols_G, use_fields)
        if key:
            gstr_dict[key] = True

    matched_g = 0
    for i in range(g_header+1, ws_g.max_row+1):
        v_raw = ws_g.cell(row=i, column=cols_G['voucher']).value
        if v_raw is None or str(v_raw).strip()=="":
            continue
        # fix numeric formatting
        for fkey in ['gross','taxable','igst','cgst','sgst']:
            col = cols_G.get(fkey)
            if col:
                val = ws_g.cell(row=i, column=col).value
                if val is not None and val != "":
                    try:
                        ws_g.cell(row=i, column=col).value = float(str(val).replace(",",""))
                        ws_g.cell(row=i, column=col).number_format = '#,##0.00'
                    except: pass
        key = build_key(ws_g, i, cols_G, use_fields)
        if key in tally_dict:
            ws_g.cell(row=i, column=cols_G['remark']).value = match_text
            matched_g+=1
        else:
            ws_g.cell(row=i, column=cols_G['remark']).value = not_match_text

    matched_t = 0
    for i in range(t_header+1, ws_t.max_row+1):
        v_raw = ws_t.cell(row=i, column=cols_T['voucher']).value
        if v_raw is None or str(v_raw).strip()=="":
            continue
        for fkey in ['gross','taxable','igst','cgst','sgst']:
            col = cols_T.get(fkey)
            if col:
                val = ws_t.cell(row=i, column=col).value
                if val is not None and val != "":
                    try:
                        ws_t.cell(row=i, column=col).value = float(str(val).replace(",",""))
                        ws_t.cell(row=i, column=col).number_format = '#,##0.00'
                    except: pass
        key = build_key(ws_t, i, cols_T, use_fields)
        if key in gstr_dict:
            ws_t.cell(row=i, column=cols_T['remark']).value = match_text
            matched_t+=1
        else:
            ws_t.cell(row=i, column=cols_T['remark']).value = not_match_text

    os.makedirs(output_dir, exist_ok=True)
    out_g = os.path.join(output_dir, "SJGroup_GSTR_Matched.xlsx")
    out_t = os.path.join(output_dir, "SJGroup_Tally_Matched.xlsx")
    wb_g.save(out_g)
    wb_t.save(out_t)
    return out_g, out_t, matched_g, matched_t, ws_g.max_row-g_header, ws_t.max_row-t_header

class App:
    def __init__(self, root):
        self.root = root
        root.title(f"{APP_NAME} v{VERSION}")
        root.geometry("1000x820")
        root.configure(bg="#0f172a")

        tk.Label(root, text="SJ GROUP", font=("Segoe UI Black", 24), bg="#0f172a", fg="#38bdf8").pack(pady=(10,0))
        tk.Label(root, text="GST Reconciler Pro - Optional Matching", font=("Segoe UI", 11, "bold"), bg="#0f172a", fg="white").pack()
        tk.Label(root, text="Developed by Kaushal Verma", font=("Segoe UI", 8), bg="#0f172a", fg="#94a3b8").pack(pady=(0,5))

        file_frame = tk.Frame(root, bg="#1e293b", padx=15, pady=8)
        file_frame.pack(padx=12, pady=5, fill="x")

        self.gstr_path = tk.StringVar()
        self.tally_path = tk.StringVar()
        self.g_header_var = tk.IntVar(value=3)
        self.t_header_var = tk.IntVar(value=3)

        tk.Label(file_frame, text="GSTR File:", bg="#1e293b", fg="white", font=("Segoe UI",9,"bold")).grid(row=0,column=0,sticky="w")
        tk.Entry(file_frame, textvariable=self.gstr_path, width=50).grid(row=0,column=1,padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_gstr, bg="#38bdf8").grid(row=0,column=2)
        tk.Label(file_frame, text="Header Row:", bg="#1e293b", fg="white").grid(row=0,column=3,padx=5)
        tk.Spinbox(file_frame, from_=1, to=10, textvariable=self.g_header_var, width=5).grid(row=0,column=4)

        tk.Label(file_frame, text="Tally File:", bg="#1e293b", fg="white", font=("Segoe UI",9,"bold")).grid(row=1,column=0,sticky="w",pady=4)
        tk.Entry(file_frame, textvariable=self.tally_path, width=50).grid(row=1,column=1,padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_tally, bg="#38bdf8").grid(row=1,column=2)
        tk.Label(file_frame, text="Header Row:", bg="#1e293b", fg="white").grid(row=1,column=3,padx=5)
        tk.Spinbox(file_frame, from_=1, to=10, textvariable=self.t_header_var, width=5).grid(row=1,column=4)

        tk.Button(file_frame, text="🔍 Load Headers & Auto-Map Columns", command=self.load_headers, bg="#f59e0b", fg="black", font=("Segoe UI",10,"bold")).grid(row=2,column=0,columnspan=5,pady=8,sticky="ew")

        # Mapping frame with scroll
        container = tk.Frame(root, bg="#0f172a")
        container.pack(padx=12, pady=2, fill="both", expand=True)

        self.map_frame = tk.Frame(container, bg="#0f172a")
        self.map_frame.pack(fill="both", expand=True)

        self.fields = [
            ('voucher', 'Voucher No./Invoice No. *', ['voucher', 'invoice'], True),
            ('gross', 'Gross Total', ['gross total', 'gross'], False),
            ('taxable', 'Taxable Value', ['taxable'], False),
            ('igst', 'IGST', ['igst'], False),
            ('cgst', 'CGST', ['cgst'], False),
            ('sgst', 'SGST', ['sgst'], False),
            ('remark', 'Remark_1 Column *', ['remark'], True),
        ]

        # Use fields dict
        self.use_vars = {}  # checkbox whether to use for matching

        self.gstr_vars = {}
        self.tally_vars = {}
        self.gstr_headers = []
        self.tally_headers = []

        # Preset buttons
        preset_frame = tk.Frame(root, bg="#0f172a")
        preset_frame.pack(pady=2)
        tk.Label(preset_frame, text="Quick Presets:", bg="#0f172a", fg="#94a3b8", font=("Segoe UI",8)).pack(side="left", padx=5)
        tk.Button(preset_frame, text="Only Invoice No.", command=lambda: self.apply_preset(['voucher']), bg="#334155", fg="white", font=("Segoe UI",8)).pack(side="left", padx=3)
        tk.Button(preset_frame, text="Invoice + Total (2 cols)", command=lambda: self.apply_preset(['voucher','gross']), bg="#334155", fg="white", font=("Segoe UI",8)).pack(side="left", padx=3)
        tk.Button(preset_frame, text="Invoice + Taxable + GST (5 cols)", command=lambda: self.apply_preset(['voucher','taxable','igst','cgst','sgst']), bg="#334155", fg="white", font=("Segoe UI",8)).pack(side="left", padx=3)
        tk.Button(preset_frame, text="All 6 cols (Full)", command=lambda: self.apply_preset(['voucher','gross','taxable','igst','cgst','sgst']), bg="#334155", fg="white", font=("Segoe UI",8)).pack(side="left", padx=3)

        self.status = tk.Label(root, text="Step 1: Select Both Files > Step 2: Load Headers > Step 3: Tick Columns For Matching Base", bg="#0f172a", fg="#e2e8f0", font=("Segoe UI",9))
        self.status.pack(pady=3)

        self.reconcile_btn = tk.Button(root, text="⚡ RECONCILE & EXPORT", command=self.reconcile, bg="#22c55e", fg="white", font=("Segoe UI", 11, "bold"), padx=20, pady=6, state="disabled")
        self.reconcile_btn.pack(pady=3)

        self.result_box = tk.Text(root, height=7, bg="#020617", fg="#22c55e", font=("Consolas",9))
        self.result_box.pack(padx=12, pady=3, fill="x")

    def apply_preset(self, keys):
        for k in self.use_vars:
            if k in ['remark']: continue
            self.use_vars[k].set(1 if k in keys else 0)
        self.status.config(text=f"Preset applied: {', '.join(keys)} will be matched with ")

    def browse_gstr(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f:
            self.gstr_path.set(f)
            self.g_header_var.set(find_header_row_guess(f))

    def browse_tally(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f:
            self.tally_path.set(f)
            self.t_header_var.set(find_header_row_guess(f))

    def load_headers(self):
        if not self.gstr_path.get() or not self.tally_path.get():
            messagebox.showerror("Error","Select the both files first")
            return
        try:
            g_h = self.g_header_var.get()
            t_h = self.t_header_var.get()
            self.gstr_headers = get_sheet_headers(self.gstr_path.get(), g_h)
            self.tally_headers = get_sheet_headers(self.tally_path.get(), t_h)

            for w in self.map_frame.winfo_children():
                w.destroy()

            # Headers
            tk.Label(self.map_frame, text="✓ Use?", bg="#0f172a", fg="#38bdf8", font=("Segoe UI",9,"bold")).grid(row=0,column=0,padx=5)
            tk.Label(self.map_frame, text="Field", bg="#0f172a", fg="#38bdf8", font=("Segoe UI",9,"bold")).grid(row=0,column=1,padx=5)
            tk.Label(self.map_frame, text="GSTR Column", bg="#0f172a", fg="#38bdf8", font=("Segoe UI",9,"bold")).grid(row=0,column=2,padx=5)
            tk.Label(self.map_frame, text="Tally Column", bg="#0f172a", fg="#38bdf8", font=("Segoe UI",9,"bold")).grid(row=0,column=3,padx=5)

            self.gstr_vars = {}
            self.tally_vars = {}
            self.use_vars = {}

            for idx, (key, label, kws, mandatory) in enumerate(self.fields, start=1):
                # Checkbox for use in matching (except remark)
                if key != 'remark':
                    use_var = tk.IntVar(value=1 if key in ['voucher','gross'] else 0)  # default: voucher+gross
                    if key == 'voucher':
                        use_var.set(1)  # voucher always default checked
                    # Mandatory fields like voucher should be always used? But make optional as per user request, but keep checked
                    chk = tk.Checkbutton(self.map_frame, variable=use_var, bg="#0f172a", activebackground="#0f172a")
                    chk.grid(row=idx,column=0)
                    self.use_vars[key] = use_var
                else:
                    tk.Label(self.map_frame, text="*", bg="#0f172a", fg="#f87171").grid(row=idx,column=0)
                    # remark not used for matching
                tk.Label(self.map_frame, text=label, bg="#0f172a", fg="white" if not mandatory else "#fbbf24", font=("Segoe UI",9)).grid(row=idx,column=1,sticky="w",padx=5,pady=2)

                auto_g = auto_map_column(self.gstr_headers, kws)
                var_g = tk.StringVar(value=auto_g)
                self.gstr_vars[key] = var_g
                opt_g = tk.OptionMenu(self.map_frame, var_g, *self.gstr_headers)
                opt_g.config(width=22, bg="#1e293b", fg="white", font=("Segoe UI",8))
                opt_g.grid(row=idx,column=2,padx=3,pady=2)

                auto_t = auto_map_column(self.tally_headers, kws)
                var_t = tk.StringVar(value=auto_t)
                self.tally_vars[key] = var_t
                opt_t = tk.OptionMenu(self.map_frame, var_t, *self.tally_headers)
                opt_t.config(width=22, bg="#1e293b", fg="white", font=("Segoe UI",8))
                opt_t.grid(row=idx,column=3,padx=3,pady=2)

            self.status.config(text=f"Loaded! Now tick the options - Like Invoice + Gross Total for matching.")
            self.reconcile_btn.config(state="normal")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status.config(text=f"Error: {e}")

    def reconcile(self):
        try:
            # Check at least one field selected for matching
            selected = [k for k,v in self.use_vars.items() if v.get()==1]
            if not selected:
                messagebox.showerror("Error","Minimum 1 column tick for the matching (Voucher No. recommended)")
                return
            self.status.config(text=f"Matching on: {', '.join(selected)}...")
            self.root.update()
            out_dir = os.path.join(os.path.expanduser("~"), "Desktop", "SJGroup_Output")

            gstr_map = {k: v.get() for k,v in self.gstr_vars.items()}
            tally_map = {k: v.get() for k,v in self.tally_vars.items()}
            use_fields = {k: bool(v.get()) for k,v in self.use_vars.items()}

            out_g, out_t, mg, mt, tg, tt = process_with_mapping(
                self.gstr_path.get(),
                self.tally_path.get(),
                gstr_map,
                tally_map,
                self.g_header_var.get(),
                self.t_header_var.get(),
                use_fields,
                out_dir
            )
            msg = f"✅ SUCCESS - Matched on: {', '.join(selected)}\nGSTR: {mg}/{tg} Matched\nTally: {mt}/{tt} Matched\n\nFiles:\n{out_g}\n{out_t}"
            self.result_box.delete(1.0, tk.END)
            self.result_box.insert(tk.END, msg)
            self.status.config(text=f"Matching Done ! {len(selected)} cols matching complete")
            messagebox.showinfo("Done", msg)
            try:
                if platform.system() == "Windows":
                    os.startfile(out_dir)
                else:
                    os.system(f'xdg-open "{out_dir}"')
            except: pass
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status.config(text=f"Error: {e}")
            import traceback
            self.result_box.delete(1.0, tk.END)
            self.result_box.insert(tk.END, traceback.format_exc())

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
