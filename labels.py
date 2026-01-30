import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime, timedelta
import hashlib, json, base64
from pathlib import Path
import os

# Heavy imports - lazy load only when needed
# from reportlab.lib.units import mm
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# from reportlab.pdfbase import pdfmetrics
# from reportlab.pdfbase.ttfonts import TTFont
# import openpyxl

class LicenseManager:
    def __init__(self):
        self.license_file = Path.home() / '.labelprinterlicense.dat'
        self.secret_salt = "LabelPrinter2025SecretKey"

    def generate_key(self, email, expiry_date):
        data = f"{email}|{expiry_date}|{self.secret_salt}"
        key = hashlib.sha256(data.encode()).hexdigest()[:24].upper()
        return '-'.join([key[i:i+4] for i in range(0, 24, 4)])

    def validate_key(self, email, key):
        for days in range(0, 730):
            test_date = (datetime.now() + timedelta(days=days)).strftime('%Y-%m-%d')
            if self.generate_key(email, test_date) == key.upper().strip():
                return test_date
        return None

    def save_license(self, email, expiry_date):
        data = {'email': email, 'expiry': expiry_date}
        encoded = base64.b64encode(json.dumps(data).encode()).decode()
        self.license_file.write_text(encoded)

    def load_license(self):
        if not self.license_file.exists(): return None
        try:
            data = json.loads(base64.b64decode(self.license_file.read_text()).decode())
            expiry = datetime.strptime(data['expiry'], '%Y-%m-%d')
            return data if expiry >= datetime.now() else None
        except:
            return None

    def get_days_remaining(self):
        d = self.load_license()
        return max(0, (datetime.strptime(d['expiry'], '%Y-%m-%d') - datetime.now()).days) if d else 0

class RegistrationDialog:
    def __init__(self, parent, license_mgr):
        self.result = None
        self.license_mgr = license_mgr
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Registracija Programa")
        self.dialog.geometry("450x250")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (250 // 2)
        self.dialog.geometry(f'450x250+{x}+{y}')
        
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill="both", expand=True)
        
        ttk.Label(frame, text="Vnesite registracijske podatke:", font=("Arial", 10, "bold")).pack(pady=(0,10))
        
        ttk.Label(frame, text="E-pošta:").pack(anchor="w")
        self.email_entry = ttk.Entry(frame, width=40)
        self.email_entry.pack(pady=(0,10), fill="x")
        
        ttk.Label(frame, text="Licenčna Ključ:").pack(anchor="w")
        self.key_entry = ttk.Entry(frame, width=40)
        self.key_entry.pack(pady=(0,15), fill="x")
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Registriraj", command=self.register).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Prekliči", command=self.cancel).pack(side="left", padx=5)
        
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
    def register(self):
        email = self.email_entry.get().strip()
        key = self.key_entry.get().strip()
        
        if not email or not key:
            messagebox.showwarning("Napačen Vnos", "Prosim vnesite e-pošto in licenčni ključ.", parent=self.dialog)
            return
        
        expiry_date = self.license_mgr.validate_key(email, key)
        if expiry_date:
            self.license_mgr.save_license(email, expiry_date)
            messagebox.showinfo("Uspeh", f"Registracija uspešna!\nLicenca veljavna do: {expiry_date}", parent=self.dialog)
            self.result = True
            self.dialog.destroy()
        else:
            messagebox.showerror("Neveljavna Licenca", "Licenčni ključ je neveljaven ali se ne ujema z e-pošto.", parent=self.dialog)
    
    def cancel(self):
        self.result = False
        self.dialog.destroy()

class LabelPrinterApp:
    def __init__(self, root):
        # Import reportlab units here (lazy load)
        from reportlab.lib.units import mm
        global mm  # Make it available to other methods
        
        self.root = root
        self.root.title("Tiskalnik Nalepk - Avery 3658")
        self.root.geometry("750x800")
        self.license_mgr = LicenseManager()
        if not self.check_license():
            root.destroy()
            return

        # 🎨 Available fonts
        self.available_fonts = ["Arial", "Arial Narrow", "Helvetica", "Times New Roman", "Courier New"]

        # 🗂 Font file map for PDF (regular and bold)
        self.font_map = {
            "Arial": [["arial.ttf", "ARIAL.TTF"], ["arialbd.ttf", "ARIALBD.TTF"]],
            "Arial Narrow": [["arialn.ttf", "ARIALN.TTF"], ["arialnb.ttf", "ARIALNB.TTF"]],
            "Helvetica": [["arial.ttf", "ARIAL.TTF"], ["arialbd.ttf", "ARIALBD.TTF"]],
            "Times New Roman": [["times.ttf", "TIMES.TTF"], ["timesbd.ttf", "TIMESBD.TTF"]],
            "Courier New": [["cour.ttf", "COUR.TTF"], ["courbd.ttf", "COURBD.TTF"]]
        }

        # Register fonts for PDF if found
        self._fonts_registered = False  # Lazy load fonts only when needed

        # Avery 3658 specs
        self.label_width = 64.6 * mm
        self.label_height = 33.8 * mm
        self.cols, self.rows = 3, 8
        self.col_gap, self.row_gap = 0 * mm, 0 * mm  # Labels touch - no physical gap
        # Calculate left margin: (A4 width - 3 labels) / 2 = (210mm - 3*64.6mm) / 2 = 8.1mm
        self.left_margin = (210 * mm - 3 * self.label_width) / 2
        self.top_margin = 13.5 * mm
        
        # Font settings
        self.lines_var = tk.IntVar(value=3)
        self.font_var = tk.StringVar(value="Arial")
        self.bold_var = tk.BooleanVar(value=False)
        self.font_size_var = tk.IntVar(value=18)
        
        # Universal horizontal padding for all labels
        self.universal_h_padding = tk.DoubleVar(value=2)  # mm from left/right edges
        
        # Additional column-specific padding adjustments
        self.left_col_extra_padding = tk.DoubleVar(value=0)   # Extra left padding for left column
        self.right_col_extra_padding = tk.DoubleVar(value=0)  # Extra right padding for right column

        self.create_widgets()

    def register_fonts_for_pdf(self):
        """Try to register available TTF fonts for PDF use (regular and bold)"""
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        
        search_paths = [
            "C:\\Windows\\Fonts",
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            str(Path.home() / ".fonts")
        ]
        for name, font_variants in self.font_map.items():
            # Register regular font
            for sp in search_paths:
                for f in font_variants[0]:
                    path = Path(sp) / f
                    if path.exists():
                        try:
                            pdfmetrics.registerFont(TTFont(name, str(path)))
                        except:
                            pass
                        break
            # Register bold font
            for sp in search_paths:
                for f in font_variants[1]:
                    path = Path(sp) / f
                    if path.exists():
                        try:
                            pdfmetrics.registerFont(TTFont(f"{name}-Bold", str(path)))
                        except:
                            pass
                        break

    def check_license(self):
        license_data = self.license_mgr.load_license()
        if license_data:
            return True
        
        # No valid license - show registration dialog
        messagebox.showwarning(
            "Potrebna Registracija", 
            "Ta program zahteva veljavno licenco.\n\nProsim vnesite registracijske podatke."
        )
        
        dialog = RegistrationDialog(self.root, self.license_mgr)
        self.root.wait_window(dialog.dialog)
        
        # Check if registration was successful
        if self.license_mgr.load_license():
            return True
        else:
            messagebox.showerror("Registration Required", "Software cannot run without a valid license.")
            return False

    def create_widgets(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="Tiskalnik Nalepk - Avery Zweckform 3658", font=("Arial", 14, "bold")).pack(pady=10)

        # 🧭 Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Datoteka", menu=file_menu)
        file_menu.add_command(label="Uvozi iz Excela", command=self.import_from_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Izhod", command=self.root.quit)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Pomoč", menu=help_menu)
        help_menu.add_command(label="Registracija Programa", command=self.show_registration_dialog)
        help_menu.add_command(label="Preveri Licenco", command=self.show_license_info)
        help_menu.add_separator()
        help_menu.add_command(label="O programu", command=self.show_about)

        # 📋 Data input
        input_frame = ttk.LabelFrame(main, text="Vnesi Podatke (ena vrstica na vrstico nalepke)", padding=10)
        input_frame.pack(fill="both", expand=True, pady=5)
        self.text_input = scrolledtext.ScrolledText(input_frame, width=70, height=10)
        self.text_input.pack(fill="both", expand=True)
        self.text_input.bind("<<Modified>>", lambda e: (self.update_preview(), self.text_input.edit_modified(False)))
        
        # Add right-click context menu
        self.context_menu = tk.Menu(self.text_input, tearoff=0)
        self.context_menu.add_command(label="Izreži", command=lambda: self.text_input.event_generate("<<Cut>>"))
        self.context_menu.add_command(label="Kopiraj", command=lambda: self.text_input.event_generate("<<Copy>>"))
        self.context_menu.add_command(label="Prilepi", command=lambda: self.text_input.event_generate("<<Paste>>"))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Izberi Vse", command=lambda: self.text_input.tag_add("sel", "1.0", "end"))
        
        def show_context_menu(event):
            self.context_menu.tk_popup(event.x_root, event.y_root)
        
        self.text_input.bind("<Button-3>", show_context_menu)  # Right-click on Windows/Linux
        self.text_input.bind("<Button-2>", show_context_menu)  # Right-click on Mac

        # ⚙ Font Settings
        font_settings = ttk.LabelFrame(main, text="Nastavitve Pisave", padding=10)
        font_settings.pack(fill="x", pady=5)

        # Row 0: Lines per label and Font
        ttk.Label(font_settings, text="Vrstic na nalepko:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Combobox(font_settings, textvariable=self.lines_var, values=[1,2,3,4,5,6], width=5, state='readonly').grid(row=0, column=1, padx=5, sticky="w")
        self.lines_var.trace_add("write", self.update_preview)
        
        ttk.Label(font_settings, text="Pisava:").grid(row=0, column=2, padx=15, pady=5, sticky="w")
        font_combo = ttk.Combobox(font_settings, textvariable=self.font_var, values=self.available_fonts, width=20, state='readonly')
        font_combo.grid(row=0, column=3, padx=5, sticky="w")
        font_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        # Row 1: Font size and Bold
        ttk.Label(font_settings, text="Velikost pisave (pt):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        font_size_spinbox = ttk.Spinbox(font_settings, from_=6, to=72, increment=1, textvariable=self.font_size_var, width=5, command=self.update_preview)
        font_size_spinbox.grid(row=1, column=1, padx=5, sticky="w")
        
        ttk.Checkbutton(font_settings, text="Krepko", variable=self.bold_var, command=self.update_preview).grid(row=1, column=2, padx=15, sticky="w")

        # 📐 Padding Settings
        padding_settings = ttk.LabelFrame(main, text="Nastavitve Odmikov (mm)", padding=10)
        padding_settings.pack(fill="x", pady=5)

        # Row 0: Universal padding
        ttk.Label(padding_settings, text="Univerzalni:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Spinbox(padding_settings, from_=0, to=15.0, increment=0.5, textvariable=self.universal_h_padding, width=8, command=self.update_preview).grid(row=0, column=1, padx=5, sticky="w")
        ttk.Label(padding_settings, text="(vse nalepke)").grid(row=0, column=2, padx=5, sticky="w")
        
        # Row 1: Left and Right column adjustments
        ttk.Label(padding_settings, text="Levi stolpec dodatno:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Spinbox(padding_settings, from_=0, to=10.0, increment=0.5, textvariable=self.left_col_extra_padding, width=8, command=self.update_preview).grid(row=1, column=1, padx=5, sticky="w")
        
        ttk.Label(padding_settings, text="Desni stolpec dodatno:").grid(row=1, column=3, padx=15, pady=5, sticky="w")
        ttk.Spinbox(padding_settings, from_=0, to=10.0, increment=0.5, textvariable=self.right_col_extra_padding, width=8, command=self.update_preview).grid(row=1, column=4, padx=5, sticky="w")

        # 👀 Preview
        preview_frame = ttk.LabelFrame(main, text="Predogled Nalepk (3 nalepke prikazane)", padding=10)
        preview_frame.pack(fill="x", pady=5)
        self.preview_canvas = tk.Canvas(preview_frame, width=700, height=140, bg="white", relief="sunken", bd=1)
        self.preview_canvas.pack(padx=10, pady=5)
        self.update_preview()

        # 🧾 Buttons
        button_frame = ttk.Frame(main)
        button_frame.pack(pady=15)
        
        ttk.Button(button_frame, text="📁 Uvozi iz Excela", command=self.import_from_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="📄 Generiraj PDF", command=self.generate_labels).pack(side="left", padx=5)
        ttk.Button(button_frame, text="🖨 Natisni", command=self.print_labels).pack(side="left", padx=5)

    def print_labels(self):
        """Generate PDF and show printer selection dialog"""
        text = self.text_input.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Ni podatkov", "Najprej vnesite ali prilepite podatke za nalepke, ali uvozite iz Excela.")
            return
        
        lines = [l.rstrip() for l in text.split("\n")]
        n = self.lines_var.get()
        labels = [lines[i:i+n] for i in range(0, len(lines), n)]
        
        if not labels:
            messagebox.showwarning("Ni Podatkov", "Ni veljavnih podatkov za tiskanje")
            return
        
        # Show printer selection dialog
        import platform
        
        if platform.system() == 'Windows':
            try:
                self.show_windows_printer_dialog(labels, n)
            except ImportError:
                messagebox.showerror("Napaka", "Potrebna je knjižnica pywin32.\nNamestite z: pip install pywin32")
        else:
            # For non-Windows, use system print command
            self.print_with_system_dialog(labels, n)
    
    def show_windows_printer_dialog(self, labels, lines_per_label):
        """Show custom printer selection dialog for Windows"""
        import win32print
        
        # Get list of printers
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        default_printer = win32print.GetDefaultPrinter()
        
        # Create dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Izberi Tiskalnik")
        dialog.geometry("450x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - 225
        y = (dialog.winfo_screenheight() // 2) - 125
        dialog.geometry(f'450x250+{x}+{y}')
        
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill="both", expand=True)
        
        ttk.Label(frame, text="Izberite tiskalnik:", font=("Arial", 10, "bold")).pack(pady=(0, 10))
        
        # Printer listbox
        listbox_frame = ttk.Frame(frame)
        listbox_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side="right", fill="y")
        
        printer_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set, height=8)
        printer_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=printer_listbox.yview)
        
        # Populate printer list
        for i, printer in enumerate(printers):
            printer_listbox.insert(tk.END, printer)
            if printer == default_printer:
                printer_listbox.selection_set(i)
                printer_listbox.see(i)
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        
        def do_print():
            selection = printer_listbox.curselection()
            if not selection:
                messagebox.showwarning("Izberi Tiskalnik", "Prosim izberite tiskalnik.", parent=dialog)
                return
            
            selected_printer = printers[selection[0]]
            dialog.destroy()
            
            # Generate PDF and print
            import tempfile
            import os
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_filename = temp_pdf.name
            temp_pdf.close()
            
            try:
                self.create_pdf(temp_filename, labels, lines_per_label)
                
                # Print to selected printer
                win32print.SetDefaultPrinter(selected_printer)
                os.startfile(temp_filename, 'print')
                
                # Restore original default printer after a short delay
                self.root.after(2000, lambda: win32print.SetDefaultPrinter(default_printer) if default_printer != selected_printer else None)
                
                messagebox.showinfo("Tiskanje", f"Dokument poslan na tiskalnik:\n{selected_printer}")
            except Exception as e:
                messagebox.showerror("Napaka pri Tiskanju", f"Napaka:\n\n{str(e)}")
        
        ttk.Button(btn_frame, text="Natisni", command=do_print).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Prekliči", command=dialog.destroy).pack(side="left", padx=5)
        
        # Bind double-click to print
        printer_listbox.bind('<Double-Button-1>', lambda e: do_print())
    
    def print_with_system_dialog(self, labels, lines_per_label):
        """Fallback for non-Windows systems"""
        import tempfile
        import subprocess
        
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_filename = temp_pdf.name
        temp_pdf.close()
        
        try:
            self.create_pdf(temp_filename, labels, lines_per_label)
            subprocess.run(['lp', temp_filename])
            messagebox.showinfo("Tiskanje", "Dokument poslan na tiskalnik.")
        except Exception as e:
            messagebox.showerror("Napaka pri Tiskanju", f"Napaka:\n\n{str(e)}")

    def show_registration_dialog(self):
        dialog = RegistrationDialog(self.root, self.license_mgr)
        self.root.wait_window(dialog.dialog)

    def show_license_info(self):
        data = self.license_mgr.load_license()
        if data:
            days = self.license_mgr.get_days_remaining()
            messagebox.showinfo(
                "Informacije o Licenci",
                f"Licenčna E-pošta: {data['email']}\n"
                f"Poteče: {data['expiry']}\n"
                f"Dni do poteka: {days}"
            )
        else:
            messagebox.showwarning("Informacije o Licenci", "Veljavna licenca ni najdena.")

    def show_about(self):
        messagebox.showinfo(
            "O Programu",
            "Tiskalnik Nalepk - Avery Zweckform 3658\n"
            "Verzija 2.0\n\n"
            "Razvil Blaž Pivk\n"
            "© 2026"
        )

    def import_from_excel(self):
        """Import data from Excel file (first column only)"""
        # Lazy import openpyxl only when needed
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("Manjkajoča Knjižnica", "openpyxl ni nameščen. Namestite z: pip install openpyxl")
            return
        
        filename = filedialog.askopenfilename(
            title="Izberi Excel Datoteko",
            filetypes=[("Excel Datoteke", "*.xlsx *.xls"), ("Vse Datoteke", "*.*")]
        )
        
        if not filename:
            return
        
        try:
            workbook = openpyxl.load_workbook(filename, data_only=True)
            sheet = workbook.active
            
            # Extract data from first column (skip empty cells)
            data = []
            for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
                cell_value = row[0].value
                if cell_value is not None:
                    # Convert to string and strip whitespace
                    data.append(str(cell_value).strip())
            
            if not data:
                messagebox.showwarning("Ni Podatkov", "V prvem stolpcu Excel datoteke ni podatkov.")
                return
            
            # Clear current text and insert imported data
            self.text_input.delete("1.0", "end")
            self.text_input.insert("1.0", "\n".join(data))
            self.update_preview()
            
            messagebox.showinfo("Uspeh", f"Uvoženih {len(data)} vrstic iz Excel datoteke.")
            
        except Exception as e:
            messagebox.showerror("Napaka pri Uvozu", f"Napaka pri uvozu Excel datoteke:\n\n{str(e)}")

    def update_preview(self, *_):
        """Show 3 labels (left, center, right) with proper scaling"""
        self.preview_canvas.delete("all")
        data = self.text_input.get("1.0", "end").strip().split("\n")
        lines_per_label = self.lines_var.get()
        
        if data and data[0]:
            text_lines = data[:lines_per_label]
        else:
            text_lines = ["Vzorčna Nalepka", "Besedilo Tukaj", "Vrstica 3"][:lines_per_label]
        
        # Calculate scale to fit labels in preview (64.6mm label width)
        scale = 3.0  # pixels per mm
        label_w = 64.6 * scale
        label_h = 33.8 * scale
        gap = 0 * scale  # Labels touch - no physical gap (Word template lines are just visual guides)
        
        # Calculate padding
        base_pad = self.universal_h_padding.get() * scale
        left_extra = self.left_col_extra_padding.get() * scale
        right_extra = self.right_col_extra_padding.get() * scale
        
        font_name = self.font_var.get()
        font_size = self.font_size_var.get()
        font_weight = "bold" if self.bold_var.get() else "normal"
        
        # Draw 3 labels
        start_x = 20
        start_y = 10
        
        for col in range(3):
            x = start_x + col * (label_w + gap)
            
            # Draw label border
            self.preview_canvas.create_rectangle(
                x, start_y, x + label_w, start_y + label_h,
                outline="lightgray", width=2
            )
            
            # Calculate padding for this column
            if col == 0:  # Left
                left_pad = base_pad + left_extra
                right_pad = base_pad
                label_text = "LEVO"
            elif col == 1:  # Center
                left_pad = base_pad
                right_pad = base_pad
                label_text = "SREDINA"
            else:  # Right
                left_pad = base_pad
                right_pad = base_pad + right_extra
                label_text = "DESNO"
            
            # Draw padding guides (light red lines)
            if left_pad > 0:
                self.preview_canvas.create_line(
                    x + left_pad, start_y, x + left_pad, start_y + label_h,
                    fill="pink", dash=(2, 2)
                )
            if right_pad > 0:
                self.preview_canvas.create_line(
                    x + label_w - right_pad, start_y, x + label_w - right_pad, start_y + label_h,
                    fill="pink", dash=(2, 2)
                )
            
            # Calculate safe width
            safe_w = label_w - left_pad - right_pad
            
            # Draw text lines
            text_y = start_y + label_h / 2 - (len(text_lines) - 1) * font_size * 0.6
            for line in text_lines:
                self.preview_canvas.create_text(
                    x + left_pad + safe_w / 2, text_y,
                    text=line,
                    font=(font_name, max(6, int(font_size * 0.8)), font_weight),
                    fill="black"
                )
                text_y += font_size * 1.2 * 0.8
            
            # Label indicator
            self.preview_canvas.create_text(
                x + label_w / 2, start_y + label_h + 10,
                text=label_text,
                font=("Arial", 8),
                fill="gray"
            )

    def generate_labels(self):
        text = self.text_input.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Ni podatkov", "Najprej vnesite ali prilepite podatke za nalepke, ali uvozite iz Excela.")
            return
        lines = [l.rstrip() for l in text.split("\n")]  # Keep empty lines, only strip trailing whitespace
        n = self.lines_var.get()
        labels = [lines[i:i+n] for i in range(0, len(lines), n)]
        if not labels:
            messagebox.showwarning("Ni Podatkov", "Ni veljavnih podatkov za tiskanje")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")], initialfile=f"nalepke_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        if not filename: return
        try:
            self.create_pdf(filename, labels, n)
            messagebox.showinfo("Uspeh", f"PDF shranjen: {filename}")
        except Exception as e:
            messagebox.showerror("Napaka", str(e))

    def create_pdf(self, filename, labels, lines_per_label):
        # Lazy import reportlab only when generating PDF
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        
        # Lazy load fonts only when generating PDF (not at startup)
        if not self._fonts_registered:
            self.register_fonts_for_pdf()
            self._fonts_registered = True
        
        c = canvas.Canvas(filename, pagesize=A4)
        font_name = self.font_var.get()
        use_bold = self.bold_var.get()
        for page_start in range(0, len(labels), self.cols * self.rows):
            if page_start > 0: c.showPage()
            for idx, label_lines in enumerate(labels[page_start:page_start + self.cols*self.rows]):
                row, col = divmod(idx, self.cols)
                x = self.left_margin + col * (self.label_width + self.col_gap)
                y = A4[1] - self.top_margin - (row + 1) * self.label_height - row * self.row_gap
                self.draw_label(c, x, y, label_lines, lines_per_label, font_name, use_bold, col)
        c.save()

    def draw_label(self, c, x, y, lines, max_lines, font_name, use_bold=False, col=1):
        actual_font_name = f"{font_name}-Bold" if use_bold else font_name
        
        # Use fixed font size from GUI (no auto-resize)
        size = self.font_size_var.get()
        line_h = size * 1.2
        total_h = len(lines) * line_h
        
        # Calculate padding - this defines the "safe area" within the label
        base_padding = self.universal_h_padding.get() * mm
        left_pad = base_padding
        right_pad = base_padding
        
        # Add column-specific extra padding
        if col == 0:  # Left column
            left_pad += self.left_col_extra_padding.get() * mm
        elif col == 2:  # Right column
            right_pad += self.right_col_extra_padding.get() * mm
        
        # Calculate the safe area width after padding
        safe_width = self.label_width - left_pad - right_pad
        
        # Center text vertically
        start_y = y + (self.label_height - total_h) / 2
        
        # Draw each line, centered within the safe area
        for i, line in enumerate(lines):
            ty = start_y + (len(lines)-1-i) * line_h
            tw = c.stringWidth(line, actual_font_name, size)
            
            # Position text: start from label left edge + left padding, 
            # then center within the safe width
            tx = x + left_pad + (safe_width - tw) / 2
            
            c.setFont(actual_font_name, size)
            c.drawString(tx, ty, line)

    def calculate_font_size(self, lines, max_w, max_h, max_lines, font_name):
        base = {1:32, 2:24, 3:18, 4:14, 5:12, 6:10}
        fs = base.get(max_lines, 10)
        while fs > 6:
            if len(lines) * fs * 1.2 > max_h: fs -= 1; continue
            if any(canvas.Canvas("t").stringWidth(l, font_name, fs) > max_w for l in lines):
                fs -= 1; continue
            break
        return fs

def main():
    root = tk.Tk()
    root.withdraw()  # Hide initially
    
    # Create app
    app = LabelPrinterApp(root)
    
    # Show window after widgets are created
    root.deiconify()
    root.mainloop()

if __name__ == "__main__":
    main()
