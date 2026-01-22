import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import openpyxl
from datetime import datetime, timedelta
import hashlib, json, base64
from pathlib import Path
import os

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
        self.dialog.title("Register Software")
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
        
        ttk.Label(frame, text="Enter your registration details:", font=("Arial", 10, "bold")).pack(pady=(0,10))
        
        ttk.Label(frame, text="Email:").pack(anchor="w")
        self.email_entry = ttk.Entry(frame, width=40)
        self.email_entry.pack(pady=(0,10), fill="x")
        
        ttk.Label(frame, text="License Key:").pack(anchor="w")
        self.key_entry = ttk.Entry(frame, width=40)
        self.key_entry.pack(pady=(0,15), fill="x")
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Register", command=self.register).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side="left", padx=5)
        
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
        
    def register(self):
        email = self.email_entry.get().strip()
        key = self.key_entry.get().strip()
        
        if not email or not key:
            messagebox.showwarning("Invalid Input", "Please enter both email and license key.", parent=self.dialog)
            return
        
        expiry_date = self.license_mgr.validate_key(email, key)
        if expiry_date:
            self.license_mgr.save_license(email, expiry_date)
            messagebox.showinfo("Success", f"Registration successful!\nLicense valid until: {expiry_date}", parent=self.dialog)
            self.result = True
            self.dialog.destroy()
        else:
            messagebox.showerror("Invalid License", "The license key is invalid or does not match the email.", parent=self.dialog)
    
    def cancel(self):
        self.result = False
        self.dialog.destroy()

class LabelPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Label Printer - Avery 3658")
        self.root.geometry("700x850")
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
        self.register_fonts_for_pdf()

        # Avery 3658 specs
        self.label_width = 64.6 * mm
        self.label_height = 33.8 * mm
        self.cols, self.rows = 3, 8
        self.col_gap, self.row_gap = 2.5 * mm, 0 * mm
        self.left_margin, self.top_margin = 4.7 * mm, 13.5 * mm
        
        # Universal horizontal padding for all labels
        self.universal_h_padding = tk.DoubleVar(value=2)  # mm from left/right edges
        
        # Additional column-specific padding adjustments
        self.left_col_extra_padding = tk.DoubleVar(value=0)   # Extra left padding for left column
        self.right_col_extra_padding = tk.DoubleVar(value=0)  # Extra right padding for right column

        self.create_widgets()

    def register_fonts_for_pdf(self):
        """Try to register available TTF fonts for PDF use (regular and bold)"""
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
            "Registration Required", 
            "This software requires a valid license to run.\n\nPlease enter your registration details."
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

        ttk.Label(main, text="Label Printer - Avery Zweckform 3658", font=("Arial", 14, "bold")).pack(pady=10)

        # 🧭 Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import from Excel", command=self.import_from_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Register Software", command=self.show_registration_dialog)
        help_menu.add_command(label="Check License", command=self.show_license_info)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

        # 📋 Data input
        input_frame = ttk.LabelFrame(main, text="Paste Data (one line per label line)", padding=10)
        input_frame.pack(fill="both", expand=True, pady=5)
        self.text_input = scrolledtext.ScrolledText(input_frame, width=70, height=10)
        self.text_input.pack(fill="both", expand=True)
        self.text_input.bind("<<Modified>>", lambda e: (self.update_preview(), self.text_input.edit_modified(False)))
        
        # Add right-click context menu
        self.context_menu = tk.Menu(self.text_input, tearoff=0)
        self.context_menu.add_command(label="Cut", command=lambda: self.text_input.event_generate("<<Cut>>"))
        self.context_menu.add_command(label="Copy", command=lambda: self.text_input.event_generate("<<Copy>>"))
        self.context_menu.add_command(label="Paste", command=lambda: self.text_input.event_generate("<<Paste>>"))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Select All", command=lambda: self.text_input.tag_add("sel", "1.0", "end"))
        
        def show_context_menu(event):
            self.context_menu.tk_popup(event.x_root, event.y_root)
        
        self.text_input.bind("<Button-3>", show_context_menu)  # Right-click on Windows/Linux
        self.text_input.bind("<Button-2>", show_context_menu)  # Right-click on Mac

        # ⚙ Settings
        settings = ttk.LabelFrame(main, text="Label Settings", padding=10)
        settings.pack(fill="x", pady=5)

        ttk.Label(settings, text="Lines per label:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.lines_var = tk.IntVar(value=3)
        ttk.Combobox(settings, textvariable=self.lines_var, values=[1,2,3,4,5,6], width=5, state='readonly').grid(row=0, column=1, sticky="w")
        self.lines_var.trace_add("write", self.update_preview)

        ttk.Label(settings, text="Font:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.font_var = tk.StringVar(value="Arial")
        font_combo = ttk.Combobox(settings, textvariable=self.font_var, values=self.available_fonts, width=25, state='readonly')
        font_combo.grid(row=1, column=1, padx=5, sticky="w")
        font_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        self.bold_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings, text="Bold", variable=self.bold_var, command=self.update_preview).grid(row=1, column=2, padx=10)

        # Universal padding
        ttk.Label(settings, text="Universal H padding (mm):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        universal_padding_spinbox = ttk.Spinbox(settings, from_=0, to=15.0, increment=0.5, textvariable=self.universal_h_padding, width=8)
        universal_padding_spinbox.grid(row=2, column=1, padx=5, sticky="w")
        ttk.Label(settings, text="(applies to all labels)").grid(row=2, column=2, padx=5, sticky="w")

        # Additional column-specific adjustments
        ttk.Label(settings, text="Left col extra padding (mm):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        left_extra_spinbox = ttk.Spinbox(settings, from_=0, to=10.0, increment=0.5, textvariable=self.left_col_extra_padding, width=8)
        left_extra_spinbox.grid(row=3, column=1, padx=5, sticky="w")
        ttk.Label(settings, text="(adds to left edge only)").grid(row=3, column=2, padx=5, sticky="w")

        ttk.Label(settings, text="Right col extra padding (mm):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        right_extra_spinbox = ttk.Spinbox(settings, from_=0, to=10.0, increment=0.5, textvariable=self.right_col_extra_padding, width=8)
        right_extra_spinbox.grid(row=4, column=1, padx=5, sticky="w")
        ttk.Label(settings, text="(adds to right edge only)").grid(row=4, column=2, padx=5, sticky="w")

        # 👀 Preview
        preview_frame = ttk.LabelFrame(main, text="Preview", padding=10)
        preview_frame.pack(fill="x", pady=5)
        self.preview_canvas = tk.Canvas(preview_frame, width=500, height=120, bg="white", relief="sunken", bd=1)
        self.preview_canvas.pack(padx=10, pady=5)
        self.update_preview()

        # 🧾 Buttons
        button_frame = ttk.Frame(main)
        button_frame.pack(pady=15)
        
        ttk.Button(button_frame, text="📁 Import from Excel", command=self.import_from_excel).pack(side="left", padx=5)
        ttk.Button(button_frame, text="📄 Generate PDF Labels", command=self.generate_labels).pack(side="left", padx=5)

    def show_registration_dialog(self):
        dialog = RegistrationDialog(self.root, self.license_mgr)
        self.root.wait_window(dialog.dialog)

    def show_license_info(self):
        data = self.license_mgr.load_license()
        if data:
            days = self.license_mgr.get_days_remaining()
            messagebox.showinfo(
                "License Info",
                f"Licensed Email: {data['email']}\n"
                f"Expires on: {data['expiry']}\n"
                f"Days remaining: {days}"
            )
        else:
            messagebox.showwarning("License Info", "No valid license found.")

    def show_about(self):
        messagebox.showinfo(
            "About",
            "Label Printer - Avery Zweckform 3658\n"
            "Version 1.4\n\n"
            "Developed by Blaž Pivk\n"
            "© 2025"
        )

    def import_from_excel(self):
        """Import data from Excel file (first column only)"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
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
                messagebox.showwarning("No Data", "No data found in the first column of the Excel file.")
                return
            
            # Clear current text and insert imported data
            self.text_input.delete("1.0", "end")
            self.text_input.insert("1.0", "\n".join(data))
            self.update_preview()
            
            messagebox.showinfo("Success", f"Imported {len(data)} lines from Excel file.")
            
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import Excel file:\n\n{str(e)}")

    def update_preview(self, *_):
        """Show the first label's text with selected font"""
        self.preview_canvas.delete("all")
        data = self.text_input.get("1.0", "end").strip().split("\n")
        lines_per_label = self.lines_var.get()
        if data and data[0]:
            text_lines = data[:lines_per_label]
        else:
            text_lines = ["Sample Label Text"]
        text = "\n".join(text_lines)
        font = (self.font_var.get(), 14, "bold" if self.bold_var.get() else "normal")
        self.preview_canvas.create_text(250, 60, text=text, font=font, justify="center", fill="black")

    def generate_labels(self):
        text = self.text_input.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("No data", "Enter or paste label data first, or import from Excel.")
            return
        lines = [l.rstrip() for l in text.split("\n")]  # Keep empty lines, only strip trailing whitespace
        n = self.lines_var.get()
        labels = [lines[i:i+n] for i in range(0, len(lines), n)]
        if not labels:
            messagebox.showwarning("No Data", "No valid data to print")
            return
        filename = filedialog.asksaveasfilename(defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")], initialfile=f"labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        if not filename: return
        try:
            self.create_pdf(filename, labels, n)
            messagebox.showinfo("Success", f"PDF saved: {filename}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def create_pdf(self, filename, labels, lines_per_label):
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
        
        # Calculate font size based on the safe area (so text respects padding)
        size = self.calculate_font_size(lines, safe_width, self.label_height, max_lines, actual_font_name)
        line_h = size * 1.2
        total_h = len(lines) * line_h
        
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
    LabelPrinterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
