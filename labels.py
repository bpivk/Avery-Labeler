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

class LabelPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Label Printer - Avery 3658")
        self.root.geometry("700x700")
        self.license_mgr = LicenseManager()
        if not self.check_license():
            root.destroy(); return

        # 🎨 Available fonts
        self.available_fonts = ["Arial", "Arial Narrow", "Helvetica", "Times New Roman", "Courier New"
        ]

        # 🗂 Font file map for PDF
        self.font_map = {
            "Arial": ["arial.ttf", "ARIAL.TTF"],
            "Arial Narrow": ["arialn.ttf", "ARIALN.TTF"],
            "Helvetica": ["arial.ttf", "ARIAL.TTF"],
            "Times New Roman": ["times.ttf", "TIMES.TTF"],
            "Courier New": ["cour.ttf", "COUR.TTF"]
        }

        # Register fonts for PDF if found
        self.register_fonts_for_pdf()

        # Avery 3658 specs
        self.label_width = 64.6 * mm
        self.label_height = 33.8 * mm
        self.cols, self.rows = 3, 8
        self.col_gap, self.row_gap = 2.5 * mm, 0
        self.left_margin, self.top_margin = 4.7 * mm, 13.5 * mm

        self.create_widgets()

    def register_fonts_for_pdf(self):
        """Try to register available TTF fonts for PDF use"""
        search_paths = [
            "C:\\Windows\\Fonts",
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            str(Path.home() / ".fonts")
        ]
        for name, files in self.font_map.items():
            for sp in search_paths:
                for f in files:
                    path = Path(sp) / f
                    if path.exists():
                        try:
                            pdfmetrics.registerFont(TTFont(name, str(path)))
                        except:
                            pass
                        break

    def check_license(self):
        if self.license_mgr.load_license(): return True
        messagebox.showinfo("License", "Demo mode: no license found.")
        return True

    def create_widgets(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="Label Printer - Avery Zweckform 3658", font=("Arial", 14, "bold")).pack(pady=10)

        # 🧭 Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Check License", command=self.show_license_info)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self.show_about)

        # 📋 Data input
        input_frame = ttk.LabelFrame(main, text="Paste Data (one line per label line)", padding=10)
        input_frame.pack(fill="both", expand=True, pady=5)
        self.text_input = scrolledtext.ScrolledText(input_frame, width=70, height=10)
        self.text_input.pack(fill="both", expand=True)
        self.text_input.bind("<<Modified>>", lambda e: (self.update_preview(), self.text_input.edit_modified(False)))

        # ⚙ Settings
        settings = ttk.LabelFrame(main, text="Label Settings", padding=10)
        settings.pack(fill="x", pady=5)

        ttk.Label(settings, text="Lines per label:").grid(row=0, column=0, padx=5, pady=5)
        self.lines_var = tk.IntVar(value=3)
        ttk.Combobox(settings, textvariable=self.lines_var, values=[1,2,3,4,5,6], width=5, state='readonly').grid(row=0, column=1)
        self.lines_var.trace_add("write", self.update_preview)

        ttk.Label(settings, text="Font:").grid(row=1, column=0, padx=5, pady=5)
        self.font_var = tk.StringVar(value="Arial")
        font_combo = ttk.Combobox(settings, textvariable=self.font_var, values=self.available_fonts, width=25, state='readonly')
        font_combo.grid(row=1, column=1, padx=5)
        font_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        self.bold_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings, text="Bold", variable=self.bold_var, command=self.update_preview).grid(row=1, column=2, padx=10)

        # 👀 Preview
        preview_frame = ttk.LabelFrame(main, text="Preview", padding=10)
        preview_frame.pack(fill="x", pady=5)
        self.preview_canvas = tk.Canvas(preview_frame, width=500, height=120, bg="white", relief="sunken", bd=1)
        self.preview_canvas.pack(padx=10, pady=5)
        self.update_preview()

        # 🧾 Generate
        ttk.Button(main, text="Generate PDF Labels", command=self.generate_labels).pack(pady=20)

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
            "Version 1.3\n\n"
            "Developed by Blaž Pivk\n"
            "© 2025"
        )

    def update_preview(self, *_):
        """Show the first label’s text with selected font"""
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
            messagebox.showwarning("No data", "Enter or paste label data first.")
            return
        lines = [l.strip() for l in text.split("\n") if l.strip()]
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
        for page_start in range(0, len(labels), self.cols * self.rows):
            if page_start > 0: c.showPage()
            for idx, label_lines in enumerate(labels[page_start:page_start + self.cols*self.rows]):
                row, col = divmod(idx, self.cols)
                x = self.left_margin + col * (self.label_width + self.col_gap)
                y = A4[1] - self.top_margin - (row + 1) * self.label_height - row * self.row_gap
                self.draw_label(c, x, y, label_lines, lines_per_label, font_name)
        c.save()

    def draw_label(self, c, x, y, lines, max_lines, font_name):
        pad = 2 * mm
        uw, uh = self.label_width - 2*pad, self.label_height - 2*pad
        size = self.calculate_font_size(lines, uw, uh, max_lines, font_name)
        line_h = size * 1.2
        total_h = len(lines) * line_h
        start_y = y + (self.label_height - total_h) / 2 + pad
        for i, line in enumerate(lines):
            ty = start_y + (len(lines)-1-i) * line_h
            tw = c.stringWidth(line, font_name, size)
            tx = x + (self.label_width - tw) / 2
            c.setFont(font_name, size)
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
