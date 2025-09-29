import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import openpyxl
from datetime import datetime, timedelta
import hashlib
import json
import base64
import os
from pathlib import Path

class LicenseManager:
    def __init__(self):
        self.license_file = Path.home() / '.labelprinterlicense.dat'
        self.secret_salt = "LabelPrinter2025SecretKey"  # Change this to your own secret
        
    def generate_key(self, email, expiry_date):
        """Generate a license key for a given email and expiry date"""
        # Format: email|YYYY-MM-DD
        data = f"{email}|{expiry_date}|{self.secret_salt}"
        hash_obj = hashlib.sha256(data.encode())
        key = hash_obj.hexdigest()[:24].upper()
        # Format as XXXX-XXXX-XXXX-XXXX-XXXX-XXXX
        formatted = '-'.join([key[i:i+4] for i in range(0, 24, 4)])
        return formatted
    
    def validate_key(self, email, key):
        """Validate a license key and return expiry date if valid"""
        # Try dates from today up to 2 years in the future
        for days in range(0, 730):
            test_date = (datetime.now() + timedelta(days=days)).strftime('%Y-%m-%d')
            expected_key = self.generate_key(email, test_date)
            if expected_key == key.upper().strip():
                return test_date
        return None
    
    def save_license(self, email, expiry_date):
        """Save license information in encrypted format"""
        license_data = {
            'email': email,
            'expiry': expiry_date
        }
        json_data = json.dumps(license_data)
        encoded = base64.b64encode(json_data.encode()).decode()
        
        with open(self.license_file, 'w') as f:
            f.write(encoded)
    
    def load_license(self):
        """Load and validate license from file"""
        if not self.license_file.exists():
            return None
        
        try:
            with open(self.license_file, 'r') as f:
                encoded = f.read()
            
            json_data = base64.b64decode(encoded.encode()).decode()
            license_data = json.loads(json_data)
            
            expiry = datetime.strptime(license_data['expiry'], '%Y-%m-%d')
            if expiry >= datetime.now():
                return license_data
            else:
                return None
        except:
            return None
    
    def get_days_remaining(self):
        """Get number of days remaining in license"""
        license_data = self.load_license()
        if not license_data:
            return 0
        
        expiry = datetime.strptime(license_data['expiry'], '%Y-%m-%d')
        days = (expiry - datetime.now()).days
        return max(0, days)


class RegistrationDialog:
    def __init__(self, parent):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Product Registration")
        self.dialog.geometry("450x250")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (250 // 2)
        self.dialog.geometry(f'+{x}+{y}')
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Product Registration Required", 
                 font=('Arial', 12, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Please enter your license information to continue.",
                 wraplength=400).pack(pady=(0, 20))
        
        # Email field
        email_frame = ttk.Frame(main_frame)
        email_frame.pack(fill=tk.X, pady=5)
        ttk.Label(email_frame, text="Email:", width=12).pack(side=tk.LEFT)
        self.email_var = tk.StringVar()
        ttk.Entry(email_frame, textvariable=self.email_var, width=35).pack(side=tk.LEFT, padx=5)
        
        # License key field
        key_frame = ttk.Frame(main_frame)
        key_frame.pack(fill=tk.X, pady=5)
        ttk.Label(key_frame, text="License Key:", width=12).pack(side=tk.LEFT)
        self.key_var = tk.StringVar()
        ttk.Entry(key_frame, textvariable=self.key_var, width=35).pack(side=tk.LEFT, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="Activate", command=self.activate).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        # Info label
        ttk.Label(main_frame, text="License format: XXXX-XXXX-XXXX-XXXX-XXXX-XXXX",
                 font=('Arial', 8), foreground='gray').pack(pady=(10, 0))
    
    def activate(self):
        email = self.email_var.get().strip()
        key = self.key_var.get().strip()
        
        if not email or not key:
            messagebox.showerror("Error", "Please enter both email and license key")
            return
        
        license_mgr = LicenseManager()
        expiry_date = license_mgr.validate_key(email, key)
        
        if expiry_date:
            license_mgr.save_license(email, expiry_date)
            expiry_obj = datetime.strptime(expiry_date, '%Y-%m-%d')
            days = (expiry_obj - datetime.now()).days
            messagebox.showinfo("Success", 
                              f"License activated successfully!\nValid until: {expiry_date}\n({days} days)")
            self.result = True
            self.dialog.destroy()
        else:
            messagebox.showerror("Error", "Invalid license key. Please check your information.")
    
    def cancel(self):
        self.result = False
        self.dialog.destroy()


class LabelPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Label Printer - Avery Zweckform 3658")
        self.root.geometry("700x600")
        
        # Check license first
        self.license_mgr = LicenseManager()
        if not self.check_license():
            self.root.destroy()
            return
        
        # Avery 3658 specifications (in mm)
        self.label_width = 64.6 * mm
        self.label_height = 33.8 * mm
        self.cols = 3
        self.rows = 8
        self.col_gap = 2.5 * mm
        self.row_gap = 0 * mm
        self.left_margin = 4.7 * mm
        self.top_margin = 13.5 * mm
        
        self.data_lines = []
        
        self.create_widgets()
        self.update_license_status()
    
    def check_license(self):
        """Check if valid license exists, show registration if not"""
        license_data = self.license_mgr.load_license()
        
        if license_data:
            return True
        
        # Show registration dialog
        dialog = RegistrationDialog(self.root)
        self.root.wait_window(dialog.dialog)
        
        return dialog.result == True
    
    def update_license_status(self):
        """Update license status in title bar"""
        days = self.license_mgr.get_days_remaining()
        if days > 30:
            status = f"Licensed"
        elif days > 0:
            status = f"License expires in {days} days"
        else:
            status = "License expired"
        
        self.root.title(f"Label Printer - Avery Zweckform 3658 - {status}")
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title = ttk.Label(main_frame, text="Label Printer - Avery Zweckform 3658", 
                         font=('Arial', 14, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Manage License", command=self.manage_license)
        help_menu.add_command(label="About", command=self.show_about)
        
        # Import section
        import_frame = ttk.LabelFrame(main_frame, text="Import Data", padding="10")
        import_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(import_frame, text="Import Excel File", 
                  command=self.import_excel).grid(row=0, column=0, padx=5)
        ttk.Button(import_frame, text="Clear Data", 
                  command=self.clear_data).grid(row=0, column=1, padx=5)
        
        # Data input section
        input_frame = ttk.LabelFrame(main_frame, text="Paste Data (one line per label line)", 
                                    padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.text_input = scrolledtext.ScrolledText(input_frame, width=70, height=15)
        self.text_input.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Settings section
        settings_frame = ttk.LabelFrame(main_frame, text="Label Settings", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(settings_frame, text="Lines per label:").grid(row=0, column=0, padx=5)
        self.lines_var = tk.IntVar(value=3)
        lines_combo = ttk.Combobox(settings_frame, textvariable=self.lines_var, 
                                   values=[1, 2, 3, 4, 5, 6], width=5, state='readonly')
        lines_combo.grid(row=0, column=1, padx=5)
        
        # Generate button
        ttk.Button(main_frame, text="Generate PDF Labels", 
                  command=self.generate_labels, 
                  style='Accent.TButton').grid(row=4, column=0, columnspan=2, pady=20)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(0, weight=1)
    
    def manage_license(self):
        """Show license management dialog"""
        license_data = self.license_mgr.load_license()
        
        dialog = tk.Toplevel(self.root)
        dialog.title("License Management")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        if license_data:
            ttk.Label(frame, text="Current License Information", 
                     font=('Arial', 12, 'bold')).pack(pady=(0, 10))
            ttk.Label(frame, text=f"Email: {license_data['email']}").pack(pady=5)
            ttk.Label(frame, text=f"Expires: {license_data['expiry']}").pack(pady=5)
            
            days = self.license_mgr.get_days_remaining()
            ttk.Label(frame, text=f"Days remaining: {days}", 
                     foreground='green' if days > 30 else 'orange').pack(pady=5)
        else:
            ttk.Label(frame, text="No active license found", 
                     font=('Arial', 12, 'bold')).pack(pady=(0, 10))
        
        ttk.Button(frame, text="Enter New License Key", 
                  command=lambda: [dialog.destroy(), self.check_license(), 
                                  self.update_license_status()]).pack(pady=20)
        ttk.Button(frame, text="Close", command=dialog.destroy).pack()
    
    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo("About", 
                           "Label Printer for Avery Zweckform 3658\n\n"
                           "Version 1.0\n\n"
                           "Creates professional labels with automatic text sizing.")
    
    def import_excel(self):
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            try:
                wb = openpyxl.load_workbook(filename, data_only=True)
                ws = wb.active
                
                data = []
                for row in ws.iter_rows(values_only=True):
                    if row[0] is not None:
                        data.append(str(row[0]))
                
                self.text_input.delete(1.0, tk.END)
                self.text_input.insert(1.0, '\n'.join(data))
                messagebox.showinfo("Success", f"Imported {len(data)} lines from Excel file")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import Excel file:\n{str(e)}")
    
    def clear_data(self):
        self.text_input.delete(1.0, tk.END)
    
    def generate_labels(self):
        # Check license before generating
        if self.license_mgr.get_days_remaining() <= 0:
            messagebox.showerror("License Expired", 
                               "Your license has expired. Please renew to continue using this software.")
            self.manage_license()
            return
        
        # Get text data
        text_data = self.text_input.get(1.0, tk.END).strip()
        if not text_data:
            messagebox.showwarning("No Data", "Please enter or import data first")
            return
        
        # Split into lines
        lines = [line.strip() for line in text_data.split('\n') if line.strip()]
        lines_per_label = self.lines_var.get()
        
        # Group lines into labels
        labels = []
        for i in range(0, len(lines), lines_per_label):
            label_lines = lines[i:i+lines_per_label]
            labels.append(label_lines)
        
        if not labels:
            messagebox.showwarning("No Data", "No valid data to print")
            return
        
        # Generate PDF
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=f"labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        
        if filename:
            try:
                self.create_pdf(filename, labels, lines_per_label)
                messagebox.showinfo("Success", 
                                   f"Created {len(labels)} labels in:\n{filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create PDF:\n{str(e)}")
    
    def create_pdf(self, filename, labels, lines_per_label):
        c = canvas.Canvas(filename, pagesize=A4)
        page_width, page_height = A4
        
        labels_per_page = self.cols * self.rows
        
        for page_num, i in enumerate(range(0, len(labels), labels_per_page)):
            if page_num > 0:
                c.showPage()
            
            page_labels = labels[i:i+labels_per_page]
            
            for idx, label_text in enumerate(page_labels):
                row = idx // self.cols
                col = idx % self.cols
                
                # Calculate label position
                x = self.left_margin + col * (self.label_width + self.col_gap)
                y = page_height - self.top_margin - (row + 1) * self.label_height - row * self.row_gap
                
                self.draw_label(c, x, y, label_text, lines_per_label)
        
        c.save()
    
    def draw_label(self, c, x, y, lines, max_lines):
        # Add small padding inside label
        padding = 2 * mm
        usable_width = self.label_width - 2 * padding
        usable_height = self.label_height - 2 * padding
        
        # Calculate optimal font size
        font_size = self.calculate_font_size(lines, usable_width, usable_height, max_lines)
        
        # Calculate total text height
        line_height = font_size * 1.2
        total_height = len(lines) * line_height
        
        # Center vertically
        start_y = y + (self.label_height - total_height) / 2 + padding
        
        # Draw each line
        for i, line in enumerate(lines):
            text_y = start_y + (len(lines) - 1 - i) * line_height
            
            # Center horizontally
            text_width = c.stringWidth(line, "Helvetica", font_size)
            text_x = x + (self.label_width - text_width) / 2
            
            c.setFont("Helvetica", font_size)
            c.drawString(text_x, text_y, line)
    
    def calculate_font_size(self, lines, max_width, max_height, max_lines):
        # Start with base font size based on number of lines
        base_sizes = {
            1: 32,
            2: 24,
            3: 18,
            4: 14,
            5: 12,
            6: 10
        }
        font_size = base_sizes.get(max_lines, 10)
        
        # Check if text fits and reduce if necessary
        while font_size > 6:
            line_height = font_size * 1.2
            total_height = len(lines) * line_height
            
            # Check height
            if total_height > max_height:
                font_size -= 1
                continue
            
            # Check width for all lines
            fits = True
            c_temp = canvas.Canvas("temp")
            for line in lines:
                text_width = c_temp.stringWidth(line, "Helvetica", font_size)
                if text_width > max_width:
                    fits = False
                    break
            
            if fits:
                break
            
            font_size -= 1
        
        return max(font_size, 6)

def main():
    root = tk.Tk()
    app = LabelPrinterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()