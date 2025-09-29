import hashlib
from datetime import datetime, timedelta

class KeyGenerator:
    def __init__(self):
        self.secret_salt = "LabelPrinter2025SecretKey"  # Must match the one in main app
    
    def generate_key(self, email, duration_days=365):
        """Generate a license key for a customer"""
        expiry_date = (datetime.now() + timedelta(days=duration_days)).strftime('%Y-%m-%d')
        
        data = f"{email}|{expiry_date}|{self.secret_salt}"
        hash_obj = hashlib.sha256(data.encode())
        key = hash_obj.hexdigest()[:24].upper()
        formatted = '-'.join([key[i:i+4] for i in range(0, 24, 4)])
        
        return formatted, expiry_date

def main():
    print("=" * 60)
    print("LICENSE KEY GENERATOR - Label Printer")
    print("=" * 60)
    print()
    
    generator = KeyGenerator()
    
    email = input("Enter customer email: ").strip()
    
    try:
        days = int(input("Enter license duration in days (default 365): ").strip() or "365")
    except ValueError:
        days = 365
    
    print("\nGenerating license key...")
    key, expiry = generator.generate_key(email, days)
    
    print("\n" + "=" * 60)
    print("LICENSE INFORMATION")
    print("=" * 60)
    print(f"Email:       {email}")
    print(f"License Key: {key}")
    print(f"Expires:     {expiry}")
    print(f"Duration:    {days} days")
    print("=" * 60)
    print("\nProvide this information to the customer.")
    print()

if __name__ == "__main__":
    main()