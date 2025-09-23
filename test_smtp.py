# test_smtp.py
import smtplib
import os
from dotenv import load_dotenv

load_dotenv()

def test_smtp_connection():
    try:
        print("Testing SMTP connection...")
        print(f"Server: {os.getenv('SMTP_SERVER')}")
        print(f"Port: {os.getenv('SMTP_PORT')}")
        print(f"Username: {os.getenv('SMTP_USERNAME')}")
        print(f"Password length: {len(os.getenv('SMTP_PASSWORD', ''))} characters")
        
        server = smtplib.SMTP(os.getenv('SMTP_SERVER'), int(os.getenv('SMTP_PORT')))
        server.ehlo()
        server.starttls()
        server.ehlo()
        
        print("TLS started, attempting login...")
        server.login(os.getenv('SMTP_USERNAME'), os.getenv('SMTP_PASSWORD'))
        print("✅ Login successful!")
        
        server.quit()
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    test_smtp_connection()