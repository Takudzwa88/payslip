import pandas as pd
from fpdf import FPDF
import os
import yagmail
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
SENDER_EMAIL = os.getenv("EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Ensure the payslips output folder exists
os.makedirs("payslips", exist_ok=True)

# Load employee data from Excel
try:
    df = pd.read_excel("employees.xlsx")
except FileNotFoundError:
    print("❌ Error: employees.xlsx file not found.")
    exit()

# Calculate net salary
df["Net Salary"] = df["Basic Salary"] + df["Allowances"] - df["Deductions"]

# Function to generate a payslip PDF for an employee
def generate_payslip(emp):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, "Payslip", ln=True, align="C")
    pdf.ln(10)

    # Employee details
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, f"Employee Name: {emp['Name']}", ln=True)
    pdf.cell(200, 10, f"Employee ID: {emp['Employee ID']}", ln=True)
    pdf.ln(5)

    # Salary breakdown
    pdf.cell(200, 10, f"Basic Salary: ${emp['Basic Salary']}", ln=True)
    pdf.cell(200, 10, f"Allowances: ${emp['Allowances']}", ln=True)
    pdf.cell(200, 10, f"Deductions: ${emp['Deductions']}", ln=True)
    pdf.cell(200, 10, f"Net Salary: ${emp['Net Salary']}", ln=True)

    # Save the PDF
    file_path = f"payslips/{emp['Employee ID']}.pdf"
    pdf.output(file_path)
    return file_path

# Function to send email with payslip
def send_email(recipient_email, name, file_path):
    try:
        yag = yagmail.SMTP(user="kanokangatakudzwa@gmail.com", password="nsmd oopt evlp vpul")
        subject = "Your Payslip for This Month"
        body = f"""
        Hi {name},

        Please find your payslip for this month attached.

        Best regards,
        HR Department
        """
        yag.send(to=recipient_email, subject=subject, contents=body, attachments=file_path)
        print(f"✅ Payslip sent to {name} ({recipient_email})")
    except Exception as e:
        print(f"❌ Error sending email to {name}: {e}")

# Process each employee row
for index, emp in df.iterrows():
    try:
        pdf_path = generate_payslip(emp)
        send_email(emp["Email"], emp["Name"], pdf_path)
    except Exception as e:
        print(f"❌ Error processing {emp['Name']}: {e}")
