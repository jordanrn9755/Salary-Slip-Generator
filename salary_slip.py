import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import os
from num2words import num2words
import math

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=[("All Files", "*.*"), ("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")]
    )
    if file_path:
        file_label.config(text=file_path)
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, engine='openpyxl') 
            else:
                messagebox.showerror("Invalid File", "Please select a valid CSV or XLSX file.")
                return

            df.columns = df.columns.str.strip()

            print("Cleaned column names:", df.columns.tolist())

            process_data(df)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")

def process_data(df):
    try:
        required_columns = ['Emp. No', 'Name', 'Designation', 'Gross Salary', 'Professional Tax', 
                            'Working days', 'Unpaid Leaves', 'Paid Leaves', 'PF UAN', 'Bank Ac No', 'Bank', 'IFSC']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise Exception(f"Missing columns in file: {', '.join(missing_columns)}")

        df['Total Working days'] = df['Working days'] + df['Paid Leaves']

        df['Basic Salary'] = ((df['Gross Salary'] * 0.5) / df['Month days']) * df['Total Working days']
        
        df['Leave Deduction'] = (df['Gross Salary'] / df['Month days']) * df['Unpaid Leaves']
        df['PF Deduction'] = (df['Basic Salary'] - df['Leave Deduction']) * 0.12
        
        df['ESIC Deduction'] = df.apply(lambda row: row['Gross Salary'] * 0.0075 if row['Gross Salary'] <= 21000 else 0, axis=1)

        df['Net Salary'] = df['Gross Salary'] - df['Professional Tax'] - df['PF Deduction'] - df['ESIC Deduction'] - df['Leave Deduction']
        
        generate_pdf(df)
        messagebox.showinfo("Success", "Salary slips generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error processing salary data: {str(e)}")


class SalarySlipPDF(FPDF):
    def employee_details(self, name, emp_no, designation, department, uan, bank, month,bank_ac_no):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Shaleemar IT Solutions Pvt. Ltd.', 0, 1, 'C')
        self.cell(0, 10, f'Salary Slip for {month} - 2024', 0, 1, 'C')
        
        name = str(name)
        emp_no = str(emp_no)
        designation = str(designation)
        department = str(department)
        uan = str(uan)
        bank = str(bank)
        bank_ac_no = str(bank_ac_no)
        
        col_widths = [40,55,40,55]  

        self.set_font('Arial', 'B', 10)
        
        rows = [
            ('Name', name, 'Emp. No', emp_no),
            ('Designation', designation, 'Department', department),
            ('','', 'Bank', bank),
            ('PF UAN', uan,'Bank Ac No',bank_ac_no),
            ('','','','')
        ]
        
        for row in rows:
            for field, detail in zip(row[::2], row[1::2]): 
                self.cell(col_widths[0], 5, field, 1)
                self.cell(col_widths[1], 5, detail, 1)
            self.ln()  
        

    def earnings_deductions(self, earnings, deductions):
        self.set_font('Arial', 'B', 10)
        
        self.cell(95, 5, 'Earnings', 1, 0, 'C')
        self.cell(95, 5, 'Deductions', 1, 1, 'C')
        
        col_widths = [47, 48, 47, 48] 
        
        self.set_font('Arial', '', 10)
        
        max_rows = max(len(earnings), len(deductions))
        
        for i in range(max_rows):
            if i < len(earnings):
                self.cell(col_widths[0], 5, f"{earnings[i][0]}", 1)
                self.cell(col_widths[1], 5, f"{earnings[i][1]}", 1)
            else:
                self.cell(col_widths[0], 5, '', 1)
                
            if i < len(deductions):
                self.cell(col_widths[2], 5, f"{deductions[i][0]}", 1)
                self.cell(col_widths[3], 5, f"{deductions[i][1]}", 1)
            else:
                self.cell(col_widths[2], 5, '', 1)            
            self.ln() 

    def gross_salary_net_pay(self, gross_salary, total_deductions, net_pay):
        col_widths = [47, 48, 47, 48] 

        self.set_font('Arial', 'B', 10)
        self.cell(col_widths[0], 5, 'Gross Salary', 1)
        self.cell(col_widths[1], 5, str(gross_salary), 1)
        self.cell(col_widths[2], 5, 'Total Deductions', 1)
        self.cell(col_widths[3], 5, str(total_deductions), 1)
        self.ln()
        
        self.cell(col_widths[0], 5, 'Net Pay', 1)
        self.cell(col_widths[1], 5, str(net_pay), 1)
        self.cell(col_widths[2], 5, '', 1)
        self.cell(col_widths[3], 5, '', 1)
        self.ln()
        
        self.cell(47,5,'Amount in words', 1)
        self.cell(143,5,f"{num2words(custom_round(net_pay), lang='en').title()} Only",1)

    def signature(self):
        self.ln(20)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'For Shaleemar IT Solutions Pvt Ltd', 0, 1, 'L')
        self.ln(20)
        self.cell(0, 10, 'HR Manager', 0, 1, 'L')
        
        self.ln(90)
        
        self.set_font("Arial", size=10)
        self.cell(200, 5, txt="Shaleemar IT Solutions Pvt. Ltd, Second Floor, Office No 03, Sneh Avishkar,", ln=True, align='C')
        self.cell(200, 5, txt="Plot No 105, Prabhat Road, Erandawa Deccan, Gymkhana, Pune 411004, Maharashtra, India", ln=True, align='C')



def generate_pdf(df):
    try:
        output_folder = "salary_slips"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        for index, row in df.iterrows():
            pdf = SalarySlipPDF()
            pdf.add_page()

            pdf.employee_details(
                name=row['Name'], 
                emp_no=row['Emp. No'], 
                designation=row['Designation'], 
                department=row['Department'], 
                uan=row['PF UAN'], 
                bank=row['Bank'],
                month=row['Month'],
                bank_ac_no=row['Bank Ac No'],
            )

            earnings = [
                ("Basic Salary", custom_round(row['Basic Salary'])),
                ("House Rent Allowance", custom_round(((row['Gross Salary'] * 0.2) / row['Month days']) * row['Working days'])),
                ("Medical Allowance", custom_round(((row['Gross Salary'] * 0.1) / row['Month days']) * row['Working days'])),
                ("Canteen Allowance", custom_round(((row['Gross Salary'] * 0.05) / row['Month days']) * row['Working days'])),
                ("Attendance Allowance", custom_round(((row['Gross Salary'] * 0.15) / row['Month days']) * row['Working days']))
            ]
            
            total_earnings = sum(amount for _, amount in earnings)

            deductions = [
                ("Professional Tax", custom_round(row['Professional Tax'])),
                ("Employee PF", custom_round(row['Basic Salary'] * 0.12)),
                ("ESIC", custom_round(total_earnings * 0.0075) if row['Gross Salary'] <= 21500 else 0),
                ("TDS", custom_round(0)),
                ("Leave Deduction", custom_round(row['Leave Deduction']))
            ]

            pdf.earnings_deductions(earnings, deductions)

            total_deductions = sum([deduction[1] for deduction in deductions])
            net_salary = custom_round(row['Gross Salary'] - total_deductions)

            pdf.gross_salary_net_pay(
                gross_salary=custom_round(row['Gross Salary']),
                total_deductions=custom_round(total_deductions),
                net_pay=net_salary
            )

            pdf.signature()

            pdf_output_path = os.path.join(output_folder, f"{row['Name']}_Salary_Slip.pdf")
            pdf.output(pdf_output_path)

            print(f"Generated salary slip for {row['Name']}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate PDFs: {str(e)}")


def custom_round(value):
    decimal_part = value - int(value)
    
    if decimal_part >= 0.50:
        return math.ceil(value)
    else:
        return math.floor(value)


root = tk.Tk()
root.title("Salary Slip Generator")

root.configure(bg='#f0f0f0') 

frame = tk.Frame(root, bg='#ffffff', padx=20, pady=20)
frame.pack(expand=True, fill=tk.BOTH)

file_label = tk.Label(frame, text="No file selected", wraplength=350, bg='#ffffff', fg='#333333')
file_label.pack(pady=10)

file_button = tk.Button(frame, text="Select CSV or XLSX File", command=select_file, bg='#4CAF50', fg='#ffffff')
file_button.pack(pady=10)

instructions_label = tk.Label(frame, text="Please select a CSV or XLSX file to generate the salary slip.", wraplength=350, bg='#ffffff', fg='#666666')
instructions_label.pack(pady=5)

root.geometry("400x200")
root.resizable(False, False)

root.mainloop()

