import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def excel_to_vcf(excel_file, output_vcf):
    # Load Excel file into a pandas DataFrame
    df = pd.read_csv(excel_file, encoding='utf-8')

    # Create and open the vCard file for writing with UTF-8 encoding
    with open(output_vcf, 'w', encoding='utf-8') as vcf_file:
        # Iterate through rows in the DataFrame
        for index, row in df.iterrows():
            # Extract information from the DataFrame
            phone = row['Telefon']
            f_name = row['Prenume']
            l_name = row['Nume']

            # Write vCard format to the file
            vcf_file.write('BEGIN:VCARD\n')
            vcf_file.write(f'FN:{f_name} {l_name}\n')
            vcf_file.write(f'N:{l_name};{f_name};;;\n')
            vcf_file.write(f'TEL:{phone}\n')
            vcf_file.write('END:VCARD\n')

def convert_button_click():
    input_excel_file = filedialog.askopenfilename(title="Select Excel/CSV File", filetypes=[("Excel/CSV Files", "*.xls;*.xlsx;*.csv")])
    if not input_excel_file:
        return  # User canceled the file dialog

    output_vcf_file = filedialog.asksaveasfilename(defaultextension=".vcf", filetypes=[("vCard Files", "*.vcf")])
    if not output_vcf_file:
        return  # User canceled the file dialog

    try:
        excel_to_vcf(input_excel_file, output_vcf_file)
        messagebox.showinfo("Conversion Complete", f"Conversion complete. vCard file saved to '{output_vcf_file}'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the main window
root = tk.Tk()
root.title("Excel to vCard Converter")

# Create a label with instructions
instructions_label = tk.Label(root, text="Please select your Excel file and choose where to save it as a vCard file.")
instructions_label.pack(pady=10)

# Create a Convert button
convert_button = tk.Button(root, text="Convert", command=convert_button_click)
convert_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
