import csv
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter, PdfFileMerger
import os
import subprocess
import argparse

script_dir = os.path.dirname(os.path.abspath(__file__))

global ghostscript_path
ghostscript_path = r'C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe'

def convert_to_pdfa_with_metadata(input_path, output_path, ghostscript_path=ghostscript_path): 
    subprocess.run([ghostscript_path, "-dPDFA", "-dBATCH", "-dNOPAUSE", "-sProcessColorModel=DeviceRGB", "-sDEVICE=pdfwrite", "-dPDFACompatibilityPolicy=1", "-sOutputFile=" + output_path, "-dEmbedAllFonts=true", "-dSubsetFonts=true", "-dAutoRotatePages=/None", input_path, 'mydocinfo.pdfmark']) 
    print(f"Converted '{input_path}' to PDF/A with metadata added and saved as '{output_path}'.")

def convert_to_pdfa(input_path, output_path, ghostscript_path=ghostscript_path): 
    subprocess.run([ghostscript_path, "-dPDFA", "-dBATCH", "-dNOPAUSE", "-sProcessColorModel=DeviceRGB", "-sDEVICE=pdfwrite", "-dPDFACompatibilityPolicy=1", "-sOutputFile=" + output_path, "-dEmbedAllFonts=true", "-dSubsetFonts=true", "-dAutoRotatePages=/None", input_path]) 
    
    print(f"Converted '{input_path}' to PDF/A and saved as '{output_path}'.")

def main(input_folder, output_folder):
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Iterate over each folder in the input folder
    for folder_name in os.listdir(input_folder):
        folder_path = os.path.join(input_folder, folder_name)
        
        if os.path.isdir(folder_path):
            # Read CSV file from each folder
            print(folder_path, folder_name + '.csv')
            csv_file = os.path.join(folder_path, folder_name + '.csv')
            
            rows = []
            if os.path.exists(csv_file):
                with open(csv_file, 'r', newline='', encoding='ISO-8859-1') as csvfile:
                    reader = csv.reader(csvfile)
                    #   rows = list(reader)
                    
                    for row in reader:
                        rows.append(row)

                rows[0] = [
                    'barcode', 'Case_Type', 'Case_Number', 'Case_Year', 'Petitioner_Name1', 'Respondent_Name1',
                    'Petitioner_Advocate_Name', 'Respondent_Advocate_Name', 'Crime_District', 'Category_Code',
                    'Judge_Name', 'Date_of_Disposal', 'section', 'Bundle_name', 'PDF_name', 'PDF Size', 'User_ID',
                    'pdfcount', 'Char_Count', 'scandate', 'verified_by', 'Date of verification', 'Remarks'
                ]


                with open(csv_file, 'w', newline='', encoding='ISO-8859-1') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerows(rows)


                # Read PDF names from CSV and convert each to PDF/A
                with open(csv_file, 'r', newline='') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        pdf_name = row['PDF_name']
                        Respondent_Name1 = row['Respondent_Name1']
                        Petitioner_Name1 = row['Petitioner_Name1']
                        Petitioner_Advocate_Name = row['Petitioner_Advocate_Name']
                        Judge_Name = row['Judge_Name']
                        Crime_District = row['Crime_District']
                        Date_of_Disposal = row['Date_of_Disposal']
                        Case_Year = row['Case_Year']
                        Case_Type = row['Case_Type']
                        Case_Number = row['Case_Number']
                        barcode = row['barcode']

                        pdf_folder = os.path.join(folder_path, pdf_name.replace('.pdf', ''))

                        for file_name in os.listdir(pdf_folder):
                            if file_name.endswith('.pdf') and file_name != pdf_name:
                                input_pdf_path = os.path.join(pdf_folder, file_name)
                                output_subfolder = os.path.join(output_folder, folder_name, pdf_name.replace('.pdf', ''))
                                output_pdfa_path = os.path.join(output_subfolder, file_name.replace('.pdf', '.pdf'))

                                # Create output subfolder if it doesn't exist
                                if not os.path.exists(output_subfolder):
                                    os.makedirs(output_subfolder)

                                convert_to_pdfa(input_pdf_path, output_pdfa_path)
                            
                            elif file_name.endswith('.pdf') and file_name.strip() == pdf_name.strip():
                                input_pdf_path = os.path.join(pdf_folder, file_name)
                                
                                output_subfolder = os.path.join(output_folder, folder_name, pdf_name.replace('.pdf', ''))
                                output_pdfa_path = os.path.join(output_subfolder, file_name.replace('.pdf', '.pdf'))

                                if not os.path.exists(output_subfolder):
                                    os.makedirs(output_subfolder)

                                content = f"""[ /Respondent_Name ({Respondent_Name1})
                                  /Petitioner_Name ({Petitioner_Name1})
                                  /Petitioner_Advocate_Name ({Petitioner_Advocate_Name})
                                  /Judge_Name ({Judge_Name})
                                  /Crime_District ({Crime_District})
                                  /Date_of_Disposal ({Date_of_Disposal})
                                  /Case_Year ({Case_Year})
                                  /Case_Type ({Case_Type})
                                  /Case_Number ({Case_Number})
                                  /barcode ({barcode})
                                  /DOCINFO
                                pdfmark"""

                                metadata_file = os.path.join(script_dir, 'mydocinfo.pdfmark')

                                with open(metadata_file, 'w') as f:
                                    f.write(content)

                                convert_to_pdfa_with_metadata(input_pdf_path, output_pdfa_path)

if __name__ == "__main__":

    input_folder=''
    parser = argparse.ArgumentParser(description='Process folder location.')
    parser.add_argument('-i', '--input_folder', type=str, help='Input folder location')
    
    args = parser.parse_args()
    if args.input_folder:
        input_folder=args.input_folder
        pass
    else:
        print("Please provide a folder location using the -i flag.")

    # Output folder for PDF/A files
    output_folder = 'modified_files'

    main(input_folder, output_folder)

    os.remove(os.path.join(script_dir, 'mydocinfo.pdfmark'))
