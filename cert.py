# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import letter
# from PyPDF2 import PdfReader, PdfWriter

# # Path to the uploaded certificate PDF
# certificate_path = 'certificate.pdf'

# # Function to create an overlay PDF with student's details
# def create_overlay_pdf(student_name, date, overlay_pdf_path):
#     c = canvas.Canvas(overlay_pdf_path, pagesize=letter)
#     # Assuming we have the positions for the text, these values might need adjustments
#     text_position_name = 400, 315# Example positions

#     text_position_date = 200, 250
    
#     c.setFont("Helvetica", 12)
#     c.drawString(text_position_name[0], text_position_name[1], f"{student_name}")
#     c.drawString(text_position_date[0], text_position_date[1], f"{date}")
    
#     c.save()
#     return overlay_pdf_path

# # Function to merge the overlay with the original certificate
# def merge_pdfs(certificate_path, overlay_pdf_path, output_pdf_path):
#     # Read the original certificate
#     reader = PdfReader(certificate_path)
#     certificate_page = reader.pages[0]
    
#     # Read the overlay PDF
#     overlay_reader = PdfReader(overlay_pdf_path)
#     overlay_page = overlay_reader.pages[0]
    
#     # Merge the overlay with the certificate
#     certificate_page.merge_page(overlay_page)
    
#     # Write the merged output
#     writer = PdfWriter()
#     writer.add_page(certificate_page)
    
#     with open(output_pdf_path, 'wb') as output_file:
#         writer.write(output_file)

# # Example student details
# student_name = "Jane Doe"
# date = "2024-02-21"

# # Paths for the overlay PDF and the final output PDF
# overlay_pdf_path = 'overlay.pdf'
# output_pdf_path = 'certificate.pdf'

# # Create an overlay PDF
# overlay_pdf = create_overlay_pdf(student_name, date, overlay_pdf_path)

# # Merge the overlay with the original certificate
# merge_pdfs(certificate_path, overlay_pdf, output_pdf_path)

# # Return the path to the generated certificate for the example student
# output_pdf_path 



import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import os

# Function to create an overlay PDF with a student's name
def create_overlay_pdf(student_name, overlay_pdf_path):
    c = canvas.Canvas(overlay_pdf_path, pagesize=letter)
    # Adjust these positions as needed
    text_position_name = 400, 315# Example positions

    text_position_date = 200, 250
    
    c.setFont("Helvetica", 12)
    c.drawString(text_position_name[0], text_position_name[1], f"{student_name}")
    c.drawString(text_position_date[0], text_position_date[1], "23-02-2024")
    
    c.save()

# Function to merge the overlay PDF with the original certificate
def merge_pdfs(certificate_path, overlay_pdf_path, output_pdf_path):
    reader = PdfReader(certificate_path)
    certificate_page = reader.pages[0]

    overlay_reader = PdfReader(overlay_pdf_path)
    overlay_page = overlay_reader.pages[0]

    certificate_page.merge_page(overlay_page)

    writer = PdfWriter()
    writer.add_page(certificate_page)

    with open(output_pdf_path, 'wb') as output_file:
        writer.write(output_file)

# Load the Excel file with student names
xls_path = 'xls.xlsx'
df = pd.read_excel(xls_path, usecols="B", skiprows=2, nrows=354)

# Directory to save the generated certificates
output_directory = 'generated_certificates'
os.makedirs(output_directory, exist_ok=True)

# Path to the uploaded certificate PDF (original template)
certificate_path = 'certificate.pdf'

# Iterate over student names and generate certificates
for index, row in df.iterrows():
    student_name = row[0]
    overlay_pdf_path = os.path.join(output_directory, f'overlay_{index}.pdf')
    output_pdf_path = os.path.join(output_directory, f'certificate_{student_name.replace(" ", "_")}.pdf')
    
    # Create an overlay PDF with the student's name
    create_overlay_pdf(student_name, overlay_pdf_path)
    
    # Merge the overlay with the original certificate
    merge_pdfs(certificate_path, overlay_pdf_path, output_pdf_path)

print("Certificates generation process is complete.")
