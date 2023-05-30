import pandas as pd
from docx import Document
import re
import os
# We need to have python installed
# We need to have pip installed (python3 pip install)
# We need to have python-docx installed (python3 pip install python-docx)
# We need to have pandas installed installed (python3 pip install pandas)
# We need to have os installed (python3 pip install os)

# Read the Excel file
df = pd.read_excel('diplomas CID.xlsx', header=None)

# Iterate over each row of the DataFrame
for index, row in df.iterrows():
    nombre = str(row.iloc[0])  # Acceder a la primera columna (A) y convertir a cadena
    cursos = str(row.iloc[1])  # Acceder a la segunda columna (B) y convertir a cadena

    # Display values to check
    print(f"Nombre: {nombre}")
    print(f"Cursos: {cursos}")

    # Upload the diploma template
    doc = Document('dip.docx')

    # Replace the labels in the template with the corresponding data
    for paragraph in doc.paragraphs:
        if 'NOMBRE_APELLIDOS' in paragraph.text:
            paragraph.text = paragraph.text.replace('NOMBRE_APELLIDOS', nombre)
        if 'CURSOS' in paragraph.text:
            paragraph.text = paragraph.text.replace('CURSOS', cursos)

    # Remove disallowed characters from the file name
    nombre_archivo = re.sub(r'[^\w.]+', '_', f'diploma_{nombre}_{cursos}.docx')

    # Gets the file extension
    nombre_base, extension = os.path.splitext(nombre_archivo)

    # If the extension is not .docx, add it to the file name
    if extension != '.docx':
        nombre_archivo = f"{nombre_base}.docx"

    # Save the diploma as an individual Word file
    doc.save(nombre_archivo)
