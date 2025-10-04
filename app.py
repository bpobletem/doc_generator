from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
from docxtpl import DocxTemplate
import os
import shutil
import tempfile

app = Flask(__name__)

PLACEHOLDERS = ["nombre", "monto_pagare", "monto_demandado", "direccion", "rut", "fecha_suscripcion", "fecha_primera_cuota", "cantidad_cuotas", "dia_pago", "comuna_exhorto"]

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files.get('excel')
        word_file = request.files.get('word')
        if not excel_file or not word_file:
            return render_template('index.html', error="Please upload both files.")

        import io
        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, 'data.xlsx')
            word_path = os.path.join(temp_dir, 'template.docx')
            output_dir = os.path.join(temp_dir, 'Output_Documents')
            os.makedirs(output_dir, exist_ok=True)
            excel_file.save(excel_path)
            word_file.save(word_path)

            data = pd.read_excel(excel_path, dtype=str)
            data.columns = data.columns.str.strip().str.lower()

            for index, row in data.iterrows():
                try:
                    doc = DocxTemplate(word_path)
                    context = {}
                    for ph in PLACEHOLDERS:
                        value = row.get(ph, "")
                        if str(value).strip().lower() in ('nan', ''):
                            value = ""
                        if (ph == "monto_pagare" or ph == "monto_demandado") and value:
                            try:
                                clean_value = "".join(filter(str.isdigit, str(value)))
                                value = f"{int(clean_value):,}".replace(",", ".")
                            except:
                                pass
                        context[ph] = value
                    doc.render(context)
                    safe_name = str(row.get("nombre", f"Documento_{index+1}")).replace(" ", "_").replace("/", "_")
                    output_path = os.path.join(output_dir, f"{safe_name}.docx")
                    doc.save(output_path)
                except Exception as e:
                    continue
            zip_filename = os.path.join(temp_dir, "Documentos_Generados.zip")
            shutil.make_archive(zip_filename.replace('.zip',''), 'zip', output_dir)
            # Read ZIP into memory before sending
            with open(zip_filename, 'rb') as f:
                zip_bytes = f.read()
            return send_file(
                io.BytesIO(zip_bytes),
                mimetype='application/zip',
                as_attachment=True,
                download_name='Documentos_Generados.zip'
            )
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)