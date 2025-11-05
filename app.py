from flask import Flask, render_template, request, send_file
import pandas as pd
from docxtpl import DocxTemplate
import os
import shutil
import tempfile
import io

app = Flask(__name__)

# --- Configuración de placeholders ---
PLACEHOLDERS_PAGARE = [
    "nombre", "monto_total_pagare", "cuotas_pagare", "valor_cuota", "dia_pago",
    "fecha_primera_cuota", "direccion", "rut", "fecha_suscripcion", "cuotas_pagadas",
    "monto_deuda_pagare", "monto_total_pagare2", "monto_deuda_pagare2", "fecha_suscripcion2",
    "cuotas_pagare2", "valor_cuota2", "dia_pago2", "fecha_primera_cuota2", "cuotas_pagadas2",
    "total_demandado", "comuna_exhorto"
]

PLACEHOLDERS_TRIBUNAL = [
    "tribunal", "apellido_demandado", "rol", "año", "fecha_pagare1",
    "monto_pagare1", "fecha_pagare2", "monto_pagare2"
]

NUMERIC_FIELDS = [
    "monto_total_pagare", "monto_deuda_pagare", "monto_total_pagare2",
    "monto_deuda_pagare2", "valor_cuota", "valor_cuota2", "total_demandado",
    "monto_pagare1", "monto_pagare2"
]


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        template_type = request.form.get('template_type')  # nuevo campo
        excel_file = request.files.get('excel')
        word_file = request.files.get('word')

        if not excel_file or not word_file:
            return render_template('index.html', error="Por favor sube ambos archivos.", template_type=template_type)

        if template_type not in ['pagare', 'tribunal']:
            return render_template('index.html', error="Selecciona un tipo de plantilla válido.", template_type=template_type)

        with tempfile.TemporaryDirectory() as temp_dir:
            excel_path = os.path.join(temp_dir, 'data.xlsx')
            word_path = os.path.join(temp_dir, 'template.docx')
            output_dir = os.path.join(temp_dir, 'Output_Documents')
            os.makedirs(output_dir, exist_ok=True)
            excel_file.save(excel_path)
            word_file.save(word_path)

            data = pd.read_excel(excel_path, dtype=str)
            data.columns = data.columns.str.strip().str.lower()

            # Seleccionar placeholders según el tipo de plantilla
            placeholders = PLACEHOLDERS_PAGARE if template_type == 'pagare' else PLACEHOLDERS_TRIBUNAL

            for index, row in data.iterrows():
                try:
                    doc = DocxTemplate(word_path)
                    context = {}
                    for ph in placeholders:
                        value = row.get(ph, "")
                        if str(value).strip().lower() in ('nan', ''):
                            value = ""
                        if (ph in NUMERIC_FIELDS) and value:
                            try:
                                clean_value = "".join(filter(str.isdigit, str(value)))
                                value = f"{int(clean_value):,}".replace(",", ".")
                            except:
                                pass
                        context[ph] = value
                    doc.render(context)

                    # Nombre de archivo de salida
                    if template_type == 'pagare':
                        safe_name = str(row.get("nombre", f"Documento_{index+1}"))
                    else:
                        safe_name = str(row.get("apellido_demandado", f"Documento_{index+1}"))
                    safe_name = safe_name.replace(" ", "_").replace("/", "_")

                    output_path = os.path.join(output_dir, f"{safe_name}.docx")
                    doc.save(output_path)
                except Exception as e:
                    print(f"Error en fila {index+1}: {e}")
                    continue

            # Crear ZIP
            zip_filename = os.path.join(temp_dir, "Documentos_Generados.zip")
            shutil.make_archive(zip_filename.replace('.zip', ''), 'zip', output_dir)

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
