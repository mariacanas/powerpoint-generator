from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches
import io
import base64
import requests
import os

app = Flask(__name__)

@app.route('/')
def home():
    return "✅ PowerPoint Generator API funcionando."

@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        # 1️⃣ Recibir JSON desde Power Automate
        data = request.get_json()

        nombre_empresa = data.get("Nombre_Empresa_Cliente", "")
        sector_empresa = data.get("Sector_Empresa_Cliente", "")
        logo_data = data.get("Logo_Empresa_Cliente", {}).get("data", "")

        # 2️⃣ URL de tu plantilla de PowerPoint en SharePoint
        # ⚠️ Si SharePoint requiere login, aquí necesitas un token o enlace público
        plantilla_url = "https://swweb1998.sharepoint.com/sites/SegurosPlantilla/Plantillas/Plantilla%20Automatizaci%C3%B3n%20Presentaciones%20Empresas.pptx"

        # Descargar plantilla desde SharePoint
        resp = requests.get(plantilla_url)
        if resp.status_code != 200:
            return jsonify({"error": f"No se pudo descargar la plantilla. Status code: {resp.status_code}"}), 400

        prs = Presentation(io.BytesIO(resp.content))

        # 3️⃣ Reemplazar texto {{Nombre_Empresa_Cliente}} y {{Sector_Empresa_Cliente}}
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = run.text.replace("{{Nombre_Empresa_Cliente}}", nombre_empresa)
                            run.text = run.text.replace("{{Sector_Empresa_Cliente}}", sector_empresa)

        # 4️⃣ Insertar logo si existe
        if logo_data:
            image_bytes = base64.b64decode(logo_data)
            image_stream = io.BytesIO(image_bytes)

            # Buscar cuadro de texto con {{Logo_Empresa_Cliente}} para posicionar la imagen
            inserted = False
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        if "{{Logo_Empresa_Cliente}}" in shape.text:
                            left = shape.left
                            top = shape.top
                            width = shape.width
                            height = shape.height
                            # Reemplaza el cuadro de texto con la imagen
                            shape.text = ""
                            slide.shapes.add_picture(image_stream, left, top, width, height)
                            inserted = True
                            break
                if inserted:
                    break
            # Si no se encontró placeholder, insertar en la primera diapositiva en posición por defecto
            if not inserted:
                first_slide = prs.slides[0]
                first_slide.shapes.add_picture(image_stream, Inches(1), Inches(1.5), Inches(2), Inches(2))

        # 5️⃣ Guardar presentación en memoria y devolver
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="Presentacion_Personalizada.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
