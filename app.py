from flask import Flask, request, jsonify
from pptx import Presentation
from pptx.util import Inches
import io
import base64
import os
import requests

app = Flask(__name__)

@app.route('/')
def home():
    return "✅ PowerPoint Generator API funcionando (Opción 2)."

@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        # 1️⃣ Recibir JSON desde Power Automate
        data = request.get_json()
        nombre_empresa = data.get("Nombre_Empresa_Cliente", "")
        sector_empresa = data.get("Sector_Empresa_Cliente", "")
        logo_data = data.get("Logo_Empresa_Cliente", {}).get("data", "")
        plantilla_data = data.get("Plantilla_Base64", "")

        if not plantilla_data:
            return jsonify({"error": "No se recibió la plantilla (Plantilla_Base64)."}), 400

        # 2️⃣ Decodificar plantilla y crear presentación
        plantilla_bytes = base64.b64decode(plantilla_data)
        prs = Presentation(io.BytesIO(plantilla_bytes))

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

            inserted = False
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame and "{{Logo_Empresa_Cliente}}" in shape.text:
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        shape.text = ""
                        slide.shapes.add_picture(image_stream, left, top, width, height)
                        inserted = True
                        break
                if inserted:
                    break
            if not inserted:
                prs.slides[0].shapes.add_picture(image_stream, Inches(1), Inches(1.5), Inches(2), Inches(2))

        # 5️⃣ Guardar presentación en memoria como Base64
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        encoded_output = base64.b64encode(output.read()).decode("utf-8")

        # 6️⃣ Devolver archivo PPTX codificado
        return jsonify({
            "status": "ok",
            "nombre": f"Presentacion_{nombre_empresa}.pptx",
            "file_content": encoded_output
        }), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
