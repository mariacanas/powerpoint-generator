from flask import Flask, request, jsonify
from pptx import Presentation
import io
import base64
import os
import traceback
from PIL import Image

app = Flask(__name__)

@app.route('/')
def home():
    return "✅ PowerPoint Generator API funcionando (conserva formato original del texto)."

@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        # 1️⃣ Recibir JSON desde Power Automate
        data = request.get_json()
        print("📥 JSON recibido:", data)

        nombre_empresa = data.get("Nombre_Empresa_Cliente", "")
        sector_empresa = data.get("Sector_Empresa_Cliente", "")
        logo_data = data.get("Logo_Empresa_Cliente", {}).get("data", "")
        plantilla_data = data.get("Plantilla_Base64", "")

        if not plantilla_data:
            return jsonify({"error": "No se recibió la plantilla (Plantilla_Base64)."}), 400

        # 2️⃣ Decodificar plantilla y crear presentación
        plantilla_bytes = base64.b64decode(plantilla_data)
        prs = Presentation(io.BytesIO(plantilla_bytes))

        # 3️⃣ Reemplazar marcadores de texto y logo sin alterar formato
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    insertar_logo = False

                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            if "{{Nombre_Empresa_Cliente}}" in run.text:
                                run.text = run.text.replace("{{Nombre_Empresa_Cliente}}", nombre_empresa)
                            if "{{Sector_Empresa_Cliente}}" in run.text:
                                run.text = run.text.replace("{{Sector_Empresa_Cliente}}", sector_empresa)
                            if "{{Logo_Empresa_Cliente}}" in run.text and logo_data:
                                run.text = run.text.replace("{{Logo_Empresa_Cliente}}", "")
                                insertar_logo = True

                    # 4️⃣ Insertar imagen si se detectó el marcador de logo
                    if insertar_logo:
                        if isinstance(logo_data, bytes):
                            logo_data = logo_data.decode('utf-8', errors='ignore')
                        logo_data = logo_data.replace("\n", "").replace("\r", "")
                        try:
                            image_bytes = base64.b64decode(logo_data)
                            image_stream = io.BytesIO(image_bytes)
                            Image.open(image_stream).verify()
                            image_stream.seek(0)
                        except Exception as e:
                            return jsonify({"error": f"Logo inválido: {str(e)}"}), 400

                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(image_stream, left, top, width, height)

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
        print("🔥 Error interno:", str(e))
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)

