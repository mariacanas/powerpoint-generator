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
    return "✅ PowerPoint Generator API funcionando (Opción 2 con validación de logo)."

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
            print("❌ Plantilla_Base64 no recibida")
            return jsonify({"error": "No se recibió la plantilla (Plantilla_Base64)."}), 400

        # 2️⃣ Decodificar plantilla y crear presentación
        plantilla_bytes = base64.b64decode(plantilla_data)
        prs = Presentation(io.BytesIO(plantilla_bytes))

        # 3️⃣ Reemplazar marcadores de texto y logo sin duplicación ni cambio de tamaño
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame

                    # 1. Reconstruir todo el texto del shape
                    full_text = ""
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            full_text += run.text

                    # 2. Reemplazar marcadores
                    full_text = full_text.replace("{{Nombre_Empresa_Cliente}}", nombre_empresa)
                    full_text = full_text.replace("{{Sector_Empresa_Cliente}}", sector_empresa)

                    # 3. Copiar estilo del primer run (si existe)
                    font_size = None
                    font_bold = None
                    font_italic = None
                    font_color = None
                    try:
                        first_paragraph = text_frame.paragraphs[0]
                        if first_paragraph.runs:
                            first_run = first_paragraph.runs[0]
                            font_size = first_run.font.size
                            font_bold = first_run.font.bold
                            font_italic = first_run.font.italic
                            if first_run.font.color and first_run.font.color.rgb:
                                font_color = first_run.font.color.rgb
                    except Exception as e:
                        print("⚠️ No se pudo copiar estilo:", str(e))

                    # 4. Eliminar todos los párrafos y crear uno nuevo
                    try:
                        while text_frame.paragraphs:
                            text_frame._element.remove(text_frame.paragraphs[0]._p)
                    except Exception as e:
                        print("⚠️ Error al eliminar párrafos:", str(e))

                    new_paragraph = text_frame.add_paragraph()
                    new_run = new_paragraph.add_run()
                    new_run.text = full_text

                    # 5. Aplicar estilo copiado
                    if font_size: new_run.font.size = font_size
                    if font_bold is not None: new_run.font.bold = font_bold
                    if font_italic is not None: new_run.font.italic = font_italic
                    if font_color: new_run.font.color.rgb = font_color

                    # 6. Insertar logo si el marcador estaba presente
                    if "{{Logo_Empresa_Cliente}}" in full_text and logo_data:
                        if isinstance(logo_data, bytes):
                            logo_data = logo_data.decode('utf-8', errors='ignore')
                        logo_data = logo_data.replace("\n", "").replace("\r", "")
                        try:
                            image_bytes = base64.b64decode(logo_data)
                            image_stream = io.BytesIO(image_bytes)
                            Image.open(image_stream).verify()
                            image_stream.seek(0)
                        except Exception as e:
                            print("❌ Error al procesar el logo:", str(e))
                            return jsonify({"error": f"Logo inválido: {str(e)}"}), 400

                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        slide.shapes.add_picture(image_stream, left, top, width, height)

        # 4️⃣ Guardar presentación en memoria como Base64
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        encoded_output = base64.b64encode(output.read()).decode("utf-8")

        # 5️⃣ Devolver archivo PPTX codificado
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
