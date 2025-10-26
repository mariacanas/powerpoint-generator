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
    return "✅ PowerPoint Generator API funcionando (versión mejorada con reemplazo múltiple y logo en todas las diapositivas)."


# --- Función simple y robusta para reemplazar texto ---
def reemplazar_marcadores_en_texto(text, nombre_empresa, sector_empresa):
    text = text.replace("{{Nombre_Empresa_Cliente}}", nombre_empresa)
    text = text.replace("{{Sector_Empresa_Cliente}}", sector_empresa)
    return text


@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        data = request.get_json()
        print("📥 JSON recibido:", data)

        nombre_empresa = data.get("Nombre_Empresa_Cliente", "")
        sector_empresa = data.get("Sector_Empresa_Cliente", "")
        logo_data = data.get("Logo_Empresa_Cliente", {}).get("data", "")
        plantilla_data = data.get("Plantilla_Base64", "")

        # Validaciones iniciales
        if not plantilla_data:
            return jsonify({"error": "No se recibió la plantilla (Plantilla_Base64)."}), 400

        plantilla_bytes = base64.b64decode(plantilla_data)
        prs = Presentation(io.BytesIO(plantilla_bytes))

        # --- Procesar el logo si se envió ---
        logo_stream = None
        if logo_data:
            try:
                if isinstance(logo_data, bytes):
                    logo_data = logo_data.decode('utf-8', errors='ignore')
                logo_data = logo_data.replace("\n", "").replace("\r", "")
                image_bytes = base64.b64decode(logo_data)
                image_stream = io.BytesIO(image_bytes)

                # Verificar que realmente sea una imagen válida
                Image.open(image_stream).verify()
                image_stream.seek(0)
                logo_stream = image_stream
                print("🖼️ Logo válido cargado correctamente.")
            except Exception as e:
                return jsonify({"error": f"Logo inválido o corrupto: {str(e)}"}), 400

        # --- Recorrer todas las diapositivas y reemplazar marcadores ---
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame

                    for paragraph in text_frame.paragraphs:
                        # Reemplazar todos los marcadores de texto
                        nuevo_texto = reemplazar_marcadores_en_texto(
                            paragraph.text,
                            nombre_empresa,
                            sector_empresa
                        )

                        # Si hay marcador de logo, insertarlo
                        if "{{Logo_Empresa_Cliente}}" in nuevo_texto and logo_stream:
                            paragraph.text = nuevo_texto.replace("{{Logo_Empresa_Cliente}}", "")
                            left, top = shape.left, shape.top
                            slide.shapes.add_picture(
                                logo_stream, left, top, width=shape.width, height=shape.height
                            )
                            logo_stream.seek(0)  # permitir reutilizarlo
                        else:
                            paragraph.text = nuevo_texto

        # --- Guardar la presentación resultante ---
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        encoded_output = base64.b64encode(output.read()).decode("utf-8")

        print("✅ Presentación generada correctamente.")

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
