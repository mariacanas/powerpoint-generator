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
    return "✅ PowerPoint Generator API funcionando (reemplazo con conservación de formato)."


# --- Función que reemplaza texto dentro de runs manteniendo formato ---
def reemplazar_marcadores_en_runs(paragraph, nombre_empresa, sector_empresa):
    """
    Reemplaza los marcadores dentro de un párrafo sin perder formato.
    """
    buffer_text = ""
    run_map = []

    # Construir texto completo y mapa de runs
    for i, run in enumerate(paragraph.runs):
        buffer_text += run.text
        run_map.append((i, run.text))

    # Si no hay marcadores, salir
    if "{{" not in buffer_text:
        return

    # Reemplazar en todo el texto plano
    nuevo_texto = buffer_text
    nuevo_texto = nuevo_texto.replace("{{Nombre_Empresa_Cliente}}", nombre_empresa)
    nuevo_texto = nuevo_texto.replace("{{Sector_Empresa_Cliente}}", sector_empresa)

    # Limpiar todos los runs
    for run in paragraph.runs:
        run.text = ""

    # Volver a escribir texto dentro de los mismos runs hasta agotar el texto nuevo
    pos = 0
    for i, run in enumerate(paragraph.runs):
        original = run_map[i][1]
        length = len(original)
        if pos >= len(nuevo_texto):
            break
        run.text = nuevo_texto[pos:pos+length]
        pos += length

    # Si queda texto sobrante (más largo que los runs originales)
    if pos < len(nuevo_texto):
        paragraph.add_run(nuevo_texto[pos:])


@app.route('/generate', methods=['POST'])
def generate_ppt():
    try:
        data = request.get_json()
        print("📥 JSON recibido:", data)

        nombre_empresa = data.get("Nombre_Empresa_Cliente", "")
        sector_empresa = data.get("Sector_Empresa_Cliente", "")
        logo_data = data.get("Logo_Empresa_Cliente", {}).get("data", "")
        plantilla_data = data.get("Plantilla_Base64", "")

        if not plantilla_data:
            return jsonify({"error": "No se recibió la plantilla (Plantilla_Base64)."}), 400

        plantilla_bytes = base64.b64decode(plantilla_data)
        prs = Presentation(io.BytesIO(plantilla_bytes))

        # --- Preparar logo si existe ---
        logo_stream = None
        if logo_data:
            try:
                if isinstance(logo_data, bytes):
                    logo_data = logo_data.decode('utf-8', errors='ignore')
                logo_data = logo_data.replace("\n", "").replace("\r", "")
                image_bytes = base64.b64decode(logo_data)
                image_stream = io.BytesIO(image_bytes)
                Image.open(image_stream).verify()
                image_stream.seek(0)
                logo_stream = image_stream
                print("🖼️ Logo válido cargado correctamente.")
            except Exception as e:
                return jsonify({"error": f"Logo inválido o corrupto: {str(e)}"}), 400

        # --- Procesar cada diapositiva ---
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        # Reemplazar marcadores sin perder formato
                        reemplazar_marcadores_en_runs(paragraph, nombre_empresa, sector_empresa)

                        # Buscar marcador de logo
                        for run in paragraph.runs:
                            if "{{Logo_Empresa_Cliente}}" in run.text and logo_stream:
                                run.text = run.text.replace("{{Logo_Empresa_Cliente}}", "")
                                left, top = shape.left, shape.top
                                slide.shapes.add_picture(
                                    logo_stream, left, top, width=shape.width, height=shape.height
                                )
                                logo_stream.seek(0)

        # --- Guardar y devolver resultado ---
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        encoded_output = base64.b64encode(output.read()).decode("utf-8")

        print("✅ Presentación generada correctamente (formato preservado).")

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
