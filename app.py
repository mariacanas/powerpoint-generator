from flask import Flask, request, jsonify
from pptx import Presentation
from pptx.util import Inches
import io
import base64
import requests
import os

app = Flask(__name__)

# ===========================
# CONFIGURACIÓN
# ===========================
TENANT_ID = "d5f3680b-66e6-45aa-a1ce-fd3e95f2fdb1"
CLIENT_ID = "0371eba3-368d-4b9b-b74f-474fc67313da"
CLIENT_SECRET = "c404782b-9250-4649-b1c2-7c6dac2ea6f0"
SITE_DOMAIN = "swweb1998.sharepoint.com"
SITE_PATH = "/sites/SegurosPlantilla"
LIBRARY_NAME = "Documentos Generados"
TEMPLATE_URL = "https://swweb1998.sharepoint.com/sites/SegurosPlantilla/Plantillas/Plantilla%20Automatización%20Presentaciones%20Empresas.pptx"

# ===========================
# FUNCIONES AUXILIARES
# ===========================
def get_graph_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    resp = requests.post(url, data=data)
    resp.raise_for_status()
    return resp.json().get("access_token")

def get_site_id(token):
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_DOMAIN}:{SITE_PATH}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]

def get_drive_id(token, site_id):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    drives = resp.json().get("value", [])
    for d in drives:
        if d["name"] == LIBRARY_NAME:
            return d["id"]
    raise Exception(f"No se encontró la biblioteca '{LIBRARY_NAME}'")

def upload_file(token, drive_id, file_name, file_bytes):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_name}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }
    resp = requests.put(url, headers=headers, data=file_bytes)
    resp.raise_for_status()
    return resp.json()["webUrl"]

# ===========================
# RUTAS FLASK
# ===========================
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

        # 2️⃣ Descargar plantilla desde SharePoint
        resp = requests.get(TEMPLATE_URL)
        if resp.status_code != 200:
            return jsonify({"error": f"No se pudo descargar la plantilla, status: {resp.status_code}"}), 400
        prs = Presentation(io.BytesIO(resp.content))

        # 3️⃣ Reemplazar texto en todos los slides
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

        # 5️⃣ Guardar presentación en memoria
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        # 6️⃣ Subir archivo a SharePoint con Graph API
        token = get_graph_token()
        site_id = get_site_id(token)
        drive_id = get_drive_id(token, site_id)
        file_name = f"Presentacion_{nombre_empresa}.pptx"
        file_url = upload_file(token, drive_id, file_name, output.getvalue())

        # 7️⃣ Devolver URL del archivo
        return jsonify({"pptx_url": file_url}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
