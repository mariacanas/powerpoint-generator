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
            try:
                first_run = text_frame.paragraphs[0].runs[0]
                font_size = first_run.font.size
                font_bold = first_run.font.bold
                font_italic = first_run.font.italic
                font_color = first_run.font.color.rgb
            except IndexError:
                font_size = None
                font_bold = None
                font_italic = None
                font_color = None

            # 4. Eliminar todos los párrafos y crear uno nuevo
            while text_frame.paragraphs:
                text_frame._element.remove(text_frame.paragraphs[0]._p)

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
                    return jsonify({"error": f"Logo inválido: {str(e)}"}), 400

                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes.add_picture(image_stream, left, top, width, height)

