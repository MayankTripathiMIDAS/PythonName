from flask import Flask, request, jsonify
import requests
from io import BytesIO
from docx import Document
import fitz
import tempfile
import os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def convert_pdf_to_html(pdf_bytes, file_extension):
    fileextension = file_extension.replace(".", "")
    doc = fitz.open(stream=pdf_bytes, filetype=fileextension)
    html_content = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        html_content += page.get_text("html")
    return html_content

def convert_docx_to_html(docx_bytes):
    document = Document(BytesIO(docx_bytes))
    html_content = "<html><body>"
    for paragraph in document.paragraphs:
        paragraph_html = "<p>"
        for run in paragraph.runs:
            paragraph_html += run.text
        paragraph_html += "</p>"
        html_content += paragraph_html
    html_content += "</body></html>"
    return html_content

def convert_doc_to_html(doc_bytes):
    try:
        # First try textract if dependencies are available
        try:
            import textract
            with tempfile.NamedTemporaryFile(suffix='.doc', delete=False) as temp_file:
                temp_file.write(doc_bytes)
                temp_path = temp_file.name
            
            text = textract.process(temp_path).decode('utf-8')
            os.unlink(temp_path)
            return format_as_html(text)
        except:
            # Fallback to pure Python approach
            from olefile import OleFileIO
            from io import BytesIO
            
            ole = OleFileIO(BytesIO(doc_bytes))
            if ole.exists('WordDocument'):
                stream = ole.openstream('WordDocument')
                text = stream.read().decode('latin-1')  # Basic extraction
                return format_as_html(text)
            else:
                return "<html><body>Unsupported .doc format</body></html>"
    except Exception as e:
        return f"<html><body>Error processing .doc file: {str(e)}</body></html>"

def format_as_html(text):
    return f"""
    <html>
        <body>
            <div style="white-space: pre-wrap; font-family: Arial, sans-serif">
                {text}
            </div>
        </body>
    </html>
    """
@app.route('/highlight-file', methods=['POST'])
def highlight_file():
    try:
        data = request.get_json()
        item_id = data.get("itemId")
        token = data.get("token")
        file_extension = data.get("fileExtension", "").lower()

        if not item_id or not token:
            return jsonify({"error": "itemId and token are required"}), 400
        
        file_content_url = f"https://graph.microsoft.com/v1.0/sites/midasconsultingmgmt.sharepoint.com,6ca0fab8-2a87-4e15-a144-d87634dcb569,1b3d5672-7447-4188-982e-126402613a10/drives/b!uPqgbIcqFU6hRNh2NNy1aXJWPRtHdIhBmC4SZAJhOhBCF-UF6RIYQ7WCbzH_wEcf/items/{item_id}/content"
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(file_content_url, headers=headers)
        
        if response.status_code != 200:
            return jsonify({"error": "Failed to fetch file", "details": response.text}), response.status_code

        file_content = response.content

        if file_extension == ".pdf":
            html_content = convert_pdf_to_html(file_content, file_extension)
        elif file_extension == ".docx":
            html_content = convert_docx_to_html(file_content)
        elif file_extension == ".doc":
            html_content = convert_doc_to_html(file_content)
        else:
            return jsonify({"message": "Unsupported file type", "extension": file_extension}), 400

        return jsonify({"html": html_content})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
