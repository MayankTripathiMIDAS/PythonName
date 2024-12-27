from flask import Flask, request, jsonify
import requests
from io import BytesIO
from docx import Document
import fitz
from flask_cors import CORS  # Import CORS

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def convert_pdf_to_html(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    html_content = ""
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        html_content += page.get_text("html")  # Extract HTML for each page

    return html_content

def convert_docx_to_html(docx_bytes):
    # Open the .docx file using python-docx
    document = Document(BytesIO(docx_bytes))
    
    # HTML content will be stored here
    html_content = "<html><body>"

    # Iterate through the paragraphs in the document
    for paragraph in document.paragraphs:
        paragraph_html = "<p>"
        
        # Check each run in the paragraph (runs are chunks of text with the same style)
        for run in paragraph.runs:
            paragraph_html += run.text
        
        paragraph_html += "</p>"
        html_content += paragraph_html

    html_content += "</body></html>"
    
    return html_content

@app.route('/highlight-file', methods=['POST'])
def highlight_file():
    try:
        # Parse request data
        data = request.get_json()
        item_id = data.get("itemId")
        token = data.get("token")
        file_extension = data.get("fileExtension", "")

        if not item_id or not token:
            return jsonify({"error": "itemId and token are required"}), 400
        
        # Construct the file content URL
        file_content_url = f"https://graph.microsoft.com/v1.0/sites/midasconsultingmgmt.sharepoint.com,6ca0fab8-2a87-4e15-a144-d87634dcb569,1b3d5672-7447-4188-982e-126402613a10/drives/b!uPqgbIcqFU6hRNh2NNy1aXJWPRtHdIhBmC4SZAJhOhBCF-UF6RIYQ7WCbzH_wEcf/items/{item_id}/content"
        print(f"File content URL: {file_content_url}")
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(file_content_url, headers=headers)
        
        if response.status_code != 200:
            return jsonify({"error": "Failed to fetch file", "details": response.text}), response.status_code

        # Get the file content as bytes
        file_content = response.content

        # Process the file without saving to disk
        if file_extension == ".docx" or file_extension == ".doc":
            html_content = convert_docx_to_html(file_content)

            # Return the HTML content as a response
            return jsonify({"html": html_content})
        
        if file_extension == ".pdf":
            html_content = convert_pdf_to_html(file_content)

            # Return the HTML content as a response
            return jsonify({"html": html_content})

        # If file extension is not .docx or .pdf, return a message
        return jsonify({"message": "File downloaded successfully, but no highlighting applied for non-docx/pdf files"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
