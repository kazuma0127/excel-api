from flask import Flask, request, jsonify, send_file
import pandas as pd
from docx import Document
from pptx import Presentation
import os

app = Flask(__name__)

@app.route('/')
def home():
    return 'Multi-file API is running!'

@app.route('/create-excel', methods=['POST'])
def create_excel():
    data = request.get_json()
    title = data.get("title", "output")
    rows = data.get("rows")

    if not rows:
        return jsonify({"error": "Excelのデータが空です"}), 400

    try:
        df = pd.DataFrame(rows[1:], columns=rows[0])
        filename = f"{title}.xlsx"
        df.to_excel(filename, index=False)
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/create-word', methods=['POST'])
def create_word():
    data = request.get_json()
    title = data.get("title", "output")
    paragraphs = data.get("paragraphs", [])

    try:
        doc = Document()
        for para in paragraphs:
            doc.add_paragraph(para)
        filename = f"{title}.docx"
        doc.save(filename)
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/create-ppt', methods=['POST'])
def create_ppt():
    data = request.get_json()
    title = data.get("title", "output")
    slides = data.get("slides", [])

    try:
        prs = Presentation()
        for slide_content in slides:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = slide_content.get("title", "")
            slide.placeholders[1].text = slide_content.get("content", "")
        filename = f"{title}.pptx"
        prs.save(filename)
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)