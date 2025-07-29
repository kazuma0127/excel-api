from flask import Flask, request, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def home():
    return 'Excel API is running!'

@app.route('/create-excel', methods=['POST'])
def create_excel():
    data = request.get_json()
    title = data.get("title", "output")
    rows = data.get("rows")

    if not rows:
        return jsonify({"error": "データが空です"}), 400

    try:
        df = pd.DataFrame(rows[1:], columns=rows[0])
        filename = f"{title}.xlsx"
        df.to_excel(filename, index=False)

        return send_file(
            filename,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
