from flask import Flask, request, jsonify
import pandas as pd
from io import BytesIO
from flask import send_file

app = Flask(__name__)

@app.route("/create-excel", methods=["POST"])
def create_excel():
    data = request.json.get("スケジュール表")

    if not data:
        return jsonify({"error": "データが空です"}), 400

    df = pd.DataFrame(data[1:], columns=data[0])

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(output, download_name="schedule.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
