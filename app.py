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
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
