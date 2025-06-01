from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
import pandas as pd
from flask_cors import CORS
from convert import PROCESS_XLXS
import os
import ast
import rotate
import time

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # Cho phép mọi nguồn truy cập

@app.route("/")
def index():
    return send_file("frontend/index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "Không có file"}), 400

    try:
        file = request.files['file']
        folder_path = request.form.get("path_folder")
        os.makedirs("data/upload", exist_ok=True)
        save_path = os.path.join("data/upload", "samples.xlsx")
        file.save(save_path)
        data = PROCESS_XLXS(save_path)
        df = data.process("SinoNom_OCR")
        if df is None:
            return jsonify({"error": "Không tìm thấy cột SinoNom_OCR"}), 400
        
        df.to_csv("info.csv", index=False)
        rotate.handle_rotate(path_folder=folder_path)
        return jsonify({"message": "File đã lưu thành công", "path": save_path})

    except Exception as e:
        print("❌ Lỗi server:", str(e))
        return jsonify({"error": str(e)}), 500


@app.route("/sequence", methods=["GET"])
def get_sequence():
    frame = pd.read_csv("info.csv")
    char = request.args.get("char", "")
    row = frame[frame["Name"] == char]

    if row.empty:
        return jsonify([])

    try:
        config_str = row["Config"].values[0]
        parsed = ast.literal_eval(config_str)
        # Convert set → list
        for item in parsed:
            if isinstance(item.get('Font'), set):
                item['Font'] = list(item['Font'])
        return jsonify(parsed)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Đọc 1 lần khi server khởi động
try:
    df_dict = pd.read_excel("dict/QuocNgu_SinoNom_Dic.xlsx")
    suggest_map = df_dict.groupby("QuocNgu")["SinoNom"].apply(list).to_dict()
except Exception as e:
    print("❌ Lỗi load từ điển:", e)
    suggest_map = {}

@app.route("/suggest", methods=["GET"])
def suggest():
    char = request.args.get("char", "")
    try:
        suggestions = suggest_map.get(char, [])
        return jsonify({
            "input": char,
            "suggestions": suggestions
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
