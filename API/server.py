# from flask import Flask, request, jsonify, send_file, send_from_directory
# from openpyxl import load_workbook
# import pandas as pd
# from flask_cors import CORS
# from convert import PROCESS_XLXS
# import os
# import ast
# import rotate
# import time
# import pathlib
# from handle_html import extract_text_with_color, write_colored_excel_from_chunks


# app = Flask(__name__, 
#     static_folder='../frontend',  # Set static folder to parent frontend directory
#     static_url_path='/frontend'   # URL path for static files
# )
# CORS(app, resources={r"/*": {"origins": "*"}})

# @app.route("/")
# def index():
#     return send_from_directory('../frontend', 'index.html')


# @app.route("/upload", methods=["POST"])
# def upload_file():
#     if 'file' not in request.files:
#         return jsonify({"error": "Không có file"}), 400

#     try:
#         file = request.files['file']
#         folder_path = request.form.get("path_folder")
#         os.makedirs("data/upload", exist_ok=True)
#         save_path = os.path.join("data/upload", "samples.xlsx")
#         file.save(save_path)
#         data = PROCESS_XLXS(save_path)
#         df = data.process("SinoNom_OCR")
#         if df is None:
#             return jsonify({"error": "Không tìm thấy cột SinoNom_OCR"}), 400
        
#         df.to_csv("info.csv", index=False)
#         rotate.handle_rotate(path_folder=folder_path)
#         return jsonify({"message": "File đã lưu thành công", "path": save_path})

#     except Exception as e:
#         print("❌ Lỗi server:", str(e))
#         return jsonify({"error": str(e)}), 500


# @app.route("/sequence", methods=["GET"])
# def get_sequence():
#     frame = pd.read_csv("info.csv")
#     char = request.args.get("char", "")
#     row = frame[frame["Name"] == char]

#     if row.empty:
#         return jsonify([])

#     try:
#         config_str = row["Config"].values[0]
#         parsed = ast.literal_eval(config_str)
#         # Convert set → list
#         for item in parsed:
#             if isinstance(item.get('Font'), set):
#                 item['Font'] = list(item['Font'])
#         return jsonify(parsed)
#     except Exception as e:
#         return jsonify({"error": str(e)}), 500

# # Đọc 1 lần khi server khởi động
# try:
#     df_dict = pd.read_excel("dict/QuocNgu_SinoNom_Dic.xlsx")
#     suggest_map = df_dict.groupby("QuocNgu")["SinoNom"].apply(list).to_dict()
# except Exception as e:
#     print("❌ Lỗi load từ điển:", e)
#     suggest_map = {}

# @app.route("/suggest", methods=["GET"])
# def suggest():
#     char = request.args.get("char", "")
#     try:
#         suggestions = suggest_map.get(char, [])
#         return jsonify({
#             "input": char,
#             "suggestions": suggestions
#         })
#     except Exception as e:
#         return jsonify({"error": str(e)}), 500

# @app.route("/save_table", methods=["POST"])
# def save_table():
#     data = request.get_json()
#     table_html = data.get("table_html", "")
#     path_excel = data.get("path_excel", "result.xlsx")  # ✅ lấy path từ client
    
#     if not table_html:
#         return jsonify({"error": "Không có bảng gửi lên"}), 400

#     try:
#         # Lưu HTML tạm thời để xử lý
#         os.makedirs("data", exist_ok=True)
#         html_path = "data/result.html"
#         with open(html_path, "w", encoding="utf-8") as f:
#             f.write(table_html)

#         # Gọi hàm xử lý
#         data = extract_text_with_color(html_path)
#         write_colored_excel_from_chunks(data, output_path=path_excel)

#         return jsonify({
#             "status": "ok",
#             "message": f"✅ Đã lưu bảng vào {path_excel}"
#         })
#     except Exception as e:
#         return jsonify({"error": str(e)}), 500

# if __name__ == "__main__":
#     app.run(debug=True)


from flask import Flask, request, jsonify, send_file, send_from_directory
from openpyxl import load_workbook
import pandas as pd
from flask_cors import CORS
from convert import PROCESS_XLXS
import os
import ast
import rotate
import time
import pathlib
from handle_html import extract_text_with_color, write_colored_excel_from_chunks


# app = Flask(__name__)
# CORS(app, resources={r"/*": {"origins": "*"}})  # Cho phép mọi nguồn truy cập

app = Flask(__name__, 
    static_folder='../frontend',  # Set static folder to parent frontend directory
    static_url_path='/frontend'   # URL path for static files
)
CORS(app, resources={r"/*": {"origins": "*"}})

# Global variable để lưu đường dẫn ảnh
current_image_path = ""

# @app.route("/")
# def index():
#     return send_file("frontend/index.html")

@app.route("/")
def index():
    return send_from_directory('../frontend', 'index.html')


@app.route("/upload", methods=["POST"])
def upload_file():
    global current_image_path
    if 'file' not in request.files:
        return jsonify({"error": "Không có file"}), 400

    try:
        file = request.files['file']
        folder_path = request.form.get("path_folder")
        current_image_path = folder_path  # Lưu đường dẫn ảnh
        
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

# ✅ Thêm endpoint để serve ảnh
@app.route("/image/<filename>")
def serve_image(filename):
    global current_image_path
    try:
        if not current_image_path or not os.path.exists(current_image_path):
            return jsonify({"error": "Đường dẫn ảnh không tồn tại"}), 404
            
        image_path = os.path.join(current_image_path, filename)
        if not os.path.exists(image_path):
            return jsonify({"error": f"Ảnh {filename} không tồn tại"}), 404
            
        return send_file(image_path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/save_table", methods=["POST"])
def save_table():
    data = request.get_json()
    table_html = data.get("table_html", "")
    path_excel = data.get("path_excel", "result.xlsx")  # ✅ lấy path từ client
    
    if not table_html:
        return jsonify({"error": "Không có bảng gửi lên"}), 400

    try:
        # Lưu HTML tạm thời để xử lý
        os.makedirs("data", exist_ok=True)
        html_path = "data/result.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(table_html)

        # Gọi hàm xử lý
        data = extract_text_with_color(html_path)
        write_colored_excel_from_chunks(data, output_path=path_excel)

        return jsonify({
            "status": "ok",
            "message": f"✅ Đã lưu bảng vào {path_excel}"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)