from flask import Flask, request, jsonify
import json
import pandas as pd
from flask_cors import CORS 

import ast



app = Flask(__name__)
CORS(app) 

# Load từ điển
with open("SinoNom_Similar_Dic_v2.json", encoding="utf-8") as f:
    raw = json.load(f)

# Chuyển từ pandas-export sang dict thực
input_chars = list(raw["Input Character"].values())
similar_lists = [eval(x) for x in raw["Top 20 Similar Characters"].values()]
similar_dict = dict(zip(input_chars, similar_lists))

frame = pd.read_csv("info.csv")

@app.route("/sequence", methods=["GET"])
def getSequence():
    char = request.args.get("char", "")
    lst = frame.Config[frame.Name == char]
    if lst.empty:
        return jsonify([])

    try:
        # Parse chuỗi từ CSV thành list object an toàn
        parsed = ast.literal_eval(lst.values[0])
        # Chuyển set → list nếu có Font là {'Arial'}
        for item in parsed:
            if isinstance(item.get('Font'), set):
                item['Font'] = list(item['Font'])
        return jsonify(parsed)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route("/suggest", methods=["GET"])
def suggest():
    char = request.args.get("char", "")
    suggestions = similar_dict.get(char, [])
    return jsonify({
        "input": char,
        "suggestions": suggestions
    })

if __name__ == "__main__":
    app.run(debug=True)
