from flask import Flask, request, jsonify
import json
from flask_cors import CORS 

app = Flask(__name__)
CORS(app) 

# Load từ điển
with open("SinoNom_Similar_Dic_v2.json", encoding="utf-8") as f:
    raw = json.load(f)

# Chuyển từ pandas-export sang dict thực
input_chars = list(raw["Input Character"].values())
similar_lists = [eval(x) for x in raw["Top 20 Similar Characters"].values()]
similar_dict = dict(zip(input_chars, similar_lists))

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
