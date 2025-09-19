from flask import Flask, request, render_template
import msoffcrypto
import io
import pandas as pd

app = Flask(__name__)

PASSWORD = "m3@9B$#*52K&692v"

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "没有文件被上传", 400
    f = request.files['file']
    filename = f.filename
    if not (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls')):
        return "请上传 Excel 文件 .xlsx 或 .xls", 400

    try:
        office_file = msoffcrypto.OfficeFile(f.stream)
        office_file.load_key(password=PASSWORD)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        df = pd.read_excel(decrypted)

        carton_col = df.columns[2]
        code_col = df.columns[5]

        df_unique = df.drop_duplicates(subset=carton_col)
        counts = df_unique[code_col].value_counts().sort_index()

                # 映射编号到城市名
        city_map = {
            850: "WEST VALLEY", 855: "WEST VALLEY",
            940: "SAN FRANCISCO", 949: "SAN FRANCISCO",
            829: "SALT LAKE CITY", 840: "SALT LAKE CITY",
            920: "SAN DIEGO",
            890: "LAS VEGAS",
            932: "BAKERSFIELD",
            980: "WA", 982: "WA", 983: "WA",
            970: "OR"
        }

        result_lines = []

        for code, count in counts.items():
            try:
                code_int = int(code)
            except:
                continue  # 忽略无法转换为数字的编号

            city = city_map.get(code_int)
            if city:
                result_lines.append(f"{city} {code_int}-{count}")

        # 按城市和编号排序（可选）
        result_lines.sort()

        # 返回居中格式 HTML
        return f"""
        <div style="text-align: center; font-family: monospace; white-space: pre-line;">
        {'\n'.join(result_lines)}
        </div>
        """


    except Exception as e:
        return f"❌ 错误：{str(e)}", 500

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))   # 从环境变量获取端口，默认5000
    app.run(debug=False, host='0.0.0.0', port=port)  # 监听所有IP



