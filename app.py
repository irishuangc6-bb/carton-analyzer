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

        result_lines = ["各编号对应的不重复 Carton 数量："]
        for code, count in counts.items():
            result_lines.append(f"{code} ：{count}")
        return "<pre>" + "\n".join(result_lines) + "</pre>"

    except Exception as e:
        return f"❌ 错误：{str(e)}", 500
