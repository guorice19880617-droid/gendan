from flask import Flask, render_template, request, send_file
import openpyxl
import os
import zipfile
import re

app = Flask(__name__)

TEMPLATE_FOLDER = "excel_templates"
OUTPUT_FOLDER = "output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# =========================
# 扫描Excel模板变量
# =========================
def get_variables():

    variables = set()

    for file in os.listdir(TEMPLATE_FOLDER):

        if file.endswith(".xlsx"):

            path = os.path.join(TEMPLATE_FOLDER, file)

            wb = openpyxl.load_workbook(path)

            sheet = wb.active

            for row in sheet.iter_rows():

                for cell in row:

                    if cell.value:

                        found = re.findall(r"\{\{(.*?)\}\}", str(cell.value))

                        for f in found:

                            variables.add(f)

    return list(variables)


# =========================
# 首页
# =========================
@app.route("/")
def index():

    variables = get_variables()

    return render_template("index.html", variables=variables)


# =========================
# 生成Excel文件
# =========================
@app.route("/generate", methods=["POST"])
def generate():

    data = request.form.to_dict()

    generated_files = []

    for template in os.listdir(TEMPLATE_FOLDER):

        if not template.endswith(".xlsx"):
            continue

        template_path = os.path.join(TEMPLATE_FOLDER, template)

        wb = openpyxl.load_workbook(template_path)

        sheet = wb.active

        for row in sheet.iter_rows():

            for cell in row:

                if cell.value:

                    for key, value in data.items():

                        cell.value = str(cell.value).replace("{{"+key+"}}", value)

        output_path = os.path.join(OUTPUT_FOLDER, template)

        wb.save(output_path)

        generated_files.append(output_path)


    # 打包ZIP
    zip_path = os.path.join(OUTPUT_FOLDER, "documents.zip")

    with zipfile.ZipFile(zip_path, "w") as zipf:

        for file in generated_files:

            zipf.write(file, os.path.basename(file))

    return send_file(zip_path, as_attachment=True)


# =========================
# 启动程序（Render必须）
# =========================
if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(host="0.0.0.0", port=port)
