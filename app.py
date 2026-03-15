from flask import Flask, render_template, request, send_file
from docx import Document
import os
import re
import zipfile

app = Flask(__name__)

TEMPLATE_FOLDER = "doc_templates"
OUTPUT_FOLDER = "output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# =========================
# 提取变量
# =========================

def get_variables():

    variables = set()

    for file in os.listdir(TEMPLATE_FOLDER):

        if file.endswith(".docx"):

            doc = Document(os.path.join(TEMPLATE_FOLDER, file))

            for p in doc.paragraphs:

                found = re.findall(r"\{\{(.*?)\}\}", p.text)

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
# 生成文件
# =========================

@app.route("/generate", methods=["POST"])
def generate():

    data = request.form.to_dict()

    generated_files = []

    for template in os.listdir(TEMPLATE_FOLDER):

        if not template.endswith(".docx"):
            continue

        doc = Document(os.path.join(TEMPLATE_FOLDER, template))

        for p in doc.paragraphs:

            for k,v in data.items():

                p.text = p.text.replace("{{"+k+"}}", v)

        output_file = os.path.join(OUTPUT_FOLDER, template)

        doc.save(output_file)

        generated_files.append(output_file)


    # 打包ZIP
    zip_path = os.path.join(OUTPUT_FOLDER, "documents.zip")

    with zipfile.ZipFile(zip_path, "w") as zipf:

        for file in generated_files:

            zipf.write(file, os.path.basename(file))

    return send_file(zip_path, as_attachment=True)


if __name__ == "__main__":
    app.run()
