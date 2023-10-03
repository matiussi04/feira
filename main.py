from flask import Flask, render_template, request, redirect
from json2docx.JsonToDocx import JsonToDocx
import os

app = Flask(__name__)


@app.route("/")
def home():
    return render_template("home.html")


@app.route('/data', methods=['POST'])
def data():
    editor = request.form['editor']

    data = {
        "fields": [
            {
                "key": "content",
                "value": editor,
            },
        ]
    }

    output_docx = os.getcwd() + "/new_document.docx"
    output_pdf = os.getcwd() + "/new_document.pdf"
    json2docx = JsonToDocx("template.docx", data, output_docx, output_pdf)
    json2docx.convert()

    return redirect('/')


if __name__ == '__main__':
    app.run(debug=True)
