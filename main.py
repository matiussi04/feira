from flask import Flask, render_template, request, redirect
from json_to_docx.JsonToDocx import JsonToDocx
import os

app = Flask(__name__)


@app.route("/")
def home():
    return render_template("home.html")


@app.route('/data', methods=['POST'])
def data():
    coordenador = request.form['coordenador']
    titulo = request.form['titulo']
    aluno = request.form['aluno']
    orientador = request.form['orientador']
    conteudo1 = request.form['conteudo1']
    conteudo2 = request.form['conteudo2']

    data = {
        "fields": [
            {
                "key": "coordenador",
                "value": coordenador,
            },
            {
                "key": "titulo",
                "value": titulo,
            },
            {
                "key": "aluno",
                "value": aluno,
            },
            {
                "key": "orientador",
                "value": orientador,
            },
            {
                "key": "conteudo1",
                "value": editor,
            },
            {
                "key": "conteudo2",
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
