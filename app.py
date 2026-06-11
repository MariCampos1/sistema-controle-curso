from flask import Flask, render_template, request, redirect, send_file
from datetime import datetime
from openpyxl import Workbook
import json
import os
import uuid

app = Flask(__name__)

ARQUIVO = "dados.json"


# -------------------------
# CARREGAR DADOS
# -------------------------
def carregar_dados():
    if os.path.exists(ARQUIVO):
        with open(ARQUIVO, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


# -------------------------
# SALVAR DADOS
# -------------------------
def salvar_dados(alunos):
    with open(ARQUIVO, "w", encoding="utf-8") as f:
        json.dump(alunos, f, indent=4, ensure_ascii=False)


# -------------------------
# TELA PRINCIPAL
# -------------------------
@app.route("/")
def index():
    alunos = carregar_dados()

    alunos = sorted(alunos, key=lambda x: x.get("nome", "").lower())

    curso_filtro = request.args.get("curso", "")

    professor_filtro = request.args.get("professor", "")

    busca = request.args.get("busca", "").strip()

    cursos = sorted(list(set(
        aluno.get("curso", "")
        for aluno in alunos
        if aluno.get("curso")
    )))

    professores = sorted(list(set(
    aluno.get("professor", "")
    for aluno in alunos
    if aluno.get("professor")
)))

    if curso_filtro:
        alunos = [
            aluno for aluno in alunos
            if aluno.get("curso") == curso_filtro
        ]

    if professor_filtro:
        alunos = [
            aluno for aluno in alunos
            if aluno.get("professor") == professor_filtro
        ]

    if busca:
        alunos = [
            aluno for aluno in alunos
            if busca.lower() in aluno.get("nome", "").lower()
        ]    

    total_alunos = len(alunos)

    total_presentes = sum(
        1 for aluno in alunos
        if aluno.get("presente", False)
    )

    total_pagamentos = sum(
        1 for aluno in alunos
        if aluno.get("pagamento", False)
    )

    total_alimentos = sum(
        1 for aluno in alunos
        if aluno.get("alimento", False)
    )

    return render_template(
        "index.html",
        alunos=alunos,
        cursos=cursos,
        curso_filtro=curso_filtro,
        professores=professores,
        professor_filtro=professor_filtro,
        busca=busca,
        total_alunos=total_alunos,
        total_presentes=total_presentes,
        total_pagamentos=total_pagamentos,
        total_alimentos=total_alimentos

    )


# -------------------------
# ADICIONAR ALUNO
# -------------------------
@app.route("/adicionar", methods=["POST"])
def adicionar():
    alunos = carregar_dados()

    nome = request.form["nome"]
    curso = request.form["curso"]
    professor = request.form["professor"]

    novo_aluno = {
        "id": str(uuid.uuid4()),
        "nome": nome,
        "curso": curso,
        "professor": professor,
        "presente": False,
        "data_presenca": "",
        "pagamento": False,
        "alimento": False
    }

    alunos.append(novo_aluno)
    salvar_dados(alunos)

    return redirect("/")


# -------------------------
# PRESENÇA
# -------------------------
@app.route("/presenca/<id_aluno>")
def presenca(id_aluno):
    alunos = carregar_dados()

    for aluno in alunos:
        if aluno.get("id") == id_aluno:
            if aluno.get("presente", False):
                aluno["presente"] = False
                aluno["data_presenca"] = ""
            else:
                aluno["presente"] = True
                aluno["data_presenca"] = datetime.now().strftime("%d/%m/%Y")
            break

    salvar_dados(alunos)
    return redirect("/")

# -------------------------
# PAGAMENTO
# -------------------------
@app.route("/pagamento/<id_aluno>")
def pagamento(id_aluno):
    alunos = carregar_dados()

    for aluno in alunos:
        if aluno.get("id") == id_aluno:
            aluno["pagamento"] = not aluno.get("pagamento", False)
            break

    salvar_dados(alunos)
    return redirect("/")

# -------------------------
# ALIMENTO
# -------------------------
@app.route("/alimento/<id_aluno>")
def alimento(id_aluno):
    alunos = carregar_dados()

    for aluno in alunos:
        if aluno.get("id") == id_aluno:
            aluno["alimento"] = not aluno.get("alimento", False)
            break

    salvar_dados(alunos)
    return redirect("/")

@app.route("/excluir/<id_aluno>")
def excluir(id_aluno):
    alunos = carregar_dados()

    alunos = [
        aluno for aluno in alunos
        if aluno.get("id") != id_aluno
    ]

    salvar_dados(alunos)
    return redirect("/")

# -------------------------
# RODAR SERVIDOR
# -------------------------
@app.route("/exportar")
def exportar():
    alunos = carregar_dados()

    wb = Workbook()
    ws = wb.active
    ws.title = "Alunos"

    ws.append([
        "Nome",
        "Curso",
        "Professor",
        "Presença",
        "Data da Presença",
        "Pagamento",
        "Alimento"
    ])

    for aluno in alunos:
        ws.append([
            aluno.get("nome", ""),
            aluno.get("curso", ""),
            aluno.get("professor", ""),
            "Presente" if aluno.get("presente", False) else "Ausente",
            aluno.get("data_presenca", ""),
            "Pago" if aluno.get("pagamento", False) else "Pendente",
            "Entregue" if aluno.get("alimento", False) else "Pendente"
        ])

    nome_arquivo = "relatorio_alunos.xlsx"
    wb.save(nome_arquivo)

    return send_file(nome_arquivo, as_attachment=True)

@app.route("/nova_chamada")
def nova_chamada():
    alunos = carregar_dados()

    for aluno in alunos:
        aluno["presente"] = False
        aluno["data_presenca"] = ""

    salvar_dados(alunos)

    return redirect("/")

if __name__ == "__main__":
    app.run(debug=True)

    