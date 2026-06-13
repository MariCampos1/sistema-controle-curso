from flask import Flask, render_template, request, redirect, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO
import os
import uuid


app = Flask(__name__)


# --------------------------------------------------
# CONFIGURAÇÃO DO BANCO DE DADOS
# --------------------------------------------------

database_url = os.environ.get("DATABASE_URL")

# No computador, enquanto você estiver testando,
# será criado um banco SQLite local.
# No Render, será utilizada a DATABASE_URL do Supabase.
if not database_url:
    database_url = "sqlite:///cadastro_local.db"

# Garante que o SQLAlchemy use o driver psycopg2.
if database_url.startswith("postgresql://"):
    database_url = database_url.replace(
        "postgresql://",
        "postgresql+psycopg2://",
        1
    )

app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# Testa a conexão antes de reutilizá-la.
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True
}

db = SQLAlchemy(app)


# --------------------------------------------------
# TABELA DE ALUNOS
# --------------------------------------------------

class Aluno(db.Model):
    __tablename__ = "alunos"

    id = db.Column(
        db.String(36),
        primary_key=True,
        default=lambda: str(uuid.uuid4())
    )

    nome = db.Column(
        db.String(150),
        nullable=False
    )

    curso = db.Column(
        db.String(150),
        nullable=False
    )

    professor = db.Column(
        db.String(150),
        nullable=False
    )

    presente = db.Column(
        db.Boolean,
        nullable=False,
        default=False
    )

    data_presenca = db.Column(
        db.String(10),
        nullable=False,
        default=""
    )

    pagamento = db.Column(
        db.Boolean,
        nullable=False,
        default=False
    )

    alimento = db.Column(
        db.Boolean,
        nullable=False,
        default=False
    )


# Cria a tabela caso ainda não exista.
with app.app_context():
    db.create_all()


# --------------------------------------------------
# TELA PRINCIPAL
# --------------------------------------------------

@app.route("/")
def index():
    curso_filtro = request.args.get("curso", "").strip()
    professor_filtro = request.args.get("professor", "").strip()
    busca = request.args.get("busca", "").strip()

    # Busca todos os cursos antes de aplicar os filtros.
    cursos_resultado = (
        db.session.query(Aluno.curso)
        .filter(Aluno.curso != "")
        .distinct()
        .order_by(Aluno.curso)
        .all()
    )

    cursos = [
        resultado[0]
        for resultado in cursos_resultado
    ]

    # Busca todos os professores antes de aplicar os filtros.
    professores_resultado = (
        db.session.query(Aluno.professor)
        .filter(Aluno.professor != "")
        .distinct()
        .order_by(Aluno.professor)
        .all()
    )

    professores = [
        resultado[0]
        for resultado in professores_resultado
    ]

    consulta = Aluno.query

    if curso_filtro:
        consulta = consulta.filter(
            Aluno.curso == curso_filtro
        )

    if professor_filtro:
        consulta = consulta.filter(
            Aluno.professor == professor_filtro
        )

    if busca:
        consulta = consulta.filter(
            Aluno.nome.ilike(f"%{busca}%")
        )

    alunos = consulta.order_by(
        Aluno.nome.asc()
    ).all()

    total_alunos = len(alunos)

    total_presentes = sum(
        1 for aluno in alunos
        if aluno.presente
    )

    total_pagamentos = sum(
        1 for aluno in alunos
        if aluno.pagamento
    )

    total_alimentos = sum(
        1 for aluno in alunos
        if aluno.alimento
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


# --------------------------------------------------
# ADICIONAR ALUNO
# --------------------------------------------------

@app.route("/adicionar", methods=["POST"])
def adicionar():
    nome = request.form.get("nome", "").strip()
    curso = request.form.get("curso", "").strip()
    professor = request.form.get("professor", "").strip()

    if not nome or not curso or not professor:
        return redirect("/")

    novo_aluno = Aluno(
        nome=nome,
        curso=curso,
        professor=professor,
        presente=False,
        data_presenca="",
        pagamento=False,
        alimento=False
    )

    db.session.add(novo_aluno)
    db.session.commit()

    return redirect("/")


# --------------------------------------------------
# PRESENÇA
# --------------------------------------------------

@app.route("/presenca/<id_aluno>")
def presenca(id_aluno):
    aluno = db.session.get(Aluno, id_aluno)

    if aluno is None:
        return "Aluno não encontrado", 404

    if aluno.presente:
        aluno.presente = False
        aluno.data_presenca = ""
    else:
        aluno.presente = True
        aluno.data_presenca = datetime.now().strftime(
            "%d/%m/%Y"
        )

    db.session.commit()

    return redirect("/")


# --------------------------------------------------
# PAGAMENTO
# --------------------------------------------------

@app.route("/pagamento/<id_aluno>")
def pagamento(id_aluno):
    aluno = db.session.get(Aluno, id_aluno)

    if aluno is None:
        return "Aluno não encontrado", 404

    aluno.pagamento = not aluno.pagamento

    db.session.commit()

    return redirect("/")


# --------------------------------------------------
# ALIMENTO
# --------------------------------------------------

@app.route("/alimento/<id_aluno>")
def alimento(id_aluno):
    aluno = db.session.get(Aluno, id_aluno)

    if aluno is None:
        return "Aluno não encontrado", 404

    aluno.alimento = not aluno.alimento

    db.session.commit()

    return redirect("/")


# --------------------------------------------------
# EXCLUIR ALUNO
# --------------------------------------------------

@app.route("/excluir/<id_aluno>")
def excluir(id_aluno):
    aluno = db.session.get(Aluno, id_aluno)

    if aluno is None:
        return "Aluno não encontrado", 404

    db.session.delete(aluno)
    db.session.commit()

    return redirect("/")


# --------------------------------------------------
# EXPORTAR EXCEL
# --------------------------------------------------

@app.route("/exportar")
def exportar():
    alunos = Aluno.query.order_by(
        Aluno.nome.asc()
    ).all()

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
            aluno.nome,
            aluno.curso,
            aluno.professor,
            "Presente" if aluno.presente else "Ausente",
            aluno.data_presenca,
            "Pago" if aluno.pagamento else "Pendente",
            "Entregue" if aluno.alimento else "Pendente"
        ])

    arquivo = BytesIO()
    wb.save(arquivo)
    arquivo.seek(0)

    return send_file(
        arquivo,
        as_attachment=True,
        download_name="relatorio_alunos.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# --------------------------------------------------
# NOVA CHAMADA
# --------------------------------------------------

@app.route("/nova_chamada")
def nova_chamada():
    alunos = Aluno.query.all()

    for aluno in alunos:
        aluno.presente = False
        aluno.data_presenca = ""

    db.session.commit()

    return redirect("/")


# --------------------------------------------------
# EDITAR PROFESSOR DE UM ALUNO
# --------------------------------------------------

@app.route(
    "/editar_professor/<id_aluno>",
    methods=["GET", "POST"]
)
def editar_professor(id_aluno):
    aluno = db.session.get(Aluno, id_aluno)

    if aluno is None:
        return "Aluno não encontrado", 404

    if request.method == "POST":
        novo_professor = request.form.get(
            "professor",
            ""
        ).strip()

        if novo_professor:
            aluno.professor = novo_professor
            db.session.commit()

        return redirect("/")

    return render_template(
        "editar_professor.html",
        aluno=aluno
    )


# --------------------------------------------------
# RODAR SERVIDOR
# --------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True)