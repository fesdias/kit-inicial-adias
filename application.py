import datetime
import os

from flask import Flask, render_template, request
from docx import Document

app = Flask(__name__)

def dataAtual(doc):
    meses = {'1':"Janeiro", '2':"Fevereiro", '3':"Março", '4':"Abril", '5':"Maio", '6':"Junho", '7':"Julho", '8':"Agosto", '9':"Setembro", '10':"Outubro", '11':"Novembro", '12':"Dezembro"}
    data_atual = datetime.datetime.now()

    x = data_atual.strftime("%d")
    y = str(data_atual.strftime("%m"))

    if doc == "ADjudicia":
        z = int(data_atual.strftime("%Y"))+1
    else:
        z = data_atual.strftime("%Y")

    data = f"Jundiaí, {x} de {meses.get(y)} de {z}."

    return data

def preencher(doc, nomeDocumento, nome2, text):

    document = Document(doc)
    titulo = f"{nome2} - {nomeDocumento}.docx"
    doc = f"{nome2}/{titulo}"

    for paragraph in document.paragraphs:
        if "%data%" in paragraph.text:
            paragraph.text = dataAtual(nomeDocumento)

        if "%nome%" in paragraph.text:
            if nomeDocumento == "Contrato":
                paragraph.text = nome2.upper()
            else:
                paragraph.text = nome2

        for i in document.paragraphs:

            inline = i.runs
            for j in range(len(inline)):

                if inline[j].text == "cabecalhoaqui":
                    inline[j].text = text

                    try:
                        if inline[j-1].text == "%" or inline[j-1].text == "ü":
                            inline[j-1].text = ""

                        if inline[j+1].text == "%" or inline[j+1].text == "ü":
                            inline[j+1].text = ""

                    except:
                        pass


    document.save(doc)

def preencherTermo(nome, cpf, rg, end, cidade, cep):

    document = Document("Intranet/Kitinicial/termo.docx")
    titulo = f"{nome} - Termo de Representação.docx"
    doc = f"{nome}/{titulo}"

    for paragraph in document.paragraphs:

        inline = paragraph.runs
        for j in range(len(inline)):

            if inline[j].text == "nomecompleto":
                inline[j].text = nome
                inline[j-1].text = ""
                inline[j+1].text = ""

            if inline[j].text == "numerocpf":
                inline[j].text = cpf
                inline[j-1].text = ""
                inline[j+1].text = ""

            if inline[j].text == "numerorg":
                inline[j].text = rg
                inline[j-1].text = ""
                inline[j+1].text = ""

            if inline[j].text == "enderecocompleto":
                inline[j].text = end
                inline[j-1].text = ""
                inline[j+1].text = ""

            if inline[j].text == "nomecidade":
                inline[j].text = cidade
                inline[j-1].text = ""
                inline[j+1].text = ""

            if inline[j].text == "numerocep":
                inline[j].text = cep
                inline[j-1].text = ""
                inline[j+1].text = ""

        document.save(doc)

@app.route('/')
def index():
    return render_template("Acervo.html")

    if __name__ == "__main__":
        port = int(os.environ.get("PORT", 5000))
        app.run(host='0.0.0.0', port=port)

@app.route("/enviar", methods = ["POST"])
def inserir():

    nome = request.form.get("nome")
    cpf = request.form.get("CPF")
    rg = request.form.get("RG")
    tel = request.form.get("tel")
    nacao = request.form.get("nacao")
    estCiv = request.form.get("estCiv")
    prof = request.form.get("prof")
    nasc = request.form.get("nasc")
    end = request.form.get("end")
    city = request.form.get("city")
    state = request.form.get("state")
    cep = request.form.get("cep")
    sexo = request.form.get("sexo")

    tipo = request.form.get("tipo")
    anda = request.form.get("andamento")
    status = request.form.get("status")

    nome = nome.upper()
    nome2 = nome.title()
    nacao = nacao.lower()
    estCiv = estCiv.lower()
    prof = prof.lower()
    end = end.title()
    city = city.title()

    if len(rg) == 8:
        rg = f"{rg[0]}{rg[1]}.{rg[2]}{rg[3]}{rg[4]}.{rg[5]}{rg[6]}{rg[7]}"

    if len(rg) == 9:
        rg = f"{rg[0]}{rg[1]}.{rg[2]}{rg[3]}{rg[4]}.{rg[5]}{rg[6]}{rg[7]}-{rg[8]}"

    if len(rg) == 10:
        rg = f"{rg[0]}{rg[1]}.{rg[2]}{rg[3]}{rg[4]}.{rg[5]}{rg[6]}{rg[7]}-{rg[8]}{rg[9]}"

    cpf = f"{cpf[0]}{cpf[1]}{cpf[2]}.{cpf[3]}{cpf[4]}{cpf[5]}.{cpf[6]}{cpf[7]}{cpf[8]}-{cpf[9]}{cpf[10]}"
    text = f"{nome}, {nacao}, {estCiv}, {prof}, portador do RG sob nº {rg}, e do CPF {cpf}, nascido em {nasc}, residente e domiciliado à {end}, {city}/{state} - CEP: {cep}"

    try:
        os.mkdir(nome2)
        pass
    except Exception as e:
        print("Já existe a pasta")
        pass

    preencher("Intranet/Kitinicial/procuracao.docx", "Procuração", nome2, text)
    preencher("Intranet/Kitinicial/pobreza.docx", "Pobreza", nome2, text)
    preencher("Intranet/Kitinicial/adjudicia.docx", "ADjudicia", nome2, text)
    preencherTermo(nome2, cpf, rg, end, city, cep)

    if (sexo == "f"):
        preencher("Intranet/Kitinicial/contrato-fem.docx", "Contrato", nome2, text)

    else:
        preencher("Intranet/Kitinicial/contrato-masc.docx", "Contrato", nome2, text)

    path = os.getcwd()
    link = f"{path}/{nome2}"

    concluido = f"Encontre os documentos de {nome2} em"

    return render_template("confirmacao.html", text = concluido, cab = text, link = link)

@app.route("/buscar", methods = ["POST", "GET"])
def buscar():

    buscar = request.form.get("buscar")

    clientes = {}
    return render_template("naoencontrado.html")
