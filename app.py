from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, send_file, session
from Termo_de_Responsabilidade import gerar_termos
from Script_Termo_Individual import criar_termo_responsabilidade
from termo_devolucao import gerar_termo_devolucao
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'chave-super-secreta'

def caminho_excel(nome):
    return os.path.join(os.path.dirname(__file__), nome)

def carregar_centros_de_custos():
    df = pd.read_excel(caminho_excel('acervo.xlsx'), sheet_name='responsavel')
    return df['ccustos'].unique()

def carregar_nomes_individuais():
    df = pd.read_excel('geral.xlsx', sheet_name='dados')
    return df['Nome'].unique()

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/gerar', methods=['POST'])
def gerar():
    ccusto_escolhido = request.form.get('ccusto')
    try:
        gerar_termos(filtrar_ccusto=ccusto_escolhido)
        flash(f"Termo para {ccusto_escolhido} gerado com sucesso!", 'success')
        return redirect(url_for('centro_custos', ccusto_gerado=ccusto_escolhido))
    except Exception as e:
        flash(f"Ocorreu um erro: {str(e)}", 'error')
        return redirect(url_for('centro_custos'))

@app.route('/download/<path:nome_arquivo>')
def download(nome_arquivo):
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), nome_arquivo, as_attachment=True)

@app.route('/centro-custos')
def centro_custos():
    ccustos = carregar_centros_de_custos()
    ccusto_gerado = request.args.get('ccusto_gerado')
    return render_template('centro_custos.html', ccustos=ccustos, ccusto_gerado=ccusto_gerado)

@app.route('/termos-individuais')
def termos_individuais():
    nomes = carregar_nomes_individuais()
    nome_gerado = request.args.get('nome_gerado')
    return render_template('termos_individuais.html', nomes=nomes, nome_gerado=nome_gerado)

@app.route('/gerar-individual', methods=['POST'])
def gerar_individual():
    nome = request.form.get('nome')
    df = pd.read_excel('geral.xlsx', sheet_name='dados')
    grupo = df[df['Nome'] == nome]
    try:
        criar_termo_responsabilidade(nome, grupo)
        flash(f"Termo de {nome} gerado com sucesso!", 'success')
        return redirect(url_for('termos_individuais', nome_gerado=nome))
    except Exception as e:
        flash(str(e), 'error')
        return redirect(url_for('termos_individuais'))

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        arquivo = request.files.get('arquivo')
        if not arquivo:
            flash('Nenhum arquivo selecionado.', 'error')
            return redirect(url_for('upload'))

        nome_arquivo = arquivo.filename.lower()
        if nome_arquivo.endswith('.xlsx'):
            if 'geral' in nome_arquivo:
                nome_seguro = 'geral.xlsx'
            elif 'acervo' in nome_arquivo:
                nome_seguro = 'acervo.xlsx'
            else:
                flash('Nome de arquivo inválido. Use "geral" ou "acervo" no nome.', 'error')
                return redirect(url_for('upload'))

            arquivo.save(caminho_excel(nome_seguro))
            flash(f'Arquivo {nome_seguro} atualizado com sucesso!', 'success')
        else:
            flash('Por favor, envie um arquivo .xlsx válido.', 'error')
        return redirect(url_for('upload'))

    return render_template('upload.html')

@app.route('/download-planilha/<nome_arquivo>')
def download_planilha(nome_arquivo):
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), nome_arquivo, as_attachment=True)

@app.route("/termo_devolucao", methods=["GET", "POST"])
def termo_devolucao():
    df_nomes = pd.read_excel("geral.xlsx", sheet_name="nomes")
    df_bens = pd.read_excel("geral.xlsx", sheet_name="base")

    nomes = df_nomes['responsavel'].dropna().tolist()
    nome_selecionado = request.form.get("nome") or session.get("nome_selecionado")

    if request.method == "POST":
        if nome_selecionado:
            session['nome_selecionado'] = nome_selecionado

        if "gerar" in request.form:
            bens_ids = session.get("bens_selecionados", [])
            bens_final = df_bens[df_bens['Número Bem'].astype(str).isin(bens_ids)]
            nome_arquivo = gerar_termo_devolucao(nome_selecionado, bens_final.to_dict(orient='records'))
            session.pop("bens_selecionados", None)
            flash("Termo gerado com sucesso!", "success")
            return redirect(url_for("termo_devolucao", nome_gerado=nome_selecionado))

        elif "remover" in request.form:
            numero_remover = request.form.get("remover")
            if numero_remover and numero_remover in session.get("bens_selecionados", []):
                session["bens_selecionados"].remove(numero_remover)
                session.modified = True

        else:
            numero_bem = request.form.get("numero_bem")
            if not numero_bem:
                flash("Digite o número do bem.")
            else:
                bem = df_bens[df_bens['Número Bem'].astype(str) == numero_bem]
                if bem.empty:
                    flash("Bem não encontrado. Verifique o número digitado.")
                else:
                    session.setdefault("bens_selecionados", [])
                    if numero_bem not in session["bens_selecionados"]:
                        session["bens_selecionados"].append(numero_bem)
                        session.modified = True

    bens_ids = session.get("bens_selecionados", [])
    bens_selecionados = df_bens[df_bens['Número Bem'].astype(str).isin(bens_ids)]
    total = bens_selecionados['Valor Atual'].sum()
    nome_gerado = request.args.get("nome_gerado")

    return render_template("termo_devolucao.html",
                           nomes=nomes,
                           nome_selecionado=nome_selecionado,
                           bens_selecionados=bens_selecionados.to_dict(orient='records'),
                           total=total,
                           nome_gerado=nome_gerado)


if __name__ == '__main__':
    app.run(debug=True)