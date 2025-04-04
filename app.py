from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from Termo_de_Responsabilidade import gerar_termos
from Script_Termo_Individual import criar_termo_responsabilidade
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'chave-super-secreta'

# Lê os centros de custos uma vez (ou você pode ler em tempo real)
def carregar_centros_de_custos():
    diretorio_atual = os.path.dirname(__file__)
    caminho_excel = os.path.join(diretorio_atual, 'acervo.xlsx')
    df_responsavel = pd.read_excel(caminho_excel, sheet_name='responsavel')
    return df_responsavel['ccustos'].unique()

def carregar_nomes_individuais():
    df = pd.read_excel('geral.xlsx', sheet_name='dados')
    return df['Nome'].unique()

@app.route('/')
def home():
    return render_template('index.html')  # ← essa será a home com os botões

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
    pasta = os.path.dirname(os.path.abspath(__file__))
    return send_from_directory(pasta, nome_arquivo, as_attachment=True)

@app.route('/centro-custos')
def centro_custos():
    ccustos = carregar_centros_de_custos()
    ccusto_gerado = request.args.get('ccusto_gerado')
    return render_template('centro_custos.html', ccustos=ccustos, ccusto_gerado=ccusto_gerado)

@app.route('/termos-individuais')
def termos_individuais():
    nomes = carregar_nomes_individuais()
    return render_template('termos_individuais.html', nomes=nomes)

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
from werkzeug.utils import secure_filename

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

            caminho = os.path.join(os.path.dirname(__file__), nome_seguro)
            arquivo.save(caminho)
            flash(f'Arquivo {nome_seguro} atualizado com sucesso!', 'success')
            return redirect(url_for('upload'))
        else:
            flash('Por favor, envie um arquivo .xlsx válido.', 'error')
            return redirect(url_for('upload'))

    return render_template('upload.html')

@app.route('/download-planilha/<nome_arquivo>')
def download_planilha(nome_arquivo):
    pasta = os.path.dirname(os.path.abspath(__file__))
    return send_from_directory(pasta, nome_arquivo, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
