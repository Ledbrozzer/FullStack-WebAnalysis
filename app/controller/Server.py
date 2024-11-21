#venv\Scripts\activate  --Ativar ambiente virtual
#pip install flask pandas openpyxl werkzeug
#pip show flask  --Verificar se flask foi installed right
#pip install xlsxwriter
#pip install lxml
#pip install openpyxl
#pip install openpyxl==3.0.10
#python app/controller/Server.py
from flask import Flask, request, jsonify, send_file, redirect, url_for, render_template_string
from werkzeug.utils import secure_filename
import pandas as pd
import io
from openpyxl import Workbook
app = Flask(__name__, static_url_path='/static', static_folder='../view')

@app.route('/')
def index():
    return render_template_string(open('app/view/index.html', encoding='utf-8').read())

@app.route('/login', methods=['POST'])
def login():
    usuario = request.form['usuario']
    senha = request.form['senha']
    #Credenciais-exemplo
    users = {"Jose Mario": "1234", "Saulo": "5678", "Gesse": "9123"}
    
    if usuario in users and users[usuario] == senha:
        return redirect(url_for('app_page'))
    else:
        return 'Credenciais inválidas!'

@app.route('/app')
def app_page():
    return render_template_string(open('app/view/App.html', encoding='utf-8').read())

@app.route('/process_csv', methods=['POST'])
def process_csv():
    try:
        data = request.json['data']
        df = pd.read_csv(io.StringIO(data), delimiter=';', on_bad_lines='skip')
        #Remov colunas indesejadas
        colunas_para_excluir = ["Requisição", "Hora Abast.", "Obs.", "Abast. Externo", "Combustível"]
        colunas_existentes = [col for col in colunas_para_excluir if col in df.columns]
        if colunas_existentes:
            df = df.drop(columns=colunas_existentes)
        #Converter colunas p/n°
        df['Km Rodados'] = pd.to_numeric(df['Km Rodados'], errors='coerce')
        df['Litros'] = pd.to_numeric(df['Litros'].astype(str).str.replace(',', '').str[:-2], errors='coerce')
        df['Vlr. Total'] = pd.to_numeric(df['Vlr. Total'].astype(str).str.replace(',', ''), errors='coerce')
        df['Horim. Equip.'] = pd.to_numeric(df['Horim. Equip.'], errors='coerce')
        #Calcular colunas adicionais
        df['Km por Litro'] = df['Km Rodados'] / df['Litros']
        df['Lucro'] = df['Km Rodados'] - df['Vlr. Total']
        df['Horim por Litro'] = df['Horim. Equip.'] / df['Litros']
        #Convertendo results p/ HTML
        result_html = df.to_html()
        return jsonify(result=result_html)
    except Exception as e:
        return jsonify(error=str(e)), 400
@app.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        data = request.json['data']
        #Envolver string em 1objeto StringIO
        df = pd.read_html(io.StringIO(data))[0]
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados Filtrados')
        #Reposicionar o ponteiro no início
        output.seek(0)
        #Enviar com extensão correta
        response = send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='dados_filtrados.xlsx'  #Usar extensão .xlsx
        )
        return response
    except Exception as e:
        return jsonify(error=str(e)), 400

if __name__ == '__main__':
    app.run(debug=True, port=5001)