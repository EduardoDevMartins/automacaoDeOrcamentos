from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
import pandas as pd

app = Flask(__name__)

# Configuração para o diretório de uploads
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Caminho para a planilha Excel de entrada
input_excel_path = 'automacao.xlsx'

# Caminho para a planilha Excel de saída (onde os dados serão salvos)
output_excel_path = 'feedback_orcamentos.xlsx'

# Carregando os dados das planilhas existentes
df_environments = pd.read_excel(input_excel_path, sheet_name='Ambientes_Problemas')
df_budget_options = pd.read_excel(input_excel_path, sheet_name='Problemas_Orcamento')




@app.route('/')
def index():
    return render_template('index.html', environments=df_environments['Ambiente'])


@app.route('/problems', methods=['POST'])
def problems():
    environment = request.form['environment']
    problems_list = df_environments[df_environments['Ambiente'] == environment]['Problema'].tolist()
    return render_template('problems.html', problems=problems_list)


@app.route('/budget', methods=['POST'])
def budget():
    problem = request.form['problem']
    selected_budget = df_budget_options[df_budget_options['Problema'] == problem]['Orçamento'].iloc[0]
    return render_template('budget.html', budget=selected_budget)


@app.route('/confirm_budget', methods=['POST'])
def confirm_budget():
    if request.method == 'POST':
        budget = request.form['budget']
        confirm = request.form['confirm']
        observations = request.form['observations']
        photo_path = None  # Inicializa photo_path como None
        
        # Verifica se o campo de arquivo 'photo' está presente no request
        if 'photo' in request.files:
            photo = request.files['photo']
            # Salvar a foto em um diretório de uploads
            if photo.filename != '':
                photo_path = os.path.join(app.config['UPLOAD_FOLDER'], photo.filename)
                photo.save(photo_path)

        # Salvar os dados na planilha Excel de feedback
        save_to_excel(budget, confirm, observations, photo_path)

        # Exibir os dados no painel de visualização
        return render_template('confirmation.html', budget=budget, confirm=confirm, observations=observations, photo_filename=photo_path)

    return redirect(url_for('index'))


def save_to_excel(budget, confirm, observations, photo_path):
    # Cria um DataFrame com os dados a serem adicionados após envio
    new_data = pd.DataFrame({
        'Orçamento': [budget],
        'Confirmado': [confirm],
        'Observações': [observations],
        'Foto': [photo_path]
    })

    # Carrega a planilha Excel existente ou criar uma nova se não existir
    try:
        excel_data = pd.read_excel(output_excel_path, sheet_name='Feedback')
    except FileNotFoundError:
        excel_data = pd.DataFrame()

    # Concatenar o novo DataFrame com os dados existentes
    updated_data = pd.concat([excel_data, new_data], ignore_index=True)

    # Salvar de volta na planilha Excel de feedback
    updated_data.to_excel(output_excel_path, index=False, sheet_name='Feedback')


# Rota para exibir a foto
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


if __name__ == '__main__':
    app.run(debug=True)
