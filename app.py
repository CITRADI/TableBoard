from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def upload_file():
    uploaded_file = request.files['file']
    df = pd.read_excel(uploaded_file)
    
    df.columns =['identificador', 'nombreCita', 'documento', 'gruposDocumentos',
       'contenidoCita', 'comentario', 'codigos', 'referencia', 'densidad',
       'modificadoPor', 'creado', 'modificado']
    # Agrupar los valores duplicados por documento y código de cita
    agrupado = df.groupby(['documento', 'codigos'], as_index=False)['contenidoCita'].first()
    # Filtrar los datos de la hoja AdministradorCitas
    filtro = agrupado[agrupado['documento'] != '']
    datos_filtrados = filtro[['documento', 'codigos', 'contenidoCita']]
    # Transformar los datos para tener los códigos de cita como columnas
    datos_pivoteados = datos_filtrados.pivot(index='documento', columns='codigos', values='contenidoCita')

    # Rellenar los valores faltantes con 'No existe'
    #datos_pivoteados = datos_pivoteados.fillna('No existe')

    # Exportar los datos pivoteados como un archivo CSV
    # Generar el archivo CSV
    csv = datos_pivoteados.to_csv('datos_pivoteados.csv', sep=';', index=True)
    with open('datos_pivoteados.csv', 'rb') as f:
        csv = BytesIO(f.read())

    # Devolver el archivo CSV como descarga
    return send_file(
        csv,
        mimetype='text/csv',
        attachment_filename='resultado.csv',
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(debug=True)
