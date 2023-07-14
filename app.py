from flask import Flask, render_template, request, send_file
from flask_wtf.csrf import CSRFProtect
from flask_talisman import Talisman
import pandas as pd
from io import BytesIO
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address



app = Flask(__name__)
app.config['SECRET_KEY'] = 'citraditboard2023+'  # Necesario para CSRF y Flask-Session, cambia 'your-secret-key' por una clave aleatoria segura
app.config['SESSION_COOKIE_SECURE'] = True  # Configurar la cookie de sesión para que solo se envíe a través de HTTPS
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Max file size of 16MB

csrf = CSRFProtect(app)  # Habilitar CSRF protection
Talisman(app)  # Habilitar Flask-Talisman para encabezados de seguridad HTTP
limiter = Limiter(key_func=get_remote_address)
limiter.init_app(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
@csrf.exempt # Exención de la protección CSRF para esta ruta
@limiter.limit("10/minute")  # 10 requests per minute
def upload_file():
    uploaded_file = request.files['file']
    if not uploaded_file.filename.endswith('.xlsx'):
        return 'Invalid file type', 400
    
    df = pd.read_excel(uploaded_file)
    
    df.columns =['id', 'nombreCita', 'documento', 'gruposDocumentos',
       'contenidoCita', 'comentario', 'codigos', 'referencia', 'densidad', 'extension','creadoPor',
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
    app.run(host='0.0.0.0', port=5000, debug=False)
