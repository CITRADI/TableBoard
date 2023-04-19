# Usa una imagen de Python como base
FROM python:3.8-slim-buster

# Establece el directorio de trabajo en /app
WORKDIR /app

# Copia el archivo de requisitos a la imagen
COPY requirements.txt .

# Instala las dependencias especificadas en el archivo de requisitos
RUN pip install --no-cache-dir -r requirements.txt

# Copia el código fuente a la imagen
COPY . .

# Expone el puerto 5000
EXPOSE 5000

# Ejecuta la aplicación cuando se inicia el contenedor
CMD [ "python", "app.py" ]
