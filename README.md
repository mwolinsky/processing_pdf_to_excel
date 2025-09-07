# PDF Table Extractor

Una aplicación web Flask que extrae tablas de archivos PDF y genera archivos Excel con cálculos específicos.

## Características

- Extracción de tablas de PDF usando pdfplumber
- Procesamiento de datos con pandas
- Cálculos automáticos de precios e IVA
- Generación de archivos Excel con resultados
- Interfaz web amigable

## Requisitos

- Python 3.9+
- Flask
- pdfplumber
- pandas
- xlsxwriter
- werkzeug

## Instalación

1. Clonar el repositorio:
```bash
git clone [URL_DEL_REPOSITORIO]
cd [NOMBRE_DEL_DIRECTORIO]
```

2. Crear un entorno virtual e instalar dependencias:
```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

1. Iniciar la aplicación:
```bash
python app.py
```

2. Abrir un navegador web y visitar:
```
http://127.0.0.1:5000
```

3. Subir un archivo PDF y recibir el Excel procesado

## Estructura del Proyecto

```
.
├── app.py              # Aplicación principal Flask
├── requirements.txt    # Dependencias del proyecto
├── templates/         
│   └── index.html     # Plantilla HTML para la interfaz web
├── Dockerfile         # Configuración para Docker
└── gunicorn.conf.py   # Configuración de Gunicorn para producción
```

## Despliegue

El proyecto está configurado para ser desplegado en Render.com usando Docker. Ver la documentación de despliegue para más detalles.

## Licencia

[MIT License](https://opensource.org/licenses/MIT)
