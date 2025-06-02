# Codigo_Correos_Automaticos

🧠 Descripción general del sistema
Este proyecto automatiza el análisis, clasificación, y respuesta de correos electrónicos institucionales utilizando la API de OpenAI y Microsoft Graph. También integra el procesamiento de datos desde archivos Excel, permitiendo mantener la información de preguntas frecuentes, etiquetas y respuestas actualizadas dinámicamente.

⚙️ Tecnologías utilizadas
Python 3.10+
Microsoft Graph API
OpenAI API (GPT-3.5 Turbo)
openpyxl para manejo de Excel
pandas para análisis de datos
watchdog para monitoreo de archivos
requests, BeautifulSoup para manejo de contenido HTML

📦 ProyectoCorreoAutomatizado/
├── PruebaConModeloEnviarCorreo v12.py          # Script principal del sistema
├── respuestas2.xlsx                            # Base de datos de etiquetas, preguntas frecuentes y respuestas
├── correos_procesados.xlsx                     # Exportación de correos procesados
├── conteo_respuestas_etiquetas.xlsx            # Estadísticas de uso de etiquetas
├── correos_procesados.json                     # Correos procesados con fechas
├── hilos_procesados.json                       # Hilos procesados con fechas
├── correos_contados_temp.json                  # IDs temporales usados para conteo
└── README.md                                   # Documentación del proyecto

🔄 Funcionalidades principales
✅ Autenticación con Microsoft 365 vía msal
✅ Lectura asincrónica de correos electrónicos de la bandeja de entrada
✅ Clasificación inteligente usando GPT según categorías definidas
✅ Identificación de nuevas etiquetas cuando no hay coincidencias
✅ Generación y envío de respuestas en HTML (con negrilla y enlaces)
✅ Exportación de información a Excel y eliminación de duplicados
✅ Conteo acumulativo de respuestas por etiqueta
✅ Monitoreo en tiempo real de cambios en el archivo de Excel

🗂️ Archivos clave
respuestas2.xlsx
Contiene:
Palabras clave (columna A)
Pregunta frecuente (columna C)
Respuesta (columna D)
Sinónimos (columna E)
Hoja credenciales y hoja excluir para gestionar acceso y filtros
PruebaConModeloEnviarCorreo v12.py
Código completo del sistema, incluye:
Clasificación con OpenAI
Interacción con Graph API
Exportación a Excel
Monitoreo de archivos y persistencia de datos procesados

🧪 Requisitos para ejecución
Python 3.10+

Instalar dependencias:

bash
Copiar
Editar
pip install openai openpyxl pandas watchdog requests msal beautifulsoup4 unidecode matplotlib
Configurar archivo .env con la clave de OpenAI:

ini
Copiar
Editar
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxx
Asegúrate de tener acceso a la cuenta de Microsoft 365 para autenticación con usuario y contraseña.

🚀 Ejecución
El script inicia automáticamente la lectura de correos y el monitoreo del archivo Excel al ejecutarlo:

bash
Copiar
Editar
python PruebaConModeloEnviarCorreo\ con\ chat\ gtp\ -excel\ -Etiqueta\ -excel\ -VariasRespuestas\ -Modificacion\ v12.py

📊 Resultados
Correos clasificados y respondidos automáticamente
Excel actualizado con etiquetas y respuestas enviadas
Conteo total por etiqueta disponible para seguimiento









