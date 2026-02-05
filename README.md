Proyecto bot inventarios

configuracion incial:
crear archivo .env de de la siguiente manera:

'''''
#CONFIGURACION USUARIO
USER_NAME= xxxxxx # espcificar usuario de sap
PASSWORD =  xxxxx # especificar contraseña sap

##CONFIGURACION SAP
SAP_PATH = C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe
CUBO= 1. Grupo Nutresa_ERP_PRD
'''''

Descargar ambiente embedido.
python-3.13-embed/ 

Seguir claramente la estrucutra del proyecto, alli se especifica donde van cada archivo


Estrucutura:
Bot_inventarios/
├── helpers/                 # Funciones auxiliares y herramientas comunes
│   ├── __init__.py
│   └── utils.py
├── logs/                    # Registro de actividad del bot (errores y eventos)
│   └── *.log
├── modulos/                 # Lógica de negocio y conexión externa
│   ├── __init__.py
│   └── conexion_sap.py      # Script principal de interacción con SAP
├── python-3.13-embed/       # Entorno de ejecución portable (Python embebido)
├── salidas/                 # Carpeta destino para reportes y archivos exportados
├── .env                     # Variables de entorno (Credenciales, rutas críticas)
├── .gitignore               # Archivos omitidos por Git (logs, .env, .venv)
├── install_packages.bat     # Setup inicial del entorno embebido
├── main.py                  # Orquestador principal (Punto de entrada)
├── pyproject.toml           # Configuración original de dependencias (uv)
├── README.md                # Documentación del proyecto
├── requirements.txt         # Librerías necesarias para el entorno portátil
└── run.bat


INSTRUCCCIONES EN SAP:
    transacción: "SQ01"
    transaccón_busqueda "ZISSC_CENT_MAT"
    fecha inicio: dia actual
    fecha fin: dia actual
    hora inicio: hora antes de la hora actual
    hora fin: hora redondeada por abajo hora acual
    archivo generado: informe_fecha_hora.xlsx
    Guardado: carpeta salidas.

Primeros pasos.
    1. instalar librerias:
        clic install_packages.bat
    2. ejecutar bot:
        clic run.bat