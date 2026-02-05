'''
Modulo para realizar interacciones con SAP GUI mediante scripting.
'''
import win32com.client
import subprocess
import time
import os
import sys
import psutil
from pathlib import Path
from datetime import datetime, timedelta
from helpers import configurar_logger
from dotenv import load_dotenv

def obtener_mensaje(session):
    return session.findById("wnd[0]/sbar").Text.lower()

def convertir_valor_sap(texto):
    '''
    Función para convertir valores de texto a float, manejando formatos comunes en SAP.
    Parámetros: texto: str
    Retorna: float
    '''
    texto = texto.strip() #Quitar espacios
    # Mover el signo negativo al frente si está al final (ej: '913-' -> '-913')
    if texto.endswith('-'):
        texto = '-' + texto[:-1]
    # Quitar separadores de miles y ajustar coma decimal
    texto = texto.replace('.', '').replace(',', '.')
    # Si queda vacío o no es numérico, devolver 0
    try:
        return float(texto)
    except ValueError:
        return 0.0

class SapSession(object):
    """ Clase con funciones para moverse en SAP."""
    def __init__(self):
        super().__init__()
    # 🔹 NUEVO MÉTODO: cerrar sesiones de SAP si están abiertas
    def cerrar_sap(self):
        """Cierra cualquier proceso activo de SAP Logon antes de abrir uno nuevo."""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if 'saplogon.exe' in proc.info['name'].lower():
                    print(f"🛑 Cerrando proceso SAP Logon (PID {proc.info['pid']})...")
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass

        # Esperar unos segundos para asegurar que el proceso se cerró
        print("⏳ Esperando 5 segundos antes de continuar...")
        time.sleep(5)

    def sap_gui(self):
        """ Abre SAP de manera recursiva """
        try: 
            return win32com.client.GetObject('SAPGUI')
        except Exception:
            time.sleep(1)
            return self.sap_gui()

    def sap_login(self, relative_path, PRD):       
        """ Logueo en SAP """
        self.cerrar_sap() #Cierra SAP
        path = relative_path  # se debe poner la ruta de saplogon.exe, el metodo lo recibe como parametro.
        try:
            subprocess.Popen(path)
        except Exception:
            print('❗Fallo en conexión a sap, Asegurate de configurar correctamente la ruta completa de sap en el archivo .env')
            sys.exit()
        SapGuiAuto = self.sap_gui()
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return
        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            print('no hay conexión')
            SapGuiAuto = None
            return
        try:
            connection = application.Children(0)
        except:
            connection = application.OpenConnection(PRD, True)  # nombre de cubo de PRD
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return
        # # Manejar la ventana de usuario duplicado
        return session

class loguearse_sap(object):
    def __init__(self):
        super().__init__()  
    def connect_sap(self):
        """Conecta con SAP GUI y devuelve la sesión activa"""
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            print("❌ No se pudo obtener SAP GUI. Verifica que SAP Logon esté abierto.")
            return None

        application = SapGuiAuto.GetScriptingEngine
        if application.Children.Count == 0:
            print("❌ No hay conexiones abiertas en SAP Logon.")
            return None

        connection = application.Children(0)
        if connection.Children.Count == 0:
            print("❌ No hay sesión activa en SAP.")
            return None

        session = connection.Children(0)
        print("✅ Conectado a SAP.")
        return session
  
    def logger_sap(self):

        """Ejecuta la lógica del proceso"""
        # Cargar credenciales desde .env
        load_dotenv() #Importante tener archivo .env
        sap_user = os.getenv("USER_NAME") # se requiere archivo .env con credenciales.
        sap_pass = os.getenv("PASSWORD")

        session = self.connect_sap()
        if not session:
            return False
        else:
            #Login en SAP
            session.findById("wnd[0]").maximize()
            try:
                session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = sap_user
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = sap_pass
                session.findById("wnd[0]").sendVKey(0)  # Enter
                return session
            except Exception:
                print("ups al parecer hay un error con el usuario y contraseña, revisa tu .env y e intenta de nuevo")
                sys.exit()

def ejecutar_bot():
    """
    bot eue descarga información de SAP EN excel
    """
    now = datetime.now()
    fecha_sap = now.strftime("%d.%m.%Y")

    hora_fin = now.replace(minute=0, second=0, microsecond=0)
    hora_inicio = hora_fin - timedelta(hours=1)
    
    # horas de busqueda de consulta    
    hora_inicio_sap = hora_inicio.strftime("%H:%M:%S")
    hora_fin_sap = hora_fin.strftime("%H:%M:%S")

    # conectarse a sap
    session_obj = SapSession()
    session = session_obj.sap_login(os.getenv("SAP_PATH"),os.getenv("CUBO"))
    creando_objeto = loguearse_sap()
    session = creando_objeto.logger_sap()
    logger = configurar_logger()


    #ENTRANDO A LA TRANSACCION SQ01
    session.StartTransaction("SQ01")
    #TRANSACCION 
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "ZISSC_CENT_MAT"
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    #PARAMETROS DE SELECCION
    session.findById("wnd[0]/usr/ctxtSP$00019-LOW").text = fecha_sap
    session.findById("wnd[0]/usr/ctxtSP$00019-HIGH").text = fecha_sap
    session.findById("wnd[0]/usr/ctxtSP$00017-LOW").text = hora_inicio_sap
    session.findById("wnd[0]/usr/ctxtSP$00017-HIGH").text = hora_fin_sap
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # COMIENZA EXPORTACION A EXCEL
    grid = session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
    grid.setFocus()
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    
    #GUARDA EL ARCHIVO EXCEL.
    fecha_str = now.strftime("%d.%m.%Y")
    hora_str = now.strftime("%H%M")  # 0730
    nombre_archivo = f"informe_{fecha_str}_{hora_str}.xlsx"
    ruta_proyecto = Path.cwd()
    ruta_excel = ruta_proyecto / 'salidas' / nombre_archivo

    excel = None
    for _ in range(60):  # hasta 30 intentos (~30 segundos)
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            if excel.Workbooks.Count > 0:
                break
        except:
            pass
        time.sleep(1)
    if excel is None:
        raise Exception("Excel no se abrió después de la exportación XXL")
    
    workbook = excel.ActiveWorkbook
    # Esperar a que Excel termine de calcular / cargar
    while excel.CalculationState != 0:
        time.sleep(1)

    workbook.SaveAs(str(ruta_excel))
    workbook.Close(SaveChanges=False)
    excel.Quit()
    
    excel = None
    workbook = None
    logger.info(f"✅Archivo {nombre_archivo}  guardado exitosamente verificar carpeta salidas👍")
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    return True

