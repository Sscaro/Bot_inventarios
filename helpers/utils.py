import logging
import os
from datetime import datetime
def configurar_logger():
    # Crear carpeta "logs" si no existe
    os.makedirs("logs", exist_ok=True)
    # Crear nombre del archivo con fecha y hora actual
    fecha_actual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_filename = f"logs/log_{fecha_actual}.log"

    # Configurar el logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Evitar agregar múltiples handlers si se llama varias veces
    if not logger.handlers:
        # Formato del mensaje
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        # Handler para archivo
        file_handler = logging.FileHandler(log_filename, encoding="utf-8")
        file_handler.setFormatter(formatter)

        # Handler para consola
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)

        # Agregar ambos handlers al logger
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger