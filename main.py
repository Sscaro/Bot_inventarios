import sys
import os
directorio_raiz = os.path.dirname(os.path.abspath(__file__))
sys.path.append(directorio_raiz)
from dotenv import load_dotenv
from modulos import ejecutar_bot
from helpers import configurar_logger

logger = configurar_logger()
def main():
    load_dotenv()
    ejecutar_bot()
    logger.info("Hola bot-inventarios!👍 ")

if __name__ == "__main__":
    main()
