import logging
import time
from src.hda_web.client import HDAClient
from src.common.config import get_settings
from src.common.system import kill_processes

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("test_suspend")

def main():
    settings = get_settings()
    client = HDAClient(settings)
    
    logger.info("Limpiando procesos...")
    kill_processes(["msedge.exe", "msedgedriver.exe"])
    time.sleep(2)

    try:
        logger.info("Iniciando navegador...")
        client.start()
        
        logger.info("Iniciando sesión...")
        client.login()
        
        logger.info("Navegando a Payments...")
        client.click_payments_tile()
        
        TICKET_ID = "728372G"
        REASONS = ["Test de suspensión paso a paso."]
        
        logger.info("Abriendo ticket %s...", TICKET_ID)
        client.open_ticket_by_id(TICKET_ID)
        time.sleep(3)
        
        logger.info("Iniciando suspensión...")
        client.suspend_ticket(TICKET_ID, REASONS)
        
        logger.info("Proceso terminado. El navegador se quedará abierto para inspección manual.")
        input("Presiona ENTER para cerrar el navegador y finalizar...")
        
    except Exception as e:
        logger.error("Error durante la prueba: %s", e)
        input("Presiona ENTER para cerrar el navegador y finalizar...")
    finally:
        if client.driver:
            client.driver.quit()

if __name__ == "__main__":
    main()