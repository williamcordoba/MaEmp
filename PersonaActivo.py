import os
import time
import glob
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# WebDriver Manager (opcional)
try:
    from webdriver_manager.chrome import ChromeDriverManager
    WEBDRIVER_MANAGER_AVAILABLE = True
except Exception:
    WEBDRIVER_MANAGER_AVAILABLE = False


class DescargaMaEmp:
    def __init__(self, headless=True):
        self.url = "https://permoda.midasoft.co/NGMidasoft/login/10"
        self.usuario = "WILLIAMCC"
        self.password = "Koaj2025.."
        self.download_dir = os.path.join(os.getcwd(), "descargas_maemp")
        os.makedirs(self.download_dir, exist_ok=True)

        self.driver = None
        self.wait = None
        self.archivo_descargado = None
        self.headless = headless

    # ---------------------------------------------------------
    # Utilidades
    # ---------------------------------------------------------
    def limpiar_carpeta_descargas(self):
        for patron in ("*.xls", "*.xlsx", "*.csv", "*.xlsm", "*.crdownload"):
            for archivo in glob.glob(os.path.join(self.download_dir, patron)):
                try:
                    os.remove(archivo)
                except Exception:
                    pass
        return set(os.listdir(self.download_dir))

    def renombrar_archivo_descargado(self, archivo_original):
        try:
            fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
            ext = os.path.splitext(archivo_original)[1]
            nuevo = f"Informe_MaEmp_{fecha}{ext}"
            os.rename(
                os.path.join(self.download_dir, archivo_original),
                os.path.join(self.download_dir, nuevo)
            )
            return nuevo
        except Exception:
            return archivo_original

    # ---------------------------------------------------------
    # Driver
    # ---------------------------------------------------------
    def configurar_driver(self):
        chrome_options = Options()

        if self.headless:
            chrome_options.add_argument("--headless=new")

        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")

        # 🔴 RUTA FIJA A CHROME (Windows)
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            raise Exception(
                "Google Chrome no encontrado.\n"
                f"Instálalo o verifica la ruta:\n{chrome_path}"
            )
        chrome_options.binary_location = chrome_path

        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "plugins.always_open_pdf_externally": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        if WEBDRIVER_MANAGER_AVAILABLE:
            service = Service(ChromeDriverManager().install())
        else:
            service = Service()

        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 40)

    # ---------------------------------------------------------
    # Flujo
    # ---------------------------------------------------------
    def login(self):
        self.driver.get(self.url)
        time.sleep(2)

        usuario = password = None
        for inp in self.driver.find_elements(By.TAG_NAME, "input"):
            t = inp.get_attribute("type")
            if t == "text" and not usuario:
                usuario = inp
            elif t == "password" and not password:
                password = inp

        if not usuario or not password:
            raise Exception("Inputs de login no encontrados")

        usuario.send_keys(self.usuario)
        password.send_keys(self.password)

        for btn in self.driver.find_elements(By.TAG_NAME, "button"):
            if "INGRESAR" in (btn.text or "").upper():
                btn.click()
                break

        time.sleep(4)

    def buscar_informe(self):
        campo = None
        for inp in self.driver.find_elements(By.TAG_NAME, "input"):
            if inp.is_displayed() and inp.is_enabled():
                campo = inp
                break

        if not campo:
            raise Exception("Campo de búsqueda no encontrado")

        campo.clear()
        campo.send_keys("INFORME MAEMP")
        time.sleep(1)
        campo.send_keys(Keys.ENTER)
        time.sleep(3)

    def entrar_iframe(self):
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            self.driver.switch_to.frame(iframes[0])
            time.sleep(1)

    def ejecutar_informe(self):
        for btn in self.driver.find_elements(By.TAG_NAME, "button"):
            texto = (btn.text or "").upper()
            if any(k in texto for k in ("CONSULTAR", "GENERAR", "EJECUTAR")):
                btn.click()
                time.sleep(5)
                return
        raise Exception("Botón ejecutar no encontrado")

    def exportar_excel(self):
        iniciales = self.limpiar_carpeta_descargas()

        for btn in self.driver.find_elements(By.TAG_NAME, "button"):
            if "EXCEL" in (btn.text or "").upper():
                btn.click()
                break

        inicio = time.time()
        while time.time() - inicio < 40:
            actuales = set(os.listdir(self.download_dir))
            nuevos = actuales - iniciales
            for f in nuevos:
                if f.lower().endswith((".xls", ".xlsx", ".csv")):
                    nuevo = self.renombrar_archivo_descargado(f)
                    self.archivo_descargado = nuevo
                    return True
            time.sleep(1)

        return False

    # ---------------------------------------------------------
    # Ejecución
    # ---------------------------------------------------------
    def ejecutar(self):
        try:
            self.configurar_driver()
            self.login()
            self.buscar_informe()
            self.entrar_iframe()
            self.ejecutar_informe()
            self.exportar_excel()
        except Exception as e:
            print(f"ERROR_CRITICO: {e}")
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass

            if self.archivo_descargado:
                ruta = os.path.abspath(
                    os.path.join(self.download_dir, self.archivo_descargado)
                )
                print(f"RESULT_PATH::{ruta}")
            else:
                print("RESULT_PATH::")


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    bot = DescargaMaEmp(headless=True)
    bot.ejecutar()
