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

# Intentar usar webdriver-manager si está disponible (más fácil en entornos sin chromedriver en PATH)
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

    def limpiar_carpeta_descargas(self):
        patrones = ["*.xls", "*.xlsx", "*.csv", "*.xlsm"]
        for patron in patrones:
            for archivo in glob.glob(os.path.join(self.download_dir, patron)):
                try:
                    os.remove(archivo)
                except:
                    pass
        return set()

    def renombrar_archivo_descargado(self, archivo_original):
        try:
            ahora = datetime.now()
            fecha_formato = ahora.strftime("%Y%m%d_%H%M%S")
            extension = os.path.splitext(archivo_original)[1]
            nuevo_nombre = f"Informe_MaEmp_{fecha_formato}{extension}"
            ruta_original = os.path.join(self.download_dir, archivo_original)
            ruta_nueva = os.path.join(self.download_dir, nuevo_nombre)
            os.rename(ruta_original, ruta_nueva)
            return nuevo_nombre
        except:
            return archivo_original

   def configurar_driver(self):
    chrome_options = Options()

    # Headless moderno
    if self.headless:
        chrome_options.add_argument("--headless=new")

    # Estabilidad
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    # 🔴 RUTA FIJA A CHROME (OBLIGATORIO EN WINDOWS)
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    if not os.path.exists(chrome_path):
        raise Exception(
            "Google Chrome no está instalado o no se encontró en la ruta esperada:\n"
            f"{chrome_path}"
        )

    chrome_options.binary_location = chrome_path

    # Preferencias de descarga
    prefs = {
        "download.default_directory": self.download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # WebDriver
    if WEBDRIVER_MANAGER_AVAILABLE:
        service = Service(ChromeDriverManager().install())
    else:
        service = Service()

    self.driver = webdriver.Chrome(service=service, options=chrome_options)
    self.wait = WebDriverWait(self.driver, 40)


    def login(self):
        self.driver.get(self.url)
        time.sleep(2)
        try:
            usuario_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder*='Usuario' i]")
            password_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder*='Contraseña' i]")
        except:
            inputs = self.driver.find_elements(By.TAG_NAME, "input")
            usuario_input = password_input = None
            for inp in inputs:
                tipo = inp.get_attribute("type")
                if tipo == "text" and usuario_input is None:
                    usuario_input = inp
                elif tipo == "password" and password_input is None:
                    password_input = inp
        if not usuario_input or not password_input:
            raise Exception("No se encontraron inputs de usuario/contraseña")
        usuario_input.send_keys(self.usuario)
        password_input.send_keys(self.password)
        boton_ingresar = None
        for btn in self.driver.find_elements(By.TAG_NAME, "button"):
            if "INGRESAR" in (btn.text or "").upper():
                boton_ingresar = btn
                break
        if not boton_ingresar:
            for inp in self.driver.find_elements(By.CSS_SELECTOR, "input[type='submit']"):
                if "INGRESAR" in (inp.get_attribute("value") or "").upper():
                    boton_ingresar = inp
                    break
        if not boton_ingresar:
            raise Exception("No se encontró el botón INGRESAR")
        boton_ingresar.click()
        try:
            self.wait.until(lambda d: "menu" in d.current_url or "vista-clasica" in d.current_url)
        except:
            time.sleep(3)
        return True

    def buscar_informe(self):
        time.sleep(1)
        campo_busqueda = None
        selectores_busqueda = [
            "input[type='search']",
            "input[type='text']",
            "input[placeholder*='buscar' i]",
            "input.search-input",
            "input[role='searchbox']",
            "input.global-search"
        ]
        for selector in selectores_busqueda:
            try:
                elementos = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for elem in elementos:
                    if elem.is_displayed() and elem.is_enabled():
                        campo_busqueda = elem
                        break
                if campo_busqueda:
                    break
            except:
                continue
        if not campo_busqueda:
            all_inputs = self.driver.find_elements(By.TAG_NAME, "input")
            for inp in all_inputs:
                try:
                    if inp.is_displayed() and inp.is_enabled() and inp.get_attribute("type") != "hidden":
                        campo_busqueda = inp
                        break
                except:
                    continue
        if not campo_busqueda:
            raise Exception("No se encontró campo de búsqueda")
        campo_busqueda.clear()
        campo_busqueda.send_keys("INFORME MAEMP")
        time.sleep(1)
        campo_busqueda.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)
        campo_busqueda.send_keys(Keys.ENTER)
        time.sleep(3)

    def entrar_iframe(self):
        time.sleep(1)
        try:
            self.wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
        except:
            time.sleep(2)
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        if not iframes:
            return False
        self.driver.switch_to.frame(iframes[0])
        time.sleep(1)
        return True

    def ejecutar_informe(self):
        time.sleep(1)
        boton_ejecutar = None
        selectores_verde = [
            "input.btn.btn-success[type='submit']",
            "button.btn.btn-success",
            "input.btn-success",
            "button.btn-success",
            ".btn-success",
            "[class*='btn-success']"
        ]
        for selector in selectores_verde:
            try:
                elementos = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for elem in elementos:
                    if elem.is_displayed() and elem.is_enabled():
                        boton_ejecutar = elem
                        break
                if boton_ejecutar:
                    break
            except:
                continue
        if not boton_ejecutar:
            botones = self.driver.find_elements(By.TAG_NAME, "button")
            for btn in botones:
                try:
                    if btn.is_displayed() and btn.is_enabled():
                        texto = (btn.text or "").upper()
                        if any(palabra in texto for palabra in ["CONSULTAR", "GENERAR", "EJECUTAR", "BUSCAR", "FILTRAR"]):
                            boton_ejecutar = btn
                            break
                except:
                    continue
        if not boton_ejecutar:
            raise Exception("No se encontró botón para ejecutar el informe")
        try:
            boton_ejecutar.click()
        except:
            try:
                self.driver.execute_script("arguments[0].click();", boton_ejecutar)
            except:
                ActionChains(self.driver).move_to_element(boton_ejecutar).click().perform()
        time.sleep(5)

    def exportar_excel(self):
        archivos_iniciales = self.limpiar_carpeta_descargas()
        time.sleep(1)
        boton_exportar = None
        selectores_export = [
            "input[value*='Excel' i]",
            "input[value*='Export' i]",
            "button:has(svg)",
            "button.btn-link",
            ".btn-success"
        ]
        for selector in selectores_export:
            try:
                elementos = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for elem in elementos:
                    if elem.is_displayed() and elem.is_enabled():
                        boton_exportar = elem
                        break
                if boton_exportar:
                    break
            except:
                continue
        if not boton_exportar:
            raise Exception("No se encontró botón de exportación")
        try:
            boton_exportar.click()
        except:
            try:
                self.driver.execute_script("arguments[0].click();", boton_exportar)
            except:
                ActionChains(self.driver).move_to_element(boton_exportar).click().perform()
        return self._esperar_descarga_estricta(archivos_iniciales, timeout=40)

    def _esperar_descarga_estricta(self, archivos_iniciales, timeout=30):
        inicio = time.time()
        archivos_descargados = []
        time.sleep(2)
        while time.time() - inicio < timeout:
            archivos_actuales = set(os.listdir(self.download_dir))
            nuevos_archivos = archivos_actuales - archivos_iniciales
            if nuevos_archivos:
                for archivo in list(nuevos_archivos):
                    ruta_completa = os.path.join(self.download_dir, archivo)
                    if archivo.endswith('.crdownload') or archivo.endswith('.part'):
                        continue
                    if archivo.lower().endswith((".xls", ".xlsx", ".csv", ".xlsm")) and os.path.exists(ruta_completa):
                        tamaño = os.path.getsize(ruta_completa)
                        if tamaño > 1000:
                            if archivo not in [a[0] for a in archivos_descargados]:
                                archivos_descargados.append((archivo, tamaño))
                                archivo_renombrado = self.renombrar_archivo_descargado(archivo)
                                self.archivo_descargado = archivo_renombrado
                                time.sleep(2)
            temporales = [f for f in os.listdir(self.download_dir) if f.endswith('.crdownload') or f.endswith('.part')]
            if self.archivo_descargado and not temporales:
                return True
            time.sleep(1)
        return bool(archivos_descargados)

    def crear_resumen_descarga(self):
        if not self.archivo_descargado:
            return
        try:
            resumen_path = os.path.join(self.download_dir, "resumen_descarga.txt")
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ruta_archivo = os.path.join(self.download_dir, self.archivo_descargado)
            if os.path.exists(ruta_archivo):
                tamaño = os.path.getsize(ruta_archivo)
                with open(resumen_path, 'w', encoding='utf-8') as f:
                    f.write("=" * 60 + "\n")
                    f.write("RESUMEN DE DESCARGA - INFORME MAEMP\n")
                    f.write("=" * 60 + "\n\n")
                    f.write(f"Fecha y hora de ejecución: {fecha_actual}\n")
                    f.write(f"Archivo descargado: {self.archivo_descargado}\n")
                    f.write(f"Tamaño: {tamaño:,} bytes ({tamaño/1024/1024:.2f} MB)\n")
                    f.write(f"Ubicación: {ruta_archivo}\n")
                    f.write(f"Usuario: {self.usuario}\n")
                    f.write("\n" + "=" * 60 + "\n")
        except:
            pass

    def ejecutar(self):
        try:
            self.limpiar_carpeta_descargas()
            self.configurar_driver()
            self.login()
            self.buscar_informe()
            self.entrar_iframe()
            self.ejecutar_informe()
            self.exportar_excel()
            self.crear_resumen_descarga()
        except Exception as e:
            print(f"ERROR: {e}")
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            if self.archivo_descargado:
                ruta = os.path.abspath(os.path.join(self.download_dir, self.archivo_descargado))
                print(f"RESULT_PATH::{ruta}")
            else:
                print("RESULT_PATH::")


if __name__ == "__main__":
    bot = DescargaMaEmp(headless=True)

    bot.ejecutar()
