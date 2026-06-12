import csv
import logging
import time
import os
import re
import base64
import requests
from datetime import datetime
from io import BytesIO
from dotenv import load_dotenv
import psycopg2
import pytesseract
from PIL import Image
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ==========================================
# CONFIGURACIÓN
# ==========================================
URL_SEARCH_COMPANY = 'https://appscvsgen.supercias.gob.ec/consultaCompanias/societario/busquedaCompanias.jsf'
BASE_URL = 'https://appscvsgen.supercias.gob.ec'

LIMITE_DESCARGA = 50
CARPETA_PDFS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdfs_descargados')
CSV_MAPEO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'mapeo_pdfs.csv')
CSV_ERRORES = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'download_errores.csv')

load_dotenv()

DB_HOST = os.getenv("DB_HOST")
DB_PORT = os.getenv("DB_PORT")
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logging.captureWarnings(True)
logging.getLogger('urllib3').setLevel(logging.ERROR)


# ==========================================
# NAVEGACIÓN Y CAPTCHA
# ==========================================
def process_captcha(driver):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-panel-content.ui-widget-content>table>tbody>tr:nth-child(4)>td img"))
    )
    captcha_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-panel-content.ui-widget-content>table>tbody>tr:nth-child(4)>td img"))
    )
    WebDriverWait(driver, 10).until(lambda d: captcha_element.size['width'] > 0)

    captcha_bytes = captcha_element.screenshot_as_png
    image = Image.open(BytesIO(captcha_bytes))
    image = image.convert('L')
    width, height = image.size
    image = image.resize((width * 2, height * 2), Image.Resampling.LANCZOS)

    text = pytesseract.image_to_string(image, config='--psm 7').strip()
    logging.info(f"Captcha extraído: {text}")

    captcha_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-panel-content.ui-widget-content>table>tbody>tr:nth-child(4)>td input"))
    )
    captcha_input.send_keys(text)

    boton_buscar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-button-text.ui-c"))
    )
    boton_buscar.click()
    time.sleep(5)

    return captcha_bytes


def process_captcha_modal(driver):
    try:
        captcha_dialog = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#dlgCaptcha"))
        )
        if not captcha_dialog.is_displayed():
            return False

        logging.info("Captcha modal detectado")

        captcha_img = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#frmCaptcha\\:captchaImage"))
        )
        WebDriverWait(driver, 10).until(lambda d: captcha_img.size['width'] > 0)

        captcha_bytes = captcha_img.screenshot_as_png
        image = Image.open(BytesIO(captcha_bytes))
        image = image.convert('L')
        width, height = image.size
        image = image.resize((width * 2, height * 2), Image.Resampling.LANCZOS)

        text = pytesseract.image_to_string(image, config='--psm 7').strip()
        text = re.sub(r'\D', '', text)
        logging.info(f"Captcha modal: {text}")

        captcha_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#frmCaptcha\\:captcha"))
        )
        captcha_input.clear()
        captcha_input.send_keys(text)

        boton_verificar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#frmCaptcha\\:btnPresentarContenido"))
        )
        boton_verificar.click()
        time.sleep(4)
        return True

    except TimeoutException:
        return False


def _navegar_a_compania(driver, expediente):
    driver.get(URL_SEARCH_COMPANY)
    time.sleep(2)

    tipo_busqueda = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-selectoneradio.ui-widget td:nth-child(1)>div"))
    )
    tipo_busqueda.click()
    time.sleep(1)

    parametro_busqueda = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-autocomplete>input"))
    )
    parametro_busqueda.send_keys(str(expediente))

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-autocomplete-item.ui-autocomplete-list-item.ui-corner-all.ui-state-highlight"))
    )

    auto_complete_item = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-autocomplete-item.ui-autocomplete-list-item.ui-corner-all.ui-state-highlight"))
    )
    auto_complete_item.click()
    time.sleep(2)

    try:
        captcha_img = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-panel-content.ui-widget-content>table>tbody>tr:nth-child(4)>td img"))
        )
        if captcha_img.is_displayed():
            process_captcha(driver)
    except TimeoutException:
        pass

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='frmMenu:menuInformacionAnualPresentada']"))
    )
    logging.info(f"Compañía {expediente} cargada")


def navigate_to_informacion_anual(driver):
    try:
        menu_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='frmMenu:menuInformacionAnualPresentada']"))
        )
        menu_item.click()
        time.sleep(3)

        process_captcha_modal(driver)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='frmInformacionCompanias:tblInformacionAnual_data']"))
        )
        return True

    except TimeoutException:
        logging.warning("No se pudo cargar tabla de información anual")
        return False


def find_auditoria_externa_row(driver):
    page = 1
    while True:
        logging.info(f"Buscando Auditoría Externa en página {page}")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='frmInformacionCompanias:tblInformacionAnual_data']//tr"))
            )
            rows = driver.find_elements(By.XPATH, "//*[@id='frmInformacionCompanias:tblInformacionAnual_data']//tr[@data-ri]")

            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) < 5:
                        continue
                    codigo_text = cells[1].text.strip()
                    nombre_text = cells[2].text.strip()

                    if codigo_text == '3.1.2' or 'Auditoria Externa' in nombre_text:
                        logging.info(f"Fila encontrada: código={codigo_text}, nombre={nombre_text}")
                        anio_fila = None
                        try:
                            anio_text = cells[0].text.strip()
                            anio_fila = int(anio_text) if anio_text.isdigit() else None
                        except Exception:
                            pass
                        pdf_link = cells[4].find_element(By.CSS_SELECTOR, "a.ui-commandlink")
                        return pdf_link, anio_fila

                except NoSuchElementException:
                    continue

        except TimeoutException:
            logging.warning("Timeout esperando filas")
            return None, None

        try:
            next_btn = driver.find_element(
                By.XPATH,
                "//*[@id='frmInformacionCompanias:tblInformacionAnual_paginator_bottom']//a[contains(@class,'ui-paginator-next')]"
            )
            if 'ui-state-disabled' in next_btn.get_attribute('class'):
                logging.info("No hay más páginas. Auditoría externa no encontrada.")
                return None, None
            next_btn.click()
            time.sleep(2)
            page += 1
        except NoSuchElementException:
            return None, None


def _is_visible(driver, xpath):
    try:
        el = driver.find_element(By.XPATH, xpath)
        return el.is_displayed()
    except Exception:
        return False


def extract_pdf_url_from_modal(driver):
    try:
        WebDriverWait(driver, 15).until(
            lambda d: (
                _is_visible(d, "//*[@id='dlgPresentarDocumentoPdf']") or
                _is_visible(d, "//*[@id='dlgPresentarDocumentoPdfConFirmasElectronicas']")
            )
        )
    except TimeoutException:
        logging.warning("Ningún modal de PDF apareció")
        return None

    try:
        dlg_normal = driver.find_element(By.XPATH, "//*[@id='dlgPresentarDocumentoPdf']")
        if dlg_normal.is_displayed():
            try:
                obj_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[@id='panelInternoPresentarDocumentoPdf']//object[@type='application/pdf']")
                    )
                )
                pdf_data = obj_element.get_attribute("data")
                if pdf_data:
                    pdf_url = pdf_data.split('?')[0]
                    if not pdf_url.startswith('http'):
                        pdf_url = BASE_URL + pdf_url
                    logging.info(f"URL PDF (modal normal): {pdf_url}")
                    return pdf_url
            except TimeoutException:
                pass
    except NoSuchElementException:
        pass

    try:
        dlg_firma = driver.find_element(By.XPATH, "//*[@id='dlgPresentarDocumentoPdfConFirmasElectronicas']")
        if dlg_firma.is_displayed():
            boton_aceptar = dlg_firma.find_element(
                By.XPATH, ".//button[contains(@onclick,'window.open')]"
            )
            onclick = boton_aceptar.get_attribute("onclick")
            match = re.search(r"window\.open\(['\"]([^'\"]+\.pdf)", onclick)
            if match:
                raw_url = match.group(1)
                pdf_url = raw_url if raw_url.startswith('http') else BASE_URL + raw_url
                logging.info(f"URL PDF (firma electrónica): {pdf_url}")
                try:
                    close_btn = dlg_firma.find_element(
                        By.XPATH, ".//a[contains(@class,'ui-dialog-titlebar-close')]"
                    )
                    close_btn.click()
                    time.sleep(0.5)
                except Exception:
                    pass
                return pdf_url
    except NoSuchElementException:
        pass

    return None


def close_pdf_modal(driver):
    try:
        close_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='dlgPresentarDocumentoPdf']//a[contains(@class,'ui-dialog-titlebar-close')]"))
        )
        close_btn.click()
        time.sleep(1)
    except TimeoutException:
        pass
    try:
        close_btn = driver.find_element(
            By.XPATH, "//*[@id='dlgPresentarDocumentoPdfConFirmasElectronicas']//a[contains(@class,'ui-dialog-titlebar-close')]"
        )
        if close_btn.is_displayed():
            close_btn.click()
            time.sleep(1)
    except NoSuchElementException:
        pass


# ==========================================
# DESCARGA Y GUARDADO EN DISCO
# ==========================================
def download_and_save_pdf(driver, pdf_url, ruta_destino):
    """Descarga el PDF usando la sesión del browser y lo guarda en disco."""
    # Método 1: fetch() desde el browser
    try:
        b64 = driver.execute_async_script(
            """
            var url = arguments[0];
            var done = arguments[1];
            fetch(url, {credentials: 'include'})
                .then(r => r.arrayBuffer())
                .then(buf => {
                    var bytes = new Uint8Array(buf);
                    var binary = '';
                    for (var i = 0; i < bytes.byteLength; i++) {
                        binary += String.fromCharCode(bytes[i]);
                    }
                    done(btoa(binary));
                })
                .catch(e => done(null));
            """,
            pdf_url
        )
        if b64:
            pdf_content = base64.b64decode(b64)
            if len(pdf_content) > 1000:
                with open(ruta_destino, 'wb') as f:
                    f.write(pdf_content)
                logging.info(f"PDF guardado ({len(pdf_content)} bytes): {ruta_destino}")
                return True
    except Exception as e:
        logging.warning(f"fetch() falló: {e}, intentando con requests...")

    # Método 2: requests con cookies de Selenium
    try:
        selenium_cookies = driver.get_cookies()
        session = requests.Session()
        for cookie in selenium_cookies:
            session.cookies.set(cookie['name'], cookie['value'])

        headers = {
            'User-Agent': driver.execute_script("return navigator.userAgent;"),
            'Referer': driver.current_url,
        }

        response = session.get(pdf_url, headers=headers, timeout=30)
        response.raise_for_status()

        with open(ruta_destino, 'wb') as f:
            f.write(response.content)
        logging.info(f"PDF guardado via requests ({len(response.content)} bytes): {ruta_destino}")
        return True

    except Exception as e:
        logging.error(f"Error descargando PDF: {e}")
        return False


# ==========================================
# LÓGICA PRINCIPAL
# ==========================================
def procesar_descarga(driver, expediente, id_compania, mapeo_writer, error_writer):
    """Navega a la empresa, encuentra el PDF de Auditoría Externa y lo descarga."""
    try:
        _navegar_a_compania(driver, expediente)

        if not navigate_to_informacion_anual(driver):
            logging.warning(f"[{expediente}] No se pudo cargar Información Anual")
            error_writer.writerow([expediente, id_compania, 'ERROR_NAVEGACION', ''])
            return False

        pdf_link, anio = find_auditoria_externa_row(driver)
        if not pdf_link:
            logging.info(f"[{expediente}] Sin documento de Auditoría Externa")
            error_writer.writerow([expediente, id_compania, 'SIN_DOCUMENTO', anio or ''])
            return False

        pdf_link.click()
        time.sleep(3)

        if process_captcha_modal(driver):
            time.sleep(3)

        pdf_url = extract_pdf_url_from_modal(driver)
        if not pdf_url:
            logging.warning(f"[{expediente}] No se pudo obtener URL del PDF")
            close_pdf_modal(driver)
            error_writer.writerow([expediente, id_compania, 'ERROR_URL', anio or ''])
            return False

        nombre_archivo = f"{expediente}.pdf"
        ruta_destino = os.path.join(CARPETA_PDFS, nombre_archivo)

        ok = download_and_save_pdf(driver, pdf_url, ruta_destino)
        close_pdf_modal(driver)

        if ok:
            mapeo_writer.writerow([expediente, id_compania, ruta_destino, anio or ''])
            logging.info(f"[{expediente}] OK — año: {anio}")
            return True
        else:
            error_writer.writerow([expediente, id_compania, 'ERROR_DESCARGA', anio or ''])
            return False

    except Exception as e:
        logging.error(f"[{expediente}] Error inesperado: {e}")
        error_writer.writerow([expediente, id_compania, f'EXCEPCION: {e}', ''])
        try:
            close_pdf_modal(driver)
        except Exception:
            pass
        return False


def main():
    os.makedirs(CARPETA_PDFS, exist_ok=True)

    try:
        conn = psycopg2.connect(
            host=DB_HOST, port=DB_PORT,
            database=DB_NAME, user=DB_USER, password=DB_PASSWORD
        )
        cursor = conn.cursor()
        cursor.execute("""
            SELECT expediente, id
            FROM crm.adm_companias
            WHERE estado_proceso = 'T'
              AND fecha_actualizacion = '2026-05-26'
              AND activos >= 642000
            LIMIT %s;
        """, (LIMITE_DESCARGA,))
        records = cursor.fetchall()
        cursor.close()
        conn.close()
        logging.info(f"Registros a procesar: {len(records)}")
    except Exception as e:
        logging.error(f"Error conectando a BD: {e}")
        return

    options = Options()
    options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    driver = webdriver.Firefox(options=options)

    descargados = 0
    errores = 0

    try:
        with open(CSV_MAPEO, 'a', newline='', encoding='utf-8') as f_mapeo, \
             open(CSV_ERRORES, 'a', newline='', encoding='utf-8') as f_err:

            mapeo_writer = csv.writer(f_mapeo, delimiter=';')
            error_writer = csv.writer(f_err, delimiter=';')

            # Escribir encabezados si los archivos están vacíos
            if os.path.getsize(CSV_MAPEO) == 0:
                mapeo_writer.writerow(['expediente', 'id_compania', 'ruta_pdf', 'anio'])
            if os.path.getsize(CSV_ERRORES) == 0:
                error_writer.writerow(['expediente', 'id_compania', 'motivo', 'anio'])

            for i, (expediente, id_compania) in enumerate(records, 1):
                logging.info(f"--- [{i}/{len(records)}] Expediente: {expediente} ---")

                ok = procesar_descarga(driver, expediente, id_compania, mapeo_writer, error_writer)
                if ok:
                    descargados += 1
                else:
                    errores += 1

    except KeyboardInterrupt:
        logging.info("Interrumpido por el usuario.")
    finally:
        driver.quit()
        logging.info(f"Fin. Descargados: {descargados} | Errores: {errores}")
        logging.info(f"PDFs guardados en: {CARPETA_PDFS}")
        logging.info(f"Mapeo: {CSV_MAPEO}")


if __name__ == "__main__":
    main()
