"""
Captura screenshots de Jenkins y los inserta en el documento Word.
Requiere Jenkins corriendo en localhost:8080
"""
import time
import pathlib
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE      = pathlib.Path(__file__).parent
DOCX_FILE = BASE / "entrega2_semana5.docx"
SHOTS_DIR = BASE / "screenshots"
SHOTS_DIR.mkdir(exist_ok=True)

JENKINS_URL = "http://localhost:8080"
USER        = "admin"
PASSWORD    = "Admin1234!"


def make_driver():
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--window-size=1400,900")
    opts.add_argument("--disable-gpu")
    return webdriver.Edge(options=opts)


def login(driver):
    driver.get(f"{JENKINS_URL}/login")
    wait = WebDriverWait(driver, 15)
    wait.until(EC.presence_of_element_located((By.ID, "j_username")))
    driver.find_element(By.ID, "j_username").send_keys(USER)
    driver.find_element(By.NAME, "j_password").send_keys(PASSWORD)
    driver.find_element(By.NAME, "Submit").click()
    time.sleep(3)
    # Verificar que ya no estamos en /login
    if "/login" not in driver.current_url:
        print(f"Login OK — {driver.current_url}")
    else:
        raise Exception(f"Login falló, URL: {driver.current_url}")


def shot(driver, name: str, url: str = None, wait_sel: str = None, delay=2) -> pathlib.Path:
    if url:
        driver.get(url)
    if wait_sel:
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, wait_sel))
            )
        except Exception:
            pass
    time.sleep(delay)
    path = SHOTS_DIR / f"{name}.png"
    driver.save_screenshot(str(path))
    print(f"  Captura: {path.name}  ({path.stat().st_size//1024} KB)")
    return path


def capture_all():
    driver = make_driver()
    screenshots = {}
    try:
        login(driver)

        # 1. Dashboard principal
        screenshots["dashboard"] = shot(
            driver, "01_dashboard",
            f"{JENKINS_URL}/",
            wait_sel=".job-status-nobuilt, .jenkins-jobs-list, #projectstatus",
        )

        # 2. Job view (historial de builds)
        screenshots["job_view"] = shot(
            driver, "02_historial_builds",
            f"{JENKINS_URL}/job/inventario-pipeline/",
            wait_sel=".build-row, #buildHistory, .jenkins-build-caption",
        )

        # 3. Stage View (build #5 exitoso)
        screenshots["stage_view"] = shot(
            driver, "03_stage_view",
            f"{JENKINS_URL}/job/inventario-pipeline/5/",
            wait_sel=".pipeline-stage-view, .stage-flownode, h1",
            delay=3,
        )

        # 4. Console output build #5 (SUCCESS)
        screenshots["console_success"] = shot(
            driver, "04_consola_build5_success",
            f"{JENKINS_URL}/job/inventario-pipeline/5/console",
            wait_sel=".console-output, pre.console-output",
            delay=2,
        )

        # 5. Console output build #3 (FAILURE)
        screenshots["console_failure"] = shot(
            driver, "05_consola_build3_failure",
            f"{JENKINS_URL}/job/inventario-pipeline/3/console",
            wait_sel=".console-output, pre.console-output",
            delay=2,
        )

        # 6. Credentials
        screenshots["credentials"] = shot(
            driver, "06_credenciales",
            f"{JENKINS_URL}/manage/credentials/store/system/domain/_/",
            wait_sel=".credentials-store-actions, table, .jenkins-table",
            delay=2,
        )

        # 7. Plugins instalados
        screenshots["plugins"] = shot(
            driver, "07_plugins_instalados",
            f"{JENKINS_URL}/manage/pluginManager/installed",
            wait_sel="#filter-box, .jenkins-search__input, table",
            delay=3,
        )

        # 8. Manage Jenkins (overview)
        screenshots["manage"] = shot(
            driver, "08_manage_jenkins",
            f"{JENKINS_URL}/manage/",
            wait_sel=".jenkins-manage-entry, h1",
        )

    finally:
        driver.quit()

    return screenshots


def insert_screenshots_in_docx(screenshots: dict):
    doc = Document(str(DOCX_FILE))

    # ── Buscar sección 7 (Evidencias) y agregar imágenes ──────────────────────
    # Encontrar el último párrafo y agregar sección nueva de evidencias visuales

    # Separador
    doc.add_paragraph()
    sep = doc.add_paragraph()
    sep_run = sep.add_run("─" * 60)
    sep_run.font.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
    sep_run.font.size = Pt(8)

    # Título sección capturas
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(12)
    r = h.add_run("7.6  Capturas de Pantalla — Interfaz Jenkins")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)

    captions = {
        "dashboard":       ("Figura 1", "Dashboard principal de Jenkins con el job inventario-pipeline"),
        "job_view":        ("Figura 2", "Historial de builds — 5 ejecuciones registradas"),
        "stage_view":      ("Figura 3", "Vista de etapas (Stage View) — Build #5 exitoso: Checkout | Build | Test"),
        "console_success": ("Figura 4", "Salida de consola — Build #5 (SUCCESS): 6 pruebas pasadas + integración OK"),
        "console_failure": ("Figura 5", "Salida de consola — Build #3 (FAILURE): fallo intencional en test_health_retorna_ok"),
        "credentials":     ("Figura 6", "Jenkins Credentials Store — credenciales inventario-db-password e inventario-secret-key"),
        "plugins":         ("Figura 7", "Plugins instalados: Pipeline, Git, GitHub, Docker Pipeline, JCasC, Job DSL"),
        "manage":          ("Figura 8", "Panel de administración Jenkins (Manage Jenkins)"),
    }

    for key, path in screenshots.items():
        if not path.exists():
            continue
        label, caption_text = captions.get(key, ("Figura", key))

        # Párrafo con la imagen
        p_img = doc.add_paragraph()
        p_img.paragraph_format.space_before = Pt(10)
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_img = p_img.add_run()
        run_img.add_picture(str(path), width=Inches(5.8))

        # Caption
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(14)
        r1 = p_cap.add_run(f"{label}: ")
        r1.bold = True
        r1.font.size = Pt(9)
        r1.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        r2 = p_cap.add_run(caption_text)
        r2.italic = True
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    doc.save(str(DOCX_FILE))
    print(f"\nDocumento actualizado con capturas: {DOCX_FILE}")
    print(f"Tamaño final: {DOCX_FILE.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    print("=== Capturando screenshots de Jenkins ===")
    shots = capture_all()
    print(f"\n{len(shots)} capturas tomadas")
    print("\n=== Insertando en documento Word ===")
    insert_screenshots_in_docx(shots)
    print("\nCompletado.")
