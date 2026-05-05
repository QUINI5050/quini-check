import re
import time
import json
import os
from datetime import datetime, timedelta

URL_OFICIAL = "https://apps.loteriasantafe.gov.ar:8443/Extractos/paginas/mostrarQuini6.xhtml?display=0"
HISTORIAL_FILE = "historial_sorteos.json"

MAPEO_MODALIDADES = {
    "TRADICIONAL PRIMER SORTEO": "Tradicional",
    "TRADICIONAL LA SEGUNDA DEL QUINI": "La Segunda",
    "REVANCHA": "Revancha",
    "SIEMPRE SALE": "Siempre Sale",
    "PREMIO EXTRA": "Premio Extra"
}

# ============================================================
# HISTORIAL
# ============================================================
def cargar_historial():
    if os.path.exists(HISTORIAL_FILE):
        try:
            with open(HISTORIAL_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def guardar_historial(historial):
    with open(HISTORIAL_FILE, "w", encoding="utf-8") as f:
        json.dump(historial, f, indent=2, ensure_ascii=False)

def guardar_sorteo_actual(resultados, pozos, info_sorteo):
    historial = cargar_historial()
    numero = info_sorteo.get("numero", "").replace("N° ", "").strip()
    if numero and numero.isdigit():
        historial[numero] = {
            "resultados": resultados,
            "pozos": pozos,
            "info_sorteo": info_sorteo,
            "fecha_consulta": datetime.now().strftime("%d/%m/%Y %H:%M")
        }
        guardar_historial(historial)
        return True
    return False

def obtener_sorteo_historial(numero_sorteo):
    historial = cargar_historial()
    numero = str(numero_sorteo).strip()
    if numero in historial:
        datos = historial[numero]
        return datos.get("resultados", {}), datos.get("pozos", {}), datos.get("info_sorteo", {})
    return {}, {}, {}

def obtener_sorteos_guardados():
    return sorted(cargar_historial().keys(), reverse=True)

# ============================================================
# EXCEL
# ============================================================
def importar_desde_excel(archivo_excel="sorteos_historicos.xlsx"):
    importados = 0
    errores = []
    try:
        import pandas as pd
        df = pd.read_excel(archivo_excel)
        if "numero_sorteo" not in df.columns or "fecha" not in df.columns:
            return 0, ["Faltan columnas obligatorias: numero_sorteo, fecha"]
        historial = cargar_historial()
        for _, fila in df.iterrows():
            try:
                numero = str(int(fila["numero_sorteo"])).strip()
                fecha = str(fila["fecha"]).strip()
                resultados = {}
                pozos = {}
                for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]:
                    if mod in df.columns and pd.notna(fila[mod]):
                        nums_str = str(fila[mod]).strip()
                        numeros = []
                        for n in nums_str.split(","):
                            n = n.strip()
                            if n.isdigit():
                                num = int(n)
                                if 1 <= num <= 45 and num not in numeros:
                                    numeros.append(num)
                        if mod == "Premio Extra":
                            if len(numeros) >= 6:
                                resultados[mod] = sorted(numeros[:18]) if len(numeros) >= 18 else sorted(numeros)
                        elif len(numeros) == 6:
                            resultados[mod] = sorted(numeros)
                for mod, col in [("Tradicional","pozo_tradicional"),("La Segunda","pozo_segunda"),("Revancha","pozo_revancha"),("Siempre Sale","pozo_siempre_sale"),("Premio Extra","pozo_extra")]:
                    pozos[mod] = {"monto": str(fila.get(col, "No disponible")), "estado": "desconocido", "ganadores": 0, "aciertos_ganadores": 0}
                if len(resultados) >= 4:
                    historial[numero] = {
                        "resultados": resultados, "pozos": pozos,
                        "info_sorteo": {"numero": f"N° {numero}", "fecha": fecha, "texto_completo": f"Sorteo N° {numero} — {fecha}"},
                        "fecha_consulta": datetime.now().strftime("%d/%m/%Y %H:%M")
                    }
                    importados += 1
            except Exception as e:
                errores.append(str(e))
        guardar_historial(historial)
    except FileNotFoundError:
        return 0, [f"Archivo '{archivo_excel}' no encontrado"]
    except Exception as e:
        return 0, [str(e)]
    return importados, errores

def exportar_historial_a_excel(archivo_salida="historial_completo.xlsx"):
    import pandas as pd
    historial = cargar_historial()
    if not historial: return False, "Sin datos"
    filas = []
    for num, datos in historial.items():
        fila = {"numero_sorteo": int(num), "fecha": datos.get("info_sorteo", {}).get("fecha", "")}
        for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]:
            fila[mod] = ", ".join(str(n).zfill(2) for n in datos.get("resultados", {}).get(mod, []))
        for mod, col in [("Tradicional","pozo_tradicional"),("La Segunda","pozo_segunda"),("Revancha","pozo_revancha"),("Siempre Sale","pozo_siempre_sale"),("Premio Extra","pozo_extra")]:
            p = datos.get("pozos", {}).get(mod, {})
            fila[col] = p.get("monto", "") if isinstance(p, dict) else str(p)
        filas.append(fila)
    pd.DataFrame(filas).sort_values("numero_sorteo", ascending=False).to_excel(archivo_salida, index=False)
    return True, archivo_salida

# ============================================================
# SCRAPING CON PLAYWRIGHT
# ============================================================
def _extraer_numeros_de_lista(texto):
    todos = re.findall(r'\b(\d{1,2})\b', texto)
    numeros = []
    for n in todos:
        num = int(n)
        if 1 <= num <= 45 and num not in numeros:
            numeros.append(num)
    return numeros

def _buscar_modalidad_en_texto(texto_completo, nombre_modalidad):
    pos = texto_completo.lower().find(nombre_modalidad.lower())
    if pos == -1: return None
    return _extraer_numeros_de_lista(texto_completo[pos:pos+600])

def _analizar_estado_pozo(texto_bloque, modalidad):
    resultado = {"monto": "No disponible", "estado": "desconocido", "ganadores": 0, "aciertos_ganadores": 0}
    try:
        m = re.search(r'1[°º]\s*Premio\s+([\d\.]+,\d{2})', texto_bloque)
        if m: resultado["monto"] = f"$ {m.group(1)}"
        if re.search(r'VACANTE', texto_bloque, re.IGNORECASE):
            resultado["estado"] = "VACANTE"; return resultado
        for linea in texto_bloque.split('\n'):
            if '1° Premio' in linea or '1º Premio' in linea:
                if modalidad == "Siempre Sale":
                    m = re.search(r'1[°º]\s*Premio\s+([\d\.]+,\d{2})\s+(\d+)\s+(\d+)\s+([\d\.]+,\d{2})', linea)
                    if m:
                        resultado["monto"] = f"$ {m.group(1)}"
                        resultado["aciertos_ganadores"] = int(m.group(2))
                        resultado["ganadores"] = int(m.group(3))
                        resultado["estado"] = "GANADO"
                else:
                    m = re.search(r'1[°º]\s*Premio\s+([\d\.]+,\d{2})\s+(\d+)\s+([\d\.]+,\d{2})', linea)
                    if m:
                        resultado["monto"] = f"$ {m.group(1)}"
                        resultado["ganadores"] = int(m.group(2))
                        resultado["aciertos_ganadores"] = 6
                        resultado["estado"] = "GANADO"
                break
        if resultado["estado"] == "desconocido": resultado["estado"] = "VACANTE"
    except: pass
    return resultado

def _analizar_estado_pozo_extra(texto_bloque):
    resultado = {"monto": "No disponible", "estado": "GANADO", "ganadores": 0, "aciertos_ganadores": 6}
    try:
        for linea in texto_bloque.split('\n'):
            m = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})\s+([\d\.]+)\s+([\d\.]+,\d{2})', linea.strip())
            if m:
                resultado["monto"] = f"$ {m.group(1)}"
                resultado["ganadores"] = int(m.group(2).replace('.', ''))
                break
    except: pass
    return resultado

def _scrape_page(page, url):
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    time.sleep(2)
    
    body_text = page.inner_text("body")
    lineas = body_text.split('\n')
    
    resultados = {}
    pozos = {}
    numero_sorteo_str = "No disponible"
    fecha_sorteo = "No disponible"
    
    mes_anio = ""
    for linea in lineas:
        if re.match(r'(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+\d{4}', linea, re.IGNORECASE):
            mes_anio = linea.strip()
        match = re.match(r'(Miércoles|Miercoles|Domingo)\s+(\d{1,2})\s*-\s*(\d{3,5})', linea, re.IGNORECASE)
        if match:
            numero_sorteo_str = f"N° {match.group(3)}"
            fecha_sorteo = f"{match.group(1).capitalize()} {match.group(2)}"
            if mes_anio: fecha_sorteo += f" de {mes_anio}"
    
    h3s = page.query_selector_all("h3")
    for h3 in h3s:
        titulo = h3.inner_text().strip().upper()
        modalidad_nombre = next((v for k, v in MAPEO_MODALIDADES.items() if k in titulo), None)
        if modalidad_nombre:
            numeros = []
            try:
                bs = h3.evaluate("""(el) => {
                    const p = el.parentElement;
                    return Array.from(p.querySelectorAll('b')).map(b => b.textContent.trim());
                }""")
                for t in bs:
                    if t.isdigit():
                        num = int(t)
                        if 1 <= num <= 45 and num not in numeros:
                            numeros.append(num)
            except: pass
            
            if modalidad_nombre != "Premio Extra" and len(numeros) < 6:
                txt = _buscar_modalidad_en_texto(body_text, modalidad_nombre)
                if txt:
                    for n in txt:
                        if 1 <= n <= 45 and n not in numeros and len(numeros) < 6:
                            numeros.append(n)
            
            if modalidad_nombre == "Premio Extra":
                resultados[modalidad_nombre] = sorted(numeros[:18]) if len(numeros) >= 18 else sorted(numeros)
            elif len(numeros) == 6:
                resultados[modalidad_nombre] = sorted(numeros)
    
    bloques = {}
    actual = None
    for linea in lineas:
        for clave, valor in MAPEO_MODALIDADES.items():
            if clave in linea.strip().upper():
                actual = valor
                if actual not in bloques: bloques[actual] = []
                break
        if actual: bloques[actual].append(linea)
    
    for mod, bloque in bloques.items():
        txt = '\n'.join(bloque)
        pozos[mod] = _analizar_estado_pozo_extra(txt) if mod == "Premio Extra" else _analizar_estado_pozo(txt, mod)
    
    return resultados, pozos, numero_sorteo_str, fecha_sorteo

def obtener_resultados():
    resultados, pozos, info_sorteo = obtener_resultados_por_numero()
    if resultados and len(resultados) >= 4:
        guardar_sorteo_actual(resultados, pozos, info_sorteo)
    return resultados, pozos, info_sorteo

def obtener_resultados_por_numero(numero_sorteo=None):
    if numero_sorteo and str(numero_sorteo).strip().isdigit():
        res, poz, inf = obtener_sorteo_historial(str(numero_sorteo).strip())
        if res and len(res) >= 4: return res, poz, inf
    
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            url = URL_OFICIAL if not numero_sorteo else f"{URL_OFICIAL}&sorteo={numero_sorteo}"
            resultados, pozos, ns, fs = _scrape_page(page, url)
            browser.close()
    except ImportError:
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
            from webdriver_manager.chrome import ChromeDriverManager
            opts = Options()
            opts.add_argument("--headless=new"); opts.add_argument("--no-sandbox"); opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-gpu"); opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--ignore-certificate-errors")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
            driver.get(URL_OFICIAL if not numero_sorteo else f"{URL_OFICIAL}&sorteo={numero_sorteo}")
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            WebDriverWait(driver, 30).until(EC.presence_of_element_located(("tag name", "b")))
            time.sleep(2)
            body_text = driver.find_element("tag name", "body").text
            lineas = body_text.split('\n')
            resultados, pozos, ns, fs = {}, {}, "No disponible", "No disponible"
            mes_anio = ""
            for linea in lineas:
                if re.match(r'(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+\d{4}', linea, re.IGNORECASE): mes_anio = linea.strip()
                m = re.match(r'(Miércoles|Miercoles|Domingo)\s+(\d{1,2})\s*-\s*(\d{3,5})', linea, re.IGNORECASE)
                if m:
                    ns = f"N° {m.group(3)}"; fs = f"{m.group(1).capitalize()} {m.group(2)}"
                    if mes_anio: fs += f" de {mes_anio}"
            h3s = driver.find_elements("tag name", "h3")
            for h3 in h3s:
                titulo = h3.text.strip().upper()
                mn = next((v for k, v in MAPEO_MODALIDADES.items() if k in titulo), None)
                if mn:
                    nums = []
                    try:
                        for b in h3.find_element("xpath", "./..").find_elements("tag name", "b"):
                            if b.text.strip().isdigit():
                                num = int(b.text.strip())
                                if 1 <= num <= 45 and num not in nums: nums.append(num)
                    except: pass
                    if mn != "Premio Extra" and len(nums) < 6:
                        txt = _buscar_modalidad_en_texto(body_text, mn)
                        if txt:
                            for n in txt:
                                if 1 <= n <= 45 and n not in nums and len(nums) < 6: nums.append(n)
                    if mn == "Premio Extra": resultados[mn] = sorted(nums[:18]) if len(nums) >= 18 else sorted(nums)
                    elif len(nums) == 6: resultados[mn] = sorted(nums)
            bloques = {}; actual = None
            for linea in lineas:
                for clave, valor in MAPEO_MODALIDADES.items():
                    if clave in linea.strip().upper():
                        actual = valor
                        if actual not in bloques: bloques[actual] = []
                        break
                if actual: bloques[actual].append(linea)
            for mod, bloque in bloques.items():
                txt = '\n'.join(bloque)
                pozos[mod] = _analizar_estado_pozo_extra(txt) if mod == "Premio Extra" else _analizar_estado_pozo(txt, mod)
            driver.quit()
        except Exception as e:
            print(f"Error Selenium: {e}")
            resultados, pozos, ns, fs = {}, {}, "No disponible", "No disponible"
    
    if len(resultados) < 4:
        resultados, pozos, ns, fs = _fuente_alternativa()
    
    info = {"numero": ns, "fecha": fs, "texto_completo": f"Sorteo {ns} — {fs}"}
    if resultados and len(resultados) >= 4:
        guardar_sorteo_actual(resultados, pozos, info)
    return resultados, pozos, info

def _fuente_alternativa():
    import requests
    from bs4 import BeautifulSoup
    res, poz, num, fec = {}, {}, "No disponible", "No disponible"
    try:
        r = requests.get("https://www.quinielas.com.ar/resultados-quini-6.html", headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
        t = BeautifulSoup(r.text, "html.parser").get_text()
        p = re.search(r'[Ss]orteo\s*(?:N[°º]?\s*)?(\d{3,5})', t)
        if p: num = f"N° {p.group(1)}"
        p = re.search(r'(?:domingo|miércoles)\s+(\d{2}/\d{2}/\d{4})', t, re.IGNORECASE)
        if p: fec = p.group(0).strip()
        for mod in ["Tradicional","La Segunda","Revancha","Siempre Sale"]:
            m = re.search(rf'{mod}[\s\S]*?(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})', t, re.IGNORECASE)
            if m: res[mod] = sorted([int(m.group(i)) for i in range(1,7)]); poz[mod] = {"monto":"Ver web","estado":"?","ganadores":0,"aciertos_ganadores":0}
        ex = _buscar_modalidad_en_texto(t, "Premio Extra")
        if ex and len(ex)>=6: res["Premio Extra"] = sorted(ex[:18]) if len(ex)>=18 else sorted(ex); poz["Premio Extra"] = {"monto":"Ver web","estado":"GANADO","ganadores":0,"aciertos_ganadores":6}
    except: pass
    return res if len(res)>=4 else {}, poz, num, fec

def obtener_fecha_sorteo_actual():
    hoy = datetime.now()
    d, h = hoy.weekday(), hoy.hour
    if d == 2 and h >= 21: return "Hoy (miércoles)"
    if d == 6 and h >= 21: return "Hoy (domingo)"
    if d in [3,4,5]: return f"Último: miércoles {(hoy - timedelta(days=d-2)).strftime('%d/%m')}"
    if d in [0,1]: return f"Último: domingo {(hoy - timedelta(days=d+1 if d==0 else 1)).strftime('%d/%m')}"
    return "Hoy a las 21:15"
