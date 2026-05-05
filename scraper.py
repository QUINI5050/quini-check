import re
import time
import json
import os
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

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
# FUNCIONES DE HISTORIAL LOCAL
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
        return (
            datos.get("resultados", {}),
            datos.get("pozos", {}),
            datos.get("info_sorteo", {})
        )
    return {}, {}, {}

def obtener_sorteos_guardados():
    historial = cargar_historial()
    return sorted(historial.keys(), reverse=True)

# ============================================================
# FUNCIONES DE IMPORTACIÓN/EXPORTACIÓN EXCEL
# ============================================================

def importar_desde_excel(archivo_excel="sorteos_historicos.xlsx"):
    importados = 0
    errores = []
    
    try:
        import pandas as pd
    except ImportError:
        return 0, ["Se requiere pandas. Instalalo con: pip install pandas openpyxl"]
    
    try:
        df = pd.read_excel(archivo_excel)
        
        columnas_obligatorias = ["numero_sorteo", "fecha"]
        for col in columnas_obligatorias:
            if col not in df.columns:
                return 0, [f"Falta la columna obligatoria: '{col}'"]
        
        historial = cargar_historial()
        
        for index, fila in df.iterrows():
            try:
                numero = str(int(fila["numero_sorteo"])).strip()
                fecha = str(fila["fecha"]).strip()
                
                resultados = {}
                pozos = {}
                
                modalidades_nombres = ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]
                
                for mod in modalidades_nombres:
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
                        else:
                            if len(numeros) == 6:
                                resultados[mod] = sorted(numeros)
                
                pozo_columnas = {
                    "Tradicional": "pozo_tradicional",
                    "La Segunda": "pozo_segunda",
                    "Revancha": "pozo_revancha",
                    "Siempre Sale": "pozo_siempre_sale",
                    "Premio Extra": "pozo_extra"
                }
                
                for mod, col_pozo in pozo_columnas.items():
                    if col_pozo in df.columns and pd.notna(fila[col_pozo]):
                        pozos[mod] = {
                            "monto": str(fila[col_pozo]).strip(),
                            "estado": "desconocido",
                            "ganadores": 0,
                            "aciertos_ganadores": 0
                        }
                    else:
                        pozos[mod] = {
                            "monto": "No disponible",
                            "estado": "desconocido",
                            "ganadores": 0,
                            "aciertos_ganadores": 0
                        }
                
                if len(resultados) >= 4:
                    info_sorteo = {
                        "numero": f"N° {numero}",
                        "fecha": fecha,
                        "texto_completo": f"Sorteo N° {numero} — {fecha}"
                    }
                    
                    historial[numero] = {
                        "resultados": resultados,
                        "pozos": pozos,
                        "info_sorteo": info_sorteo,
                        "fecha_consulta": datetime.now().strftime("%d/%m/%Y %H:%M")
                    }
                    importados += 1
                else:
                    errores.append(f"Fila {index+2}: Sorteo {numero} - Solo {len(resultados)} modalidades")
                    
            except Exception as e:
                errores.append(f"Fila {index+2}: Error - {str(e)}")
        
        guardar_historial(historial)
        
    except FileNotFoundError:
        return 0, [f"Archivo '{archivo_excel}' no encontrado."]
    except Exception as e:
        return 0, [f"Error al leer el archivo: {str(e)}"]
    
    return importados, errores


def exportar_historial_a_excel(archivo_salida="historial_completo.xlsx"):
    import pandas as pd
    
    historial = cargar_historial()
    
    if not historial:
        return False, "No hay sorteos en el historial"
    
    filas = []
    for numero, datos in historial.items():
        fila = {
            "numero_sorteo": int(numero),
            "fecha": datos.get("info_sorteo", {}).get("fecha", ""),
        }
        
        resultados = datos.get("resultados", {})
        for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]:
            if mod in resultados:
                fila[mod] = ", ".join(str(n).zfill(2) for n in resultados[mod])
            else:
                fila[mod] = ""
        
        pozos = datos.get("pozos", {})
        pozo_columnas = {
            "Tradicional": "pozo_tradicional",
            "La Segunda": "pozo_segunda",
            "Revancha": "pozo_revancha",
            "Siempre Sale": "pozo_siempre_sale",
            "Premio Extra": "pozo_extra"
        }
        for mod, col in pozo_columnas.items():
            if mod in pozos:
                fila[col] = pozos[mod].get("monto", "") if isinstance(pozos[mod], dict) else str(pozos[mod])
            else:
                fila[col] = ""
        
        filas.append(fila)
    
    df = pd.DataFrame(filas)
    df = df.sort_values("numero_sorteo", ascending=False)
    df.to_excel(archivo_salida, index=False)
    
    return True, archivo_salida

# ============================================================
# FUNCIONES DE SCRAPING
# ============================================================

def _crear_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--user-agent=Mozilla/5.0")
    chrome_options.add_argument("--ignore-certificate-errors")
    
    # Intentar usar Chromium (típico en Streamlit Cloud / Linux)
    try:
        chrome_options.binary_location = "/usr/bin/chromium-browser"
        service = Service("/usr/bin/chromedriver")
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception:
        # Fallback para Windows local con ChromeDriver
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=chrome_options)def _extraer_numeros_de_lista(texto):
    todos = re.findall(r'\b(\d{1,2})\b', texto)
    numeros = []
    for n in todos:
        num = int(n)
        if 1 <= num <= 45 and num not in numeros:
            numeros.append(num)
    return numeros

def _buscar_modalidad_en_texto(texto_completo, nombre_modalidad):
    texto_normalizado = texto_completo.lower()
    nombre_normalizado = nombre_modalidad.lower()
    pos = texto_normalizado.find(nombre_normalizado)
    if pos == -1:
        return None
    fragmento = texto_completo[pos:pos+600]
    return _extraer_numeros_de_lista(fragmento)

def _analizar_estado_pozo(texto_bloque, modalidad):
    resultado = {
        "monto": "No disponible",
        "estado": "desconocido",
        "ganadores": 0,
        "aciertos_ganadores": 0
    }
    
    try:
        match_monto = re.search(r'1[°º]\s*Premio\s+([\d\.]+,\d{2})', texto_bloque)
        if match_monto:
            resultado["monto"] = f"$ {match_monto.group(1)}"
        
        if re.search(r'VACANTE', texto_bloque, re.IGNORECASE):
            resultado["estado"] = "VACANTE"
            return resultado
        
        lineas = texto_bloque.split('\n')
        for linea in lineas:
            if '1° Premio' in linea or '1º Premio' in linea:
                
                if modalidad == "Siempre Sale":
                    match = re.search(
                        r'1[°º]\s*Premio\s+([\d\.]+,\d{2})\s+(\d+)\s+(\d+)\s+([\d\.]+,\d{2})',
                        linea
                    )
                    if match:
                        resultado["monto"] = f"$ {match.group(1)}"
                        resultado["aciertos_ganadores"] = int(match.group(2))
                        resultado["ganadores"] = int(match.group(3))
                        resultado["estado"] = "GANADO"
                else:
                    match = re.search(
                        r'1[°º]\s*Premio\s+([\d\.]+,\d{2})\s+(\d+)\s+([\d\.]+,\d{2})',
                        linea
                    )
                    if match:
                        resultado["monto"] = f"$ {match.group(1)}"
                        resultado["ganadores"] = int(match.group(2))
                        resultado["aciertos_ganadores"] = 6
                        resultado["estado"] = "GANADO"
                break
        
        if resultado["estado"] == "desconocido":
            resultado["estado"] = "VACANTE"
    
    except Exception as e:
        print(f"Error analizando estado del pozo para {modalidad}: {e}")
    
    return resultado

def _analizar_estado_pozo_extra(texto_bloque):
    resultado = {
        "monto": "No disponible",
        "estado": "GANADO",
        "ganadores": 0,
        "aciertos_ganadores": 6
    }
    
    try:
        lineas = texto_bloque.split('\n')
        for linea in lineas:
            match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})\s+([\d\.]+)\s+([\d\.]+,\d{2})', linea.strip())
            if match:
                resultado["monto"] = f"$ {match.group(1)}"
                resultado["ganadores"] = int(match.group(2).replace('.', ''))
                resultado["estado"] = "GANADO"
                break
        
        if resultado["ganadores"] == 0:
            for linea in lineas:
                if re.search(r'\d{1,3}(?:\.\d{3})*,\d{2}', linea):
                    partes = linea.strip().split()
                    numeros_validos = [p for p in partes if re.match(r'^[\d\.]+(?:,\d{2})?$', p)]
                    if len(numeros_validos) >= 2:
                        if ',' in numeros_validos[0]:
                            resultado["monto"] = f"$ {numeros_validos[0]}"
                        gan_str = numeros_validos[1].replace('.', '')
                        if gan_str.isdigit():
                            resultado["ganadores"] = int(gan_str)
                            resultado["estado"] = "GANADO"
                            break
    except Exception as e:
        print(f"Error analizando pozo Premio Extra: {e}")
    
    return resultado

def obtener_resultados():
    resultados, pozos, info_sorteo = obtener_resultados_por_numero(numero_sorteo=None)
    if resultados and len(resultados) >= 4:
        guardar_sorteo_actual(resultados, pozos, info_sorteo)
    return resultados, pozos, info_sorteo

def obtener_resultados_por_numero(numero_sorteo=None):
    resultados = {}
    pozos = {}
    numero_sorteo_str = "No disponible"
    fecha_sorteo = "No disponible"
    
    if numero_sorteo and str(numero_sorteo).strip().isdigit():
        num = str(numero_sorteo).strip()
        res_hist, pozos_hist, info_hist = obtener_sorteo_historial(num)
        if res_hist and len(res_hist) >= 4:
            print(f"📂 Sorteo {num} encontrado en historial")
            return res_hist, pozos_hist, info_hist
    
    driver = None
    try:
        driver = _crear_driver()
        driver.get(URL_OFICIAL if not numero_sorteo else f"{URL_OFICIAL}&sorteo={numero_sorteo}")
        
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "b")))
        time.sleep(2)
        
        body_text = driver.find_element(By.TAG_NAME, "body").text
        lineas = body_text.split('\n')
        
        mes_anio = ""
        for linea in lineas:
            if re.match(r'(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+\d{4}', linea, re.IGNORECASE):
                mes_anio = linea.strip()
            match = re.match(r'(Miércoles|Miercoles|Domingo)\s+(\d{1,2})\s*-\s*(\d{3,5})', linea, re.IGNORECASE)
            if match:
                numero_sorteo_str = f"N° {match.group(3)}"
                fecha_sorteo = f"{match.group(1).capitalize()} {match.group(2)}"
                if mes_anio:
                    fecha_sorteo += f" de {mes_anio}"
        
        h3_elements = driver.find_elements(By.TAG_NAME, "h3")
        for h3 in h3_elements:
            titulo = h3.text.strip().upper()
            modalidad_nombre = None
            for clave, valor in MAPEO_MODALIDADES.items():
                if clave in titulo:
                    modalidad_nombre = valor
                    break
            
            if modalidad_nombre:
                numeros = []
                
                # Método 1: Extraer de etiquetas <b>
                try:
                    padre = h3.find_element(By.XPATH, "./..")
                    elementos_b = padre.find_elements(By.TAG_NAME, "b")
                    for b in elementos_b:
                        texto = b.text.strip()
                        if texto.isdigit():
                            num = int(texto)
                            if 1 <= num <= 45 and num not in numeros:
                                numeros.append(num)
                except:
                    pass
                
                # Método 2: Si no encontró 6 números, buscar en el texto
                if modalidad_nombre != "Premio Extra" and len(numeros) < 6:
                    texto_cercano = _buscar_modalidad_en_texto(body_text, modalidad_nombre)
                    if texto_cercano and len(texto_cercano) >= 6:
                        # Tomar solo los primeros 6 números válidos
                        numeros_alt = [n for n in texto_cercano if 1 <= n <= 45 and n not in numeros]
                        for n in numeros_alt:
                            if len(numeros) < 6:
                                numeros.append(n)
                
                if modalidad_nombre == "Premio Extra":
                    if len(numeros) >= 18:
                        resultados[modalidad_nombre] = sorted(numeros[:18])
                    elif len(numeros) >= 6:
                        resultados[modalidad_nombre] = sorted(numeros)
                elif len(numeros) == 6:
                    resultados[modalidad_nombre] = sorted(numeros)
                else:
                    print(f"⚠️ {modalidad_nombre}: solo {len(numeros)} números encontrados: {numeros}")
        
        # Extraer pozos
        bloques_modalidad = {}
        modalidad_actual = None
        for linea in lineas:
            for clave, valor in MAPEO_MODALIDADES.items():
                if clave in linea.strip().upper():
                    modalidad_actual = valor
                    if modalidad_actual not in bloques_modalidad:
                        bloques_modalidad[modalidad_actual] = []
                    break
            if modalidad_actual:
                bloques_modalidad[modalidad_actual].append(linea)
        
        for modalidad, bloque in bloques_modalidad.items():
            texto_bloque = '\n'.join(bloque)
            if modalidad == "Premio Extra":
                pozos[modalidad] = _analizar_estado_pozo_extra(texto_bloque)
            else:
                pozos[modalidad] = _analizar_estado_pozo(texto_bloque, modalidad)
        
        print(f"✅ Modalidades: {list(resultados.keys())}")
        for mod, nums in resultados.items():
            print(f"   {mod}: {nums}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        resultados = {}
        pozos = {}
    finally:
        if driver:
            driver.quit()
    
    if len(resultados) < 4:
        print("→ Intentando fuente alternativa...")
        resultados, pozos, numero_sorteo_str, fecha_sorteo = _fuente_alternativa()
    
    info_sorteo = {
        "numero": numero_sorteo_str,
        "fecha": fecha_sorteo,
        "texto_completo": f"Sorteo {numero_sorteo_str} — {fecha_sorteo}"
    }
    
    if resultados and len(resultados) >= 4:
        guardar_sorteo_actual(resultados, pozos, info_sorteo)
    
    return resultados, pozos, info_sorteo

def _fuente_alternativa():
    import requests
    from bs4 import BeautifulSoup
    
    resultados = {}
    pozos = {}
    numero = "No disponible"
    fecha = "No disponible"
    
    try:
        url = "https://www.quinielas.com.ar/resultados-quini-6.html"
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")
        texto = soup.get_text()
        
        patron = re.search(r'[Ss]orteo\s*(?:N[°º]?\s*)?(\d{3,5})', texto)
        if patron:
            numero = f"N° {patron.group(1)}"
        
        patron_f = re.search(r'(?:domingo|miércoles)\s+(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
        if patron_f:
            fecha = patron_f.group(0).strip()
        
        for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale"]:
            patron_mod = rf'{mod}[\s\S]*?(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})\s+(\d{{1,2}})'
            match = re.search(patron_mod, texto, re.IGNORECASE)
            if match:
                resultados[mod] = sorted([int(match.group(i)) for i in range(1, 7)])
                pozos[mod] = {"monto": "Ver web", "estado": "desconocido", "ganadores": 0, "aciertos_ganadores": 0}
        
        nums_extra = _buscar_modalidad_en_texto(texto, "Premio Extra")
        if nums_extra and len(nums_extra) >= 6:
            resultados["Premio Extra"] = sorted(nums_extra[:18]) if len(nums_extra) >= 18 else sorted(nums_extra)
            pozos["Premio Extra"] = {"monto": "Ver web", "estado": "GANADO", "ganadores": 0, "aciertos_ganadores": 6}
        
    except Exception as e:
        print(f"Error plan B: {e}")
    
    return resultados if len(resultados) >= 4 else {}, pozos, numero, fecha

def obtener_fecha_sorteo_actual():
    hoy = datetime.now()
    dia = hoy.weekday()
    hora = hoy.hour
    
    if dia == 2 and hora >= 21:
        return "Hoy (miércoles)"
    elif dia == 6 and hora >= 21:
        return "Hoy (domingo)"
    elif dia in [3, 4, 5]:
        fecha = hoy - timedelta(days=dia-2)
        return f"Último: miércoles {fecha.strftime('%d/%m')}"
    elif dia in [0, 1]:
        fecha = hoy - timedelta(days=dia+1) if dia == 0 else hoy - timedelta(days=1)
        return f"Último: domingo {fecha.strftime('%d/%m')}"
    return "Hoy a las 21:15"