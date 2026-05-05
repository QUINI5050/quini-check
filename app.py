import streamlit as st
import streamlit.components.v1 as components
import json
from datetime import datetime
from checker import revisar_premios
from notificar import enviar_correo_con_pdf  # Lo usamos para enviar HTML también
from scraper import (
    obtener_resultados,
    obtener_resultados_por_numero,
    obtener_fecha_sorteo_actual,
    obtener_sorteos_guardados,
    importar_desde_excel,
    exportar_historial_a_excel
)

st.set_page_config(page_title="🎰 Quini 6 Checker", page_icon="🎯", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stButton > button { width: 100%; }
    .numero-bola { display: inline-block; width: 40px; height: 40px; line-height: 40px; border-radius: 50%; background: #FFD700; color: #000; font-weight: bold; margin: 3px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    iframe { border: none !important; }
    .badge-vacante { display: inline-block; padding: 4px 12px; border-radius: 12px; background: #e74c3c; color: white; font-weight: bold; font-size: 0.8em; margin-top: 5px; }
    .badge-ganado { display: inline-block; padding: 4px 12px; border-radius: 12px; background: #27ae60; color: white; font-weight: bold; font-size: 0.8em; margin-top: 5px; }
    .badge-desconocido { display: inline-block; padding: 4px 12px; border-radius: 12px; background: #95a5a6; color: white; font-weight: bold; font-size: 0.8em; margin-top: 5px; }
</style>
""", unsafe_allow_html=True)

st.title("🎰 Quini 6 Checker")
st.markdown("### Resultados y control de jugadas")

for key, val in [("ultimo_chequeo", None), ("resultados_cache", None), ("pozos_cache", None), ("info_sorteo_cache", None), ("sorteo_seleccionado", "ultimo"), ("mostrar_detalle", False)]:
    if key not in st.session_state: st.session_state[key] = val

@st.cache_data(ttl=60)
def cargar_jugadas():
    try:
        with open("jugadas.json", "r", encoding="utf-8") as f: return json.load(f)
    except:
        return [
            {"nombre": "Adrian", "numeros": [7, 15, 18, 23, 33, 34]},
            {"nombre": "Carlos", "numeros": [5, 6, 22, 24, 39, 40]},
            {"nombre": "Maxi", "numeros": [2, 9, 10, 14, 19, 41]},
            {"nombre": "Ruben", "numeros": [3, 7, 17, 21, 33, 36]},
        ]

def guardar_jugadas(j):
    with open("jugadas.json", "w", encoding="utf-8") as f: json.dump(j, f, indent=2, ensure_ascii=False)
    st.cache_data.clear()

# ---------- HTML PARA TABLERO EN PANTALLA ----------
def _construir_tablero_html(detalle, resultados):
    modalidades = ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]
    html = """<html><head><meta charset="UTF-8"><style>
*{box-sizing:border-box}body{margin:0;padding:10px;background:#fff;font-family:sans-serif}
.tabla-jugador{width:100%;border-collapse:collapse;margin-bottom:20px;border:2px solid #1a3a5c}
.tabla-jugador th.numero{background:#1a3a5c;color:#fff;padding:14px 6px;font-size:1.15em;font-weight:bold;min-width:48px}
.tabla-jugador th.nombre{background:#1a3a5c;color:#fff;padding:14px 15px;font-size:1.2em;font-weight:bold;letter-spacing:1px}
.tabla-jugador th.aciertos-header{background:#1a3a5c;color:#fff;padding:14px 10px;font-size:1em;font-weight:bold}
.tabla-jugador td{padding:11px 6px;font-size:1.05em;color:#000;font-weight:600;background:#fff;min-width:48px;border:1px solid #d5d8dc}
.tabla-jugador td.modalidad{text-align:left;padding-left:15px;font-weight:bold;color:#1a3a5c;background:#eaf0f8;border-right:3px solid #1a3a5c}
.celda-ok{background:#1b7a3d!important;color:#fff!important;font-weight:bold}
.celda-vacia{background:#fff!important;color:#ccc}
.total-premio{background:#1b7a3d!important;color:#fff!important;font-weight:bold;font-size:1.2em!important}
.total-medio{background:#e67e22!important;color:#fff!important;font-weight:bold;font-size:1.1em!important}
.total-bajo{background:#fff!important;color:#333!important;font-weight:bold}
.total-cero{background:#f5f5f5!important;color:#ccc!important}
</style></head><body>"""
    for jd in detalle:
        nombre = jd["nombre"]
        nums_j = list(jd["modalidades"].values())[0]["numeros_jugados"]
        html += f'<table class="tabla-jugador"><thead><tr><th class="nombre">{nombre}</th>'
        for n in nums_j: html += f'<th class="numero">{str(n).zfill(2)}</th>'
        html += '<th class="aciertos-header">Aciertos</th></tr></thead><tbody>'
        for mod in modalidades:
            if mod in jd["modalidades"]:
                d = jd["modalidades"][mod]; a = d["aciertos"]
                ct = "total-premio" if a>=4 else ("total-medio" if a>=3 else ("total-bajo" if a>0 else "total-cero"))
                nm = "Extra" if mod=="Premio Extra" else mod
                html += f'<tr><td class="modalidad">{nm}</td>'
                for n in nums_j: html += '<td class="celda-ok">OK</td>' if n in d["numeros_sorteados"] else '<td class="celda-vacia">-</td>'
                html += f'<td class="{ct}">{a}</td></tr>'
        html += '</tbody></table>'
    return html + '</body></html>'

# ---------- HTML PARA CORREO (COMPLETO) ----------
def _construir_html_completo(resultados, pozos, info, detalle):
    css = """<style>
body{font-family:Arial,sans-serif;background:#fff;color:#333;margin:0;padding:20px}
h1{color:#1a3a5c;text-align:center;font-size:2em}
h2{color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:5px;font-size:1.4em;margin-top:25px}
h3{color:#2c5aa0;font-size:1.2em}
.contenedor-modalidades{display:flex;flex-wrap:wrap;gap:10px;margin:15px 0}
.tarjeta-modalidad{flex:1;min-width:22%;background:linear-gradient(135deg,#1a3a5c,#2c5aa0);color:#fff;padding:18px;border-radius:12px;text-align:center}
.tarjeta-modalidad h4{font-size:1.2em;text-transform:uppercase;margin:0 0 10px;color:#fff}
.bolita{display:inline-block;width:35px;height:35px;line-height:35px;border-radius:50%;background:#FFD700;color:#000;font-weight:bold;margin:2px;font-size:.9em}
.bolita-extra{display:inline-block;width:28px;height:28px;line-height:28px;border-radius:50%;background:#FFD700;color:#000;font-weight:bold;margin:1px;font-size:.7em}
.pozo-monto{font-size:1.1em;margin-top:10px;font-weight:bold;color:#FFD700;background:rgba(0,0,0,.3);padding:5px 10px;border-radius:8px;display:inline-block}
.info-ganadores{font-size:.85em;margin-top:6px;color:#fff}
.badge-vacante{display:inline-block;padding:5px 14px;border-radius:10px;background:#e74c3c;color:#fff;font-weight:bold;font-size:.9em;margin-top:8px}
.badge-ganado{display:inline-block;padding:5px 14px;border-radius:10px;background:#27ae60;color:#fff;font-weight:bold;font-size:.9em;margin-top:8px}
.tarjeta-extra{background:linear-gradient(135deg,#6d1a8a,#9b59b6);color:#fff;padding:18px;border-radius:12px;text-align:center;margin:15px 0}
.tabla-pdf{width:100%;border-collapse:collapse;margin:10px 0;border:2px solid #1a3a5c}
.tabla-pdf th{background:#1a3a5c;color:#fff;padding:10px 8px;font-weight:bold}
.tabla-pdf td{padding:10px 8px;background:#fff;color:#000;font-weight:600;border:1px solid #d5d8dc}
.tabla-pdf td.mod{text-align:left;color:#1a3a5c;background:#eaf0f8;font-weight:bold}
.ok-pdf{background:#1b7a3d!important;color:#fff!important;font-weight:bold}
.total-verde{background:#1b7a3d!important;color:#fff!important;font-weight:bold;font-size:1.1em}
.total-naranja{background:#e67e22!important;color:#fff!important;font-weight:bold}
.separador{border:none;border-top:1px dashed #ccc;margin:15px 0}
.pie{text-align:center;font-size:.8em;color:#999;margin-top:20px}
</style>"""
    
    html = f"""<html><head><meta charset="UTF-8">{css}</head><body>
<h1>🎰 Quini 6 Checker</h1>
<p style="text-align:center">📌 {info['texto_completo']}</p>
<h2>🏆 NÚMEROS GANADORES</h2>
<div class="contenedor-modalidades">"""
    
    for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale"]:
        if mod in resultados:
            nums = resultados[mod]
            pd = pozos.get(mod, {})
            monto = pd.get("monto","N/D") if isinstance(pd, dict) else str(pd)
            estado = pd.get("estado","?") if isinstance(pd, dict) else "?"
            gan = pd.get("ganadores",0) if isinstance(pd, dict) else 0
            ac = pd.get("aciertos_ganadores",0) if isinstance(pd, dict) else 0
            
            badge = '<span class="badge-vacante">⚠️ VACANTE</span>' if estado=="VACANTE" else (f'<span class="badge-ganado">✅ {gan} gan. ({ac} ac.)</span>' if (estado=="GANADO" and gan>0 and ac>0 and ac<6) else (f'<span class="badge-ganado">✅ {gan} gan.</span>' if (estado=="GANADO" and gan>0) else '<span class="badge-ganado">✅ GANADO</span>'))
            info_gan = f"{gan} gan. con {ac} aciertos" if (estado=="GANADO" and gan>0 and ac>0) else (f"{gan} ganadores" if (estado=="GANADO" and gan>0) else ("Pozo vacante" if estado=="VACANTE" else ""))
            bolitas = " ".join([f'<span class="bolita">{str(n).zfill(2)}</span>' for n in nums])
            html += f'<div class="tarjeta-modalidad"><h4>{mod}</h4><div>{bolitas}</div><div class="pozo-monto">🏆 {monto}</div><div class="info-ganadores">{info_gan}</div><div>{badge}</div></div>'
    
    html += "</div>"
    
    if "Premio Extra" in resultados:
        nums = resultados["Premio Extra"]
        pd = pozos.get("Premio Extra", {})
        monto = pd.get("monto","N/D") if isinstance(pd, dict) else str(pd)
        gan = pd.get("ganadores",0) if isinstance(pd, dict) else 0
        badge = f'<span class="badge-ganado">✅ {gan} ganadores</span>' if gan>0 else '<span class="badge-ganado">✅ Con ganadores</span>'
        bolitas = " ".join([f'<span class="bolita-extra">{str(n).zfill(2)}</span>' for n in nums])
        html += f'<div class="tarjeta-extra"><h4>🎟️ Premio Extra</h4><div>{bolitas}</div><div class="pozo-monto">🏆 {monto}</div><div class="info-ganadores">{gan} ganadores</div><div>{badge}</div></div>'
    
    html += '<h2>📋 TABLERO DE JUGADAS</h2>'
    
    for jd in detalle:
        nombre = jd["nombre"]
        nums_j = list(jd["modalidades"].values())[0]["numeros_jugados"]
        html += f'<h3>🔹 {nombre}</h3><table class="tabla-pdf"><thead><tr><th>Modalidad</th>'
        for n in nums_j: html += f'<th>{str(n).zfill(2)}</th>'
        html += '<th>Aciertos</th></tr></thead><tbody>'
        for mod in ["Tradicional", "La Segunda", "Revancha", "Siempre Sale", "Premio Extra"]:
            if mod in jd["modalidades"]:
                d = jd["modalidades"][mod]; a = d["aciertos"]
                ct = "total-verde" if a>=4 else ("total-naranja" if a>=3 else "")
                nm = "Extra" if mod=="Premio Extra" else mod
                html += f'<tr><td class="mod" style="color:#1a3a5c;background:#eaf0f8;font-weight:bold;text-align:left">{nm}</td>'
                for n in nums_j: html += '<td class="ok-pdf">OK</td>' if n in d["numeros_sorteados"] else '<td>-</td>'
                html += f'<td class="{ct}">{a}</td></tr>'
        html += '</tbody></table><br>'
    
    html += f'<hr class="separador"><p class="pie">Quini 6 Checker · {info["texto_completo"]} · Generado el {datetime.now().strftime("%d/%m/%Y %H:%M")}</p></body></html>'
    return html

# ---------- BARRA LATERAL ----------
with st.sidebar:
    st.header("⚙️ Configuración")
    tabs = st.tabs(["📋 Jugadas", "📧 Email", "🔍 Sorteos", "📥 Importar", "ℹ️ Info"])
    
    with tabs[0]:
        st.subheader("Tus jugadas")
        jugadas = cargar_jugadas()
        for i, j in enumerate(jugadas):
            with st.expander(f"🔹 {j['nombre']}"):
                nn = st.text_input("Nombre", value=j['nombre'], key=f"nom_{i}")
                ns = ", ".join(str(n) for n in j["numeros"])
                nn2 = st.text_input("6 números (1-45)", value=ns, key=f"nums_{i}")
                try:
                    nl = [int(n.strip()) for n in nn2.split(",") if n.strip().isdigit()]
                    if len(nl)==6 and all(1<=n<=45 for n in nl): jugadas[i]["nombre"]=nn; jugadas[i]["numeros"]=nl
                    elif nn2!=ns: st.warning("6 números entre 1 y 45")
                except: st.error("Formato inválido")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Guardar", use_container_width=True): guardar_jugadas(jugadas); st.success("¡Guardadas!")
        with c2:
            if st.button("➕ Nueva", use_container_width=True): jugadas.append({"nombre":f"Jugada {len(jugadas)+1}","numeros":[1,2,3,4,5,6]}); guardar_jugadas(jugadas); st.rerun()
    
    with tabs[1]:
        st.subheader("Configurar email")
        st.session_state["destinatario"] = st.text_input("Mail destinatario", value="desktop.share2021@gmail.com")
        st.session_state["remitente"] = st.text_input("Tu Gmail", value="bertaad736@gmail.com")
        st.session_state["password"] = st.text_input("Contraseña de app", type="password", value = "hmtw lcaq nlni ejqc")
        if st.button("📝 Guardar email"): st.success("Guardado")
    
    with tabs[2]:
        st.subheader("🔍 Sorteos anteriores")
        sg = obtener_sorteos_guardados()
        if sg:
            st.caption(f"📂 {len(sg)} sorteos")
            se = st.selectbox("Seleccioná:", ["Último sorteo"] + [f"Sorteo N° {s}" for s in sg])
            if st.button("📂 Cargar", use_container_width=True):
                if se == "Último sorteo":
                    with st.spinner("Cargando..."): res, poz, inf = obtener_resultados()
                else:
                    with st.spinner(f"Cargando N° {se.replace('Sorteo N° ','')}..."): res, poz, inf = obtener_resultados_por_numero(se.replace("Sorteo N° ",""))
                if res and len(res)>=4:
                    st.session_state.update({"resultados_cache":res, "pozos_cache":poz, "info_sorteo_cache":inf, "ultimo_chequeo":datetime.now().strftime("%d/%m %H:%M"), "mostrar_detalle":False})
                st.rerun()
        
        sb = st.text_input("N° de sorteo", placeholder="3368")
        if st.button("🔍 Buscar") and sb.strip().isdigit():
            with st.spinner(f"Buscando N° {sb}..."): res, poz, inf = obtener_resultados_por_numero(sb.strip())
            if res and len(res)>=4:
                st.session_state.update({"resultados_cache":res, "pozos_cache":poz, "info_sorteo_cache":inf, "ultimo_chequeo":datetime.now().strftime("%d/%m %H:%M"), "mostrar_detalle":False})
                st.rerun()
            else: st.error(f"No encontrado: {sb}")
    
    with tabs[3]:
        st.subheader("📥 Importar sorteos")
        af = st.file_uploader("Subir Excel", type=["xlsx"])
        if af:
            with open("temp.xlsx","wb") as f: f.write(af.getbuffer())
            if st.button("📥 Importar"):
                cant, err = importar_desde_excel("temp.xlsx")
                if cant>0: st.success(f"✅ {cant} sorteos"); st.rerun()
        if st.button("📤 Exportar historial"):
            ok, r = exportar_historial_a_excel()
            st.success(r) if ok else st.error(r)
    
    with tabs[4]:
        st.info(f"📅 {obtener_fecha_sorteo_actual()}")
        st.caption("Sorteos: miércoles y domingos 21:15 hs")

# ============================================================
# PANEL PRINCIPAL
# ============================================================
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    if st.button("🎯 CARGAR ÚLTIMO SORTEO", type="primary", use_container_width=True):
        st.session_state["sorteo_seleccionado"] = "ultimo"
        st.session_state["mostrar_detalle"] = False
        with st.spinner("🔍 Obteniendo..."):
            try:
                res, poz, inf = obtener_resultados()
                if res and len(res) >= 4:
                    st.session_state["resultados_cache"] = res
                    st.session_state["pozos_cache"] = poz
                    st.session_state["info_sorteo_cache"] = inf
                    st.session_state["ultimo_chequeo"] = datetime.now().strftime("%d/%m %H:%M")
                    st.rerun()
                else:
                    st.error("No se pudieron obtener los resultados. Revisá los logs.")
            except Exception as e:
                import traceback
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())

if st.session_state["resultados_cache"]:
    resultados = st.session_state["resultados_cache"]
    pozos = st.session_state["pozos_cache"]
    info = st.session_state["info_sorteo_cache"]
    
    st.markdown("---")
    st.markdown(f"## 🏆 NÚMEROS GANADORES — {info['texto_completo']}")
    
    mods_p = {k:v for k,v in resultados.items() if k!="Premio Extra"}
    mods_e = {k:v for k,v in resultados.items() if k=="Premio Extra"}
    cols = st.columns(4)
    
    for i, mod in enumerate(["Tradicional", "La Segunda", "Revancha", "Siempre Sale"]):
        if mod in mods_p:
            with cols[i]:
                nums = mods_p[mod]
                pd = pozos.get(mod,{})
                monto = pd.get("monto","N/D") if isinstance(pd,dict) else str(pd)
                estado = pd.get("estado","?") if isinstance(pd,dict) else "?"
                gan = pd.get("ganadores",0) if isinstance(pd,dict) else 0
                ac = pd.get("aciertos_ganadores",0) if isinstance(pd,dict) else 0
                
                badge = '<span class="badge-vacante">⚠️ VACANTE</span>' if estado=="VACANTE" else (f'<span class="badge-ganado">✅ {gan} gan. ({ac} ac.)</span>' if (gan>0 and ac>0 and ac<6) else (f'<span class="badge-ganado">✅ {gan} gan.</span>' if gan>0 else '<span class="badge-ganado">✅ GANADO</span>'))
                bol = " ".join([f'<span class="numero-bola">{str(n).zfill(2)}</span>' for n in nums])
                st.markdown(f"""<div style="background:linear-gradient(135deg,#1a3a5c,#2c5aa0);padding:15px;border-radius:12px;text-align:center;color:#fff;margin-bottom:10px"><div style="font-size:.9em;opacity:.9;margin-bottom:8px;text-transform:uppercase">{mod}</div><div style="margin:10px 0">{bol}</div><div style="font-size:.8em;margin-top:8px">🏆 Pozo: {monto}</div><div style="margin-top:5px">{badge}</div></div>""", unsafe_allow_html=True)
    
    if mods_e:
        st.markdown("---")
        for mod, nums in mods_e.items():
            pd = pozos.get(mod,{})
            monto = pd.get("monto","N/D") if isinstance(pd,dict) else str(pd)
            gan = pd.get("ganadores",0) if isinstance(pd,dict) else 0
            badge = f'<span class="badge-ganado">✅ {gan} ganadores</span>' if gan>0 else '<span class="badge-ganado">✅ Con ganadores</span>'
            bol = " ".join([f'<span class="numero-bola" style="font-size:.8em;width:35px;height:35px;line-height:35px">{str(n).zfill(2)}</span>' for n in nums])
            st.markdown(f"""<div style="background:linear-gradient(135deg,#6d1a8a,#9b59b6);padding:15px;border-radius:12px;text-align:center;color:#fff;margin-bottom:10px"><div style="font-size:1em;margin-bottom:8px">🎟️ {mod}</div><div style="margin:10px 0">{bol}</div><div style="font-size:.8em;margin-top:8px">🏆 Pozo: {monto}</div><div style="margin-top:5px">{badge}</div></div>""", unsafe_allow_html=True)
    
    st.markdown("---")
    col_d1, col_d2, col_d3 = st.columns([1, 2, 1])
    with col_d2:
        if not st.session_state["mostrar_detalle"]:
            if st.button("🔍 VER MIS JUGADAS", type="secondary", use_container_width=True): st.session_state["mostrar_detalle"]=True; st.rerun()
        else:
            if st.button("🔼 OCULTAR MIS JUGADAS", type="secondary", use_container_width=True): st.session_state["mostrar_detalle"]=False; st.rerun()
    
    if st.session_state["mostrar_detalle"]:
        st.markdown("---"); st.markdown("## 📋 TABLERO DE JUGADAS")
        jugadas = cargar_jugadas(); detalle = revisar_premios(jugadas, resultados)
        components.html(_construir_tablero_html(detalle, resultados), height=80+len(detalle)*310, scrolling=True)
        tp = sum(1 for j in detalle for d in j["modalidades"].values() if d["aciertos"]>=4)
        if tp>0: st.balloons(); st.success(f"🚨 ¡TOTAL: {tp} premios!")
    
    st.markdown("---")
    col_m1, col_m2, col_m3 = st.columns([1, 2, 1])
    with col_m2:
        if st.button("✉️ ENVIAR POR MAIL", use_container_width=True):
            if not st.session_state.get("password"): st.error("Configurá la contraseña en la barra lateral")
            else:
                with st.spinner("📧 Preparando..."):
                    jugadas = cargar_jugadas(); detalle = revisar_premios(jugadas, resultados)
                    html_completo = _construir_html_completo(resultados, pozos, info, detalle)
                    
                    # Guardar HTML
                    archivo_html = f"quini6_{info['numero'].replace('N° ','').strip()}.html"
                    with open(archivo_html, "w", encoding="utf-8") as f: f.write(html_completo)
                    
                    # Adjuntar HTML como si fuera un archivo (mismo método que PDF)
                    with st.spinner("📧 Enviando mail..."):
                        try:
                            enviar_correo_con_pdf(
                                st.session_state["destinatario"],
                                f"🎰 Quini 6 — {info['texto_completo']}",
                                f"🎯 Resultados Quini 6 - {info['texto_completo']}\n\nAdjuntamos el archivo HTML con el detalle.\nAbrilo en el navegador para verlo correctamente.\n\n¡Saludos!",
                                archivo_html,
                                st.session_state["remitente"],
                                st.session_state["password"]
                            )
                            st.success(f"✅ Mail enviado con {archivo_html}")
                        except Exception as e: st.error(f"❌ Error: {e}")

else:
    st.info("👆 Hacé clic en **CARGAR ÚLTIMO SORTEO**")
    st.markdown("""<div style="text-align:center;padding:40px;background:linear-gradient(135deg,#1a3a5c,#2c5aa0);border-radius:20px;color:#fff;margin:20px 0"><div style="font-size:5em">🎰</div><div style="font-size:1.5em;margin:20px 0">Quini 6 Checker</div></div>""", unsafe_allow_html=True)

st.markdown("---")
st.caption(f"🕐 {st.session_state['ultimo_chequeo'] or 'Nunca'} | 📅 {obtener_fecha_sorteo_actual()}")
