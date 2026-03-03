import streamlit as st
from streamlit_local_storage import LocalStorage
from O365 import Account
from datetime import datetime, timedelta
import os

# --- 1. CONFIGURACIÓN DE SEGURIDAD (Secrets) ---
try:
    # Cargar credenciales desde st.secrets
    client_id = st.secrets["client_id"]
    client_secret = st.secrets["client_secret"]
    tenant_id = st.secrets["tenant_id"]
    
    credentials = (client_id, client_secret)
    protocolo_scopes = ['mail.send', 'calendars.readwrite']
    
    # LÓGICA PARA EL TOKEN EN LA NUBE
    # Si el token existe en Secrets, lo escribimos en un archivo temporal para la librería O365
    if "token" in st.secrets:
        with open("o365_token.txt", "w") as f:
            f.write(st.secrets["token"]["token_data"])
    
    account = Account(credentials, tenant_id=tenant_id)
    
except Exception as e:
    st.error(f"⚠️ Error de configuración: {e}")
    st.stop()

# --- 2. OPCIONES DEL FORMULARIO ---
AREAS = ["Contable", "Logística", "Recursos humanos", "Director global de programas", "Gestión de calidad", "Control interno", "Proyectos", "Género", "Internacional", "Administrativo", "Seguridad", "Gestión del conocimiento", "Comunicaciones", "Desarrollo", "AIV", "ERAE", "Compras", "Operaciones"]
TRANSPORTES = ["Aéreo", "Terrestre (vehículo CCCM)", "Terrestre (vehículo particular)", "Moto (CCCM)", "Moto (particular)", "Terrestre (transporte mular)", "Terrestre (a pie)", "Fluvial/marítimo (CCCM)", "Fluvial/marítimo (particular)", "Transporte publico (Ej. bus)"]
RIESGOS = ["Bajo (Ruta habitual)", "Medio (Requiere monitoreo)", "Alto (Zona Crítica)"]

# --- 3. INTERFAZ DE USUARIO ---
st.set_page_config(page_title="Movilidad Terreno", layout="wide", page_icon="📍")
st.title("📍 Sistema de Reporte de Movilidad")

localS = LocalStorage()

with st.form("main_form"):
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Información del Personal")
        nombres = st.text_input("1. Nombres y apellidos")
        area = st.selectbox("2. Área", options=AREAS)
        jefe = st.text_input("3. Jefe inmediato")
        st.subheader("Tiempos de Misión")
        f_salida = st.date_input("5. Fecha Salida", value=datetime.now())
        f_llegada = st.datetime_input("6. Llegada estimada (Alarma)", value=datetime.now() + timedelta(hours=2))
        f_retorno = st.datetime_input("7. Finalización (Retorno)", value=datetime.now() + timedelta(days=1))
    with col2:
        st.subheader("Destino y Motivo")
        origen = st.text_input("9. Origen")
        destino = st.text_input("10. Destino")
        motivo = st.text_area("8. Objetivo")
        riesgo = st.selectbox("13. Riesgo", options=RIESGOS)
    st.divider()
    modo = st.selectbox("11. Transporte", options=TRANSPORTES)
    detalles_t = st.text_input("12. Detalles Transporte")
    emergencia = st.text_input("16. Emergencia (Nombre/Tel)")
    correos_v = st.text_input("17. Correo adicional (Opcional)")
    
    enviar = st.form_submit_button("💾 GUARDAR REPORTE (MODO OFFLINE)")
    if enviar:
        reporte_id = f"mov_{datetime.now().timestamp()}"
        datos = {
            "nombres": nombres, "area": area, "jefe": jefe,
            "salida": f_salida.strftime('%d/%m/%Y'),
            "llegada_iso": f_llegada.isoformat(),
            "retorno": f_retorno.strftime('%d/%m/%Y %H:%M'),
            "origen": origen, "destino": destino, "motivo": motivo,
            "transporte": f"{modo} - {detalles_t}",
            "riesgo": riesgo, "emergencia": emergencia, "correos": correos_v
        }
        localS.setItem(reporte_id, datos)
        st.success(f"✅ Guardado localmente. Alarma: {f_llegada.strftime('%H:%M')}")

# --- 4. SINCRONIZACIÓN ---
st.sidebar.header("Panel de Sincronización")
if st.sidebar.button("🔄 Sincronizar con Outlook y Enviar"):
    pendientes = localS.getAll()
    if pendientes:
        # Intentar autenticar usando el archivo generado por los Secrets
        if not account.is_authenticated:
            st.sidebar.warning("Autenticando... Por favor espera.")
            account.authenticate(scopes=protocolo_scopes)
            
        for clave, r in list(pendientes.items()):
            if not isinstance(r, dict) or 'llegada_iso' not in r: continue
            try:
                dt_llegada = datetime.fromisoformat(r['llegada_iso'])
                # CALENDARIO
                cal = account.schedule().get_default_calendar()
                ev = cal.new_event()
                ev.subject = f"🚨 LLEGADA: {r.get('nombres')}"; ev.start = dt_llegada; ev.end = dt_llegada + timedelta(minutes=30); ev.save()
                # CORREO
                msg = account.new_message()
                destinos = [c.strip() for c in r.get('correos', '').split(",") if "@" in c]
                destinos.append("gerente.seguridad@colombiasinminas.org")
                msg.to.add(destinos); msg.subject = f"REPORTE: {r.get('nombres')} -> {r.get('destino')}"
                msg.body = f"<html><body><h2>Reporte de Movilidad</h2><table border='1' style='border-collapse: collapse; width: 100%;'><tr><td><b>Funcionario</b></td><td>{r.get('nombres')}</td></tr><tr><td><b>Área</b></td><td>{r.get('area')}</td></tr><tr><td><b>Destino</b></td><td>{r.get('destino')}</td></tr><tr style='background:#fff3cd;'><td><b>Llegada</b></td><td><b>{dt_llegada.strftime('%d/%m/%Y %H:%M')}</b></td></tr><tr><td><b>Riesgo</b></td><td>{r.get('riesgo')}</td></tr><tr><td><b>Transporte</b></td><td>{r.get('transporte')}</td></tr><tr><td><b>Emergencia</b></td><td>{r.get('emergencia')}</td></tr></table></body></html>"
                msg.content_subtype = 'html'; msg.send()
                localS.deleteItem(clave)
                st.sidebar.success(f"✔️ {r.get('nombres')} enviado")
            except Exception as e:
                st.sidebar.error(f"Error en envío {clave}: {e}")
    else:
        st.sidebar.info("Sin reportes pendientes.")