import streamlit as st
from streamlit_local_storage import LocalStorage
from O365 import Account
from datetime import datetime, timedelta
import os
import json

# --- 1. CONFIGURACIÓN DE PÁGINA Y ESTÉTICA ---
st.set_page_config(page_title="Movilidad Terreno CCCM", layout="wide", page_icon="📍")

# Ocultar menús de Share, GitHub, Edit y Manage App para los usuarios
st.markdown("""
    <style>
        .viewerBadge_container__1QSob {display: none !important;}
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .stAppToolbar {display: none !important;}
        .block-container {padding-top: 2rem;}
        /* Estilo para el botón de sincronización en el sidebar */
        .stButton>button {width: 100%; border-radius: 5px; height: 3em; background-color: #f0f2f6;}
    </style>
""", unsafe_allow_html=True)

# --- 2. CONFIGURACIÓN DE SEGURIDAD (Azure & Token) ---
try:
    client_id = st.secrets["client_id"]
    client_secret = st.secrets["client_secret"]
    tenant_id = st.secrets["tenant_id"]
    
    credentials = (client_id, client_secret)
    protocolo_scopes = ['mail.send', 'calendars.readwrite']
    
    # Inyectar el token desde Secrets al sistema de archivos de la nube
    if "token" in st.secrets:
        with open("o365_token.txt", "w") as f:
            f.write(st.secrets["token"]["token_data"])
    
    account = Account(credentials, tenant_id=tenant_id)
    
except Exception as e:
    st.error(f"⚠️ Error de configuración técnica: {e}")
    st.stop()

# --- 3. OPCIONES DEL FORMULARIO ---
AREAS = ["Contable", "Logística", "Recursos humanos", "Director global de programas", "Gestión de calidad", "Control interno", "Proyectos", "Género", "Internacional", "Administrativo", "Seguridad", "Gestión del conocimiento", "Comunicaciones", "Desarrollo", "AIV", "ERAE", "Compras", "Operaciones"]
TRANSPORTES = ["Aéreo", "Terrestre (vehículo CCCM)", "Terrestre (vehículo particular)", "Moto (CCCM)", "Moto (particular)", "Terrestre (transporte mular)", "Terrestre (a pie)", "Fluvial/marítimo (CCCM)", "Fluvial/marítimo (particular)", "Transporte publico (Ej. bus)"]
RIESGOS = ["Bajo (Ruta habitual)", "Medio (Requiere monitoreo)", "Alto (Zona Crítica)"]

# --- 4. INTERFAZ DE USUARIO ---
st.title("📍 Sistema de Reporte de Movilidad")
st.caption("Organización Colombiana de Respuesta a Minas (CCCM)")

localS = LocalStorage()

with st.form("main_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Información del Personal")
        nombres = st.text_input("1. Nombres y apellidos")
        area = st.selectbox("2. Área a la que pertenece", options=AREAS)
        jefe = st.text_input("3. Jefe inmediato")
        st.subheader("⏱️ Tiempos de Misión")
        f_salida = st.date_input("5. Fecha inicio de misión (salida)", value=datetime.now())
        f_llegada = st.datetime_input("6. Fecha y hora estimada de llegada", value=datetime.now() + timedelta(hours=2))
        f_retorno = st.datetime_input("7. Fecha finalización misión (retorno)", value=datetime.now() + timedelta(days=1))
    
    with col2:
        st.subheader("🗺️ Destino y Motivo")
        origen = st.text_input("9. Lugar de Origen")
        destino = st.text_input("10. Destino (Municipio, veredas, etc.)")
        motivo = st.text_area("8. Objetivo de la misión")
        riesgo = st.selectbox("13. Evaluación de riesgo de la ruta", options=RIESGOS)
    
    st.divider()
    st.subheader("🚌 Transporte y Notificación")
    c1, c2 = st.columns(2)
    with c1:
        modo = st.selectbox("11. Tipo de transporte", options=TRANSPORTES)
        detalles_t = st.text_input("12. Detalles del transporte (Placa, modelo o número)")
    with c2:
        emergencia = st.text_input("16. Contacto de emergencia (Nombre y Teléfono)")
        correos_v = st.text_input("17. Correo adicional para notificación (opcional)")
    
    enviar = st.form_submit_button("💾 GUARDAR REPORTE (MODO OFFLINE)")
    
    if enviar:
        if not nombres or not destino:
            st.error("Por favor completa los campos obligatorios (Nombre y Destino).")
        else:
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
            st.success(f"✅ Reporte guardado en el dispositivo. Alarma programada para las {f_llegada.strftime('%H:%M')}")

# --- 5. PANEL DE CONTROL (SIDEBAR) ---
st.sidebar.image("https://www.colombiasinminas.org/wp-content/uploads/2021/04/Logo-CCCM-Web.png", width=150)
st.sidebar.header("Sincronización")
st.sidebar.info("Use este botón cuando tenga conexión a internet para enviar sus reportes guardados.")

if st.sidebar.button("🔄 SINCRONIZAR AHORA"):
    pendientes = localS.getAll()
    if pendientes:
        if not account.is_authenticated:
            account.authenticate(scopes=protocolo_scopes)
            
        for clave, r in list(pendientes.items()):
            if not isinstance(r, dict) or 'llegada_iso' not in r: continue
            try:
                dt_llegada = datetime.fromisoformat(r['llegada_iso'])
                
                # A. CREAR EVENTO EN CALENDARIO (ALARMA)
                cal = account.schedule().get_default_calendar()
                ev = cal.new_event()
                ev.subject = f"🚨 ALARMA LLEGADA: {r.get('nombres')}"
                ev.start = dt_llegada
                ev.end = dt_llegada + timedelta(minutes=30)
                ev.remind_before_minutes = 15
                ev.save()
                
                # B. ENVIAR CORREO ELECTRÓNICO CON TABLA PROFESIONAL
                msg = account.new_message()
                destinos = [c.strip() for c in r.get('correos', '').split(",") if "@" in c]
                destinos.append("gerente.seguridad@colombiasinminas.org")
                
                msg.to.add(destinos)
                msg.subject = f"NUEVO REPORTE DE MOVILIDAD: {r.get('nombres')}"
                
                tabla_html = f"""
                <html>
                <body style="font-family: Arial, sans-serif; color: #333;">
                    <div style="max-width: 600px; border: 1px solid #1a4a7a; padding: 20px; border-radius: 10px;">
                        <h2 style="color: #1a4a7a; text-align: center;">Reporte de Movilización Terreno</h2>
                        <table style="width: 100%; border-collapse: collapse;">
                            <tr><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Funcionario:</b></td><td>{r.get('nombres')}</td></tr>
                            <tr style="background:#f2f2f2;"><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Área:</b></td><td>{r.get('area')}</td></tr>
                            <tr><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Jefe Inmediato:</b></td><td>{r.get('jefe')}</td></tr>
                            <tr style="background:#f2f2f2;"><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Fecha Salida:</b></td><td>{r.get('salida')}</td></tr>
                            <tr style="background:#fff3cd;"><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Llegada Estimada:</b></td><td><b>{dt_llegada.strftime('%d/%m/%Y %H:%M')}</b></td></tr>
                            <tr style="background:#f2f2f2;"><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Destino:</b></td><td>{r.get('destino')}</td></tr>
                            <tr><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Transporte:</b></td><td>{r.get('transporte')}</td></tr>
                            <tr style="background:#f2f2f2;"><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Riesgo:</b></td><td>{r.get('riesgo')}</td></tr>
                            <tr><td style="padding: 8px; border-bottom: 1px solid #ddd;"><b>Emergencia:</b></td><td>{r.get('emergencia')}</td></tr>
                        </table>
                        <p style="font-size: 0.9em; color: #666; margin-top: 15px;">Objetivo: {r.get('motivo')}</p>
                    </div>
                </body>
                </html>
                """
                msg.body = tabla_html
                msg.content_subtype = 'html'
                msg.send()

                localS.deleteItem(clave)
                st.sidebar.success(f"✅ Enviado: {r.get('nombres')}")
            except Exception as e:
                st.sidebar.error(f"Error al procesar reporte: {e}")
    else:
        st.sidebar.info("No hay reportes pendientes por sincronizar.")