import streamlit as st
from streamlit_local_storage import LocalStorage
from O365 import Account
from datetime import datetime, timedelta
import os
import json

# --- 1. CONFIGURACIÓN DE PÁGINA Y ESTÉTICA ---
st.set_page_config(page_title="Movilidad Terreno CCCM", layout="wide", page_icon="📍")

st.markdown("""
    <style>
        .viewerBadge_container__1QSob {display: none !important;}
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .stAppToolbar {display: none !important;}
        .block-container {padding-top: 2rem;}
        /* Estilos para que los botones se vean bien en móvil */
        .stButton>button {width: 100%; border-radius: 8px; font-weight: bold;}
    </style>
""", unsafe_allow_html=True)

# --- 2. CONFIGURACIÓN DE SEGURIDAD (Azure & Token) ---
try:
    client_id = st.secrets["client_id"]
    client_secret = st.secrets["client_secret"]
    tenant_id = st.secrets["tenant_id"]
    credentials = (client_id, client_secret)
    protocolo_scopes = ['mail.send', 'calendars.readwrite']
    
    if "token" in st.secrets:
        with open("o365_token.txt", "w") as f:
            f.write(st.secrets["token"]["token_data"])
    
    account = Account(credentials, tenant_id=tenant_id)
except Exception as e:
    st.error(f"⚠️ Error de configuración técnica: {e}")
    st.stop()

# --- 3. VARIABLES Y OPCIONES ---
AREAS = ["Contable", "Logística", "Recursos humanos", "Director global de programas", "Gestión de calidad", "Control interno", "Proyectos", "Género", "Internacional", "Administrativo", "Seguridad", "Gestión del conocimiento", "Comunicaciones", "Desarrollo", "AIV", "ERAE", "Compras", "Operaciones"]
TRANSPORTES = ["Aéreo", "Terrestre (vehículo CCCM)", "Terrestre (vehículo particular)", "Moto (CCCM)", "Moto (particular)", "Terrestre (transporte mular)", "Terrestre (a pie)", "Fluvial/marítimo (CCCM)", "Fluvial/marítimo (particular)", "Transporte publico (Ej. bus)"]
RIESGOS = ["Bajo (Ruta habitual)", "Medio (Requiere monitoreo)", "Alto (Zona Crítica)"]

localS = LocalStorage()

# --- 4. INTERFAZ PRINCIPAL ---
st.title("📍 Sistema de Reporte de Movilidad")
st.caption("Organización Colombiana de Respuesta a Minas (CCCM)")

with st.form("main_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("👤 Información del Personal")
        nombres = st.text_input("1. Nombres y apellidos")
        area = st.selectbox("2. Área a la que pertenece", options=AREAS)
        jefe = st.text_input("3. Jefe inmediato")
        st.subheader("⏱️ Tiempos de Misión")
        f_salida = st.date_input("5. Fecha inicio de misión", value=datetime.now())
        f_llegada = st.datetime_input("6. Fecha y hora estimada de llegada", value=datetime.now() + timedelta(hours=2))
        f_retorno = st.datetime_input("7. Fecha finalización misión", value=datetime.now() + timedelta(days=1))
    
    with col2:
        st.subheader("🗺️ Destino y Motivo")
        origen = st.text_input("9. Lugar de Origen")
        destino = st.text_input("10. Destino (Municipio, veredas, etc.)")
        motivo = st.text_area("8. Objetivo de la misión")
        riesgo = st.selectbox("13. Evaluación de riesgo", options=RIESGOS)
    
    st.divider()
    st.subheader("🚌 Transporte y Notificación")
    c1, c2 = st.columns(2)
    with c1:
        modo = st.selectbox("11. Tipo de transporte", options=TRANSPORTES)
        detalles_t = st.text_input("12. Detalles del transporte")
    with c2:
        emergencia = st.text_input("16. Contacto de emergencia")
        correos_v = st.text_input("17. Correo adicional (opcional)")
    
    # BOTÓN DE GUARDAR DENTRO DEL FORMULARIO
    enviar_local = st.form_submit_button("1️⃣ GUARDAR EN EL CELULAR (MODO OFFLINE)")
    
    if enviar_local:
        if not nombres or not destino:
            st.error("⚠️ Completa Nombre y Destino antes de guardar.")
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
            st.success(f"✅ ¡Guardado! Ahora presiona el botón de abajo para enviar.")

# --- 5. BOTÓN DE SINCRONIZACIÓN (FUERA DEL FORMULARIO PARA VER CAMBIOS) ---
st.write("---")
st.subheader("📤 Paso Final: Enviar a la Central")
st.info("Presiona este botón cuando tengas conexión a Internet.")

if st.button("2️⃣ ENVIAR REPORTES Y SINCRONIZAR CON OUTLOOK"):
    pendientes = localS.getAll()
    if pendientes:
        if not account.is_authenticated:
            account.authenticate(scopes=protocolo_scopes)
            
        for clave, r in list(pendientes.items()):
            if not isinstance(r, dict) or 'llegada_iso' not in r: continue
            try:
                dt_llegada = datetime.fromisoformat(r['llegada_iso'])
                
                # Calendario
                cal = account.schedule().get_default_calendar()
                ev = cal.new_event()
                ev.subject = f"🚨 ALARMA LLEGADA: {r.get('nombres')}"
                ev.start = dt_llegada; ev.end = dt_llegada + timedelta(minutes=30); ev.save()
                
                # Correo
                msg = account.new_message()
                destinos = [c.strip() for c in r.get('correos', '').split(",") if "@" in c]
                destinos.append("gerente.seguridad@colombiasinminas.org")
                msg.to.add(destinos)
                msg.subject = f"REPORTE DE MOVILIDAD: {r.get('nombres')}"
                
                tabla_html = f"""
                <html><body style="font-family: Arial; color: #333;">
                <div style="max-width: 600px; border: 2px solid #1a4a7a; padding: 20px; border-radius: 10px;">
                <h2 style="color: #1a4a7a; text-align: center;">Reporte de Movilización</h2>
                <table style="width: 100%; border-collapse: collapse;">
                <tr><td><b>Funcionario:</b></td><td>{r.get('nombres')}</td></tr>
                <tr style="background:#f2f2f2;"><td><b>Área:</b></td><td>{r.get('area')}</td></tr>
                <tr style="background:#fff3cd;"><td><b>Llegada Estimada:</b></td><td><b>{dt_llegada.strftime('%d/%m/%Y %H:%M')}</b></td></tr>
                <tr><td><b>Destino:</b></td><td>{r.get('destino')}</td></tr>
                <tr style="background:#f2f2f2;"><td><b>Riesgo:</b></td><td>{r.get('riesgo')}</td></tr>
                </table>
                <p><b>Objetivo:</b> {r.get('motivo')}</p>
                </div></body></html>
                """
                msg.body = tabla_html; msg.content_subtype = 'html'; msg.send()

                localS.deleteItem(clave)
                st.success(f"🚀 Reporte de {r.get('nombres')} enviado con éxito.")
            except Exception as e:
                st.error(f"❌ Error al enviar: {e}")
    else:
        st.warning("🧐 No hay reportes guardados para enviar. Primero llena el formulario y dale a 'Guardar'.")

# Logotipo institucional al final
st.image("https://www.colombiasinminas.org/wp-content/uploads/2021/04/Logo-CCCM-Web.png", width=120)