import streamlit as st
from streamlit_local_storage import LocalStorage
from O365 import Account
from datetime import datetime, timedelta

# --- 1. CONFIGURACIÓN DE SEGURIDAD (Secrets) ---
try:
    credentials = (st.secrets["client_id"], st.secrets["client_secret"])
    account = Account(credentials, tenant_id=st.secrets["tenant_id"])
    protocolo_scopes = ['mail.send', 'calendars.readwrite']
except Exception:
    st.error("⚠️ Error: Configura 'client_id', 'client_secret' y 'tenant_id' en los Secrets de Streamlit.")
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
        area = st.selectbox("2. Área a la que pertenece", options=AREAS)
        jefe = st.text_input("3. Jefe inmediato")
        
        st.subheader("Tiempos de Misión")
        f_salida = st.date_input("5. Fecha inicio de misión (salida)", value=datetime.now())
        f_llegada = st.datetime_input("6. Fecha y hora estimada de llegada", value=datetime.now() + timedelta(hours=2))
        f_retorno = st.datetime_input("7. Fecha finalización misión (retorno)", value=datetime.now() + timedelta(days=1))

    with col2:
        st.subheader("Destino y Motivo")
        origen = st.text_input("9. Lugar de Origen")
        destino = st.text_input("10. Destino (Municipio, veredas, etc.)")
        motivo = st.text_area("8. Objetivo de la misión")
        riesgo = st.selectbox("13. Evaluación de riesgo de la ruta", options=RIESGOS)

    st.divider()
    st.subheader("Transporte y Notificación")
    c1, c2 = st.columns(2)
    with c1:
        modo = st.selectbox("11. Tipo de transporte", options=TRANSPORTES)
        detalles_t = st.text_input("12. Detalles del transporte (Placa, modelo o número)")
    with c2:
        emergencia = st.text_input("16. Contacto de emergencia (Nombre y Teléfono)")
        correos_v = st.text_input("17. Correo adicional para notificación (opcional)")

    enviar = st.form_submit_button("💾 GUARDAR REPORTE (MODO OFFLINE)")

    if enviar:
        reporte_id = f"mov_{datetime.now().timestamp()}"
        datos = {
            "nombres": nombres, 
            "area": area, 
            "jefe": jefe,
            "salida": f_salida.strftime('%d/%m/%Y'),
            "llegada_iso": f_llegada.isoformat(),
            "retorno": f_retorno.strftime('%d/%m/%Y %H:%M'),
            "origen": origen, 
            "destino": destino, 
            "motivo": motivo,
            "transporte": f"{modo} - {detalles_t}",
            "riesgo": riesgo, 
            "emergencia": emergencia, 
            "correos": correos_v
        }
        localS.setItem(reporte_id, datos)
        st.success(f"✅ Guardado localmente. Alarma para las {f_llegada.strftime('%H:%M')}")

# --- 4. LÓGICA DE SINCRONIZACIÓN ---
st.sidebar.header("Panel de Sincronización")
if st.sidebar.button("🔄 Sincronizar con Outlook y Enviar"):
    pendientes = localS.getAll()
    
    if pendientes:
        if not account.is_authenticated:
            account.authenticate(scopes=protocolo_scopes)
            
        for clave, r in list(pendientes.items()):
            if not isinstance(r, dict) or 'llegada_iso' not in r:
                continue
                
            try:
                dt_llegada = datetime.fromisoformat(r['llegada_iso'])
                
                # A. CALENDARIO OUTLOOK
                cal = account.schedule().get_default_calendar()
                ev = cal.new_event()
                ev.subject = f"🚨 LLEGADA TERRENO: {r.get('nombres')}"
                ev.start = dt_llegada
                ev.end = dt_llegada + timedelta(minutes=30)
                ev.remind_before_minutes = 15
                ev.save()

                # B. ENVÍO DE CORREO
                msg = account.new_message()
                lista_correos = [c.strip() for c in r.get('correos', '').split(",") if "@" in c]
                lista_correos.append("gerente.seguridad@colombiasinminas.org")
                
                msg.to.add(lista_correos)
                msg.subject = f"REPORTE MOVILIDAD: {r.get('nombres')} -> {r.get('destino')}"
                
                # TABLA CON TODOS LOS CAMPOS DILIGENCIADOS
                tabla_html = f"""
                <html>
                <body style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
                    <div style="max-width: 650px; border: 1px solid #eee; padding: 20px; border-radius: 10px;">
                        <h2 style="color: #1a4a7a; border-bottom: 2px solid #1a4a7a; padding-bottom: 10px;">Detalles de Movilización Reportada</h2>
                        
                        <table border="0" style="width: 100%; border-collapse: collapse;">
                            <tr style="background-color: #f2f2f2;"><td style="padding: 8px; width: 40%;"><b>1. Funcionario</b></td><td style="padding: 8px;">{r.get('nombres')}</td></tr>
                            <tr><td style="padding: 8px;"><b>2. Área</b></td><td style="padding: 8px;">{r.get('area')}</td></tr>
                            <tr style="background-color: #f2f2f2;"><td style="padding: 8px;"><b>3. Jefe inmediato</b></td><td style="padding: 8px;">{r.get('jefe')}</td></tr>
                            <tr><td style="padding: 8px;"><b>5. Fecha Salida</b></td><td style="padding: 8px;">{r.get('salida')}</td></tr>
                            <tr style="background-color: #fff3cd;"><td style="padding: 8px;"><b>6. Llegada Estimada (Alarma)</b></td><td style="padding: 8px;"><b>{dt_llegada.strftime('%d/%m/%Y %H:%M')}</b></td></tr>
                            <tr><td style="padding: 8px;"><b>7. Finalización/Retorno</b></td><td style="padding: 8px;">{r.get('retorno')}</td></tr>
                            <tr style="background-color: #f2f2f2;"><td style="padding: 8px;"><b>8. Objetivo de Misión</b></td><td style="padding: 8px;">{r.get('motivo')}</td></tr>
                            <tr><td style="padding: 8px;"><b>9. Origen</b></td><td style="padding: 8px;">{r.get('origen')}</td></tr>
                            <tr style="background-color: #f2f2f2;"><td style="padding: 8px;"><b>10. Destino</b></td><td style="padding: 8px;">{r.get('destino')}</td></tr>
                            <tr><td style="padding: 8px;"><b>11. Transporte</b></td><td style="padding: 8px;">{r.get('transporte')}</td></tr>
                            <tr style="background-color: #f2f2f2;"><td style="padding: 8px;"><b>13. Evaluación de Riesgo</b></td><td style="padding: 8px;">{r.get('riesgo')}</td></tr>
                            <tr><td style="padding: 8px;"><b>16. Emergencia</b></td><td style="padding: 8px;">{r.get('emergencia')}</td></tr>
                            <tr style="background-color: #f2f2f2;"><td style="padding: 8px;"><b>17. Notificados Adicionales</b></td><td style="padding: 8px;">{r.get('correos') if r.get('correos') else 'Ninguo'}</td></tr>
                        </table>
                        
                        <br>
                        <p style="font-size: 0.8em; color: #888; text-align: center;">Este es un reporte automático generado por el Sistema de Movilidad Terreno.</p>
                    </div>
                </body>
                </html>
                """
                msg.body = tabla_html
                msg.content_subtype = 'html'
                msg.send()

                localS.deleteItem(clave)
                st.sidebar.success(f"✔️ Enviado: {r.get('nombres')}")
                
            except Exception as e:
                if "subscriptable" not in str(e):
                    st.sidebar.error(f"Error en envío {clave}: {e}")
    else:
        st.sidebar.info("No hay reportes por sincronizar.")