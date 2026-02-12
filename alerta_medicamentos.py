"""
SISTEMA DE ALERTAS DE MEDICAMENTOS
Versión con diseño personalizado según PDF
Lee datos desde Google Drive
"""

import openpyxl
from datetime import datetime, date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys
import requests

# Configuración desde variables de entorno
GMAIL_USUARIO = os.environ.get('GMAIL_USUARIO')
GMAIL_PASSWORD = os.environ.get('GMAIL_PASSWORD')
EMAIL_DESTINO = os.environ.get('EMAIL_DESTINO')
WHATSAPP_API_KEY = os.environ.get('WHATSAPP_API_KEY', '')

# Archivo Excel
RUTA_EXCEL = "medicamentos_alertas.xlsx"

# Configuración
DIAS_ALERTA = 3
FILA_INICIO = 18

def log(mensaje):
    """Registrar mensajes con timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {mensaje}")

def leer_info_paciente(sheet):
    """Lee la información del paciente desde las celdas específicas"""
    try:
        paciente = sheet['B5'].value or "No especificado"
        responsable = sheet['B9'].value or "No especificado"
        telefono_whatsapp = sheet['I9'].value or ""
        
        if telefono_whatsapp:
            telefono_whatsapp = str(telefono_whatsapp).replace(" ", "").replace("-", "").replace("+", "")
        
        return {
            'paciente': paciente,
            'responsable': responsable,
            'telefono': telefono_whatsapp
        }
    except Exception as e:
        log(f"Error al leer información del paciente: {e}")
        return {
            'paciente': "No especificado",
            'responsable': "No especificado",
            'telefono': ""
        }

def leer_excel_y_buscar_alertas(ruta_archivo):
    """Lee el archivo Excel y busca fechas próximas en columna J desde fila 18"""
    try:
        log(f"Abriendo archivo Excel: {ruta_archivo}")
        workbook = openpyxl.load_workbook(ruta_archivo, data_only=True)
        sheet = workbook.active
        
        info_paciente = leer_info_paciente(sheet)
        log(f"Paciente: {info_paciente['paciente']}")
        log(f"Responsable: {info_paciente['responsable']}")
        
        alertas = []
        fecha_hoy = date.today()
        columna_fecha = 10
        
        log(f"Revisando columna J desde fila {FILA_INICIO}")
        log(f"Buscando fechas con menos de {DIAS_ALERTA} días...")
        
        for fila in range(FILA_INICIO, sheet.max_row + 1):
            celda = sheet.cell(row=fila, column=columna_fecha)
            valor = celda.value
            
            if isinstance(valor, datetime):
                fecha_celda = valor.date()
                dias_restantes = (fecha_celda - fecha_hoy).days
                
                if 0 <= dias_restantes < DIAS_ALERTA:
                    nombre_medicamento = sheet.cell(row=fila, column=1).value or "Medicamento sin nombre"
                    uso_medicamento = sheet.cell(row=fila, column=2).value or "Uso no especificado"
                    
                    alerta = {
                        'fila': fila,
                        'fecha': fecha_celda,
                        'dias_restantes': dias_restantes,
                        'medicamento': str(nombre_medicamento),
                        'uso': str(uso_medicamento)
                    }
                    alertas.append(alerta)
                    log(f"  ⚠️ Alerta: {nombre_medicamento} - Fila {fila}, Fecha: {fecha_celda}, Días: {dias_restantes}")
        
        workbook.close()
        log(f"Total de alertas encontradas: {len(alertas)}")
        return alertas, info_paciente
    
    except FileNotFoundError:
        log(f"❌ ERROR: No se encontró el archivo: {ruta_archivo}")
        return None, None
    except Exception as e:
        log(f"❌ ERROR al leer Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

def crear_html_email_personalizado(alertas, info_paciente):
    """Crea email HTML según diseño del PDF con Font Awesome"""
    num_alertas = len(alertas)
    fecha_revision = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    alertas_hoy = [a for a in alertas if a['dias_restantes'] == 0]
    alertas_manana = [a for a in alertas if a['dias_restantes'] == 1]
    alertas_proximas = [a for a in alertas if a['dias_restantes'] >= 2]
    
    html = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Alertas de Medicamentos</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Segoe UI', sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; }}
        .container {{ max-width: 900px; margin: 0 auto; background: white; border-radius: 20px; overflow: hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }}
        .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px 30px; text-align: center; }}
        .header h1 {{ font-size: 2.2rem; margin-bottom: 10px; }}
        .header p {{ font-size: 1.1rem; opacity: 0.9; }}
        .info-paciente {{ display: grid; grid-template-columns: 150px 1fr; gap: 30px; padding: 40px; background: linear-gradient(to right, #f8f9fa, white); border-bottom: 3px solid #667eea; }}
        .foto-placeholder {{ width: 120px; height: 120px; border-radius: 50%; background: linear-gradient(135deg, #667eea, #764ba2); display: flex; align-items: center; justify-content: center; font-size: 4rem; color: white; border: 4px solid white; box-shadow: 0 5px 15px rgba(0,0,0,0.2); }}
        .datos-paciente {{ display: flex; flex-direction: column; justify-content: center; gap: 15px; }}
        .dato-item {{ display: flex; align-items: center; gap: 15px; }}
        .dato-icon {{ width: 40px; height: 40px; background: linear-gradient(135deg, #667eea, #764ba2); border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: 1.2rem; }}
        .dato-label {{ font-size: 0.85rem; color: #6c757d; text-transform: uppercase; margin-bottom: 3px; }}
        .dato-valor {{ font-size: 1.2rem; font-weight: bold; color: #212529; }}
        .alert-banner {{ background: linear-gradient(135deg, #ff6b6b, #ee5a6f); color: white; padding: 30px; text-align: center; }}
        .alert-banner h2 {{ font-size: 2.5rem; margin-bottom: 10px; }}
        .alertas-container {{ padding: 40px; }}
        .seccion {{ margin-bottom: 40px; }}
        .seccion-titulo {{ font-size: 1.5rem; font-weight: bold; padding: 15px 20px; border-radius: 10px; margin-bottom: 20px; display: flex; align-items: center; gap: 15px; }}
        .seccion-hoy .seccion-titulo {{ background: linear-gradient(135deg, #ff6b6b, #ee5a6f); color: white; }}
        .seccion-manana .seccion-titulo {{ background: linear-gradient(135deg, #ffa502, #ff7f50); color: white; }}
        .seccion-proximas .seccion-titulo {{ background: linear-gradient(135deg, #ffd93d, #ffc107); color: #000; }}
        .medicamento-card {{ background: white; border-radius: 15px; padding: 25px; margin-bottom: 20px; box-shadow: 0 5px 20px rgba(0,0,0,0.1); border-left: 6px solid; }}
        .card-hoy {{ border-color: #ff6b6b; background: linear-gradient(to right, #fff5f5, white); }}
        .card-manana {{ border-color: #ffa502; background: linear-gradient(to right, #fff8f0, white); }}
        .card-proxima {{ border-color: #ffd93d; background: linear-gradient(to right, #fffef0, white); }}
        .medicamento-header {{ display: flex; justify-content: space-between; align-items: start; margin-bottom: 15px; flex-wrap: wrap; gap: 15px; }}
        .medicamento-info {{ flex: 1; min-width: 250px; }}
        .medicamento-nombre {{ font-size: 1.5rem; font-weight: bold; color: #212529; margin-bottom: 8px; }}
        .medicamento-uso {{ font-size: 1rem; color: #6c757d; }}
        .badge-urgencia {{ padding: 10px 20px; border-radius: 50px; font-weight: bold; font-size: 0.95rem; display: inline-flex; align-items: center; gap: 8px; }}
        .badge-hoy {{ background: #ff6b6b; color: white; }}
        .badge-manana {{ background: #ffa502; color: white; }}
        .badge-proxima {{ background: #ffd93d; color: #000; }}
        .medicamento-footer {{ display: flex; justify-content: space-between; margin-top: 15px; padding-top: 15px; border-top: 1px solid #e9ecef; flex-wrap: wrap; gap: 10px; }}
        .fecha-info {{ display: flex; align-items: center; gap: 8px; color: #495057; font-size: 0.95rem; }}
        .footer {{ background: #f8f9fa; padding: 30px; text-align: center; color: #6c757d; border-top: 3px solid #e9ecef; }}
        .footer-info {{ display: flex; justify-content: center; gap: 30px; margin-bottom: 15px; flex-wrap: wrap; }}
        .footer-item {{ display: flex; align-items: center; gap: 8px; font-size: 0.9rem; }}
        @media (max-width: 768px) {{
            body {{ padding: 10px; }}
            .header h1 {{ font-size: 1.6rem; }}
            .info-paciente {{ grid-template-columns: 1fr; gap: 20px; padding: 30px 20px; text-align: center; }}
            .alertas-container {{ padding: 20px; }}
            .medicamento-header {{ flex-direction: column; }}
            .alert-banner h2 {{ font-size: 1.8rem; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-hospital"></i> Sistema de Alertas de Medicamentos</h1>
            <p>Control y Seguimiento Automatizado</p>
        </div>
        
        <div class="info-paciente">
            <div class="foto-placeholder"><i class="fas fa-user-circle"></i></div>
            <div class="datos-paciente">
                <div class="dato-item">
                    <div class="dato-icon"><i class="fas fa-user"></i></div>
                    <div>
                        <div class="dato-label">Paciente</div>
                        <div class="dato-valor">{info_paciente['paciente']}</div>
                    </div>
                </div>
                <div class="dato-item">
                    <div class="dato-icon"><i class="fas fa-user-nurse"></i></div>
                    <div>
                        <div class="dato-label">Responsable</div>
                        <div class="dato-valor">{info_paciente['responsable']}</div>
                    </div>
                </div>
                <div class="dato-item">
                    <div class="dato-icon"><i class="fas fa-phone"></i></div>
                    <div>
                        <div class="dato-label">Contacto</div>
                        <div class="dato-valor">{info_paciente['telefono'] or 'No especificado'}</div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="alert-banner">
            <h2><i class="fas fa-exclamation-triangle"></i> {num_alertas} Medicamentos</h2>
            <p>Requieren Revisión Inmediata</p>
        </div>
        
        <div class="alertas-container">
    """
    
    if alertas_hoy:
        html += '<div class="seccion seccion-hoy"><div class="seccion-titulo"><i class="fas fa-exclamation-circle"></i><span>URGENTE - Revisión HOY</span></div>'
        for alerta in alertas_hoy:
            html += f"""
                <div class="medicamento-card card-hoy">
                    <div class="medicamento-header">
                        <div class="medicamento-info">
                            <div class="medicamento-nombre"><i class="fas fa-pills"></i> {alerta['medicamento']}</div>
                            <div class="medicamento-uso"><i class="fas fa-notes-medical"></i> {alerta['uso']}</div>
                        </div>
                        <span class="badge-urgencia badge-hoy"><i class="fas fa-clock"></i> HOY - Acción Inmediata</span>
                    </div>
                    <div class="medicamento-footer">
                        <div class="fecha-info"><i class="fas fa-calendar-alt"></i><strong>Fecha:</strong> {alerta['fecha'].strftime('%d/%m/%Y')}</div>
                        <div class="fecha-info"><i class="fas fa-hourglass-end"></i><strong>Días:</strong> 0</div>
                    </div>
                </div>
            """
        html += "</div>"
    
    if alertas_manana:
        html += '<div class="seccion seccion-manana"><div class="seccion-titulo"><i class="fas fa-bell"></i><span>IMPORTANTE - Revisión MAÑANA</span></div>'
        for alerta in alertas_manana:
            html += f"""
                <div class="medicamento-card card-manana">
                    <div class="medicamento-header">
                        <div class="medicamento-info">
                            <div class="medicamento-nombre"><i class="fas fa-pills"></i> {alerta['medicamento']}</div>
                            <div class="medicamento-uso"><i class="fas fa-notes-medical"></i> {alerta['uso']}</div>
                        </div>
                        <span class="badge-urgencia badge-manana"><i class="fas fa-clock"></i> MAÑANA - 1 día</span>
                    </div>
                    <div class="medicamento-footer">
                        <div class="fecha-info"><i class="fas fa-calendar-alt"></i><strong>Fecha:</strong> {alerta['fecha'].strftime('%d/%m/%Y')}</div>
                        <div class="fecha-info"><i class="fas fa-hourglass-half"></i><strong>Días:</strong> 1</div>
                    </div>
                </div>
            """
        html += "</div>"
    
    if alertas_proximas:
        html += '<div class="seccion seccion-proximas"><div class="seccion-titulo"><i class="fas fa-calendar-check"></i><span>PRÓXIMAMENTE - Planificar Revisión</span></div>'
        for alerta in alertas_proximas:
            html += f"""
                <div class="medicamento-card card-proxima">
                    <div class="medicamento-header">
                        <div class="medicamento-info">
                            <div class="medicamento-nombre"><i class="fas fa-pills"></i> {alerta['medicamento']}</div>
                            <div class="medicamento-uso"><i class="fas fa-notes-medical"></i> {alerta['uso']}</div>
                        </div>
                        <span class="badge-urgencia badge-proxima"><i class="fas fa-clock"></i> {alerta['dias_restantes']} días</span>
                    </div>
                    <div class="medicamento-footer">
                        <div class="fecha-info"><i class="fas fa-calendar-alt"></i><strong>Fecha:</strong> {alerta['fecha'].strftime('%d/%m/%Y')}</div>
                        <div class="fecha-info"><i class="fas fa-hourglass-start"></i><strong>Días:</strong> {alerta['dias_restantes']}</div>
                    </div>
                </div>
            """
        html += "</div>"
    
    html += f"""
        </div>
        <div class="footer">
            <div class="footer-info">
                <div class="footer-item"><i class="fas fa-clock"></i><span>Revisión: {fecha_revision}</span></div>
                <div class="footer-item"><i class="fas fa-robot"></i><span>Sistema Automatizado</span></div>
                <div class="footer-item"><i class="fas fa-cloud"></i><span>GitHub Actions</span></div>
            </div>
            <p style="margin-top: 15px; font-size: 0.85rem;">Este correo fue generado automáticamente</p>
        </div>
    </div>
</body>
</html>
    """
    
    return html

def enviar_whatsapp(telefono, mensaje, info_paciente):
    """Envía mensaje por WhatsApp"""
    try:
        if not telefono:
            return False
        
        texto = f"🏥 ALERTA MEDICAMENTOS\n👤 {info_paciente['paciente']}\n👨‍⚕️ {info_paciente['responsable']}\n\n{mensaje}"
        url = "https://api.callmebot.com/whatsapp.php"
        params = {'phone': telefono, 'text': texto, 'apikey': WHATSAPP_API_KEY}
        response = requests.get(url, params=params, timeout=10)
        return response.status_code == 200
    except:
        return False

def crear_mensaje_whatsapp(alertas):
    """Crea mensaje resumido para WhatsApp"""
    mensaje = f"⚠️ {len(alertas)} medicamentos requieren revisión:\n\n"
    for i, alerta in enumerate(alertas[:5], 1):
        urgencia = "🔴 HOY" if alerta['dias_restantes'] == 0 else f"🟡 {alerta['dias_restantes']} días"
        mensaje += f"{i}. {alerta['medicamento']}\n   {urgencia} - {alerta['fecha'].strftime('%d/%m/%Y')}\n\n"
    if len(alertas) > 5:
        mensaje += f"...y {len(alertas) - 5} más."
    return mensaje

def enviar_email(destinatario, asunto, cuerpo_html, archivo_adjunto=None):
    """Envía email vía Gmail SMTP"""
    try:
        log("Preparando email...")
        mensaje = MIMEMultipart()
        mensaje['From'] = GMAIL_USUARIO
        mensaje['To'] = destinatario
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo_html, 'html', 'utf-8'))
        
        if archivo_adjunto and os.path.exists(archivo_adjunto):
            with open(archivo_adjunto, 'rb') as archivo:
                parte = MIMEBase('application', 'octet-stream')
                parte.set_payload(archivo.read())
                encoders.encode_base64(parte)
                parte.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(archivo_adjunto)}')
                mensaje.attach(parte)
        
        log("Conectando con Gmail...")
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(GMAIL_USUARIO, GMAIL_PASSWORD)
        servidor.sendmail(GMAIL_USUARIO, destinatario, mensaje.as_string())
        servidor.quit()
        
        log("✅ Email enviado exitosamente!")
        return True
    except Exception as e:
        log(f"❌ ERROR al enviar email: {str(e)}")
        return False

def main():
    """Función principal"""
    log("="*70)
    log("SISTEMA DE ALERTAS DE MEDICAMENTOS - VERSIÓN PERSONALIZADA")
    log("="*70)
    
    if not all([GMAIL_USUARIO, GMAIL_PASSWORD, EMAIL_DESTINO]):
        log("❌ ERROR: Faltan variables de entorno")
        sys.exit(1)
    
    if not os.path.exists(RUTA_EXCEL):
        log(f"❌ ERROR: No se encontró el archivo Excel: {RUTA_EXCEL}")
        sys.exit(1)
    
    alertas, info_paciente = leer_excel_y_buscar_alertas(RUTA_EXCEL)
    
    if alertas is None:
        log("❌ No se pudo leer el archivo Excel")
        sys.exit(1)
    
    if len(alertas) > 0:
        log(f"\n🚨 Se encontraron {len(alertas)} alertas")
        
        cuerpo_html = crear_html_email_personalizado(alertas, info_paciente)
        asunto = f"🏥 ALERTAS: {len(alertas)} Medicamentos - {info_paciente['paciente']}"
        
        if enviar_email(EMAIL_DESTINO, asunto, cuerpo_html, RUTA_EXCEL):
            log(f"✅ Email enviado a: {EMAIL_DESTINO}")
        
        if info_paciente['telefono']:
            mensaje_wa = crear_mensaje_whatsapp(alertas)
            enviar_whatsapp(info_paciente['telefono'], mensaje_wa, info_paciente)
    else:
        log("✅ No se encontraron alertas")
    
    log("="*70)
    log("PROCESO FINALIZADO")
    log("="*70)

if __name__ == "__main__":
    main()
