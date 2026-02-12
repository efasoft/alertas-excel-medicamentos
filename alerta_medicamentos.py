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
    """Crea email HTML según diseño del PDF proporcionado"""
    num_alertas = len(alertas)
    fecha_revision = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    # Diccionarios para traducir meses y días
    meses_es = {
        1: 'ENE', 2: 'FEB', 3: 'MAR', 4: 'ABR', 5: 'MAY', 6: 'JUN',
        7: 'JUL', 8: 'AGO', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DIC'
    }
    
    dias_es = {
        0: 'LUN', 1: 'MAR', 2: 'MIÉ', 3: 'JUE', 4: 'VIE', 5: 'SÁB', 6: 'DOM'
    }
    
    html = f"""
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Control de Medicamentos</title>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ 
            font-family: 'Arial', sans-serif; 
            background: linear-gradient(135deg, #e0e5ec 0%, #f5f7fa 100%); 
            padding: 20px; 
        }}
        .container {{ 
            max-width: 1200px; 
            margin: 0 auto; 
            background: #f5f7fa; 
            border-radius: 20px; 
            overflow: hidden; 
            box-shadow: 0 10px 40px rgba(0,0,0,0.15); 
        }}
        
        /* Header con imagen de píldoras */
        .header {{ 
            background: linear-gradient(rgba(0,0,0,0.3), rgba(0,0,0,0.3)), 
                        url('https://images.unsplash.com/photo-1587854692152-cbe660dbde88?w=1200') center/cover;
            color: white; 
            padding: 80px 30px; 
            text-align: center; 
            position: relative;
        }}
        .header h1 {{ 
            font-size: 3rem; 
            font-weight: bold; 
            text-shadow: 2px 2px 8px rgba(0,0,0,0.5);
            letter-spacing: 2px;
        }}
        
        /* Contenedor de tarjetas de info */
        .info-cards {{ 
            display: grid; 
            grid-template-columns: 1fr 1fr; 
            gap: 30px; 
            padding: 40px; 
            background: #f5f7fa;
        }}
        
        /* Tarjeta del paciente (verde) */
        .card-paciente {{ 
            background: linear-gradient(135deg, #2d5f3f 0%, #3a7d52 100%);
            color: white;
            border-radius: 20px;
            padding: 30px;
            display: flex;
            align-items: center;
            gap: 25px;
            box-shadow: 0 8px 25px rgba(45, 95, 63, 0.3);
        }}
        .card-paciente .foto {{ 
            width: 120px; 
            height: 120px; 
            border-radius: 50%; 
            background: white;
            overflow: hidden;
            border: 4px solid rgba(255,255,255,0.3);
            flex-shrink: 0;
        }}
        .card-paciente .foto img {{ 
            width: 100%; 
            height: 100%; 
            object-fit: cover; 
        }}
        .card-paciente .info {{ 
            flex: 1;
        }}
        .card-paciente .label {{ 
            font-size: 1.1rem; 
            font-weight: bold; 
            color: #d4ff00;
            margin-bottom: 8px;
            text-transform: uppercase;
        }}
        .card-paciente .valor {{ 
            font-size: 1.8rem; 
            font-weight: bold; 
        }}
        
        /* Tarjeta del responsable (azul) */
        .card-responsable {{ 
            background: linear-gradient(135deg, #1e5a8e 0%, #2874b5 100%);
            color: white;
            border-radius: 20px;
            padding: 30px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            gap: 15px;
            box-shadow: 0 8px 25px rgba(30, 90, 142, 0.3);
        }}
        .card-responsable .label {{ 
            font-size: 1.1rem; 
            font-weight: bold; 
            color: #00d4ff;
            text-transform: uppercase;
        }}
        .card-responsable .valor {{ 
            font-size: 1.8rem; 
            font-weight: bold; 
        }}
        .card-responsable .telefono {{ 
            font-size: 1.4rem; 
            margin-top: 5px;
        }}
        
        /* Banner amarillo de advertencia */
        .alert-banner {{ 
            background: linear-gradient(135deg, #f4c430 0%, #ffd700 100%);
            color: #2c2c2c;
            padding: 35px 40px;
            margin: 0 40px 30px 40px;
            border-radius: 20px;
            display: flex;
            align-items: center;
            gap: 25px;
            box-shadow: 0 8px 25px rgba(244, 196, 48, 0.3);
        }}
        .alert-banner .icon {{ 
            font-size: 5rem;
        }}
        .alert-banner .texto {{ 
            flex: 1;
            font-size: 1.6rem;
            font-weight: bold;
            line-height: 1.4;
            text-transform: uppercase;
        }}
        
        /* Container de medicamentos */
        .medicamentos-container {{ 
            padding: 0 40px 40px 40px; 
        }}
        
        /* Tarjeta de medicamento */
        .medicamento-card {{ 
            background: white;
            border-radius: 20px;
            margin-bottom: 25px;
            box-shadow: 0 8px 25px rgba(0,0,0,0.1);
            overflow: hidden;
            display: flex;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .medicamento-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 12px 35px rgba(0,0,0,0.15);
        }}
        
        /* Calendario lateral (rojo) */
        .calendario {{ 
            background: linear-gradient(135deg, #c41e3a 0%, #e63946 100%);
            color: white;
            width: 140px;
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
            flex-shrink: 0;
        }}
        .calendario .dia-semana {{ 
            font-size: 1rem; 
            font-weight: bold; 
            margin-bottom: 5px;
        }}
        .calendario .dia {{ 
            font-size: 4rem; 
            font-weight: bold; 
            line-height: 1;
            margin-bottom: 5px;
        }}
        .calendario .mes {{ 
            font-size: 1.2rem; 
            font-weight: bold; 
        }}
        
        /* Contenido del medicamento */
        .medicamento-contenido {{ 
            flex: 1;
            padding: 30px;
            display: flex;
            flex-direction: column;
            gap: 15px;
        }}
        .medicamento-nombre {{ 
            font-size: 2rem; 
            font-weight: bold; 
            color: #2c2c2c;
        }}
        .medicamento-uso {{ 
            font-size: 1.1rem; 
            color: #666;
        }}
        
        /* Badge de días restantes */
        .badge-dias {{ 
            display: inline-block;
            background: linear-gradient(135deg, #ff6b35 0%, #ff8c42 100%);
            color: white;
            padding: 12px 25px;
            border-radius: 50px;
            font-size: 1.2rem;
            font-weight: bold;
            margin-top: 10px;
        }}
        
        /* Footer */
        .footer {{ 
            background: #2c3e50;
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .footer-info {{ 
            display: flex;
            justify-content: center;
            gap: 30px;
            margin-bottom: 15px;
            flex-wrap: wrap;
            font-size: 0.95rem;
        }}
        
        /* Responsive */
        @media (max-width: 768px) {{
            body {{ padding: 10px; }}
            .header h1 {{ font-size: 2rem; }}
            .info-cards {{ grid-template-columns: 1fr; gap: 20px; padding: 20px; }}
            .alert-banner {{ 
                flex-direction: column; 
                margin: 0 20px 20px 20px;
                text-align: center;
            }}
            .alert-banner .icon {{ font-size: 3rem; }}
            .alert-banner .texto {{ font-size: 1.2rem; }}
            .medicamentos-container {{ padding: 0 20px 20px 20px; }}
            .medicamento-card {{ flex-direction: column; }}
            .calendario {{ width: 100%; padding: 15px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header con imagen de píldoras -->
        <div class="header">
            <h1>CONTROL DE MEDICAMENTOS</h1>
        </div>
        
        <!-- Tarjetas de información del paciente y responsable -->
        <div class="info-cards">
            <!-- Tarjeta verde del paciente -->
            <div class="card-paciente">
                <div class="foto">
                    <img src="https://via.placeholder.com/120/3a7d52/ffffff?text=Paciente" alt="Foto paciente">
                </div>
                <div class="info">
                    <div class="label">PACIENTE</div>
                    <div class="valor">{info_paciente['paciente']}</div>
                </div>
            </div>
            
            <!-- Tarjeta azul del responsable -->
            <div class="card-responsable">
                <div>
                    <div class="label">RESPONSABLE</div>
                    <div class="valor">{info_paciente['responsable']}</div>
                </div>
                <div class="telefono">{info_paciente['telefono'] or 'Sin teléfono'}</div>
            </div>
        </div>
        
        <!-- Banner amarillo de advertencia -->
        <div class="alert-banner">
            <div class="icon">✋</div>
            <div class="texto">
                MEDICAMENTOS QUE ESTÁN PRÓXIMOS AGOTARSE Y REQUIEREN ATENCIÓN
            </div>
        </div>
        
        <!-- Lista de medicamentos -->
        <div class="medicamentos-container">
"""
    
    # Generar tarjetas de medicamentos
    for alerta in alertas:
        fecha = alerta['fecha']
        dia_semana = dias_es[fecha.weekday()]
        dia = fecha.day
        mes = meses_es[fecha.month]
        dias_texto = f"Quedan {alerta['dias_restantes']:02d} días" if alerta['dias_restantes'] > 0 else "VENCE HOY"
        
        html += f"""
            <div class="medicamento-card">
                <!-- Calendario lateral -->
                <div class="calendario">
                    <div class="dia-semana">{dia_semana}</div>
                    <div class="dia">{dia}</div>
                    <div class="mes">{mes}</div>
                </div>
                
                <!-- Contenido del medicamento -->
                <div class="medicamento-contenido">
                    <div class="medicamento-nombre">{alerta['medicamento']}</div>
                    <div class="medicamento-uso">{alerta['uso']}</div>
                    <div class="badge-dias">{dias_texto}</div>
                </div>
            </div>
        """
    
    html += f"""
        </div>
        
        <!-- Footer -->
        <div class="footer">
            <div class="footer-info">
                <div>📅 Revisión: {fecha_revision}</div>
                <div>🤖 Sistema Automatizado</div>
                <div>☁️ GitHub Actions</div>
            </div>
            <p style="margin-top: 15px; font-size: 0.85rem; opacity: 0.8;">
                Este correo fue generado automáticamente
            </p>
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
