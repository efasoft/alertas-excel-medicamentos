# 🏥 Sistema de Alertas de Medicamentos

Sistema automatizado que monitorea fechas de vencimiento de medicamentos desde un archivo Excel y envía alertas visuales profesionales por email.

![Vista Desktop](docs/screenshot_desktop.png)
![Vista Email](docs/screenshot_email.png)

---

## 📋 Índice

- [Objetivo del Proyecto](#-objetivo-del-proyecto)
- [Características Principales](#-características-principales)
- [Vista Previa](#-vista-previa)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Configuración](#️-configuración)
- [Uso](#-uso)
- [Estructura del Excel](#-estructura-del-excel)
- [Diseño del Email](#-diseño-del-email)
- [Automatización con GitHub Actions](#-automatización-con-github-actions)
- [Desarrollo](#-desarrollo)
- [Créditos](#-créditos)

---

## 🎯 Objetivo del Proyecto

Crear un sistema automatizado que:
1. Lee datos de medicamentos desde un archivo Excel
2. Identifica medicamentos próximos a vencerse (dentro de 3 días)
3. Extrae la foto del paciente desde el Excel
4. Genera un email HTML con diseño moderno y profesional
5. Envía alertas por email y opcionalmente por WhatsApp
6. Se ejecuta automáticamente con GitHub Actions

---

## ✨ Características Principales

### 🎨 Diseño Visual
- ✅ Header profesional con título "CONTROL DE MEDICAMENTOS"
- ✅ Tarjetas apiladas con separación óptima
- ✅ Foto del paciente extraída automáticamente del Excel
- ✅ Colores sólidos para máxima compatibilidad con clientes de email
- ✅ Tipografía moderna (Montserrat & Raleway)
- ✅ Diseño responsive para móviles
- ✅ Enlaces de teléfono sin formato visual (blancos, sin subrayado)

### 📊 Funcionalidades
- ✅ Lectura automática de datos desde Excel
- ✅ Detección de fechas próximas a vencer (configurable)
- ✅ Extracción de imagen del paciente (base64)
- ✅ Generación de HTML responsive
- ✅ Envío por email vía Gmail SMTP
- ✅ Envío opcional por WhatsApp
- ✅ Logs detallados con timestamps
- ✅ Manejo robusto de errores

### 🔧 Tecnologías
- Python 3.x
- openpyxl (lectura de Excel e imágenes)
- Pillow (procesamiento de imágenes)
- smtplib (envío de emails)
- HTML5 + CSS3 (diseño responsive)
- Google Fonts (tipografías)

---

## 🖼️ Vista Previa

### Vista en Cliente de Email (Desktop)
La primera imagen muestra cómo se ve el Excel original con todos los datos estructurados y la foto del paciente insertada.

### Vista en Cliente de Email (Outlook)
La segunda imagen muestra el email recibido con:
- Header azul con título
- Tarjeta verde del paciente con foto
- Tarjeta azul del responsable con teléfono
- Banner amarillo de advertencia
- Tarjetas de medicamentos con calendarios rojos
- Footer con información del desarrollador

---

## 📦 Requisitos

### Software
```bash
Python 3.8 o superior
pip (gestor de paquetes de Python)
```

### Librerías Python
```bash
openpyxl>=3.0.0
Pillow>=9.0.0
requests>=2.28.0
```

### Cuentas Necesarias
- **Gmail**: Para envío de emails (requiere contraseña de aplicación)
- **CallMeBot** (opcional): Para notificaciones WhatsApp

---

## 🚀 Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/efasoft/alertas-excel-medicamentos.git
cd alertas-excel-medicamentos
```

### 2. Instalar dependencias
```bash
pip install openpyxl Pillow requests
```

O usando el archivo de requisitos:
```bash
pip install -r requirements.txt
```

### 3. Verificar instalación
```bash
python diagnostico_foto.py
```

---

## ⚙️ Configuración

### Variables de Entorno

Crea un archivo `.env` o configura las siguientes variables:

```bash
# Email (obligatorio)
GMAIL_USUARIO=tu_email@gmail.com
GMAIL_PASSWORD=tu_contraseña_de_aplicación
EMAIL_DESTINO=destino@email.com

# WhatsApp (opcional)
WHATSAPP_API_KEY=tu_api_key_callmebot
```

### Configurar Gmail

1. Activa la **verificación en 2 pasos** en tu cuenta de Gmail
2. Genera una **contraseña de aplicación**:
   - Ve a: https://myaccount.google.com/apppasswords
   - Crea una contraseña para "Correo"
   - Usa esa contraseña en `GMAIL_PASSWORD`

### Configurar WhatsApp (opcional)

1. Registra tu número en CallMeBot: https://www.callmebot.com/blog/free-api-whatsapp-messages/
2. Obtén tu API key
3. Configura `WHATSAPP_API_KEY`

---

## 💻 Uso

### Ejecución Manual

```bash
python alerta_medicamentos.py
```

### Ejecución con Variables de Entorno

```bash
export GMAIL_USUARIO="tu_email@gmail.com"
export GMAIL_PASSWORD="tu_contraseña"
export EMAIL_DESTINO="destino@email.com"
python alerta_medicamentos.py
```

### Configurar Parámetros

Edita las constantes al inicio de `alerta_medicamentos.py`:

```python
DIAS_ALERTA = 3        # Días de anticipación para alertas
FILA_INICIO = 18       # Primera fila con datos de medicamentos
```

---

## 📊 Estructura del Excel

### Ubicación de Datos

| Celda | Contenido |
|-------|-----------|
| B5 | Nombre del paciente |
| B9 | Nombre del responsable |
| I9 | Teléfono del responsable |
| L5:M11 | Foto del paciente (imagen insertada) |

### Tabla de Medicamentos (desde fila 18)

| Columna | Contenido |
|---------|-----------|
| A | Nombre del medicamento |
| B | Uso del medicamento |
| J | Fecha de revisión/vencimiento |

### Ejemplo de Estructura

```
Fila 5:  B5=MARIA DEL CARMEN CALDERON  | L5:M11=[FOTO]
Fila 9:  B9=OVIDIA RONDON CALDERON     | I9=611131467
Fila 18: A18=SITAGLIPINA | B18=AZUCAR  | J18=12/02/2026
Fila 19: A19=AMLODIPINO  | B19=TENSION | J19=14/02/2026
```

### Insertar la Foto del Paciente

1. Abre el Excel
2. Selecciona el rango **L5:M11**
3. Ve a **Insertar > Imagen**
4. Selecciona la foto del paciente
5. Ajusta el tamaño para que quede dentro del rango
6. **Guarda el archivo**

---

## 🎨 Diseño del Email

### Estructura Visual

```
┌─────────────────────────────────────┐
│  CONTROL DE MEDICAMENTOS (Header)  │
├─────────────────────────────────────┤
│  [FOTO] PACIENTE                    │
│         MARIA DEL CARMEN CALDERON   │
├─────────────────────────────────────┤
│         RESPONSABLE                 │
│         OVIDIA RONDON CALDERON      │
│         611131467                   │
├─────────────────────────────────────┤
│  ✋ MEDICAMENTOS PRÓXIMOS A         │
│     AGOTARSE Y REQUIEREN ATENCIÓN   │
├─────────────────────────────────────┤
│  JUE │ SITAGLIPINA                  │
│   12 │ AZUCAR                       │
│  FEB │ [VENCE HOY]                  │
├─────────────────────────────────────┤
│  SÁB │ AMLODIPINO                   │
│   14 │ TENSION                      │
│  FEB │ [QUEDAN 02 DÍAS]             │
├─────────────────────────────────────┤
│  Revisión: 12/02/2026 | Sistema    │
│  Desarrollado por: Ernesto +34...   │
└─────────────────────────────────────┘
```

### Paleta de Colores

| Elemento | Color | Hex |
|----------|-------|-----|
| Header | Azul | #667eea |
| Tarjeta Paciente | Verde esmeralda | #059669 |
| Tarjeta Responsable | Azul índigo | #4f46e5 |
| Banner Advertencia | Amarillo | #fbbf24 |
| Calendario | Rojo | #dc2626 |
| Badge Días | Naranja | #f97316 |
| Footer | Gris oscuro | #1e293b |

### Tipografía

- **Títulos**: Montserrat (700-800)
- **Contenido**: Raleway (400-600)
- **Fuente**: Google Fonts

### Responsive

El diseño se adapta automáticamente a:
- **Desktop**: >768px (diseño completo)
- **Móvil**: <768px (tarjetas apiladas, texto reducido)

---

## 🤖 Automatización con GitHub Actions

### Crear Workflow

Crea el archivo `.github/workflows/alertas.yml`:

```yaml
name: Alertas Medicamentos

on:
  schedule:
    # Ejecutar todos los días a las 8:00 AM (UTC)
    - cron: '0 8 * * *'
  workflow_dispatch:  # Permitir ejecución manual

jobs:
  enviar-alertas:
    runs-on: ubuntu-latest
    
    steps:
    - name: Checkout código
      uses: actions/checkout@v3
    
    - name: Configurar Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
    
    - name: Instalar dependencias
      run: |
        pip install openpyxl Pillow requests
    
    - name: Ejecutar script de alertas
      env:
        GMAIL_USUARIO: ${{ secrets.GMAIL_USUARIO }}
        GMAIL_PASSWORD: ${{ secrets.GMAIL_PASSWORD }}
        EMAIL_DESTINO: ${{ secrets.EMAIL_DESTINO }}
        WHATSAPP_API_KEY: ${{ secrets.WHATSAPP_API_KEY }}
      run: |
        python alerta_medicamentos.py
```

### Configurar Secrets

En tu repositorio de GitHub:

1. Ve a **Settings > Secrets and variables > Actions**
2. Click en **New repository secret**
3. Agrega cada secret:
   - `GMAIL_USUARIO`
   - `GMAIL_PASSWORD`
   - `EMAIL_DESTINO`
   - `WHATSAPP_API_KEY` (opcional)

### Frecuencias de Ejecución

```yaml
# Todos los días a las 8 AM
- cron: '0 8 * * *'

# Cada 12 horas (8 AM y 8 PM)
- cron: '0 8,20 * * *'

# Solo días laborables a las 9 AM
- cron: '0 9 * * 1-5'
```

---

## 🛠️ Desarrollo

### Estructura del Proyecto

```
alertas-excel-medicamentos/
├── alerta_medicamentos.py    # Script principal
├── diagnostico_foto.py        # Herramienta de diagnóstico
├── medicamentos_alertas.xlsx  # Archivo Excel de datos
├── test_email.html           # Vista previa del email
├── requirements.txt          # Dependencias Python
├── README.md                # Este archivo
├── .github/
│   └── workflows/
│       └── alertas.yml      # Automatización GitHub Actions
└── docs/
    ├── screenshot_desktop.png
    └── screenshot_email.png
```

### Funciones Principales

```python
# Lectura de Excel
leer_excel_y_buscar_alertas(ruta_archivo)
leer_info_paciente(sheet)
extraer_imagen_paciente(ruta_excel)

# Generación de HTML
crear_html_email_personalizado(alertas, info_paciente)

# Envío de notificaciones
enviar_email(destinatario, asunto, cuerpo_html, archivo_adjunto)
enviar_whatsapp(telefono, mensaje, info_paciente)

# Utilidades
log(mensaje)
crear_mensaje_whatsapp(alertas)
```

### Diagnóstico de Problemas

Si la foto no aparece, ejecuta:

```bash
python diagnostico_foto.py
```

El script te mostrará:
- ✅ Si encuentra imágenes en el Excel
- ✅ La posición exacta de cada imagen
- ✅ Si están en la zona correcta (L5:M11)
- ✅ El tamaño de cada imagen

---

## 🧪 Testing

### Probar Localmente

```bash
# 1. Configurar variables de entorno
export GMAIL_USUARIO="test@test.com"
export GMAIL_PASSWORD="test"
export EMAIL_DESTINO="test@test.com"

# 2. Generar HTML de prueba
python -c "
from alerta_medicamentos import *
alertas, info = leer_excel_y_buscar_alertas('medicamentos_alertas.xlsx')
html = crear_html_email_personalizado(alertas, info)
with open('test_email.html', 'w', encoding='utf-8') as f:
    f.write(html)
print('✓ test_email.html generado')
"

# 3. Abrir en navegador
open test_email.html  # macOS
xdg-open test_email.html  # Linux
start test_email.html  # Windows
```

### Verificar Compatibilidad Email

1. Envía un email de prueba
2. Verifica en diferentes clientes:
   - ✅ Gmail (web y app)
   - ✅ Outlook (desktop y web)
   - ✅ Apple Mail
   - ✅ Thunderbird

---

## 📝 Personalización

### Cambiar Colores

Edita la función `crear_html_email_personalizado()`:

```python
# Ejemplo: cambiar color de tarjeta del paciente
.card-paciente {{ 
    background: #10b981;  # Tu color personalizado
    ...
}}
```

### Cambiar Tipografía

```python
# Cambiar fuentes Google Fonts
<link href="https://fonts.googleapis.com/css2?family=TuFuente:wght@400;700&display=swap" rel="stylesheet">
```

### Ajustar Días de Alerta

```python
DIAS_ALERTA = 5  # Cambiar de 3 a 5 días
```

---

## 🐛 Solución de Problemas

### La foto no aparece

**Causa**: Imagen no insertada correctamente en Excel
**Solución**: 
1. Usa `python diagnostico_foto.py`
2. Inserta la imagen en L5:M11 usando "Insertar > Imagen"
3. Guarda el Excel

### Email no se envía

**Causa**: Credenciales incorrectas o 2FA no configurado
**Solución**:
1. Activa verificación en 2 pasos en Gmail
2. Genera contraseña de aplicación
3. Usa esa contraseña en `GMAIL_PASSWORD`

### Colores no se ven en Outlook

**Causa**: Outlook no soporta algunos CSS
**Solución**: El código ya usa colores sólidos compatibles

### Enlaces de teléfono en azul

**Causa**: Estilo por defecto del navegador/cliente
**Solución**: Ya implementado con `a[href^="tel"] { color: #ffffff !important; }`

---

## 📄 Licencia

Este proyecto es de código abierto y está disponible bajo la licencia MIT.

---

## 👨‍💻 Créditos

**Desarrollado por**: Ernesto Fernandez  
**Contacto**: +34 611131467  
**Email**: efasoft@hotmail.com  
**Fecha**: Febrero 2026  

### Tecnologías Utilizadas

- Python 3.x
- openpyxl
- Pillow (PIL)
- Gmail SMTP
- HTML5 + CSS3
- Google Fonts
- GitHub Actions

---

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

---

## 📞 Soporte

Si tienes preguntas o necesitas ayuda:

- 📧 Email: efasoft@hotmail.com
- 📱 WhatsApp: +34 611131467
- 🐛 Issues: https://github.com/efasoft/alertas-excel-medicamentos/issues

---

## 🔄 Changelog

### v1.0.0 (Febrero 2026)
- ✅ Implementación inicial
- ✅ Extracción de foto del paciente desde Excel
- ✅ Diseño responsive moderno
- ✅ Compatibilidad con clientes de email
- ✅ Automatización con GitHub Actions
- ✅ Enlaces de teléfono sin formato visual
- ✅ Footer con información del desarrollador

---

## 🎯 Roadmap

### Futuras Mejoras

- [ ] Dashboard web para visualización
- [ ] Base de datos para historial de alertas
- [ ] Notificaciones push móviles
- [ ] ML para predicción de consumo
- [ ] Gestión multi-paciente
- [ ] Integración con farmacias
- [ ] App móvil nativa
- [ ] Recordatorios de tomas diarias

---

<div align="center">

**⭐ Si este proyecto te fue útil, dale una estrella en GitHub ⭐**

[Reportar Bug](https://github.com/efasoft/alertas-excel-medicamentos/issues) · [Solicitar Feature](https://github.com/efasoft/alertas-excel-medicamentos/issues)

</div>
