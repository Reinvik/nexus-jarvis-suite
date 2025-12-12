# 🤖 Nexus JARVIS - Logistics Automation Suite

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat&logo=python&logoColor=white)
![React](https://img.shields.io/badge/React-18-61DAFB?style=flat&logo=react&logoColor=white)
![Vercel](https://img.shields.io/badge/Deployment-Vercel-000000?style=flat&logo=vercel&logoColor=white)
![SAP](https://img.shields.io/badge/SAP-GUI_Scripting-008FD3?style=flat&logo=sap&logoColor=white)
![AI](https://img.shields.io/badge/AI-Google_Gemini-8E75B2?style=flat&logo=google&logoColor=white)
![License](https://img.shields.io/badge/License-Proprietary-red?style=flat)

> **Plataforma Web Centralizada para Orquestación y Monitoreo de Bots RPA + IA en Operaciones Logísticas**

---

## 🌐 Demo en Vivo

**🚀 [Ver Aplicación Web](https://nexus-jarvis-7evytdswp-ariels-projects-c0e12d35.vercel.app/)**

---

## 💡 Visión General

**Nexus JARVIS** (Just A Rather Very Intelligent System) es una plataforma Full-Stack que transforma la gestión de procesos logísticos mediante la automatización inteligente. Centraliza y orquesta múltiples bots especializados que interactúan con SAP ERP, procesan documentos con IA y generan reportes analíticos automáticamente.

### 🎯 Problema de Negocio

En entornos logísticos complejos, la dependencia de procesos manuales genera:
- ⏱️ **Ineficiencias operativas** por tareas repetitivas
- ❌ **Errores humanos** en digitación y transcripción
- 🔍 **Falta de visibilidad** sobre el estado de operaciones
- 📊 **Datos dispersos** en múltiples sistemas sin consolidar
- 🚧 **Cuellos de botella** en procesos administrativos

### ✅ Solución Tecnológica

Nexus JARVIS proporciona:
- 🖥️ **Interfaz Web Unificada** para gestionar todos los bots
- 🔄 **Orquestación Centralizada** de flujos de trabajo
- 📈 **Monitoreo en Tiempo Real** del estado de ejecuciones
- 🤖 **Suite de Bots Especializados** para cada proceso crítico
- 🧠 **IA Integrada** para procesamiento de documentos y visión artificial

---

## 🏗️ Arquitectura del Sistema

```
┌─────────────────────────────────────────────────────────────┐
│                    INTERFAZ WEB (Vercel)                    │
│              React + Firebase + Tailwind CSS                │
└────────────────────────┬────────────────────────────────────┘
                         │ HTTP/WebSocket
┌────────────────────────▼────────────────────────────────────┐
│              ORQUESTADOR CENTRAL (Python)                   │
│           worker_sap.py + Firebase Realtime DB              │
└────────────────────────┬────────────────────────────────────┘
                         │
        ┌────────────────┼────────────────┐
        │                │                │
┌───────▼──────┐  ┌──────▼──────┐  ┌─────▼──────┐
│   SAP GUI    │  │  Google AI  │  │   Outlook  │
│  Scripting   │  │   Gemini    │  │    API     │
└──────────────┘  └─────────────┘  └────────────┘
```

---

## 🤖 Suite de Bots Automatizados

### 📦 **Bot Conciliación Email (MIGO Asistido)**
- **Función:** Automatiza carga masiva de movimientos de mercancías en SAP MIGO desde correos electrónicos
- **Tecnología:** Python + win32com (SAP GUI Scripting) + Outlook API
- **Modo:** Asistido (usuario valida antes de contabilizar)
- **Características:**
  - ✅ Extracción automática de datos desde emails
  - ✅ Generación de Excel con validación de lotes
  - ✅ Mapeo dinámico de plantas (P1/P2)
  - ✅ Prevención de duplicados con caché temporal
- **Impacto:** Reduce 90% el tiempo de digitación manual

### 📊 **Bot Auditor de Stock ("Zombies")**
- **Función:** Detecta inventario inmovilizado sin movimientos recientes
- **Tecnología:** Cruce de datos MB52 (stock) vs MB51 (movimientos)
- **Salida:** Reporte Excel con clasificación por días sin movimiento
- **Características:**
  - 🟢 FRESCO (0-2 días)
  - 🟡 PENDIENTE (3-7 días)
  - 🔴 LENTO (8-90 días)
  - 💀 CRÍTICO (>90 días)
- **Impacto:** Previene mermas por obsolescencia y libera capital

### 📐 **Bot Optimizador de Altura (Pallet)**
- **Función:** Genera mapas visuales de ubicaciones en altura desde LX02
- **Tecnología:** Extracción SAP + procesamiento Excel
- **Salida:** Reporte con coordenadas de pallets
- **Impacto:** Optimiza auditorías físicas de almacén

### 🚚 **Bot Monitor de Transporte**
- **Función:** Consolida datos de flota y despachos desde VT11/VT03N
- **Tecnología:** Scraping SAP + consolidación multi-transacción
- **Salida:** Dashboard de estado de transportes
- **Impacto:** Visibilidad en tiempo real de la cadena logística

### 🔄 **Bot Traspaso Automático (LT01)**
- **Función:** Ejecuta traspasos masivos entre ubicaciones
- **Tecnología:** Automatización de transacción LT01
- **Características:**
  - ✅ Carga desde Excel
  - ✅ Validación de stock disponible
  - ✅ Generación de documentos de traspaso
- **Impacto:** Elimina errores de digitación en traspasos

### 🧠 **Bot Visión Operacional (IA)**
- **Función:** Digitaliza información manuscrita de pizarras de andén
- **Tecnología:** Google Gemini Vision API
- **Características:**
  - 📸 Procesamiento de imágenes
  - 📝 OCR inteligente de texto manuscrito
  - 📊 Integración con Power BI
- **Impacto:** Digitaliza operaciones no sistematizadas

### 📧 **Bot Consolidación Zonales**
- **Función:** Procesa correos de reportes zonales y genera consolidados
- **Tecnología:** Outlook API + Pandas
- **Salida:** Excel consolidado con análisis multi-zonal
- **Impacto:** Automatiza reportería gerencial

### 🔢 **Bot Conversiones UMV**
- **Función:** Extrae factores de conversión de unidades desde MM03
- **Tecnología:** SAP GUI Scripting
- **Salida:** Tabla maestra de conversiones
- **Impacto:** Mantiene actualizada la base de datos de conversiones

### 📄 **Bot Lectura de Facturas (IA)**
- **Función:** Extrae datos de facturas escaneadas
- **Tecnología:** Google Gemini Vision API
- **Características:**
  - 🔍 Detección de campos clave (RUT, fecha, total)
  - 📋 Extracción de líneas de detalle
  - 📝 Transcripción de notas manuscritas
- **Impacto:** Elimina digitación manual de facturas

---

## 🛠️ Stack Tecnológico

### Frontend (Interfaz Web)
- **Framework:** React 18 + Vite
- **Estilos:** Tailwind CSS
- **Base de Datos:** Firebase Realtime Database
- **Hosting:** Vercel
- **Comunicación:** REST API + WebSocket (Firebase)

### Backend (Orquestador)
- **Lenguaje:** Python 3.10+
- **Automatización SAP:** win32com (SAP GUI Scripting)
- **IA/ML:** Google Gemini API (Vision + Text)
- **Procesamiento:** Pandas, openpyxl
- **Email:** win32com.client (Outlook)
- **Base de Datos:** Firebase Admin SDK

### Bots Workers
- **Interfaz Local:** CustomTkinter (opcional)
- **Logging:** Python logging module
- **Gestión de Estado:** JSON local + Firebase

---

## 📁 Estructura del Proyecto

```
Nexus_Jarvis/
├── 📂 Interfaz_Vercel/          # Frontend React (Vercel)
│   ├── src/
│   │   ├── components/          # Componentes React
│   │   ├── firebase.js          # Configuración Firebase
│   │   └── App.jsx              # Aplicación principal
│   ├── public/
│   ├── package.json
│   └── vercel.json
│
├── 📂 Bots/                     # Suite de Bots Workers
│   ├── Bot_Conciliacion_Email.py
│   ├── Bot_Auditor.py
│   ├── Bot_Pallet.py
│   ├── Bot_Transporte.py
│   ├── Bot_Traspaso_LT01.py
│   ├── Bot_Vision.py
│   ├── Bot_Consolidacion_Zonales.py
│   ├── Bot_Conversiones_UMV.py
│   └── Bot_Lectura_Facturas.py
│
├── 📄 worker_sap.py             # Orquestador Central
├── 📄 Logistic-Automation-Suite.py  # Interfaz Local (Legacy)
├── 📄 launcher.bat              # Script de inicio
├── 📄 installer.bat             # Instalador automático
├── 📄 requirements.txt          # Dependencias Python
├── 📄 fire.json                 # Credenciales Firebase
├── 📄 README_INSTALACION.md     # Guía de instalación
└── 📄 GUIA_RAPIDA.md            # Guía rápida de uso
```

---

## 🚀 Instalación y Despliegue

### Requisitos Previos
- ✅ Python 3.10 o superior
- ✅ Node.js 18+ (para frontend)
- ✅ SAP GUI con Scripting habilitado
- ✅ Cuenta de Firebase (Realtime Database)
- ✅ API Key de Google Gemini

### Instalación Rápida (Windows)

```bash
# 1. Clonar el repositorio
git clone https://github.com/tu-usuario/nexus-jarvis.git
cd nexus-jarvis

# 2. Ejecutar instalador automático
installer.bat

# 3. Configurar credenciales Firebase
# Editar fire.json con tus credenciales

# 4. Lanzar el sistema
launcher.bat
```

### Instalación Manual

```bash
# Backend (Python)
pip install -r requirements.txt

# Frontend (React)
cd Interfaz_Vercel
npm install
npm run dev

# Worker (Orquestador)
python worker_sap.py
```

### Despliegue en Vercel

```bash
cd Interfaz_Vercel

# Configurar variables de entorno en Vercel Dashboard:
# VITE_FIREBASE_API_KEY
# VITE_FIREBASE_AUTH_DOMAIN
# VITE_FIREBASE_PROJECT_ID
# VITE_FIREBASE_STORAGE_BUCKET
# VITE_FIREBASE_MESSAGING_SENDER_ID
# VITE_FIREBASE_APP_ID

# Desplegar
vercel --prod
```

---

## 📖 Documentación Adicional

- 📘 [Guía de Instalación Completa](README_INSTALACION.md)
- 📗 [Guía Rápida de Uso](GUIA_RAPIDA.md)
- 📙 [Instrucciones para el Equipo](INSTRUCCIONES_EQUIPO.md)

---

## 🔐 Seguridad

- 🔒 **Credenciales:** Almacenadas en variables de entorno (no versionadas)
- 🔑 **Firebase:** Autenticación y reglas de seguridad configuradas
- 🛡️ **SAP:** Acceso mediante credenciales de usuario (no almacenadas)
- 📝 **Logs:** Sin información sensible en registros

---

## 📊 Métricas de Impacto

| Métrica | Antes | Después | Mejora |
|---------|-------|---------|--------|
| Tiempo de carga MIGO | 45 min | 5 min | **90% ↓** |
| Errores de digitación | 15% | <1% | **93% ↓** |
| Auditorías de stock | 4h | 30 min | **87% ↓** |
| Procesamiento facturas | 2h | 15 min | **87% ↓** |
| Reportes zonales | 1.5h | 10 min | **89% ↓** |

---

## 🗺️ Roadmap

- [ ] Integración con Power BI API
- [ ] Dashboard de métricas en tiempo real
- [ ] Notificaciones push (Telegram/WhatsApp)
- [ ] Modo offline con sincronización diferida
- [ ] Soporte multi-idioma (ES/EN)
- [ ] API REST pública para integraciones
- [ ] Módulo de Machine Learning para predicción de stock

---

## 👨‍💻 Autor

**Ariel Mella**  
Ingeniero de Soluciones Operacionales | Logística & Datos (RPA lead/Python/SAP/AI) | Facilitador Técnico de Mejora Continua
📧 ariel.mella@cial.cl | ariel.mellag@gmail.com

---

## 📄 Licencia

Este proyecto es **propietario** y de uso interno exclusivo de CIAL Alimentos.  
Prohibida su distribución o uso comercial sin autorización.

---

## 🙏 Agradecimientos

- Google Gemini API por las capacidades de IA
- Firebase por la infraestructura en tiempo real
- Vercel por el hosting gratuito
- Comunidad de Python por las librerías open-source

---

<div align="center">
  <strong>Hecho con ❤️ y ☕ en Chile</strong>
</div>
