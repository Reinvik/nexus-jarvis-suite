# 🚀 SAP Automation Suite - Guía de Instalación

## 📋 Requisitos Previos
- Windows 10 o superior
- Acceso a OneDrive de CIAL Alimentos
- Conexión a internet (solo para instalación inicial)

---

## 🔧 Instalación (Solo Primera Vez)

### Paso 1: Verificar Python

1. Abre una ventana de **PowerShell** o **CMD**
2. Escribe: `python --version`
3. Si ves algo como `Python 3.11.x` → **Salta al Paso 2**
4. Si dice "no se reconoce el comando" → **Instala Python**:
   - Ve a: https://www.python.org/downloads/
   - Descarga **Python 3.11** o superior
   - Durante la instalación:
     - ✅ **IMPORTANTE:** Marca la casilla **"Add Python to PATH"**
     - ✅ Selecciona "Install for all users"
   - Reinicia tu PC después de instalar

### Paso 2: Ejecutar el Instalador

1. Abre la carpeta de OneDrive: `Antigravity`
2. Haz **doble clic** en `installer.bat`
3. Espera a que instale todas las dependencias (puede tardar 2-5 minutos)
4. Cuando veas "✅ INSTALACIÓN COMPLETADA", cierra la ventana

---

## ▶️ Uso Diario

### Iniciar el Sistema

1. Abre la carpeta de OneDrive: `Antigravity`
2. Haz **doble clic** en `launcher.bat`
3. Espera unos segundos y verás:
   - ✅ Una ventana negra (Worker SAP) - **NO LA CIERRES**
   - ✅ La interfaz gráfica del sistema

### Usar el Sistema

- Selecciona el bot que necesites desde la interfaz
- Carga tu archivo Excel cuando sea necesario
- El sistema procesará automáticamente en SAP

### Cerrar el Sistema

- Simplemente cierra la ventana negra del Worker
- La interfaz se cerrará automáticamente

---

## 🔄 Actualizaciones

**¡No necesitas hacer nada!** 

Como los archivos están en OneDrive:
- Cuando yo actualice el código, tú verás los cambios automáticamente
- Solo necesitas cerrar y volver a abrir el `launcher.bat`

---

## ❓ Problemas Comunes

### "Python no se reconoce como comando"
- **Solución:** Instala Python siguiendo el Paso 1 y marca "Add to PATH"

### "No se encuentra worker_sap.py"
- **Solución:** Asegúrate de estar en la carpeta correcta de OneDrive

### "Error al importar módulos"
- **Solución:** Ejecuta nuevamente `installer.bat`

### El Worker se cierra solo
- **Solución:** Revisa que `fire.json` esté en la carpeta

### La interfaz no se abre
- **Solución:** Verifica que SAP esté instalado en tu PC

---

## 📞 Soporte

Si tienes problemas, contacta a:
- **Ariel Mella** - Desarrollador del sistema

---

## 📁 Estructura de Archivos

```
Antigravity/
├── launcher.bat          ← EJECUTA ESTO para iniciar
├── installer.bat         ← Ejecuta solo la primera vez
├── requirements.txt      ← Lista de dependencias
├── worker_sap.py         ← Worker en segundo plano
├── Logistic-Automation-Suite.py  ← Interfaz gráfica
├── fire.json             ← Credenciales Firebase
├── Bot_*.py              ← Bots de automatización
└── Interfaz/             ← Archivos de la interfaz web
```

---

## ✅ Checklist de Instalación

- [ ] Python instalado (con "Add to PATH")
- [ ] Ejecutado `installer.bat` exitosamente
- [ ] `launcher.bat` abre Worker e Interfaz
- [ ] Puedo ver la interfaz gráfica
- [ ] El sistema está listo para usar

---

**¡Listo! Ahora puedes usar el sistema de automatización SAP fácilmente.** 🎉
