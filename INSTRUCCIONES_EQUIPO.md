# 🚀 Instrucciones para el Equipo - SAP Automation Suite

## 📋 ¿Qué es esto?

Un sistema de automatización SAP con:
- ✅ **Worker local** en tu PC (procesa las órdenes en SAP)
- ✅ **Interfaz web** compartida (envías órdenes desde el navegador)
- ✅ **Actualización automática** vía OneDrive

---

## 🔧 Instalación (Solo Primera Vez)

### 1️⃣ Instalar Python

Si no tienes Python instalado:

1. Ve a: https://www.python.org/downloads/
2. Descarga **Python 3.11** o superior
3. Durante la instalación:
   - ✅ **IMPORTANTE:** Marca **"Add Python to PATH"**
   - ✅ Instala para todos los usuarios
4. Reinicia tu PC

### 2️⃣ Instalar Dependencias

1. Abre la carpeta de OneDrive: `Antigravity`
2. Haz **doble clic** en `installer.bat`
3. Espera 2-5 minutos
4. Cuando veas "✅ INSTALACIÓN COMPLETADA", cierra la ventana

---

## ▶️ Uso Diario

### Iniciar el Sistema

1. Abre la carpeta de OneDrive: `Antigravity`
2. Haz **doble clic** en `launcher.bat`
3. Verás:
   - ✅ Una **ventana negra** (Worker SAP) - **NO LA CIERRES**
   - ✅ Tu **navegador** se abrirá automáticamente con la interfaz

### Usar la Interfaz Web

1. En el navegador verás: **Nexus Orchestrator**
2. Selecciona el bot que necesites
3. Carga tu archivo Excel (si es necesario)
4. Haz clic en **"EJECUTAR FLUJO"**
5. El worker procesará automáticamente en SAP

### Cerrar el Sistema

- Cierra la **ventana negra** del Worker
- Puedes cerrar el navegador cuando quieras

---

## 🌐 Acceso Directo a la Interfaz

Si tu worker ya está corriendo, puedes acceder directamente desde cualquier navegador:

**🔗 https://nexus-orchestrator.vercel.app/**

---

## 🔄 Actualizaciones

**¡No necesitas hacer nada!**

- Los archivos están en OneDrive
- Cuando se actualice el código, solo cierra y vuelve a abrir `launcher.bat`
- Verás los cambios automáticamente

---

## 📊 Arquitectura del Sistema

```
┌─────────────────────────────────────────────────────┐
│                  TU NAVEGADOR                       │
│         https://nexus-orchestrator.vercel.app/      │
│                                                     │
│  [Seleccionar Bot] [Cargar Archivo] [Ejecutar]     │
└──────────────────┬──────────────────────────────────┘
                   │
                   │ Firebase (Nube)
                   │
┌──────────────────▼──────────────────────────────────┐
│              WORKER SAP (Tu PC)                     │
│                                                     │
│  • Escucha órdenes desde Firebase                  │
│  • Procesa en SAP automáticamente                  │
│  • Reporta resultados a la interfaz                │
└─────────────────────────────────────────────────────┘
```

**Ventajas:**
- ✅ Múltiples personas pueden usar la misma interfaz
- ✅ Cada uno ejecuta su propio worker
- ✅ No hay conflictos entre usuarios
- ✅ Interfaz siempre actualizada (está en la nube)

---

## ❓ Problemas Comunes

### "Python no se reconoce como comando"
**Solución:** Instala Python y marca "Add to PATH"

### "No se encuentra worker_sap.py"
**Solución:** Asegúrate de estar en la carpeta correcta de OneDrive

### El Worker se cierra inmediatamente
**Solución:** Verifica que `fire.json` esté en la carpeta

### La interfaz web no carga
**Solución:** Verifica tu conexión a internet

### El worker no procesa órdenes
**Solución:** 
1. Verifica que la ventana negra siga abierta
2. Revisa que diga "🤖 WORKER SAP INICIADO"
3. Si hay error, ejecuta `installer.bat` de nuevo

---

## 📁 Archivos Importantes

| Archivo | Para qué sirve |
|---------|---------------|
| `launcher.bat` | Inicia worker + abre navegador |
| `installer.bat` | Instala dependencias (solo 1ra vez) |
| `worker_sap.py` | Worker que procesa en SAP |
| `fire.json` | Credenciales de Firebase |

---

## 🎯 Flujo de Trabajo Típico

1. **Llegar a la oficina** → Doble clic en `launcher.bat`
2. **Trabajar normalmente** → Usar la interfaz web cuando necesites
3. **Salir de la oficina** → Cerrar la ventana negra

---

## 📞 Soporte

Si tienes problemas, contacta a:
- **Ariel Mella** - Desarrollador del sistema

---

## ✅ Checklist de Verificación

Después de la instalación, verifica:

- [ ] Python instalado correctamente
- [ ] `installer.bat` ejecutado sin errores
- [ ] `launcher.bat` abre la ventana negra
- [ ] El navegador se abre automáticamente
- [ ] Puedes ver la interfaz Nexus Orchestrator
- [ ] El worker dice "📡 Escuchando órdenes desde la Web..."

**Si todos los checks están ✅, estás listo para usar el sistema!** 🎉
