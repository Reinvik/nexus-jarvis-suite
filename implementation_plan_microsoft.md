# Implementación de Integración Híbrida Microsoft (OneDrive/SharePoint)

## Estrategia: "OneDrive Sync"
Aprovecharemos la sincronización local de OneDrive para conectar el mundo Python con el ecosistema Microsoft sin requerir complejas configuraciones de Azure AD ni APIs costosas.

### 1. Arquitectura de Carpetas (OneDrive Usuario)
El sistema dependerá de dos carpetas clave que deben existir en el OneDrive del usuario (o un SharePoint sincronizado):

1.  **`.../Nexus_System/Dist` (Canal de Distribución)**
    - Contiene: `Nexus Jarvis.exe` (última versión), `version.txt` (ej: `1.5.0`), `changelog.txt`.
    - **Uso**: El Launcher leerá esta carpeta. Si `version.txt` local < `version.txt` nube, descargará y actualizará.

2.  **`.../Nexus_System/Logs` (Backend de Datos)**
    - Contiene: `bitacora_operaciones.csv` (o Excel).
    - **Uso**: Todos los bots escribirán aquí sus resultados ("MIGO creada", "Error SAP", etc.).
    - **Integración Power BI**: Power BI Desktop se conecta directamente a este archivo CSV/Excel en SharePoint/OneDrive para generar dashboards en tiempo real.

### 2. Componentes a Desarrollar

#### A. `nexus_updater.py` (Módulo de Actualización)
Este módulo se integrará en el `nexus_launcher.py`.
- **Lógica**:
    1.  Al iniciar, leer configuración local (dónde buscar actualizaciones).
    2.  Comparar versión local vs `version.txt` en la ruta de actualización.
    3.  Si hay nueva versión:
        - Notificar al usuario.
        - Renombrar ejecutable actual a `.bak`.
        - Copiar nuevo ejecutable desde OneDrive.
        - Reiniciar aplicación.

#### B. `nexus_logger.py` (Central de Logs)
Un módulo estandarizado para que los bots reporten actividad.
- **Formato CSV**:
    - `Timestamp`: Fecha y hora ISO 8601.
    - `Bot`: Nombre del bot (ej: "MIGO", "Pallet").
    - `Usuario`: Usuario de Windows.
    - `Accion`: "Inicio", "Fin", "Error", "Registro Creado".
    - `Detalle`: Texto libre o JSON con detalles (ej: "Doc Material 5000000123").
    - `Estado`: "OK", "ERROR", "WARN".

### 3. Integración Power BI y API
- **Power BI**: Se conecta al `bitacora_operaciones.csv` en OneDrive. Al estar en la nube, se puede programar refresh automático si se publica en Power BI Service.
- **Futura API**: Si se requiere una API real, se puede usar Power Automate para leer este CSV y exponer datos, o migrar el CSV a una SharePoint List.
en power bi extraera información de los excel generados de los Bots de transporte vt11.

## Plan de Ejecución
1.  Definir rutas en `settings.json` (para flexibilidad).
2.  Crear `nexus_logger.py` y probar escritura en OneDrive.
3.  Crear `nexus_updater.py` y simular una actualización.
