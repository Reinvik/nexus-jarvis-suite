# Plan de Distribución Nexus Jarvis

Este documento detalla la estrategia para distribuir Nexus Jarvis a dos tipos de usuarios diferentes: Tú (Admin) y los Usuarios Finales (Operadores).

## 1. El Enfoque (Un solo EXE, dos experiencias)
En lugar de crear programas diferentes, usaremos el mismo `Nexus Jarvis.exe` generado, pero entregaremos carpetas con distinto contenido.

| Característica | Versión Admin (Ariel) | Versión Usuario (Operadores) |
| :--- | :--- | :--- |
| **Ejecutable** | `Nexus Jarvis.exe` | `Nexus Jarvis.exe` |
| **Interfaz Gráfica** | ✅ Acceso total | ✅ Acceso total (MIGO, LT01, Auditor) |
| **Procesos de Fondo** | ✅ Manager, Emails, Workers | ❌ No incluidos |
| **Método de Inicio** | `start_manager.bat` (Inicia todo) | Doble clic en `Nexus Jarvis.exe` |

## 2. Estructura de Carpetas Propuesta

### 📦 Carpeta: `Nexus_Jarvis_Usuario_Final`
*Lo que le entregaremos a tus colegas.*
*   📂 **_internal/**: (Archivos del sistema, no tocar).
*   📄 **Nexus Jarvis.exe**: La aplicación.
*   📂 **Plantillas/**: Carpeta con los Excels vacíos que necesitan para trabajar.
    *   `Plantilla_MIGO.xlsx`
    *   `Plantilla_LT01.xlsx`
    *   `Plantilla_Auditor.xlsx`
*   📄 **LEEME.txt**: Instrucciones simples ("Pega tus datos en la plantilla y ejecuta").

### 🔧 Carpeta: `Nexus_Jarvis_Admin` (Tuya)
*   Todo lo anterior +
*   📄 **start_manager.bat**: Para activar tus bots de correo y workers.
*   📄 **email_commander.py**, etc.: (Ya integrados en el EXE, pero accesibles si necesitas scripts sueltos).

## 3. Discusión: Archivos de Entrada (Excels)
Mencionaste "no se sobre que hacer los archivos". Para que los usuarios usen los bots, necesitan llenar ciertos Excels.

*   **MIGO**: Requiere un Excel con columnas específicas (Material, Centro, etc.). ¿Tienes una plantilla estándar?
*   **LT01**: Requiere Excel con (Material, Cantidad, Tipo).
*   **Auditor**: ¿Requiere Excel o solo input manual de Almacén? (Parece ser manual por el código).

## 4. Próximos Pasos
1.  **Recopilar Plantillas**: Buscar o crear los Excels vacíos ("Templates") para incluirlos en la entrega.
2.  **Limpiar Distribución**: Asegurar que la carpeta de Usuario no tenga scripts basura.
3.  **Configuración**:
    *   ¿Quieres que los usuarios reporten errores a tu Firebase? (Dejar `fire.json`).
    *   ¿O prefieres que funcionen totalmente offline? (Quitar `fire.json` si es posible, aunque el código podría requerirlo).

---
**Pregunta:** ¿Te parece bien este enfoque de "mismo EXE, diferente entrega"? ¿Y tienes a mano los Excels de ejemplo para ponerlos en una carpeta de "Plantillas"?
