import pandas as pd
import os
import shutil
import sys
from datetime import datetime
import time

# Configurar salida estándar a UTF-8 para soportar emojis en Windows
try:
    if sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

class BotAnalisisZonales:
    def __init__(self):
        # Rutas
        self.onedrive_base = os.path.join(os.path.expanduser("~"), "OneDrive - CIAL Alimentos")
        self.target_folder = os.path.join(
            self.onedrive_base,
            "Archivos de Operación  Outbound CD - 16.-Inventario Critico"
        )
        self.master_file = os.path.join(self.target_folder, "Consolidado_Zonales_Master.xlsx")
        self.maestro_pasillos_file = os.path.join(os.getcwd(), "Maestro_Pasillos.xlsx")
        
    def load_maestro_pasillos(self):
        """Carga el maestro de pasillos en un diccionario {SKU: Pasillo}"""
        print("📚 Cargando Maestro de Pasillos...")
        if not os.path.exists(self.maestro_pasillos_file):
            print("⚠️ ADVERTENCIA: No se encontró 'Maestro_Pasillos.xlsx'. La detección de productos cambiados será limitada.")
            return {}
            
        try:
            df = pd.read_excel(self.maestro_pasillos_file, dtype=str)
            # Normalizar columnas
            df.columns = df.columns.astype(str).str.lower().str.strip()
            
            # Buscar columnas clave
            col_sku = next((c for c in df.columns if 'sku' in c or 'material' in c), None)
            col_pasillo = next((c for c in df.columns if 'pasillo' in c), None)
            
            if not col_sku or not col_pasillo:
                print("❌ Error: Maestro de Pasillos debe tener columnas SKU y Pasillo.")
                return {}
                
            # Crear diccionario {SKU_Limpio: Pasillo_Limpio}
            lookup = {}
            for _, row in df.iterrows():
                sku = str(row[col_sku]).strip().lstrip('0')
                pasillo = str(row[col_pasillo]).strip().upper()
                if sku and pasillo:
                    lookup[sku] = pasillo
            
            print(f"✅ Maestro cargado: {len(lookup)} SKUs mapeados.")
            return lookup
            
        except Exception as e:
            print(f"❌ Error leyendo maestro de pasillos: {e}")
            return {}

    def get_aisle(self, sku, lookup):
        """Obtiene pasillo para un SKU"""
        if not sku: return "SIN_SKU"
        s = str(sku).strip().lstrip('0')
        return lookup.get(s, "DESCONOCIDO")

    def analyze_transport(self, transport_id, df_f_sub, df_s_sub, pasillo_lookup):
        """Analiza un transporte específico en busca de cruces"""
        
        # Listas para resultados
        swapped_items = []
        remaining_faltantes = []
        remaining_sobrantes = []
        
        # Hacemos copias para no afectar el original y poder "consumir" filas
        f_rows = df_f_sub.to_dict('records')
        s_rows = df_s_sub.to_dict('records')
        
        # Agregar pasillo a cada fila para facilitar comparación
        for row in f_rows:
            row['_pasillo'] = self.get_aisle(row.get('SKU'), pasillo_lookup)
            row['_matched'] = False
            
        for row in s_rows:
            row['_pasillo'] = self.get_aisle(row.get('SKU'), pasillo_lookup)
            row['_matched'] = False
            
        # --- ALGORITMO PRODUCTO CAMBIADO ---
        # Condición: Misma Cantidad Y Mismo Pasillo
        
        for f in f_rows:
            if f['_matched']: continue
            
            f_qty = f.get('Cantidad')
            f_pasillo = f['_pasillo']
            
            # Buscar match en sobrantes
            best_match_idx = -1
            
            for i, s in enumerate(s_rows):
                if s['_matched']: continue
                
                s_qty = s.get('Cantidad')
                s_pasillo = s['_pasillo']
                
                # CRITERIO ESTRICTO
                if f_qty == s_qty and f_pasillo == s_pasillo and f_pasillo != "DESCONOCIDO":
                    best_match_idx = i
                    break
            
            if best_match_idx != -1:
                # MATCH ENCONTRADO
                s = s_rows[best_match_idx]
                f['_matched'] = True
                s['_matched'] = True
                
                swapped_items.append({
                    'Transporte_ID': transport_id,
                    'Zonal': f.get('Zonal'),
                    'Fecha': f.get('Fecha_Email'),
                    'SKU_Faltante': f.get('SKU'),
                    'Desc_Faltante': f.get('Descripcion'),
                    'SKU_Sobrante': s.get('SKU'),
                    'Desc_Sobrante': s.get('Descripcion'),
                    'Cantidad': f_qty,
                    'UM': f.get('UM', 'UN'),
                    'Pasillo': f_pasillo,
                    'Slotting': 'Mismo Pasillo Picking'
                })
        
        # Recopilar remanentes (Lo que no fue match)
        for f in f_rows:
            if not f['_matched']: remaining_faltantes.append(f)
            
        for s in s_rows:
            if not s['_matched']: remaining_sobrantes.append(s)
            
        # --- DETERMINAR ESTADO GLOBAL DEL TRANSPORTE ---
        status = "Indeterminado"
        
        has_swapped = len(swapped_items) > 0
        has_falt = len(remaining_faltantes) > 0
        has_sobr = len(remaining_sobrantes) > 0
        
        if not has_swapped and not has_falt and not has_sobr:
            status = "Sin Diferencias" # (O archivo vacío)
        elif has_swapped and not has_falt and not has_sobr:
            status = "Producto Cambiado"
        elif not has_swapped and has_falt and not has_sobr:
            status = "Faltante"
        elif not has_swapped and not has_falt and has_sobr:
            status = "Sobrante"
        elif not has_swapped and has_falt and has_sobr:
            status = "Faltante/Sobrante"
        else:
            # Combinaciones complejas
            parts = []
            if has_swapped: parts.append("Producto Cambiado")
            if has_falt: parts.append("Faltante")
            if has_sobr: parts.append("Sobrante")
            if len(parts) > 1:
                status = " / ".join(parts) # Ej: Producto Cambiado / Faltante
            else:
                status = parts[0] if parts else "Revisar"

        return status, swapped_items, remaining_faltantes, remaining_sobrantes

    def run(self):
        print("--- INICIANDO ANÁLISIS DE TRANSPORTES ZONALES ---")
        
        if not os.path.exists(self.master_file):
            print("❌ Error: No existe el archivo maestro Consolidado_Zonales_Master.xlsx")
            return

        # 1. Crear Snapshot
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        snapshot_name = f"Analisis_Zonales_{timestamp}.xlsx"
        # Guardar en escritorio para facil acceso
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        snapshot_path = os.path.join(desktop, snapshot_name)
        
        print(f"📸 Creando snapshot de trabajo: {snapshot_name}")
        try:
            shutil.copy2(self.master_file, snapshot_path)
            print("✅ Copia creada exitosamente.")
        except Exception as e:
            print(f"❌ Error copiando archivo: {e}")
            return

        # 2. Cargar Datos
        try:
            print("📖 Leyendo datos del snapshot...")
            xl = pd.ExcelFile(snapshot_path)
            
            # Cargar pestañas seguras
            df_faltantes = pd.read_excel(xl, sheet_name='Faltantes', dtype=str) if 'Faltantes' in xl.sheet_names else pd.DataFrame()
            df_sobrantes = pd.read_excel(xl, sheet_name='Sobrantes', dtype=str) if 'Sobrantes' in xl.sheet_names else pd.DataFrame()
            df_transportes_ref = pd.read_excel(xl, sheet_name='Transportes', dtype=str) if 'Transportes' in xl.sheet_names else pd.DataFrame()
            
            # Convertir Cantidad a números
            if 'Cantidad' in df_faltantes.columns:
                df_faltantes['Cantidad'] = pd.to_numeric(df_faltantes['Cantidad'], errors='coerce').fillna(0)
            if 'Cantidad' in df_sobrantes.columns:
                df_sobrantes['Cantidad'] = pd.to_numeric(df_sobrantes['Cantidad'], errors='coerce').fillna(0)
                
        except Exception as e:
             print(f"❌ Error leyendo excel: {e}")
             return

        pasillos = self.load_maestro_pasillos()
        
        # 3. Definir Universo de Transportes
        # Fuente primaria: Pestaña 'Transportes'
        # Fuente secundaria: Archivos presentes en faltantes/sobrantes
        
        all_transports = set()
        
        # 3.1 Desde referencia Transporte
        if not df_transportes_ref.empty and 'Origen_Archivo' in df_transportes_ref.columns:
             all_transports.update(df_transportes_ref['Origen_Archivo'].unique())
             
        # 3.2 Desde Discrepancias (por robustez)
        if 'Origen_Archivo' in df_faltantes.columns:
            all_transports.update(df_faltantes['Origen_Archivo'].unique())
        if 'Origen_Archivo' in df_sobrantes.columns:
            all_transports.update(df_sobrantes['Origen_Archivo'].unique())
            
        all_transports = [t for t in all_transports if str(t).lower() not in ['nan', 'none', '']]
        print(f"🔍 Analizando {len(all_transports)} transportes totales...")
        
        summary_rows = []
        all_discrepancies_unified = []
        all_swapped_details = []
        
        for transp in all_transports:
            # Filtrar Dataframes Discrepancias
            f_sub = pd.DataFrame()
            s_sub = pd.DataFrame()
            
            # Intentar obtener metadatos (Fecha, Zonal)
            meta_zonal = "Desconocido"
            meta_fecha = "Desconocido"
            
            # Prioridad Metadatos: Transporte > Faltante > Sobrante
            if not df_transportes_ref.empty and 'Origen_Archivo' in df_transportes_ref.columns:
                tr_row = df_transportes_ref[df_transportes_ref['Origen_Archivo'] == transp]
                if not tr_row.empty:
                    meta_zonal = tr_row.iloc[0].get('Zonal', meta_zonal)
                    meta_fecha = tr_row.iloc[0].get('Fecha_Email', meta_fecha)

            if 'Origen_Archivo' in df_faltantes.columns:
                f_sub = df_faltantes[df_faltantes['Origen_Archivo'] == transp].copy()
                if not f_sub.empty and meta_zonal == "Desconocido":
                    meta_zonal = f_sub.iloc[0].get('Zonal', meta_zonal)
                    meta_fecha = f_sub.iloc[0].get('Fecha_Email', meta_fecha)

            if 'Origen_Archivo' in df_sobrantes.columns:
                s_sub = df_sobrantes[df_sobrantes['Origen_Archivo'] == transp].copy()
                if not s_sub.empty and meta_zonal == "Desconocido":
                    meta_zonal = s_sub.iloc[0].get('Zonal', meta_zonal)
                    meta_fecha = s_sub.iloc[0].get('Fecha_Email', meta_fecha)

            # Ejecutar Análisis de Cruce
            status, swapped, rem_f, rem_s = self.analyze_transport(transp, f_sub, s_sub, pasillos)
            
            # Guardar Swaps
            all_swapped_details.extend(swapped)
            
            # Guardar Discrepancias Unificadas (Faltantes + Sobrantes remanentes)
            for item in rem_f:
                item['Estado'] = 'Faltante'
                item['Tipo_Diferencia'] = 'Faltante'
                if '_matched' in item: del item['_matched']
                if '_pasillo' in item: del item['_pasillo']
                all_discrepancies_unified.append(item)
                
            for item in rem_s:
                item['Estado'] = 'Sobrante'
                item['Tipo_Diferencia'] = 'Sobrante'
                if '_matched' in item: del item['_matched']
                if '_pasillo' in item: del item['_pasillo']
                all_discrepancies_unified.append(item)

            # --- NUEVO: Agregar fila para transportes "Sin Diferencias" ---
            if status == "Sin Diferencias":
                 dummy_row = {
                    'Fecha_Email': meta_fecha,
                    'Zonal': meta_zonal,
                    'Origen_Archivo': transp,
                    'Estado': 'Sin Diferencias',
                    'Tipo_Diferencia': 'Sin Diferencias',
                    'SKU': '-',
                    'Descripcion': 'Transporte Cuadrado (Sin Diferencias)',
                    'Cantidad': 0,
                    'Pasillo': '-'
                 }
                 all_discrepancies_unified.append(dummy_row)
            # -----------------------------------------------------------
            
            # Resumen Global
            summary_rows.append({
                'Fecha': meta_fecha,
                'Zonal': meta_zonal,
                'Archivo_Origen': transp,
                'Estado_Transporte': status,
                'Cant_Faltantes_Reales': len(rem_f),
                'Cant_Sobrantes_Reales': len(rem_s),
                'Cant_Productos_Cambiados': len(swapped),
                'Analisis': f"{len(rem_f)} Falt / {len(rem_s)} Sobr / {len(swapped)} Swaps"
            })
            
        # 4. Escribir Resultados
        print("💾 Escribiendo reporte final...")
        
        df_summary = pd.DataFrame(summary_rows)
        # Ordenar por fecha
        try:
            df_summary['DT'] = pd.to_datetime(df_summary['Fecha'], dayfirst=True, errors='coerce')
            df_summary = df_summary.sort_values('DT', ascending=False)
            del df_summary['DT']
        except: pass
        
        df_swapped = pd.DataFrame(all_swapped_details)
        df_discrepancies = pd.DataFrame(all_discrepancies_unified)
        
        try:
            with pd.ExcelWriter(snapshot_path, mode='a', engine='openpyxl') as writer:
                df_summary.to_excel(writer, sheet_name='Resumen_Global', index=False)
                
                if not df_discrepancies.empty:
                    # Organizar columnas para que Zonal, Fecha, Estado queden al principio
                    cols = list(df_discrepancies.columns)
                    priority_cols = ['Fecha_Email', 'Zonal', 'Estado', 'SKU', 'Descripcion', 'Cantidad', 'Pasillo']
                    final_cols = [c for c in priority_cols if c in cols] + [c for c in cols if c not in priority_cols]
                    df_discrepancies = df_discrepancies[final_cols]
                    
                    df_discrepancies.to_excel(writer, sheet_name='Analisis_Discrepancias', index=False)
                
                if not df_swapped.empty:
                    df_swapped.to_excel(writer, sheet_name='Productos_Cambiados', index=False)
            
            print(f"🚀 ¡LISTO! Reporte generado en su escritorio.")
            print(f"📂 Archivo: {snapshot_name}")
            
            # Abrir archivo automáticamente
            os.startfile(snapshot_path)
            
        except Exception as e:
            print(f"❌ Error guardando resultados: {e}")

if __name__ == "__main__":
    bot = BotAnalisisZonales()
    bot.run()
