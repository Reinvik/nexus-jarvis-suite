import os
import sys
import google.generativeai as genai
import pandas as pd
from PIL import Image
import time
from datetime import datetime

# Configurar salida para soportar emojis en Windows
try:
    sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

class FacturaBot:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.input_folder = os.path.join(self.base_dir, "Muestra Facturas")
        self.output_file = os.path.join(self.base_dir, "Consolidado_Facturas.xlsx")
        self.output_file = os.path.join(self.base_dir, "Consolidado_Facturas.xlsx")
        # Intentar leer variable de entorno, si no existe será None
        self.api_key = os.getenv("GEMINI_API_KEY")

    def setup_api(self):
        """Configura la API key de Gemini"""
        if not self.api_key:
            # Si no está en variables de entorno, pedirla
            print("⚠️ No se encontró la variable de entorno GEMINI_API_KEY.")
            print("Por favor, ingresa tu API Key de Google Gemini (o presiona Enter para salir):")
            self.api_key = input().strip()
            
        if not self.api_key:
            print("❌ No se proporcionó API Key. Saliendo...")
            return False
            
        genai.configure(api_key=self.api_key)
        return True

    def get_images(self):
        """Obtiene la lista de imágenes en la carpeta"""
        if not os.path.exists(self.input_folder):
            print(f"❌ La carpeta '{self.input_folder}' no existe.")
            return []
            
        valid_extensions = ('.jpg', '.jpeg', '.png', '.webp')
        images = [f for f in os.listdir(self.input_folder) if f.lower().endswith(valid_extensions)]
        print(f"📁 Encontradas {len(images)} imágenes en '{self.input_folder}'")
        return images

    def get_best_model(self):
        """Busca el mejor modelo de visión disponible"""
        print("🔍 Buscando modelos disponibles...")
        try:
            available_models = []
            print("   Modelos encontrados en tu cuenta:")
            for m in genai.list_models():
                if 'generateContent' in m.supported_generation_methods:
                    print(f"   - {m.name}")
                    available_models.append(m.name)
            
            # Prioridad de modelos (nombres completos con 'models/')
            priorities = [
                'models/gemini-2.0-flash',
                'models/gemini-2.0-flash-001',
                'models/gemini-2.5-flash',
                'models/gemini-1.5-flash',
                'models/gemini-flash-latest',
                'models/gemini-2.0-pro-exp',
                'models/gemini-1.5-pro',
                'models/gemini-pro-vision'
            ]
            
            for model_name in priorities:
                if model_name in available_models:
                    print(f"✅ Modelo seleccionado: {model_name}")
                    return model_name
            
            # Búsqueda parcial inteligente
            for m in available_models:
                if 'flash' in m and '2.0' in m:
                    return m
                if 'flash' in m and '1.5' in m:
                    return m
                if 'flash' in m: # Cualquier flash
                    return m
            
            # Si llegamos aquí, probamos con el primero que parezca útil
            if available_models:
                print(f"⚠️ Usando primer modelo disponible: {available_models[0]}")
                return available_models[0]
                
            print("⚠️ No se encontró un modelo ideal. Usando default 'gemini-2.0-flash'")
            return 'gemini-2.0-flash'
            
        except Exception as e:
            print(f"⚠️ Error listando modelos: {e}. Usando default.")
            return 'gemini-1.5-flash'

    def analyze_image(self, image_filename, model_name):
        """Analiza una imagen usando Gemini Vision"""
        image_path = os.path.join(self.input_folder, image_filename)
        print(f"🔍 Analizando: {image_filename}...")
        
        try:
            from PIL import Image, ImageOps
            img = Image.open(image_path)
            
            # Corregir orientación basada en EXIF (para fotos de celular)
            try:
                img = ImageOps.exif_transpose(img)
            except:
                pass
            
            # Inicializar modelo
            model = genai.GenerativeModel(model_name)
            
            prompt = """
            Analiza esta imagen de factura y extrae la información en formato JSON estructurado.
            
            Reglas de extracción:
            1. FECHA: Formato estricto DD/MM/YYYY (ejemplo: 21/11/2025). Si el año aparece como '25', conviértelo a '2025'.
            2. PROVEEDOR: Nombre completo de la empresa.
            3. ITEMS (Detalle): Extrae cada línea de producto con:
               - Material: Código o descripción del producto.
               - Cantidad: Número de unidades.
               - Kilos: Peso en kg si aparece explícitamente (ej: "15.4 kg"), sino null o vacío.
               - Total Linea: Valor total de esa línea.
            4. NOTAS MANUSCRITAS: Busca texto escrito a mano (lápiz/bolígrafo) como "falta -1 cj", "recepción conforme", etc. Transcríbelo tal cual.
            
            Responde SOLO con este JSON válido:
            {
                "numero_factura": "string",
                "fecha": "DD/MM/YYYY",
                "proveedor": "string",
                "total_factura": "string",
                "notas_manuscritas": "string",
                "items": [
                    {
                        "material": "string",
                        "cantidad": "string",
                        "kilos": "string",
                        "total_linea": "string"
                    }
                ]
            }
            """
            
            response = model.generate_content([prompt, img])
            
            # Limpiar respuesta para obtener solo JSON
            text_response = response.text.strip()
            if text_response.startswith("```json"):
                text_response = text_response[7:-3]
            elif text_response.startswith("```"):
                text_response = text_response[3:-3]
                
            import json
            data = json.loads(text_response)
            data['archivo'] = image_filename # Agregar nombre de archivo
            return data
            
        except Exception as e:
            print(f"❌ Error analizando {image_filename}: {e}")
            return None

    def save_to_excel(self, data_list):
        """Guarda los datos en Excel, aplanando los items"""
        if not data_list:
            print("⚠️ No hay datos para guardar.")
            return

        # Aplanar datos (un registro por item)
        flat_data = []
        for doc in data_list:
            # Normalizar proveedor
            prov = doc.get('proveedor', '')
            if prov and 'CONSORCIO' in prov.upper() and 'ALIMENTOS' in prov.upper():
                prov = "CONSORCIO INDUSTRIAL DE ALIMENTOS S.A."
            
            # Info base de la factura
            base_info = {
                'archivo': doc.get('archivo'),
                'numero_factura': doc.get('numero_factura'),
                'fecha': doc.get('fecha'),
                'proveedor': prov,
                'total_factura': doc.get('total_factura'),
                'notas_manuscritas': doc.get('notas_manuscritas')
            }
            
            items = doc.get('items', [])
            if not items:
                # Si no hay items, agregar fila solo con cabecera
                flat_data.append(base_info)
            else:
                for item in items:
                    row = base_info.copy()
                    row.update({
                        'material': item.get('material'),
                        'cantidad': item.get('cantidad'),
                        'kilos': item.get('kilos'),
                        'total_linea': item.get('total_linea')
                    })
                    flat_data.append(row)

        df = pd.DataFrame(flat_data)
        
        # Reordenar columnas
        cols = [
            'archivo', 'numero_factura', 'fecha', 'proveedor', 
            'material', 'cantidad', 'kilos', 'total_linea', 
            'total_factura', 'notas_manuscritas'
        ]
        
        # Asegurar que existan todas las columnas
        for col in cols:
            if col not in df.columns:
                df[col] = ""
        
        df = df[cols]
        
        try:
            df.to_excel(self.output_file, index=False)
            print(f"\n✅ Consolidado guardado exitosamente en:\n{self.output_file}")
            
            # Ajustar ancho de columnas (opcional, visual)
            try:
                from openpyxl import load_workbook
                wb = load_workbook(self.output_file)
                ws = wb.active
                for column in ws.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column[0].column_letter].width = adjusted_width
                wb.save(self.output_file)
            except:
                pass
                
        except PermissionError:
            print(f"\n❌ Error: El archivo '{os.path.basename(self.output_file)}' está abierto.")
            # Guardar con otro nombre
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_filename = f"Consolidado_Facturas_{timestamp}.xlsx"
            new_path = os.path.join(self.base_dir, new_filename)
            
            try:
                df.to_excel(new_path, index=False)
                print(f"✅ Se guardó una copia en: {new_filename}")
            except Exception as e:
                print(f"❌ No se pudo guardar la copia: {e}")
                
        except Exception as e:
            print(f"❌ Error guardando Excel: {e}")

    def run(self):
        print("="*50)
        print("🤖 BOT LECTOR DE FACTURAS (GEMINI AI)")
        print("="*50)
        
        if not self.setup_api():
            return

        images = self.get_images()
        if not images:
            return

        # Seleccionar modelo una sola vez
        model_name = self.get_best_model()

        consolidated_data = []
        
        for img_file in images:
            data = self.analyze_image(img_file, model_name)
            if data:
                consolidated_data.append(data)
                print(f"   ✅ Datos extraídos: {data.get('numero_factura', 'S/N')} | Notas: {data.get('notas_manuscritas', 'Ninguna')}")
            time.sleep(1) # Pequeña pausa para evitar rate limits
            
        self.save_to_excel(consolidated_data)
        print("\nProceso finalizado.")

if __name__ == "__main__":
    bot = FacturaBot()
    bot.run()
