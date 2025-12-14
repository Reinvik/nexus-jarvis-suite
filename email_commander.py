import time
import os
import sys
import firebase_admin
from firebase_admin import credentials, firestore
import win32com.client
import pythoncom
from datetime import datetime
import google.generativeai as genai

# --- CONFIGURACIÓN ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "fire.json")

# Intentar cargar Prompt de Alexa si existe, sino usar default
PROMPT_PATH = os.path.join(BASE_DIR, "..", "Alexa_Jarvis", "jarvis_prompt.md")
DEFAULT_PROMPT = "Eres Jarvis, un asistente de IA irónico y eficiente. Responde brevemente."

# API KEY: Prioridad Variable de Entorno -> Hardcoded (Solo Dev Local)
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY")

TRIGGER_PREFIXES = ["NEXUS:", "JARVIS:"]  # Ahora responde a ambos
AUTHORIZED_EMAILS = ["ariel.mella@cial.cl"]

# Configuración UTF-8
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

class NexusEmailCommander:
    def __init__(self):
        print("🤖 INICIANDO AGENTE DE COMANDOS POR EMAIL (NEXUS J.A.R.V.I.S.)...")
        
        # 1. Configurar AI
        self.configure_ai()

        # 2. Conexión Firebase
        try:
            if not firebase_admin._apps:
                cred = credentials.Certificate(CREDENTIALS_FILE)
                firebase_admin.initialize_app(cred)
            self.db = firestore.client()
            print("✅ Conectado a Infraestructura Nexus.")
        except Exception as e:
            print(f"❌ Error conectando a Firebase: {e}")
            sys.exit(1)

        # 3. Conexión Outlook
        self.connect_outlook()

    def configure_ai(self):
        """Configura Google Gemini y carga la personalidad"""
        try:
            genai.configure(api_key=GOOGLE_API_KEY)
            self.model = genai.GenerativeModel('gemini-1.5-flash')
            
            if os.path.exists(PROMPT_PATH):
                with open(PROMPT_PATH, "r", encoding="utf-8") as f:
                    self.system_prompt = f.read()
                print("🧠 Personalidad J.A.R.V.I.S. cargada desde archivo.")
            else:
                self.system_prompt = DEFAULT_PROMPT
                print("⚠️ Usando personalidad base (No se encontró archivo de prompt).")
                
        except Exception as e:
            print(f"⚠️ Error configurando IA: {e}. Se usará modo 'Básico'.")
            self.model = None

    def connect_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6) 
            print("✅ Enlace con Outlook establecido.")
        except Exception as e:
            print(f"❌ Error Outlook: {e}")

    def resolve_sender_email(self, msg):
        """Intenta obtener el SMTP real si es una dirección Exchange Legacy"""
        email = ""
        try:
            sender = msg.Sender
            if sender.AddressEntryUserType == 30: # ExchangeUser
                email = sender.GetExchangeUser().PrimarySmtpAddress
            else:
                email = msg.SenderEmailAddress
        except:
             pass

        # Si el email sigue siendo basura de Exchange o está vacío, FORZAR TU EMPRESARIAL
        if not email or "/O=" in email.upper():
            return AUTHORIZED_EMAILS[0]
            
        return email

    def listen_and_act(self):
        print(f"📡 Escuchando frecuencias... (Asuntos: {TRIGGER_PREFIXES})")
        
        while True:
            try:
                # 1. Leer Comandos
                items = self.inbox.Items.Restrict("[UnRead] = True")
                
                for msg in items:
                    try:
                        subject = msg.Subject.upper() if msg.Subject else ""
                        sender = msg.SenderName
                        
                        # Resolver email real
                        email_addr = self.resolve_sender_email(msg)
                        
                        is_command = any(subject.startswith(p) for p in TRIGGER_PREFIXES)
                        
                        if is_command:
                            print(f"\n📩 MENSAJE ENTRANTE de {sender} (Respuesta a: {email_addr}): {subject}")
                            msg.UnRead = False
                            
                            full_command = subject
                            for p in TRIGGER_PREFIXES:
                                full_command = full_command.replace(p, "")
                            
                            self.process_command(full_command.strip(), msg, email_addr)
                            
                    except Exception as e:
                        print(f"⚠️ Error procesando mensaje: {e}")
                
                # 2. Revisar Tareas Completadas (REPORTING)
                self.check_completions()
                
                time.sleep(5) 
                
            except Exception as e:
                print(f"⚠️ Interferencias en el ciclo principal: {e}")
                time.sleep(5)
                self.connect_outlook()

    def generate_jarvis_response(self, user_input, system_result, is_final_report=False):
        """Usa Gemini para generar la respuesta con personalidad"""
        if not self.model:
            return system_result

        try:
            role_desc = "un asistente irónico que acaba de recibir una orden" if not is_final_report else "un asistente que entrega el trabajo terminado"
            
            input_text = (
                f"{self.system_prompt}\n\n"
                f"SITUACIÓN: Actúas como {role_desc}.\n"
                f"Contexto: {user_input}\n"
                f"Datos Técnicos: '{system_result}'\n\n"
                f"Redacta el correo brevísimo para el Señor Ariel. Firma exactamente como: 'Atte. JARVIS - Asistente de Automatización de Ariel Mella'."
            )
            
            response = self.model.generate_content(input_text)
            return response.text
        except Exception as e:
            return system_result

    def process_command(self, command, msg, sender_email):
        try:
            # Intentar responder inmediatamente
            reply = msg.Reply()
            
            cmd_upper = command.upper()
            technical_result = ""
            
            # Parametros base para el bot
            extra_params = {
                "sender_email": sender_email, 
                "users_command": command,
                "origin_email_id": msg.EntryID # GUARDAMOS EL ID PARA RESPONDER DESPUES
            }
            
            if "STATUS" in cmd_upper or "ESTADO" in cmd_upper:
                technical_result = self.check_status()
                
            elif "AUDITOR" in cmd_upper:
                parts = cmd_upper.split("AUDITOR")
                almacen = "SGVT"
                if len(parts) > 1 and parts[1].strip():
                    almacen = parts[1].strip().split(" ")[0]
                
                extra_params["almacen"] = almacen
                technical_result = self.trigger_bot(
                    "AUDITOR", 
                    f"Solicitud remota para {almacen}",
                    extra_params=extra_params
                )
                
            elif "ZONALES" in cmd_upper:
                technical_result = self.trigger_bot(
                    "ZONALES", 
                    f"Solicitud remota Zonales",
                    extra_params=extra_params
                )
                
            elif "AYUDA" in cmd_upper:
                technical_result = "Comandos: AUDITOR [ALMACEN], ZONALES, STATUS."
            else:
                technical_result = f"Comando '{command}' desconocido."

            print("🧠 Generando confirmación...")
            final_body = self.generate_jarvis_response(command, technical_result)
            
            reply.Body = final_body
            reply.Send()
            print(f"📤 Respuesta de recepción enviada.")
            
        except Exception as e:
            print(f"⚠️ Error procesando comando: {e}")

    def trigger_bot(self, bot_type, description, extra_params=None):
        try:
            params = {'source': 'email_agent', 'email_reported': False}
            if extra_params:
                params.update(extra_params)
                
            doc_ref = self.db.collection('ordenes_bot').add({
                'tipo_bot': bot_type,
                'status': 'pending',
                'parametros': params,
                'fecha_creacion': firestore.SERVER_TIMESTAMP,
                'descripcion': description
            })
            return f"Orden iniciada ID: {doc_ref[1].id}."
        except Exception as e:
            return f"Error al iniciar: {e}"

    def check_status(self):
        try:
            docs = self.db.collection('ordenes_bot')\
                .order_by('fecha_creacion', direction=firestore.Query.DESCENDING)\
                .limit(5).stream()
            report = ""
            for doc in docs:
                d = doc.to_dict()
                report += f"- {d.get('tipo_bot')}: {d.get('status')}\n"
            return report
        except Exception as e:
            return f"Error leyendo DB: {e}"

    def check_completions(self):
        """Busca tareas finalizadas que requieran reporte por email"""
        try:
            # Consulta refinada: Éxito + No Reportado
            # Nota: Si falla por falta de índice compuesto, revertir a filtro en memoria
            docs = self.db.collection('ordenes_bot')\
                .where('status', '==', 'success')\
                .where('parametros.email_reported', '==', False)\
                .limit(20).stream()
            
            docs_list = list(docs)
            if len(docs_list) > 0:
                print(f"[DEBUG] Encontradas {len(docs_list)} tareas pendientes de reporte.")

            for doc in docs_list:
                d = doc.to_dict()
                # Ya filtramos en la query, pero validamos safety
                self.send_completion_email(doc.id, d)

        except Exception as e:
            # Fallback: Si error de índice, usar filtro memoria
            try:
                # print(f"⚠️ Index Error (posible), reintentando en memoria: {e}")
                docs = self.db.collection('ordenes_bot')\
                    .where('status', '==', 'success')\
                    .limit(20).stream()
                
                for doc in docs:
                    d = doc.to_dict()
                    p = d.get('parametros', {})
                    if p.get('source') == 'email_agent' and p.get('email_reported') == False:
                        self.send_completion_email(doc.id, d)
            except:
                pass

    def send_completion_email(self, doc_id, data):
        print(f"✅ Tarea completada detectada: {doc_id}. Preparando reporte...")
        
        params = data.get('parametros', {})
        target_email = params.get('sender_email')
        original_cmd = params.get('users_command', 'Tarea Solicitada')
        result_file = data.get('result_payload') 
        origin_id = params.get('origin_email_id')
        
        mail = None
        
        # 1. INTENTAR RESPONDER AL HILO ORIGINAL (REPLY)
        if origin_id:
            try:
                original_msg = self.namespace.GetItemFromID(origin_id)
                mail = original_msg.Reply()
                print(f"   ↩️ Respondiendo al hilo original de: {original_msg.SenderName}")
            except Exception as e:
                print(f"   ⚠️ No se encontró el email original para responder ({e}). Usando nuevo correo.")
        
        # 2. FALLBACK: CREAR NUEVO CORREO SI FALLA EL REPLY
        if not mail:
            if not target_email or "/O=" in target_email.upper():
                target_email = AUTHORIZED_EMAILS[0]
                print(f"   ⚠️ Email original inválido. Redirigiendo reporte a: {target_email}")
            
            mail = self.outlook.CreateItem(0)
            mail.To = target_email
            mail.Subject = f"NEXUS REPORT: {data.get('tipo_bot')} Finalizado"
            print(f"   📧 Creando nuevo correo para: {target_email}")

        try:
            # Generar cuerpo con IA
            body_text = self.generate_jarvis_response(
                f"El usuario pidió '{original_cmd}'. La tarea terminó exitosamente.", 
                f"Archivo generado: {result_file}",
                is_final_report=True
            )
            
            # Preservar el cuerpo original si es Reply para no borrar el historial?
            # En Outlook COM, mail.Body sobreescribe texto plano. 
            # Para insertar al inicio sin borrar historial, se suele concatenar.
            # Pero Gemini genera un texto completo. Vamos a concatenar simple.
            # mail.Body = body_text + "\n\n" + "-"*20 + "\n" + mail.Body 
            # (Esto a veces rompe formato HTML, pero mail.Body es texto plano usualmente en COM simple)
            
            mail.Body = body_text

            if result_file and os.path.exists(result_file):
                print(f"   📎 Adjuntando: {result_file}")
                mail.Attachments.Add(result_file)
            else:
                mail.Body += "\n\n(Nota: No se generó archivo adjunto o no se encontró en el servidor)"

            mail.Send()
            print(f"   📤 Reporte enviado exitosamente.")
            
            self.mark_reported(doc_id)
            
        except Exception as e:
            print(f"❌ Error enviando reporte final: {e}")

    def mark_reported(self, doc_id):
        self.db.collection('ordenes_bot').document(doc_id).update({
            'parametros.email_reported': True
        })

def main():
    agent = NexusEmailCommander()
    agent.listen_and_act()

if __name__ == "__main__":
    main()
