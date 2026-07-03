import imaplib
import email
import csv
import os
import re
import sys
import traceback
from datetime import datetime
from email.header import decode_header

# Configuración
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
IMAP_SERVER = 'outlook.office365.com'
IMAP_PORT = 993

# Archivo CSV
CSV_FILE = 'reporte_operaciones.csv'

def extract_caja_trailer(body):
    """Extrae Caja/Trailer del cuerpo del correo (R + dígitos o 5 dígitos)"""
    if not body:
        return ""
    
    # Busca patrones: R seguido de dígitos o 5 dígitos solos
    matches = re.findall(r'(R\d+|\b\d{5}\b)', body)
    
    if matches:
        return ", ".join(matches)
    return ""

def detect_operation_type(sender, subject):
    """Detecta el tipo de operación según remitente y asunto"""
    sender_lower = sender.lower()
    subject_lower = subject.lower()
    
    # Exportación: de Cecilia Ramón, asunto con "Loyalty Relacion CCP"
    if 'cecilia' in sender_lower and 'ramon' in sender_lower:
        if 'loyalty' in subject_lower and 'relacion' in subject_lower and 'ccp' in subject_lower:
            return 'Exportacion'
    
    # Importación: asunto con "Carta Porte", de personal Radar
    if 'carta porte' in subject_lower:
        if 'radar' in sender_lower:
            return 'Importacion'
    
    # Cruces: de personal Radar, especialmente Abraham Alvarez
    if 'radar' in sender_lower:
        if 'abraham' in sender_lower and 'alvarez' in sender_lower:
            return 'Cruces'
        elif 'carta porte' not in subject_lower:
            return 'Cruces'
    
    return 'Otra'

def decode_email_subject(subject):
    """Decodifica el asunto del correo"""
    decoded_parts = decode_header(subject)
    decoded_subject = ""
    
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            decoded_subject += part.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_subject += part
    
    return decoded_subject.strip()

def get_email_body(msg):
    """Extrae el cuerpo del correo"""
    body = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    break
                except:
                    body = part.get_payload()
                    break
    else:
        try:
            body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        except:
            body = msg.get_payload()
    
    return body.strip()

def fetch_emails():
    """Conecta a Outlook y obtiene los correos nuevos"""
    try:
        print(f"[INFO] Intentando conectar a {IMAP_SERVER}:{IMAP_PORT}")
        print(f"[INFO] Usuario: {EMAIL_USER}")
        
        # Conectar a IMAP
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        print("[SUCCESS] Conexión IMAP exitosa")
        
        mail.login(EMAIL_USER, EMAIL_PASSWORD)
        print("[SUCCESS] Login exitoso")
        
        mail.select('INBOX')
        print("[INFO] Carpeta INBOX seleccionada")
        
        # Obtener todos los correos
        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        
        print(f"[INFO] Total de correos encontrados: {len(email_ids)}")
        
        emails_data = []
        
        # Procesar últimos 50 correos
        for i, email_id in enumerate(email_ids[-50:]):
            try:
                status, msg_data = mail.fetch(email_id, '(RFC822)')
                
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        
                        # Extraer información
                        sender = msg.get('From', '')
                        subject = decode_email_subject(msg.get('Subject', ''))
                        date_str = msg.get('Date', '')
                        body = get_email_body(msg)
                        
                        # Procesar fecha
                        try:
                            from email.utils import parsedate_to_datetime
                            date_obj = parsedate_to_datetime(date_str)
                            fecha_recepcion = date_obj.strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            fecha_recepcion = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        
                        # Detectar tipo de operación
                        tipo_operacion = detect_operation_type(sender, subject)
                        
                        # Extraer Caja/Trailer
                        caja_trailer = extract_caja_trailer(body)
                        
                        # Limitar observaciones a primeras 500 caracteres
                        observaciones = body[:500] if body else ""
                        
                        emails_data.append({
                            'Fecha Recepcion': fecha_recepcion,
                            'Tipo Operacion': tipo_operacion,
                            'Asunto': subject,
                            'Caja/Trailer': caja_trailer,
                            'Observaciones': observaciones,
                            'Remitente': sender
                        })
                        
                        print(f"[✓] Correo {i+1} procesado: {tipo_operacion}")
            except Exception as e:
                print(f"[ERROR] Procesando correo {i+1}: {e}")
                traceback.print_exc()
        
        mail.close()
        mail.logout()
        
        print(f"[SUCCESS] {len(emails_data)} correos procesados correctamente")
        return emails_data
    
    except Exception as e:
        print(f"[CRITICAL ERROR] Error al conectar a Outlook: {e}")
        traceback.print_exc()
        sys.exit(1)

def save_to_csv(emails_data):
    """Guarda los correos en un CSV"""
    try:
        with open(CSV_FILE, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Fecha Recepcion', 'Tipo Operacion', 'Asunto', 'Caja/Trailer', 'Observaciones', 'Remitente']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for email_data in emails_data:
                writer.writerow(email_data)
        
        print(f"[SUCCESS] CSV guardado: {CSV_FILE}")
        print(f"[INFO] Total de correos en reporte: {len(emails_data)}")
    
    except Exception as e:
        print(f"[ERROR] Error al guardar CSV: {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    print("=" * 60)
    print("INICIANDO MONITOREO DE OUTLOOK")
    print("=" * 60)
    
    emails = fetch_emails()
    
    if emails:
        save_to_csv(emails)
        print("=" * 60)
        print("✓ REPORTE ACTUALIZADO EXITOSAMENTE")
        print("=" * 60)
    else:
        print("=" * 60)
        print("[WARNING] No se encontraron correos")
        print("=" * 60)
