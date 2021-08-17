import logging
import json
import azure.functions as func
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from binascii import a2b_base64
import base64
import smtplib
from azure.cosmosdb.table.tableservice import TableService

def main(msg: func.QueueMessage) -> None:
    logging.info('Python queue trigger function processed a queue item.')
   
    result = {
        'id': msg.id,
        'body': msg.get_body().decode('utf-8'),
        'expiration_time': (msg.expiration_time.isoformat()
                            if msg.expiration_time else None),
        'insertion_time': (msg.insertion_time.isoformat()
                           if msg.insertion_time else None),
        'time_next_visible': (msg.time_next_visible.isoformat()
                              if msg.time_next_visible else None),
        'pop_receipt': msg.pop_receipt,
        'dequeue_count': msg.dequeue_count
    }

    data = json.loads(result['body'].replace("'",'"'))
    sender = data.get('sender')
    # emails = data.get('emails')
    table_service = TableService(account_name='queuefrontera', account_key='ivU72c/zwNUs1o3uJV1J9LYUjBuoN8Papj6QhSeqnSrCcCOHVLRByFBtYC9UdKsFwgFSgNvWEhBbjzR/lG6FCA==')
    emails = table_service.query_entities('correos')
    # inicio de sesión
    sesion_smtp = smtplib.SMTP_SSL(sender['smtp'], sender['port'])
    try:
        sesion_smtp.login(sender['user'], sender['pass'])

        correos = []
        
        for email in emails:
            msg = MIMEMultipart()
            msg['From'] = sender['user']
            msg['To'] = email['to']
            msg['Cc'] = email['Cc']
            if email['Cc'] != "":
                receivers = [email['to']] + [email['Cc']]
            else:
                receivers = email['to']
            if email['Cco'] != "":
                receivers = receivers + [email['Cco']]
            else:
                pass
            msg['Subject'] = email['subject']
            # Cuerpo del correo electrónico final
            base64_message = email['body']
            base64_bytes = base64_message.encode('ascii')
            message_bytes = base64.b64decode(base64_bytes)
            body = message_bytes.decode('ascii')
            # Hidden del correo
            # hidden = f'''<div id='hidden' style='display: none'>
            #           {email['hidden']}
            #           </div>'''
            hidden = f'''<div id='campo oculto'>
                        <!-- {email['status']} -->
                        </div>'''
            email_body = body + hidden
            msg.attach(MIMEText(email_body, 'html', 'utf-8'))
            part = MIMEBase('application', 'octet-stream')
            if email['attachment'] != "":
                part.set_payload(a2b_base64(email['attachment']))
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', "attachment; filename= prueba.xlsx")
                msg.attach(part)
            texto = msg.as_string()
            sesion_smtp.sendmail(sender['user'], receivers, texto)
            correos.append(receivers)
        
        sesion_smtp.quit()

        return logging.info(f'''Los correos desde: {sender['user']}, 
                              fueron correctamente enviados a {correos}''')
    
    except smtplib.SMTPAuthenticationError:

        return logging.info(f'''Credenciales incorrectas, 
                              verifique usuario y contraseña''')
    

