#Exportar librerias a usar 
#Primarias, necesarias
import pandas as pd
from tableau_api_lib import TableauServerConnection
from tableau_api_lib.sample import sample_config
from tableau_api_lib.utils import querying
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from smtplib import SMTP

#Secundarias
from tableau_api_lib.utils.common import flatten_dict_column
from tableau_api_lib.utils import extract_pages
from typing import List

#crear variable fecha 
today = date.today()
date_today = today.strftime("%Y-%m-%d")
#print('done')

# Guardar la configuracion de Tableau
tableau_server_config = {
'tableau_env': {
    'server': 'https:*******',
    'api_version': '3.17',
    'username': '******',
    'password': '***********',
    'site_name': '',
    'site_url': ''
}
}

#Guardar e iniciar sesion de tableau
connection =TableauServerConnection(config_json=tableau_server_config, env='tableau_env')
connection.sign_in()

#Acceder a la vista la cual se va a trabajar 
worbook_view_id='************'#alerta

#Convertir la vista a un dataframe 
view_to_df= querying.get_view_data_dataframe(connection,view_id=worbook_view_id)

#Renombrar estas columnas para una mejor presentacion
view_to_df.rename(columns= {'CÃ³digo Producto':'Codigo Producto','% DesposiciÃ³n':'% Desposiocionamiento' }, inplace=True)
view_to_df =view_to_df[['Codigo Producto','Producto-formato','E-Commerce','Marca','Alerta de Desposicionamiento','% Desposiocionamiento','Precio Sugerido 70%','Precio Mayorista','Precio Mayorista']].sort_values(by='Codigo Producto',ignore_index=True)
#view_to_df.to_excel('C:/Users/blanc/yday_tableau_api.xlsx')

#Guardar un excel con la fecha y generar nombre a la hoja la cual vera el cliente
write =pd.ExcelWriter(r'C:/Users/*******/tableau_alerta'+ date_today+'.xlsx', engine='xlsxwriter')
view_to_df.to_excel(write, sheet_name='Alerta', index=None)
write.save()

#ver la vista, para validacion 
#print(view_to_df)
#print('Done')

#crear variables para el envio del correo desde Gmail
FECHA = date_today
MAIL_TO = [
        "******@******.com"
    ]  
    ## TO es CC en gmail
TO = [ 
    ] 
MAIL_SUBJET = 'Alerta Desposicionamiento Comercial Castro'+ " " + (FECHA)
MAIL_BODY = """
                        <p>Estimado/a Equipo de Comercial Castro,</p>
                        <p>Espero que se encuentre muy bien. Me pongo en contacto con usted para enviarle el archivo adjunto con la alerta de desposicionamiento</p>
                        <p>Si necesita alguna información adicional, puede revisar el reporte directamente o no dude en ponerse en contacto nuestro equipo.</p>
                        <p>Reporte = https:*******</p>
                        <p>Atentamente,</p>
                        <p><b>********.</b></p>
                        """
#IMAGE_FILE  = [['logo',r'C:\Users\blanc\DIR_IMAGE\****-LOGO-1.png']]


#se crea una funcion para enviar correo
def send_email(MAIL_TO,TO,MAIL_SUBJET,MAIL_BODY,**kwargs):
    #Definir valiables
    MAIL_FILE = None
    FILE_NAME = 'C:/Users/*****/tableau_alerta_desposicionamiento '+ date_today+'.xlsx'
    
    for key, value in kwargs.items():
        if key == 'MAIL_FILE':
            MAIL_FILE = value
        if key == 'FILE_NAME':
            FILE_NAME = value

    #configuracion mensaje
    mensaje = MIMEMultipart()
    mensaje["From"]     = "****@*****.com"
    mensaje["To"]       = "******@*******.com".join(MAIL_TO)
    mensaje["Cc"]       = ", ".join(TO)
    mensaje["Subject"]  = MAIL_SUBJET

    
    #CUERPO DE MAIL
    mensaje.attach(MIMEText(MAIL_BODY, "html"))

    with open(FILE_NAME, 'rb') as file:
        attachment = MIMEApplication(file.read(), _subtype='xlsx')
        attachment.add_header('Content-Disposition', 'attachment', filename=FILE_NAME)
        mensaje.attach(attachment)

    #configuracion smtp (ENVIO MAIL)
    smtp = SMTP("smtp.gmail.com",587)
    smtp.starttls()
    smtp.login("correo@gmail.com","password")##("******@g*****.***","*********")
    smtp.sendmail("correo@gmail.com", MAIL_TO , mensaje.as_string().encode('UTF-8'))##("*****@****.cl", MAIL_TO , mensaje.as_string().encode('UTF-8'))
    smtp.quit() 

    MAIL_SUBJET = 'Alerta de Desposicionamiento'+ " " +str(today) ### CAMBIAR SUBJECT DEL MAIL
    #IMAGE_FILE  = [['logo',r'C:\Users\blanc\DIR_IMAGE\GREGARIO-LOGO-1.png']]
    MAIL_BODY = """
                        
                        """
send_email(MAIL_TO, TO, MAIL_SUBJET, MAIL_BODY)
print("Mail enviado: Alerta de Desposicionamiento")
