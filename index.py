import win32com.client
import re
import os 

Cuenta = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

Buzon = Cuenta.GetDefaultFolder(6)
Mensajes = Buzon.Items 
directorioDeDestino = r"C:\Users\samyr\OneDrive\Documentos\descargar_excels"
print(f"Número de correos en la bandeja de entrada: {len(Mensajes)}")

codigo = "PARA: Comunidad UNADISTA"
for mensaje in Mensajes:
    cuerpo = mensaje.Body
    #re.search sirve para buscar que contenga el codigo en el cuerpo de correo 
    if re.search(rf"\b{codigo}\b",cuerpo):
        # print(f"asunto : {mensaje.Body}")
        for attachment in mensaje.Attachments:
            print(f"de descargara el archivo{attachment.FileName}")
            ruta_destino=os.path.join(directorioDeDestino,attachment.FileName)
            attachment.SaveAsFile(ruta_destino)
            break
        break
            

