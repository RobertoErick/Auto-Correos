import pythoncom
pythoncom.CoInitialize()

import win32com.client
from datetime import datetime
import os
import xlwings as xw

def exportar_correos_a_excel():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    account_name = "correo@dominio.com"

    # Crear una nueva instancia de Excel usando xlwings
    app = xw.App(visible=False)
    workbook = app.books.add()
    worksheet = workbook.sheets[0]

    worksheet.range('A1').value = ["Correo recibido", "Hora de Recepción", "Respondido", "Hora de Respuesta", "Correo Enviado", "Hora de Envío"]

    i = 2
    today_date = datetime.now().date()

    # Encuentra la cuenta especificada
    account = None
    for acc in namespace.Accounts:
        if acc.DisplayName == account_name or acc.SmtpAddress == account_name:
            account = acc
            break

    if account is None:
        print(f"Cuenta no encontrada: {account_name}")
        return

    inbox = account.DeliveryStore.GetDefaultFolder(6)  # olFolderInbox
    sent_folder = account.DeliveryStore.GetDefaultFolder(5)  # olFolderSentMail

    for item in inbox.Items:
        if item.Class == 43:  # olMailItem
            mail = item
            if mail.ReceivedTime.date() == today_date:
                responded = False
                response_time = None

                for sent_item in sent_folder.Items:
                    if sent_item.Class == 43:  # olMailItem
                        sent_mail = sent_item
                        if sent_mail.ConversationID == mail.ConversationID:
                            responded = True
                            response_time = sent_mail.SentOn
                            break

                worksheet.range(f'A{i}').value = [mail.Subject, mail.ReceivedTime, "Sí" if responded else "No", response_time, "", ""]
                i += 1

    # Agregar los correos enviados que no son respuestas
    for sent_item in sent_folder.Items:
        if sent_item.Class == 43:  # olMailItem
            sent_mail = sent_item
            if sent_mail.SentOn.date() == today_date:
                # Verificar si el correo enviado es una respuesta
                if not sent_mail.Subject.startswith("RE:"):
                    worksheet.range(f'A{i}').value = ["", "", "", "", sent_mail.Subject, sent_mail.SentOn]
                    i += 1

    # Construir la ruta del archivo
    directory = "C:\\Sistema\\Reportes"
    if not os.path.exists(directory):
        os.makedirs(directory)

    file_path = os.path.join(directory, f"{account_name}_CorreosRecibidos_{today_date.strftime('%Y-%m-%d')}.xlsx")

    # Guardar y cerrar el archivo de Excel
    try:
        workbook.save(file_path)
        print(f"Los correos de hoy han sido exportados a {file_path}.")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")
    finally:
        workbook.close()
        app.quit()

    # Enviar el archivo por correo
    enviar_correo(file_path, "correo2@dominio.com")

    # Forzar la finalización del proceso de Excel en caso de que algo haya quedado abierto
    # os.system("taskkill /f /im excel.exe")

def enviar_correo(file_path, destinatario):
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Crear un objeto MailItem directamente
    mail = outlook.CreateItem(0)
    mail.Subject = "Correos del día"
    mail.Body = "Adjunto se encuentra el archivo con los correos del día."
    mail.To = destinatario
    mail.Attachments.Add(file_path)

    # Enviar el correo
    mail.Send()
    print(f"Correo enviado a {destinatario} con el archivo adjunto.")

if __name__ == "__main__":
    exportar_correos_a_excel()
