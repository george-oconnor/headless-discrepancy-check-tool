import win32com.client as win32
import logging, smtplib, keyring
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os.path import basename
from os import makedirs

makedirs("./logs", exist_ok=True)
logger = logging.getLogger(__name__)
handler = logging.FileHandler('./logs/email.log')
formatter = logging.Formatter("%(asctime)s | %(name)s | %(levelname)s | %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)
logger.debug("-------Starting Execution-------")

def sendMail(subject:str, recipients:str, body:str, cc_recipients:str="", bcc_recipients:str="", send:bool=True) -> None:
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = recipients
        mail.CC = cc_recipients
        mail.BCC = bcc_recipients
        #attachment = mail.Attachments.Add(os.getcwd() + "\\Currencies.png")
        #attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "currency_img")
        mail.HTMLBody = body
        #mail.Attachments.Add(os.getcwd() + "\\Currencies.xlsx")
        if send==True:
            mail.Send()
            logger.info("Email successfully sent")
        else:
            logger.info("Email not sent but got details successfully")
    except Exception as e:
        logger.critical("Failed to send email:")
        logger.critical(e, exc_info=True)

def unattended_send_email(subject:str, body:str, mail_type:str, username:str, passwd:str, to:str, cc:str="", bcc:str="", files:list=None) -> None:
    type_mail = {
        'error':r'****ERROR**** :: ',
        'success':r'Successful run :: ',
        'warning':r'Warning/Info :: ',
        'none':r''
    }
    sub_prefix = type_mail[mail_type]

    server = 'smtp.outlook.com'
    port = '587'

    msg = MIMEMultipart('alternative')
    msg['Subject'] = sub_prefix+subject
    msg['From'] = username
    msg['To'] = to
    msg['CC'] = cc
    msg['BCC'] = bcc
    msg.attach(MIMEText(body, 'html'))
    
    for f in files or []:
        with open(f, "rb") as file:
            part = MIMEApplication(file.read(), Name=basename(f))
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    try:
        with smtplib.SMTP(server, port) as smtp:
            smtp.starttls()
            smtp.ehlo()
            smtp.login(username, passwd)
            smtp.sendmail(username, to, msg.as_string())
            logger.info("Sent unattended email")
            smtp.quit()
    except Exception as e:
        logger.critical("Failed to send unattended mail")
        logger.critical(e, exc_info=True)

if __name__ == "__main__":
    username =  keyring.get_password("attendance_sharepoint", "username")
    passwd = keyring.get_password("attendance_sharepoint", username)
    if username is None: logger.error("Failed to get username from keyring")
    if passwd is None: logger.error("Failed to get password from keyring")
    try:
        unattended_send_email("Unattended Email Test", "test", 'warning', username, passwd, "goconnor@instituteofeducation.ie", "oconnorgeorge99@gmail.com")
        logger.info("Successfully sent test email")
    except Exception as e:
        logger.error("Failed to send test email")
        logger.error(e, exc_info=True)

logger.debug("-------Finished Execution-------")