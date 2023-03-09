import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
import datetime

today = datetime.datetime.now()
first = today.replace(day=1)
last_month = first - datetime.timedelta(days=1)
current_month = last_month.strftime('%B')
current_year = last_month.strftime('%Y')
current_date = f"{today.strftime('%d')}-{today.strftime('%b')}-{today.strftime('%Y')}"

def send_mail(send_from, send_to,send_cc, files, invoice_number, receiver_name, total_amount, smtp_server, smtp_port, smtp_user, smtp_pass, please_confirm_text):
    assert isinstance(send_to, list)

    subject = f'Invoice {invoice_number} from Martrust Corporation Limited for {receiver_name} - Lifting Fees - {current_month} {current_year}'

    text = f"""
        Good day,

        Please find attached your Lifting Fees - {current_month} {current_year} invoice {invoice_number} for USD {total_amount}.

        {please_confirm_text}

        If you have any questions, please let us know.

        Thanks,
        Martrust Corporation Limited
    """

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Cc'] = COMMASPACE.join(send_cc)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)


    try:

        smtp = smtplib.SMTP(smtp_server, smtp_port)
        
        smtp.starttls()
        smtp.login(smtp_user, smtp_pass)
        smtp.sendmail(send_from, send_to + send_cc, msg.as_string())
        # Do something with the SMTP connection
    except smtplib.SMTPConnectError as e:
        print(f"Error connecting to SMTP server: {e}")
    except smtplib.SMTPAuthenticationError as e:
        print(f"Error authenticating with SMTP server: {e}")
    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        smtp.quit()

def send_consolidated_mail(send_from, send_to, files, smtp_server, smtp_port, smtp_user, smtp_pass):

    subject = f'Consolidated Report - {current_date}'

    text = f"""
        Good day,

        Please find the attached Consolidated reports for {current_date}

        If you have any questions, please let us know.

        Thanks,
        BOT 
    """

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)


    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(smtp_user, smtp_pass)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

def send_group_mail(send_from, send_to,send_cc, files, smtp_server, smtp_port, smtp_user, smtp_pass,receiver_name):

    subject = f'MARTRUST Invoice for {receiver_name} for the month of {current_month}-{current_year}- Lifting fees'

    text = f"""
        Good day,

        Please find attached the lifting fee invoices for the month of {current_month}-{current_year} along with the supporting information.
        Please confirm if we can deduct these lifting fees from your USD currency balances.
        
        Thanks,
        Martrust Corporation Limited
    """

    msg = MIMEMultipart('alternative')
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Cc'] = COMMASPACE.join(send_cc)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)


    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(smtp_user, smtp_pass)
    smtp.sendmail(send_from, send_to+send_cc, msg.as_string())
    smtp.close()

def send_mail_cpdesk(send_from, send_to,send_cc, files, receiver_name, smtp_server, smtp_port, smtp_user, smtp_pass):
    assert isinstance(send_to, list)
    # assert isinstance(send_cc, list)

    subject = f'Invoice and details of Charter Parties attended for {receiver_name} in {current_month} {current_year}.'

    text = f"""
        Good Day,

        Please find attached invoice and details of Charter Parties attended in {current_month} {current_year}.
          
        If you have any queries, please do not hesitate to contact us.
         
        Kindly confirm receipt of this email. 
        
        Thanks,
        Marcura Accounts
        Office: +971 4 3636200
    """

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Cc'] = COMMASPACE.join(send_cc)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)


    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(smtp_user, smtp_pass)
    smtp.sendmail(send_from, send_to+send_cc, msg.as_string())
    smtp.close()
