

import re
from datetime import datetime
import pytz

# Email Imports
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders

class Shift():

    def __init__(self):
        self.day = None
        self.start_date_time_obj = None
        self.end_date_time_obj = None
        self.role = None
    # End of def __init__()
# End of clas shift()


def move_email(mail, mail_id, location):
    result = mail.copy(mail_id, location)
    
    if result[0] == 'OK':
        mail.store(mail_id, '+FLAGS', '\\Deleted')
        mail.expunge()

# End of move_email()


def send_mail(send_from, send_to, subject, message, files=[],
              server="localhost", port=587, username='', password='',
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format("NSD_ICS.ics"))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()


def iphone_fix(mail_content):

    # Handle Iphone forwarding issues

    fixed_lines = ""
    
    for line in mail_content.split('\n'):
        if len(line) == 0:
            fixed_lines += "\n"
            continue
        # End of if len()
        if line[:4] == ">=20":
            fixed_lines += "\n"
            continue
        # End of if len()
        if line[:2] == "> ":
            line = line[2:]
        # End of if len()
        if line[-2:] == "=\r":
            fixed_lines+= line[:-2]
            continue
        # End of if len()
        
        fixed_lines += line
        fixed_lines += '\n'

    # End of for line in mail_content

    if fixed_lines[-1] == "\n":
        fixed_lines = fixed_lines[:-1]

    return fixed_lines

# End of def iphone_fix()


def parse_for_email_shifts(mail_content):

    name = None
    week = None
    shifts = []

    ptrn_name = re.compile(r'Hello ([\w ]*),')
    results_name = ptrn_name.findall(mail_content)

    if len(results_name) == 0:
        print("Unable to read the name")
        return (None, None, None)
    # End of if len(result_week)

    name = results_name[0]

    ptrn_week = re.compile(r'week of \**([\d\w, ]+)\**')
    results_week = ptrn_week.findall(mail_content)

    if len(results_week) == 0:
        print("Unable to read the week")
        return (None, None, None)
    # End of if len(result_week)

    week = results_week[0]
    
    ptrn = re.compile(r'\**\w\w\w (\w\w\w) (\d+), (\d\d\d\d)\**((?:[\s -]+\d+:\d\d \w\w - \d+:\d\d \w\w - [\w ]+)+)',re.MULTILINE)
    results = ptrn.findall(mail_content)
    
    if len(results) == 0:
        print("Unable to read the working days")
        return (None, None, None)
    # End of if len(results)

    for result in results:
        
        month = result[0]
        day = result[1]
        if len(day) == 1:
            day = '0'+day
        # End of if len(day)
        year = result[2]
        
        date_str = "%s %s, %s" % (month, day, year)
        
        day_content = result[3]
        
        ptrn_shift = re.compile(r'(\d)+:(\d\d) (\w\w) - (\d+):(\d\d) (\w\w) - ([\w ]+)',re.MULTILINE)
        results_shift = ptrn_shift.findall(day_content)
        
        # Get the working times in the days working
        
        for shift_result in results_shift:

            shift_obj = Shift()
            shift_obj.day = date_str
            
            start_hour = shift_result[0]
            if len(start_hour) == 1:
                start_hour = '0'+start_hour
            # End of if len(star_hour)
            start_min = shift_result[1]
            start_am_pm = shift_result[2]
            
            end_hour = shift_result[3]
            if len(end_hour) == 1:
                end_hour = '0'+end_hour
            # End of if len(end_hour)
            end_min = shift_result[4]
            end_am_pm = shift_result[5]
            
            role = shift_result[6]
            
            start_time_string = month+' '+day+' '+year+' '+start_hour+' '+start_min+' '+start_am_pm
            end_time_string = month+' '+day+' '+year+' '+end_hour+' '+end_min+' '+end_am_pm
            
            start_date_time_obj = datetime.strptime(start_time_string, '%b %d %Y %I %M %p')
            end_date_time_obj = datetime.strptime(end_time_string, '%b %d %Y %I %M %p')
            
            start_date_time_obj = pytz.timezone('US/Eastern').localize(start_date_time_obj)
            end_date_time_obj = pytz.timezone('US/Eastern').localize(end_date_time_obj)

            shift_obj.start_date_time_obj = start_date_time_obj
            shift_obj.end_date_time_obj = end_date_time_obj
            shift_obj.role = role
            
            shifts.append(shift_obj)
        # End of for shift_result in ...
    # End of for result in results
                        
    return (name, week, shifts)
# End of def parse_for_email_shifts()
