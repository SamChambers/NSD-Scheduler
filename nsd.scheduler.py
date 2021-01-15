
import email
import imaplib
import re
import sys
from datetime import datetime
import pytz

import time

from icalendar import Calendar, Event


import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
import tempfile


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



    

EMAIL = 'nsd.scheduler@gmail.com'
PASSWORD = 'Blue42Blue42'
SERVER = 'imap.gmail.com'



while(True):

    # connect to the server and go to its inbox
    mail = imaplib.IMAP4_SSL(SERVER)
    mail.login(EMAIL, PASSWORD)


    # we choose the inbox but you can select others
    mail.select('ToSchedule')
    # we'll search using the ALL criteria to retrieve
    # every message inside the inbox
    # it will return with its status and a list of ids
    status, data = mail.search(None, 'ALL')

    # the list returned is a list of bytes separated
    # by white spaces on this format: [b'1 2 3', b'4 5 6']
    # so, to separate it first we create an empty list
    mail_ids = []
    # then we go through the list splitting its blocks
    # of bytes and appending to the mail_ids list
    for block in data:
        # the split function called without parameter
        # transforms the text or bytes into a list using
        # as separator the white spaces:
        # b'1 2 3'.split() => [b'1', b'2', b'3']
        mail_ids += block.split()
    # End of for block in
        
    # mail_ids.reverse()
    

    # now for every id we'll fetch the email
    # to extract its content
    for i in mail_ids:
        
        # the fetch function fetch the email given its id
        # and format that you want the message to be
        status, data = mail.fetch(i, '(RFC822)')
        
        # the content data at the '(RFC822)' format comes on
        # a list with a tuple with header, content, and the closing
        # byte b')'
        for response_part in data:
            # so if its a tuple...
            if isinstance(response_part, tuple):
                # we go for the content at its second element
                # skipping the header at the first and the closing
                # at the third
                message = email.message_from_bytes(response_part[1])
                
                # with the content we can extract the info about
                # who sent the message and its subject
                mail_from = message['from']
                mail_subject = message['subject']
            
            
                mail_date = message['date']
                
                # then for the text we have a little more work to do
                # because it can be in plain text or multipart
                # if its not plain text we need to separate the message
                # from its annexes to get the text
                if message.is_multipart():
                    mail_content = ''
                    
                    # on multipart we have the text message and
                    # another things like annex, and html version
                    # of the message, in that case we loop through
                    # the email payload
                    for part in message.get_payload():
                        # if the content type is text/plain
                        # we extract it
                        if part.get_content_type() == 'text/plain':
                            mail_content += part.get_payload()
                        # End of if part.get_content_type
                    # End of for part in message.
                else:
                    # if the message isn't multipart, just extract it
                    mail_content = message.get_payload()
                # End of if/else message.is_multipart


                # Handle Iphone forwarding issues

                fixed_lines = ""
                
                for line in mail_content.split('\n'):
                    if len(line) == 0:
                        continue
                    if line == ">=20\r":
                        continue
                    if line[:2] == "> ":
                        line = line[2:]
                    
                    if line[-2:] == "=\r":
                        fixed_lines+= line[:-2]
                        continue

                    fixed_lines += line
                    fixed_lines += '\n'

                # End of for line in mail_content
                        

                # print(fixed_lines)
                # print(mail_content)


                mail_content = fixed_lines
            
                ptrn_week = re.compile(r'week of \**([\d\w, ]+)\**')
                results_week = ptrn_week.findall(mail_content)
                ptrn = re.compile(r'\**\w\w\w (\w\w\w) (\d+), (\d\d\d\d)\**((?:[\s -]+\d+:\d\d \w\w - \d+:\d\d \w\w - [\w ]+)+)',re.MULTILINE)
                results = ptrn.findall(mail_content)
                
                if len(results_week) == 0:
                    print("Unable to read the week")
                    sys.exit()
                # End of if len(result_week)
                
                week = results_week[0]
                
                if len(results) == 0:
                    print("Unable to read the working days")
                    sys.exit()
                # End of if len(results)

                
                cal = Calendar()
                dates = []
                
                for result in results:

                    month = result[0]
                    day = result[1]
                    if len(day) == 1:
                        day = '0'+day
                    # End of if len(day)
                    year = result[2]

                    dates.append("%s %s, %s" % (month, day, year))

                    day_content = result[3]

                    ptrn_shift = re.compile(r'(\d)+:(\d\d) (\w\w) - (\d+):(\d\d) (\w\w) - ([\w ]+)',re.MULTILINE)
                    results_shift = ptrn_shift.findall(day_content)
                    
                    # Get the working times in the days working
                    
                    for shift_result in results_shift:
                    
                    
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
                        
                        event = Event()
                        event.add('summary', role)
                        event.add('dtstart', start_date_time_obj)
                        event.add('dtend', end_date_time_obj)
                        #event.add('description', "Sample Description")
                        event.add('location', "North Shore Dog, Danvers MA 01923")
                        cal.add_component(event)
                    # End of for shift_result in results_shift
                    
                # End of for result in results
                
                # temp = tempfile.NamedTemporaryFile()
                filename = r"C:\Users\samjc\Documents\nsdScheduler\nsd_scheduler.ics"
                
                #C:\Users\samjc\Documents\nsdScheduler

                f = open(filename, 'wb')
                f.write(cal.to_ical())
                f.close()
                
                print(mail_from)
                
                send_from = EMAIL
                send_to = [mail_from]
                subject = "North Shore Dog Calender Invites Week Of %s" % (week)
                dates_list = "\n".join(dates)
                text = "Attached are the calender invites for the dates:\n%s" % (dates_list)
                files = [filename]
                
                print(filename)
                
                send_mail(send_from, send_to, subject, text, files, server="smtp.gmail.com", port=587, username=EMAIL, password=PASSWORD, use_tls=True)
                
                # temp.close()
                
                result = mail.copy(i, 'Completed')
                
                if result[0] == 'OK':
                    mail.store(i, '+FLAGS', '\\Deleted')
                    mail.expunge()
                    
                # End of if result[0]
            # End of if isinstance(response_part, tuple)
        
        # End of for response_part
    
    # End of for i in mail_ids:

    mail.close()

    if len(mail_ids) == 0:
        time.sleep(10)
# End of while True
