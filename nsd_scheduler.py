#!/usr/bin/env python3

# Imports
import argparse
import configparser
import email
import imaplib
import os
import pytz
import re
import sys
import time

# From Import
from icalendar import Calendar, Event
from datetime import datetime

# Email Imports
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders


import scheduler_functions as sf



###########
# MAIN CODE
###########

parser = argparse.ArgumentParser()
parser.add_argument('-c', '--config', type=str, required=True)
args = parser.parse_args()

config_path = args.config

if os.path.exists(config_path) == False:
    print("The config path does not exist")
    sys.exit(1)

config = configparser.ConfigParser()

config.read(config_path)

EMAIL = config['DEFAULT']['EMAIL']
PASSWORD = config['DEFAULT']['PASSWORD']
SERVER = config['DEFAULT']['SERVER']
FILENAME = config['DEFAULT']['FILENAME']
WAITTIME = int(config['DEFAULT']['WAITTIME'])
RUNTIME = int(config['DEFAULT']['RUNTIME'])

end_date_time = time.time()+RUNTIME

while(time.time() < end_date_time):

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

                mail_content = sf.iphone_fix(mail_content)
                (name,
                 week,
                 shifts_list) = sf.parse_for_email_shifts(mail_content)

                if (name is None) or (week is None) or (shifts_list is None):
                    sf.move_email(mail, i, 'Unable To Complete')
                    continue
                # End of if (name is None) or ...

                cal = Calendar()
                dates = []

                for shift_obj in shifts_list:
                    role = shift_obj.role
                    start_date_time_obj = shift_obj.start_date_time_obj
                    end_date_time_obj = shift_obj.end_date_time_obj
                    dates.append(shift_obj.day)
                    
                    event = Event()
                    event.add('summary', role)
                    event.add('dtstart', start_date_time_obj)
                    event.add('dtend', end_date_time_obj)
                    #event.add('description', "Sample Description")
                    event.add('location', "North Shore Dog, Danvers MA 01923")
                    cal.add_component(event)
                    # End of for shift_result in results_shift
                    
                # End of for shift_obj in shifts_list
                

                f = open(FILENAME, 'wb')
                f.write(cal.to_ical())
                f.close()
                
                
                
                send_from = EMAIL
                send_to = [mail_from]
                subject = "North Shore Dog Calender Invites Week Of %s" % (week)
                dates_list = "\n".join(dates)
                text = "Hello %s,\nAttached are the calender invites for the dates:\n%s" % (name, dates_list)
                text += "\n\n"
                text += "Gingr: nsdog.gingrapp.com\n"
                text += "Slack: northshoredog.slack.com\n"
                text += "\n\n"
                text += "This is an auto generated email report.\n"
                files = [FILENAME]

                print(mail_from)
                print(name)
                print(week)
                value = datetime.fromtimestamp(time.time())
                print(value.strftime('%Y-%m-%d %H:%M:%S'))
                
                sf.send_mail(send_from, send_to, subject, text, files, server="smtp.gmail.com", port=587, username=EMAIL, password=PASSWORD, use_tls=True)
                
                sf.move_email(mail, i, 'Completed')
                    
            # End of if isinstance(response_part, tuple)
        
        # End of for response_part
    
    # End of for i in mail_ids:

    mail.close()

    if len(mail_ids) == 0:
        time.sleep(WAITTIME)
    else:
        end_date_time = time.time()+RUNTIME
    # End of while True
