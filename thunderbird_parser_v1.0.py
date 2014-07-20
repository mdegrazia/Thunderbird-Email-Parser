#
#thunderbird_parser v1.0
#This program parses the raw emails file created by Thunderbird
#located under \Users\%USERNAME%\AppData\Roaming\Thunderbird\Profiles\[random].default\
#
#These files do not have a file extension and contain emails in MIME format
#It will parse the Header information(To, from, CC, BC, Date and Subject) into an Excel file
#and create a link to the .eml file. It will also list the attachments.
#
#It requires the xlwt library, which can be installed using easy install or the Windows installer for windows
#
#To read more about it, visit my blog at http://az4n6.blogspot.com/
#
# Copyright (C) 2014 Mari DeGrazia (arizona4n6@gmail.com)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You can view the GNU General Public License at <http://www.gnu.org/licenses/>

__author__ = 'arizona4n6@gmail.com (Mari DeGrazia)'
__version__ = '1.0'
__copyright__ = 'Copyright (C) 2014 Mari DeGrazia'
__license__ = 'GNU'

from email.parser import Parser
from email.utils import parsedate,parsedate_tz
import time
import datetime
import xlwt 
import os
import sys
import string
import sqlite3
import json
import re
parser = Parser()

#make new excel worksheet
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Email')

#format for Date and Header column in Excel
style0 = xlwt.XFStyle()
style0.num_format_str = 'yyyy-mm-dd hh:mm:ss'

style = xlwt.easyxf('font: underline single, color blue')
style_header = xlwt.easyxf('font: bold on')

#write column headers
worksheet.write(0,0,"Filename",style_header)
worksheet.write(0,1,"From", style_header)
worksheet.write(0,2,"To", style_header)
worksheet.write(0,3, "CC", style_header)
worksheet.write(0,4, "BCC", style_header)
worksheet.write(0,5, "Subject", style_header)
worksheet.write(0,6, "Raw Date", style_header)
worksheet.write(0,7, "Converted Date UTC", style_header)
worksheet.write(0,8, "Email File", style_header)
worksheet.write(0,9, "Attachments",style_header)
worksheet.write(0,10, "Read",style_header)
worksheet.write(0,11, "Deleted",style_header)
worksheet.write(0,12, "Replied To",style_header)
worksheet.write(0,13, "Forwarded",style_header)
worksheet.write(0,14, "Message-ID",style_header)

#this holds email body
email = ''

#global variable to keep track of a unique filename count and Excel rows
global count

#global variable to keep track of sqlite database values so multiple queries are not ran

global read_id
global repliedTo_id
global forwarded_id
global current_file

#add one because we already have a header and Excel entries will need start on the next row
count = 1

def remove_ascii_non_printable(str):
    return ''.join([ch for ch in str if ord(ch) > 31 and ord(ch) < 126 or ord(ch) ==9])

def process_email(email):
    global count   
    global email_folder
    global seperator
    global current_file
        
    email_p = parser.parsestr(email)
  
    #blank and corrupt subjects make it blow up
    try:
        subject = remove_ascii_non_printable(email_p.get('Subject'))
    except:
        subject = "Blank Subject"
        
    #strip characters not allowed in file names
    if subject != None:
        try:
            subject = "".join(x for x in subject if x.isalnum())
        except:
            subject = "Blank Subject"
    else:
        subject = "Blank subject"
    
    #limit file name to 200 characters
    
    if len(subject) > 200:
        subject = subject[0:100]
    
    print "Processing " + subject
      
    #write fields out to Excel Sheet
    
    worksheet.write(count,0,current_file)
    
    if email_p.get('From') != None:
        
        worksheet.write(count,1,remove_ascii_non_printable(email_p.get('From')))
    else:        
        worksheet.write(count,1,"Blank")
    
    if email_p.get('To') != None:             
        worksheet.write(count,2,remove_ascii_non_printable(email_p.get('To')))    
                
    else:        
        worksheet.write(count,2,"Blank")    
    
    if email_p.get('CC') != None:             
        worksheet.write(count,3,remove_ascii_non_printable(email_p.get('CC')))    
                
    else:        
        worksheet.write(count,3,"")  
    
    if email_p.get('BCC') != None:             
        worksheet.write(count,4,remove_ascii_non_printable(email_p.get('BCC')))    
                
    else:        
        worksheet.write(count,4,"")
    
    if email_p.get('Subject') != None:                
        subject = email_p.get('Subject')
          
        subject = remove_ascii_non_printable(subject)
        
        worksheet.write(count,5,subject)

    else:
        
        worksheet.write(count,5,"Blank")    
    
    date = ""
    if email_p.get('Date') != None:
        date = email_p.get('Date')
              
        #Date is in Long format, Fri, 14 Feb 2014 11:03:43 -0500 (EST)
        #convert to YYYY-MM-DD hh:mm:ss UTC so  it can be sorted
        
        timestamp = 0
        
        try:
            tt = parsedate(date)
        except:
            date = "Error on Conversion"
        
        try:
            tz = parsedate_tz(date)
        except:
            date = "Error on Conversion"
        try:
                timezone_offset = tz[9]
        except:
            date = "Error on Conversion"
 
        if "Error on Conversion" not in date:
 
            if timezone_offset == None:
                timezone_offset = 0

            timestamp = time.mktime(tt)
        
            #convert to UTC
            timestamp += timezone_offset
       
            #convert to formatted string
            date=(datetime.datetime.fromtimestamp(int(timestamp)).strftime('%Y-%m-%d %H:%M:%S'))
        
            #write out raw date and formatted date
            worksheet.write(count,6,remove_ascii_non_printable(email_p.get('Date')))
            worksheet.write(count,7,date,style0)
 
        else:
        
            worksheet.write(count,6,"Blank")
            worksheet.write(count,7,"Blank")
             
        
    #write out raw MIME to .eml file 
            
    #each file name will be named after the date, email subject plus a unique incremented number.
   
    if subject != None:
        try:
            #get rid of illegal file name characters:
            valid_chars = "_.() %s%s" % (string.ascii_letters, string.digits)
            valid_filename = ''.join(c for c in subject if c in valid_chars)
            #limit file name to 200 characters
 
        except:
            valid_filename = "Blank Subject"
    else:
        valid_filename = "Blank subject"
            
    if len(valid_filename) > 80:
        valid_filename = valid_filename[0:80]
    
    filename = date.replace(':',".") + "-"+ valid_filename +"_" + str(count) + '.eml'
    email_file = email_folder + filename
            
    isDeleted = False
    
    MessageID = email_p.get("Message-ID")
   
    if MessageID == None:
        worksheet.write(count,10,"Data not Available")
        worksheet.write(count,11,"Data not Available")
        worksheet.write(count,12,"Data not Available")
        worksheet.write(count,13,"Data not Available")
                
    else:
        #trim the <> characters
        MessageID = MessageID[1:-1]
        
        #sometimes there are spaces, so make sure they are gone
        MessageID = MessageID.replace('<', '')
        MessageID = MessageID.replace('>', '')
       
        worksheet.write(count,14,MessageID)
        
        #look for the globals sqlite database, this holds information like if the message was read, etc
        database = options.directory + seperator + "global-messages-db.sqlite"
        
        #if the database does not exist, we can't get the flags
        if os.path.isfile(database) == False:
                worksheet.write(count,10,"Data not Available")
                worksheet.write(count,11,"Data not Available")
                worksheet.write(count,12,"Data not Available")
                worksheet.write(count,13,"Data not Available")
        
        #if it does exist, connect and see if the flags are there
        else:    
            conn = sqlite3.connect(database)
           
            #check to see if MessageID is in database
            try:
                cursor = conn.execute("select * from messages where headerMessageID = '%s'" % MessageID)
                if cursor.fetchone():
                    isDeleted = False
                
                else:
                    #if the messages ID is not there, check the msf file. Junk email is not always in the db, but can be found in the corresponding msf file
                    msf_file = current_file + ".msf"
                    
                    if MessageID in open(msf_file).read():
                        isDeleted = False
                        
                    else:
                        #the msf (mork) file format is jacked up. It stores the message ID with slashes. strip the slashes and look for a match
                        #reset flag
                        isDeleted = True
                        with open(msf_file,"r") as f:                
                            for line in f:
                                strippedline = line.replace('\\', '')
                                
                                if MessageID[:60] in strippedline:
                                    isDeleted = False
                                    break
                                
                        if isDeleted == True:  
                            worksheet.write(count,11,"Deleted (Verify)")
                                     
                        
            except:
                worksheet.write(count,11,"Data not Available")
               
                #if there were issues connecting to db, then no use trying other queries, so set isDeleted flag to true
                isDeleted = True
               
           
            if read_id != None and isDeleted != True:
                try:
                    cursor = conn.execute("select jsonAttributes from messages where headerMessageID = '%s'" % MessageID)
                    jsonAttributes = cursor.fetchone()[0]
                    
                    if jsonAttributes != "":                   
                        ja = json.loads(jsonAttributes)
                        worksheet.write(count,10,ja[str(read_id)])
                    else:
                        worksheet.write(count,10,"Data not Available")
                except:
                                      
                    worksheet.write(count,10,"Data not Available")
            else:
                worksheet.write(count,10,"Data not Available")
                
    
            if repliedTo_id != None and isDeleted != True:
                try:
                    cursor = conn.execute("select jsonAttributes from messages where headerMessageID = '%s'" % MessageID)
                    jsonAttributes = cursor.fetchone()[0]
                    ja = json.loads(jsonAttributes)
                    worksheet.write(count,12,ja[str(repliedTo_id)])
                except:
                    worksheet.write(count,12,"Data not Available")
            else:
                    worksheet.write(count,12,"Data not Available")
            
            if forwarded_id != None and isDeleted != True:
                try:
                    cursor = conn.execute("select jsonAttributes from messages where headerMessageID = '%s'" % MessageID)
                    jsonAttributes = cursor.fetchone()[0]
                   
                    ja = json.loads(jsonAttributes)
                                    
                    worksheet.write(count,13,ja[str(forwarded_id)])
                except:
                    worksheet.write(count,13,"Data not Available")
            else:
                    worksheet.write(count,13,"Data not Available")

    counter = 0
    attachments = ""
        
    #find attachments in email, and if there are any, list them 
    for part in email_p.walk():
       
        c_disp = part.get("Content-Disposition")
        
        if c_disp != None:
            if counter == 0:            
                
                this_filename = part.get_filename()
                if this_filename != None:
                    attachments = this_filename
                counter += 1
            else:
                this_filename = part.get_filename()
                if this_filename != None:
                    attachments = attachments + "," + this_filename
    
    #write out attachment names to column
    
    worksheet.write(count,9,attachments)
    
    #write out the email
    outemail = open(email_file,"w")
    outemail.write(email)
    outemail.close()
           
    
    #write out filname
    click = "emails\\" + filename
    worksheet.write(count,8, xlwt.Formula('HYPERLINK("%s";"%s")' % (click,filename)),style)

    count = count + 1
    email = ''
    return email

######################### Main ###################################################

usage = "\n\nThis program parses a Thunderbird profile directory for the raw email files created by Thunderbird.\
 Header information (To, From, Date etc) are stored in an Excel file and a link is created to an exported\
 .eml file. Choose the directory containing the profile and an output directory to contain the report and email.\n\n\
NO TRAILING SLASHES\n\n\
Examples:\n\
eamil_parser.py -d /home/sansforensics/thunderbird_profile/9tdq9zg0.default -o /home/sansforensics/documents/parsed_emails"

from optparse import OptionParser

input_parser = OptionParser(usage=usage)

input_parser.add_option("-d", "--dir", dest = "directory", help = "process all files in directory and subdirectories", metavar = "/home/sansforensics/thunderbird_profile/9tdq9zg0.default")
input_parser.add_option("-o", "--output", dest = "output", help = "empty directory to output report and emails to", metavar = "/home/sansforensics/emails")

(options,args)=input_parser.parse_args()

#no arguments given by user,exit
if len(sys.argv) == 1:
    input_parser.print_help()
    exit(0)

#process all files directory        
if options.directory != None:
    
    #check to see if the directory exists, if not, silly user.. go find the right directory!
    if os.path.isdir(options.directory) == False:
        print ("Could not locate directory. Please check path and try again")
        exit (0)
    
    #crap, now we need to check to see if the path is a Windows or Linux path
    
    if '\\' in options.directory:
        seperator = "\\"
    if '/' in options.directory:
        seperator = "/"

    #create the output directory to hold the report and exported emails
    if not os.path.exists(options.output):
        os.makedirs(options.output)
    
    if not os.path.exists(options.output + seperator + "emails"):
        os.makedirs(options.output + seperator + "emails")
    
    email_folder = options.output + seperator + "emails" + seperator
    
    #create log file
    log_file = open(options.output + seperator + "log.txt", "w")
    
    print "Looking for global-messages-db.sqlite Database..."
    
    #look for the globals sqlite database, this holds information like if the message was read, etc
    database = options.directory + seperator + "global-messages-db.sqlite"
    
    if os.path.isfile(database) == True:
        print "Database Found"
        conn = sqlite3.connect(database)
             
        try:
            cursor = conn.execute("SELECT id from attributeDefinitions where name = 'read'")
            read_id = cursor.fetchone()[0]
        except:
            read_id = None
        
        try:
            cursor = conn.execute("SELECT id from attributeDefinitions where name = 'repliedTo'")
            repliedTo_id = cursor.fetchone()[0]
        except:
            repliedTo_id = None
            
        try:
            cursor = conn.execute("SELECT id from attributeDefinitions where name = 'forwarded'")
            forwarded_id = cursor.fetchone()[0]
        except:
            forwarded_id = None
    else:
        print "Database Not located"

    #loop through each file
    for subdir, dirs, files in os.walk(options.directory):
        for fname in files:
   
            current_file = os.path.join(subdir,fname)
            print current_file
            
        #try to open the files, if not there bail out
            try:
                f = open(current_file, "rb")
                                    
            except IOError as e:
                print 'File Not Found :' + current_file
                exit(0)
    
            #read "header" if not "From" not MIME format
            
            file_header = f.read(4)
            if str(file_header) != 'From':
                print fname + ' does not appear to be mailbox format'
                log_file.write ("NOT Processed: " + current_file + "\n\r")
                f.close()
            else:
                print "Processing " + fname
                log_file.write("Processed" + current_file + "\n\r")

                all_emails = current_file
                email = ""
                
                hexpattern = "From\x20\x0D"
                hexpattern2 = "From\x20\x0A"
            
                
                with open(all_emails,"r") as f:
                
                    for line in f:
                        hex1 = re.match(hexpattern,line)
                        hex2 = re.match(hexpattern,line)
                        #catches emails in Inbox, etc
                        if "From - " in line and email != "":
                                          
                            email = process_email(email)
                        
                        #catches emails in sent and draft folders - local emails composed from Thunderbird
                        #are formatted differently
                        
                        elif line == "From \n" and email != "":
                            email = process_email(email)
                           
                        elif line == "From" and email != "":    
                            email = process_email(email)
                        
                        elif hex1 and email != "":
                            email = process_email(email)
                        
                        elif hex2 and email != "":
                            email = process_email(email)
                        
                        else:
                            #build up the email body
                            email = email + line
                        
                #print last email in the que            
                email = process_email(email)
            
    #take away one from the count for excel sheet header
    print "Processed " + str(count-1) + " Emails"
    
    log_file.write("Processed " + str(count-1) + " Emails")

workbook.save(options.output + seperator + "report.xls")

log_file.close()
