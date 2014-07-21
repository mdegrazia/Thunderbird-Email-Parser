Thunderbird-email-parser
========================

This script parses the raw emails, including deleted emails in files created by Thunderbird
located under \Users\%USERNAME%\AppData\Roaming\Thunderbird\Profiles\[random].default\

These files do not have a file extension and contain emails in MIME format.
It will parse the Header information(To, from, CC, BC, Date and Subject) into an Excel file
and create a link to the .eml file. It will also list the attachments.


####Required Library 
  Install the xlwt library on Linux/OS X using:
  
    sudo easy_install xlwt
    
  Windows:
  
  Use the xlwt installer located at https://pypi.python.org/pypi/xlwt/0.7.2
      

####Usage

    thunderbird_parser.py -d /home/sansforensics/thunderbird_profile/9tdq9zg0.default -o /home/sansforensics/documents/parsed_emails

####More Information

View the blog post at http://az4n6.blogspot.com/2014/04/whats-word-thunderbird-parser-that-is.html


Email Mari > arizona4n6 at gmail dot com for help/questions/bugs

