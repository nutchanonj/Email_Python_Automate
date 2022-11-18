import smtplib
from email.message import EmailMessage

EMAIL_ADDRESS = 'example@gmail.com' # Your email.
EMAIL_PASSWORD = 'password' # Your email's password.

import mimetypes

from csv import reader

if __name__ == '__main__':

    # open file in read mode
    with open('Job_Fair_Email_Info.csv', 'r', encoding='utf-8') as read_obj:
    # pass the file object to reader() to get the reader object
        csv_reader = reader(read_obj)
        header = next(csv_reader)
        # Iterate over each row in the csv using reader object
        n = 1
        
        for row in csv_reader:
            
            msg = EmailMessage()
            msg['Subject'] = 'Announcing your organization\'s Zoom meeting URL and appointment time for the faculty\'s Job Fair 2022'
            msg['From'] = EMAIL_ADDRESS 
            msg['To'] = row[1].split()
            msg.set_content("""
                
            To the HR of %s
            \n
            Regarding the faculty\'s Job Fair, Your organization\'s time slot is allocated for you as follow:
            %s January 2022, %s, at Zoom meeting room number %s
            \n
            The associated Zoom URL is provided below:
            %s
            Meeting ID: %s
            Password: %s
            Please standby in this online meeting 15 minutes prior to the appointed time. 
            \n
            For emergency case, if we cannot contact you in the Zoom meeting,
            we will contact you back through your phone, using this number: %s
            or your LINE account: %s
            \n
            We also provide you the briefing as the pdf file in this email\'s attachment.
            \n
            We are very appreciated for your participation in our  2022 Job Fair. Thank you very much!
            See you!
            \n
            If you have any question, please contact 083-XXX-XXXX or 086-XXX-XXXX.
                            
            """ % (row[0], row[4], row[5], row[6], row[7], row[8], row[9], row[2], row[3]))
            
            
            with open('Briefing.pdf', 'rb') as content_file:
                content = content_file.read()
                msg.add_attachment(content, maintype='application', subtype='pdf', filename='Briefing.pdf')
                
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD) 
                smtp.send_message(msg)
                print("Success", n)
                n += 1
            