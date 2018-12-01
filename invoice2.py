import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont
from string import Template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('invoice.json',scope)
client = gspread.authorize(creds)

sh = client.open ('invoice data')
sheet = sh.get_worksheet(0)

def variables():
    name = sheet.cell(row_count, 3).value
    email = sheet.cell(row_count, 4).value
    event = sheet.cell(row_count, 5).value
    amount = str(sheet.cell(row_count, 6).value)
    date = sheet.cell(row_count, 1).value
    reg_num = str(row_count - 1)
    sheet.update_cell(row_count,7,reg_num)

    return name, email, event, amount, date, reg_num

def invoice(name,event,amount,date,reg_num):
    image = Image.open('invoice2.png')
    draw = ImageDraw.Draw(image)
    date_font = ImageFont.truetype('arial.ttf', size=12)
    reg_font = ImageFont.truetype('arial.ttf', size=18)
    name_font = ImageFont.truetype('arial.ttf', size=22)
    colour = 'rgb(0,0,0)'
    
    draw.text((350,30), date, fill=colour, font=date_font)
    draw.text((430,90), '#'+reg_num, fill=colour, font=reg_font)
    draw.text((125,182), name, fill=colour, font=name_font)
    draw.text((125,241), event, fill=colour, font=name_font)
    draw.text((143,300), amount+" ₹", fill=colour, font=name_font)
    
    image.save("#"+reg_num+'.png')

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def send_mail():
                 #sending mails
                 message_template = read_template('greeting_template.txt')
                 # set up the SMTP server
                 s = smtplib.SMTP(host='smtp.gmail.com', port=587)
                 s.starttls()
                 s.login("your email" , "password")

                 msg = MIMEMultipart()
                 message = message_template.substitute(name=name.title())
                 msg['From']="your email"
                 msg['To']=email
                 msg['Subject']="Invoice"
                 msg.attach(MIMEText(message, 'plain'))
                 filename = reg_num+".png"
                 attachment = open("#"+reg_num+".png", "rb")
                 p = MIMEBase('application', 'octet-stream')
                 p.set_payload((attachment).read())
                 encoders.encode_base64(p)
                 p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                 msg.attach(p)
                 s.send_message(msg)
                 del msg
                 s.quit()
                 sheet.update_cell(row_count,8,"sent")

if __name__=="__main__":
     while (True):
         cell = sheet.find('')
         row_count = cell.row
         name, email, event, amount, date, reg_num = variables()
         invoice(name,event,amount, date,reg_num)
         status = str(sheet.cell(row_count,8).value)
         print('done')
         if (status != "sent"):
             send_mail()



         
          
         
         
    
    
    
    
