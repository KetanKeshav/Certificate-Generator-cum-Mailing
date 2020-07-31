from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd 
import qrcode
import time
import smtplib
from email.message import EmailMessage

def main():
    server=smtplib.SMTP_SSL("smtp.gmail.com",465)
    server.login("your-email-id","your-password")
    
    path=input("Enter Path of Excel File\n")
    wb = xlrd.open_workbook(path) 
    sheet = wb.sheet_by_index(0)
    R=sheet.nrows
    C=sheet.ncols
    basewidth = 80
    hsize=80;
    selectFont = ImageFont.truetype(r'C:\Users\ketan\Desktop\emailcertificate\font.ttf', 65)
    for i in range(R): 
        fname=sheet.cell_value(i,0)
        fullname=fname
        mail=sheet.cell_value(i,1)

        #declaring no. of white spaces
        max=36
        half=max//2
        namel=len(fullname)
        if namel<max: 
            w=half-(namel//2)
            ws=" "*(int(w*2))
            fullname=ws+fullname

        #take a image
        img = Image.open(r"C:\Users\ketan\Desktop\emailcertificate\certificate.jpg")
        draw = ImageDraw.Draw(img)
        draw.text( (200,600),fullname, (0,0,0),font=selectFont)
        img.save(r"C:\Users\ketan\Desktop\emailcertificate\\"+fname+".pdf", "PDF", resolution=100.0)

        msg=EmailMessage()

        msg['Subject']="Thank you for attending the workshop on \"How to Choose Right Journal\""
        msg['From']='ketan.keshav79@gmail.com'
        msg['To']=mail
        body="""
Dear Sir/Ma'am,

PFA the certificate for the workshop on "How to Choose Right Journal?" conducted on 14 July 2020. 

Thanks
Ketan
"""

        msg.set_content(body)
        with open(fname+".pdf",'rb') as f:
            file_data=f.read()
            file_name=f.name
        msg.add_attachment( file_data , maintype='application' , subtype='octet-stream' , filename=file_name )

        server.send_message(msg)

    server.quit()

if __name__ == "__main__":
    main()
    input("Please Enter to Exit")