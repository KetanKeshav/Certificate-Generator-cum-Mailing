# Automated Certificate Generator and Mailing using Python

Automation script for Generating Certificates and mailing them by using content from an excel file

# Dependencies needed:
pip install Pillow
pip install qrcode
pip install xlrd

# Change the path names for the below accordingly

1. selectFont = ImageFont.truetype(r'C:\Users\ketan\Desktop\emailcertificate\font.ttf', 65)

2. img = Image.open(r"C:\Users\ketan\Desktop\emailcertificate\certificate.jpg")

3. img.save(r"C:\Users\ketan\Desktop\emailcertificate\\"+fname+".pdf", "PDF", resolution=100.0)

#Run the program 

#Enter Path of Excel File (better if all the files are part of the same folder)

#Please Enter to Exit
