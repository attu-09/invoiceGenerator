from flask import Flask, render_template, url_for ,redirect, request
import xlsxwriter
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app=Flask(__name__)
app.config['SECRET_KEY']="asdadvadfsdfs"


#--------------------- Utility Function -----------------------------------------------------------
def sendMail(recieverEmail):
    subject = "Vendor Innvoice"
    body = "Automated generated innvoice"
    sender_email = "ace.webinnovation@gmail.com"
    receiver_email = recieverEmail
    password = "TheAce's1"

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    # message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = "Invoice.xlsx"  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= Invoice.xlsx",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    with smtplib.SMTP("smtp.gmail.com",587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)

def excelMaker(data1,data2):
    # predefined columns
    first = ['Invoice No','Invoice Date','Vendor Code','Invoice Amount','Tax Amount','P.O. No.']
    second = ['P.O. Line Item No.','QTY','Value','GST']
    
    # opening excel first, if not found then create it.
    workbook = xlsxwriter.Workbook('Invoice.xlsx')
    worksheet = workbook.add_worksheet()
    
    # first
    for col,item in enumerate(first):
        worksheet.write(0,col,item)
    
    # row_counter indicates on which column we are in.
    row_counter = 1
    
    # inserting row in first_part
    for col,value in enumerate(data1):
        worksheet.write(row_counter,col,value)
    row_counter += 1
        
    # second_part column 
    row_counter+=1
    for col,item, in enumerate(second):
        worksheet.write(row_counter,col,item)
    row_counter+=1
        
    # inserting row in second_part
    for row_value in data2:
        for col,value in enumerate(row_value):
            worksheet.write(row_counter,col,value)
        row_counter+=1
        
    workbook.close()
#--------------------------------------------------------------------------------------------------


@app.route('/',methods=["Get","POST"])
def home():
    msg=''
    if request.method=="POST":
        data1= (request.form['invoiceNo'],request.form['invoiceDate'],request.form['venderCode'],request.form['invoiceAmount'],request.form['taxAmount'],request.form['poNo.'])
        data2=[]
        poLine,quant,value,gst=request.form.getlist('poLineItem'),request.form.getlist('quantity'),request.form.getlist('value'),request.form.getlist('gst')
        valueTemp=gstTemp=0
        for i in range(len(poLine)):
            v,g=int(value[i]),int(gst[i])
            valueTemp+=v
            gstTemp+=g
            data2.append((poLine[i],quant[i],v,g))
        if int(data1[3])==valueTemp and int(data1[4])==gstTemp:
            excelMaker(data1,data2)
            try:                
                sendMail(request.form['recieverEmail'].strip())
            except:
                msg="Invalid email"
                return render_template('form.html',msg=msg)
            msg="Mail sent"
        else:
            msg="Amount dosen't match"
    return render_template('form.html',msg=msg)

if __name__=="__main__":
    app.run(debug=True)