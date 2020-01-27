import config
import xlrd
import smtplib #module used for handling email related tasks in python
def sendmail(msg,name,password):
        server=smtplib.SMTP('smtp.gmail.com:587')#creates an smtp object that connects to smtp server og gmail and 587 port no
        server.ehlo()#for setting up contact with server (just like a hello greet message , if after printing  250 is the first code its the code for success
        server.starttls()#enables trans layer encryption for our program 
        server.login(config.email_address,password)#used for logging in to the senders email 
        server.sendmail(config.email_address,email,"Subject: CBT RESULTS\n\n "+msg)#we need to seperate subject from body using newline twice 
        server.quit()
        print("email sent!")
        
##excel files

file_location="C://Users//Akhil Kumar//Desktop//book2.xlsx"
workbook=xlrd.open_workbook(file_location)
sheet=workbook.sheet_by_name('Sheet1') #we can even do sheet_by_index(0) will give us first sheet 
#print(sheet.cell_value(0, 0))
for row in range(1,sheet.nrows):
        email=[]
        email.append((sheet.cell_value(row,1)))
        name=sheet.cell_value(row,0)
        p_marks=sheet.cell_value(row, 2)
        c_marks=sheet.cell_value(row, 3)
        m_marks=sheet.cell_value(row, 4)
        cm_marks=sheet.cell_value(row, 5)
        b_marks=sheet.cell_value(row, 6)
        #marks=[p_marks,c_marks,m_marks]
        msg='Hi ' +name+' your marks are:\nPhysics: '+str(p_marks)+"\nChemistry: "+str(c_marks)+"\nMaths: "+str(m_marks)+"\nComputers: "+str(cm_marks)+"\nBiology: "+str(b_marks)
        if(sum([p_marks,c_marks,m_marks,cm_marks,b_marks])<=250):
                msg=msg+"\n\nYou have failed. Better luck next time!"
        else:
                msg=msg+"\n\nYou have managed to pass. Congratulations and keep it up!"
        #print(msg)
        password=config.password
        sendmail(msg,name,password)
        




