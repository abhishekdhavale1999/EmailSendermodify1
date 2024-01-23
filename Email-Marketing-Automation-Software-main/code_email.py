import smtplib  


def Email_send_function(to, subject, message, uname, pasw):
    s = smtplib.SMTP("smtp.gmail.com", 587) 
    s.starttls()  
    s.login(uname, pasw)
    msg = "Subject: {}\n\n{}".format(subject, message)
    s.sendmail(uname, to, msg)
    x = s.ehlo()
    if x[0] == 250:
        return "s"
    else:
        return "f"
    s.close()
