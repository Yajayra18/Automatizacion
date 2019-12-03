#!/usr/bin/python3
import json
import re
import smtplib
import subprocess
import sys

list_package = []

try:
    import getpass#
except ImportError:
    print("WARNING: The package getpass is not installed.")
    list_package.append("getpass")    
try:
    import pandas as pd#
except ImportError:
    print("WARNING: The package pandas is not installed.")
    list_package.append("pandas")

if len(list_package):
    flag = input("Do you want install the following packages\n\t-"+"\n\t-".join(list_package) +"\n( [y] / n ):")
    if flag == "" or flag.upper() == "Y":
        if sys.platform == 'win32':
            subprocess.run("pip install " + " ".join(list_package) , stdout=subprocess.PIPE, universal_newlines=True)        
        elif sys.platform == 'linux':
            subprocess.run("sudo pip3 install " + " ".join(list_package), stdout=subprocess.PIPE, universal_newlines=True)
        else :
            print("WARNING: For the %s platform you must install manualy the packages."%(sys.platform))
            sys.exit(1)
    elif flag.upper() == "N":
        sys.exit(1)
    else:
        sys.exit(1)

import src.claseMensaje as msg

# Regular Expression
email_patron = re.compile(r'\b[\w.%+-]+@[\w.-]+\.[a-zA-Z]{2,6}\b')
url_patron = re.compile(r"^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$")
html_patron = re.compile(r"\b[\w.%+-]+\.html")

# Import Configuration from json
with open(r"config.json", "r") as config_file:
    config = json.load(config_file)

# SMTP Connection
with smtplib.SMTP(config["host"][config["host_opc"]], config["port"]) as server:
    server.starttls()

    print("***** Login *****")
    # Username
    while True:
        temp_email = input("Login ( [%s] ): "%config["email"])
        if temp_email != "":
            if email_patron.match(temp_email):
                config["email"] = temp_email
                break
            else:
                print("WARNING: write an able email.")
        else:
            break

    # Password
    pass_personal = getpass.getpass()
    if pass_personal == "":
        print("WARNING: please write your password.")
        sys.exit(1)

    # Login 
    try:
        server.login(config["email"], pass_personal)
    except smtplib.SMTPAuthenticationError:
        print("WARNING: Could not login to the smtp server please check your username and password.")
        sys.exit(1)
    else:
        print("Login successfully.")
    
    
    print("\n***** Email *****")
    # Message class
    new_message = msg.Mensaje(config["email"])

    # To
    print("Importing from excel's file the emails.")
    data = pd.read_excel(r"excel/correos.xlsm", skiprows=3)
    df = pd.DataFrame(data)

    if "Nombre" in df.columns:
        for i in df.index:
            new_message.add_email(df["Correo"][i], df["Nombre"][i])
    else:
        for i in df.index:
            new_message.add_email(df["Correo"][i])
    
    # Subject
    temp_subject = input("Subject: ")
    if temp_subject == "":
        temp = input(r"Are you sure that you want an empty subject? ( [y]/n ):")
        if temp.upper() == "N":
            temp_subject = input("Subject: ")
    new_message.add_subject(temp_subject)

    # HTML
    while True:
        temp_html = input("HTML ( [%s] ): "%config["html"])
        if temp_html != "":
            if html_patron.match(temp_html):
                config["html"] = temp_html
                break
            else:
                print("WARNING: write an able html file.")
        else:
            break
    new_message.add_message_html(config["html"])

    # opciones: oculto, personalizado, grupa
    print("\n***** Sending (in process) *****")
    if new_message.To_name:
        for name, email in zip(new_message. To_name,new_message.To):
            server.sendmail(new_message.From, [email], new_message.get_message(name,email).encode("utf8"))
            print("Email sent successfully to " + email)
    else:
        for email in new_message.To:
            server.sendmail(new_message.From, [email], new_message.get_message(email=email).encode("utf8"))
            print("Email sent successfully to " + email)

# Save Configuration into json
with open(r"config.json", "w") as outfile:
    print("\nSave configuration ... in process")
    json.dump(config, outfile, sort_keys=True, indent=4)
    print("Save configuration ... OK\n")

