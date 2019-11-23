# Clase mensaje
class Mensaje:
    # Constructor
    def __init__(self, From):
        self.From = From
        self.To_name = []
        self.To = []
        self.Subject = ""
        self.MessageHTML = ""
        print("New Message was created.")
    
    def add_email(self, email, name=None):
        if name: 
            self.To_name.append(name)
        self.To.append(email)

    def add_subject(self, subject):
        self.Subject = subject
    
    def add_message_html(self,html_file):
        with open(r"html/" + html_file, "r", encoding='utf-8') as html:
            self.MessageHTML = html.read()

    def get_message(self, name=None, email=None):
        temp = "Subject: " + self.Subject + "\n"
        temp += "From: <" + self.From + ">\n"

        if name and email:
            temp += "To: " + name + " <" + email + ">\n"
        elif name==None and email:
            temp += "To: " + email + "\n"
        else:
            temp += "To: \n"
        
        temp += "Content-Type: text/html; charset=UTF-8\n\n"
        temp += "%s"%self.MessageHTML
        
        return temp