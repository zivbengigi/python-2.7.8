import socket
import os
import inspect
import shutil

class Server:
    gate = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    # This allows the address/port to be reused immediately
    # instead of it being stuck, waiting for late packets to arrive.
    gate.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    port = 0
    count = 0

    def __init__(self, port):
        self.port = port
        self.count = 0
        self.gate.bind(('', self.port))

        self.currentDir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))) + "\\output"

        if os.path.exists(self.currentDir):
            shutil.rmtree(self.currentDir)
        os.mkdir(self.currentDir)

    def listen(self):
        self.gate.listen(10)
        while True:
            conn,address = self.gate.accept()
            filename = self.getName(conn)
            self.receiveFile(conn, filename)

    def getName(self, sock):
        data = sock.recv(1024)
        sock.send("OK")
        return data
    
    def receiveFile(self, sock, filename):
        fileData= open(os.path.join(self.currentDir, filename), 'wb')
        while True:
            data = sock.recv(1024)
            # Exit if disconnected
            if not data: break
            # End of file
            if data[-4:] == "END!":
                fileData.write(data[:-4])
                fileData.close()
                break
            else:
                fileData.write(data)
        sock.send("ACK")

def Main():
    s = Server(9966)
    s.listen()

if __name__ == "__main__":
    Main()
