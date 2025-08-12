# Serial_Terminal
This project is a serial terminal application developed in Python with a Tkinter GUI. It allows users to communicate with serial devices by sending and receiving data in real-time.

Overview:

This is a serial terminal application developed in Python with a Tkinter graphical user interface. It allows users to communicate with serial devices via serial ports, send commands, and view responses in real-time. The program features:

<img width="828" height="536" alt="image" src="https://github.com/user-attachments/assets/887cf519-e9bd-4201-a8a0-1259a0f7462c" />

Listing and selecting available serial ports

Configurable baud rate and connection parameters

Sending commands and receiving data asynchronously

Timestamped output display with scrollable text area

Logging communication to rotating log files

Save 10 commands in json file for multiple used commands

send batch commands file

Routing:

Routing USB COM Port with hub4com and com2tcp to another comports or another host by telnet for monitoring with multiple points

What are hub4com and com2tcp?

com0com is open source software yo can find

http://sourceforge.net/projects/com0com/

The hub4com is a Windows application and is a part of the com0com project.
The hub4com can be used for sharing, joining, redirecting, multiplexing and encrypting serial and TCP/IP port data.
It allows to receive data and signals from one port, modify and send it to a number of other ports and vice versa. com2tcp is a utility that redirects a COM port over TCP to another host.

Setup for using 2 application in one PC windows:

![Your paragraph text](https://github.com/user-attachments/assets/5e25b56a-d7e3-4585-b3ea-5bae11b1ca28)

1. Install Com0Com
2. Open setup configure same as picture

<img width="447" height="393" alt="Screenshot 2025-08-12 135127" src="https://github.com/user-attachments/assets/58ecd126-d8fa-479c-aeec-56d811e8711d" />

3. On Serial Terminal choose Source Com1 and Dest 1 Com20 and Dest 2 Com30
4. press Start Routing
5. now you can use COM2 on app1 and on COM3 app2

Setup for using 2 application in two PC windows in same network:

![Your paragraph text (1)](https://github.com/user-attachments/assets/5ba9a0bc-e683-4110-9428-06c3b64da24a)


1. Install Com0Com
2. Open setup configure same as picture

<img width="447" height="393" alt="Screenshot 2025-08-12 141217" src="https://github.com/user-attachments/assets/720fe1eb-fe91-45f7-8636-197bdc51efff" />


3. On Serial Terminal choose Source Com1 and Dest 1 Com20 and Dest 2 Com30 and Dest 3 COM40
4. press Start Routing
5. PC1 consider as a server and PC2 as a client 
6.  now you can use COM2 on app1 and on COM3 app2 in PC1 and use any telnet software using PC1 IP (like Putty) and port 23 to get serial data routed from PC1

🙋‍♂️ Author: Amin Khalafi

amin_khalafi@yahoo.com

Reach out via GitHub or email if you'd like to collaborate!
