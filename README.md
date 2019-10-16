# hw2serial
Sending hardware sensors data to serial port (for Arduino and other external devices with serial)

This program is able to send custom hardware information to serial port using OpenHardwareMonitor library, e.g. it can be used to create PC Monitor based on Arduino.

To create .exe file you should install Pyinstaller and run the following code:
pyinstaller.exe --onefile --icon=hw2serial.ico --noconsole --add-data="hw2serial.ico;." hw2serial.py
