# PM2220
### Video demo: (https://youtu.be/HAgHDzlEVsY)
### Description: A Windows desktop program to display to real time value of  electric Voltage, Current from the Schneider power meter PM2220 using Modbus protocol

### Hardware requirments
1. Schneider PM2220 power meter.
2. A serial to usb module (if your desktop/laptop does not have serial port).

For wiring the PM2220 power meter, check out the [datasheet](https://download.schneider-electric.com/files?p_enDocType=User+guide&p_File_Name=NHA2778902-08-EN.pdf&p_Doc_Ref=NHA2778902-01)
from Schneider.

PM2220 Modbus register address (https://www.se.com/ng/en/faqs/FA410489/).

**You must be a professional/certificated electrical technician to perform wiring the power meter**

### Software

This program is written in C sharp in Visual studio.

It uses Easy Mobus library to connect, and ZedGrap library to visualize the data from the power meter.

It also record the data to an Exel workbook every 10 seconds.

To install program to your PC, you need to install the .NET framework 4.7.
