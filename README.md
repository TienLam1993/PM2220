# PM2220
### Video demo: (https://youtu.be/HAgHDzlEVsY)
### Description: 

A Windows desktop program to display to real time value of  electric Voltage, Current from the Schneider power meter PM2220 using Modbus protocol

#### Hardware requirments
1. Schneider PM2220 power meter.
2. A serial to usb module (if your desktop/laptop does not have serial port).

For wiring the PM2220 power meter, check out the [datasheet](https://download.schneider-electric.com/files?p_enDocType=User+guide&p_File_Name=NHA2778902-08-EN.pdf&p_Doc_Ref=NHA2778902-01)
from Schneider.

About Modbus protocol (https://modbus.org/docs/PI_MBUS_300.pdf).

PM2220 Modbus register address (https://www.se.com/ng/en/faqs/FA410489/).

**You must be a professional/certificated electrical technician to perform wiring the power meter**.

#### Software

This program is a Winform application written in C sharp in Visual studio.

I use Easy Mobus library.

Because the power meter only support modbus RTU, and my laptop does not have a serial port so I have to use an RS485-USB converter.
Another problem is you need to monitor the data in real time from serial port. So you have to do a backgroud worker and Invoke function.
You can find more about background worker class in Microsoft docs [here](https://docs.microsoft.com/en-us/dotnet/api/system.componentmodel.backgroundworker?view=net-5.0).

To visualize the data, I use the ZedGraph library. You can read more about it [here](https://sourceforge.net/projects/zedgraph/).

The application also records the data to an Exel workbook every 10 seconds using Microsoft Office Excel interop.

The application use the .NET framework 4.8.
