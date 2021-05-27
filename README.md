# PM2220
### Video demo: (https://youtu.be/HAgHDzlEVsY)
### Description: 

A Windows desktop program to display to real time value of  electric Voltage, Current from the Schneider power meter PM2220 using Modbus protocol

#### Hardware requirments
1. Schneider PM2220 power meter.
2. A serial to usb module (if your desktop/laptop does not have serial ports).

Modbus is an open protocol derived from the Master/Slave architecture originally developed by Modicon (now Schneider Electric). 
It is a widely accepted in Industrial Automation System due to its ease of use and reliability.
There are 3 main types of Modbus: Modbus TCP/IP, Modbus ASCII and Modbus RTU.
Modbus RTU (Remote Terminal Unit) — This is used in serial communication and makes use of a compact, binary representation of the data for protocol communication.
The RTU format follows the commands/data with a cyclic redundancy check checksum as an error check mechanism to ensure the reliability of data. 
Modbus RTU is the most common implementation available for Modbus.
A Modbus RTU message must be transmitted continuously without inter-character hesitations. Modbus messages are framed (separated) by idle (silent) periods.


Because PM2220 only support Modbus RTU, and my laptop does not have any Serial port, so I have to use a Serial to USB converter module.


For wiring the PM2220 power meter, check out the [User Manual](https://download.schneider-electric.com/files?p_enDocType=User+guide&p_File_Name=NHA2778902-08-EN.pdf&p_Doc_Ref=NHA2778902-01)
from Schneider.

About Modbus protocol (https://modbus.org/docs/PI_MBUS_300.pdf).

PM2220 Modbus register address (https://www.se.com/ng/en/faqs/FA410489/).

**You must be a professional/certificated electrical technician to perform wiring the power meter**.

#### Software

This program is a Winform application written in C sharp in Visual studio.

At the beginning of my project, I had to choose between Python and C sharp, but after doing some research, I decided to use C sharp because the Modbus library
on C sharp is more easy to use and more stable(according to some answers on stackoverflow). Besides this, I want to build a Windows application so C# seem to be a nature choice.

I use  the Easy Mobus .NET library. It supports all the functions of the Modbus RTU protocol.

- Read Coils (FC1)
- Read Discrete Inputs (FC2)
- Read Holding Registers (FC3)
- Read Input Registers (FC4)
- Write Single Coil (FC5)
- Write Single Register (FC6)
- Write Multiple Coils (FC15)
- Write Multiple Registers (FC16)

You can find more information about the library [here](http://easymodbustcp.net/en/modbusclient-methods).

When I created a Winform application, Visual studio automactically built the project structure for me. So I only have to focus on coding.
The main program is done of Form1.cs

A big problem I faced is the need of monitoring the data in real time.
This is the time I started to learn about multi-threaded programming. 
We often use background threads when a time-consuming process needed to be executed in the background without affecting the responsiveness of the user interface. 
This is where a BackgroundWorker component comes into play.

You can find more about background worker class in Microsoft docs [here](https://docs.microsoft.com/en-us/dotnet/api/system.componentmodel.backgroundworker?view=net-5.0).

To visualize the data, I use the ZedGraph library. You can read more about it [here](https://sourceforge.net/projects/zedgraph/).

The application also records the data to an Exel workbook every 10 seconds using Microsoft Office Excel interop.

The application use the .NET framework 4.8.
