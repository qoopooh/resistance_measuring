# resistance_measuring
This project is product tester software that work with jig tester (Hardware + Arduino).

## Technician setup instruction
Before use this software in production line, technician will setup the line by connecting jig tester to PC.
1. Open the software
1. Select COM port of jig tester in dropdown list (Combobox) e.g. COM1, COM2, COM3,..
1. If there is no COM port in dropdown list, please reconnect the USB cable then press "Refresh" button


## Work instruction for staff (in production line)
1. Press "Check" button to query data from Arduino
1. Record both "Pass" And "Fail" value
1. "Pass" shows green screen and "Fail" show red


## Software deployment
This software will be setup on Windows machine. We will deploy the software as .exe file.

### Create execution file command
Windows 10
```sh

pyinstaller --onefile --noconfirm --noconsole --icon=favicon.ico --name resistance tester.pyw

```
Mac OS
```sh

pyinstaller --onefile --noconfirm --noconsole --icon=mac.icns --name resistance tester.pyw

```
