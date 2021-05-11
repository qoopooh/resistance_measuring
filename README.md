# resistance_measuring software
This project is product tester software that work with jig tester (Hardware + Arduino).


## Business requirements
1. Staff can use this software to validate product in production line
1. The validation result will show "Pass" with green screen and "Fail" is red
1. Software will record both of "Pass" And "Fail" value, and then can export data later


## Technician setup instruction
Before use this software in production line, technician will setup the line by connecting jig tester to PC.
1. Open the software
1. Select COM port of jig tester in dropdown list (Combobox) e.g. COM1, COM2, COM3,..
1. If there is no COM port in dropdown list, please reconnect the USB cable then press "Refresh" button
1. Uncheck "Lock" checkbox to edit Lot number
1. Edit Lot number
1. Check "Lock" checkbox


## Work instruction for staff (in production line)
1. Connect DUT to jig tester 
1. Press "Check" button to test DUT
1. If software show green screen with "PASS" text means DUT is OK
1. If the result show red screen with "FAIL" text means DUT is not good and put it in scrap box


# Software deployment
This software will be setup on Windows machine. We will deploy the software as .exe file.

## Setup software environment for deployment
```sh
pip install -r requirements.txt
```

## Create execution file command
Windows 10
```sh

pyinstaller --onefile --noconfirm --noconsole --icon=favicon.ico --name resistance tester.pyw

```
Mac OS (for testing)
```sh

pyinstaller --onefile --noconfirm --noconsole --icon=mac.icns --name resistance tester.pyw

```
