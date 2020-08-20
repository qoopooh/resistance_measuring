# resistance_measuring

## Requirements
1. Press "Check" button to query data from Arduino
1. Record both "Pass" And "Fail" value
1. "Pass" shows green screen and "Fail" show red

## Packing
```sh

pyinstaller --onefile --noconfirm --noconsole --icon=favicon.ico --name resistance tester.pyw

```
mac
```sh

pyinstaller --onefile --noconfirm --noconsole --icon=mac.icns --name resistance tester.pyw

```
