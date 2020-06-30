# Server/WebConsole Health Check

This is just a demonstration on how we can manuipulate selenium to our will to automate anything on Web Browser.

## Cause:

Main motive or reason behind the script is to continously/at specific intervals monitor Perimeter Servers/Adapters/ConsoleStatus via web interface.

## Function:

This script just shows how versatile we can configure selenium and it also goes deep on fundamental/different features available in Selenium Library using Python.

## Modules/Packages used:

These are the modules used in the script -

* Selenium
* OS
* PIL
* io
* pytesseract
* win32com.client
* logging

`Selenium` module is core of this script since this controls whole webpage.

`OS` module is used to act as interface with underlying operating system depending on the OS in user's machine.

`PIL` library is used to manipulate and save the image which is being captured by pytesseract module.

`io` module provides functionality for dealing with various types of I/O (input output operations).

`pytesseract` it's a optical character recognition tool for python which is used to recognize and read the text embedded in the images.

`win32com.client` used to provide access to outlook/emails which is used as medium to convey the status report.

`logging` standard livrary modile built in python to provide access to create teh logs while executing the scripts.

# License

Copyright (C) 2020 Kailash Prabhu

This is made public just to show on versatile Selenium and its vast functionality. It's currently being used in our organization to generate Health check status reports.

Happy to accept any pull and recode it based on your organization the main reason why this is made public.