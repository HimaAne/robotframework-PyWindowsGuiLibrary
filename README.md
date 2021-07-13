# PyWindowsGui Library
This library is created to automate testing the WINDOW GUI (desktop) application using the Robot Framework. 
It uses pywinauto and pyautogui wrappers to interact with GUI interface.

## Installation
PyWindowGuiLibrary can be found on PyPI: https://pypi.org/project/robotframework-PyWindowsGuiLibrary.

To install, simply use pip:

```dos
pip install robotframework-PyWindowGuiLibrary
```

Dependencies are automatically installed.

## Importing in Robot Framework
As soon as installation has succeeded, you can import the library in Robot Framework:

```robot
*** Settings ***
Library  PyWindowsGuiLibrary
```

## Usage

First of all make sure application is up and running.

To open a new instance of an application from given path use Launch Application or Focus Application Window to connect to a instance that is already opened.

This library provides support even for opened applications. To connect an existing application use Focus Application Window without using Launch Application.

## Finding properties
Initially to know about the application tittle property use AutoIT inspector.

To get AutoIT inspector follow this link and use: https://www.autoitscript.com/cgi-bin/getfile.pl?autoit3/autoit-v3-setup.exe

To get the Controls and properties use Print Current Window Page Object Properties and Print Specific Object Properties On Current Window keywords.

To find exact element property/control, analyse the hierarchy and find from printed controls identifiers.

## Keyword documentation
For the keyword documentation [go here](https://himaane.github.io/robotframework-PyWindowsGuiLibrary/PyWindowsGuiLibrary.html).
