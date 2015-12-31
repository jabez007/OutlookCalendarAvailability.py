## Synopsis

This was a quick project I put together to easily export what availability I have over any particular date range

## Code Example

Select the start and end dates for the date range you want to export your availability for, then select the time zone
you want to export your availability for.

![example workflow](https://cloud.githubusercontent.com/assets/14042259/12064065/29e2ab3c-af7f-11e5-8181-bed662f5b1b1.png)

## Motivation

I found myself spending entirely too much time manually going through my Outlook calendar trying to find good times to
schedule meetings with people who don't have access to my Outlook calendar. 
This utility is primarily meant has a time saver

## Installation

Only tested on python 2.7
Download all the .py files from this repository to the same local directory
Install the python packages found in requirements.txt

## Known Bugs

1. If a date included in your range has no availability, the script will crash
2. If your Outlook is set to something other than Central time, the timezone switch wont work properly

## Contributors

* Credit for ttkcalendar.py goes to [evandrix](https://github.com/evandrix/cPython-2.7.3/blob/master/Demo/tkinter/ttk/ttkcalendar.py)

## License

The MIT License (MIT)

Copyright (c) 2015 Jimmy McCann

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.