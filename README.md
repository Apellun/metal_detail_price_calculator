# A metal detail making cost calculator

<b>Это английская версия презентации, выберите ветку presentation_rus, чтобы посмотреть версию на русском.</b>

Technologies: Pandas, PySimpleGui

## How to launch:

Download the repository and run the main.py file.
(soon) Or download the archive with an executable for your system, unpack it and run the executable.

## What it does

The main function of the app is to automate the calculation of the price for a metal detail. Previously, the managers of the company had to manually look through tables and calculate the prices.

The main window allows the user to input all detail data. When the user can call up for the demonstration of the prices and other detail info, save the detail data to the spreadsheet with previous accounts, or save a separate spreadsheet with the current calculation only for print.

Also, the user can change the file the application takes the data from (it has to be an excel-file with the structure programmed in the app), the file where it writes the data of the calculations, and turn those back to the default settings.

## Difficulties I encountered while building it

### Having to work with spreadsheets

The application is built to work with excel spreadsheets, since it was most convenient for the client an there were no time and money to create a database for it. I may add the database in the future, to avoid all the complications that come with the app being dependend on external files.

For now, I tried my best to create workarounds for any possible malfunction that could happen because of it. For example, the application checks at the startup if the spreadsheet it needs for work (with the prices of metals, cutting and inset) is in place and if it's readable. If it's not, itasks the user to manually select the spreadsheet to get data from.

### Python GUI

I understand that Python is not the most popular language for building desktop apps but I did not have the time to learn another language :\) Actually, having to find out about existing GUI frameworks was interesting. I started with Tkinter and built some of the interface with it, but at some point it became cumbersome and hard to read, so I switched to PySimpleGUI. There was another complication with that, the lack of documentation (or maybe I just wasn't lucky to find a decent one). But ChatGPT helped me with that.

Overall, working on this app was an interesting and enriching experience, even if an unconvential one.