# A metal detail making cost calculator

<b>Это английская версия презентации, выберите ветку presentation_rus, чтобы посмотреть версию на русском.</b>

Technologies: Pandas, PySimpleGui

## How to launch:

Download the repository and run the main.py file.
(soon) Or download the archive with an executable for your system, unpack it and run the executable.

## What it does

The main function of the app is to automate the calculation of the price for a metal detail. Previously, the managers of the company had to manually look through tables and calculate the prices. Now the user can input all detail data in the main window of the application and make a calculation with one click.

The user can get the demonstration of the prices and other detail info, save the detail data to the spreadsheet with previous calculations, or save a separate spreadsheet with the current calculation for print. Also, the user can choose the file the application takes the prices data from (it has to be an Excel file with the structure programmed in the app), the file where it saves the calculations, or change them back to the default settings.

## Difficulties I encountered

### Having to work with spreadsheets

The application is built to work with excel spreadsheets, since it was most convenient for the client an there were no time and money to create a database. I may add the database in the future, to avoid all the complications that come with the app being dependend on external files.

For now, I tried my best to create workarounds for any possible malfunction. For example, the application checks at the startup if the spreadsheet it needs for work (with the prices of metals, cutting and inset) is in place and if it's readable. If it's not, the application asks the user to manually select the spreadsheet to work with.

### Python GUI

Python is not the most popular language for building desktop apps but I did not have the time to learn another language :\) Learning about existing GUI frameworks was interesting. I started with Tkinter and built some of the interface with it, but at some point it became cumbersome and hard to read, so I switched to PySimpleGUI. There was another complication with it, the lack of documentation (or maybe I just wasn't lucky to find a decent one). But ChatGPT helped me with that.

Overall, working on this app was an interesting and enriching experience.