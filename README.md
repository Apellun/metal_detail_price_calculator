# A metal detail making cost calculator

<b>This is the english version of the app, please switch to presentation_rus for the Russian version</b>

Technologies: Pandas, PySimpleGui

## How to launch:

Download the repository and run the main.py file.
Or download the archive with an executable for your system, unpack it and run the executable.

## What it does

The main function of the app is to automate the calculation of the price for a metal detail. Previously, the managers of the company had to manually look through tables, get all variables and calculate the prices.

The main window allows the user to input all detail data. When the user can call up for the demonstration of the prices and other detail info they might need, save the data to the spreadsheet with previous accounts, or save a separate spreadsheet with the current calculation only for print.

Also, the user can change the files where the application looks for the data for the calculation and also where it writes the data of the calculation (and turn it back to the default settings).

## Difficulties I encountered while building it

### Having to work with spreadsheets

The application is built to work with excel spreadsheets, since it was most convenient for the client was and creating the database for it was not in alignement with the time and money the company had for the development. I may add the database in the future, to avoid all the complications that come with the app being dependend on external files.

For now, I tried my best to create workarounds for any possible malfunction that could happen because of it. For example, the application checks at the startup if the spreadsheet it needs for work is in place. If it's not, it provides the user with the opportunity to manually select the spreadsheet to get data from.

### Python GUI

I understand that Python is not the most popular language for building desktop apps, but again, considering the time and the budget, I did not have the time to learn another language :\) Actually, having to find out about existing GUI frameworks was enriching to say the least. I started with Tkinter and built some of the interface with it, but at some point it became cumbersome and hard to read, so I switched to PySimpleGUI. There was another complication with that, the lack of documentation (or maybe I just wasn't lucky to find a decent one). But ChatJPT helped me with that.

Overall, working on this app was an interesting and enriching experience, even if an unconvential one.