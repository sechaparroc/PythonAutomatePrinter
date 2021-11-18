#Simple script to save / load printer settings based on Folder
#Based on 
# https://stackoverflow.com/questions/39249360/python-print-pdf-file-with-win32print
# https://stackoverflow.com/questions/11045704/change-printer-tray-with-pywin32

#required libraries to manage printers
from tkinter import filedialog
import win32api
import win32print
from tkinter import *
from tkinter import font # * doesn't import font or messagebox
from tkinter import messagebox
import dill as pickle
import os.path
from pathlib import Path

DM_OUT_BUFFER = 0x02
DM_IN_BUFFER = 0x08
DM_IN_PROMPT = 0x04

def loadSettings():
    path = "./config/config.pkl"
    if os.path.isfile(path):
        pickle_file = open(path, 'rb')
        settings = pickle.load(pickle_file)
        pickle_file.close()
        return settings
    return dict()

def saveSettings(settings):
    pickle_file = open("./config/config.pkl", 'wb')
    pickle.dump(settings, pickle_file)
    pickle_file.close()

def chooseFolder(printerText, text):
    printerText.delete(1.0, END)
    global _folderName
    _folderName.set(filedialog.askdirectory())
    print(_folderName.get())
    settingsInfo(settings, _folderName.get(), text)    


def devmodeInfo(devmode):
    info = dict()
    for name in dir(devmode):  
        if name.startswith("_") or name in ["DriverData", "Clear", "BitsPerPel"] :
            continue
        info[name] = getattr(devmode,name)
    return info

def settingsInfo(settings, folder, text):
    text.delete(1.0, END)
    if folder not in settings:
        return
    else :
        for printer in settings[folder]:
            devmode = settings[folder][printer]
            text.insert(END, '---------------------')
            text.insert(END, printer)    
            for prop in devmode:
                text.insert(END, '\n\t' + prop + ": " + str(devmode[prop]))    
            text.insert(END, '\n')    
            text.insert(END, '---------------------\n')    

def setPrinterSettings(printer, text = None):
    global _devmode
    #Open a Dialog to Set printer settings
    pHandle = win32print.OpenPrinter(printer.get(), {"DesiredAccess":win32print.PRINTER_ALL_ACCESS})  
    # Get the default properties for the printer
    properties = win32print.GetPrinter(pHandle, 2)
    _devmode = properties['pDevMode']
    win32print.DocumentProperties(0, pHandle, printer.get(), _devmode, _devmode, DM_IN_BUFFER | DM_OUT_BUFFER | DM_IN_PROMPT)
    # place info in given textArea
    if text is not None:
        text.delete(1.0, END)
        info = devmodeInfo(_devmode)
        for prop in info:
            text.insert(END, '\n\t' + prop + ": " + str(info[prop]))    
        text.insert(END, '\n')    

def savePrinterSettings(settings, folder, printer, devmode, text):
    if folder in settings:
        settings[folder][printer] = devmodeInfo(devmode)
    else:
        settings[folder] = { printer : devmodeInfo(devmode) }
    #persist
    saveSettings(settings)
    settingsInfo(settings, folder, text)

def deletePrinterSettings(settings, folder, printer, text, printerText):
    if folder in settings:
        if printer in settings[folder]:
            del settings[folder][printer]   
    
    #persist
    saveSettings(settings)
    settingsInfo(settings, folder, text)
    printerText.delete(1.0, END)

def initWindow(width, height):
    window = Tk()
    def_font = font.nametofont("TkDefaultFont")
    def_font.config(size=14)
    window.title("Printing Configuration Script")
    window.geometry(str(width) + "x" + str(height))
    window.tk.call('encoding', 'system', 'utf-8')
    window.grid_columnconfigure(0, weight=1)
    window.grid_columnconfigure(1, weight=1)
    window.grid_columnconfigure(2, weight=1)

    return window

def onPrinterChanges(*args):
    printerLabel.config(text = _printer.get().upper() + " OPTIONS")
    detailsText.delete(1.0, END)

Path('./config/').mkdir(parents=True, exist_ok=True)

# Variables
width, height = 600, 800
_devmode = None

# Get settings
settings = loadSettings()
window = initWindow(width, height)

#Control Variables
_folderName = StringVar("")
_printer = StringVar(window)


#Printer options

# link function to change dropdown
choices = [printer[2] for printer in win32print.EnumPrinters(2)]

popupMenu = OptionMenu(window, _printer, *choices)
Label(window, text="SELECT PRINTER").grid(row = 1, column = 0)
popupMenu.grid(row = 1, column =1, columnspan=2)

printerLabel = Label(window, text = str(_printer.get()).upper() + " OPTIONS" )
printerLabel.grid(row = 2, column = 0, columnspan=3, pady= 15)

settingsLabel = Label(window, text = "Printer Settings Details" )
settingsLabel.grid(row = 4, column = 0, columnspan=3, pady=20)
detailsText = Text(window, height=10, width=60)
detailsText.grid(row = 5, column = 0, columnspan=3)


settingsAllLabel = Label(window, text = "Saved Folder Settings Details" )
settingsAllLabel.grid(row = 6, column = 0, columnspan=3, pady=20)
detailsAllText = Text(window, height=14, width=60)
detailsAllText.grid(row = 7, column = 0, columnspan=3)

# Allows to select a Folder
Label(window, text="SELECT FOLDER").grid(row = 0, column = 0, pady=20, padx = 20)
Entry(window, textvariable = _folderName).grid(row = 0, column = 1, pady=20, ipady=10, ipadx = 10)
chooseButton = Button(window, text=u"\uD83D\uDCC1" ' BROWSE', command= lambda : chooseFolder(detailsText, detailsAllText), bg = "white")
chooseButton.grid(row = 0, column = 2, padx = 20, pady=20)


deleteButton = Button(window, text="Delete Settings", command=lambda:deletePrinterSettings(settings, _folderName.get(), _printer.get(), detailsAllText, detailsText), bg = "white")
deleteButton.grid(row = 3, column = 0)

setButton = Button(window, text="Set Settings", command=lambda : setPrinterSettings(_printer, detailsText), bg = "white")
setButton.grid(row = 3, column = 1)

saveButton = Button(window, text="Save Settings", command=lambda : savePrinterSettings(settings, _folderName.get(), _printer.get(), _devmode, detailsAllText), bg = "white")
saveButton.grid(row = 3, column = 2)



_printer.trace('w', onPrinterChanges)
_printer.set(win32print.GetDefaultPrinter()) # set the default option


window.mainloop()