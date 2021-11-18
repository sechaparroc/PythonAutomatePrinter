#Script that keeps looking when a Folder has changed and applies the action defined by the config file
#required libraries to manage printers
#based on http://timgolden.me.uk/python/win32_how_do_i/watch_directory_for_changes.html

import win32api
import win32print
import dill as pickle
import os
import os.path
import os, time
import sys
import ghostscript
import locale
import math
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


# Not used since is not working as Expected

#Taken from https://stackoverflow.com/questions/27195594/python-silent-print-pdf-to-specific-printer
# Use GhostScript API to silent print .pdf and .ps. Use win32api to print .txt. Return a error if printing failed or
# file ext doesn't match.
# def printPDF(printer, folder, file, settings):
#     filepath = folder + "/" + file
#     #create if not exists
#     Path(f'{folder}/tmp').mkdir(parents=True, exist_ok=True)
# 
#     collate = "/Collate " + ("true" if settings["Collate"] == 1 else "false")
#     collate = ""
#     copies = f"/NumCopies {settings['Copies']}"
#     duplex = "" if settings["Duplex"] <= 1 else "/Duplex true"
#     # DMORIENT_PORTRAIT (1) or DMORIENT_LANDSCAPE (2)
#     #  0 (portrait), 3 (landscape), 2 (upside down), and 1 (seascape)
#     orientation = "/Orientation" + (" 0" if settings['Orientation'] == 1 else  "3") 
# 
#     # PaperLength -> Specified in 1/10 millimeters
#     # PaperWidth  -> Specified in 1/10 millimeters
#     # https://www.ghostscript.com/doc/9.25/Use.htm#Known_paper_sizes
# 
#     # Not required since PDF contains this info
#     # https://stackoverflow.com/questions/7446552/resizing-a-pdf-using-ghostscript
#     paperLength = millimetersToPoints(settings['PaperLength'] / 10)
#     paperWidth = millimetersToPoints(settings['PaperWidth'] / 10)
# 
#     scale = ""
#     #Not tested, hopefully works
#     if str(settings['Scale']) != '100':
#         scale = "/Install { " + settings['Scale'] / 100.0 + " " + settings['Scale'] / 100.0 + " scale } bind"
# 
#     args = [
#         "-dPrinted", "-dBATCH", "-dNOSAFER", "-dNOPAUSE", "-dNOPROMPT", "-q",
#         "-sDEVICE#ps2write",
#         #f'-dDEVICEHEIGHTPOINTS#{paperLength}',
#         #f'-dDEVICEWIDTHPOINTS#{paperWidth}',                
#         f'-sOutputFile#"{folder}/tmp/{file}.ps"',
#         f'-c', f'"<</PSDocOptions (<<{collate} {duplex} {copies} {orientation} {scale}>> setpagedevice)>> setdistillerparams"', 
#         f'-f',
#         f'-r{settings["PrintQuality"]}',
#         f'"{filepath}"'
#     ]
#
#    #Check https://www.ghostscript.com/doc/current/VectorDevices.htm
#    if settings["Color"] == 0: 
#        args.append(f'-sColorConversionStrategy#Gray')
#    encoding = locale.getpreferredencoding()
#    args = [a.encode(encoding) for a in args]
#    ghostscript.Ghostscript(*args)
#    #print temporal ps and then remove it
#    printPS(printer, f'{folder}/tmp', f'{file}.ps', settings)

def printPS(printer, folder, file, settings):
    filepath = folder + "/" + file
    print("...Printing File", filepath)
    for i in range(settings["Copies"]):
        args = [
            "-dPrinted", "-dBATCH", "-dNOSAFER", "-dNOPAUSE", "-dNOPROMPT", "-q",
            #f'-dNumCopies#{settings["Copies"]}',  -- Not working properly
            "-sDEVICE#mswinpr2",
            f'-sOutputFile#"%printer%{printer}"',
            f'-r{settings["PrintQuality"]}',
            f'"{filepath}"'
        ]
        #Check https://www.ghostscript.com/doc/current/VectorDevices.htm
        if settings["Color"] == 0: 
            args.append(f'-sColorConversionStrategy=Gray')
        encoding = locale.getpreferredencoding()
        args = [a.encode(encoding) for a in args]
        print(args)
        ghostscript.Ghostscript(*args)

def printFile(printer, folder, file, settings):
    filepath = folder + "/" + file
    try:
        if os.path.splitext(filepath)[1] in [".pdf"]:
            printPS(printer, folder, file, settings)
        elif os.path.splitext(filepath)[1] in [".ps"]:
            printPS(printer, folder, file, settings)
        else : 
            win32api.ShellExecute(0, "printto", '"%s"' % filepath, '"%s"' % printer, ".", 0)
        return True
    except:
        print("Printing error for file: ", '"%s"' % filepath, "| Printer: ", '"%s"' % printer)
        return False

def millimetersToPoints(x):
    inches = math.ceil(x * 0.0393701)
    points = inches * 72
    return points


def executeSettings(settings, folder, files):
    print("----")

    folder_settings = settings[folder]
    for printer in folder_settings:
        for file in files: #Print each file using settings
            pHandle = win32print.OpenPrinter(printer, {"DesiredAccess":win32print.PRINTER_ALL_ACCESS})  
            # Get the default properties for the printer
            properties = win32print.GetPrinter(pHandle, 2)
            #Set properties according to settings
            for prop in folder_settings[printer]:
                try : 
                    setattr(properties['pDevMode'], prop, folder_settings[printer][prop])                
                except : 
                    print("pass", getattr(properties['pDevMode'], prop))
                    pass
            win32print.SetPrinter(pHandle, 2, properties, 0)
            #Print
            try:
                #win32print.SetDefaultPrinter(printer)
                #print(folder + '/' + file)
                #win32api.ShellExecute(0, "printto", f'"{folder}/{file}"', f'"{printer}"', ".", 0)
                printFile(printer, folder , file, folder_settings[printer])

            except Exception as e:
                print(e)
                print("Error", "There was an error printing the file :(")

            win32print.ClosePrinter(pHandle)


def setState(settings):
    state = dict()
    for folder in settings:
        try: 
            state[folder] = set(f for f in os.listdir (folder))
        except:
            print("Folder:", folder, "not exists")
    return state

def execute(prevState, settings):
    currentState = setState(settings)
    print("-"*20)
    for folder in currentState:
        added = set (f for f in currentState[folder] if not f in prevState[folder])
        if added: 
            #Print added files
            executeSettings(settings, folder, added)            
            print("Watching", folder, "Status: new files added")
        else:
            print("Watching", folder, "Status: no changes found")
    return currentState

settings = loadSettings()
print(settings)
state = setState(settings)

while 1:
    state = execute(state, settings)
    time.sleep (10)
