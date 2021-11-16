#Script that keeps looking when a Folder has changed and applies the action defined by the config file
#required libraries to manage printers
#based on http://timgolden.me.uk/python/win32_how_do_i/watch_directory_for_changes.html

import win32api
import win32print
import dill as pickle
import os.path
import os, time

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

def executeSettings(settings, folder, files):
    print("----")
    print(files)

    folder_settings = settings[folder]
    for printer in folder_settings:
        for file in files: #Print each file using settings
            pHandle = win32print.OpenPrinter(printer, {"DesiredAccess":win32print.PRINTER_ALL_ACCESS})  
            # Get the default properties for the printer
            properties = win32print.GetPrinter(pHandle, 2)
            #Set properties according to settings
            print(properties['pDevMode'])
            print("////")
            print(folder_settings[printer])
            for prop in folder_settings[printer]:
                print(prop, folder_settings[printer][prop])
                try : 
                    setattr(properties['pDevMode'], prop, folder_settings[printer][prop])                
                except : 
                    print("pass", getattr(properties['pDevMode'], prop))
                    pass
            win32print.SetPrinter(pHandle, 2, properties, 0)

            #Print
            try:
                #win32print.SetDefaultPrinter(printer)
                print(folder + '/' + file)
                win32api.ShellExecute(0, "printto", f'"{folder}/{file}"', f'"{printer}"', ".", 0)

            except:
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
