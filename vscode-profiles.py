# VS Code Profiles Menu / Add profiles to context menu
# Author: rtxa
#
# Add profiles in the context menu for:
# 1. Right-click on a file
# 2. Right-click on a folder
# 3. Right-click inside the background of a folder
#
# To do:
# - Create shortcuts in the same dir (https://www.blog.pythonlibrary.org/2010/01/23/using-python-to-create-shortcuts/)
# - Add in settings.json of every profile the name of the current profile
# - Expand enviroment variables if present os.path.expandvars(path)
# - Use Github Actions to build this script to .exe https://github.com/marketplace/actions/pyinstaller-windows

import winreg
import json
import os

def main():
    configFile = {
        "profiles-dir": "C:\\Users\\rtxa\\.vscode\\profiles",
        "vscode-exe": "C:\\Users\\rtxa\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
        "vscode-menu-icon": "C:\\Users\\rtxa\\.vscode\\icons\\icon-blue.ico",
        "vscode-menu-dir": "vscode-profiles",
        "vscode-menu-name": "Open with a Code profile",
        "profiles": []
    }
    
    # Load profiles config
    path = os.path.dirname(os.path.realpath(__file__))
    with open("{path}\\vscode-profiles.json".format(path=path), "r") as jsonFile:
        configFile = json.load(jsonFile)
    
    cmc = ContextMenuCreator()

    # Clean up old keys before creating new menus (thus to avoid having old items in the mnu)
    try:
        regDelNode(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}".format(filetype=cmc.FT_ALLFILES, folder=configFile["vscode-menu-dir"]))
        regDelNode(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}".format(filetype=cmc.FT_DIRECTORY, folder=configFile["vscode-menu-dir"]))
        regDelNode(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}".format(filetype=cmc.FT_DIRECTORY_BG, folder=configFile["vscode-menu-dir"]))
    except Exception:
        pass

    # Now create the menus for every operation (opening a file, a directory and the inside of a dir)
    cmc.createMenu(cmc.FT_ALLFILES, configFile["vscode-menu-dir"], configFile["vscode-menu-name"], configFile["vscode-menu-icon"])
    cmc.createMenu(cmc.FT_DIRECTORY, configFile["vscode-menu-dir"], configFile["vscode-menu-name"], configFile["vscode-menu-icon"])
    cmc.createMenu(cmc.FT_DIRECTORY_BG, configFile["vscode-menu-dir"], configFile["vscode-menu-name"], configFile["vscode-menu-icon"])

    # Now create the profiles items in the menu
    for profileData in configFile["profiles"]:
        # Right-click on file
        cmc.addItem(
            ContextMenuCreator.FT_ALLFILES, 
            configFile["vscode-menu-dir"], 
            profileData["name"], 
            profileData["name-ui"], 
            profileData["icon"],
            vscodeCmd(configFile["vscode-exe"], "%1", configFile["profiles-dir"], profileData["name"])
        )  

        # Right-click on directory
        cmc.addItem(
            ContextMenuCreator.FT_DIRECTORY, 
            configFile["vscode-menu-dir"], 
            profileData["name"], 
            profileData["name-ui"], 
            profileData["icon"],
            vscodeCmd(configFile["vscode-exe"], "%1", configFile["profiles-dir"], profileData["name"])
        )

        # Right-click on background of directory
        cmc.addItem(
            ContextMenuCreator.FT_DIRECTORY_BG, 
            configFile["vscode-menu-dir"], 
            profileData["name"], 
            profileData["name-ui"], 
            profileData["icon"],
            vscodeCmd(configFile["vscode-exe"], "%V", configFile["profiles-dir"], profileData["name"])
        )

    print("Profiles have been added with success!")

# This is the cmd to open vscode with a profile and a selected file
# contentType: %1 is a file to open
#              %V is a folder
def vscodeCmd(vscodeExePath, contentType, profilesPath, profileName):
    strToFormat = "\"{vscodeExePath}\" \"{contentType}\" --extensions-dir \"{profilesPath}\\{profileName}\\extensions\" --user-data-dir \"{profilesPath}\\{profileName}\\data\""
    return strToFormat.format(vscodeExePath=vscodeExePath, contentType=contentType, profilesPath=profilesPath, profileName=profileName)


# Class to create menus in the context menu
class ContextMenuCreator():
    FT_ALLFILES = "*"
    FT_DIRECTORY = "Directory"
    FT_DIRECTORY_BG = "Directory\\Background"

    def createMenu(self, filetype, folder, title, iconPath=""):
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}".format(filetype=filetype, folder=folder)) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, title)
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, iconPath)


    def addItem(self, filetype, folder, name, nameUI, iconPath, cmd):
        shellKey = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}\\shell\\{name}".format(filetype=filetype, folder=folder, name=name))
        shellCmdKey = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}\\shell\\{name}\\command".format(filetype=filetype, folder=folder, name=name))
        winreg.SetValueEx(shellKey, "", 0, winreg.REG_SZ, nameUI)
        winreg.SetValueEx(shellKey, "Icon", 0, winreg.REG_SZ, iconPath)
        winreg.SetValueEx(shellCmdKey, "", 0, winreg.REG_SZ, cmd)

# Deletes a registry key and all its subkeys / values.
# Source: https://stackoverflow.com/questions/38205784/python-how-to-delete-registry-key-and-subkeys-from-hklm-getting-error-5
def regDelNode(key0, current_key, arch_key=0, debug=0):
    open_key = winreg.OpenKey(key0, current_key, 0, winreg.KEY_ALL_ACCESS | arch_key)
    info_key = winreg.QueryInfoKey(open_key)
    for x in range(0, info_key[0]):
        # NOTE:: This code is to delete the key and all sub_keys.
        sub_key = winreg.EnumKey(open_key, 0)
        try:
            winreg.DeleteKey(open_key, sub_key)
            if debug:
                print("Removed %s\\%s " % (current_key, sub_key))
        except OSError:
            regDelNode(key0, "\\".join([current_key,sub_key]), arch_key)
            # No extra delete here since each call
            # to delete_sub_key will try to delete itself when its empty.

    winreg.DeleteKey(open_key, "")
    open_key.Close()
    if debug:
        print("Removed %s" % current_key)
    return

if __name__ == "__main__":
    main()