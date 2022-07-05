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
import sys

def main():
    config_file = {
        "profiles-dir": "C:\\Users\\rtxa\\.vscode\\profiles",
        "vscode-exe": "C:\\Users\\rtxa\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
        "vscode-menu-icon": "icons\\icon-blue.ico",
        "vscode-menu-dir": "vscode-profiles",
        "vscode-menu-name": "Open with a Code profile",
        "profiles": []
    }
    
    # Change dir to where python script is located
    # This will make easier to work with relative paths
    os.chdir(get_dir_exe())

    # Load profiles config
    with open("vscode-profiles.json", "r") as file:
        config_file = json.load(file)
    
    # Clean up old keys before creating new menus (thus to avoid having old items in the mnu)
    ContextMenu.clear_menu(ContextMenu.FT_ALLFILES, config_file["vscode-menu-dir"])
    ContextMenu.clear_menu(ContextMenu.FT_DIRECTORY, config_file["vscode-menu-dir"])
    ContextMenu.clear_menu(ContextMenu.FT_DIRECTORY_BG, config_file["vscode-menu-dir"])

    # Now create the menus for every operation (opening a file, a directory and the inside of a dir)
    ContextMenu.create_menu(ContextMenu.FT_ALLFILES, config_file["vscode-menu-dir"], config_file["vscode-menu-name"], config_file["vscode-menu-icon"])
    ContextMenu.create_menu(ContextMenu.FT_DIRECTORY, config_file["vscode-menu-dir"], config_file["vscode-menu-name"], config_file["vscode-menu-icon"])
    ContextMenu.create_menu(ContextMenu.FT_DIRECTORY_BG, config_file["vscode-menu-dir"], config_file["vscode-menu-name"], config_file["vscode-menu-icon"])

    # Now create the profiles items in the menu
    for profile in config_file["profiles"]:
        # Right-click on file
        ContextMenu.add_item(
            ContextMenu.FT_ALLFILES, 
            config_file["vscode-menu-dir"], 
            profile["name"], 
            profile["name-ui"], 
            profile["icon"],
            vscode_cmd(config_file["vscode-exe"], "%1", config_file["profiles-dir"], profile["name"])
        )  

        # Right-click on directory
        ContextMenu.add_item(
            ContextMenu.FT_DIRECTORY, 
            config_file["vscode-menu-dir"], 
            profile["name"], 
            profile["name-ui"], 
            profile["icon"],
            vscode_cmd(config_file["vscode-exe"], "%1", config_file["profiles-dir"], profile["name"])
        )

        # Right-click on background of directory
        ContextMenu.add_item(
            ContextMenu.FT_DIRECTORY_BG, 
            config_file["vscode-menu-dir"], 
            profile["name"], 
            profile["name-ui"], 
            profile["icon"],
            vscode_cmd(config_file["vscode-exe"], "%V", config_file["profiles-dir"], profile["name"])
        )

    print("Profiles have been added with success!")

# Gets the directory path of the running python script
def get_dir_exe():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

# This is the cmd to open vscode with a profile and a selected file
# content_type: %1 is a file to open
#               %V is a folder
def vscode_cmd(vscode_exe, content_type, profiles_path, profile_name):
    string = "\"{vscode_exe}\" \"{content_type}\" --extensions-dir \"{profiles_path}\\{profile_name}\\extensions\" --user-data-dir \"{profiles_path}\\{profile_name}\\data\""
    return string.format(vscode_exe=vscode_exe, content_type=content_type, profiles_path=profiles_path, profile_name=profile_name)


# Class to create menus in the context menu
class ContextMenu:
    FT_ALLFILES = "*"
    FT_DIRECTORY = "Directory"
    FT_DIRECTORY_BG = "Directory\\Background"

    @staticmethod
    def create_menu(filetype, folder, title, iconPath=""):
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}".format(filetype=filetype, folder=folder)) as key:
            winreg.SetValueEx(key, "MUIVerb", 0, winreg.REG_SZ, title)
            winreg.SetValueEx(key, "SubCommands", 0, winreg.REG_SZ, "")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, os.path.abspath(iconPath))

    @staticmethod
    def add_item(filetype, folder, name, nameUI, iconPath, cmd):
        shellKey = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}\\shell\\{name}".format(filetype=filetype, folder=folder, name=name))
        shellCmdKey = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}\\shell\\{name}\\command".format(filetype=filetype, folder=folder, name=name))
        winreg.SetValueEx(shellKey, "", 0, winreg.REG_SZ, nameUI)
        winreg.SetValueEx(shellKey, "Icon", 0, winreg.REG_SZ, os.path.abspath(iconPath))
        winreg.SetValueEx(shellCmdKey, "", 0, winreg.REG_SZ, cmd)

    @staticmethod
    def clear_menu(filetype, folder):
        try:
            reg_del_node(winreg.HKEY_CLASSES_ROOT, "{filetype}\\shell\\{folder}".format(filetype=filetype, folder=folder))
        except:
            pass

# Deletes a registry key and all its subkeys / values.
# Source: https://stackoverflow.com/questions/38205784/python-how-to-delete-registry-key-and-subkeys-from-hklm-getting-error-5
def reg_del_node(key, current_key, arch_key=0, debug=0):
    open_key = winreg.OpenKey(key, current_key, 0, winreg.KEY_ALL_ACCESS | arch_key)
    info_key = winreg.QueryInfoKey(open_key)
    # Delete the key and all sub_keys.
    for x in range(0, info_key[0]):
        sub_key = winreg.EnumKey(open_key, 0)
        try:
            winreg.DeleteKey(open_key, sub_key)
        except OSError:
            reg_del_node(key, "\\".join([current_key,sub_key]), arch_key)
            # No extra delete here since each call
            # to delete_sub_key will try to delete itself when its empty.

    winreg.DeleteKey(open_key, "")
    open_key.Close()

if __name__ == "__main__":
    main()