# VSCode Profiles Script 

![Author](https://img.shields.io/badge/Author-rtxa-9cf "Author") ![Version](https://img.shields.io/badge/Version-0.1-blue "Version")

Script to add profiles to context menu and desktop shortcuts for ease of use.

![](https://imgur.com/7chsQX6.png)

This script works by using the command-line options `--extensions-dir` and `--user-data-dir` provided by VSCode to run isolated instances.

# Requirements

Only if you want to run the Python script instead of the bundled executable.

- Python 3
- Module `pywin32`

# How to use it

Before you use it, check out that `vscode-profiles.json` fields are correct (VSCode executable location, profiles directory, etc.). This should work out of the box.

1. Run the `.py` script or the bundled executable.
2. Now you are ready to use the profiles. No extra configuration needed :)

If you want to add more profiles, just modify the `profiles` section in `vscode-profiles.json` and run again the script.

# Credits

@equiman - Creator of the icons and owner of the [tutorial series](https://dev.to/equiman/series/8983) I follow to create your own profiles. 

