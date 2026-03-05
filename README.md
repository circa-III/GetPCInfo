# GetPCInfo
A PowerShell script to fetch info about a computer on an enterprise domain. 

## How to use
Download the GetPCInfo.zip file, right click and Extract All, and choose the C:\ drive. The shortcuts (.lnk file) can go anywhere like the desktop or drag to your taskbar. Double click the shortcut to run.

**If the script gives a warning when launching it, right click the .ps1 file, go to Properties, check the Unblock option at the bottom, click apply.**

Wait for it to finish pull the info. If there is red error text saying "Attempting to perform the InitializeDefaultDrives operation on the 'FileSystem' provider failed" you can ignore it if it still pulls info. It checks 4 different registry keys for info and some computers won't have them all.

It should list the name, serial, current logged in user, asset, IP address, current logged in user, Printers installed, software installed (minus a bunch of pre-installed software and drivers), Monitors plugged in, OneDrive sync status, and the AD location and groups.

If the files are not in the C:\ folder the shortcuts won't work until you change the path of the shortcut to point to where you put the files.\

If you have any questions, reach out to Brad Linder (blinder@ecommunity)

## CHANGELOG
-===[ UPDATE 26.03.05]===-
- Replace the previous Get-CimInstance owner lookup to avoid failures when GetOwner cannot be retrieved with OneDrive and silently continue on error.
- update the OneDrive error message to surface the exception message and display it in yellow for clearer diagnostics
-===[ UPDATE 26.03 ]===-
- Cleaned up code, added sections with color
- Added OneDrive sync status
- Added monitor section, shows all plugged in monitors
- Re-wrote AD section
- Gives warning when name isn't pinged or found
  
-===[ UPDATE 26.02 ]===-
- Removed McAfee from the exceptions list so we can see what still has McAfee
- Added ANCILE to list of exceptions now that it's pushed to all machines
- Changed versioning scheme, now is the year.month to simplify
- Added version number at script run
- Removed version from file name so future versions can more easily be replaced

-===[ UPDATE Apr 2025 ]===-
- Added Beyond Trust, OneDrive printer, PaperCut to exceptions
- Added a new file for searching by asset tag (AD description)
- Added shortcuts that should run the files as admin by double clicking (must be in C:\ or changed path in shortcut if not)

-===[ UPDATE Apr 2024 ]===-
- Completely redid the software report to use registry instead of Win32_Product to be more accurate
- Removed the .txt file requirement, instead asks within the terminal window
- Included user-specific installs not just machine-wide
- Filtered Microsoft published, but exceptions for 365, Teams, Visio, Access, PowerBI
- Included MAC address

-===[ UPDATE Feb 2023 ]===-
- Added BloxOne and Cortex to the filter so they no longer are listed
- Added IP address and Asset tag (Computer Description) to the items pulled
- Fixed the capitalization in username
