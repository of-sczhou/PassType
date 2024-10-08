# PassType
Handy GUI-script for KeePass databases.
This graphical script for Windows allows you to work with multiple KeePass databases. A simple graphical interface and constant presence in the system tray make it easy and fast to select and use KeePass entries.

This script is designed for system adiministrators, it allows printing credentials from KeePass databases to any consoles (VMware vmrc, Microsoft Hyper-V, HP iLo, Dell iDrac, or whatever you use).

Each KeePass entry corresponds to its own button. When a button is pressed, the script programmatically types (like keyboard touches) the login name and password into the desired authetication prompt. Secret information does not fit on the clipboard. Pressing the button with Ctrl only sends the password.

The script does not require Keepass software installed, local or network access to KeePass databases is sufficient.

To operate the program, an external PowerShell module poshkeepass is used. The original module page can be found at https://github.com/PSKeePass/PoShKeePass. For simplicity, the current version of the poshkeepass module is located in the script folder under the same name.
