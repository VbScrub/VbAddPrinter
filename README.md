# VbAddPrinter

Valid arguments:

**VbAddPrinter.exe add printer_name driver_name port_name**

Adds a new local printer 
Example: 
VbAddPrinter.exe add "My Printer" "HP LaserJet 1100 PCL6" "lpt1:"


**VbAddPrinter.exe enum**

Lists currently installed printers


**VbAddPrinter.exe addunc \\\\server\\printer**

Adds a network printer shared from another machine
Example:
VbAddPrinter.exe addunc "\\\\192.168.0.12\\HP LaserJet 1100"


More info on why I created this: https://vbscrub.video.blog/2020/03/24/ricoh-printer-exploit-priv-esc-to-local-system/
