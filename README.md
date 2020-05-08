# Meraki Device Update
A python script that uses the Meraki API to scan though an excel workbook and update device attributes and update switchports.  Created to speed up the provsioning of new Meraki switches, while at the same time providing a port mapping spreadsheet that can be used during the cutover to the new switches. 

The script uses both Meraki SDK and XLRD libraries, both of which can be installed via PIP.  
