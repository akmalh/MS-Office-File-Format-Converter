# Import library modules
import os
import win32com.client

# Get current directory path 
path = os.getcwd()

# Open MS Excel application
exl= win32com.client.Dispatch("Excel.Application")
exl.visible = 0

# Function to convert file format	
def convertXslToXlsx(filename):
	exlFile = exl.Workbooks.Open(filename)
	convertedFilename = filename+"x"
	print convertedFilename
	exlFile.SaveAs(convertedFilename, FileFormat = 51)
	try:
		exlFile.close(True)
	except:
		print "Exception Found"

# Browse through all subdirectories			
for root, dirs, files in os.walk(path):
    for name in files:
        if name.endswith((".xls")):
			filePath = root + "\\" + name
			print filePath
			convertXslToXlsx(filePath)

# Close MS Excel application				
exl.Quit()
