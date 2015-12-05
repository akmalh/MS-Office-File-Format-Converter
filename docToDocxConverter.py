# Import library modules
import os
import win32com.client

# Get current directory path 
path = os.getcwd()

# Open MS Word application
wrd= win32com.client.Dispatch("Word.Application")
wrd.visible = 0

# Function to convert file format
def convertDocToDocx(filename):
	wrdFile = wrd.Documents.Open(filename)
	convertedFilename = filename+"x"
	print convertedFilename
	wrdFile.SaveAs(convertedFilename, FileFormat = 12)
	wrdFile.Close(True)

# Browse through all subdirectories	
for root, dirs, files in os.walk(path):
    for name in files:
        if name.endswith((".doc")):
			filePath = root + "\\" + name
			print filePath
			convertDocToDocx(filePath)

# Close MS Word application			
wrd.Quit()
