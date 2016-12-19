' Include CommonFunctions
sFilePath = "C:\Users\2troy\Desktop\OldDropboxFiles\Projects\VBS Scripting\In Progress\01 001 CommonFunctions\CommonFunctions.vbs"
Execute CreateObject("Scripting.FileSystemObject").openTextFile(sFilePath).readAll()
