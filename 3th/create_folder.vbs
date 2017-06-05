Option Explicit 
Dim objFSO, objFolder, strDirectory, i 
strDirectory = "C:\harsono\upload\" 

Set objFSO = CreateObject("Scripting.FileSystemObject") 
i = 1  '' <===== CHANGED!
While i < 9000
    Set objFolder = objFSO.CreateFolder(strDirectory & i) 
    i = i+1 
    ''WScript.Quit '' <===== COMMENTED OUT!
Wend 