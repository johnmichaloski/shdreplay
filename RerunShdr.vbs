'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VB Script to get SHDR Log file and then C++ app to replay on port 7878
' VB Script GUI not flexible enough to allow multiple widget dialog box
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Read MTConnect SHDR Logfile then echo to port 7878
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim sFileSelected
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
'wscript.echo sFileSelected



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Determine script location for VBScript
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim sScriptDir : sScriptDir = oFSO.GetParentFolderName(WScript.ScriptFullName)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Create Config.ini file, with SHDR log file as input
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ini

ini=ini&"[GLOBALS]" & vbCrLf
ini=ini&"Filename= " &  sFileSelected  & vbCrLf
ini=ini&"Repeat=1" & vbCrLf
ini=ini&"Wait=1" & vbCrLf
ini=ini&"PortNum=7878" & vbCrLf
ini=ini&"IP=127.0.0.1" & vbCrLf
ini=ini&"TimeMultipler=1.0" & vbCrLf


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output new COnfig.ini file
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim objFIle
Set objFSO=CreateObject("Scripting.FileSystemObject")
Dim outFile: outFile=sScriptDir  & "\Config.ini"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write ini
objFile.Close

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Run SHDR Echo C++ executable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
wShell.Run(sScriptDir  & "\SHDR_Replay.exe")
' side effect will pop up console
Set wShell = Nothing
' Application will run asynchronously