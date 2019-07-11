
' Output file for report
fileOutput = "Allow_RelayIP.txt"
' The Vitual SMTP Server ID
virtualSMTPServer = "IIS://localhost/smtpsvc/1"
Set fso = CreateObject("Scripting.FileSystemObject")
Set allowFile = fso.opentextfile(fileOutput,2,true)
 
Set objSMTP = GetObject(virtualSMTPServer)
Set objRelayIpList = objSMTP.Get("RelayIpList")
objCurrentList = objRelayIpList.IPGrant

i = 0
For Each objIP in objCurrentList
	arrayFields = Split(objIP, ", ")
    allowFile.WriteLine arrayFields(0)
	Wscript.Echo arrayFields(0)
    i = i + 1
Next
Wscript.Echo "Saved in " & fileOutput
If i = 0 Then
    Wscript.Echo "No entries found"
End If
