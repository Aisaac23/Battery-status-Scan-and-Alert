
Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory 'We use the current directory to save evrything so be sure to have this script in the appropriate path

'******************************* We create one file for detailed information and one for the script to use; 2 is for write
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath & "\chargeCapacities - detailed.txt", 2, True) 
Set objFileToWrite2 = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath & "\chargeCapacities.txt", 2, True)

objFileToWrite.Close
objFileToWrite2.Close
'*****************************************************

'************We prepare a WMI object and get the full charged capacity
Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

Set colItems = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")
For Each objItem In colItems
	Full = objItem.FullChargedCapacity
Next
					
'******************************* The process is repeated to register the battery changes as your laoptop discharges.
While (1)

	Set colItems = objWMIService.ExecQuery("Select * From BatteryStatus")	
	For Each objItem In colItems
		If objItem.RemainingCapacity > 0 Then
			remaining = objItem.RemainingCapacity
		End If
	Next
	

	iPercent = ((remaining / Full) * 100) Mod 100

	Set objShell = CreateObject("Wscript.Shell")
	strPath = objShell.CurrentDirectory

	'*************************** 8 is for appending
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath & "\chargeCapacities - detailed.txt", 8, False)
	Set objFileToWrite2 = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath & "\chargeCapacities.txt", 8, False)

	WScript.sleep 30000 ' 30 secs
	'************************************* Date and time ## remaining as raw data ## remaining in porcentage
	objFileToWrite.WriteLine (FormatDateTime(Now) & " ## " & "Plain Remaining:" & remaining & " ## " & "Porcent remaining: " & iPercent)
	objFileToWrite2.WriteLine (iPercent)
	objFileToWrite.Close
	objFileToWrite2.Close

Wend


