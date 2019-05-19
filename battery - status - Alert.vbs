'********** Reading from file to get the laptop lowest charge capacity

Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory

' 1 for read
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(strPath & "\chargeCapacities.txt",1)

do while not objFileToRead.AtEndOfStream

     strLine = objFileToRead.ReadLine()
loop

lowestP = strLine 

objFileToRead.Close
Set objFileToRead = Nothing
'**************************************************************

While (1)

Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

'***********To get the remaining charge capacity of the laptop 
Set colItems = objWMIService.ExecQuery("Select * From BatteryStatus")

band = false
For Each objItem In colItems

    If objItem.RemainingCapacity > 0 and band = false Then

        remaining = objItem.RemainingCapacity
		charching = objItem.Charging
		plugged =  objItem.PowerOnline
		band = True
    End If

Next

'***************************************

Set colItems = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")

For Each objItem In colItems
    Full = objItem.FullChargedCapacity
Next

iPercent = ((remaining / Full) * 100) Mod 100

'Feel free to modify the number that's going to be added to lowestP deppending on how much battery 
'you want to have left before it reaches its lowestPoint.
If Abs(iPercent) < Abs(lowestP + 10) and Not charging and Not plugged Then
   
    MsgBox "Battery is at " & iPercent & "%." & " User, please connect your charger.", 4096 + 48, "Battery monitor"
    
End If
wscript.sleep 30000 ' 30 secs 

Wend

' * To alwasy run at boot, make a shortcut of your script and paste it in the following folder:
' C:\Users\your-user-name\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
' remember to replace "your-user-name" with your actual user-name
' *Remember that in the same folder where you place this script you need to have the file previously created with 
' the "battery - lowest capacity scan" script. 