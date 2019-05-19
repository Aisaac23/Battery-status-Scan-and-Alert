'This script reads the "designed charge capacity" of your laptop's battery and the "Full charged capacity" and 
' it provides a percentage, giving a hint of the battery health.  

Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

Set colItems = objWMIService.ExecQuery("Select * From BatteryFullChargedCapacity")

For Each objItem In colItems

    wscript.Echo "FullChargedCapacity: " & objItem.FullChargedCapacity
    Full = objItem.FullChargedCapacity
    
Next

'This method IS NOT always accurate but still could be a good tool to know if your battery needs to be replaced or 
' will need in the future.
'I recommend to execute this script with your charger unplugged

Set colItems = objWMIService.ExecQuery("Select * From BatteryStaticData")

For Each objItem In colItems
    wscript.Echo "DesignedCapacity: " & objItem.DesignedCapacity
    designed = objItem.DesignedCapacity
Next


iPercent = ((Full/designed) * 100) 
MsgBox "Battery LIFE is at " & FormatNumber(iPercent, 2) & "%", vbInformation, "Life Status"
wscript.sleep 30000 ' 5 minutes

