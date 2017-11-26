Sub SysInfo()
Dim strCom As String
Dim objWMIService As Object
Dim colAdapters As Object
Dim objAdapter As Object
strCom = "."
Set objWMIService = GetObject _
("winmgmts:" & "!\\" & strCom & "\root\cimv2")
Set colAdapters = objWMIService.ExecQuery _
("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objAdapter In colAdapters
MsgBox "Host name: " & objAdapter.DNSHostName & vbCrLf & _
           "Description: " & objAdapter.Description & vbCrLf & _
           "Physical(MAC) address: " & objAdapter.MACAddress & vbCrLf & _
           "IP address: " & objAdapter.IPAddress(i) & vbCrLf & _
           "Subnet:  " & objAdapter.IPSubnet(i) & vbCrLf & _
           "DNS Suffix: " & objAdapter.DNSDomain & vbCrLf & _
           "Primary WINS server: " & objAdapter.WINSPrimaryServer & vbCrLf & _
           "Secondary WINS server: " & objAdapter.WINSSecondaryServer
Next objAdapter
End Sub
