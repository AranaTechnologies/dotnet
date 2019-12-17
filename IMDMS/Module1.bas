Attribute VB_Name = "Module1"
Public Function SHOW_ERROR(ENO As Integer)

Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
CNN.Open "DSN=from oracle; PROVIDER=MSDASQL; UID=imdms; PWD=imdms1"
RST.Open "SELECT * FROM ERR WHERE ENO = " & ENO, CNN, adOpenStatic, adLockOptimistic, adCmdText
MsgBox (" ERROR CODE : " & RST.Fields("ENO") & " :: " & RST.Fields("MSG"))
RST.Close
CNN.Close

End Function

 

