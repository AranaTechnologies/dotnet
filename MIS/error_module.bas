Attribute VB_Name = "Module1"
Public Function AUTO_ERROR(Errcode As Integer)

Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
CNN.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
RST.Open "SELECT * FROM ERROR_MESSAGE_file WHERE ERROR_CODE = " & Errcode, CNN, adOpenStatic, adLockOptimistic, adCmdText
MsgBox (" ERROR CODE : " & RST.Fields("ERROR_CODE") & " :: " & RST.Fields("ERROR_MESSAGE"))
RST.Close
CNN.Close

End Function

