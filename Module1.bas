Attribute VB_Name = "Module1"
Public c As ADODB.Connection
Public r As ADODB.Recordset
Public sql As String
Public Function conn()
Set c = New ADODB.Connection
c.Open "provider = msdaora.1;user id = shivam/kumar;persist sequrity info =false"
Set r = New ADODB.Recordset
End Function
