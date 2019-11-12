Attribute VB_Name = "Module1"
Public r As ADODB.Recordset
Public c1 As ADODB.Recordset
Public cbook As ADODB.Recordset
Public D1 As ADODB.Recordset
Public book4 As ADODB.Recordset
Public c As ADODB.Connection
Public A As String
Public cname As String

Public Sub connect()
Set c = New ADODB.Connection
c.Open "provider = msdaora.1;user id=tour/travel;presist security info =true"
Set r = New ADODB.Recordset
End Sub

Public Sub unload2()
Unload travell
newindex.Show
End Sub

Public Function selectText(txtbox As TextBox)
  txtbox.SetFocus
  txtbox.SelStart = 0
  txtbox.SelLength = Len(txtbox.Text)
End Function
