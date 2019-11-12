VERSION 5.00
Begin VB.Form packag 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PACKAGE"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18315
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "package.frx":0000
   ScaleHeight     =   10455
   ScaleWidth      =   18315
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8385
      Left            =   1305
      TabIndex        =   0
      Top             =   630
      Width           =   14775
      Begin VB.CommandButton Command10 
         Height          =   555
         Left            =   10125
         Picture         =   "package.frx":1CD36
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7380
         Width           =   1500
      End
      Begin VB.CommandButton close2 
         Height          =   555
         Left            =   12600
         Picture         =   "package.frx":1DDBF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7110
         Width           =   1455
      End
      Begin VB.CommandButton new5 
         Height          =   555
         Left            =   5850
         Picture         =   "package.frx":1EC4D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7020
         Width           =   1500
      End
      Begin VB.CommandButton add 
         Height          =   555
         Left            =   7875
         Picture         =   "package.frx":1FAD8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6975
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton update 
         Height          =   555
         Left            =   5355
         Picture         =   "package.frx":24E46
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5850
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11070
         TabIndex        =   2
         Top             =   1575
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.OptionButton Opt2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Private Packages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   9945
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         Height          =   6990
         Left            =   450
         Top             =   945
         Width           =   13965
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         Height          =   6990
         Left            =   360
         Top             =   1080
         Width           =   13965
      End
   End
End
Attribute VB_Name = "packag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim s As String
Dim r As ADODB.Recordset
'Dim c As ADODB.Connection


Private Sub acharge_KeyPress(KeyAscii As Integer)

If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
'If acharge.Text = "" Then
'MsgBox "Enter Adult Charge", vbCritical, "Tour & Travels"
'acharge.SetFocus
'Else
ccharge.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
'End If
End Sub

Private Sub Add_Click()

If Opt1.Value = True Then

If M() = True Then
Exit Sub
End If


Text2.Visible = False
connect
s = "insert into package values('" + id2.Text + "', '" + Combo6.Text + "','" + Combo5.Text + "','" + Combo10.Text + "'," + tourc.Text + ")"
MsgBox s
Set r = c.Execute(s)
MsgBox "Data added succesfully", vbOKOnly, "ADD"
Clear
Unload Me
packag.Show
End If

If Opt2.Value = True Then

If Message() = True Then
Exit Sub
End If
'Else
Text1.Visible = False
s = "insert into cpackage values ('" & Text6.Text & "'," & Text5.Text & ",'" + Combo8.Text + "','" + Combo4.Text + "','" & Combo7.Text & "','" & Combo9.Text & "'," & Text2.Text & "," & Text1.Text & ")"
MsgBox s
Set r = c.Execute(s)
MsgBox "Data added Succesfully", vbOKOnly, "ADD"
clear1
End If
End Sub

Private Sub ccharge_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
tcharge.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub
Private Sub charge_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
tourc.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub close2_Click()
Unload Me
newindex.Show
End Sub

Private Sub Combo1_Click()
Set r = New ADODB.Recordset
A = "select * from package where id='" & Combo1.Text & "'"
MsgBox A
Set r = c.Execute(A)
Combo6.Text = r.Fields(1)
Combo5.Text = r.Fields(2)
Combo10.Text = r.Fields(3)
tourc.Text = r.Fields(4)
'tcharge.Text = r.Fields(9)

End Sub

Private Sub Combo2_Click()
A = "select * from Cpackage where id='" & Combo2.Text & "'"
MsgBox A
Set r = c.Execute(A)
Text5.Text = r.Fields(1)
Combo2.Text = r.Fields(0)
Combo8.Text = r.Fields(2)
Combo4.Text = r.Fields(3)
Combo7.Text = r.Fields(4)
Combo9.Text = r.Fields(5)
Text2.Text = r.Fields(6)
Text1.Text = r.Fields(7)
'charge.Text = r.Fields(8)
'tcharge.Text = r.Fields(9)
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Combo3.Text = "" Then
MsgBox "Enter Hotel Name", vbCritical, "Tour & Travels"
Combo3.SetFocus
Else
Text8.SetFocus
End If
End If
End Sub

Private Sub Combo5_Click()
Set r = New ADODB.Recordset
A = "select * from BUS where NAME='" & Combo5.Text & "'"
MsgBox A
Set r = c.Execute(A)
Combo10.Text = r.Fields(2)
End Sub

Private Sub Combo6_click()
Set r = New ADODB.Recordset
A = "select  *from car where place = '" & Combo6.Text & "'"
Set r = c.Execute(A)
While r.EOF <> True
Combo6.AddItem r.Fields(0)
r.MoveNext
Wend
'A = "select * from hotel where place='" & Combo6.Text & "'"
'MsgBox A
'Set r = c.Execute(A)

End Sub

Private Sub Combo6_LostFocus()
s = "select count(id) from package"
Set r = c.Execute(s)
id2.Text = UCase(Combo6.Text & r.Fields(0) + 1)
End Sub





Private Sub Combo8_Click()
Combo7.Clear
Combo9.Clear
Set r = New ADODB.Recordset
A = "select  *from car where place = '" & Combo8.Text & "'"
Set r = c.Execute(A)
While r.EOF <> True
Combo7.AddItem r.Fields(0)
Combo9.Text = r.Fields(4)
r.MoveNext
Wend


End Sub

Private Sub Combo8_LostFocus()
Combo4.Clear
A = "select  *from hotel where place = '" & Combo8.Text & "'"
Set r = c.Execute(A)
While r.EOF <> True
Combo4.AddItem r.Fields(2)
r.MoveNext
Wend
End Sub

Private Sub Command10_Click()
On Error GoTo q:
If Opt1.Value = True Then
If d() = True Then
Exit Sub
End If

    s = "delete from package where ID='" & Combo1.Text & "'"
    Set r = c.Execute(s)
    Combo1.RemoveItem Combo1.ListIndex
    MsgBox "Data Deleted Succesfully", , "Delete"
    Clear
    End If
If Opt2.Value = True Then
If f() = True Then
Exit Sub
End If

    s = "delete from cpackage where ID='" & Combo2.Text & "'"
    Set r = c.Execute(s)
    Combo2.RemoveItem Combo2.ListIndex
    MsgBox "Data Deleted Succesfully", , "Delete"
    clear1
    End If
q:
End Sub

Private Sub id2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub opt1_Click()
Combo1.Visible = True
Combo1.Enabled = True
Combo2.Visible = True
Combo2.Enabled = False
id2.Visible = False
new5.Visible = True
Combo1.SetFocus
clear1
End Sub

Private Sub opt2_Click()
Clear
Combo2.Visible = True
Combo2.Enabled = True
Combo2.SetFocus
id2.Visible = False
Combo1.Visible = True
Combo1.Enabled = False
lblpack.Visible = True
Combo1.Visible = True
new5.Visible = True
add.Visible = False

End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
add.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"

End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Text1.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"

End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text7.Text = "" Then
MsgBox "Enter Place Name", vbCritical, "Tour & Travels"
Text7.SetFocus
Else
Combo3.SetFocus
End If
End If
End Sub

Private Sub text8_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
vehicle.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
Text8.Text = ""
End If
End Sub
Private Sub Form_Load()
Dim B As String
Dim U As String

connect
s = "select count(id) from package"
Set r = c.Execute(s)
id2.Text = UCase(Combo6.Text & r.Fields(0) + 1)
s = "select count(id) from cpackage"
Set r = c.Execute(s)
Text6.Text = UCase(Combo8.Text) & r.Fields(0) + 1
A = "select id from package"
Set r = c.Execute(A)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select distinct place from hotel"
Set r = c.Execute(A)
While r.EOF <> True
Combo6.AddItem r.Fields(0)
'Combo8.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select ID from CPACKAGE"
Set r = c.Execute(A)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select distinct NAME from BUS"
Set r = c.Execute(A)
While r.EOF <> True
Combo5.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select distinct place from CAR"
Set r = c.Execute(A)
While r.EOF <> True
Combo8.AddItem r.Fields(0)
r.MoveNext
Wend

Exit Sub
Opt1.SetFocus
End Sub

Private Sub new5_Click()
If Opt1.Value = False And Opt2.Value = False Then
MsgBox "choose the package", vbOKOnly, "TOUR & TRAVEL"
'Opt1.SetFocus
End If
If Opt1.Value = True Then
lblpack.Visible = False
Combo1.Visible = False
id2.Visible = True
p(0).Visible = True
id2.SetFocus
new5.Visible = False
add.Visible = True
'Combo1.SetFocus
Clear
End If
If Opt2.Value = True Then
new5.Visible = False
pack(1).Visible = True
Text6.Visible = True
Text6.SetFocus
Combo2.Visible = False
Label7.Visible = False
add.Visible = True
clear1
'Combo2.Visible = False
'Text6.Visible = True
'lblpack.Visible = False
'Combo1.Visible = False
'p(o).Visible = True
'id2.Visible = True
'Text8.Enabled = True
'vn.Enabled = True
'vehicle.Enabled = True
'tourc.Enabled = True
'acharge.Enabled = True
'ccharge.Enabled = True
'tcharge.Enabled = True
'add.Visible = True
'new5.Visible = False
'Text8.SetFocus
End If
End Sub

'Private Sub Opt2_Click()
'Text6.Visible = True
'Combo2.Visible = True
'Text5.Visible = True
'Text9.Visible = True
'Combo4.Visible = True
'Text4.Visible = True '
'Text3.Visible = True
'Text2.Visible = True
''Text1.Visible = True
'pack(1).Visible = True
'Label1.Visible = True
'Label(0).Visible = True
'hn(5).Visible = True
'car(2).Visible = True
'cn(0).Visible = True
'tc.Visible = True
'pc.Visible = True



'End Sub

Private Sub tcharge_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
'add.Visible = True
add.Enabled = True
add.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub tourc_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
acharge.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub
Private Sub update_Click()
'On Error GoTo c:
If Opt1.Value = True Then
If d() = True Then
Exit Sub
End If
'Else
s = "update package set place='" & Combo6.Text & "',vehicle='" & Combo5.Text & "',no='" & Combo10.Text & "',tourc=" & tourc.Text & "where id ='" & Combo1.Text & "'"
Set r = c.Execute(s)
MsgBox "Data Updated Succesfully", , "Update"
Clear
End If
If Opt2.Value = True Then

If f() = True Then
Exit Sub
End If

s = "update cpackage set days=" & Text5.Text & ",PLACE='" & Combo8.Text & "',HNAME= '" & Combo4.Text & "',VEHICLE ='" & Combo7.Text & "',NO='" & Combo9.Text & "',tourc=" & Text2.Text & ",packc=" & Text1.Text & "where id='" & Combo2.Text & "' "
MsgBox s
Set r = c.Execute(s)
MsgBox "Data Updated Succesfully", , "Update"
'new1.Enabled = True
clear1
End If
'c:
End Sub
Function Clear() As Boolean
Combo6.Text = ""
Combo10.Text = ""
'Text8.Text = ""
Combo5.Text = ""
'Text3.Text = ""
tourc.Text = ""
'acharge.Text = ""
'ccharge.Text = ""
'tcharge.Text = ""
End Function
Function clear1() As Boolean
Combo7.Text = ""

Text5.Text = ""
'    Text4.Text = ""
    Combo5.Text = ""
    Text2.Text = ""
    Text1.Text = ""
   Combo8.Text = ""
    Combo4.Text = ""
End Function

Private Sub vehicle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If vehicle.Text = "" Then
MsgBox "Enter Bus Name", vbCritical, "Tour & Travels"
vehicle.SetFocus
Else
vn.SetFocus
End If
End If
End Sub

Private Sub vn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If vn.Text = "" Then
MsgBox "Enter Bus No", vbCritical, "Tour & Travels"
vn.SetFocus
Else
tourc.SetFocus
End If
End If
End Sub
Function Message() As Boolean
If Combo2.Text = "Select Package" Then
MsgBox "Please Select Package ", vbCritical, "TOUR & TRVEL"
Combo2.SetFocus
Message = True
ElseIf Text5.Text = "" Then
MsgBox "Please Enter Day", vbCritical, "TOUR & TRVEL"
Text5.SetFocus
Message = True
ElseIf Combo8.Text = "" Then
MsgBox "Please select place", vbCritical, "TOUR & TRVEL"
Combo8.SetFocus
Message = True
ElseIf Combo4.Text = "" Then
MsgBox "Please select hotel name", vbCritical, "TOUR & TRVEL"
Combo4.SetFocus
Message = True
ElseIf Combo7.Text = "" Then
MsgBox "Please select car name", vbCritical, "TOUR & TRVEL"
Combo7.SetFocus
Message = True

ElseIf Combo9.Text = "" Then
MsgBox "Please select car no", vbCritical, "TOUR & TRVEL"
Combo9.SetFocus
Message = True
ElseIf Text2.Text = "" Then
MsgBox "Please enter tour cost", vbCritical, "TOUR & TRVEL"
Text2.SetFocus
Message = True

ElseIf Text1.Text = "" Then
MsgBox "Please enter package charge", vbCritical, "TOUR & TRVEL"
Text1.SetFocus
Message = True
End If
End Function
Function M() As Boolean


If Combo6.Text = "" Then
MsgBox "PLEASE SELECT PLACE", vbCritical, "TOUR & TRVEL"
Combo6.SetFocus
M = True

ElseIf Combo5.Text = "" Then
MsgBox "PLEASE SELECT BUS NAME", vbCritical, "TOUR & TRVEL"
Combo5.SetFocus
M = True
ElseIf Combo10.Text = "" Then
MsgBox "PLEASE SELECT CAR NO", vbCritical, "TOUR & TRVEL"
Combo10.SetFocus
M = True
ElseIf tourc.Text = "" Then
MsgBox "PLEASE ENTER TOUR CHARGES", vbCritical, "TOUR & TRVEL"
Combo10.SetFocus
M = True
End If
End Function
Function f() As Boolean
If Text5.Text = "" Then
MsgBox "Please Select Pckage", vbCritical, "TOUR & TRVEL"
Combo2.SetFocus
c = True
End If
End Function
Function d() As Boolean
If tourc.Text = "" Then
MsgBox "PLEASE select packages", vbCritical, "TOUR & TRVEL"
Combo1.SetFocus
d = True
End If
End Function
