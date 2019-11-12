VERSION 5.00
Begin VB.Form inde 
   Caption         =   "INDEX"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "index.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   11700
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2970
      Width           =   2475
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4950
      Width           =   2475
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   11700
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4545
      Width           =   2475
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   855
      Left            =   11700
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   5805
      Width           =   2490
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   11700
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5175
      Width           =   2475
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2295
      Width           =   2475
   End
   Begin VB.OptionButton Opt2 
      Caption         =   "CAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   810
      Width           =   1995
   End
   Begin VB.OptionButton Option1 
      Caption         =   "BUS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   810
      Width           =   1950
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   11700
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3690
      Width           =   2475
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   11685
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2295
      Width           =   2475
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   435
      Left            =   11700
      TabIndex        =   17
      Text            =   "select package"
      Top             =   1665
      Width           =   2490
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   6015
      TabIndex        =   15
      Text            =   "select package"
      Top             =   1650
      Width           =   2490
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4365
      Width           =   2475
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3690
      Width           =   2475
   End
   Begin VB.TextBox B1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   765
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   8805
      Width           =   16275
   End
   Begin VB.CommandButton clos 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9900
      Picture         =   "index.frx":1A822
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8055
      Width           =   1485
   End
   Begin VB.CommandButton detail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7830
      Picture         =   "index.frx":1FC3A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8055
      Width           =   1485
   End
   Begin VB.CommandButton bookin 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5850
      Picture         =   "index.frx":25368
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8055
      Width           =   1485
   End
   Begin VB.TextBox hinfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   855
      Left            =   6030
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6255
      Width           =   2490
   End
   Begin VB.TextBox busn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5580
      Width           =   2475
   End
   Begin VB.TextBox day 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000004&
      Height          =   495
      Left            =   6030
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3015
      Width           =   2475
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   7035
      Left            =   2700
      Top             =   450
      Width           =   11850
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   6990
      Left            =   2790
      Top             =   540
      Width           =   11805
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Cost"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   9180
      TabIndex        =   35
      Top             =   2970
      Width           =   1290
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Cost"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   3555
      TabIndex        =   33
      Top             =   4950
      Width           =   1290
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9180
      TabIndex        =   30
      Top             =   3690
      Width           =   675
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hotel information"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   9180
      TabIndex        =   28
      Top             =   6165
      Width           =   2265
   End
   Begin VB.Label lbldept 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Place"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   9180
      TabIndex        =   26
      Top             =   5220
      Width           =   735
   End
   Begin VB.Label lbldept 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Place"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   3555
      TabIndex        =   24
      Top             =   2385
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Car name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9180
      TabIndex        =   20
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Package charges"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9180
      TabIndex        =   18
      Top             =   2385
      Width           =   2175
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Packages"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   9180
      TabIndex        =   16
      Top             =   1665
      Width           =   2100
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Child charges"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3555
      TabIndex        =   13
      Top             =   4455
      Width           =   1785
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Adult charges"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3555
      TabIndex        =   11
      Top             =   3780
      Width           =   1770
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   825
      Left            =   585
      Shape           =   4  'Rounded Rectangle
      Top             =   8685
      Width           =   16665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Top             =   3150
      Width           =   675
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hotel information"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   3555
      TabIndex        =   8
      Top             =   6300
      Width           =   2265
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Packages"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   3555
      TabIndex        =   7
      Top             =   1710
      Width           =   2100
   End
   Begin VB.Label LblSelect 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bus Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   3600
      TabIndex        =   6
      Top             =   5580
      Width           =   1305
   End
End
Attribute VB_Name = "inde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim r As New ADODB.Recordset
Private Sub bookin_Click()
If Option1.Value = True Then
Unload Me
book.Show

Exit Sub
End If
If Opt2.Value = True Then
Unload Me
CAR_BOOK.Show

End If

End Sub

Private Sub bookin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ToolTipText = " Ticket Booking menu"
A = ToolTipText
B1.Text = A
End Sub
Private Sub clos_Click()
End
End Sub
Private Sub clos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ToolTipText = "Exit from application"
A = ToolTipText
B1.Text = A
End Sub
Private Sub Combo1_Click()
Set r = New ADODB.Recordset
A = "select * from package where place='" & Combo1.Text & "'"
Set r = c.Execute(A)
Day.Text = r.Fields(3)
busn.Text = r.Fields(4)
Text10.Text = r.Fields(6)
Text5.Text = r.Fields(1)
Text2.Text = r.Fields(8)
Text1.Text = r.Fields(7)
hinfo.Text = r.Fields(2)
End Sub



Private Sub Combo2_Click()
Set r = New ADODB.Recordset
A = "select * from cpackage where place='" & Combo2.Text & "'"
Set r = c.Execute(A)
Text9.Text = r.Fields(0)
Text11.Text = r.Fields(6)
Text6.Text = r.Fields(2)
Text3.Text = r.Fields(7)
Text4.Text = r.Fields(4)
Text7.Text = r.Fields(3)
End Sub

Private Sub detail_Click()
Unload Me
cust.Show
End Sub
Private Sub detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ToolTipText = " Ticket Booking Details"
A = ToolTipText
B1.Text = A
End Sub
Private Sub Form_Load()
connect
A = "select place from package"
Set r = c.Execute(A)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select place from cpackage"
Set r = c.Execute(A)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
B1.Text = " "

End Sub
Private Sub id_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
id.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub opt2_Click()
Combo2.Enabled = True
Combo1.Text = ""
Text10.Text = ""
Text5.Text = ""
Day.Text = ""
Text1.Text = ""
Text2.Text = ""
busn.Text = ""
hinfo.Text = ""
'Combo1.Enabled = False
End Sub

Private Sub Option1_Click()
Combo1.Enabled = True
Combo2.Text = ""
Text11.Text = ""
Text3.Text = ""
Text4.Text = ""
Text9.Text = ""
Text6.Text = ""
Text7.Text = ""
'Combo2.Enabled = False
End Sub

Private Sub schedul_Click()
Unload Me
'schedu.Show
End Sub
Private Sub schedul_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ToolTipText = "Get Date and time of every package"
A = ToolTipText
B1.Text = A
End Sub
Private Sub SHOW1_LostFocus()
'B1.Text = ""
End Sub

Private Sub SHOW1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ToolTipText = "Show all records in Ticket Booking"
A = ToolTipText
B1.Text = A
End Sub

Private Sub SHOW1_Click()
MSFl.Visible = True

End Sub
