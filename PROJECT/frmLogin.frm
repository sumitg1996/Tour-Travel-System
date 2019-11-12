VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOGIN"
   ClientHeight    =   3570
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2109.273
   ScaleMode       =   0  'User
   ScaleWidth      =   7351.947
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txtpass 
      DataField       =   "USER_ID"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3720
      Width           =   2325
   End
   Begin VB.TextBox Txtusername 
      DataField       =   "USER_ID"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2910
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1200
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4050
      Picture         =   "frmLogin.frx":4818
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2205
      Width           =   1170
   End
   Begin VB.TextBox txtpassword 
      DataField       =   "PASS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2910
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1710
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password ?"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      MousePointer    =   4  'Icon
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   0
      Left            =   1590
      TabIndex        =   0
      Top             =   1230
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   1
      Left            =   1590
      TabIndex        =   5
      Top             =   1740
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim p As String
Dim q As String

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
connect
s = "select user_nm ,pass from login"
Set r = c.Execute(s)
q = r.Fields(0)
p = r.Fields(1)
r.MoveNext
If errMSG = True Then
Exit Sub
Else
 If Txtusername.Text = "admin" And txtpassword.Text = Txtpass Then
        LoginSucceeded = True
        Unload Me
        MDI.Show
        Main.Show
        pass.Txtpass.Visible = True
    ElseIf Txtusername.Text = "employee" And txtpassword.Text = Txtpass Then
        LoginSucceeded = True
        Unload Me
        MDI.Show
        Main.Show
       Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtpassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End If
End Sub
Function errMSG() As Boolean
  If Txtusername.Text = "" Then
  MsgBox "Enter UserName"
  Txtusername.SetFocus
  errMSG = True
  ElseIf txtpassword.Text = "" Then
  MsgBox "Enter Password", vbOKOnly, "TOUR & TRAVEL"
  txtpassword.SetFocus
  errMSG = True
  End If
  End Function

Private Sub Form_Load()
connect
s = "Select *from login"
Set r = c.Execute(s)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlack
End Sub

Private Sub Label1_Click()
Unload Me
pass.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbRed
End Sub

Private Sub txtUserName_LostFocus()
On Error GoTo abc
s = "Select *From login where User_nm='" & Txtusername.Text & "'"
Set r = c.Execute(s)
If r.RecordCount = 0 Then
MsgBox "Invalid username", vbOKOnly, "TOUR & TRAVEL"
Txtusername.SetFocus
SendKeys "{Home}+{End}"
Else
Txtpass.Text = r.Fields(2)
End If
r.close
Exit Sub
abc:
MsgBox "invalid username"
Txtusername.Text = ""
Txtusername.SetFocus
End Sub

