VERSION 5.00
Begin VB.Form pass 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORGET PASSWORD"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   269.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   529.5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txtpass 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1035
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton close 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "EXIT"
      Height          =   480
      Left            =   4635
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4635
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5370
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10590
      Begin VB.TextBox usern 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   4950
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2160
         Width           =   2835
      End
      Begin VB.CommandButton update 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "UPDATE"
         Height          =   465
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4635
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox confpass 
         Appearance      =   0  'Flat
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4935
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   3945
         Width           =   2820
      End
      Begin VB.TextBox npass 
         Appearance      =   0  'Flat
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4950
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3330
         Width           =   2820
      End
      Begin VB.TextBox opass 
         Appearance      =   0  'Flat
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4950
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2745
         Width           =   2820
      End
      Begin VB.ComboBox Cmduser 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1665
         Width           =   2805
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "************* if you don't know old password then concern to admin **************"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   45
         TabIndex        =   14
         Top             =   540
         Width           =   10365
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Left            =   1890
         TabIndex        =   13
         Top             =   2250
         Width           =   1905
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Conform Password"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Left            =   1905
         TabIndex        =   12
         Top             =   4005
         Width           =   2745
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Left            =   1905
         TabIndex        =   11
         Top             =   3420
         Width           =   2205
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         Left            =   1905
         TabIndex        =   10
         Top             =   2835
         Width           =   2205
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Username"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   1890
         TabIndex        =   9
         Top             =   1665
         Width           =   2715
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   1395
      Top             =   5715
   End
End
Attribute VB_Name = "pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim c As New ADODB.Connection
'Dim WithEvents adoPrimaryRS As ADODB.Recordset
'Dim r As New ADODB.Recordset
'Dim passrs As New ADODB.Recordset
Dim r As New ADODB.Recordset
Dim s As String
Public Sub X()
Cmduser.SetFocus
usern.Text = ""
npass.Text = ""
opass.Text = ""
confpass.Text = ""
'Cmduser.Text = ""
pass.Show
Txtpass.Text = ""
End Sub

Private Sub close_Click()
Unload MDI
Unload Me
frmLogin.Show
End Sub

Private Sub Cmduser_Click()
'With r 'passrs
UPDATE.Visible = True
s = "Select *from login where User_nm='" & Cmduser.Text & "'"
Set r = c.Execute(s)
If Not r.EOF Then
usern = r.Fields(1)
Txtpass = r.Fields(2)
End If
'r.update
r.close

End Sub

Private Sub confpass_LostFocus()
If confpass.Text = npass.Text Then
Else
MsgBox "Reenter conform Password", vbCritical, "Tour & Travels"
confpass.Text = ""
confpass.SetFocus
End If
End Sub

Private Sub Form_Load()
connect
s = "select user_nm,pass from login"
Set r = c.Execute(s)
Do While Not r.EOF
Cmduser.AddItem r.Fields(0)
r.MoveNext
Loop
Label2.Caption = Label2.Caption & Space(30)
End Sub

Private Sub opass_LostFocus()
If opass.Text = Txtpass.Text Then
Else
MsgBox "Invalid Password", vbCritical, "Tour & Travels"
opass.Text = ""
opass.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Dim str As String
str = pass.Label2.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
pass.Label2.Caption = str
End Sub

Private Sub update_Click()
If confpass.Text = "" And opass.Text = "" Then
MsgBox "Enter old Password & Then new password", vbCritical, "Tour & Travels"
opass.SetFocus
Else
connect
s = "update  login set pass = '" + confpass.Text + "'where user_nm= '" + Cmduser.Text + "'"
Set r = c.Execute(s)
MsgBox "Password Update", vbOKOnly, "Tour & Travels"
End If
X
End Sub
Private Sub txtUserName_LostFocus()
s = "Select *From login where User_nm='" & Txtusername.Text & "'"
If r.RecordCount = 0 Then
MsgBox "Invalid username"
Txtusername.SetFocus
SendKeys "{Home}+{End}"
Else
Txtpass = r.Fields(1)
End If
r.close

End Sub
