VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form newindex 
   BackColor       =   &H00DDD3D2&
   Caption         =   "INDEX "
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "newindex.frx":0000
   ScaleHeight     =   10890
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1890
      Top             =   1845
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1125
      Top             =   2025
   End
   Begin VB.Timer Timer1 
      Interval        =   120
      Left            =   315
      Top             =   2745
   End
   Begin VB.TextBox addr 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   9945
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5445
      Width           =   3390
   End
   Begin VB.TextBox EMAIL2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   14895
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4950
      Width           =   3360
   End
   Begin VB.TextBox PH2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   14895
      MaxLength       =   10
      TabIndex        =   7
      Top             =   4455
      Width           =   3360
   End
   Begin VB.TextBox NAME2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11070
      MaxLength       =   20
      TabIndex        =   8
      Top             =   4905
      Width           =   2280
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   150
      Left            =   7245
      TabIndex        =   5
      Top             =   5130
      Width           =   195
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   150
      Left            =   5670
      TabIndex        =   4
      Top             =   5130
      Width           =   195
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5265
      MaxLength       =   4
      TabIndex        =   18
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6795
      MaxLength       =   4
      TabIndex        =   19
      Top             =   4995
      Width           =   375
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   360
      ItemData        =   "newindex.frx":246FF
      Left            =   12600
      List            =   "newindex.frx":24701
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1845
      Width           =   2670
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   360
      ItemData        =   "newindex.frx":24703
      Left            =   9810
      List            =   "newindex.frx":24705
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1845
      Width           =   2670
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   360
      ItemData        =   "newindex.frx":24707
      Left            =   7020
      List            =   "newindex.frx":24709
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1845
      Width           =   2670
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   360
      ItemData        =   "newindex.frx":2470B
      Left            =   4230
      List            =   "newindex.frx":24712
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1845
      Width           =   2670
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   45
      TabIndex        =   12
      Text            =   "Avail"
      Top             =   495
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Car Booking"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   14175
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   1770
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   -180
      Top             =   7335
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=TOUR/TRAVEL;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=TOUR/TRAVEL;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"newindex.frx":24727
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   465
      Left            =   1260
      TabIndex        =   6
      Top             =   990
      Width           =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   16065
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1680
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "***************** TOUR AND TRAVEL SYSTEM********************"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   24
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   180
      TabIndex        =   39
      Top             =   180
      Width           =   19905
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   4785
      Left            =   1530
      Shape           =   4  'Rounded Rectangle
      Top             =   2925
      Width           =   6900
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Package Details"
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
      Left            =   1620
      TabIndex        =   38
      Top             =   3195
      Width           =   6765
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Details"
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
      Left            =   8640
      TabIndex        =   37
      Top             =   3285
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   4785
      Left            =   8505
      Shape           =   4  'Rounded Rectangle
      Top             =   2925
      Width           =   10320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PACKAGE CHARGES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2070
      TabIndex        =   36
      Top             =   6345
      Width           =   3030
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  CHARGES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2070
      TabIndex        =   35
      Top             =   6885
      Width           =   3435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TOUR CHARGES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2070
      TabIndex        =   34
      Top             =   5670
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHARGES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2070
      TabIndex        =   33
      Top             =   5085
      Width           =   1635
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11070
      TabIndex        =   32
      Top             =   4455
      Width           =   2265
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   31
      Top             =   6795
      Width           =   1725
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   30
      Top             =   6255
      Width           =   1725
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   29
      Top             =   5670
      Width           =   1725
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5940
      TabIndex        =   28
      Top             =   4995
      Width           =   825
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4410
      TabIndex        =   27
      Top             =   5040
      Width           =   825
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5625
      TabIndex        =   26
      Top             =   3870
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   135
      Top             =   1485
      Width           =   510
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   4590
      TabIndex        =   25
      Top             =   3915
      Width           =   1230
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer id"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8640
      TabIndex        =   24
      Top             =   4455
      Width           =   1545
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email id"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13590
      TabIndex        =   23
      Top             =   4995
      Width           =   1080
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone no"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13590
      TabIndex        =   22
      Top             =   4500
      Width           =   1200
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8595
      TabIndex        =   21
      Top             =   5625
      Width           =   1035
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8640
      TabIndex        =   20
      Top             =   4950
      Width           =   750
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Per/Day"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   420
      Left            =   6750
      TabIndex        =   17
      Top             =   6795
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DOUBLE BED"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5850
      TabIndex        =   16
      Top             =   4545
      Width           =   1230
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SINGLE BED"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4365
      TabIndex        =   15
      Top             =   4590
      Width           =   1230
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Per package"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   420
      Left            =   6750
      TabIndex        =   13
      Top             =   6795
      Width           =   1320
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   960
      Left            =   90
      Top             =   0
      Width           =   20310
   End
End
Attribute VB_Name = "newindex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String

Private Sub addr_LostFocus()
addr.Text = UCase(addr.Text)
End Sub

Private Sub Combo1_Click()
Combo2.Clear
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
connect
 If Combo1.Text = "PRIVATE PACKAGE" Then
 Command2.Visible = True
 Label9.Visible = True
 Label4.Visible = False
 A = "select distinct place from cpackage"
 Set r = c.Execute(A)
 While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
End If
Exit Sub
End Sub

Private Sub Combo2_Click()
connect
Combo3.Clear
If Combo1.Text = "PRIVATE PACKAGE" Then
s = "SELECT h_name FROM HOTEL WHERE PLACE = '" & Combo2.Text & "'"
Set r = c.Execute(s)
While r.EOF <> True
Combo3.AddItem r.Fields(0)
r.MoveNext
Wend
Exit Sub
End If
End Sub

Private Sub Combo2_LostFocus()
On Error GoTo A:
connect
s = "SELECT  DAYS Cpackage WHERE PLACE = '" & Combo2.Text & "'"
Set r = c.Execute(s)
Label11.Caption = r.Fields(0)
Text6.Text = Text7.Text
Text5.Text = Text7.Text

A:
End Sub

Private Sub Combo3_Click()
Combo4.Clear
s = "SELECT *FROM hotel WHERE h_name= '" & Combo3.Text & "'"
Set r = c.Execute(s)
While r.EOF <> True
Combo4.AddItem r.Fields(3)
Combo4.AddItem r.Fields(4)
r.MoveNext
Wend
End Sub

Private Sub Combo4_Click()
Select Case Combo4.ListIndex
Case 0:
s = "SELECT *FROM hotel WHERE AC= '" & Combo4.Text & "'"
Set r = c.Execute(s)
Label12.Caption = r.Fields(5)
Label13.Caption = r.Fields(6)
Case 1:
s = "SELECT *FROM hotel WHERE NAC= '" & Combo4.Text & "'"
Set r = c.Execute(s)
Label12.Caption = r.Fields(7)
Label13.Caption = r.Fields(8)
 End Select
End Sub

Private Sub Combo4_LostFocus()
 If Combo1.Text = "PRIVATE PACKAGE" Then
 s = "SELECT tourc,packc,days FROM CPACKAGE WHERE hname= '" & Combo3.Text & "'"
 Set r = c.Execute(s)
 If r.EOF <> True Or r.BOF <> True Then
  Label14.Caption = r.Fields(0)
  Label15.Caption = r.Fields(1)
  Label11.Caption = r.Fields(2)
  Text6.Text = Label11.Caption
  Text5.Text = Label11.Caption
  End If
   End If
 Label4.Caption = Format(Label14.Caption, "000.00")
   Label15.Caption = Format(Label15.Caption, "000.00")
End Sub


Private Sub Command1_Click()
Unload Me
newindex.Show
End Sub

Private Sub Command2_Click()
If Message = True Then
Exit Sub
Else
connect
CAR_BOOK.TC2.Text = newindex.Label14.Caption
CAR_BOOK.Text10.Text = newindex.Label17.Caption
CAR_BOOK.TOC2.Text = newindex.Label16.Caption
CAR_BOOK.packa.Text = newindex.Combo2.Text
CAR_BOOK.PC.Text = newindex.Label15.Caption
CAR_BOOK.Show
End If
End Sub

Private Sub Command3_Click()
If Message = True Then
Exit Sub
Else
book.Text9.Text = newindex.Label17.Caption
book.extra.Text = newindex.Label16.Caption
book.packa.Text = newindex.Combo2.Text
book.Show
End If
End Sub



Private Sub Form_Load()
Label6.Caption = Label6.Caption & Space(40)
Label19.Caption = Label19.Caption & Space(60)
Label18.Caption = Label18.Caption & Space(100)
Timer1.Enabled = True
'DTPicker1.Value = Date
connect
B = "CUST00"
s = "select count(CID) from CUSTOMER"
Set r = c.Execute(s)
Label17.Caption = B & r.Fields(0) + 1


End Sub

Function Message() As Boolean
If NAME2.Text = "" Then
MsgBox "Please Enter Name ", vbCritical, "TOUR & TRVEL"
NAME2.SetFocus
Message = True
ElseIf addr.Text = "" Then
MsgBox "Please Enter Address", vbCritical, "TOUR & TRVEL"
addr.SetFocus
Message = True
ElseIf PH2.Text = "" Then
MsgBox "Please Enter Contact No", vbCritical, "TOUR & TRVEL"
PH2.SetFocus
Message = True
ElseIf EMAIL2.Text = "" Then
MsgBox "Please Enter Email Id", vbCritical, "TOUR & TRVEL"
EMAIL2.SetFocus
Message = True
End If
End Function

Private Sub Label12_Change()
Label12.Caption = Format(Label12.Caption, "0.00")
End Sub

Private Sub Label13_Change()
Label13.Caption = Format(Label13.Caption, "0.00")
End Sub

Private Sub Label14_Change()
Label14.Caption = Format(Label14.Caption, "0.00")
End Sub

Private Sub name2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 And KeyAscii <> 8 And KeyAscii <> 32 Then
MsgBox " please enter only A to Z", vbCritical, "Tour & Travels"
KeyAscii = 0
NAME2.SetFocus
Else
End If
If KeyAscii = 13 Then
End If
End Sub

Private Sub name2_LostFocus()
NAME2.Text = UCase(NAME2.Text)
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Label16.Caption = "0"
Label16.Caption = Val(Label14.Caption) + Val(Label15.Caption) + (Val(Label16.Caption) + Val(Label12.Caption) * Val(Text6.Text))
Label16.Caption = Format(Label16.Caption, "000.00")
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label16.Caption = "0"
Label16.Caption = Val(Label14.Caption) + Val(Label15.Caption) + Val(Label16.Caption) + Val(Label13.Caption) * Val(Text6.Text)
Label16.Caption = Format(Label16.Caption, "000.00")
End If
End Sub

Private Sub PH2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
EMAIL2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub Timer1_Timer()
Dim str As String
str = newindex.Label6.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
newindex.Label6.Caption = str
End Sub

Private Sub Timer2_Timer()
Dim str As String
str = newindex.Label19.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
newindex.Label19.Caption = str
End Sub

Private Sub Timer3_Timer()
Dim str As String
str = newindex.Label18.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
newindex.Label18.Caption = str
End Sub
