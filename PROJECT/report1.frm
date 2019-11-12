VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form report1 
   Caption         =   "REPORT"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "report1.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   18075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BackColor       =   &H00808080&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7290
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00808080&
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7290
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00808080&
      Caption         =   "HOTEL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7290
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00808080&
      Caption         =   "BOOKING"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7290
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "CAR"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7290
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3105
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   645
      Left            =   630
      Top             =   5940
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1138
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
      Connect         =   "Provider=MSDAORA.1;User ID=tour/travel;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=tour/travel;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from cpackage where 1>2"
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
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "report1.frx":114C58
      Left            =   7065
      List            =   "report1.frx":114C62
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2025
      Width           =   2760
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "report1.frx":114C79
      Left            =   7065
      List            =   "report1.frx":114C7B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2610
      Width           =   2760
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   510
      Left            =   495
      Top             =   4725
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   900
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
      Connect         =   "Provider=MSDAORA.1;User ID=tour/travel;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=tour/travel;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from package WHERE 1>2"
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "report1.frx":114C7D
      Height          =   2760
      Left            =   4815
      TabIndex        =   4
      Top             =   3825
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4868
      _Version        =   393216
      BackColor       =   8421504
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DAYS"
         Caption         =   "DAYS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PLACE"
         Caption         =   "PLACE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "HNAME"
         Caption         =   "HNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "VEHICLE"
         Caption         =   "VEHICLE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "NO"
         Caption         =   "NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "TOURC"
         Caption         =   "TOURC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "PACKC"
         Caption         =   "PACKC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   4920
      Left            =   4725
      Top             =   1755
      Width           =   7125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTS PACKAGE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7110
      TabIndex        =   0
      Top             =   540
      Width           =   5460
   End
End
Attribute VB_Name = "report1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
connect
Select Case Combo1.ListIndex
Case 0:
    s = "select distinct id from package "
    Combo2.Clear
    Set r = c.Execute(s)
   While r.EOF <> True
   Combo2.AddItem r.Fields(0)
   r.MoveNext
   Wend
Case 1:
  Combo1.Text = "place"
s = "select distinct place  from package "
 Combo2.Clear
    Combo2.Text = " select place "
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
End Select
End Sub

Private Sub Combo2_Click()
connect
Adodc1.RecordSource = "SELECT *FROM package WHERE ID ='" & Combo2.Text & "'OR place='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource

End Sub

Private Sub Combo4_Click()
connect
Select Case Combo4.ListIndex
Case 0:
    s = "select distinct id from cpackage "
    Combo3.Clear
    Set r = c.Execute(s)
   While r.EOF <> True
   Combo3.AddItem r.Fields(0)
   r.MoveNext
   Wend
Case 1:
  'Combo1.Text = "place"
s = "select distinct place  from cpackage "
 Combo3.Clear
'    Combo3.Text = " select place "
Set r = c.Execute(s)
While r.EOF <> True
Combo3.AddItem r.Fields(0)
r.MoveNext
Wend
End Select
End Sub

Private Sub Combo3_Click()
connect
Adodc2.RecordSource = "SELECT *FROM cpackage WHERE ID ='" & Combo3.Text & "' OR place='" & Combo3.Text & "' "
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource
End Sub



Private Sub Command2_Click()
connect
If DataEnvironment1.rsCPACKAGE.State = 1 Then DataEnvironment1.rsCPACKAGE.close
DataEnvironment1.rsCPACKAGE.Open " SELECT * FROM CPACKAGE WHERE ID ='" & Combo3.Text & "' or PLACE = '" & Combo3.Text & "'"  ' or sal = '" & Combo3.Text & "'"
CPACKAGE.Refresh
CPACKAGE.Show
DataEnvironment1.rsCPACKAGE.close
End Sub

Private Sub Command3_Click()
bus.Show
End Sub

Private Sub Command4_Click()
carS.Show
End Sub

Private Sub Command5_Click()
Busbook.Show
End Sub

Private Sub Command6_Click()
carbooking.Show
End Sub

Private Sub Command7_Click()
hotel.Show
End Sub

Private Sub Command8_Click()
customer3.Show
End Sub

Private Sub Command9_Click()
CANCE.Show
End Sub
