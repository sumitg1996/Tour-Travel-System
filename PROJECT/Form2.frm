VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form emp 
   Caption         =   "EMPLOYEE DETAILS"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   DrawWidth       =   2
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "REFRESH"
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
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8640
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000C&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3510
      Width           =   1410
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   690
      Left            =   -675
      Top             =   7110
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1217
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
      RecordSource    =   "select *from driver"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   -585
      Top             =   6660
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   661
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
      RecordSource    =   "select  *from employe3"
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
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Form2.frx":3474F
      Left            =   12420
      List            =   "Form2.frx":3475F
      Style           =   2  'Dropdown List
      TabIndex        =   17
      ToolTipText     =   "* Select information of employee"
      Top             =   2340
      Width           =   2760
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Form2.frx":34785
      Left            =   12420
      List            =   "Form2.frx":34787
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2925
      Width           =   2760
   End
   Begin VB.TextBox add 
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
      Height          =   615
      Left            =   5040
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Please Enter Address"
      Top             =   2565
      Width           =   3210
   End
   Begin VB.CommandButton add1 
      Height          =   555
      Left            =   3195
      Picture         =   "Form2.frx":34789
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8550
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton update1 
      Height          =   555
      Left            =   4770
      Picture         =   "Form2.frx":39AF7
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8550
      Width           =   1530
   End
   Begin VB.CommandButton exit 
      Height          =   555
      Left            =   6390
      Picture         =   "Form2.frx":3F178
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8550
      Width           =   1500
   End
   Begin VB.CommandButton delete1 
      Height          =   555
      Left            =   8010
      Picture         =   "Form2.frx":400C8
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8550
      Width           =   1500
   End
   Begin VB.CommandButton new1 
      Height          =   555
      Left            =   3195
      Picture         =   "Form2.frx":41151
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8550
      Width           =   1455
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      TabIndex        =   1
      Top             =   765
      Width           =   1905
   End
   Begin VB.TextBox licence 
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
      Height          =   435
      Left            =   5040
      MaxLength       =   15
      TabIndex        =   12
      ToolTipText     =   "Enter Licence No Or Aadhar No"
      Top             =   7425
      Width           =   3315
   End
   Begin VB.TextBox saler 
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
      Height          =   450
      Left            =   5040
      MaxLength       =   6
      TabIndex        =   11
      ToolTipText     =   "Enter Sallery Of Employee"
      Top             =   6840
      Width           =   3315
   End
   Begin VB.TextBox quali 
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
      Height          =   405
      Left            =   5040
      MaxLength       =   11
      TabIndex        =   9
      ToolTipText     =   "Enter Employee Qualification"
      Top             =   5715
      Width           =   3315
   End
   Begin VB.TextBox status 
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
      Height          =   465
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   8
      ToolTipText     =   "Enter Employee Designation"
      Top             =   5085
      Width           =   2235
   End
   Begin VB.TextBox conn 
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
      Height          =   435
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "Enter A valid Contact No"
      Top             =   3825
      Width           =   3225
   End
   Begin VB.TextBox email 
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
      Height          =   435
      Left            =   5040
      MaxLength       =   100
      TabIndex        =   5
      ToolTipText     =   "Please Enter A valid Email id"
      Top             =   3285
      Width           =   3240
   End
   Begin VB.TextBox name1 
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
      Height          =   435
      Left            =   5040
      MaxLength       =   20
      TabIndex        =   3
      ToolTipText     =   "Please Enter The Name"
      Top             =   2070
      Width           =   3210
   End
   Begin VB.TextBox Text1 
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
      Height          =   450
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1440
      Width           =   2505
   End
   Begin VB.ComboBox select1 
      DataSource      =   "Adodc1"
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
      Left            =   5040
      TabIndex        =   31
      Text            =   "Select Id"
      Top             =   1485
      Width           =   2505
   End
   Begin MSComCtl2.DTPicker BDate 
      Height          =   420
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Choose Date Of Birth"
      Top             =   4500
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   68878339
      CurrentDate     =   40185
   End
   Begin MSComCtl2.DTPicker JDate 
      Height          =   420
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   "Choose Joining Date"
      Top             =   6255
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   68878339
      CurrentDate     =   40259
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "Form2.frx":41FDC
      Height          =   4155
      Left            =   10260
      TabIndex        =   34
      Top             =   4410
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   7329
      _Version        =   393216
      BackColor       =   16744576
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "EID"
         Caption         =   "Employee Id"
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
         DataField       =   "NAME"
         Caption         =   "Name"
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
         DataField       =   "ADDR"
         Caption         =   "Address"
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
         DataField       =   "EMAIL_ID"
         Caption         =   "Eail id"
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
         DataField       =   "PHNO"
         Caption         =   "Contact"
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
         DataField       =   "DOB"
         Caption         =   "DOB"
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
         DataField       =   "STATUS"
         Caption         =   "Desig"
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
         DataField       =   "QUALI"
         Caption         =   "Qualification"
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
      BeginProperty Column08 
         DataField       =   "J_DATE"
         Caption         =   "J_DATE"
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
      BeginProperty Column09 
         DataField       =   "SAL"
         Caption         =   "Sallery"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   8520
      Left            =   1530
      Top             =   675
      Width           =   8610
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   5
      Height          =   8745
      Left            =   1440
      Top             =   585
      Width           =   8790
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1995
      Left            =   12060
      Top             =   2205
      Width           =   3435
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Employee"
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
      Height          =   525
      Left            =   2610
      TabIndex        =   32
      Top             =   675
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      Left            =   2970
      TabIndex        =   30
      Top             =   1485
      Width           =   1710
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3015
      TabIndex        =   29
      Top             =   2115
      Width           =   750
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3015
      TabIndex        =   28
      Top             =   2655
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   3015
      TabIndex        =   27
      Top             =   5130
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date"
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
      Left            =   3015
      TabIndex        =   26
      Top             =   4545
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Join date"
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
      Left            =   3015
      TabIndex        =   25
      Top             =   6300
      Width           =   1170
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
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
      Left            =   3015
      TabIndex        =   24
      Top             =   6885
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
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
      Left            =   3015
      TabIndex        =   23
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
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
      Left            =   3015
      TabIndex        =   22
      Top             =   5715
      Width           =   1695
   End
   Begin VB.Label lblselect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select  ID"
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
      Left            =   2970
      TabIndex        =   21
      Top             =   1485
      Width           =   1290
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail ID"
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
      Left            =   3015
      TabIndex        =   20
      Top             =   3375
      Width           =   1260
   End
   Begin VB.Label lblLIc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licence No"
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
      Left            =   3015
      TabIndex        =   0
      Top             =   7470
      Width           =   1515
   End
End
Attribute VB_Name = "emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As String
Dim con As New ADODB.Connection
Dim A As String
Dim s As String

Private Sub add1_Click()
If Message = True Then
Else
If Opt1.Value = True Then
'Text2.Visible = False
s = "insert into employe3 values ('" & Text1.Text & "','" & name1.Text & "','" & add.Text & "','" & email.Text & "'," & conn.Text & ",'" & Format(BDate.Value, "dd-mmm-yyyy") & "','" & status.Text & "','" & quali.Text & "','" & Format(JDate.Value, "dd-mmm-yyyy") & "'," & saler.Text & ",'" & licence.Text & "')"
MsgBox s
Set r = c.Execute(s)
MsgBox "Data added Succesfully", , "Add"
new1.Visible = True
select1.Visible = True
delete1.Enabled = True
UPDATE1.Enabled = True
End If
End If
End Sub

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0:
    s = "select eid from employe3 "
    Combo2.Clear
    Set r = c.Execute(s)
   While r.EOF <> True
   Combo2.AddItem r.Fields(0)
   r.MoveNext
   Wend
Case 1:
s = "select name  from employe3 "
 Combo2.Clear
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
Case 2:
'  Combo1.Text = "Salary"
    s = "select sal  from employe3 "
     Combo2.Clear
'    Combo2.Text = " select salary "
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
Case 3:
'Combo1.Text = "desig"
    s = "select DISTINCT status  from employe3 "
     Combo2.Clear
'    Combo2.Text = " select desig "
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
End Select
End Sub

Private Sub Combo2_Click()
connect
If Combo1.Text = "Employee Id" Then
Adodc1.RecordSource = "SELECT *FROM employe3 WHERE EID ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
ElseIf Combo1.Text = "Name" Then
Adodc1.RecordSource = "SELECT *FROM employe3 WHERE name='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
ElseIf Combo1.Text = "Salary" Then
Adodc1.RecordSource = "SELECT *FROM employe3 WHERE sal ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
Else

Adodc1.RecordSource = "SELECT *FROM employe3 WHERE status ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub Command1_Click()
Unload Me
emp.Show
End Sub



Private Sub Command3_Click()
connect
DataEnvironment1.rsCommand1.Open " SELECT * FROM employe3 WHERE EID   ='" & Combo2.Text & "' or name = '" & Combo2.Text & "'or sal = '" & Combo2.Text & "'OR STATUS ='" & Combo2.Text & "'"
'employee.Refresh
employee.Show
DataEnvironment1.rsCommand1.close
End Sub

Private Sub conn_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
BDate.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub delete1_Click()
If Opt1.Value = True Then
s = "delete from Employe3 where EID='" & select1.Text & "'"
Set r = c.Execute(s)
MsgBox "Data Deleted Succesfully", , "Delete"
name1.Text = " "
add.Text = " "
email.Text = " "
 conn.Text = " "
 status.Text = " "
 quali.Text = " "
 saler.Text = " "
 licence.Text = ""
End If
Unload Me
emp.Show
End Sub

Private Sub exit_Click()
Unload Me
Main.Show
End Sub

Private Sub Form_Load()
Label2.Visible = False
Text1.Visible = False
connect
A = "select distinct eid from employe3"
Set r = c.Execute(A)
While r.EOF <> True
select1.AddItem r.Fields(0)
r.MoveNext
Wend
A = "emp"
B = "driver"
s = "select count(eid) from employe3"
Set r = c.Execute(s)
Text1.Text = UCase(A & r.Fields(0) + 1)

End Sub

Private Sub new1_Click()
enb
Clear
A = "emp"
'B = "driver"
name1.SetFocus
Label2.Visible = True
Text1.Visible = True
lblselect.Visible = False
'select2.Visible = False
select1.Visible = False
add1.Visible = True
'If Opt1.Value = False And Opt2.Value = False Then
'MsgBox "Please Select Category from Driver or Employee"

If Opt1.Value = True Then
select1.Visible = False
End If
'If Opt2.Value = True Then
'select2.Visible = False
'Text2.Visible = True
licence.Text = ""

new1.Visible = False
add.Visible = True
Exit Sub

End Sub

Function enb()
 Text1.Locked = False
 name1.Locked = False
 add.Locked = False
 conn.Locked = False
 status.Locked = False
 quali.Locked = False
End Function

Function Clear() As Boolean
name1 = ""
add = ""
conn = ""
email = ""
status = ""
quali = ""
saler = ""
End Function

Private Sub BDateDTP_Change()
'BDate.Text = BDateDTP
End Sub

Private Sub saler_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
add1.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub select1_Click()
Set r = New ADODB.Recordset
A = "select * from employe3 where eid='" & select1.Text & "'"
MsgBox A
Set r = c.Execute(A)
With r
name1.Text = .Fields(1)
add.Text = .Fields(2)
email.Text = .Fields(3)
 conn.Text = .Fields(4)
 BDate.Value = .Fields(5)
 status.Text = .Fields(6)
 quali.Text = .Fields(7)
 JDate.Value = .Fields(8)
 saler.Text = .Fields(9)
 licence.Text = .Fields(10)
End With
End Sub

Private Sub UPDATE1_Click()
connect
s = "update employe3 set  name='" & name1.Text & "',addr='" & add.Text & "',email_id='" & email.Text & "',phno=" & conn.Text & ",dob='" & Format(BDate.Value, "dd-mmm-yyyy") & "',status='" & status.Text & "',quali='" & quali.Text & "',j_date='" & Format(JDate.Value, "dd-mmm-yyyy") & "',sal=" & saler.Text & "where eid = '" & select1.Text & "'"
Set r = c.Execute(s)
MsgBox "Data Updated Succesfully", , "Update"
Clear

End Sub
Private Sub opt1_Click()
Clear
Grid1.Visible = True
If Opt1.Value = True Then
UPDATE1.Visible = True
lblselect.Caption = "Select Employee ID"
select1.Visible = True
End If
End Sub

Private Sub JDateDTP_Change()
'JDate.Text = JDateDTP
End Sub

Function Message() As Boolean
   If name1.Text = "" Then
       MsgBox "Please Enter Name", vbCritical, "TOUR & TRAVEL"
       name1.SetFocus
       Message = True
    ElseIf add.Text = "" Then
       MsgBox "Please Enter Address", vbCritical, "TOUR & TRAVEL"
       add.SetFocus
       Message = True
    ElseIf email.Text = "" Then
       MsgBox "Please Enter Email ID", vbCritical, "TOUR & TRAVEL"
       email.SetFocus
       Message = True
    ElseIf conn.Text = "" Then
       MsgBox "Please Enter Contact", vbCritical, "TOUR & TRAVEL"
       conn.SetFocus
       Message = True
    ElseIf status.Text = "" Then
       MsgBox "Please Enter Status ", vbCritical, "TOUR & TRAVEL"
       status.SetFocus
       Message = True
    ElseIf quali.Text = "" Then
       MsgBox "Please Enter Qualification", vbCritical, "TOUR & TRAVEL"
       quali.SetFocus
       Message = True
    ElseIf JDate.Value = "" Then
       MsgBox "Please Enter joining date", vbCritical, "TOUR & TRAVEL"
       JDate.SetFocus
       Message = True
    ElseIf saler.Text = "" Then
       MsgBox "please Enter Salary", vbCritical, "TOUR & TRAVEL"
       saler.SetFocus
       Message = True
       End If
End Function

