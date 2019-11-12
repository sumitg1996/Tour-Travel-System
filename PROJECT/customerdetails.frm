VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cust 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Customer details"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "customerdetails.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
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
      Height          =   510
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5355
      Width           =   1545
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "CAR BOOKING"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1980
      TabIndex        =   10
      Top             =   5490
      Width           =   195
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00808080&
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
      Left            =   5175
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5400
      Width           =   2760
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00808080&
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
      ItemData        =   "customerdetails.frx":B1787
      Left            =   2340
      List            =   "customerdetails.frx":B1794
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5400
      Width           =   2715
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00808080&
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
      ItemData        =   "customerdetails.frx":B17B5
      Left            =   2070
      List            =   "customerdetails.frx":B17BF
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1260
      Width           =   2715
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00808080&
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
      Left            =   4905
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1260
      Width           =   2760
   End
   Begin VB.CommandButton Command4 
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
      Height          =   510
      Left            =   7785
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1215
      Width           =   1545
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000007&
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
      Height          =   195
      Left            =   1620
      TabIndex        =   1
      Top             =   1350
      Width           =   195
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   465
      Left            =   45
      Top             =   3420
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   820
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
      RecordSource    =   "select *from cbook"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   510
      Left            =   45
      Top             =   4050
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
      Connect         =   "Provider=MSDAORA.1;Password=travel;User ID=tour;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=travel;User ID=tour;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT *FROM Customer"
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
   Begin MSDataGridLib.DataGrid Grid2 
      Bindings        =   "customerdetails.frx":B17D6
      Height          =   3525
      Left            =   3780
      TabIndex        =   3
      Top             =   1755
      Visible         =   0   'False
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   6218
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   32768
      ForeColor       =   12632256
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CID"
         Caption         =   "CUSTOMER ID"
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
         Caption         =   "NAME"
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
         Caption         =   "ADDRESS"
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
         DataField       =   "PHNO"
         Caption         =   "CONTACT NO"
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
         DataField       =   "EMAIL_ID"
         Caption         =   "EMAIL_ID"
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
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton exit 
      Height          =   555
      Left            =   17730
      Picture         =   "customerdetails.frx":B17EB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   405
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   510
      Left            =   45
      Top             =   4635
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
      Connect         =   "Provider=MSDAORA.1;Password=travel;User ID=tour;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=travel;User ID=tour;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT *FROM BOOK"
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
   Begin MSDataGridLib.DataGrid Grid3 
      Bindings        =   "customerdetails.frx":B2679
      Height          =   3870
      Left            =   405
      TabIndex        =   4
      Top             =   5985
      Width           =   19500
      _ExtentX        =   34396
      _ExtentY        =   6826
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   -2147483647
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
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
      ColumnCount     =   23
      BeginProperty Column00 
         DataField       =   "BID"
         Caption         =   "BOOKING ID"
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
         DataField       =   "CID"
         Caption         =   "CUSTOMER ID"
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
         DataField       =   "ID"
         Caption         =   "PACKAGE"
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
         DataField       =   "DDATE"
         Caption         =   "DDATE"
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
         DataField       =   "RDATE"
         Caption         =   "RDATE"
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
         DataField       =   "DTIME"
         Caption         =   "DTIME"
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
         Caption         =   "STATUS"
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
         DataField       =   "NO"
         Caption         =   "CAR NO"
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
         DataField       =   "NP"
         Caption         =   "NO PERSON"
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
         DataField       =   "BOARD"
         Caption         =   "BOARDING"
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
      BeginProperty Column10 
         DataField       =   "TC"
         Caption         =   "TOUR CHARGE"
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
      BeginProperty Column11 
         DataField       =   "PC"
         Caption         =   "P CHARGES"
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
      BeginProperty Column12 
         DataField       =   "DISC"
         Caption         =   "DISCOUNT"
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
      BeginProperty Column13 
         DataField       =   "TOTAL_C"
         Caption         =   "TOTAL_C"
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
      BeginProperty Column14 
         DataField       =   "ADV"
         Caption         =   "ADVANCE"
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
      BeginProperty Column15 
         DataField       =   "DUE"
         Caption         =   "DUE AMOUNT"
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
      BeginProperty Column16 
         DataField       =   "CASH"
         Caption         =   "CASH"
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
      BeginProperty Column17 
         DataField       =   "BANK"
         Caption         =   "BANK"
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
      BeginProperty Column18 
         DataField       =   "CHEQUE"
         Caption         =   "CHEQUE"
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
      BeginProperty Column19 
         DataField       =   "BANK_N"
         Caption         =   "BANK_N"
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
      BeginProperty Column20 
         DataField       =   "BRANCH"
         Caption         =   "BRANCH"
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
      BeginProperty Column21 
         DataField       =   "PAMOUNT"
         Caption         =   "PAY AMOUNT"
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
      BeginProperty Column22 
         DataField       =   "DATE1"
         Caption         =   "DATE1"
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
         SizeMode        =   1
         BeginProperty Column00 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   " Car Booking"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   420
      Left            =   450
      TabIndex        =   12
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Customer 
      BackColor       =   &H80000007&
      Caption         =   " Customer"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   420
      Left            =   360
      TabIndex        =   5
      Top             =   1260
      Width           =   1635
   End
End
Attribute VB_Name = "cust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As String
Dim customer1 As String
Dim customer2 As String


Public Function TicketsGrid()
MSFGCAR.Text = ""
s = "SELECT *FROM CBOOK"
Set r = c.Execute(s)
'Do While Not r.EOF
MSFGCAR.Rows = 1

MSFGCAR.AddItem r.Fields(0) & vbTab & r.Fields(2) & vbTab & r.Fields(4) & vbTab & r.Fields(1) & vbTab & r.Fields(3) & vbTab & r.Fields(6) & vbTab & r.Fields(8) & vbTab & r.Fields(9) & vbTab & r.Fields(14) & vbTab & r.Fields(15)
r.MoveNext
'Loop
End Function

Public Function totalsGrid()
MSFGCAR.Text = ""
s = "SELECT *FROM BOOK"
Set r = c.Execute(s)
Do While Not r.EOF
MSFGCAR.Rows = 1

MSFGCAR.AddItem r.Fields(0) & vbTab & r.Fields(2) & vbTab & r.Fields(4) & vbTab & r.Fields(1) & vbTab & r.Fields(3) & vbTab & r.Fields(6) & vbTab & r.Fields(8) & vbTab & r.Fields(9) & vbTab & r.Fields(14) & vbTab & r.Fields(15)
r.MoveNext
Loop
End Function
Function nosel() As Boolean
If SelectType.Text = "Select Data" Then
MsgBox "First Select Type"
nosel = True
SelectType.SetFocus
End If
End Function



Private Sub cmbcustomerid_Click()
Set r = New ADODB.Recordset
A = "select * from cbook where cid='" & cmbcustomerid.Text & "'"
MsgBox A
Set r = c.Execute(A)
With r
txtcustid.Text = .Fields(0)
txtcustname.Text = .Fields(2)
txtcustadd.Text = .Fields(3)
txtcusttelephone.Text = .Fields(4)
txtcustemailid.Text = .Fields(5)
txtPackage.Text = .Fields(1)
'txtBusName = !Bus_Name
txtBusNo.Text = .Fields(8)
'txtSeats. = !Seat_numbers
txtDeptdate.Text = .Fields(6)
txtDeptTime.Text = .Fields(7)
'txtNoSeats.Text  = r.Fields (
txtFares.Text = .Fields(15)
txtpiadAmount.Text = .Fields(23)
txtChargesPKM = "Not Applicable"
'txtReturnDate = !Return_date
txtBoarding.Text = .Fields(12)
txtCash.Text = .Fields(19)
txtBank.Text = .Fields(20)
txtchequeNo.Text = .Fields(21)
txtBranchName.Text = .Fields(23)
txtBankName.Text = .Fields(22)
End With
End Sub

Private Sub cmbCustomerID2_Click()
Set r = New ADODB.Recordset
A = "select * from book where cid='" & cmbCustomerID2.Text & "'"
MsgBox A
Set r = c.Execute(A)
With r
txtcustid.Text = .Fields(0)
txtcustname.Text = .Fields(2)
txtcustadd.Text = .Fields(3)
txtcusttelephone.Text = .Fields(4)
txtcustemailid.Text = .Fields(5)
txtPackage.Text = .Fields(1)
'txtBusName = !Bus_Name
txtBusNo.Text = .Fields(8)
'txtSeats. = !Seat_numbers
txtDeptdate.Text = .Fields(6)
txtDeptTime.Text = .Fields(7)
'txtNoSeats.Text  = r.Fields (
txtFares.Text = .Fields(18)
txtpiadAmount.Text = .Fields(26)
txtChargesPKM = "Not Applicable"
'txtReturnDate = !Return_date
txtBoarding.Text = .Fields(13)
txtCash.Text = .Fields(21)
txtBank.Text = .Fields(22)
txtchequeNo.Text = .Fields(23)
txtBranchName.Text = .Fields(25)
txtBankName.Text = .Fields(24)
End With
End Sub

Private Sub SelectType_CLICK()
If SelectType.Text = "Car tickets" Then
carTickets.Visible = True
'MSFGCAR.Visible = False
Txtsearch2.Visible = False
CMDSEARCH2.Visible = False
CMDsearch.Visible = True
txtSearch.Visible = True
End If
If SelectType.Text = "Bus Tickets" Then
busticketb.Visible = True
carTickets.Visible = False
busticketb.Visible = True
Txtsearch2.Visible = True
CMDSEARCH2.Visible = True
CMDsearch.Visible = False
txtSearch.Visible = False
End If
End Sub

Private Sub cmdEdit_Click()
update.Visible = True
'cmdUpdate.Visible = True
lblDistanceCalci.Visible = True
lbltotal.Visible = True
lblPayMent.Visible = True
txtDistanceCalci.Visible = True
txtUpdCalci.Visible = True
txtTotalPayment.Visible = True
End Sub

Private Sub CMDsearch_Click()
If txtSearch.Text = "" Then
MsgBox "Select Search option", , "Select"
Else
connect
s = "Select " & customer1 & "  from cbook where " & customer1 & "='" & txtSearch.Text & "'"
MsgBox s
Set r = c.Execute(s)
If r.RecordCount = 0 Then
MsgBox "Record Not Found", , "Not Found"
txtSearch.SetFocus
End If
Call TicketsGrid
End If
End Sub

Private Sub CMDSEARCH2_Click()

If Txtsearch2.Text = "" Then
MsgBox "Select option", , "Search"
Else
connect

s = "Select * from book "
MsgBox s
Set r = c.Execute(s)
If customer2 = r.Fields(0) Then
If r.RecordCount = 0 Then
MsgBox "Record not found", , "Not Found"
End If
Call totalsGrid
End If
End If
End Sub

Private Sub B_Click()
If b.Value = True Then
Text1.Text = "ENTER BOOKING ID"
Frame1.Enabled = False
Frame2.Enabled = True
Frame3.Enabled = False
Grid1.Visible = True
Grid2.Visible = False
Grid3.Visible = False
End If
End Sub

Private Sub Combo1_Click()
connect
 
Select Case Combo1.ListIndex
Case 0:
    s = "select bid from book "
    Combo2.Clear
    Combo2.Text = " select Booking id "
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
Case 1:
  Combo1.Text = "Package"

s = "select id  from book "
 Combo2.Clear
    Combo2.Text = " select package "
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
Case 2:

 Combo1.Text = "Bus No"
       s = "select no  from book "
        Combo2.Clear
    Combo2.Text = " select bus no "
Set r = c.Execute(s)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
End Select


End Sub

Private Sub Combo2_Click()
connect
connect
If Combo1.Text = "Booking Id" Then

Adodc1.RecordSource = "SELECT *FROM BOOK WHERE BID ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource


ElseIf Combo1.Text = "Package" Then

Adodc1.RecordSource = "SELECT *FROM BOOK WHERE id ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
ElseIf Combo1.Text = "Name" Then

Adodc1.RecordSource = "SELECT *FROM BOOK WHERE name ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
Else
Combo1.Text = "Bus No"
Adodc1.RecordSource = "SELECT *FROM BOOK WHERE no ='" & Combo2.Text & "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
'''ElseIf Combo1.Text = "Package" Then
'''Combo2.Clear
'''s = "select NAME  from book "
'''Set r = c.Execute(s)
'''While r.EOF <> True
'''Combo2.AddItem r.Fields(0)
''Adodc1.RecordSource = "SELECT *FROM BOOK WHERE CID ='" & Combo2.Text & "'"
''Adodc1.Refresh
''Adodc1.Caption = Adodc1.RecordSource
''Select Case Combo2.ListIndex
''Case 0:
''  Combo1.Text = "Customer Id"
''Adodc1.RecordSource = "SELECT *FROM BOOK WHERE CID ='" & Combo2.Text & "'"
''Adodc1.Refresh
''Adodc1.Caption = Adodc1.RecordSource
''Case 1:
''Adodc1.RecordSource = "SELECT *FROM BOOK WHERE id ='" & Combo2.Text & "'"
''Adodc1.Refresh
''Adodc1.Caption = Adodc1.RecordSource
'''Case 2:
'  Combo4.Text = "Name"
'    s = "select NAME  from cbook "
'     Combo3.Clear
'    Combo3.Text = " select customer name "
'Set r = c.Execute(s)
'While r.EOF <> True
'Combo3.AddItem r.Fields(0)
'r.MoveNext
'Wend
'Case 3:
'     Combo4.Text = "Departure Date"
'        s = "select ddate  from cbook "
'         Combo3.Clear
'    Combo3.Text = " select departure date "
'Set r = c.Execute(s)
'While r.EOF <> True
'Combo3.AddItem r.Fields(0)
'r.MoveNext
'Wend
'Case 4:
' Combo4.Text = "Car No"
'       s = "select car  from cbook "
'        Combo3.Clear
'    Combo3.Text = " select bus no "
'Set r = c.Execute(s)
'While r.EOF <> True
'Combo3.AddItem r.Fields(0)
'r.MoveNext
'Wend
'End Select
End If

End Sub

Private Sub Combo3_Click()
If Combo4.Text = "Customer Id" Then
Adodc2.RecordSource = "SELECT *FROM Customer WHERE CID ='" & Combo3.Text & "'"
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource
'ElseIf Combo4.Text = "Package" Then
'Adodc2.RecordSource = "SELECT *FROM CBOOK WHERE id ='" & Combo3.Text & "'"
'Adodc2.Refresh
'Adodc2.Caption = Adodc2.RecordSource
ElseIf Combo4.Text = "Name" Then
Adodc2.RecordSource = "SELECT *FROM Customer WHERE name ='" & Combo3.Text & "'"
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource
'Else
'Combo4.Text = "Bus No"
'Adodc2.RecordSource = "SELECT *FROM CBOOK WHERE CAR ='" & Combo3.Text & "'"
'Adodc2.Refresh
'Adodc2.Caption = Adodc2.RecordSource
End If
End Sub

Private Sub Combo4_Click()
connect
Select Case Combo4.ListIndex
Case 0:
    s = "select cid from customer "
    Combo3.Clear
'    Combo3.Text = " select customer id "
Set r = c.Execute(s)
While r.EOF <> True
Combo3.AddItem r.Fields(0)
r.MoveNext
Wend
Case 1:
'  Combo4.Text = "Name"
    s = "select name  from customer "
     Combo3.Clear
'    Combo3.Text = " select customer name "
Set r = c.Execute(s)
While r.EOF <> True
Combo3.AddItem r.Fields(0)
r.MoveNext
Wend
End Select
End Sub

Private Sub Combo5_Click()
connect
Select Case Combo5.ListIndex
Case 0:
    s = "select bid from cbook "
    'Combo6.Clear
'    Combo6.Text = " select Booking id "
Set r = c.Execute(s)
While r.EOF <> True
Combo6.AddItem r.Fields(0)
r.MoveNext
Wend
Case 1:
  Combo5.Text = "Package"

s = "select id  from cbook "
 Combo6.Clear
    Combo6.Text = " select package "
Set r = c.Execute(s)
While r.EOF <> True
Combo6.AddItem r.Fields(0)
r.MoveNext
Wend
Case 2:

 Combo5.Text = "Car no"
       s = "select no  from cbook "
        Combo2.Clear
    Combo6.Text = " select car no "
Set r = c.Execute(s)
While r.EOF <> True
Combo6.AddItem r.Fields(0)
r.MoveNext
Wend
End Select
End Sub

Private Sub Combo6_click()
If Combo5.Text = "Booking Id" Then

Adodc3.RecordSource = "SELECT *FROM cBOOK WHERE BID ='" & Combo6.Text & "'"
Adodc3.Refresh
Adodc3.Caption = Adodc3.RecordSource


ElseIf Combo5.Text = "Package" Then

Adodc3.RecordSource = "SELECT *FROM cBOOK WHERE id ='" & Combo6.Text & "'"
Adodc3.Refresh
Adodc3.Caption = Adodc3.RecordSource
ElseIf Combo5.Text = "Name" Then

Adodc3.RecordSource = "SELECT *FROM cBOOK WHERE name ='" & Combo6.Text & "'"
Adodc3.Refresh
Adodc3.Caption = Adodc3.RecordSource
Else
Combo5.Text = "CAR No"
Adodc3.RecordSource = "SELECT *FROM cBOOK WHERE no ='" & Combo6.Text & "'"
Adodc3.Refresh
Adodc3.Caption = Adodc3.RecordSource
End If
End Sub


Private Sub Command4_Click()
connect
If DataEnvironment1.rscustomer.State = 1 Then DataEnvironment1.rscustomer.close
DataEnvironment1.rscustomer.Open " SELECT * FROM CUSTOMER WHERE CID ='" & Combo3.Text & "' or NAME = '" & Combo3.Text & "'"
customer3.Refresh
customer3.Show
DataEnvironment1.rscustomer.close
End Sub
Private Sub Command5_Click()
If DataEnvironment1.rsCBOOK.State = 1 Then DataEnvironment1.rsCBOOK.close
DataEnvironment1.rsCBOOK.Open " SELECT * FROM cbook WHERE bID ='" & Combo6.Text & "' or no = '" & Combo6.Text & "' or id = '" & Combo6.Text & "'"
carbooking.Refresh
carbooking.Show
DataEnvironment1.rsCBOOK.close
End Sub

Private Sub DTPicker2_Click()
Adodc1.RecordSource = "SELECT  *FROM BOOK WHERE  DDATE='" & Format(DTPicker2.Value, "dd-mmm-yyyy") & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox " RECORD NOT FOUND,CHOOSE ANY  DEPARTURE DATE ", vbCritical, "TOUR & TRAEL"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub exit_Click()
Unload Me
Main.Show
End Sub

Private Sub optTickets_Click()
If optTickets.Value = True Then
cmbcustomerid.Visible = True
cmbCustomerID2.Visible = False
txtcustid.Text = ""
txtcustname.Text = ""
txtcustadd.Text = ""
txtcusttelephone.Text = ""
txtcustemailid.Text = ""
txtPackage.Text = ""
txtBusNo.Text = ""
txtDeptdate.Text = ""
txtDeptTime.Text = ""
txtFares.Text = ""
txtpiadAmount.Text = ""
txtChargesPKM = ""
txtBoarding.Text = ""
txtCash.Text = ""
txtBank.Text = ""
txtchequeNo.Text = ""
txtBranchName.Text = ""
txtBankName.Text = ""
End If
End Sub

Private Sub optTotal_Click()
If optTotal.Value = True Then
cmbCustomerID2.Visible = True
cmbcustomerid.Visible = False
txtcustid.Text = ""
txtcustname.Text = ""
txtcustadd.Text = ""
txtcusttelephone.Text = ""
txtcustemailid.Text = ""
txtPackage.Text = ""
txtBusNo.Text = ""
txtDeptdate.Text = ""
txtDeptTime.Text = ""
txtFares.Text = ""
txtpiadAmount.Text = ""
txtChargesPKM = ""
txtBoarding.Text = ""
txtCash.Text = ""
txtBank.Text = ""
txtchequeNo.Text = ""
txtBranchName.Text = ""
txtBankName.Text = ""
End If
End Sub

Private Sub txtDistanceCalci_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
txtDistanceCalci.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub


Private Sub txtTotalPayment_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
txtTotalPayment.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub txtUpdCalci_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
txtUpdCalci.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub Form_Load()
Grid3.Visible = False
End Sub

Private Sub Option1_Click()
Grid2.Visible = True
Grid3.Visible = False
End Sub

Private Sub Option2_Click()
Grid2.Visible = False
Grid3.Visible = True
End Sub

Private Sub Option3_Click()
Grid1.Visible = False
End Sub
