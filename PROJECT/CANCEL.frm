VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BOOKINGC 
   Caption         =   " Cancel Form"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CANCEL.frx":0000
   ScaleHeight     =   9075
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   1350
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   5040
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   180
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   5040
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CANCEL.frx":771C9
      Height          =   3435
      Left            =   1080
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   6059
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   -2147483645
      Enabled         =   -1  'True
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   19
      TabAcrossSplits =   -1  'True
      TabAction       =   2
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CAR BOOKING DETAILS"
      ColumnCount     =   28
      BeginProperty Column00 
         DataField       =   "CID"
         Caption         =   "Customer Id"
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
         DataField       =   "BID"
         Caption         =   "Booking Id"
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
         DataField       =   "ID"
         Caption         =   "Package"
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
         DataField       =   "DDATE"
         Caption         =   "Ddate"
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
         DataField       =   "RDATE"
         Caption         =   "Rdate"
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
         DataField       =   "NO"
         Caption         =   "Bus No"
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
         DataField       =   "NP"
         Caption         =   "No Of Person"
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
         DataField       =   "BOARD"
         Caption         =   "Boarding"
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
         DataField       =   "TC"
         Caption         =   "Tour Charge"
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
         DataField       =   "PC"
         Caption         =   "Package Charge"
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
      BeginProperty Column12 
         DataField       =   "DISC"
         Caption         =   "Discount"
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
         Caption         =   "Total Charge"
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
         Caption         =   "Advance"
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
         Caption         =   "Due Amount"
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
         Caption         =   "Cash"
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
         DataField       =   "CHEQUE"
         Caption         =   "Cheque"
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
         DataField       =   "PAMOUNT"
         Caption         =   "Paid Amount"
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
         DataField       =   "DATE1"
         Caption         =   "Booking Date"
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
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      BeginProperty Column23 
         DataField       =   ""
         Caption         =   ""
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
      BeginProperty Column24 
         DataField       =   ""
         Caption         =   ""
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
      BeginProperty Column25 
         DataField       =   ""
         Caption         =   ""
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
      BeginProperty Column26 
         DataField       =   ""
         Caption         =   ""
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
      BeginProperty Column27 
         DataField       =   ""
         Caption         =   ""
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
         Size            =   3
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column23 
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   315
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   -135
      TabIndex        =   10
      Top             =   7290
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
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
      Left            =   9045
      Picture         =   "CANCEL.frx":771DE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5625
      Width           =   1320
   End
   Begin VB.OptionButton Option2 
      Caption         =   "CAR BOOKING"
      Height          =   330
      Left            =   8055
      TabIndex        =   1
      Top             =   225
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   645
      Left            =   4995
      Top             =   7470
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      RecordSource    =   "select *from customer,book where 1>2"
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
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
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
      Left            =   9495
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5085
      Width           =   1410
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   8055
      TabIndex        =   3
      Top             =   5085
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   5625
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   7
      Top             =   7290
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   555
      Left            =   8235
      Top             =   7560
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   979
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
      RecordSource    =   "SELECT  *FROM CUSTOMER ,CBOOK WHERE 1>2"
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6165
      TabIndex        =   13
      Top             =   5085
      Width           =   1725
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   510
      Left            =   15615
      TabIndex        =   12
      Top             =   225
      Width           =   2220
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "           %"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7605
      TabIndex        =   11
      Top             =   4635
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Amount"
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
      Left            =   9495
      TabIndex        =   9
      Top             =   4635
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
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
      Left            =   6525
      TabIndex        =   8
      Top             =   4635
      Width           =   1365
   End
End
Attribute VB_Name = "BOOKINGC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim s As String
Dim rS As New ADODB.Recordset

Private Sub Command1_Click()
connect

 Dim ans As String
    ans = InputBox("ENTER NAME OR BOOKING ID ", "SEARCH")
 If ans = "" Then
       Text1.Text = "No message"
    Else
        Text1.Text = ans
        Text6.Text = ans
        If Option2.Value = True Then
       s = " SELECT C.BID,C.ID,B.NAME FROM CBOOK C,CUSTOMER B WHERE B.NAME ='" & Text1.Text & "' AND C.CID= B.CID "
       Set r = c.Execute(s)
       If r.EOF <> True Or r.BOF <> True Then
       Text6.Text = r.Fields(0)
       Text2.Text = r.Fields(1)
       End If
       End If
       
    End If
End Sub

Private Sub Command2_Click()
If Label4.Caption > Date Then
MsgBox "NOT ALLOWED IN BACK DATE ", vbCritical
Else
Dim h As String
connect
Set rS = New ADODB.Recordset
s = "SELECT N.NAME,D.BID,D.ID,D.NO,D.DDATE,D.TOTAL_C,D.DISC,D.PAMOUNT FROM CUSTOMER N,cBOOK D WHERE D.BID='" & Text6.Text & "'AND N.CID=D.CID OR N.NAME ='" & Text6.Text & "' AND N.CID=D.CID"
Set rS = c.Execute(s)
Set r = New ADODB.Recordset
h = "insert into carcan values ('" & Text5.Text & "'," & Text7.Text & "," & Text4.Text & ",'" & Text2.Text & "')"
Set r = c.Execute(h)
'Set ds = New ADODB.Recordset
'A = "update CBOOK set  STATUS='" & Text8.Text & "'"
'Set ds = c.Execute(A)
Frm_cancelbill.lblCustomerNAme.Caption = rS.Fields(0)
Frm_cancelbill.lblBookinID.Caption = rS.Fields(1)
Frm_cancelbill.lblPackage.Caption = rS.Fields(2)
Frm_cancelbill.lblBusNo.Caption = rS.Fields(3)
Frm_cancelbill.lblDeptDate.Caption = rS.Fields(4)
Frm_cancelbill.lblToalAmt.Caption = rS.Fields(5)
Frm_cancelbill.lblDiscount.Caption = rS.Fields(6)
Frm_cancelbill.lblAmountPaid.Caption = rS.Fields(7)
Frm_cancelbill.Label3.Caption = Text7.Text
Frm_cancelbill.Label4.Caption = Text4.Text
Frm_cancelbill.Show
End If

End Sub
Private Sub Form_Load()
Label4.Caption = Format(Date, "dd-mmm-yyyy")
End Sub
Private Sub Label5_Change()
Label5.Caption = Format(Label5.Caption, ".00")
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
DataGrid1.ClearSelCols
DataGrid1.Visible = True
Text3.Text = ""

End If



End Sub

Private Sub Text1_Change()
If Option2.Value = True Then
Adodc1.Refresh
Adodc1.RecordSource = "SELECT *FROM CUSTOMER C,CBOOK B WHERE  b.BID='" & UCase(Text1.Text) & "'and B.cid=C.cid OR C.NAME='" & UCase(Text1.Text) & "'  and B.cid=C.cid "
Adodc1.Refresh
s = "select  *from cbook where bid='" & UCase(Text1.Text) & "'"
Set r = c.Execute(s)
 If r.EOF <> True Or r.BOF <> True Then
 Label5.Caption = r.Fields(21)
End If
If Adodc1.Recordset.EOF Then
MsgBox " RECORD NOT FOUND,CHOOSE OTHER BOOKING DATE ", vbCritical, "TOUR & TRAEL"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If
End Sub

Private Sub Text3_Change()
Text7.Text = (Val(Label5.Caption) * Val(Text3.Text) / 100)
Text4.Text = Val(Label5.Caption) - Val(Text7.Text)
End Sub

Private Sub Text3_LostFocus()
s = "select   ID,CID from cbook where bid='" & UCase(Text1.Text) & "'"
Set r = c.Execute(s)
 If r.EOF <> True Or r.BOF <> True Then
 Text2.Text = r.Fields(0)
 Text5.Text = r.Fields(1)
 End If
End Sub

Private Sub Text4_Change()
Text4.Text = Format(Text4.Text, ".00")
End Sub

Private Sub Text6_Change()
Text6.Text = UCase(Text6.Text)
End Sub
