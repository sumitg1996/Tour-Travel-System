VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CAR_BOOK 
   Caption         =   "BOOKING"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19500
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CAR_BOOK.frx":0000
   ScaleHeight     =   10800
   ScaleWidth      =   19500
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CAR_BOOK.frx":4B308
      Height          =   1635
      Left            =   2970
      TabIndex        =   53
      Top             =   7605
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   2884
      _Version        =   393216
      BackColor       =   8421504
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   12
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
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   645
      Left            =   630
      Top             =   7920
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
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
      Connect         =   "Provider=MSDAORA.1;User ID=tour/travel;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=tour/travel;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"CAR_BOOK.frx":4B31D
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
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Left            =   9675
      MaxLength       =   8
      TabIndex        =   32
      Top             =   5040
      Width           =   2265
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4365
      TabIndex        =   31
      ToolTipText     =   "Choose Car No"
      Top             =   4230
      Width           =   2310
   End
   Begin VB.OptionButton cheque_opt 
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   30
      Top             =   4590
      Width           =   1320
   End
   Begin VB.OptionButton Optcash 
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9675
      TabIndex        =   29
      Top             =   4590
      Width           =   1050
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   9675
      MaxLength       =   8
      TabIndex        =   28
      Text            =   "0"
      Top             =   4095
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   9675
      MaxLength       =   8
      TabIndex        =   27
      Text            =   "0"
      Top             =   3600
      Width           =   2265
   End
   Begin VB.TextBox TOC2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   26
      Top             =   3060
      Width           =   2280
   End
   Begin VB.TextBox B_ID 
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
      Left            =   4365
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      Top             =   1665
      Width           =   2280
   End
   Begin VB.TextBox BOARDING2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4365
      MaxLength       =   15
      TabIndex        =   24
      Top             =   5265
      Width           =   2190
   End
   Begin VB.TextBox NP2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4365
      TabIndex        =   23
      Text            =   "0"
      Top             =   4770
      Width           =   495
   End
   Begin VB.TextBox PC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   22
      Top             =   2025
      Width           =   2280
   End
   Begin VB.TextBox TC2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9675
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   21
      Top             =   1530
      Width           =   2280
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Height          =   780
      Left            =   4770
      TabIndex        =   18
      Top             =   6030
      Width           =   4245
      Begin VB.CommandButton submit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Width           =   1515
      End
      Begin VB.CommandButton Command5 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   1545
      End
   End
   Begin VB.Frame frameCheque 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cheque "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   11970
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   4380
      Begin VB.TextBox txtBranch 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtBank_Name 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtchequeNo 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblBank 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   225
         TabIndex        =   17
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblBank 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   225
         TabIndex        =   16
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label lblChkNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -945
         TabIndex        =   15
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   225
         TabIndex        =   14
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9675
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2565
      Width           =   2280
   End
   Begin VB.TextBox Text10 
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
      Left            =   4365
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2160
      Width           =   2280
   End
   Begin VB.TextBox packa 
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
      Left            =   4365
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2655
      Width           =   2280
   End
   Begin VB.TextBox Text9 
      Height          =   420
      Left            =   16695
      TabIndex        =   5
      Top             =   7560
      Width           =   1995
   End
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "REG"
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtCash 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1260
      TabIndex        =   1
      Top             =   10845
      Width           =   1335
   End
   Begin VB.TextBox txtBank 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   -180
      TabIndex        =   0
      Top             =   10845
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   375
      Left            =   9675
      TabIndex        =   33
      Top             =   5535
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
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
      Format          =   68485123
      CurrentDate     =   42779
   End
   Begin MSComCtl2.DTPicker DTP 
      Height          =   375
      Left            =   4365
      TabIndex        =   34
      ToolTipText     =   "Choose Departure Date"
      Top             =   3195
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   661
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
      Format          =   68485123
      CurrentDate     =   42773
   End
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   375
      Left            =   4365
      TabIndex        =   35
      ToolTipText     =   "Choose Departure Time"
      Top             =   3690
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   661
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
      CustomFormat    =   "MM:HH"
      Format          =   68485122
      CurrentDate     =   42773
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Package Charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   52
      Top             =   2115
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Cost "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   51
      Top             =   1665
      Width           =   1200
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Charges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   50
      Top             =   3150
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Pay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   49
      Top             =   3690
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   48
      Top             =   4185
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   47
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7380
      TabIndex        =   46
      Top             =   5175
      Width           =   1845
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   45
      Top             =   5670
      Width           =   1365
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7380
      TabIndex        =   44
      Top             =   2655
      Width           =   1035
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H004459EE&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   6840
      Top             =   990
      Width           =   2670
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1845
      TabIndex        =   43
      Top             =   1710
      Width           =   1245
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Package "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1845
      TabIndex        =   42
      Top             =   2745
      Width           =   1170
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1845
      TabIndex        =   41
      Top             =   3285
      Width           =   2190
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1845
      TabIndex        =   40
      Top             =   3825
      Width           =   2220
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Person"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1845
      TabIndex        =   39
      Top             =   4905
      Width           =   1740
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Boarding"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   1845
      TabIndex        =   38
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1845
      TabIndex        =   37
      Top             =   2205
      Width           =   1755
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Car No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1845
      TabIndex        =   36
      Top             =   4365
      Width           =   960
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H004459EE&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   1485
      Top             =   990
      Width           =   2670
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   675
      Top             =   7110
      Width           =   12660
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   6135
      Left            =   1485
      Top             =   990
      Width           =   10770
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   675
      Top             =   540
      Width           =   12660
   End
   Begin VB.Label bookingdate 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
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
      Left            =   18000
      TabIndex        =   6
      Top             =   135
      Width           =   1410
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date"
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
      Left            =   16110
      TabIndex        =   3
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "CAR_BOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As New ADODB.Recordset
Dim s As String

Private Sub ADD2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
PH2.SetFocus
End If
End Sub

Private Sub BOARDING2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TC2.SetFocus
End If
End Sub

Private Sub cheque_opt_Click()
If cheque_opt.Value = True Then
frameCheque.Visible = True
Optcash.Enabled = False
txtBranch.Text = ""
txtchequeNo.Text = ""
txtBank_Name.Text = ""
txtBank.Text = "Yes"
txtCash = "No"
Optcash.Enabled = True
End If
End Sub

Private Sub Command7_Click()

End Sub

Private Sub cheque_opt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cheque_opt.SetFocus
End If

End Sub

Private Sub Combo2_Click()
A = "select * from CBOOK where BID ='" & Combo2.Text & "'"
MsgBox A
Set r = c.Execute(A)
B_ID.Text = r.Fields(0)
Text10.Text = r.Fields(1)
packa.Text = r.Fields(2)
DTP.Value = r.Fields(3)
DTP2.Value = r.Fields(4)
Combo1.Text = r.Fields(7)
NP2.Text = r.Fields(8)
BOARDING2.Text = r.Fields(9)
TC2.Text = r.Fields(10)
PC.Text = r.Fields(11)
Text6.Text = r.Fields(12)
TOC2.Text = r.Fields(13)
Text1.Text = r.Fields(14)
Text3.Text = r.Fields(21)
Text8.Text = (Text3.Text * Text7.Text) / 100
 Text9.Text = Val(Text3.Text) - Val(Text8.Text)
End Sub

Private Sub Command2_Click()
On Error GoTo w:

s = "insert into CARCAN values ('" & B_ID.Text & "','" & Text10.Text & "','" & packa.Text & "','" & Format(DTP.Value, "dd-mmm-yyyy") & "','" & Combo1.Text & "', " & TOC2.Text & "," & Text8.Text & "," & Text9.Text & ")"
Set r = c.Execute(s)
MsgBox "DATA ADDED SUCCESSFULLY IN CANCILATION", vbOKOnly, "TOUR & TRAVEL"
Command2.Enabled = False
Frm_cancelbill.lblBookinID.Caption = B_ID.Text
Frm_cancelbill.lblCustomerNAme.Caption = Text10.Text
Frm_cancelbill.lblPackage.Caption = packa.Text
Frm_cancelbill.lblBusNo.Caption = Combo2.Text
Frm_cancelbill.lblDeptDate.Caption = DTP.Value
Frm_cancelbill.lblToalAmt.Caption = TOC2.Text
Frm_cancelbill.lblDiscount.Caption = Text1.Text
Frm_cancelbill.lblAmountPaid.Caption = Text3.Text
Frm_cancelbill.Label3.Caption = Text8.Text
Frm_cancelbill.Label4.Caption = Text9.Text
Frm_cancelbill.Show
Unload Me
w:
End Sub

Private Sub Command5_Click()
Unload Me
MAIN.Show
End Sub

Private Sub DTP_Change()
Combo1.Clear
If DTP.Value < Date Then
MsgBox "NOT ALLOWED IN BACK DATE BOOKING", vbCritical, "Tour & Travel"
DTP.Value = Date
End If
s = "select no from car where not exists(select no from cbook where '" & Format(DTP.Value, "dd-MMM-yyyy") & "'<= rdate  and car.no=cbook.no and car.place=cbook.id) "
Set r = c.Execute(s)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
Adodc1.CommandType = adCmdUnknown
    Adodc1.RecordSource = "select no,status,place,mxs from car where not exists(select no from cbook where'" & Format(DTP.Value, "dd-mmm-yyyy") & "' <= rdate and car.no=cbook.no)"
    Adodc1.Refresh
End Sub

Private Sub DTP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
time2.SetFocus
End If
End Sub

Private Sub EMAIL2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DTP.SetFocus
End If
End Sub

Private Sub DTP1_Change()
If DTP1.Value < Date Then
MsgBox "NOT ALLOWED ", vbCritical, "Tour & Travel"
DTP1.Value = Date
End If
End Sub

Private Sub Form_Load()
DTP.Value = Date
DTP1.Value = Date
bookingdate.Caption = Format(Date, "dd-MMM-yyyy")
connect
B = "BOOKING"
s = "select count(Bid) from cbook"
Set r = c.Execute(s)
B_ID.Text = B & r.Fields(0) + 1
Dim A As String
s = "select no from car where not exists(select no from cbook where '" & Format(bookingdate.Caption, "dd-mmm-yyyy") & "'<= rdate  and car.no=cbook.no and car.place=cbook.id) "
Set r = c.Execute(s)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub

Private Sub name2_LostFocus()
NAME2.Text = UCase(NAME2.Text)
End Sub

Private Sub NP2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
NA2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub optCash_Click()
If Optcash.Value = True Then
frameCheque.Visible = False
txtCash.Text = "Yes"
txtBank.Text = "No"
txtBranch.Text = "No"
txtchequeNo.Text = "No"
txtBank_Name.Text = "No"
cheque_opt.Enabled = True
End If
End Sub

Private Sub PACKA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NAME2.SetFocus
End If
End Sub

Private Sub PACKA_LostFocus()
'A = "select distinct no from car where place = '" + packa.Text + "'"
'Set r = c.Execute(A)
'While r.EOF <> True
'Combo1.AddItem r.Fields(0)
'r.MoveNext
'Wend
End Sub

Private Sub PC_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
TOC2.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
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

Private Sub submit_Click()
If Message = True Then
Exit Sub
End If
connect
s = "insert into cbook values ('" & B_ID.Text & "','" & Text10.Text & "','" & packa.Text & "','" & Format(DTP.Value, "dd-mmm-yyyy") & "','" & Format(DTP1.Value, "dd-mmm-yyyy") & "','" & Format(DTP2.Value, "dd-mmm-yyyy") & "','" & Text4.Text & "','" & Combo1.Text & "'," & NP2.Text & ",'" & BOARDING2.Text & "'," & TC2.Text & "," & PC.Text & "," & Text6.Text & "," & TOC2.Text & "," & Text1.Text & "," & Text2.Text & ",'" & txtCash.Text & "','" & txtBank.Text & "','" & txtchequeNo.Text & "','" & txtBank_Name.Text & "','" & txtBranch.Text & "'," & Text3.Text & ", '" & Format(bookingdate.Caption, "dd-mmm-yyyy") & "','no')"
MsgBox s
Set r = c.Execute(s)
MsgBox "DATA ADDED SUCCESFULLY ", vbOKOnly, "TOUR & TRAVEL"
s = "insert into CUSTOMER values ('" & newindex.Label17.Caption & "','" & newindex.NAME2.Text & "','" & newindex.addr.Text & "'," & newindex.PH2.Text & ",'" & newindex.EMAIL2.Text & "')"
Set r = c.Execute(s)
frmTicketBill.lblBookinID.Caption = B_ID.Text
frmTicketBill.lblCustomerNAme.Caption = newindex.NAME2.Text
frmTicketBill.lblPackage.Caption = packa.Text
frmTicketBill.lblBusName.Caption = newindex.PH2.Text
frmTicketBill.lblBusNo.Caption = Combo1.Text
frmTicketBill.lblDeptDate.Caption = DTP.Value
frmTicketBill.lblDeptTime.Caption = DTP2.Value
frmTicketBill.lblChild.Caption = NP2.Text
frmTicketBill.lblToalAmt.Caption = TOC2.Text
frmTicketBill.lblAmountPaid.Caption = Text3.Text
If Optcash.Value = True Then
frmTicketBill.lblBy.Caption = "Cash"
End If
Unload Me
Unload newindex
frmTicketBill.Show
MAIN.WindowState = 2
End Sub

Private Sub TC2_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
PC.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub Text1_Change()
Text1.Text = Format(Text1.Text, "0.00")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Text2.SetFocus
Text2.Text = Val(TOC2.Text) - Val(Text1.Text)
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub Text1_LostFocus()
Text2.Text = Val(TOC2.Text) - Val(Text1.Text)
End Sub

Private Sub Text2_Change()
Text2.Text = Format(Text2.Text, "000.00")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Optcash.SetFocus
End If
End Sub

Private Sub Text3_Change()
Text3.Text = Format(Text3.Text, "000.00")
End Sub

Private Sub Text6_LostFocus()
Text5.Text = Val(TOC2.Text) * Val(Text6.Text) / 100
TOC2.Text = Val(TOC2.Text) - Val(Text5.Text)
TOC2.Text = Format(TOC2.Text, "000.00")
End Sub

Private Sub TIME2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NP2.SetFocus
End If
End Sub

Private Sub TOC2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
End Sub

Function Message() As Boolean
If packa.Text = "Select Package" Then
MsgBox "Please Select Package "
packa.SetFocus
Message = True
End If
End Function

