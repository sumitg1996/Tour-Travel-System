VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form book 
   BackColor       =   &H80000009&
   Caption         =   "Booking"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "booking.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   20190
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   11025
      MaxLength       =   5
      TabIndex        =   51
      Top             =   1260
      Width           =   570
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
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   50
      Top             =   2295
      Width           =   2280
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   13725
      TabIndex        =   49
      Top             =   6255
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   46
      Top             =   1755
      Width           =   2280
   End
   Begin VB.TextBox Text7 
      Height          =   420
      Left            =   17415
      TabIndex        =   45
      Top             =   6570
      Width           =   1995
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   -135
      TabIndex        =   44
      Text            =   "Resv"
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   7965
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1845
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   10620
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1845
      Width           =   1785
   End
   Begin VB.TextBox b_id 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   4320
      MaxLength       =   20
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1260
      Width           =   2280
   End
   Begin VB.TextBox dam 
      Alignment       =   2  'Center
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
      Height          =   420
      Left            =   9090
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   11
      Text            =   "0"
      Top             =   3735
      Width           =   1785
   End
   Begin VB.TextBox Total 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   9090
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2790
      Width           =   1785
   End
   Begin VB.TextBox advance 
      Alignment       =   2  'Center
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
      Left            =   9090
      MaxLength       =   8
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3240
      Width           =   1785
   End
   Begin VB.TextBox extra 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   9090
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1260
      Width           =   1785
   End
   Begin VB.OptionButton optCash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "cash"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9090
      TabIndex        =   12
      Top             =   4455
      Width           =   1020
   End
   Begin VB.OptionButton optCheque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "cheque"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10080
      TabIndex        =   13
      Top             =   4455
      Width           =   1290
   End
   Begin VB.TextBox disc 
      Alignment       =   2  'Center
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
      Left            =   9090
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "0"
      Top             =   2295
      Width           =   1770
   End
   Begin VB.Frame frameCheque 
      BackColor       =   &H00404040&
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
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   11250
      TabIndex        =   26
      Top             =   4770
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtBranch 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtBank_Name 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtchequeNo 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   2
         Left            =   225
         TabIndex        =   30
         Top             =   585
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   29
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
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblBank 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Index           =   0
         Left            =   -945
         TabIndex        =   27
         Top             =   -135
         Width           =   1215
      End
   End
   Begin VB.TextBox Text2 
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
      Left            =   9090
      TabIndex        =   14
      Text            =   "0"
      Top             =   4995
      Width           =   1815
   End
   Begin VB.TextBox bookingdate 
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
      Left            =   18765
      TabIndex        =   23
      Top             =   90
      Width           =   1500
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4590
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Height          =   915
      Left            =   4860
      TabIndex        =   21
      Top             =   5670
      Width           =   4290
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESET"
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
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
         Width           =   1545
      End
      Begin VB.CommandButton submit 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   18
         Top             =   315
         Width           =   1515
      End
   End
   Begin VB.TextBox txtBank 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   -45
      TabIndex        =   20
      Top             =   10575
      Width           =   1335
   End
   Begin VB.TextBox txtCash 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1395
      TabIndex        =   0
      Top             =   10575
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   4335
      TabIndex        =   2
      Top             =   2835
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   57802755
      CurrentDate     =   42716
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   420
      Left            =   4320
      TabIndex        =   3
      Top             =   3915
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   57802755
      CurrentDate     =   42716
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   420
      Left            =   4320
      TabIndex        =   48
      Top             =   3375
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "mm:hh"
      Format          =   57802754
      CurrentDate     =   42716
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   5955
      Left            =   1440
      Top             =   855
      Width           =   11355
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   5955
      Left            =   1620
      Top             =   1035
      Width           =   11355
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5955
      Left            =   1530
      Top             =   945
      Width           =   11355
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking id"
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
      Left            =   2025
      TabIndex        =   47
      Top             =   1260
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return date"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2025
      TabIndex        =   43
      Top             =   4005
      Width           =   1485
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due amount"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6930
      TabIndex        =   42
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label23 
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
      Left            =   2025
      TabIndex        =   41
      Top             =   1800
      Width           =   1665
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charges"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6930
      TabIndex        =   40
      Top             =   1350
      Width           =   1065
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay mode"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6930
      TabIndex        =   39
      Top             =   4455
      Width           =   1275
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advance pay"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6930
      TabIndex        =   38
      Top             =   3285
      Width           =   1695
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total amount"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6930
      TabIndex        =   37
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departure time"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2025
      TabIndex        =   36
      Top             =   3420
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departure date"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2025
      TabIndex        =   35
      Top             =   2880
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Package id"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2025
      TabIndex        =   34
      Top             =   2295
      Width           =   1440
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6930
      TabIndex        =   33
      Top             =   2340
      Width           =   1155
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Bus no"
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
      Left            =   2025
      TabIndex        =   32
      Top             =   4590
      Width           =   1230
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
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
      Left            =   6930
      TabIndex        =   31
      Top             =   5130
      Width           =   1995
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   6930
      TabIndex        =   25
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   9900
      TabIndex        =   24
      Top             =   1890
      Width           =   675
   End
   Begin VB.Label Label29 
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
      Left            =   17010
      TabIndex        =   22
      Top             =   135
      Width           =   1545
   End
End
Attribute VB_Name = "book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Private Sub boarding_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
disc.SetFocus
End If
End Sub

Private Sub boarding_LostFocus()
boarding.Text = UCase(boarding.Text)
End Sub


Private Sub advance_Change()
advance.Text = Format(advance.Text, "000.00")
End Sub

Private Sub advance_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Total.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub
Private Sub booked_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
reman.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub












Private Sub advance_LostFocus()
dam.Text = Val(Total.Text) - Val(advance.Text)
dam.Text = Format(dam.Text, "000.00")
End Sub

Private Sub DTPicker3_Change()
Dim TDate As Date
TDate = CDate(Format(Date, "dd-mmm-yy"))
If DTPicker3.Value = DateAdd("d", 3, TDate) Then
Frame2.Visible = True
ElseIf DTPicker3.Value = DateAdd("d", 6, TDate) Then
FrameSleeper.Visible = ture
Else
MsgBox "error"
End If
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub







Private Sub Combo1_Click()
'Dim B As String


A = "select * from BOOK where bID ='" & Combo1.Text & "'"
MsgBox A
Set r = c.Execute(A)
B_ID.Text = r.Fields(0)
Text9.Text = r.Fields(1)
packa.Text = r.Fields(2)
DTPicker1.Value = r.Fields(3)
'Depar.Value = r.Fields(7)
Combo2.Text = r.Fields(7)
Text3.Text = r.Fields(8)
Text1.Text = r.Fields(9)
'adult.Text = r.Fields(13)
extra.Text = r.Fields(10)
'Child.Text = r.Fields(15)
disc.Text = r.Fields(11)
Total.Text = r.Fields(12)
advance.Text = r.Fields(13)
dam.Text = r.Fields(14)
Text2.Text = r.Fields(20)
Text8.Text = (Total.Text * Text6.Text) / 100
 Text7.Text = Val(Total.Text) - Val(Text8.Text)
End Sub

Private Sub Command1_Click()
If Opt1.Value = True Then
s = "delete from BOOK where CID='" & customer1.Text & "'"
Set r = c.Execute(s)
CUST_ID.Text = " "
packa.Text = " "
name1.Text = " "
address.Text = " "
phno.Text = " "
email.Text = " "
DTPicker1.Value = " "
Depar.Text = " "
Text1.Text = " "
seat.Text = " "
na.Text = " "
NC.Text = " "
np.Text = " "
boarding.Text = " "
adult.Text = " "
extra.Text = " "
Child.Text = " "
disc.Text = " "
Total.Text = " "
advance.Text = " "
dam.Text = " "
Text2.Text = " "
Combo1.RemoveItem ListIndex
End If
End Sub

Private Sub Command2_Click()
s = "delete from BOOK where CID='" & Combo1.Text & "'"
Set r = c.Execute(s)
CUST_ID.Text = " "
packa.Text = " "
name1.Text = " "
address.Text = " "
phno.Text = " "
email.Text = " "

Text1.Text = " "
Text3.Text = ""

extra.Text = " "
disc.Text = " "
Total.Text = " "
advance.Text = " "
dam.Text = " "
Text2.Text = " "
End Sub

Private Sub Command3_Click()
s = "insert into carCAN values ('" & B_ID.Text & "', '" & Text9.Text & "','" & packa.Text & "','" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "','" & Combo2.Text & "', " & Total.Text & "," & Text7.Text & "," & Text8.Text & ")"
MsgBox s
Set r = c.Execute(s)
MsgBox "DATA ADDED SUCCESSFULLY IN CANCILATION", vbOKOnly, "TOUR & TRAVEL"
Frm_cancelbill.lblBookinID.Caption = B_ID.Text
Frm_cancelbill.lblCustomerNAme.Caption = Text9.Text
Frm_cancelbill.lblPackage.Caption = packa.Text
Frm_cancelbill.lblBusNo.Caption = Combo2.Text
Frm_cancelbill.lblDeptDate.Caption = DTPicker1.Value
Frm_cancelbill.lblToalAmt.Caption = Total.Text
Frm_cancelbill.lblDiscount.Caption = advance.Text
Frm_cancelbill.lblAmountPaid.Caption = Text2.Text
Frm_cancelbill.Label3.Caption = Text8.Text
Frm_cancelbill.Label4.Caption = Text7.Text
Frm_cancelbill.Show
Unload Me
End Sub

Private Sub Command5_Click()
Unload Me
book.Show
End Sub

Private Sub dam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Optcash.SetFocus
End If
End Sub

Private Sub depar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
seat.SetFocus
End If
End Sub

Private Sub disc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
advance.SetFocus
End If
End Sub

Private Sub disc_LostFocus()
Text5.Text = Val(extra.Text) * Val(disc.Text) / 100
Total.Text = Val(extra.Text) - Val(Text5.Text)
Total.Text = Format(Total.Text, ".00")
'Total.Text = (Val(adult.Text) + Val(extra.Text) + Val(Child.Text)) - Val(disc.Text)
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Depar.SetFocus
End If
End Sub

Private Sub email_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DTPicker1.SetFocus
End If
End Sub

Private Sub email_LostFocus()
Dim i, j, k As String
i = email.Text
j = InStr(i, "@")
k = InStr(i, ".")
If j = 0 Or k = 0 Then
MsgBox "Invalid email ID", , "Invalid"
email.Text = ""
name1.SetFocus
End If
End Sub

'Private Sub na_Change()
'adult.Text = Val(na.Text) * Val(adult.Text)
'End Sub



Private Sub na_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NC.SetFocus
End If
End Sub


Private Sub na_LostFocus()
adult.Text = (na.Text) * (r.Fields(7))
End Sub

Private Sub name1_LostFocus()
name1.Text = UCase(name1.Text)
End Sub


Private Sub nc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
np.SetFocus
End If
End Sub





Private Sub NC_LostFocus()
np.Text = Val(NC.Text) + Val(na.Text)
Child.Text = Val(NC.Text) * Val(r.Fields(8))
End Sub

Private Sub np_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
boarding.SetFocus
End If
End Sub

Private Sub np_LostFocus()
extra.Text = Val(np.Text) * Val(r.Fields(6))
End Sub

Private Sub extra_Change()
'extra.Text = (Text10.Text) * extra.Text
End Sub

Private Sub optCash_Click()
If Optcash.Value = True Then
frameCheque.Visible = False
txtCash.Text = "Yes"
txtBank.Text = "No"
txtBranch.Text = "No"
txtchequeNo.Text = "No"
txtBank_Name.Text = "No"
optCheque.Enabled = True
End If
End Sub

Private Sub optCheque_Click()
If optCheque.Value = True Then
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

'

Private Sub departureti_LostFocus()
FrameSleeper.Visible = True
End Sub

Private Sub exit_Click()

unload2
End Sub

Private Sub Form_Load()
bookingdate.Text = Format(Date, "dd-MMM-yyyy")
DTPicker1.Value = Date
DTPicker2.Value = Date
connect
connect
B = "BOOKING"
s = "select count(bid) from book"
Set r = c.Execute(s)
B_ID.Text = B & r.Fields(0) + 1
A = "select BID from BOOK"
Set r = c.Execute(A)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
Exit Sub
End Sub

Private Sub PACKA_LostFocus()
A = "select no from bus where place = '" + packa.Text + "'"
Set r = c.Execute(A)
While r.EOF <> True
Combo2.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub

Private Sub submit_Click()
If Message = True Then
Exit Sub
End If
connect
s = "insert into book values ('" & B_ID.Text & "','" & Text9.Text & "','" & packa.Text & "','" & Format(DTPicker1.Value, "dd-mmm-yyyy") & "','" & Format(DTPicker3.Value, "dd-mmm-yyyy") & "','" & Text4.Text & "','" & Format(DTPicker2.Value, "dd-mmm-yyyy") & "','" & Combo2.Text & "','" & Text3.Text & "','" & Text1.Text & "'," & extra.Text & "," & disc.Text & ", " & Total.Text & "," & advance.Text & "," & dam.Text & ",'" & txtCash.Text & "','" & txtBank.Text & "','" & txtchequeNo.Text & "','" & txtBank_Name.Text & "', '" & txtBranch.Text & "'," & Text2.Text & ",'" & Format(bookingdate.Text, "dd-mmm-yyyy") & "','no')"
MsgBox s
Set r = c.Execute(s)
MsgBox "DATA ADDEDE SUCCESFULLY", vbOKOnly, "TOUR & TRAVEL"
s = "insert into CUSTOMER values ('" & newindex.CUST_ID.Text & "','" & newindex.NAME2.Text & "','" & newindex.addr.Text & "'," & newindex.PH2.Text & ",'" & newindex.EMAIL2.Text & "')"
Set r = c.Execute(s)
's = "update bus set STATUS= '" & Text4.Text & "',rdate = '" & Format(DTPicker2.Value, "dd-MMM-yyyy") & "' where no='" & Combo2.Text & "' "
'Set r = c.Execute(s)

BUSSBILL.lblBookinID.Caption = B_ID.Text
BUSSBILL.lblCustomerNAme.Caption = newindex.NAME2.Text
BUSSBILL.lblPackage.Caption = packa.Text
BUSSBILL.lblBusName.Caption = newindex.PH2.Text
BUSSBILL.lblBusNo.Caption = Combo2.Text
BUSSBILL.lblDeptDate.Caption = DTPicker1.Value
BUSSBILL.lblreturndate.Caption = DTPicker2.Value
BUSSBILL.lblDeptTime.Caption = Text5.Text
'BUSSBILL.lblseatNumber.Caption = seat.Text
BUSSBILL.lblToalAmt.Caption = Total.Text
BUSSBILL.lblDiscount.Caption = disc.Text
BUSSBILL.lblAmountPaid.Caption = dam.Text
BUSSBILL.Label16.Caption = Text3.Text
BUSSBILL.Label12.Caption = Text1.Text
If Optcash.Value = True Then
BUSSBILL.lblBy.Caption = "Cash"
End If
If optCheque.Value = True Then
BUSSBILL.lblBy.Caption = "Cheque"
End If
BUSSBILL.Show
Unload Me
End Sub





Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text10_LostFocus()
Total.Text = (extra.Text) * (Text10.Text)
Total.Text = Format(Total.Text, "000.00")
End Sub

Private Sub Text2_Change()
Text2.Text = Format(Text2.Text, ".00")
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)
End Sub

Private Sub Total_Change()
Text5.Text = (Val(extra.Text) * Val(disc.Text) * Val(Text10.Text)) / 100
'Total.Text = Val(extra.Text) - Val(Text5.Text)
Total.Text = Format(Total.Text, ".00")
End Sub

Private Sub Total_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
Optcash.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub
Function Message() As Boolean
If packa.Text = "Select Package" Then
MsgBox "Please Select Package ", vbCritical, "TOUR & TRVEL"
packa.SetFocus
Message = True

End If
End Function

