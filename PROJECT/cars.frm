VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form cars 
   Caption         =   "cars"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15855
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15855
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Bus"
      TabPicture(0)   =   "cars.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(9)=   "Combo1"
      Tab(0).Control(10)=   "Text8"
      Tab(0).Control(11)=   "Text7"
      Tab(0).Control(12)=   "Text6"
      Tab(0).Control(13)=   "Text5"
      Tab(0).Control(14)=   "Text4"
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(16)=   "Text2"
      Tab(0).Control(17)=   "Text1"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Car"
      TabPicture(1)   =   "cars.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label16"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label17"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label18"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Combo2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text9"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text11"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text12"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text13"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text14"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text15"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text16"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.TextBox Text16 
         Height          =   495
         Left            =   7680
         TabIndex        =   27
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   7680
         TabIndex        =   26
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   7680
         TabIndex        =   25
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   7680
         TabIndex        =   24
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   7680
         TabIndex        =   23
         Top             =   4440
         Width           =   3495
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   7680
         TabIndex        =   22
         Top             =   5160
         Width           =   3495
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   7680
         TabIndex        =   21
         Top             =   5880
         Width           =   3495
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   7680
         TabIndex        =   20
         Top             =   6600
         Width           =   3495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7680
         TabIndex        =   19
         Text            =   "BUS NO.."
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   -66960
         TabIndex        =   9
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   -66960
         TabIndex        =   8
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   -66960
         TabIndex        =   7
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   -66960
         TabIndex        =   6
         Top             =   4560
         Width           =   3495
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   -66960
         TabIndex        =   5
         Top             =   5280
         Width           =   3495
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   -66960
         TabIndex        =   4
         Top             =   6000
         Width           =   3495
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   -66960
         TabIndex        =   3
         Top             =   6720
         Width           =   3495
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   -66960
         TabIndex        =   2
         Top             =   7440
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -66960
         TabIndex        =   1
         Text            =   "BUS NO.."
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Bus No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   36
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   35
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   34
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/c:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   33
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Milage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   32
         Top             =   3840
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kmph:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   31
         Top             =   4560
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max seats:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   30
         Top             =   5280
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   29
         Top             =   6000
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Per Km:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   28
         Top             =   6720
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Bus No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   18
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bus No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   17
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   16
         Top             =   3240
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/c:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   15
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Milage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   14
         Top             =   4680
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kmph:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   13
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max seats:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   12
         Top             =   6120
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   11
         Top             =   6840
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Per Km:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69720
         TabIndex        =   10
         Top             =   7560
         Width           =   1950
      End
   End
End
Attribute VB_Name = "cars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_RESIZE()
SSTab1.Width = cars.ScaleWidth

End Sub

