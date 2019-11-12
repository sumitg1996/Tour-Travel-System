VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form travell 
   Caption         =   "Travelling mode"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H8000000C&
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
      Height          =   555
      Left            =   18675
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   585
      Width           =   1410
   End
   Begin VB.TextBox nac 
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
      Left            =   11745
      MaxLength       =   4
      TabIndex        =   52
      Text            =   "0"
      Top             =   7740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox ac 
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
      Left            =   10305
      MaxLength       =   4
      TabIndex        =   51
      Text            =   "0"
      Top             =   7740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtremai 
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Text            =   "0"
      Top             =   10350
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10140
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   20355
      _ExtentX        =   35904
      _ExtentY        =   17886
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   758
      BackColor       =   -2147483639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Package details"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture3"
      Tab(0).Control(1)=   "add1"
      Tab(0).Control(2)=   "UPDATE1"
      Tab(0).Control(3)=   "close"
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(5)=   "new1"
      Tab(0).Control(6)=   "Command2"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Car details"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "UPDATE"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "close2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "new2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Hotel Details"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(1)=   "Command6"
      Tab(2).Control(2)=   "Command4"
      Tab(2).Control(3)=   "delete"
      Tab(2).Control(4)=   "new3"
      Tab(2).Control(5)=   "exit"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton Command7 
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
         Left            =   8145
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1530
         Width           =   1320
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
         Height          =   510
         Left            =   -65505
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1710
         Width           =   1545
      End
      Begin VB.CommandButton new1 
         Height          =   555
         Left            =   -70950
         Picture         =   "Form1.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7020
         Width           =   1545
      End
      Begin VB.CommandButton Command3 
         Height          =   555
         Left            =   -67800
         Picture         =   "Form1.frx":0EDF
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7020
         Width           =   1455
      End
      Begin VB.CommandButton close 
         Height          =   555
         Left            =   -66315
         Picture         =   "Form1.frx":1F68
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7020
         Width           =   1500
      End
      Begin VB.CommandButton new2 
         Height          =   555
         Left            =   2790
         Picture         =   "Form1.frx":2DF6
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6930
         Width           =   1500
      End
      Begin VB.CommandButton close2 
         Height          =   555
         Left            =   7515
         Picture         =   "Form1.frx":3C81
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6930
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Height          =   555
         Left            =   5940
         Picture         =   "Form1.frx":4B0F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6930
         Width           =   1500
      End
      Begin VB.CommandButton exit 
         Height          =   555
         Left            =   -67845
         Picture         =   "Form1.frx":5B98
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   8100
         Width           =   1500
      End
      Begin VB.CommandButton new3 
         Height          =   555
         Left            =   -73020
         Picture         =   "Form1.frx":6A26
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   8100
         Width           =   1500
      End
      Begin VB.CommandButton delete 
         Height          =   555
         Left            =   -69420
         Picture         =   "Form1.frx":78B1
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   8100
         Width           =   1500
      End
      Begin VB.CommandButton UPDATE 
         Height          =   555
         Left            =   4365
         Picture         =   "Form1.frx":893A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6930
         Width           =   1500
      End
      Begin VB.CommandButton UPDATE1 
         Height          =   555
         Left            =   -69375
         Picture         =   "Form1.frx":995D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7020
         Width           =   1545
      End
      Begin VB.CommandButton add1 
         Height          =   555
         Left            =   -70905
         Picture         =   "Form1.frx":A980
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   7020
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Height          =   555
         Left            =   2790
         Picture         =   "Form1.frx":FCEE
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   6930
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Height          =   555
         Left            =   -72615
         Picture         =   "Form1.frx":1505C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   8100
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton Command6 
         Height          =   555
         Left            =   -71040
         Picture         =   "Form1.frx":1A3CA
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   8100
         Width           =   1500
      End
      Begin VB.PictureBox Picture1 
         Height          =   9645
         Left            =   -75000
         Picture         =   "Form1.frx":1B3ED
         ScaleHeight     =   9585
         ScaleWidth      =   20250
         TabIndex        =   55
         Top             =   450
         Width           =   20310
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   5895
            MaxLength       =   15
            TabIndex        =   30
            Top             =   1860
            Width           =   3030
         End
         Begin VB.TextBox id2 
            Appearance      =   0  'Flat
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
            Left            =   5895
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   29
            Top             =   1215
            Width           =   3030
         End
         Begin VB.TextBox lice 
            Appearance      =   0  'Flat
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
            Left            =   5985
            MaxLength       =   15
            TabIndex        =   43
            Top             =   6765
            Width           =   3030
         End
         Begin VB.TextBox addr 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   6030
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   5535
            Width           =   3030
         End
         Begin VB.TextBox charge1 
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
            Left            =   4905
            MaxLength       =   4
            TabIndex        =   34
            Text            =   "0"
            Top             =   4155
            Width           =   1095
         End
         Begin VB.TextBox name3 
            Appearance      =   0  'Flat
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
            Left            =   5895
            MaxLength       =   20
            TabIndex        =   31
            Top             =   2535
            Width           =   3030
         End
         Begin VB.ComboBox select3 
            Appearance      =   0  'Flat
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
            Left            =   5895
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1260
            Width           =   3030
         End
         Begin VB.OptionButton Option1 
            Caption         =   "AC"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5895
            TabIndex        =   32
            Top             =   3105
            Width           =   960
         End
         Begin VB.TextBox sbed4 
            Appearance      =   0  'Flat
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
            Left            =   9090
            MaxLength       =   4
            TabIndex        =   41
            Top             =   4815
            Width           =   1095
         End
         Begin VB.TextBox sbed3 
            Appearance      =   0  'Flat
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
            Left            =   7695
            MaxLength       =   4
            TabIndex        =   40
            Top             =   4830
            Width           =   1095
         End
         Begin VB.TextBox charge4 
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
            Left            =   9090
            MaxLength       =   4
            TabIndex        =   37
            Text            =   "0"
            Top             =   4110
            Width           =   1095
         End
         Begin VB.TextBox charge3 
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
            Left            =   7695
            MaxLength       =   4
            TabIndex        =   36
            Text            =   "0"
            Top             =   4110
            Width           =   1095
         End
         Begin VB.TextBox sbed2 
            Appearance      =   0  'Flat
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
            Left            =   6255
            MaxLength       =   4
            TabIndex        =   39
            Top             =   4830
            Width           =   1095
         End
         Begin VB.TextBox sbed1 
            Appearance      =   0  'Flat
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
            Left            =   4905
            MaxLength       =   4
            TabIndex        =   38
            Top             =   4830
            Width           =   1095
         End
         Begin VB.TextBox charge2 
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
            Left            =   6255
            MaxLength       =   4
            TabIndex        =   35
            Text            =   "0"
            Top             =   4110
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "NON AC"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7785
            TabIndex        =   33
            Top             =   3150
            Width           =   1140
         End
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
            Height          =   555
            Left            =   14535
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   2490
            Width           =   1545
         End
         Begin VB.ComboBox Combo5 
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
            ItemData        =   "Form1.frx":5080A
            Left            =   13860
            List            =   "Form1.frx":5080C
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1890
            Width           =   2760
         End
         Begin VB.ComboBox Combo6 
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
            ItemData        =   "Form1.frx":5080E
            Left            =   13860
            List            =   "Form1.frx":5081E
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1320
            Width           =   2805
         End
         Begin VB.Label id4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Id :-"
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
            Left            =   2595
            TabIndex        =   72
            Top             =   1305
            Width           =   555
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hotel license No:-"
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
            Left            =   2700
            TabIndex        =   71
            Top             =   6810
            Width           =   2370
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Address :-"
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
            Left            =   2655
            TabIndex        =   70
            Top             =   5685
            Width           =   1305
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Charge :-"
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
            Left            =   2610
            TabIndex        =   69
            Top             =   4110
            Width           =   1215
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Room Type :-"
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
            Left            =   2565
            TabIndex        =   68
            Top             =   3165
            Width           =   1770
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hotel Name :-"
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
            Left            =   2565
            TabIndex        =   67
            Top             =   2535
            Width           =   1815
         End
         Begin VB.Label id 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hotel Id :-"
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
            Left            =   2565
            TabIndex        =   66
            Top             =   1305
            Width           =   1350
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Place"
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
            Left            =   2565
            TabIndex        =   65
            Top             =   1905
            Width           =   735
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
            Height          =   330
            Left            =   4815
            TabIndex        =   64
            Top             =   3660
            Width           =   1230
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
            Height          =   330
            Left            =   6210
            TabIndex        =   63
            Top             =   3645
            Width           =   1230
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Available room"
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
            Left            =   2565
            TabIndex        =   62
            Top             =   4875
            Width           =   1980
         End
         Begin VB.Label Label16 
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
            Height          =   330
            Left            =   8955
            TabIndex        =   61
            Top             =   3645
            Width           =   1230
         End
         Begin VB.Label Label24 
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
            Height          =   330
            Left            =   7695
            TabIndex        =   60
            Top             =   3645
            Width           =   1230
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   2085
            Left            =   13545
            Top             =   1140
            Width           =   3435
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   9645
         Left            =   45
         Picture         =   "Form1.frx":50846
         ScaleHeight     =   9585
         ScaleWidth      =   20205
         TabIndex        =   73
         Top             =   450
         Width           =   20265
         Begin VB.TextBox mss2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   4770
            MaxLength       =   2
            TabIndex        =   23
            Top             =   5310
            Width           =   1590
         End
         Begin VB.TextBox milage2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4770
            MaxLength       =   2
            TabIndex        =   22
            Top             =   4770
            Width           =   1590
         End
         Begin VB.TextBox color2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4770
            MaxLength       =   10
            TabIndex        =   20
            Top             =   4095
            Width           =   1590
         End
         Begin VB.TextBox ac2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4770
            MaxLength       =   3
            TabIndex        =   18
            Top             =   3510
            Width           =   1590
         End
         Begin VB.TextBox no2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4770
            MaxLength       =   10
            TabIndex        =   17
            Top             =   2925
            Width           =   2355
         End
         Begin VB.TextBox name2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4770
            MaxLength       =   20
            TabIndex        =   15
            Top             =   1665
            Width           =   3075
         End
         Begin VB.ComboBox select2 
            DataField       =   "NO"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1080
            Width           =   3075
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "NO"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2295
            Width           =   3075
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
            ItemData        =   "Form1.frx":7DAA6
            Left            =   7875
            List            =   "Form1.frx":7DAB0
            TabIndex        =   19
            Text            =   "Select status"
            Top             =   3465
            Width           =   1860
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
            Left            =   7875
            TabIndex        =   21
            Top             =   4005
            Width           =   1860
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Color :-"
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
            Left            =   2595
            TabIndex        =   83
            Top             =   4275
            Width           =   1020
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "A/c :-"
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
            Left            =   2640
            TabIndex        =   82
            Top             =   3645
            Width           =   720
         End
         Begin VB.Label sel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Select Car No:-"
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
            Left            =   2205
            TabIndex        =   81
            Top             =   1125
            Width           =   2010
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name:-"
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
            Left            =   2445
            TabIndex        =   80
            Top             =   1815
            Width           =   945
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Millage :-"
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
            Left            =   2595
            TabIndex        =   79
            Top             =   4860
            Width           =   1275
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Max seats :-"
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
            Left            =   2550
            TabIndex        =   78
            Top             =   5445
            Width           =   1545
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "No :-"
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
            Left            =   2640
            TabIndex        =   77
            Top             =   3105
            Width           =   675
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Place"
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
            Left            =   2550
            TabIndex        =   76
            Top             =   2385
            Width           =   735
         End
         Begin VB.Label Label27 
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
            Height          =   465
            Left            =   6750
            TabIndex        =   75
            Top             =   3510
            Width           =   960
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Rdate"
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
            Left            =   6750
            TabIndex        =   74
            Top             =   4050
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   9645
         Left            =   -75000
         Picture         =   "Form1.frx":7DAC1
         ScaleHeight     =   9585
         ScaleWidth      =   20250
         TabIndex        =   84
         Top             =   450
         Width           =   20310
         Begin VB.TextBox Text6 
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
            ForeColor       =   &H80000007&
            Height          =   450
            Left            =   6750
            MaxLength       =   5
            TabIndex        =   94
            Top             =   1260
            Width           =   2520
         End
         Begin VB.ComboBox Combo11 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   405
            ItemData        =   "Form1.frx":BFC24
            Left            =   6750
            List            =   "Form1.frx":BFC26
            TabIndex        =   3
            Top             =   3195
            Width           =   2520
         End
         Begin VB.ComboBox Combo10 
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
            ForeColor       =   &H80000007&
            Height          =   405
            Left            =   6750
            TabIndex        =   13
            Top             =   1350
            Width           =   2520
         End
         Begin VB.TextBox Text5 
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
            ForeColor       =   &H80000007&
            Height          =   405
            Left            =   6750
            MaxLength       =   2
            TabIndex        =   1
            Top             =   1935
            Width           =   900
         End
         Begin VB.TextBox Text4 
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
            ForeColor       =   &H80000007&
            Height          =   450
            Left            =   6795
            MaxLength       =   5
            TabIndex        =   6
            Top             =   5220
            Width           =   2460
         End
         Begin VB.TextBox Text3 
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
            ForeColor       =   &H80000007&
            Height          =   450
            Left            =   6795
            MaxLength       =   5
            TabIndex        =   7
            Top             =   5895
            Width           =   2460
         End
         Begin VB.ComboBox Combo7 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   405
            ItemData        =   "Form1.frx":BFC28
            Left            =   6750
            List            =   "Form1.frx":BFC2A
            TabIndex        =   4
            Top             =   3870
            Width           =   2535
         End
         Begin VB.ComboBox Combo8 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   405
            ItemData        =   "Form1.frx":BFC2C
            Left            =   6750
            List            =   "Form1.frx":BFC2E
            TabIndex        =   2
            Top             =   2565
            Width           =   2490
         End
         Begin VB.ComboBox Combo9 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   405
            ItemData        =   "Form1.frx":BFC30
            Left            =   6750
            List            =   "Form1.frx":BFC32
            TabIndex        =   5
            Top             =   4500
            Width           =   2535
         End
         Begin VB.Label hn 
            BackStyle       =   0  'Transparent
            Caption         =   "Hotel Name"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   420
            Index           =   5
            Left            =   4140
            TabIndex        =   93
            Top             =   3195
            Width           =   1680
         End
         Begin VB.Label place 
            BackStyle       =   0  'Transparent
            Caption         =   "Place"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   420
            Index           =   4
            Left            =   4140
            TabIndex        =   92
            Top             =   2610
            Width           =   1455
         End
         Begin VB.Label pc 
            BackStyle       =   0  'Transparent
            Caption         =   "package Charges"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   375
            Left            =   4140
            TabIndex        =   91
            Top             =   5940
            Width           =   2535
         End
         Begin VB.Label pack 
            BackStyle       =   0  'Transparent
            Caption         =   "Package Id"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   1
            Left            =   4185
            TabIndex        =   90
            Top             =   1395
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Package"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   4140
            TabIndex        =   89
            Top             =   1395
            Width           =   1575
         End
         Begin VB.Label car 
            BackStyle       =   0  'Transparent
            Caption         =   "car"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   420
            Index           =   2
            Left            =   4140
            TabIndex        =   88
            Top             =   3870
            Width           =   945
         End
         Begin VB.Label cn 
            BackStyle       =   0  'Transparent
            Caption         =   "car no"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   4140
            TabIndex        =   87
            Top             =   4500
            Width           =   1815
         End
         Begin VB.Label Label 
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
            ForeColor       =   &H80000007&
            Height          =   420
            Index           =   0
            Left            =   4185
            TabIndex        =   86
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label tc 
            BackStyle       =   0  'Transparent
            Caption         =   "Tour cost "
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   495
            Left            =   4140
            TabIndex        =   85
            Top             =   5220
            Width           =   3105
         End
      End
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No :-"
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
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   675
   End
End
Attribute VB_Name = "travell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim r As New ADODB.Recordset

Private Sub ac1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Color.SetFocus
End If
End Sub

Private Sub ac1_LostFocus()
ac1.Text = UCase(ac1.Text)
End Sub

Private Sub ac2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
color2.SetFocus
End If
End Sub

Private Sub ac2_LostFocus()
ac2.Text = UCase(ac2.Text)
End Sub

Private Sub add1_Click()
If M() = True Then
Exit Sub
Else
connect
On Error GoTo d:
s = "insert into cpackage values ('" & Text6.Text & "'," & Text5.Text & ",'" + Combo8.Text + "','" + Combo11.Text + "','" & Combo7.Text & "','" & Combo9.Text & "'," & Text4.Text & "," & Text3.Text & ")"
MsgBox s
Set r = c.Execute(s)
MsgBox "Data added Succesfully", vbOKOnly, "ADD"
U
d:
End If
End Sub

Private Sub addr_LostFocus()
addr.Text = UCase(addr.Text)
End Sub

Private Sub charge1_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 32 Then
If KeyAscii = 13 Then
addr.SetFocus
Exit Sub
End If
KeyAscii = 0
MsgBox "Only valid  Number like a 0 To 9", vbCritical, "Tour & Travels"
End If
End Sub

Private Sub close_Click()
unload2
End Sub

Private Sub close2_Click()
unload2
End Sub

Private Sub color_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MILAGE.SetFocus
End If
End Sub

Private Sub color_LostFocus()
'Color.Text = UCase(Color.Text)
End Sub

Private Sub color2_LostFocus()
color2.Text = UCase(color2.Text)
End Sub

Private Sub Combo10_Click()
Set DE = New ADODB.Recordset
A = "select days,id,place,hname,vehicle,no,tourc,packc from Cpackage where id='" & Combo10.Text & "'"
Set DE = c.Execute(A)
Text5.Text = DE.Fields(0)
Combo10.Text = DE.Fields(1)
Combo8.Text = DE.Fields(2)
Combo11.Text = DE.Fields(3)
Combo7.Text = DE.Fields(4)
Combo9.Text = DE.Fields(5)
Text4.Text = DE.Fields(6)
Text3.Text = DE.Fields(7)

End Sub

Private Sub Combo6_click()
Select Case Combo6.ListIndex
Case 0:
    s = "select distinct place from hotel "
    Combo5.Clear
Set r = c.Execute(s)
While r.EOF <> True
Combo5.AddItem r.Fields(0)
r.MoveNext
Wend
Case 1:

s = "select distinct H_name from hotel "
 Combo5.Clear
Set r = c.Execute(s)
While r.EOF <> True
Combo5.AddItem r.Fields(0)
r.MoveNext
Wend
Case 2:
    s = "select distinct lice_no from hotel "
     Combo5.Clear
Set r = c.Execute(s)
While r.EOF <> True
Combo5.AddItem r.Fields(0)
r.MoveNext
Wend
Case 3:
End Select
End Sub

Private Sub Combo8_Click()
Combo11.Clear
Combo7.Clear
Combo9.Clear
    A = "select no from car where not exists(select no from cpackage where car.name=cpackage.vehicle)"
    Set r = c.Execute(A)
    
While r.EOF <> True
Combo9.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select h_name from hotel where place='" & Combo8.Text & "' "
Set r = c.Execute(A)
While r.EOF <> True
Combo11.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select distinct name from Car"
Set r = c.Execute(A)
While r.EOF <> True
Combo7.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub

Private Sub Command1_Click()
connect
On Error GoTo V1:
s = "insert into car (name,place,status,no,type,color,mileage,mxs)values ('" + NAME2.Text + "','" + Combo1.Text + "','" + Combo2.Text + "','" & no2.Text + "','" + ac2.Text + "','" + color2.Text + "','" + milage2.Text + "','" + mss2.Text + "')"
Set r = c.Execute(s)
MsgBox "DATA ADDED SUCCESFULLY", vbOKOnly, "TOUR & TRAVEL"
Command1.Visible = False
U
V1:
End Sub

Private Sub Command10_Click()
On Error GoTo M1:
If Message = True Then
Exit Sub
Else
A = "delete  from car where NO='" & select2.Text & "'"
MsgBox A
Set r = c.Execute(A)
MsgBox " DATA DELETED SUCCESFULLY", vbOKOnly, "TOUR & TRAVELS"
select2.RemoveItem select2.ListIndex
U
End If
M1:
End Sub

Private Sub Command11_Click()
U
update.Visible = True
Command11.Visible = False
End Sub



Private Sub Command3_Click()
On Error GoTo X1:
connect
If MSG = True Then
Exit Sub
Else
 s = "delete from cpackage where ID='" & Combo10.Text & "'"
    Set r = c.Execute(s)
    Combo10.RemoveItem Combo10.ListIndex
    MsgBox "Data Deleted Succesfully", , "Delete"
U
End If
X1:
End Sub

Private Sub Command4_Click()
connect
On Error GoTo D1:
s = "insert into hotel values ('" & id2.Text & "','" & Text1.Text & "','" & name3.Text & "','" & ac.Text & "','" & nac.Text & "'," & charge1.Text & "," & charge2.Text & "," & charge3.Text & "," & charge4.Text & "," & sbed1.Text & "," & sbed2.Text & "," & sbed3.Text & "," & sbed4.Text & ",'" & addr.Text & "','" & lice.Text & "')"
MsgBox s
Set r = c.Execute(s)
MsgBox "DATA ADDED SUCCESFULLY", vbOKOnly, "TOUR & TRAVEL"
Unload Me
travell.Show
U
D1:
End Sub

Private Sub Command5_Click()
connect
 If DataEnvironment1.rsHOTEL.State = 1 Then DataEnvironment1.rsHOTEL.close
DataEnvironment1.rsHOTEL.Open " SELECT * FROM hotel WHERE lice_no   = '" & Combo5.Text & "' or H_name = '" & Combo5.Text & "' or place = '" & Combo5.Text & "'"
hotel.Refresh
hotel.Show
DataEnvironment1.rsHOTEL.close
If Combo5.Text = "All" Then
If DataEnvironment1.rsHOTEL.State = 1 Then DataEnvironment1.rsHOTEL.close
hotel.Show
DataEnvironment1.rsHOTEL.close
End If
End Sub

Private Sub Command6_Click()
If MSG1 = True Then
Exit Sub
Else
s = "update hotel set id= '" & select3.Text & "',h_name='" & name3.Text & "',charge1=" & charge1.Text & ",charge2=" & charge2.Text & ",charge3=" & charge3.Text & ",charge4=" & charge4.Text & ",sbed1=" & sbed1.Text & ",sbed2=" & sbed2.Text & ",sbed3=" & sbed3.Text & ",sbed4=" & sbed4.Text & ",addr='" & addr.Text & "',lice_no='" & lice.Text & "'where id='" & select3.Text & "' "
MsgBox s
Set r = c.Execute(s)
MsgBox "update"
U
End If
End Sub

Private Sub Command7_Click()
connect
If DataEnvironment1.rsCAR.State = 1 Then DataEnvironment1.rsCAR.close
DataEnvironment1.rsCAR.Open " SELECT * FROM CAR WHERE NO   ='" & select2.Text & "' "
carS.Refresh
carS.Show
DataEnvironment1.rsCAR.close
End Sub

Private Sub Command8_Click()
Unload Me
travell.Show

End Sub

Private Sub delete_Click()
If MSG1 = True Then
Exit Sub
Else
A = "delete  from hotel where id='" & select3.Text & "'"
MsgBox A
Set r = c.Execute(A)
MsgBox " DATA DELETED SUCCESFULLY", vbOKOnly, "TOUR & TRAVELS"
select3.RemoveItem select3.ListIndex
U
End If
End Sub

Private Sub exit_Click()
unload2
End Sub

Private Sub Form_Load()
Text2.Text = Format(Date, "DD-mmm-yyyy")
Text6.Visible = False
connect
connect
A = "select NO from CAR"
Set r = c.Execute(A)
While r.EOF <> True
select2.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select id from hotel"
Set r = c.Execute(A)
While r.EOF <> True
select3.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select distinct(place) from hotel"
Set r = c.Execute(A)
While r.EOF <> True
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select ID from CPACKAGE"
Set r = c.Execute(A)
While r.EOF <> True
Combo10.AddItem r.Fields(0)
r.MoveNext
Wend
A = "select distinct(place) from car"
Set r = c.Execute(A)
While r.EOF <> True
Combo8.AddItem r.Fields(0)
r.MoveNext
Wend
id4.Visible = False
id2.Visible = False
Dim B As String
B = "hotel"
s = "select count(id) from hotel"
Set r = c.Execute(s)
id2.Text = UCase(B & r.Fields(0) + 1)
g = "pack"
s = "select count(id) from cpackage"
Set r = c.Execute(s)
Text6.Text = UCase(g & r.Fields(0) + 1)
End Sub

Private Sub Form_resize()
SSTab1.Width = travell.ScaleWidth
SSTab1.Height = travell.ScaleHeight
End Sub

Private Sub Image2_RESIZE()
Image2.Width = travell.ScaleWidth
Image2.Height = travell.ScaleHeight

End Sub

Private Sub lice_LostFocus()
lice.Text = UCase(lice.Text)
End Sub

Private Sub MILAGE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mss.SetFocus
End If
End Sub

Private Sub mss_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
add1.Enabled = True
add1.SetFocus
End If
End Sub

Private Sub name1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
no.SetFocus
End If
End Sub

Private Sub name1_LostFocus()
name1.Text = UCase(name1.Text)
End Sub

Private Sub name2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
n02.SetFocus
End If
End Sub

Private Sub name2_LostFocus()
NAME2.Text = UCase(NAME2.Text)
End Sub

Private Sub name3_LostFocus()
name3.Text = UCase(name3.Text)
End Sub

Private Sub new1_Click()
Text6.Visible = True
new1.Visible = False
add1.Visible = True
Combo10.Visible = False
Combo10.Text = ""
Combo8.Text = ""
Combo11.Text = ""
Combo7.Text = ""
Combo9.Text = ""
Text4.Text = ""
Text3.Text = ""
End Sub

Private Sub new2_Click()
new2.Visible = False
Command1.Visible = True
sel.Visible = False
select2.Visible = False
NAME2.Locked = False
no2.Locked = False
ac2.Locked = False
color2.Locked = False
milage2.Locked = False
mss2.Locked = False
Command1.Visible = True
NAME2.Text = ""
ac2.Text = " "
color2.Text = " "
no2.Text = ""
milage2.Text = ""
mss2.Text = ""

End Sub

Private Sub new3_Click()
new3.Visible = False
Command4.Visible = True
id.Visible = False
id2.Visible = True
id4.Visible = True
select3.Visible = False
name3.Locked = False
id2.Locked = False
addr.Locked = False
lice.Locked = False
name3.Text = " "
charge1.Text = ""
addr.Text = ""
lice.Text = ""
End Sub

Private Sub no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ac1.SetFocus
End If
End Sub

Private Sub no_LostFocus()
no.Text = UCase(no.Text)
End Sub

Private Sub no2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ac2.SetFocus
End If
End Sub

Private Sub no2_LostFocus()
no2.Text = UCase(no2.Text)
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
ac.Text = "AC"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
nac.Text = "NON AC"
End If
End Sub

Private Sub select2_Click()
On Error GoTo a1:
Set r = New ADODB.Recordset
A = "select * from CAR where NO='" & select2.Text & "'"
MsgBox A
Set r = c.Execute(A)
NAME2.Text = r.Fields(0)
Combo1.Text = r.Fields(1)
Combo2.Text = r.Fields(2)
no2.Text = r.Fields(4)
ac2.Text = r.Fields(5)
color2.Text = r.Fields(6)
milage2.Text = r.Fields(7)
mss2.Text = r.Fields(8)
a1:
End Sub

Private Sub select3_Click()
Set r = New ADODB.Recordset
A = "select * from hotel where id='" & select3.Text & "'"
MsgBox A
Set r = c.Execute(A)
Text1.Text = r.Fields(1)
name3.Text = r.Fields(2)
ac.Text = r.Fields(3)
nac.Text = r.Fields(4)
charge1.Text = r.Fields(5)
charge2.Text = r.Fields(6)
charge3.Text = r.Fields(7)
charge4.Text = r.Fields(8)
sbed1.Text = r.Fields(9)
sbed2.Text = r.Fields(10)
sbed3.Text = r.Fields(11)
sbed4.Text = r.Fields(12)
addr.Text = r.Fields(13)
lice.Text = r.Fields(14)
End Sub

Private Sub t1_LostFocus()
t1.Text = UCase(t1.Text)
End Sub

Private Sub Text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub txtphoto_Change()
If txtphoto.Text = "" Then
txtphoto.Text = "bg_sai.jpg"
Else
carImg.Picture = LoadPicture(App.Path + "\imgs\car\" + txtphoto.Text)
End If
End Sub

Private Sub update_Click()
On Error GoTo F1:
If Message = True Then
Exit Sub
Else
s = "update CAR set name= '" + NAME2.Text + "',place = '" + Combo1.Text + "',status = '" + Combo2.Text + "',rdate ='" & Format(Text2.Text, "dd-MMM-yyyy") & "',no= '" + no2.Text + "' ,type='" + ac2.Text + "',color='" + color2.Text + "',mileage='" + milage2.Text + "',mxs='" + mss2.Text + "'where no='" & select2.Text & "' "
MsgBox s
Set r = c.Execute(s)
MsgBox "update"
NAME2.Text = " "
no2.Text = ""
ac2.Text = ""
color2.Text = ""
milage2.Text = ""
mss2.Text = ""
U
End If
F1:
End Sub

Private Sub UPDATE1_Click()
On Error GoTo B1:
If MSG = True Then
Exit Sub
Else
s = "update cpackage set days=" & Text5.Text & ",PLACE='" & Combo8.Text & "',HNAME= '" & Combo11.Text & "',VEHICLE ='" & Combo7.Text & "',NO='" & Combo9.Text & "',tourc=" & Text4.Text & ",packc=" & Text3.Text & "where id='" & Combo10.Text & "' "

Set r = c.Execute(s)
MsgBox "Data Updated Succesfully", , "Update"
U
End If
B1:
End Sub
Function U() As Boolean()
Unload Me
travell.Show
End Function
Function Message() As Boolean
On Error GoTo E1:
   If NAME2.Text = "" Then
       MsgBox "PLEASE SELECT CAR NO", vbCritical, "TOUR & TRAVEL"
       select2.SetFocus
       Message = True
   End If
E1:
   End Function
       
   Function MSG1() As Boolean
   If Text1.Text = "" Then
       MsgBox "PLEASE SELECT CAR NO", vbCritical, "TOUR & TRAVEL"

       MSG1 = True
   End If
   End Function

Function M() As Boolean
   If Text5.Text = "" Then
       MsgBox "Please Enter Days"
       Text5.SetFocus
       M = True
    ElseIf Combo8.Text = "" Then
       MsgBox "Please Choose Place"
       Combo8.SetFocus
       M = True
    ElseIf Combo11.Text = "" Then
       MsgBox "Please choose hotel name"
       Combo11.SetFocus
       M = True
    ElseIf Combo7.Text = "" Then
       MsgBox "Please Choose Car"
       Combo7.SetFocus
       M = True
    ElseIf Combo9.Text = "" Then
       MsgBox "Please choose car no "
       Combo9.SetFocus
       M = True
        ElseIf Text4.Text = "" Then
       MsgBox "Please Enter Tour Cost"
       Text4.SetFocus
       M = True
    ElseIf Text3.Text = "" Then
       MsgBox "Please Enter Package Charges  "
       Text3.SetFocus
       M = True
       End If
       End Function
