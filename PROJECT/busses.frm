VERSION 5.00
Begin VB.Form busses 
   Caption         =   "bus"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16365
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   Picture         =   "busses.frx":0000
   ScaleHeight     =   10755
   ScaleWidth      =   16365
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000D&
      Caption         =   "Command5"
      Height          =   615
      Left            =   480
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   8520
      Picture         =   "busses.frx":10841
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   13920
      Picture         =   "busses.frx":18521
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   11880
      Picture         =   "busses.frx":204CD
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   10200
      Picture         =   "busses.frx":28B31
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   8
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   7
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   6
      Top             =   4800
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   5
      Top             =   5520
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   4
      Top             =   6240
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   3
      Top             =   6960
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   2
      Top             =   7680
      Width           =   3495
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   11760
      TabIndex        =   1
      Top             =   8400
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   11760
      TabIndex        =   0
      Text            =   "BUS NO.."
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      Height          =   8535
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   7335
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Bus No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   17
      Top             =   2640
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bus No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   16
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   15
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "A/c:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   14
      Top             =   4800
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Milage:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   13
      Top             =   5520
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Kmph:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   12
      Top             =   6240
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Max seats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   11
      Top             =   6960
      Width           =   1500
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Colour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   10
      Top             =   7680
      Width           =   945
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Charge Per Km:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9240
      TabIndex        =   9
      Top             =   8520
      Width           =   2235
   End
End
Attribute VB_Name = "busses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
