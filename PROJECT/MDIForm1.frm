VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H80000004&
   Caption         =   "TOUR & TRAVELS"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16695
   LinkMode        =   1  'Source
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "7:16 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "6/24/2017"
         EndProperty
      EndProperty
   End
   Begin VB.Menu fil 
      Caption         =   "File"
      Begin VB.Menu MAIN2 
         Caption         =   "MAIN"
      End
      Begin VB.Menu INDE 
         Caption         =   "INDEX..."
      End
      Begin VB.Menu EXIT 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu TRAVEL 
      Caption         =   "Travelling"
      Index           =   0
      Begin VB.Menu EMP32 
         Caption         =   "EMPLOYEE"
      End
      Begin VB.Menu BOOKING 
         Caption         =   "BOOKING CANCEL"
      End
      Begin VB.Menu TRAVELLLL 
         Caption         =   "TRAVELLING MODE"
      End
   End
   Begin VB.Menu repor 
      Caption         =   "Report"
      Begin VB.Menu al 
         Caption         =   "All"
      End
      Begin VB.Menu package 
         Caption         =   "package"
      End
      Begin VB.Menu customer 
         Caption         =   "Customer"
      End
      Begin VB.Menu hotel 
         Caption         =   "Hotel"
      End
      Begin VB.Menu ca 
         Caption         =   "Car"
      End
      Begin VB.Menu cancel 
         Caption         =   "Cancellation"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu cal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu about 
         Caption         =   "About Us"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub admin1_Click()
Unload newindex
Unload Me
frmLogin.Show
End Sub

Private Sub backup_Click()
End Sub

Private Sub al_Click()
 report1.Show
Unload newindex
Unload travell
'Unload book
Unload cust
Unload emp
Unload pass
'Unload About1
End Sub

Private Sub BOOKING_Click()
Unload report1
Unload newindex
Unload travell
'Unload book
Unload cust
Unload emp
Unload pass
'Unload about
BOOKINGC.Show
End Sub

Private Sub ca_Click()
Unload report1
Unload newindex
travell.Show
'Unload book
Unload cust
Unload emp
Unload pass
'Unload About1

End Sub

Private Sub cal_Click()
Shell "calc.exe", vbNormalNoFocus
End Sub

Private Sub cancel_Click()
Unload newindex
Unload report1
Unload cust
Unload emp
Unload pass
report1.Show
End Sub



Private Sub customerdetails_Click()
Unload newindex
Unload report1
Unload packag
Unload book
Unload travell
Unload pass
cust.Show
Unload Main
End Sub

Private Sub editpassword_Click()
Unload newindex
Unload report1
Unload INDE
Unload packag
Unload book
Unload cust
Unload emp
Unload travell
Unload About1
pass.Show
Unload Main

End Sub

Private Sub EMP_Click()
emp.Show
Unload report1
Unload newindex
DataReport1.Show
Unload Main
End Sub

Private Sub employeedetails_Click()
Unload report1
Unload newindex
Unload cust
Unload travell
Unload pass
emp.Show
Unload Main
End Sub

Private Sub customer_Click()
Unload report1
Unload newindex
Unload travell
Unload book
 cust.Show
Unload emp
Unload pass
Unload About1
Unload Main
End Sub

Private Sub EMP32_Click()
Unload newindex
Unload report1
Unload cust
Unload pass
emp.Show
Unload travell
Unload Main
End Sub

Private Sub exit_Click()
Dim B As String
B = MsgBox("DO YOU WANT TO EXIT ", vbYesNo, "Tour & Travel System")
If B = vbYes Then



End
End If
End Sub







Private Sub travelling_Click()
Unload newindex
Unload report1
Unload cust
Unload pass
travell.Show
Unload Main
End Sub


Private Sub hotel_Click()
Unload report1
Unload newindex
travell.Show
Unload book
Unload cust
Unload emp
Unload pass
Unload About1
Unload Main
End Sub

Private Sub INDE_Click()
Unload report1
newindex.Show
Unload travell
Unload CAR_BOOK
Unload cust
Unload emp
Unload pass
Unload pass
Unload Main
End Sub

Private Sub MAIN_Click()
Unload report1
Unload newindex
 Main.Show
Unload cust
Unload emp
Unload pass
End Sub

Private Sub MAIN2_Click()
Unload report1
Unload newindex
 Main.Show
'Unload book
Unload cust
Unload emp
Unload pass
Unload travell
End Sub

Private Sub package_Click()
Unload report1
Unload newindex
 travell.Show
'Unload book
Unload cust
Unload emp
Unload pass
'Unload About1
End Sub

Private Sub TRAVELLLL_Click()
Unload report1
Unload newindex
 travell.Show
'Unload book
Unload cust
Unload emp
Unload pass
'Unload About1
End Sub
