VERSION 5.00
Begin VB.MDIForm Main_Menu 
   BackColor       =   &H0080C0FF&
   Caption         =   "CONTROL WINDOW"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8070
   Icon            =   "Main_Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main_Menu.frx":0442
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11820
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.TextBox Text1 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   2280
         TabIndex        =   1
         Text            =   "HOTEL  MANAGEMENT  SYSTEM"
         Top             =   0
         Width           =   7575
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   11040
         Picture         =   "Main_Menu.frx":1786
         Stretch         =   -1  'True
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         Picture         =   "Main_Menu.frx":3578
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   15
         Left            =   3960
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Menu cs 
      Caption         =   "Customer"
      Begin VB.Menu cde 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu cbe 
         Caption         =   "Browser Entry"
      End
      Begin VB.Menu customerfeedback 
         Caption         =   "Feedback"
         Begin VB.Menu customerfeedbackde 
            Caption         =   "Data Entry"
         End
         Begin VB.Menu cfbe 
            Caption         =   "Browser Entry"
         End
      End
   End
   Begin VB.Menu emp 
      Caption         =   "Employee"
      Begin VB.Menu ede 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu ebe 
         Caption         =   "Browser Entry"
      End
      Begin VB.Menu epayslip 
         Caption         =   "Payslip"
         Begin VB.Menu epde 
            Caption         =   "Data Entry"
         End
         Begin VB.Menu epbe 
            Caption         =   "Browser Entry"
         End
      End
      Begin VB.Menu epayreport 
         Caption         =   "Payreport"
         Begin VB.Menu epade 
            Caption         =   "Data Entry"
         End
         Begin VB.Menu epabe 
            Caption         =   "Browser Entry"
         End
      End
   End
   Begin VB.Menu rm 
      Caption         =   "Room"
      Begin VB.Menu rde 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu rbe 
         Caption         =   "Browser Entry"
      End
   End
   Begin VB.Menu foodmenu 
      Caption         =   "Food"
      Begin VB.Menu fmc 
         Caption         =   "Menu_Card"
         Begin VB.Menu fmcde 
            Caption         =   "Data entry"
         End
         Begin VB.Menu fmcbe 
            Caption         =   "Browser Entry"
         End
      End
   End
   Begin VB.Menu bookingmenu 
      Caption         =   "Booking"
      Begin VB.Menu bde 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu bbe 
         Caption         =   "Browser Entry"
      End
      Begin VB.Menu bbookingroom 
         Caption         =   "Bookingroom"
      End
      Begin VB.Menu Bbcl 
         Caption         =   "bookingcancel"
         Begin VB.Menu bcde 
            Caption         =   "Data Entry"
         End
         Begin VB.Menu bcbe 
            Caption         =   "Browser Entry"
         End
      End
   End
   Begin VB.Menu ordermenu 
      Caption         =   "Order"
      Begin VB.Menu ode 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu obe 
         Caption         =   "Browser Entry"
      End
   End
   Begin VB.Menu Paymentmenu 
      Caption         =   "Payment"
      Begin VB.Menu pde 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu pbe 
         Caption         =   "Browser Entry"
      End
   End
   Begin VB.Menu mgt 
      Caption         =   "Management"
      Begin VB.Menu maduty 
         Caption         =   "ADuty"
      End
      Begin VB.Menu mq 
         Caption         =   "Query"
      End
   End
   Begin VB.Menu pd 
      Caption         =   "Password"
   End
   Begin VB.Menu hp 
      Caption         =   "Help"
   End
   Begin VB.Menu aboutmenu 
      Caption         =   "About"
      Begin VB.Menu atd 
         Caption         =   "About The Developer"
      End
      Begin VB.Menu ats 
         Caption         =   "About the System"
      End
   End
   Begin VB.Menu noticemenu 
      Caption         =   "Notice"
      Begin VB.Menu nde 
         Caption         =   "Data Entry"
      End
      Begin VB.Menu nbe 
         Caption         =   "Browser Entry"
      End
   End
   Begin VB.Menu exitmenu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub afrmSplash_Click()
frmSplash.Show
End Sub

Private Sub atd_Click()
Adeveloper.Show
End Sub

Private Sub ats_Click()
Asystem.Show
End Sub

Private Sub bbe_Click()
Booking_Browser.Show
End Sub
Private Sub bbookingroom_Click()
Bookingroom.Show
End Sub

Private Sub bcbe_Click()
Bookingcancel_Browser.Show
End Sub
Private Sub bcde_Click()
Bookingcancel.Show
End Sub
Private Sub bde_Click()
Booking.Show
End Sub

Private Sub cbe_Click()
Customer_Browser.Show
End Sub

Private Sub cde_Click()
CUSTOMER.Show
End Sub

Private Sub cfbe_Click()
Cfeedback_Browser.Show
End Sub

Private Sub customerfeedbackde_Click()
Cfeedback.Show
End Sub

Private Sub ebe_Click()
Employee_Browser.Show
End Sub

Private Sub ede_Click()
Employee.Show
End Sub

Private Sub epabe_Click()
Payreport_Browser.Show
End Sub

Private Sub epade_Click()
Payreport.Show
End Sub

Private Sub epbe_Click()
Payslip_Browser.Show
End Sub

Private Sub epde_Click()
Payslip.Show
End Sub

Private Sub exitmenu_Click()
Dim response As Integer
response = MsgBox("Do you really want to quit PARK CHAIN OF HOTEL", 36, "Are You Sure")
If response = 6 Then
End
Else
Main_Menu.Show
End If
End Sub
Private Sub mabe_Click()
ADuty_Browser.Show
End Sub

Private Sub mabentry_Click()
ADuty_Browser.Show
End Sub

Private Sub mad_Click()
ADuty.Show
End Sub

Private Sub made_Click()
ADuty.Show
End Sub

Private Sub madentry_Click()
ADuty.Show
End Sub
Private Sub fmcbe_Click()
MenuCard_Browser.Show
End Sub

Private Sub fmcde_Click()
Menu_Card.Show
End Sub

Private Sub hp_Click()
Help.Show
End Sub

Private Sub maduty_Click()
ADuty_Browser.Show
End Sub

Private Sub mcbe_Click()
MenuCard_Browser.Show
End Sub

Private Sub mcde_Click()
Menu_Card.Show
End Sub
Private Sub mquerybookingroom_Click()
QueryBookingroom.Show
End Sub

Private Sub mquerycustomer_Click()
QueryCustomer.Show
End Sub

Private Sub mqueryemployee_Click()
QueryEmployee.Show
End Sub

Private Sub mquerypayreport_Click()
QueryPayreport.Show
End Sub

Private Sub mq_Click()
Query.Show
End Sub

Private Sub nbe_Click()
Notice_Browser.Show
End Sub

Private Sub nde_Click()
Notice.Show
End Sub

Private Sub obe_Click()
Order_Browser.Show
End Sub

Private Sub ode_Click()
Order.Show
End Sub

Private Sub pbe_Click()
Payment_Browser.Show
End Sub

Private Sub pd_Click()
Password.Show
End Sub

Private Sub pde_Click()
Payment.Show
End Sub

Private Sub rbe_Click()
Room_Browser.Show
End Sub

Private Sub rde_Click()
Room.Show
End Sub
