VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Control Menu"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image3 
      Height          =   975
      Left            =   240
      Picture         =   "Form 1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   360
      Picture         =   "Form 1.frx":B9340
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   11415
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   1080
      Picture         =   "Form 1.frx":319264
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Menu Patientmnu 
      Caption         =   "&Patient"
      Begin VB.Menu PatientDetailsentrymnu 
         Caption         =   "Patient Details Entry"
      End
      Begin VB.Menu barfield1 
         Caption         =   "-"
      End
      Begin VB.Menu browsepatientdetailsmnu 
         Caption         =   "Browse Patient Details"
      End
      Begin VB.Menu Barfield3 
         Caption         =   "-"
      End
      Begin VB.Menu Testreportmnu 
         Caption         =   "Test Report"
      End
      Begin VB.Menu barfield6mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Employeemnu 
      Caption         =   "Employee"
      Begin VB.Menu Barfield37mnu 
         Caption         =   "-"
      End
      Begin VB.Menu employeeinformationmnu 
         Caption         =   "Employee Information Entry"
      End
      Begin VB.Menu Barfield38mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Browseemployeemnu 
         Caption         =   "Browse Employee Information"
      End
      Begin VB.Menu Barfield39mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Dutyallotmentinformation 
         Caption         =   "Duty Allotment Information"
      End
      Begin VB.Menu barfield7mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Suppliermnu 
      Caption         =   "Supplier"
      Begin VB.Menu supplierinformationentrymnu 
         Caption         =   "Supplier Information Entry"
      End
      Begin VB.Menu barfield8 
         Caption         =   "-"
      End
      Begin VB.Menu browsesupplierinformationmnu 
         Caption         =   "Browse Supplier Information"
      End
      Begin VB.Menu barfield10 
         Caption         =   "-"
      End
      Begin VB.Menu defectiveitemmnu 
         Caption         =   "Defective Item"
         Begin VB.Menu Replacementrequestmnu 
            Caption         =   "Replacement Request"
         End
         Begin VB.Menu barfield28 
            Caption         =   "-"
         End
      End
      Begin VB.Menu barfield12 
         Caption         =   "-"
      End
      Begin VB.Menu Tenderinvitationentrymnu 
         Caption         =   "Tender Invitation Entry"
      End
      Begin VB.Menu barfield53mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Tenderentrymnu 
         Caption         =   "Tender Entry"
      End
      Begin VB.Menu barfield54mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Tenderdateentrymnu 
         Caption         =   "Tender Date Entry"
      End
      Begin VB.Menu Supplieditementrymnu 
         Caption         =   "Supplied Item Entry"
      End
      Begin VB.Menu barfield42 
         Caption         =   "-"
      End
      Begin VB.Menu supplierenlistmentmnu 
         Caption         =   "Supplier Enlistment"
      End
      Begin VB.Menu barfield56mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Complainentrymnu 
         Caption         =   "Complain Entry"
      End
      Begin VB.Menu barfield14mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Inventoryitemmnu 
      Caption         =   "Inventory Item"
      Begin VB.Menu iteminformationentrymnu 
         Caption         =   "Item Information Entry"
      End
      Begin VB.Menu barfield15 
         Caption         =   "-"
      End
      Begin VB.Menu browseiteminformationmnu 
         Caption         =   "Browse Item Information"
      End
      Begin VB.Menu barfield16mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Departmnu 
      Caption         =   "Department"
      Begin VB.Menu departmentinformationentrymnu 
         Caption         =   "Department Information Entry"
      End
      Begin VB.Menu barfield17 
         Caption         =   "-"
      End
      Begin VB.Menu Browsedepartmentinformationmnu 
         Caption         =   "Browse Department Information"
      End
      Begin VB.Menu barfield18mnu 
         Caption         =   "-"
      End
      Begin VB.Menu departmenttestinfomnu 
         Caption         =   "Department_Test_info"
      End
      Begin VB.Menu barfield55mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu testmnu 
      Caption         =   "Test"
      Begin VB.Menu testinformationmnu 
         Caption         =   "Test Information Entry"
      End
      Begin VB.Menu barfield19 
         Caption         =   "-"
      End
      Begin VB.Menu browsetestinformationmnu 
         Caption         =   "Browse Test Information"
      End
      Begin VB.Menu barfield20 
         Caption         =   "-"
      End
      Begin VB.Menu Testrequisitionentrymnu 
         Caption         =   "Test Requisition Entry"
      End
      Begin VB.Menu barfield21 
         Caption         =   "-"
      End
      Begin VB.Menu Testreagententrymnu 
         Caption         =   "Test Reagent Entry"
      End
      Begin VB.Menu barfield22mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Quartermnu 
      Caption         =   "Quarter"
      Begin VB.Menu quarterinformationentrymnu 
         Caption         =   "Quarter Information Entry"
      End
      Begin VB.Menu barfield23 
         Caption         =   "-"
      End
      Begin VB.Menu Browsequarterinformationmnu 
         Caption         =   "Browse Quarter Information"
      End
      Begin VB.Menu barfield24 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Managementmnu 
      Caption         =   "Management"
      Begin VB.Menu Querymnu 
         Caption         =   "Query"
      End
      Begin VB.Menu barfield26 
         Caption         =   "-"
      End
      Begin VB.Menu publishmnu 
         Caption         =   "Publish News"
      End
      Begin VB.Menu barfield34mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Dutyallotmentmnu 
         Caption         =   "Duty Allotment"
         Begin VB.Menu barfield29 
            Caption         =   "-"
         End
         Begin VB.Menu dutyroster 
            Caption         =   "Duty Roster"
         End
         Begin VB.Menu barfield30mnu 
            Caption         =   "-"
         End
      End
      Begin VB.Menu barfield51mnu 
         Caption         =   "-"
      End
      Begin VB.Menu Salaryrulesmnu 
         Caption         =   "Salary Rules"
      End
      Begin VB.Menu barfield41mnu 
         Caption         =   "-"
      End
      Begin VB.Menu taxmnu 
         Caption         =   "Tax"
      End
      Begin VB.Menu barfield52mnu 
         Caption         =   "-"
      End
      Begin VB.Menu EMail 
         Caption         =   "E-Mail"
      End
      Begin VB.Menu barfield40mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Changepasswordmnu 
      Caption         =   "Change Password"
   End
   Begin VB.Menu Aboutus 
      Caption         =   "About Us"
      Begin VB.Menu barfield31 
         Caption         =   "-"
      End
      Begin VB.Menu aboutthedevelopermnu 
         Caption         =   "About The Developer"
      End
      Begin VB.Menu barfield32 
         Caption         =   "-"
      End
      Begin VB.Menu copyrightinformationmnu 
         Caption         =   "Copy Right Information"
      End
      Begin VB.Menu barfield33mnu 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Helpmnu 
      Caption         =   "Help"
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Staffmnu_Click()
End Sub

Private Sub aboutthedevelopermnu_Click()
Form21.Show
End Sub

Private Sub Acknowledgement_Click()
Form39.Show
End Sub

Private Sub Browsedepartmentinformationmnu_Click()
Form10.Show
End Sub

Private Sub Browseemployeemnu_Click()
Form24.Show
End Sub

Private Sub browseiteminformationmnu_Click()
Form16.Show
End Sub

Private Sub browsepatientdetailsmnu_Click()
Form6.Show
End Sub

Private Sub Browsequarterinformationmnu_Click()
Form26.Show
End Sub

Private Sub browsesupplierinformationmnu_Click()
Form25.Show
End Sub

Private Sub browsetestinformationmnu_Click()
Form8.Show

End Sub

Private Sub Browsetestreportmnu_Click()
End Sub

Private Sub Changepasswordmnu_Click()
Form42.Show
End Sub

Private Sub Complainentrymnu_Click()
Form34.Show
End Sub

Private Sub copyrightinformationmnu_Click()
frmAbout.Show
End Sub

Private Sub departmentinformationentrymnu_Click()
Form9.Show
End Sub

Private Sub Dutyallotmentletter_Click()
Form28.Show
End Sub

Private Sub departmenttestinfomnu_Click()
Form17.Show
End Sub

Private Sub Dutyallotmentinformation_Click()
Form28.Show
End Sub

Private Sub dutyroster_Click()
Form27.Show
End Sub

Private Sub EMail_Click()
Form13.Show
End Sub

Private Sub employeeinformationmnu_Click()
Form23.Show
End Sub

Private Sub EXIT_Click()
Dim response As Integer
response = MsgBox("Do you really want to quit the system?", 36, "are you sure?")
If response = 6 Then
End
End If
End Sub

Private Sub Helpmnu_Click()
Form20.Show
End Sub

Private Sub iteminformationentrymnu_Click()
Form15.Show
End Sub

Private Sub MoneyReceiptmnu_Click()
Form6.Show
End Sub

Private Sub PatientDetailsentrymnu_Click()
Form4.Show
End Sub

Private Sub Paymentbillmnu_Click()
Form38.Show
End Sub

Private Sub publishmnu_Click()
HTML.Show

End Sub

Private Sub quarterallotmentinformationmnu_Click()
Form14.Show
End Sub

Private Sub Purchaseordermnu_Click()
Form31.Show
End Sub

Private Sub quarterallotmentletter_Click()

End Sub

Private Sub quarterallotmentlettermnu_Click()
Form26.Show
End Sub

Private Sub quarterinformationentrymnu_Click()
Form14.Show
End Sub
 
Private Sub Querymnu_Click()
Form19.Show
End Sub

Private Sub Replacementrequestmnu_Click()
Form36.Show
End Sub

Private Sub Salaryrulesmnu_Click()
Form29.Show
End Sub

Private Sub SalarySlipmnu_Click()
Form37.Show
End Sub

Private Sub Supplieditementrymnu_Click()
Form41.Show
End Sub

Private Sub supplierenlistmentmnu_Click()
Form18.Show
End Sub

Private Sub supplierinformationentrymnu_Click()
Form3.Show
End Sub

Private Sub taxmnu_Click()
Form30.Show
End Sub

Private Sub Testfeesmnu_Click()
Form17.Show
End Sub

Private Sub Tenderdateentrymnu_Click()
Form40.Show
End Sub

Private Sub Tenderentrymnu_Click()
Form33.Show
End Sub

Private Sub Tenderinvitationentrymnu_Click()
Form32.Show
End Sub

Private Sub testinformationmnu_Click()
Form7.Show

End Sub

Private Sub Testreagententrymnu_Click()
Form11.Show
End Sub

Private Sub Testreportmnu_Click()
Form22.Show
End Sub

Private Sub Testrequisitionentrymnu_Click()
Form5.Show
End Sub
