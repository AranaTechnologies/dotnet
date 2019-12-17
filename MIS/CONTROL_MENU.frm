VERSION 5.00
Begin VB.Form CONTROL_MENU 
   BackColor       =   &H00004080&
   Caption         =   "CONTROL WINDOW"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu PMNU 
      Caption         =   "&Product"
      Begin VB.Menu PEMNU 
         Caption         =   "PRODUCT ENTRY"
      End
      Begin VB.Menu PBMNU 
         Caption         =   "BROWSE PRODUCT"
      End
      Begin VB.Menu PSMNU 
         Caption         =   "PRODUCT STOCK ENTRY"
      End
   End
   Begin VB.Menu RMNU 
      Caption         =   "Raw-&Materials"
      Begin VB.Menu REMNU 
         Caption         =   "RAW MATERIALS ENTRY"
      End
      Begin VB.Menu RBMNU 
         Caption         =   "BROWSE RAW MATERIALS"
      End
      Begin VB.Menu RSEMNU 
         Caption         =   "RAW MATERIALS STOCK ENTRY"
      End
      Begin VB.Menu TFEMNU 
         Caption         =   "TENDER FORM ENTRY"
      End
   End
   Begin VB.Menu CMNU 
      Caption         =   "&Customer"
      Begin VB.Menu CEMNU 
         Caption         =   "CUSTOMER ENTRY"
      End
      Begin VB.Menu CBMNU 
         Caption         =   "BROWSE CUSTOMER"
      End
      Begin VB.Menu COMNU 
         Caption         =   "CUSTOMER ORDER ENTRY"
      End
      Begin VB.Menu BCOMNU 
         Caption         =   "BROWSE CUSTOMER ORDER"
      End
      Begin VB.Menu CHEMNU 
         Caption         =   "CHALLAN ENTRY"
      End
      Begin VB.Menu BCMNU 
         Caption         =   "BROWSE CHALLAN"
      End
      Begin VB.Menu CPEMNU 
         Caption         =   "CUSTOMER PAYMENT ENTRY"
      End
      Begin VB.Menu BCPMNU 
         Caption         =   "BROWSE CUSTOMER PAYMENT"
      End
      Begin VB.Menu pomnu 
         Caption         =   "Process Order"
      End
   End
   Begin VB.Menu VMNU 
      Caption         =   "&Vendor"
      Begin VB.Menu VEMNU 
         Caption         =   "VENDOR ENTRY"
      End
      Begin VB.Menu VBMNU 
         Caption         =   "VENDOR BROWSE"
      End
      Begin VB.Menu POEMNU 
         Caption         =   "PURCHASE ORDER ENTRY"
      End
      Begin VB.Menu BEMNU 
         Caption         =   "BILL ENTRY"
      End
      Begin VB.Menu BBMNU 
         Caption         =   "BROWSE BILL"
      End
      Begin VB.Menu VPEMNU 
         Caption         =   "VENDOR PAYMENT ENTRY"
      End
      Begin VB.Menu STFEMNU 
         Caption         =   "SUBMITTED TENDER FORM ENTRY"
      End
   End
   Begin VB.Menu RSMNU 
      Caption         =   "&Report-section"
      Begin VB.Menu monymnu 
         Caption         =   "Money Reciept"
      End
   End
   Begin VB.Menu QMNU 
      Caption         =   "&Query"
   End
   Begin VB.Menu PASSMNU 
      Caption         =   "Pass&word"
   End
   Begin VB.Menu HMNU 
      Caption         =   "&Help"
   End
   Begin VB.Menu EMNU 
      Caption         =   "E&xit"
   End
   Begin VB.Menu amnu 
      Caption         =   "&About"
      Begin VB.Menu admnu 
         Caption         =   "About the Developer"
      End
      Begin VB.Menu asmnu 
         Caption         =   "About the Software"
      End
   End
End
Attribute VB_Name = "CONTROL_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BBMNU_Click()
BROWSE_BILL.Show
End Sub

Private Sub BCMNU_Click()
BROWSE_CHALLAN.Show
End Sub

Private Sub BCOMNU_Click()
BROWSE_CUSTOMER_ORDER.Show
End Sub

Private Sub BCPMNU_Click()
browse_customer_payment
End Sub

Private Sub BEMNU_Click()
bill.Show
End Sub

Private Sub CBMNU_Click()
BROWSE_CUSTOMER.Show
End Sub

Private Sub CEMNU_Click()
CUSTOMER.Show
End Sub

Private Sub CHEMNU_Click()
challan.Show
End Sub

Private Sub COMNU_Click()
customer_order.Show
End Sub

Private Sub CPEMNU_Click()
customer_payment.Show
End Sub

Private Sub monymnu_Click()
frmmoney_receipt.Show
End Sub

Private Sub PBMNU_Click()
BROWSE_PRODUCT.Show
End Sub

Private Sub PEMNU_Click()
product.Show
End Sub

Private Sub POEMNU_Click()
purchase_order.Show
End Sub

Private Sub pomnu_Click()
frmprocess_order.Show
End Sub

Private Sub PSMNU_Click()
product_stock.Show
End Sub

Private Sub RBMNU_Click()
BROWSE_RAW_MATERIALS.Show
End Sub

Private Sub REMNU_Click()
raw_materials.Show
End Sub

Private Sub RSEMNU_Click()
RAW_MATERIAL_STOCK.Show
End Sub

Private Sub STFEMNU_Click()
SUBMITTED_TENDER_FORM.Show
End Sub

Private Sub TFEMNU_Click()
TENDER_FORM.Show
End Sub

Private Sub VBMNU_Click()
BROWSE_VENDOR.Show
End Sub

Private Sub VEMNU_Click()
VENDOR.Show
End Sub

Private Sub VPEMNU_Click()
vendor_payment.Show
End Sub
