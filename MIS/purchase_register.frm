VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form purchase_register 
   Caption         =   "PURCHASE REGISTER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d-MMM-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "purchase_register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject

Private Sub Command1_Click()
'code of submit button
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select rs.stock_date,rs.raw_material_id,rs.description,p.purchase_order_no,rs.quantity,s.price,rs.quantity * s.price as Value from raw_material_stock rs,purchase_order_info p,submitted_tender_form s Where upper(rs.transaction_id) = upper(p.purchase_order_no) And upper(p.tender_no) = upper(s.tender_no) and upper(rs.raw_material_id)=upper(s.raw_material_id) and rs.transaction_code='PURCHASED'and rs.stock_date>='" & Text1.Text & "' and rs.stock_date<='" & Text2.Text & "' order by rs.stock_date", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
End Sub

Private Sub Command2_Click()
'code of print button
Dim OUTSTREAM As TextStream
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\mis\reports\purchase_register.DOC", True, False)
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(50) & "PURCHASE REGISTER"
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
OUTSTREAM.WriteLine
prec = Space(5) & "From Date :" & Text1.Text & Space(50) & "To Date :" & Text2.Text
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "Date" & Space(5) & "Item Code" & Space(2) & "Description" & Space(22) & "Order No." & Space(2) & "Quantity" & Space(2) & "Price" & Space(10) & "Value"
OUTSTREAM.WriteLine prec
prec = Space(5) & "----------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
While rst.EOF = False
prec = Space(5) & rst("stock_date") & Space(2) & rst("raw_material_id") & Space(11 - Len(rst("raw_material_id"))) & rst("description") & Space(33 - Len(rst("description"))) & rst("purchase_order_no") & Space(19 - Len(rst("quantity")) - Len(rst("purchase_order_no"))) & rst("quantity") & Space(7 - Len(rst("price"))) & rst("price") & Space(15 - Len(rst("Value"))) & rst("Value")
OUTSTREAM.WriteLine prec
rst.MoveNext
Wend
MsgBox ("Printing is over")
End Sub
