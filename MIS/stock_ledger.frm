VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form stock_ledger 
   Caption         =   "Form5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   8895
      _ExtentX        =   15690
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
   Begin VB.CommandButton command1 
      Caption         =   "SUBMIT"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "stock_ledger.frx":0000
      Left            =   3480
      List            =   "stock_ledger.frx":0028
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT THE MONTH"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "stock_ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub Command1_Click()
'CODE OF SUBMIT BUTTON
rst.Open "select * from raw_materials", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
rst1.Open "select * from raw_material_stock where raw_material_id='" & rst("raw_material_id") & "' and stock_date>='" & "01-" & Combo1.Text & "-2003" & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText

End Sub
