VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form graph 
   Caption         =   "GRAPH"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4048
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT RAW MATERIAL ID"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
'CODE OF SUBMIT BUTTON
Dim rst1 As New ADODB.Recordset
rst1.Open "select * from raw_material_stock  where raw_material_id='" & Combo1.Text & "' and stock_date>='" & "01-" & Text1.Text & "-2003" & "' order by stock_date", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst1
End Sub

Private Sub Command2_Click()
'code of exit(frame3) button
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub

Private Sub Form_Load()
cnn.Open "DSN=MIS; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from raw_materials", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst.Fields("raw_material_id")
rst.MoveNext
Wend
End Sub
