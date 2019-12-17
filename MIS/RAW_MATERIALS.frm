VERSION 5.00
Begin VB.Form raw_materials 
   BackColor       =   &H00C0FFFF&
   Caption         =   "RAW MATERIALS INFORMATION ENTRY"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form9"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   25
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   24
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF00&
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "&DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "&NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "&PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "&SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "&ADD NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CARRYING COST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER ORDERING COST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RAW MATERIALS INFORMATION ENTRY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   720
      TabIndex        =   21
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER RAW MATERIAL ID."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER RAW MATERIAL DESCRIPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER UNIT OF MEASUREMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENT STOCK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER LEAD TIME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " REORDER LEVEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ECONOMIC ORDER QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   5280
      Width           =   4095
   End
End
Attribute VB_Name = "raw_materials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
'ADD NEW RECORD BUTTON
Private Sub Command1_Click()
If Not rst.BOF Or Not rst.EOF Then
rst.MoveLast
End If
rst.AddNew
Command1.Enabled = False
Command2.Enabled = True
End Sub
'SUBMIT BUTTON
Private Sub Command2_Click()
rst.Update
Command2.Enabled = False
Command1.Enabled = True
End Sub
'PREVIOUS BUTTON
Private Sub Command3_Click()
rst.MovePrevious
If rst.BOF Then
MsgBox ("You are on the First Record")
If rst.RecordCount <> 0 Then
rst.MoveFirst
End If
End If
End Sub
'NEXT BUTTON
Private Sub Command4_Click()
rst.MoveNext
If rst.EOF Then
MsgBox ("You are on the Last Record")
If rst.RecordCount <> 0 Then
rst.MoveLast
End If
End If
End Sub
'DELETE BUTTON
Private Sub Command5_Click()
Dim response As Integer
Dim message As String
message = "Delete the record of " & UCase(Text1.Text) & "?"
response = MsgBox(message, 36, "Delete Record")
If response = 6 Then
If rst.EOF = True Then
MsgBox ("Eof has occured")
Else
rst.Delete
rst.Update
If Not rst.BOF Or Not rst.EOF Then
If rst.RecordCount > 1 Then
rst.MoveNext
End If
End If

If rst.EOF = True Then
If rst.RecordCount > 1 Then
rst.MovePrevious
End If
End If
End If
End If
rst.Close
rst.Open "select * from raw_materials", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Set Text7.DataSource = rst
End Sub


Private Sub Command7_Click()
'code of cancel button
rst.CancelUpdate
Text1.Text = rst("raw_material_id")
Text2.Text = rst("description")
Text3.Text = rst("unit_of_measurement")
Text4.Text = rst("current_stock")
Text5.Text = rst("lead_time")
Text6.Text = rst("reorder_level")
Text7.Text = rst("eoq")
End Sub

'EXIT BUTTON
Private Sub Command8_Click()
rst.Close
cnn.Close
Unload Me
CONTROL_MENU.Show
End Sub
Private Sub Form_Load()
Dim rst1 As New ADODB.Recordset
Dim lead_time As Double
Dim reorder_level, avg_demand As Integer
Dim eoq As Double
Dim o_cost As Integer
Dim C_cost As Integer
cnn.Open "DSN=mis; provider=MSDASQL; uid=mis; pwd=mis1"
rst.Open "select * from raw_materials", cnn, adOpenStatic, adLockOptimistic, adCmdText

While rst.EOF = False
sql = "select sum(quantity) from raw_material_stock where raw_material_id='" & rst("raw_material_id") & "' and transaction_code='ISSUED'"
Set rst1 = cnn.Execute(sql)
avg_demand = rst1("sum(quantity)") / 6
lead_time = rst("lead_time") / 30
reorder_level = lead_time * avg_demand
o_cost = rst.Fields("O_COST")
C_cost = rst.Fields("C_COST")

If rst("reorder_level") = Null Then
sql = "update raw_materials set reorder_level=" & reorder_level & " where raw_material_id='" & rst("raw_material_id") & "'"
Set r = cnn.Execute(sql)
End If
MsgBox (avg_demand)
MsgBox (o_cost)
MsgBox (C_cost)
If Len(Text7.Text) = 0 Then
'Text7.Text = str(((2 * o_cost * avg_demand) / C_cost))
MsgBox (((2 * o_cost * avg_demand)))
eoq = Sqr((2 * o_cost * avg_demand * 1.1) / C_cost)
sql = "update raw_materials set eoq=" & eoq & " where raw_material_id='" & rst1("raw_material_id") & "'"
Set r = cnn.Execute(sql)
End If
rst.MoveNext
Wend
rst.Close
rst.Open "select * from raw_materials", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Set Text7.DataSource = rst
Set Text8.DataSource = rst
Set Text9.DataSource = rst

Text1.DataField = "raw_material_id"
Text2.DataField = "description"
Text3.DataField = "unit_of_measurement"
Text4.DataField = "current_stock"
Text5.DataField = "lead_time"
Text6.DataField = "reorder_level"
Text7.DataField = "eoq"
Text8.DataField = "o_cost"
Text9.DataField = "c_cost"
End Sub


