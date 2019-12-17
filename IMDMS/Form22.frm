VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H0080C0FF&
   Caption         =   "Department Login"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O K"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Department Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Dim CNN As New ADODB.Connection
Dim RST As New ADODB.Recordset
Private Sub Command1_Click()
Form34.Label1.Caption = Combo1.Text
Form34.Show

Unload Me

End Sub

Private Sub Form_Load()
Dim sql As String
CNN.Open "DSN=from oracle; provider=MSDASQL; uid=imdms; pwd=imdms1"
RST.Open "select * from DEPARTMENT", CNN, adOpenStatic, adLockOptimistic, adCmdText
While Not RST.EOF
Combo1.AddItem RST.Fields("DEPARTMENT_NAME")
RST.MoveNext
Wend
RST.Close




'sql = "select * from Department_Test_info where Department_name='" & Label1.Caption & "'"
' , adOpenStatic, adLockOptimistic, adCmdText
End Sub

