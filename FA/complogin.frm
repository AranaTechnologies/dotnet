VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00004000&
   Caption         =   "Company Login"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6165
   LinkTopic       =   "Form12"
   ScaleHeight     =   8535
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Submit"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
Dim rst2 As New ADODB.Recordset
rst2.Open "select company_id from COMPANY where name='" & Combo1.Text & "'", cnn, adOpenStatic, adLockOptimistic, adCmdText
Form11.Label2.Caption = Combo1.Text
Form15.Text1.Text = rst2.Fields("company_id")
Form15.Text2.Text = Combo1.Text
'Form15.Show
rst2.Close
rst.Close
cnn.Close
Form11.Show
Unload Me

End Sub

Private Sub Command2_Click()
rst.Close
cnn.Close
Unload Me

End Sub

Private Sub Form_Load()
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=FA; pwd=FA1"
rst.Open "select * from COMPANY", cnn, adOpenStatic, adLockOptimistic, adCmdText
While rst.EOF = False
Combo1.AddItem rst.Fields("name")
rst.MoveNext
Wend
rst.MoveFirst
End Sub
