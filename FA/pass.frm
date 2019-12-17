VERSION 5.00
Begin VB.Form passfrm 
   BackColor       =   &H00004040&
   Caption         =   "Password"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   Picture         =   "pass.frx":0000
   ScaleHeight     =   2610
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Submit"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Password :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "passfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SUBMIT BITTON
Private Sub Command1_Click()
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=fa; pwd=fa1"
rst.Open "select * from pass", cnn, adOpenStatic, adLockOptimistic, adCmdText
If rst.Fields(0) = Text1.Text Then
Unload Me
Form14.Show
Else
Unload Me
passfrm.Show

End If
End Sub

'CANCEL BUTTON
Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

