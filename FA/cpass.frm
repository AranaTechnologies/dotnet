VERSION 5.00
Begin VB.Form cpassfrm 
   BackColor       =   &H00004040&
   Caption         =   "Change Password"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6180
   LinkTopic       =   "Form8"
   Picture         =   "cpass.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Submit"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the new Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The old Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "cpassfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=fa; pwd=fa1"
rst.Open "select * from pass", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text2.DataSource = rst

If rst.Fields("word") = Text1.Text Then
If Len(Text2.Text) <> 0 Then
rst.Fields("word") = Text2.Text
rst.Update
MsgBox ("Password has been Changed ")
Unload Me
Else
MsgBox ("New Password can not empty ")
End If
Else
MsgBox ("You have entered WORNG old Password ")
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

