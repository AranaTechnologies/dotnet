VERSION 5.00
Begin VB.Form Password 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1140
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Password.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Password.frx":08CA
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MouseIcon       =   "Password.frx":1194
      MousePointer    =   99  'Custom
      Picture         =   "Password.frx":149E
      TabIndex        =   2
      Top             =   240
      Width           =   495
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
      ForeColor       =   &H000000FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c%
Dim s$
Option Explicit
Private Sub Form_Load()
s = "HTLMS"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
If UCase(Text1.Text) = UCase(s) Then
Main_Menu.Show
Unload Me
Else
MsgBox "WRONG PSSWORD TRY AGAIN", vbCritical, "SORRY"
c = c + 1
If c = 5 Then
MsgBox "YOU ARE NOT A REAL USER", vbInformation, "DONT DISTURB ME....."
Unload Me
Else
Text1.Text = ""
Text1.SetFocus
End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Text1.Text = ""
Text1.Text = s
End If
End Sub
