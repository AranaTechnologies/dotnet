VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4065
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Display Settings - Screen Area ( 640 by 480 ) pixels "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Presented By MOU  MAITEE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   3690
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hotel Management System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3285
         Left            =   240
         TabIndex        =   4
         Top             =   -120
         Width           =   6180
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform - WINDOWS-95 / 98 / ME"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   3360
         Width           =   4080
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version - 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   2880
         Width           =   1410
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Warning.... DONT COPY RIGHT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   3660
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   9000
         Left            =   0
         Picture         =   "FRONTPage.frx":0000
         Top             =   -2160
         Width           =   12000
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t%

Private Sub Form_DblClick()
    Password.Show
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Password.Show
    Unload Me
End Sub

Private Sub Form_Load()
        lblVersion.Caption = "Version " & "-" & " " & "1.0"
    lblProductName.Caption = "Hotel Management System"
End Sub
Private Sub Frame1_DblClick()
    Password.Show
    Unload Me
End Sub

Private Sub Timer1_Timer()
t = t + 1
If t = 5 Then
Timer1.Enabled = False
Password.Show
Unload Me
End If
End Sub
