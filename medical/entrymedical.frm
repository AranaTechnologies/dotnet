VERSION 5.00
Begin VB.Form entrymedical 
   BackColor       =   &H00FFFF80&
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10800
      TabIndex        =   32
      Top             =   5760
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   31
      Top             =   5760
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reference Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7800
      TabIndex        =   21
      Top             =   7080
      Width           =   7095
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3120
         TabIndex        =   25
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3120
         TabIndex        =   23
         Top             =   240
         Width           =   2895
      End
      Begin VB.Image Image3 
         Height          =   1215
         Left            =   120
         Picture         =   "entrymedical.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference date"
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Refered By"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Contec Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   16
      Top             =   7080
      Width           =   7095
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3720
         TabIndex        =   20
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFFF&
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
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   1005
         Left            =   360
         Picture         =   "entrymedical.frx":230C0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   14175
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "entrymedical.frx":27329
         Left            =   9720
         List            =   "entrymedical.frx":27333
         TabIndex        =   42
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         TabIndex        =   35
         Top             =   2880
         Width           =   5655
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Option4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   39
            Top             =   120
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Option3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   38
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label21 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Economic Clinic"
            Height          =   255
            Left            =   3120
            TabIndex        =   37
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pay Clinic"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7560
         TabIndex        =   29
         Top             =   2280
         Width           =   5655
         Begin VB.Label Label19 
            BackColor       =   &H00FFFFC0&
            Caption         =   "New Patient"
            Height          =   255
            Left            =   3120
            TabIndex        =   34
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Old Patient"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFFF&
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
         Left            =   9720
         TabIndex        =   14
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFFF&
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
         Left            =   9720
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
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
         Left            =   9720
         TabIndex        =   9
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   960
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   1470
         Left            =   120
         Picture         =   "entrymedical.frx":27345
         Top             =   1320
         Width           =   1590
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   375
         Left            =   8400
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   375
         Left            =   8400
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF80&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   8400
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14040
      Top             =   2640
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Contac Us: scriptonova@yahoo.com"
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   11160
      TabIndex        =   41
      Top             =   10320
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By: ScriptoNova"
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   11160
      TabIndex        =   40
      Top             =   10080
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Old Patient"
      Height          =   255
      Left            =   9480
      TabIndex        =   30
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "New Patient"
      Height          =   375
      Left            =   9120
      TabIndex        =   28
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Patient"
      Height          =   255
      Left            =   9720
      TabIndex        =   27
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   2415
      Left            =   4200
      Picture         =   "entrymedical.frx":2936A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   26
      Top             =   240
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      Height          =   1335
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   8880
      Width           =   7575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   11880
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "entrymedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
Dim sql As String
Dim ptype As String
Dim ntype As String
cnn.Open "DSN=fromaccess"
If Option1.Value = True Then
ptype = "Old"
End If
If Option2.Value = True Then
ptype = "New"
End If
If Option3.Value = True Then
ntype = "Pay"
End If
If Option4.Value = True Then
ntype = "Economic"
End If
sql = "insert into patient (  date_reg , pt_name , age , sex , address , ph , mobile_no , ptype , ntype ,   refd_by , date_ref , arrival_time ) values (  '" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Combo1.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & ptype & "','" & ntype & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "')"
MsgBox (sql)

cnn.Execute (sql)
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
cnn.Close

End Sub

Private Sub Command2_Click()

'cnn.Close

Unload Me

End Sub

Private Sub Form_Load()
Text2.Text = Date
Text12.Text = Date
Text13.Text = Time



End Sub

Private Sub Timer1_Timer()
Label6.Caption = Time


End Sub
