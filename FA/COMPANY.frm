VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "COMPANY BROWSER FORM"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "P&RINT HORIZONTAL"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "P&RINT VERTICAL"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   12648447
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Browser Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim fsys As New FileSystemObject
Private Sub Command1_Click()
Dim response As Integer
response = MsgBox(" Do You Realy Want to Quit ? ", 36, " Are You Sure ? ")
If response = 6 Then
End
End If

End Sub
Private Sub Command2_Click()
Dim OUTSTREAM As TextStream
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\fa\reports\VERTICAL" & ".PRN", True, False)
prec = Space(5) & "...................................................................."
OUTSTREAM.WriteLine prec
While rst.EOF = False
prec = Space(5) & "Company ID." & rst.Fields("company_id")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Name." & rst.Fields("Name")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Alias." & rst.Fields("Alias")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Email Address." & rst.Fields("email_address")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Street." & rst.Fields("Street")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "City." & rst.Fields("City")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Building NO." & rst.Fields("Building_no")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Pin No." & rst.Fields("PIN_NO")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "IT NO." & rst.Fields("IT_NO")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "ST NO." & rst.Fields("ST_NO")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Financial Year From." & rst.Fields("financial_year_from")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "Pass Word." & rst.Fields("pass_word")
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
prec = Space(5) & "...................................................................."
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
rst.MoveNext
Wend
MsgBox ("Printing is Over")
End Sub

Private Sub Command3_Click()
Dim OUTSTREAM As TextStream
Dim prec As String
Set OUTSTREAM = fsys.CreateTextFile("C:\mca\fa\Reports\HORIZONTAL" & ".PRN", True, False)
prec = Space(5) & "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
prec = Space(5) & "COMPANY_ID  NAME                            ALIAS       EMAIL            STREET           CITY    BUILDING_NO    PIN_NO      IT_NO     ST_NO       FINANCIAL_YEAR_FROM       PASS_WORD"
OUTSTREAM.WriteLine prec
prec = Space(5) & "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
While rst.EOF = False
prec = Space(5) & rst.Fields("COMPANY_ID") & Space(12 - Len(rst.Fields("COMPANY_ID"))) & rst.Fields("NAME") & Space(32 - Len(rst.Fields("NAME"))) & rst.Fields("ALIAS") & Space(12 - Len(rst.Fields("ALIAS"))) & rst.Fields("EMAIL_ADDRESS") & Space(17 - Len(rst.Fields("EMAIL_ADDRESS"))) & rst.Fields("STREET") & Space(17 - Len(rst.Fields("STREET"))) & rst.Fields("CITY") & Space(10 - Len(rst.Fields("CITY"))) & rst.Fields("BUILDING_NO") & Space(13 - Len(rst.Fields("BUILDING_NO"))) & rst.Fields("PIN_NO") & Space(12 - Len(rst.Fields("PIN_NO"))) & rst.Fields("IT_NO") & Space(10 - Len(rst.Fields("IT_NO"))) & rst.Fields("ST_NO") & Space(12 - Len(rst.Fields("ST_NO"))) & rst.Fields("FINANCIAL_YEAR_FROM") & Space(26 - Len(rst.Fields("FINANCIAL_YEAR_FROM"))) & rst.Fields("PASS_WORD") & Space(8 - Len(rst.Fields("PASS_WORD")))
OUTSTREAM.WriteLine prec
OUTSTREAM.WriteLine
rst.MoveNext
Wend
prec = Space(5) & "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
OUTSTREAM.WriteLine prec
MsgBox ("PRINTING IS OVER")
End Sub



Private Sub Form_Load()
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=FA; pwd=FA1"
rst.Open "select * from COMPANY", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set DataGrid1.DataSource = rst
End Sub

