VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Conbook 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Booking1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Text            =   "Text6"
      Top             =   4800
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   5160
      TabIndex        =   8
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4471
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4935
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Text            =   "Text4"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cmbemployee_id 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ROOM_ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "B_Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER_ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "  DAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "MONTH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " YEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   1560
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1680
         Picture         =   "Conbook.frx":0000
         Top             =   2640
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
End
Attribute VB_Name = "Conbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
cnn.Open "DSN=fromoracle; provider=MSDASQL; uid=HTLMS; pwd=HTLMS1"
rst.Open "select * from bookconroom", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
Text1.DataField = "CUSTOMER_ID"
Text2.DataField = "DAY"
Text3.DataField = "MONTH"
Text4.DataField = "YEAR"
Text5.DataField = "TIME"
Text6.DataField = "DURATION"
End Sub
Private Sub Form_Unload(Cancel As Integer)
'rst.Close
cnn.Close
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''TO ADD NEW RECORD
If Button.Index = 1 Then
    If Not rst.BOF Or Not rst.EOF Then
        rst.MoveLast
    End If
        rst.AddNew
End If
''TO SUBMIT THE RECORD
If Button.Index = 2 Then
rst.Update
End If
''TO SEE THE NEXT RECORD
If Button.Index = 3 Then
         If Not rst.BOF Or Not rst.EOF Then
            rst.MovePrevious
         End If
         If rst.BOF Then
            MsgBox ("You are on the First Record")
                If rst.RecordCount <> 0 Then
                        rst.MoveFirst
                End If
                If Not rst.BOF Or Not rst.EOF Then
                        rst.MoveFirst
                End If
          End If
End If
''TO SEE THE PRIVIOUS RECORD
If Button.Index = 4 Then
    If Not rst.BOF Or Not rst.EOF Then
        rst.MoveNext
    End If
    If rst.EOF Then
        MsgBox ("You are on the Last Record")
            If rst.RecordCount <> 0 Then
                  rst.MoveLast
            End If
    End If
End If
''TO DELETE A RECORD
If Button.Index = 5 Then
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
rst.Open "select * from bookconroom", cnn, adOpenStatic, adLockOptimistic, adCmdText
Set Text1.DataSource = rst
Set Text2.DataSource = rst
Set Text3.DataSource = rst
Set Text4.DataSource = rst
Set Text5.DataSource = rst
Set Text6.DataSource = rst
End If
''TO EXIT FROM THE WINDOW
If Button.Index = 6 Then
Unload Me
End If
End Sub



