VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form s11commercemarks 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   15540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "DELEATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   24
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   23
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "INSERT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   22
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      TabIndex        =   21
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Search Student"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Caption         =   "Student Marks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2400
      TabIndex        =   8
      Top             =   4560
      Width           =   10815
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   2760
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   2760
         TabIndex        =   11
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   615
         Left            =   8640
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   615
         Left            =   8640
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "     MATH :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   18
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "BUSINESS ST :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   2280
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "   ACCOUNTS :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "ECONOMICS :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   15
         Top             =   1320
         Width           =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "                ENGLISH :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         TabIndex        =   14
         Top             =   2640
         Width           =   3060
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C000C0&
      Caption         =   "Student Details :"
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
      Left            =   7200
      TabIndex        =   3
      Top             =   2880
      Width           =   7815
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "D.O.B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4680
         TabIndex        =   6
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label9 
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label10 
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Select Exam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   2400
      TabIndex        =   1
      Top             =   2880
      Width           =   4695
      Begin VB.ComboBox exam 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   2
         Text            =   "SELECT EXAM NAME"
         Top             =   480
         Width           =   4095
      End
   End
   Begin VB.ComboBox sid 
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
      Left            =   120
      TabIndex        =   0
      Text            =   "STUDENT ID"
      Top             =   2160
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   480
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database3.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database3.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   240
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database1.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database1.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   2130
      Left            =   2160
      Picture         =   "s11commercemarks.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   13140
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   26
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "11 cOMMERCE MARKS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   5400
      TabIndex        =   25
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "s11commercemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Private Sub Command10_Click()
rs2.MoveFirst
rs2.Find ("XENTRY='" & CStr(Text7.Text) & "'")
If rs2.BOF = True Or rs2.EOF = True Then
MsgBox "DATA NOT FOUND", vbCritical
Else
rs2.Delete
MsgBox "Current Data Has Been Deleated", vbInformation
rs2.Update
End If
End Sub

Private Sub Command5_Click()
rs1.MoveFirst
rs1.Find ("SENTRY='" & CStr(Text7.Text) & "'")
If (rs1.EOF = True) Then
    Label9.Caption = ""
    Label10.Caption = ""
    MsgBox "PLEASE ENTER A VALID ID", vbCritical
Else
Label9.Caption = rs1.Fields("STU NAME").Value
Label10.Caption = rs1.Fields("STU DOB").Value
End If

If rs2.BOF = True Then
    MsgBox "NO any DATA FOUND PLAS INSERT DATA", vbExclamation
Else
    rs2.MoveFirst
    rs2.Find ("XENTRY='" & CStr(Text7.Text) & "'")
    If rs2.BOF = True Or rs2.EOF = True Then
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    Else
    Text3.Text = rs2.Fields("SUBJECT1").Value
    Text4.Text = rs2.Fields("SUBJECT2").Value
    Text5.Text = rs2.Fields("SUBJECT3").Value
    Text6.Text = rs2.Fields("SUBJECT4").Value
    Text8.Text = rs2.Fields("SUBJECT5").Value
    End If
End If
End Sub

Private Sub Command7_Click()
s11sciencemarks.Hide
sadminoption.Show
End Sub

Private Sub Command8_Click()
rs2.AddNew
rs1.Find ("SENTRY='" & CStr(Text7.Text) & "'")
rs2.Fields("XNAME").Value = rs1.Fields("STU NAME").Value
rs2.Fields("XDOB").Value = rs1.Fields("STU DOB").Value
rs2.Fields("XCLASS").Value = rs1.Fields("STU CLASS").Value
rs2.Fields("XSTREAM").Value = rs1.Fields("STU STREAM").Value
rs2.Fields("XENTRY").Value = rs1.Fields("SENTRY").Value
rs2.Fields("SUBJECT1").Value = Text3.Text
rs2.Fields("SUBJECT2").Value = Text4.Text
rs2.Fields("SUBJECT3").Value = Text5.Text
rs2.Fields("SUBJECT4").Value = Text6.Text
rs2.Fields("SUBJECT5").Value = Text8.Text
MsgBox "Marks Has Been Updated Successfully", vbInformation
rs2.Update
End Sub

Private Sub Command9_Click()
rs2.MoveFirst
rs2.Find ("XENTRY='" & CStr(Text7.Text) & "'")
If rs2.BOF = True Or rs2.EOF = True Then
MsgBox "DATA NOT FOUND FOR UPDATE", vbCritical
Else
rs2.Fields("SUBJECT1").Value = Text3.Text
rs2.Fields("SUBJECT2").Value = Text4.Text
rs2.Fields("SUBJECT3").Value = Text5.Text
rs2.Fields("SUBJECT4").Value = Text6.Text
rs2.Fields("SUBJECT5").Value = Text8.Text
MsgBox "Marks Has Been Updated Successfully", vbInformation
rs2.Update
End If

End Sub

Private Sub exam_Click()
If exam.Text = "SUMMATIVE ASSESSMENT-2" Then
cn2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database3.accdb;Persist Security Info=False"
cn2.Open
rs2.ActiveConnection = cn2
rs2.CursorType = adOpenDynamic
rs2.LockType = adLockOptimistic
rs2.Source = "stusa2"
rs2.Open
End If
If exam.Text = "SUMMATIVE ASSESSMENT-1" Then
cn2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database3.accdb;Persist Security Info=False"
cn2.Open
rs2.ActiveConnection = cn2
rs2.CursorType = adOpenDynamic
rs2.LockType = adLockOptimistic
rs2.Source = "stusa1"
rs2.Open
End If
End Sub

Private Sub Form_Load()
cn1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database1.accdb;Persist Security Info=False"
cn1.Open
rs1.ActiveConnection = cn1
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
rs1.Source = "11commerce"
rs1.Open

exam.AddItem ("SUMMATIVE ASSESSMENT-1")
exam.AddItem ("SUMMATIVE ASSESSMENT-2")

'cn2.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database3.accdb;Persist Security Info=False"
'cn2.Open
'rs2.ActiveConnection = cn2
'rs2.CursorType = adOpenDynamic
'rs2.LockType = adLockOptimistic
'rs2.Source = "stusa1"
'rs2.Open
Do While rs1.EOF = False
sid.AddItem (rs1.Fields("SENTRY").Value)
rs1.MoveNext
Loop
rs1.MoveFirst
End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub sid_Click()
Text7.Text = sid.Text
rs1.MoveFirst
rs1.Find ("SENTRY='" & CStr(Text7.Text) & "'")

Label9.Caption = rs1.Fields("STU NAME").Value
Label10.Caption = rs1.Fields("STU DOB").Value

If rs2.BOF = True Then
    MsgBox "NO any DATA FOUND PLAS INSERT DATA", vbExclamation
Else
    rs2.MoveFirst
    rs2.Find ("XENTRY='" & CStr(Text7.Text) & "'")
    If rs2.BOF = True Or rs2.EOF = True Then
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    Else
    Text3.Text = rs2.Fields("SUBJECT1").Value
    Text4.Text = rs2.Fields("SUBJECT2").Value
    Text5.Text = rs2.Fields("SUBJECT3").Value
    Text6.Text = rs2.Fields("SUBJECT4").Value
    Text8.Text = rs2.Fields("SUBJECT5").Value
    End If
End If
End Sub

