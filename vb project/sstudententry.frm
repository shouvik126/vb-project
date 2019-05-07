VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sstudententry 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      TabIndex        =   40
      Top             =   7800
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   8880
      TabIndex        =   39
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox board 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8880
      TabIndex        =   38
      Text            =   "Select Board"
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   6480
      TabIndex        =   36
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   6960
      Width           =   5415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      MaskColor       =   &H0000FF00&
      TabIndex        =   32
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   31
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      MaskColor       =   &H000000C0&
      TabIndex        =   30
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13080
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   11640
      ScaleHeight     =   2355
      ScaleWidth      =   2835
      TabIndex        =   29
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Select the (class/stream) u want to study"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   11415
      Begin VB.ComboBox stream 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         TabIndex        =   28
         Text            =   "Select Stream"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox class 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   26
         Text            =   "Select Class"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Stream :-"
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
         Left            =   6960
         TabIndex        =   27
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Class :-"
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
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   8040
      TabIndex        =   23
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2280
      TabIndex        =   21
      Top             =   5760
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   5160
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   11160
      TabIndex        =   15
      Top             =   3000
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   110755841
      CurrentDate     =   42859
   End
   Begin VB.ComboBox gender 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   11
      Text            =   "Select Gender"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox caste 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   9
      Text            =   "Select Caste"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   9600
      TabIndex        =   7
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   5655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2160
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\vb project\database\Database1.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\vb project\database\Database1.accdb;Persist Security Info=False"
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Board/University"
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
      Left            =   8880
      TabIndex        =   37
      Top             =   6600
      Width           =   2280
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Marks(%)"
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
      Left            =   6600
      TabIndex        =   35
      Top             =   6480
      Width           =   1290
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Name of School/College(Last Attended)"
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
      Left            =   120
      TabIndex        =   33
      Top             =   6480
      Width           =   5535
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Occupation :-"
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
      Left            =   6000
      TabIndex        =   22
      Top             =   5880
      Width           =   1905
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Mother Name :-"
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
      Left            =   0
      TabIndex        =   20
      Top             =   5880
      Width           =   2190
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Occupation :-"
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
      Left            =   6000
      TabIndex        =   18
      Top             =   5400
      Width           =   1905
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Father Name :-"
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
      Left            =   0
      TabIndex        =   16
      Top             =   5280
      Width           =   2115
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Mobile :-"
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
      Left            =   9720
      TabIndex        =   14
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "D.O.B :-"
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
      Left            =   6600
      TabIndex        =   12
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Gender :-"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Caste :-"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Email :-"
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
      Left            =   8400
      TabIndex        =   6
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Address :-"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Student Name :-"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   2280
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Entry No :-"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   10680
      Picture         =   "sstudententry.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Caption         =   "STUDENT ADMISSION"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   360
      Width           =   5205
   End
End
Attribute VB_Name = "sstudententry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim no As Integer
Dim str As String
Dim max As Integer
Dim cn As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
Dim rs7 As New ADODB.Recordset

Private Sub Command1_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Command3_Click()
If (class.Text = "11" And stream.Text = "SCIENCE") Then
    rs1.AddNew
    rs1.Fields("STU NAME").Value = Text1.Text
    rs1.Fields("STU ADDRESS").Value = Text2.Text
    rs1.Fields("STU EMAIL").Value = Text3.Text
    rs1.Fields("STU CAST").Value = caste.Text
    rs1.Fields("STU GENDER").Value = gender.Text
    rs1.Fields("STU DOB").Value = DTPicker1.Value
    rs1.Fields("STU MOBILE").Value = Text4.Text
    rs1.Fields("STU CLASS").Value = class.Text
    rs1.Fields("STU STREAM").Value = stream.Text
    rs1.Fields("FATHER NAME").Value = Text5.Text
    rs1.Fields("FATHER OCCUPATION").Value = Text6.Text
    rs1.Fields("MOTHER NAME").Value = Text7.Text
    rs1.Fields("MOTHER OCCUPATION").Value = Text8.Text
    rs1.Fields("SCHOOL(LAST ATTEND)").Value = Text9.Text
    rs1.Fields("MARKS").Value = Text10.Text
    rs1.Fields("BOARD").Value = board.Text
    rs1.Fields("PICTURE").Value = str
    rs1.Fields("SENTRY").Value = Text11.Text
    MsgBox "Data is saved successfully", vbInformation
    rs1.Update
    
    rs2.AddNew
    rs2.Fields("STU ENTRYNO").Value = Val(Text11.Text)
    rs2.Fields("STU ID").Value = Val(Text11.Text)
    rs2.Fields("STU PASSWORD").Value = Text4.Text
    MsgBox "YOUR ID=" & rs2.Fields("STU ID").Value & "| PASSWORD=" & rs2.Fields("STU PASSWORD").Value, vbInformation
    rs2.Update
End If

If (class.Text = "11" And stream.Text = "ARTS") Then
    rs3.AddNew
    rs3.Fields("STU NAME").Value = Text1.Text
    rs3.Fields("STU ADDRESS").Value = Text2.Text
    rs3.Fields("STU EMAIL").Value = Text3.Text
    rs3.Fields("STU CAST").Value = caste.Text
    rs3.Fields("STU GENDER").Value = gender.Text
    rs3.Fields("STU DOB").Value = DTPicker1.Value
    rs3.Fields("STU MOBILE").Value = Text4.Text
    rs3.Fields("STU CLASS").Value = class.Text
    rs3.Fields("STU STREAM").Value = stream.Text
    rs3.Fields("FATHER NAME").Value = Text5.Text
    rs3.Fields("FATHER OCCUPATION").Value = Text6.Text
    rs3.Fields("MOTHER NAME").Value = Text7.Text
    rs3.Fields("MOTHER OCCUPATION").Value = Text8.Text
    rs3.Fields("SCHOOL(LAST ATTEND)").Value = Text9.Text
    rs3.Fields("MARKS").Value = Text10.Text
    rs3.Fields("BOARD").Value = board.Text
    rs3.Fields("PICTURE").Value = str
    rs3.Fields("SENTRY").Value = Text11.Text
    MsgBox "Data is saved successfully", vbInformation
    rs3.Update
        
    rs2.AddNew
    rs2.Fields("STU ENTRYNO").Value = Val(Text11.Text)
    rs2.Fields("STU ID").Value = Val(Text11.Text)
    rs2.Fields("STU PASSWORD").Value = Text4.Text
    MsgBox "YOUR ID=" & rs2.Fields("STU ID").Value & "| PASSWORD=" & rs2.Fields("STU PASSWORD").Value, vbInformation
    rs2.Update
End If

If (class.Text = "11" And stream.Text = "COMMERSE") Then
    rs4.AddNew
    rs4.Fields("STU NAME").Value = Text1.Text
    rs4.Fields("STU ADDRESS").Value = Text2.Text
    rs4.Fields("STU EMAIL").Value = Text3.Text
    rs4.Fields("STU CAST").Value = caste.Text
    rs4.Fields("STU GENDER").Value = gender.Text
    rs4.Fields("STU DOB").Value = DTPicker1.Value
    rs4.Fields("STU MOBILE").Value = Text4.Text
    rs4.Fields("STU CLASS").Value = class.Text
    rs4.Fields("STU STREAM").Value = stream.Text
    rs4.Fields("FATHER NAME").Value = Text5.Text
    rs4.Fields("FATHER OCCUPATION").Value = Text6.Text
    rs4.Fields("MOTHER NAME").Value = Text7.Text
    rs4.Fields("MOTHER OCCUPATION").Value = Text8.Text
    rs4.Fields("SCHOOL(LAST ATTEND)").Value = Text9.Text
    rs4.Fields("MARKS").Value = Text10.Text
    rs4.Fields("BOARD").Value = board.Text
    rs4.Fields("PICTURE").Value = str
    rs4.Fields("SENTRY").Value = Text11.Text
    MsgBox "Data is saved successfully", vbInformation
    rs4.Update
        
    rs2.AddNew
    rs2.Fields("STU ENTRYNO").Value = Val(Text11.Text)
    rs2.Fields("STU ID").Value = Val(Text11.Text)
    rs2.Fields("STU PASSWORD").Value = Text4.Text
    MsgBox "YOUR ID=" & rs2.Fields("STU ID").Value & "| PASSWORD=" & rs2.Fields("STU PASSWORD").Value, vbInformation
    rs2.Update
End If

If (class.Text = "12" And stream.Text = "SCIENCE") Then
    rs5.AddNew
    rs5.Fields("STU NAME").Value = Text1.Text
    rs5.Fields("STU ADDRESS").Value = Text2.Text
    rs5.Fields("STU EMAIL").Value = Text3.Text
    rs5.Fields("STU CAST").Value = caste.Text
    rs5.Fields("STU GENDER").Value = gender.Text
    rs5.Fields("STU DOB").Value = DTPicker1.Value
    rs5.Fields("STU MOBILE").Value = Text4.Text
    rs5.Fields("STU CLASS").Value = class.Text
    rs5.Fields("STU STREAM").Value = stream.Text
    rs5.Fields("FATHER NAME").Value = Text5.Text
    rs5.Fields("FATHER OCCUPATION").Value = Text6.Text
    rs5.Fields("MOTHER NAME").Value = Text7.Text
    rs5.Fields("MOTHER OCCUPATION").Value = Text8.Text
    rs5.Fields("SCHOOL(LAST ATTEND)").Value = Text9.Text
    rs5.Fields("MARKS").Value = Text10.Text
    rs5.Fields("BOARD").Value = board.Text
    rs5.Fields("PICTURE").Value = str
    rs5.Fields("SENTRY").Value = Text11.Text
    MsgBox "Data is saved successfully", vbInformation
    rs5.Update
        
    rs2.AddNew
    rs2.Fields("STU ENTRYNO").Value = Val(Text11.Text)
    rs2.Fields("STU ID").Value = Val(Text11.Text)
    rs2.Fields("STU PASSWORD").Value = Text4.Text
    MsgBox "YOUR ID=" & rs2.Fields("STU ID").Value & "| PASSWORD=" & rs2.Fields("STU PASSWORD").Value, vbInformation
    rs2.Update
End If

If (class.Text = "12" And stream.Text = "ARTS") Then
    rs6.AddNew
    rs6.Fields("STU NAME").Value = Text1.Text
    rs6.Fields("STU ADDRESS").Value = Text2.Text
    rs6.Fields("STU EMAIL").Value = Text3.Text
    rs6.Fields("STU CAST").Value = caste.Text
    rs6.Fields("STU GENDER").Value = gender.Text
    rs6.Fields("STU DOB").Value = DTPicker1.Value
    rs6.Fields("STU MOBILE").Value = Text4.Text
    rs6.Fields("STU CLASS").Value = class.Text
    rs6.Fields("STU STREAM").Value = stream.Text
    rs6.Fields("FATHER NAME").Value = Text5.Text
    rs6.Fields("FATHER OCCUPATION").Value = Text6.Text
    rs6.Fields("MOTHER NAME").Value = Text7.Text
    rs6.Fields("MOTHER OCCUPATION").Value = Text8.Text
    rs6.Fields("SCHOOL(LAST ATTEND)").Value = Text9.Text
    rs6.Fields("MARKS").Value = Text10.Text
    rs6.Fields("BOARD").Value = board.Text
    rs6.Fields("PICTURE").Value = str
    rs6.Fields("SENTRY").Value = Text11.Text
    MsgBox "Data is saved successfully", vbInformation
    rs6.Update
        
    rs2.AddNew
    rs2.Fields("STU ENTRYNO").Value = Val(Text11.Text)
    rs2.Fields("STU ID").Value = Val(Text11.Text)
    rs2.Fields("STU PASSWORD").Value = Text4.Text
    MsgBox "YOUR ID=" & rs2.Fields("STU ID").Value & "| PASSWORD=" & rs2.Fields("STU PASSWORD").Value, vbInformation
    rs2.Update
End If

If (class.Text = "12" And stream.Text = "COMMERSE") Then
    rs7.AddNew
    rs7.Fields("STU NAME").Value = Text1.Text
    rs7.Fields("STU ADDRESS").Value = Text2.Text
    rs7.Fields("STU EMAIL").Value = Text3.Text
    rs7.Fields("STU CAST").Value = caste.Text
    rs7.Fields("STU GENDER").Value = gender.Text
    rs7.Fields("STU DOB").Value = DTPicker1.Value
    rs7.Fields("STU MOBILE").Value = Text4.Text
    rs7.Fields("STU CLASS").Value = class.Text
    rs7.Fields("STU STREAM").Value = stream.Text
    rs7.Fields("FATHER NAME").Value = Text5.Text
    rs7.Fields("FATHER OCCUPATION").Value = Text6.Text
    rs7.Fields("MOTHER NAME").Value = Text7.Text
    rs7.Fields("MOTHER OCCUPATION").Value = Text8.Text
    rs7.Fields("SCHOOL(LAST ATTEND)").Value = Text9.Text
    rs7.Fields("MARKS").Value = Text10.Text
    rs7.Fields("BOARD").Value = board.Text
    rs7.Fields("PICTURE").Value = str
    rs7.Fields("SENTRY").Value = Text11.Text
    MsgBox "Data is saved successfully", vbInformation
    rs7.Update
    
    rs2.AddNew
    rs2.Fields("STU ENTRYNO").Value = Val(Text11.Text)
    rs2.Fields("STU ID").Value = Val(Text11.Text)
    rs2.Fields("STU PASSWORD").Value = Text4.Text
    MsgBox "YOUR ID=" & rs2.Fields("STU ID").Value & "| PASSWORD=" & rs2.Fields("STU PASSWORD").Value, vbInformation
    rs2.Update
End If
End Sub

Private Sub Command4_Click()
sstudententry.Hide
smainpage.Show
End Sub

Private Sub Form_Load()
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database1.accdb;Persist Security Info=False"
cn.Open
rs1.ActiveConnection = cn
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
rs1.Source = "11science"
rs1.Open
rs2.ActiveConnection = cn
rs2.CursorType = adOpenDynamic
rs2.LockType = adLockOptimistic
rs2.Source = "stulogin"
rs2.Open
rs3.ActiveConnection = cn
rs3.CursorType = adOpenDynamic
rs3.LockType = adLockOptimistic
rs3.Source = "11arts"
rs3.Open
rs4.ActiveConnection = cn
rs4.CursorType = adOpenDynamic
rs4.LockType = adLockOptimistic
rs4.Source = "11commerce"
rs4.Open
rs5.ActiveConnection = cn
rs5.CursorType = adOpenDynamic
rs5.LockType = adLockOptimistic
rs5.Source = "12science"
rs5.Open
rs6.ActiveConnection = cn
rs6.CursorType = adOpenDynamic
rs6.LockType = adLockOptimistic
rs6.Source = "12arts"
rs6.Open
rs7.ActiveConnection = cn
rs7.CursorType = adOpenDynamic
rs7.LockType = adLockOptimistic
rs7.Source = "12commerce"
rs7.Open
board.AddItem ("STATE BOARD")
board.AddItem ("IB")
board.AddItem ("ICSE")
board.AddItem ("CBSE")

stream.AddItem ("COMMERSE")
stream.AddItem ("ARTS")
stream.AddItem ("SCIENCE")
class.AddItem ("12")
class.AddItem ("11")
caste.AddItem ("GENERAL")
caste.AddItem ("SC")
caste.AddItem ("ST")
caste.AddItem ("OBC")
caste.AddItem ("OTHERS")

gender.AddItem ("MALE")
gender.AddItem ("FEMALE")
gender.AddItem ("OTHERs")

rs2.MoveFirst
max = 0
Do While rs2.EOF = False
If (rs2.Fields("STU ENTRYNO").Value > max) Then
max = rs2.Fields("STU ENTRYNO").Value
End If
rs2.MoveNext
Loop
Text11.Text = max + 1
End Sub
