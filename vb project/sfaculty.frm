VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form sfaculty 
   BackColor       =   &H0080FFFF&
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   240
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database2.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database2.accdb;Persist Security Info=False"
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
   Begin VB.CommandButton Command10 
      Caption         =   "DELETE"
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
      Left            =   12960
      TabIndex        =   35
      Top             =   6960
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
      Left            =   12960
      TabIndex        =   34
      Top             =   6000
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
      Left            =   12960
      TabIndex        =   33
      Top             =   5040
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
      Left            =   12960
      TabIndex        =   32
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh"
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
      TabIndex        =   31
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Search Faculty"
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
      TabIndex        =   30
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   1215
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Nevigate"
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
      Left            =   2400
      TabIndex        =   23
      Top             =   6600
      Width           =   10335
      Begin VB.CommandButton Command4 
         Caption         =   "LAST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   27
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   26
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FIRST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2280
      TabIndex        =   16
      Top             =   4200
      Width           =   10455
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   22
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   20
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Mob NO :"
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
         Left            =   6960
         TabIndex        =   21
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Email :"
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
         Left            =   3000
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Address :"
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
         Top             =   480
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "General Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   12735
      Begin VB.CommandButton Command11 
         Caption         =   "UPLOAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   37
         Top             =   2400
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         Height          =   2055
         Left            =   10200
         ScaleHeight     =   1995
         ScaleWidth      =   2355
         TabIndex        =   36
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox subject 
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
         Left            =   6360
         TabIndex        =   15
         Text            =   "SELECT SUBJECT"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   13
         Top             =   1800
         Width           =   3015
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
         Left            =   240
         TabIndex        =   11
         Text            =   "SELECT GENDER"
         Top             =   1800
         Width           =   2175
      End
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
         Left            =   7680
         TabIndex        =   9
         Text            =   "SELECT STREAM"
         Top             =   720
         Width           =   2295
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
         Left            =   5400
         TabIndex        =   7
         Text            =   "SELECT CLASS"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Allocate Subject :"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   2460
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Highest Qualification :"
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
         Left            =   2640
         TabIndex        =   12
         Top             =   1440
         Width           =   3075
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Gender :"
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
         TabIndex        =   10
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Allocate Stream :"
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
         Left            =   7680
         TabIndex        =   8
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Allocate Class :"
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
         TabIndex        =   6
         Top             =   360
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "Age :"
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
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
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
         TabIndex        =   2
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Mob NO"
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
      Left            =   480
      TabIndex        =   38
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "Enter To Search :"
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
      Left            =   0
      TabIndex        =   28
      Top             =   2160
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "FACULTY MEMBER DETAILS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "sfaculty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim cn As New ADODB.Connection
Dim rs1 As New ADODB.Recordset

Private Sub Command1_Click()
rs1.MoveFirst
Text1.Text = (rs1.Fields("FNAME").Value)
Text2.Text = (rs1.Fields("FAGE").Value)
class.Text = (rs1.Fields("FCLASS").Value)
stream.Text = (rs1.Fields("FSTREAM").Value)
gender.Text = (rs1.Fields("FGENDER").Value)
Text3.Text = (rs1.Fields("FQUALIFICATION").Value)
subject.Text = (rs1.Fields("FSUBJECT").Value)
Text4.Text = (rs1.Fields("FADDRESS").Value)
Text5.Text = (rs1.Fields("FEMAIL").Value)
Text6.Text = (rs1.Fields("FMOBILE").Value)
Picture1.Picture = LoadPicture(rs1!FPHOTO)
End Sub

Private Sub Command10_Click()
'rs1.Find ("FMOBILE='" + CStr(Text7.Text) + "'")
rs1.Delete
MsgBox "DATA HA BEEN DELEATED", vbInformation
rs1.Update
End Sub

Private Sub Command11_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub

Private Sub Command2_Click()
rs1.MovePrevious
If rs1.BOF = True Then
rs1.MoveLast
End If
Text1.Text = (rs1.Fields("FNAME").Value)
Text2.Text = (rs1.Fields("FAGE").Value)
class.Text = (rs1.Fields("FCLASS").Value)
stream.Text = (rs1.Fields("FSTREAM").Value)
gender.Text = (rs1.Fields("FGENDER").Value)
Text3.Text = (rs1.Fields("FQUALIFICATION").Value)
subject.Text = (rs1.Fields("FSUBJECT").Value)
Text4.Text = (rs1.Fields("FADDRESS").Value)
Text5.Text = (rs1.Fields("FEMAIL").Value)
Text6.Text = (rs1.Fields("FMOBILE").Value)
Picture1.Picture = LoadPicture(rs1!FPHOTO)
End Sub

Private Sub Command3_Click()
rs1.MoveNext
If rs1.EOF = True Then
    rs1.MoveFirst
End If

Text1.Text = (rs1.Fields("FNAME").Value)
Text2.Text = (rs1.Fields("FAGE").Value)
class.Text = (rs1.Fields("FCLASS").Value)
stream.Text = (rs1.Fields("FSTREAM").Value)
gender.Text = (rs1.Fields("FGENDER").Value)
Text3.Text = (rs1.Fields("FQUALIFICATION").Value)
subject.Text = (rs1.Fields("FSUBJECT").Value)
Text4.Text = (rs1.Fields("FADDRESS").Value)
Text5.Text = (rs1.Fields("FEMAIL").Value)
Text6.Text = (rs1.Fields("FMOBILE").Value)
Picture1.Picture = LoadPicture(rs1!FPHOTO)
End Sub

Private Sub Command4_Click()
rs1.MoveLast
Text1.Text = (rs1.Fields("FNAME").Value)
Text2.Text = (rs1.Fields("FAGE").Value)
class.Text = (rs1.Fields("FCLASS").Value)
stream.Text = (rs1.Fields("FSTREAM").Value)
gender.Text = (rs1.Fields("FGENDER").Value)
Text3.Text = (rs1.Fields("FQUALIFICATION").Value)
subject.Text = (rs1.Fields("FSUBJECT").Value)
Text4.Text = (rs1.Fields("FADDRESS").Value)
Text5.Text = (rs1.Fields("FEMAIL").Value)
Text6.Text = (rs1.Fields("FMOBILE").Value)
Picture1.Picture = LoadPicture(rs1!FPHOTO)
End Sub

Private Sub Command5_Click()
rs1.MoveFirst
rs1.Find ("FMOBILE='" & CStr(Text7.Text) & "'")
If rs1.EOF = True Then
    Text1.Text = ""
    Text2.Text = ""
    class.Text = "SELECT CLASS"
    stream.Text = "SELECT STREAM"
    gender.Text = "SELECT GENDER"
    Text3.Text = ""
    subject.Text = "SELECT SUBJECT"
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    MsgBox "ENTER A VALID ID", vbCritical
    'Picture1.Picture = ""
    
Else
Text1.Text = (rs1.Fields("FNAME").Value)
Text2.Text = (rs1.Fields("FAGE").Value)
class.Text = (rs1.Fields("FCLASS").Value)
stream.Text = (rs1.Fields("FSTREAM").Value)
gender.Text = (rs1.Fields("FGENDER").Value)
Text3.Text = (rs1.Fields("FQUALIFICATION").Value)
subject.Text = (rs1.Fields("FSUBJECT").Value)
Text4.Text = (rs1.Fields("FADDRESS").Value)
Text5.Text = (rs1.Fields("FEMAIL").Value)
Text6.Text = (rs1.Fields("FMOBILE").Value)
Picture1.Picture = LoadPicture(rs1!FPHOTO)
End If
End Sub

Private Sub Command6_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command7_Click()
sfaculty.Hide
sadminoption.Show
End Sub

Private Sub Command8_Click()
rs1.AddNew
rs1.Fields("FNAME").Value = Text1.Text
rs1.Fields("FAGE").Value = Text2.Text
rs1.Fields("FCLASS").Value = class.Text
rs1.Fields("FSTREAM").Value = stream.Text
rs1.Fields("FGENDER").Value = gender.Text
rs1.Fields("FQUALIFICATION").Value = Text3.Text
rs1.Fields("FSUBJECT").Value = subject.Text
rs1.Fields("FADDRESS").Value = Text4.Text
rs1.Fields("FEMAIL").Value = Text5.Text
rs1.Fields("FMOBILE").Value = Text6.Text
rs1.Fields("FPHOTO").Value = str
MsgBox "data save successfully!!!", vbInformation
rs1.Update

End Sub

Private Sub Command9_Click()
'rs1.Find ("FMOBILE='" & CStr(Text7.Text) & "'")
rs1.Fields("FNAME").Value = Text1.Text
rs1.Fields("FAGE").Value = Text2.Text
rs1.Fields("FCLASS").Value = class.Text
rs1.Fields("FSTREAM").Value = stream.Text
rs1.Fields("FGENDER").Value = gender.Text
rs1.Fields("FQUALIFICATION").Value = Text3.Text
rs1.Fields("FSUBJECT").Value = subject.Text
rs1.Fields("FADDRESS").Value = Text4.Text
rs1.Fields("FEMAIL").Value = Text5.Text
rs1.Fields("FMOBILE").Value = Text6.Text
rs1.Fields("FPHOTO").Value = str
MsgBox "data updated successfully!!!", vbInformation
rs1.Update
End Sub

Private Sub Form_Load()
gender.AddItem ("MALE")
gender.AddItem ("FEMALE")
gender.AddItem ("OTHERS")
class.AddItem ("11")
class.AddItem ("12")
stream.AddItem ("SCIENCE")
stream.AddItem ("ARTS")
stream.AddItem ("COMMERCE")

cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database2.accdb;Persist Security Info=False"
cn.Open
rs1.ActiveConnection = cn
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
rs1.Source = "facultyrecord"
rs1.Open
End Sub

Private Sub stream_Click()
subject.Clear
If stream.Text = "SCIENCE" Then
    subject.AddItem ("PHYSICS")
    subject.AddItem ("CHEMISTRY")
    subject.AddItem ("MATH")
    subject.AddItem ("COMPUTER SCIENCE")
    subject.AddItem ("ENGLISH")
'End If
ElseIf stream.Text = "ARTS" Then
    subject.AddItem ("HISTORY")
    subject.AddItem ("GEOGRAPHY")
    subject.AddItem ("POLITICAL SCIENCE")
    subject.AddItem ("ECONOMICS")
    subject.AddItem ("ENGLISH")
'End If

Else 'stream.Text = "COMMERCE" Then

    subject.AddItem ("MATH")
    subject.AddItem ("BUSINESS STUDIES")
    subject.AddItem ("ACCOUNTS")
    subject.AddItem ("ECONOMICS")
    subject.AddItem ("ENGLISH")
End If
End Sub
