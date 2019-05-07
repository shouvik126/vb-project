VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adarts 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   15570
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   13920
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
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
   Begin VB.CommandButton Command2 
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
      Left            =   13680
      TabIndex        =   93
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   13680
      TabIndex        =   92
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "11(GEOGRAPHY)"
      Height          =   615
      Left            =   2040
      TabIndex        =   79
      Top             =   1800
      Width           =   2055
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "12((HISTORY)"
      Height          =   615
      Left            =   2040
      TabIndex        =   77
      Top             =   2520
      Width           =   2055
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   78
         Text            =   "Combo2"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000000FF&
      Caption         =   "11(HISTORY)"
      Height          =   615
      Left            =   2040
      TabIndex        =   75
      Top             =   3240
      Width           =   2055
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   120
         TabIndex        =   76
         Text            =   "Combo3"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(GEOGRAPHY)"
      Height          =   615
      Left            =   2040
      TabIndex        =   73
      Top             =   3960
      Width           =   2055
      Begin VB.ComboBox Combo14 
         Height          =   315
         Left            =   120
         TabIndex        =   74
         Text            =   "Combo14"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H000000FF&
      Caption         =   "11(ENGLISH)"
      Height          =   615
      Left            =   2040
      TabIndex        =   71
      Top             =   4680
      Width           =   2055
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   120
         TabIndex        =   72
         Text            =   "Combo4"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ECONOMICS)"
      Height          =   615
      Left            =   2040
      TabIndex        =   69
      Top             =   5400
      Width           =   2055
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   120
         TabIndex        =   70
         Text            =   "Combo5"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H000000FF&
      Caption         =   "11(POL SC)"
      Height          =   615
      Left            =   2040
      TabIndex        =   67
      Top             =   6120
      Width           =   2055
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   240
         TabIndex        =   68
         Text            =   "Combo6"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(GEOGRAPHY)"
      Height          =   615
      Left            =   2040
      TabIndex        =   65
      Top             =   6840
      Width           =   2055
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   120
         TabIndex        =   66
         Text            =   "Combo7"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H000000FF&
      Caption         =   "11(HISTORY)"
      Height          =   615
      Left            =   2040
      TabIndex        =   63
      Top             =   7560
      Width           =   2055
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Text            =   "Combo8"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ECONOMICS)"
      Height          =   615
      Left            =   2040
      TabIndex        =   61
      Top             =   8280
      Width           =   2055
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Text            =   "Combo9"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H000000FF&
      Caption         =   "11(HISTORY)"
      Height          =   615
      Left            =   4440
      TabIndex        =   59
      Top             =   1800
      Width           =   2055
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   120
         TabIndex        =   60
         Text            =   "Combo10"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(POL SC)"
      Height          =   615
      Left            =   4440
      TabIndex        =   57
      Top             =   2520
      Width           =   2055
      Begin VB.ComboBox Combo11 
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Text            =   "Combo11"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H000000FF&
      Caption         =   "11(POL SC)"
      Height          =   615
      Left            =   4440
      TabIndex        =   55
      Top             =   3240
      Width           =   2055
      Begin VB.ComboBox Combo12 
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Text            =   "Combo12"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(POL SC)"
      Height          =   615
      Left            =   4440
      TabIndex        =   53
      Top             =   3960
      Width           =   2055
      Begin VB.ComboBox Combo13 
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Text            =   "Combo13"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H000000FF&
      Caption         =   "11(POL SCIENCE)"
      Height          =   615
      Left            =   4440
      TabIndex        =   51
      Top             =   4680
      Width           =   2055
      Begin VB.ComboBox Combo15 
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Text            =   "Combo15"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(GEOGRAPHY)"
      Height          =   615
      Left            =   4440
      TabIndex        =   49
      Top             =   5400
      Width           =   2055
      Begin VB.ComboBox Combo16 
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Text            =   "Combo16"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame17 
      BackColor       =   &H000000FF&
      Caption         =   "11(GEOGRAPHY)"
      Height          =   615
      Left            =   4440
      TabIndex        =   47
      Top             =   6120
      Width           =   2055
      Begin VB.ComboBox Combo17 
         Height          =   315
         Left            =   240
         TabIndex        =   48
         Text            =   "Combo17"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(POL SC)"
      Height          =   615
      Left            =   4440
      TabIndex        =   45
      Top             =   6840
      Width           =   2055
      Begin VB.ComboBox Combo18 
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Text            =   "Combo18"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame19 
      BackColor       =   &H000000FF&
      Caption         =   "11(ECONOMICS)"
      Height          =   615
      Left            =   4440
      TabIndex        =   43
      Top             =   7560
      Width           =   2055
      Begin VB.ComboBox Combo19 
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Text            =   "Combo19"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(HISTORY)"
      Height          =   615
      Left            =   4440
      TabIndex        =   41
      Top             =   8280
      Width           =   2055
      Begin VB.ComboBox Combo20 
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Text            =   "Combo20"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame21 
      BackColor       =   &H000000FF&
      Caption         =   "11(POL SC)"
      Height          =   615
      Left            =   9120
      TabIndex        =   39
      Top             =   1800
      Width           =   1935
      Begin VB.ComboBox Combo21 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame22 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ECONOMICS)"
      Height          =   615
      Left            =   9120
      TabIndex        =   37
      Top             =   2520
      Width           =   1935
      Begin VB.ComboBox Combo22 
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Text            =   "Combo22"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame23 
      BackColor       =   &H000000FF&
      Caption         =   "11(ENGLISH)"
      Height          =   615
      Left            =   9120
      TabIndex        =   35
      Top             =   3240
      Width           =   1935
      Begin VB.ComboBox Combo23 
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Text            =   "Combo23"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame24 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ENGLISH)"
      Height          =   615
      Left            =   9120
      TabIndex        =   33
      Top             =   3960
      Width           =   1935
      Begin VB.ComboBox Combo24 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Text            =   "Combo24"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame25 
      BackColor       =   &H000000FF&
      Caption         =   "11(GEOGRAPHY)"
      Height          =   615
      Left            =   9120
      TabIndex        =   31
      Top             =   4680
      Width           =   1935
      Begin VB.ComboBox Combo25 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Text            =   "Combo25"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame26 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(HISTORY)"
      Height          =   615
      Left            =   9120
      TabIndex        =   29
      Top             =   5400
      Width           =   1935
      Begin VB.ComboBox Combo26 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Text            =   "Combo26"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame27 
      BackColor       =   &H000000FF&
      Caption         =   "11(ECONOMICS)"
      Height          =   615
      Left            =   9120
      TabIndex        =   27
      Top             =   6120
      Width           =   1935
      Begin VB.ComboBox Combo27 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Text            =   "Combo27"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame28 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ENGLISH)"
      Height          =   615
      Left            =   9120
      TabIndex        =   25
      Top             =   6840
      Width           =   1935
      Begin VB.ComboBox Combo28 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Text            =   "Combo28"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame29 
      BackColor       =   &H000000FF&
      Caption         =   "11(ENGLISH)"
      Height          =   615
      Left            =   9120
      TabIndex        =   23
      Top             =   7560
      Width           =   1935
      Begin VB.ComboBox Combo29 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Text            =   "Combo29"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame30 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(GEOGRAPHY)"
      Height          =   615
      Left            =   9120
      TabIndex        =   21
      Top             =   8280
      Width           =   1935
      Begin VB.ComboBox Combo30 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Text            =   "Combo30"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame31 
      BackColor       =   &H000000FF&
      Caption         =   "11(ECONOMICS)"
      Height          =   615
      Left            =   11400
      TabIndex        =   19
      Top             =   1800
      Width           =   2055
      Begin VB.ComboBox Combo31 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Text            =   "Combo31"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame32 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ENGLISH)"
      Height          =   615
      Left            =   11400
      TabIndex        =   17
      Top             =   2520
      Width           =   2055
      Begin VB.ComboBox Combo32 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Text            =   "Combo32"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame33 
      BackColor       =   &H000000FF&
      Caption         =   "11(ECONOMICS)"
      Height          =   615
      Left            =   11400
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
      Begin VB.ComboBox Combo33 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Text            =   "Combo33"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame34 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(HISTORY)"
      Height          =   615
      Left            =   11400
      TabIndex        =   13
      Top             =   3960
      Width           =   2055
      Begin VB.ComboBox Combo34 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Text            =   "Combo34"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame35 
      BackColor       =   &H000000FF&
      Caption         =   "11(HISTORY)"
      Height          =   615
      Left            =   11400
      TabIndex        =   11
      Top             =   4680
      Width           =   2055
      Begin VB.ComboBox Combo35 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Text            =   "Combo35"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame36 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(ENGLISH)"
      Height          =   615
      Left            =   11400
      TabIndex        =   9
      Top             =   5400
      Width           =   2055
      Begin VB.ComboBox Combo36 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "Combo36"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame37 
      BackColor       =   &H000000FF&
      Caption         =   "11(ENGLISH)"
      Height          =   615
      Left            =   11400
      TabIndex        =   7
      Top             =   6120
      Width           =   2055
      Begin VB.ComboBox Combo37 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "Combo37"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame38 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(HISTORY)"
      Height          =   615
      Left            =   11400
      TabIndex        =   5
      Top             =   6840
      Width           =   2055
      Begin VB.ComboBox Combo38 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Text            =   "Combo38"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame39 
      BackColor       =   &H000000FF&
      Caption         =   "11(GEOGRAPHY)"
      Height          =   615
      Left            =   11400
      TabIndex        =   3
      Top             =   7560
      Width           =   2055
      Begin VB.ComboBox Combo39 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "Combo39"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame40 
      BackColor       =   &H0080FFFF&
      Caption         =   "12(POL SC)"
      Height          =   615
      Left            =   11400
      TabIndex        =   1
      Top             =   8280
      Width           =   2055
      Begin VB.ComboBox Combo40 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "Combo40"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "MON"
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
      Left            =   0
      TabIndex        =   91
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "TUES"
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
      Left            =   0
      TabIndex        =   90
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "WED"
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
      Left            =   0
      TabIndex        =   89
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "THURS"
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
      Left            =   0
      TabIndex        =   88
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "FRI"
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
      Left            =   0
      TabIndex        =   87
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "(11-12)PM"
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
      Left            =   2040
      TabIndex        =   86
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "(12-1)PM"
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
      Left            =   4440
      TabIndex        =   85
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "(1-1:30)PM"
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
      Left            =   6840
      TabIndex        =   84
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "(1:30-2:30)"
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
      Left            =   9120
      TabIndex        =   83
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "(2:30-3:30)"
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
      Left            =   11400
      TabIndex        =   82
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "BREAK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   6840
      TabIndex        =   81
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "ARTS TIME TABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   4440
   End
End
Attribute VB_Name = "adarts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim i As Integer
Dim n As Integer


Private Sub Command1_Click()
sadminoption.Show
adarts.Hide
End Sub

Private Sub Command2_Click()
rs2.MoveFirst
Do While rs2.EOF = False
rs2.Delete
rs2.MoveNext
Loop
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo1.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo2.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo3.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo14.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo4.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo5.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo6.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo7.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo8.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo9.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo10.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo11.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo12.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo13.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo15.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo16.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo17.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo18.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo19.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo20.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo21.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo22.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo23.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo24.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo25.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo26.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo27.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo28.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo29.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo30.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo31.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo32.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo33.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo34.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo35.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo36.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo37.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo38.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo39.Text
    rs2.AddNew
    rs2.Fields("TEACHER").Value = Combo40.Text
rs2.Update
MsgBox "YOUR DATA HAS BEEN SAVED SUCCESSFULLY!!!", vbInformation
End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\vb project\database\Database2.accdb;Persist Security Info=False"
cn.Open
rs1.ActiveConnection = cn
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
rs1.Source = "facultyrecord"
rs1.Open
rs2.ActiveConnection = cn
rs2.CursorType = adOpenDynamic
rs2.LockType = adLockOptimistic
rs2.Source = "artstable"
rs2.Open

Combo1.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo2.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo3.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo14.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo4.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo5.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo6.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo7.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo8.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo9.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo10.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo11.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo12.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo13.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo15.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo16.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo17.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo18.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo19.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo20.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo21.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo22.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo23.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo24.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo25.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo26.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo27.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo28.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo29.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo30.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo31.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo32.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo33.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo34.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo35.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo36.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo37.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo38.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo39.Text = rs2.Fields("TEACHER").Value
    rs2.MoveNext
    Combo40.Text = rs2.Fields("TEACHER").Value

rs2.MoveFirst


Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "11" And rs1.Fields("FSUBJECT") = "HISTORY") Then
    Combo10.AddItem (rs1.Fields("FNAME"))
    Combo3.AddItem (rs1.Fields("FNAME"))
    Combo35.AddItem (rs1.Fields("FNAME"))
    Combo8.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
Loop
rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "11" And rs1.Fields("FSUBJECT") = "GEOGRAPHY") Then
    Combo1.AddItem (rs1.Fields("FNAME"))
    Combo25.AddItem (rs1.Fields("FNAME"))
    Combo17.AddItem (rs1.Fields("FNAME"))
    Combo39.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
    Loop
    rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "11" And rs1.Fields("FSUBJECT") = "POLITICAL SCIENCE") Then
    Combo21.AddItem (rs1.Fields("FNAME"))
    Combo12.AddItem (rs1.Fields("FNAME"))
    Combo15.AddItem (rs1.Fields("FNAME"))
    Combo6.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
    Loop
    rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "11" And rs1.Fields("FSUBJECT") = "ECONOMICS") Then
    Combo31.AddItem (rs1.Fields("FNAME"))
    Combo33.AddItem (rs1.Fields("FNAME"))
    Combo27.AddItem (rs1.Fields("FNAME"))
    Combo19.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
    Loop
    rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "11" And rs1.Fields("FSUBJECT") = "ENGLISH") Then
    Combo23.AddItem (rs1.Fields("FNAME"))
    Combo4.AddItem (rs1.Fields("FNAME"))
    Combo37.AddItem (rs1.Fields("FNAME"))
    Combo29.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
Loop

rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "12" And rs1.Fields("FSUBJECT") = "HISTORY") Then
    Combo2.AddItem (rs1.Fields("FNAME"))
    Combo34.AddItem (rs1.Fields("FNAME"))
    Combo26.AddItem (rs1.Fields("FNAME"))
    Combo38.AddItem (rs1.Fields("FNAME"))
    Combo20.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
Loop
rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "12" And rs1.Fields("FSUBJECT") = "GEOGRAPHY") Then
    Combo14.AddItem (rs1.Fields("FNAME"))
    Combo16.AddItem (rs1.Fields("FNAME"))
    Combo7.AddItem (rs1.Fields("FNAME"))
    Combo30.AddItem (rs1.Fields("FNAME"))
    End If
rs1.MoveNext
    Loop
    rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "12" And rs1.Fields("FSUBJECT") = "POLITICAL SCIENCE") Then
    Combo11.AddItem (rs1.Fields("FNAME"))
    Combo13.AddItem (rs1.Fields("FNAME"))
    Combo18.AddItem (rs1.Fields("FNAME"))
    Combo40.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
    Loop
    rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "12" And rs1.Fields("FSUBJECT") = "ECONOMICS") Then
    Combo22.AddItem (rs1.Fields("FNAME"))
    Combo5.AddItem (rs1.Fields("FNAME"))
    Combo9.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
    Loop
    rs1.MoveFirst
Do While rs1.EOF = False
    If (rs1.Fields("FCLASS").Value = "12" And rs1.Fields("FSUBJECT") = "ENGLISH") Then
    Combo32.AddItem (rs1.Fields("FNAME"))
    Combo24.AddItem (rs1.Fields("FNAME"))
    Combo36.AddItem (rs1.Fields("FNAME"))
    Combo28.AddItem (rs1.Fields("FNAME"))
    End If
    rs1.MoveNext
    Loop
    
    rs1.MoveFirst
    
End Sub

