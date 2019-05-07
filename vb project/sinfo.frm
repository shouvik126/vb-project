VERSION 5.00
Begin VB.Form sinfo 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   2265
   ClientTop       =   4530
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   15900
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
      Height          =   975
      Left            =   6480
      TabIndex        =   8
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      Caption         =   "GUIDED BY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   15375
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Soumen pal"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Spm maam"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "DESIGNED BY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   15495
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Sayan"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12000
         TabIndex        =   5
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Nilanjan"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   4
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Sawan"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Sunetra "
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1875
      End
   End
End
Attribute VB_Name = "sinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sinfo.Hide
smainpage.Show
End Sub

