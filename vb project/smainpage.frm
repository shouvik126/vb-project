VERSION 5.00
Begin VB.Form smainpage 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   15090
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   5
      Top             =   7440
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ADMISSION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   3495
      Left            =   3720
      Picture         =   "smainpage.frx":0000
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   9495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   4680
      Picture         =   "smainpage.frx":39922
      Top             =   1320
      Width           =   6450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "STUDY CENTRE  MANAGEMENT system"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi Cond"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   10455
   End
End
Attribute VB_Name = "smainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
smainpage.Hide
sstudententry.Show
End Sub

Private Sub Command2_Click()
smainpage.Hide
sadminlogin.Show
End Sub

Private Sub Command3_Click()
smainpage.Hide
sinfo.Show
End Sub

Private Sub Command4_Click()
End

End Sub


