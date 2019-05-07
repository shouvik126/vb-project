VERSION 5.00
Begin VB.Form sadminoption 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox mark 
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
      Left            =   8520
      TabIndex        =   6
      Text            =   "SELECT (CLASS & STREAM)"
      Top             =   5160
      Width           =   4215
   End
   Begin VB.ComboBox stream 
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
      Left            =   4440
      TabIndex        =   4
      Text            =   "SELECT STREAM"
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
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
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Faculty Management"
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
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Students Marks Management"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   4560
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Time Table Management"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   3330
      Left            =   240
      Picture         =   "sadminoption.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   14700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "ADMINISTRATOR OPTIONS"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   6750
   End
End
Attribute VB_Name = "sadminoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    sfaculty.Show
    sadminoption.Hide
End Sub

Private Sub Command2_Click()
sadminoption.Hide
smainpage.Show

End Sub

Private Sub Form_Load()
mark.AddItem ("11-SCIENCE")
mark.AddItem ("11-ARTS")
mark.AddItem ("11-COMMERCE")
mark.AddItem ("12-SCIENCE")
mark.AddItem ("12-ARTS")
mark.AddItem ("12-COMMERCE")
stream.AddItem ("SCIENCE")
stream.AddItem ("ARTS")
stream.AddItem ("COMMERCE")
End Sub


Private Sub mark_Click()
If mark.Text = "11-SCIENCE" Then
sadminoption.Hide
s11sciencemarks.Show
End If
If mark.Text = "11-ARTS" Then
sadminoption.Hide
s11artsmarks.Show
End If
If mark.Text = "11-COMMERCE" Then
sadminoption.Hide
s11commercemarks.Show
End If

If mark.Text = "12-SCIENCE" Then
sadminoption.Hide
s12sciencemarks.Show
End If
If mark.Text = "12-ARTS" Then
sadminoption.Hide
s12artsmarks.Show
End If
If mark.Text = "12-COMMERCE" Then
sadminoption.Hide
s12commercemarks.Show
End If
End Sub

Private Sub stream_Click()
If stream.Text = "SCIENCE" Then
    sadminoption.Hide
    adscience.Show
End If
If stream.Text = "ARTS" Then
    sadminoption.Hide
    adarts.Show
    End If
If stream.Text = "COMMERCE" Then
    sadminoption.Hide
    adcommerce.Show
    End If


End Sub
