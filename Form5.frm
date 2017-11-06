VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13845
   LinkTopic       =   "Form5"
   ScaleHeight     =   8040
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "RATE MOVIE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.TextBox Text2 
         Height          =   3075
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   4200
         Width           =   11715
      End
      Begin VB.OptionButton Option10 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   16
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton Option9 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   14
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option7 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8880
         TabIndex        =   5
         Top             =   1080
         Width           =   3585
      End
      Begin VB.TextBox Text1 
         Height          =   585
         Left            =   2760
         TabIndex        =   4
         Top             =   960
         Width           =   5475
      End
      Begin VB.CommandButton Command2 
         Caption         =   "RATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8880
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "YOUR REVIEW : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   17
         Top             =   3480
         Width           =   2595
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "YOUR RATING : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   2715
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "NAME OF MOVIE : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3075
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Unload Form5
Load Form5
Form5.Hide
 Unload Form1
Load Form1
  LoadDatabase
    CurrentMoviesArray
    LoadIntialPictures
     Initializations
     
     
 Form1.Show
 Form1.SSTab1 = 2
  

End Sub

Private Sub Command2_Click()
AddRating
End Sub

Private Sub Command3_Click()

updatesearch
End Sub

Private Sub Form_Load()
Form5.Text1.Locked = False
Form5.Text1.Text = ""
Form5.Text2.Text = ""
Form5.Command2.Visible = False
Form5.Command3.Visible = True
End Sub

