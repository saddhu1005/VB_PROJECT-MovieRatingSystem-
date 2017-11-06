VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   5160
   ClientLeft      =   7470
   ClientTop       =   3435
   ClientWidth     =   8880
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5025
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8745
      Begin VB.CommandButton Command2 
         Caption         =   "REGISTER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4320
         TabIndex        =   6
         Top             =   3450
         Width           =   3585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   570
         TabIndex        =   5
         Top             =   3450
         Width           =   3585
      End
      Begin VB.TextBox Text2 
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
         IMEMode         =   3  'DISABLE
         Left            =   3990
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2070
         Width           =   3945
      End
      Begin VB.TextBox Text1 
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
         Left            =   3990
         TabIndex        =   3
         Top             =   930
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PASSWORD : "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   570
         TabIndex        =   2
         Top             =   2070
         Width           =   3165
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USERNAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   570
         TabIndex        =   1
         Top             =   930
         Width           =   3165
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    UserLogin
End Sub

Private Sub Command2_Click()
    Form4.Show
    Unload Me
End Sub

