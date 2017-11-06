VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090
   LinkTopic       =   "Form4"
   ScaleHeight     =   7830
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " REGISTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7635
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8895
      Begin VB.CommandButton Command2 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   4410
         TabIndex        =   8
         Top             =   5760
         Width           =   2955
      End
      Begin VB.CommandButton Command1 
         Caption         =   "REGISTER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1110
         TabIndex        =   7
         Top             =   5790
         Width           =   2955
      End
      Begin VB.TextBox Text5 
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
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   4740
         Width           =   3465
      End
      Begin VB.TextBox Text4 
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
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   3810
         Width           =   3465
      End
      Begin VB.TextBox Text3 
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
         Left            =   4080
         TabIndex        =   4
         Top             =   2970
         Width           =   3465
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
         Left            =   4080
         TabIndex        =   3
         Top             =   2010
         Width           =   3465
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
         Left            =   4080
         TabIndex        =   2
         Top             =   1080
         Width           =   3465
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONFIRM PASSWORD : "
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
         Left            =   1020
         TabIndex        =   12
         Top             =   4740
         Width           =   3075
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PASSWORD : "
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
         Left            =   1020
         TabIndex        =   11
         Top             =   3810
         Width           =   3075
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USERNAME : "
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
         Left            =   1020
         TabIndex        =   10
         Top             =   2970
         Width           =   3075
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LAST NAME : "
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
         Left            =   1020
         TabIndex        =   9
         Top             =   2010
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FIRST NAME : "
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
         Left            =   1020
         TabIndex        =   1
         Top             =   1080
         Width           =   3075
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    RegisterUser
End Sub

Private Sub Command2_Click()
    Unload Me
    Form3.Show
End Sub
