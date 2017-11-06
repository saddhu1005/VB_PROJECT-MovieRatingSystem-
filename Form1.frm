VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   3930
   ClientTop       =   -420
   ClientWidth     =   16605
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   16605
   Begin VB.PictureBox Picture3 
      Height          =   2595
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   4605
      TabIndex        =   9
      Top             =   7440
      Width           =   4665
   End
   Begin VB.PictureBox Picture2 
      Height          =   2595
      Left            =   -120
      ScaleHeight     =   2535
      ScaleWidth      =   4605
      TabIndex        =   7
      Top             =   4080
      Width           =   4665
   End
   Begin VB.Frame Frame2 
      Height          =   10740
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   11475
      Begin TabDlg.SSTab SSTab1 
         Height          =   10680
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   18838
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "MOVIE DESCRIPTION"
         TabPicture(0)   =   "Form1.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label5"
         Tab(0).Control(1)=   "Label6"
         Tab(0).Control(2)=   "Label7"
         Tab(0).Control(3)=   "Label8"
         Tab(0).Control(4)=   "Label9"
         Tab(0).Control(5)=   "Label10"
         Tab(0).Control(6)=   "Label11"
         Tab(0).Control(7)=   "Label12"
         Tab(0).Control(8)=   "Label13"
         Tab(0).Control(9)=   "Picture4"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "RATINGS AND REVIEWS"
         TabPicture(1)   =   "Form1.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label14"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Picture5"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame3"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "SEARCH MOVIES"
         TabPicture(2)   =   "Form1.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "searchtxt"
         Tab(2).Control(1)=   "searchbttn"
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame3 
            Caption         =   "RATINGS"
            Height          =   3975
            Left            =   5280
            TabIndex        =   25
            Top             =   1920
            Width           =   5445
            Begin VB.Label Label16 
               Caption         =   "Label16"
               Height          =   3495
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   5175
            End
         End
         Begin VB.PictureBox Picture5 
            Height          =   5145
            Left            =   360
            ScaleHeight     =   5085
            ScaleWidth      =   4575
            TabIndex        =   24
            Top             =   720
            Width           =   4635
         End
         Begin VB.CommandButton searchbttn 
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
            Height          =   615
            Left            =   -67080
            TabIndex        =   22
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox searchtxt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -72000
            TabIndex        =   21
            Top             =   1440
            Width           =   4695
         End
         Begin VB.PictureBox Picture4 
            Height          =   5415
            Left            =   -74760
            ScaleHeight     =   5355
            ScaleWidth      =   4815
            TabIndex        =   11
            Top             =   720
            Width           =   4875
         End
         Begin VB.Frame Frame4 
            Caption         =   "REVIEWS"
            Height          =   3615
            Left            =   360
            TabIndex        =   26
            Top             =   6000
            Width           =   10395
            Begin VB.Label Label15 
               Caption         =   "Label15"
               Height          =   3255
               Left            =   240
               TabIndex        =   28
               Top             =   240
               Width           =   9975
            End
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label14"
            Height          =   645
            Left            =   5400
            TabIndex        =   23
            Top             =   720
            Width           =   5325
         End
         Begin VB.Label Label13 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label13"
            Height          =   2985
            Left            =   -74760
            TabIndex        =   20
            Top             =   6720
            Width           =   10725
         End
         Begin VB.Label Label12 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SYNOPSIS : "
            Height          =   435
            Left            =   -74760
            TabIndex        =   19
            Top             =   6240
            Width           =   3375
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label11"
            Height          =   495
            Left            =   -69720
            TabIndex        =   18
            Top             =   5640
            Width           =   5685
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            Height          =   945
            Left            =   -69720
            TabIndex        =   17
            Top             =   4560
            Width           =   5685
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label9"
            Height          =   555
            Left            =   -69720
            TabIndex        =   16
            Top             =   3870
            Width           =   5685
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label8"
            Height          =   555
            Left            =   -69720
            TabIndex        =   15
            Top             =   3210
            Width           =   5685
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label7"
            Height          =   555
            Left            =   -69720
            TabIndex        =   14
            Top             =   2550
            Width           =   5685
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label6"
            Height          =   555
            Left            =   -69720
            TabIndex        =   13
            Top             =   1890
            Width           =   5685
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label5"
            Height          =   1035
            Left            =   -69720
            TabIndex        =   12
            Top             =   750
            Width           =   5685
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   11925
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   5265
      Begin VB.PictureBox Picture1 
         Height          =   2595
         Left            =   90
         ScaleHeight     =   2535
         ScaleWidth      =   4605
         TabIndex        =   5
         Top             =   840
         Width           =   4665
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   11805
         Left            =   4830
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   615
         Left            =   90
         TabIndex        =   10
         Top             =   11070
         Width           =   4665
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   495
         Left            =   90
         TabIndex        =   8
         Top             =   6840
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         Height          =   495
         Left            =   90
         TabIndex        =   6
         Top             =   3480
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOVIES : NOW RUNNING"
         Height          =   435
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   4665
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu AddNewMovie 
         Caption         =   "Add New Movie"
      End
      Begin VB.Menu DeleteMovie 
         Caption         =   "Delete Movie"
      End
      Begin VB.Menu RateaMovie 
         Caption         =   "Rate a Movie"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddNewMovie_Click()
Unload Form2
Load Form2
Form2.Show
Form1.Hide
Form2.SSTab1.TabVisible(0) = True
Form2.SSTab1.TabVisible(1) = False
Form2.SSTab1.TabVisible(2) = False
Form2.SSTab1 = 0
    
End Sub



Private Sub DeleteMovie_Click()
Form1.Hide
Unload Form2
Load Form2
Form2.Show
Form2.SSTab1.TabVisible(0) = False
Form2.SSTab1.TabVisible(1) = False
Form2.SSTab1.TabVisible(2) = True
Form2.SSTab1 = 2
End Sub

Private Sub Edit_Click()
Form1.Hide
Unload Form2
Load Form2
Form2.Show
Form2.SSTab1.TabVisible(0) = False
Form2.SSTab1.TabVisible(2) = False
Form2.SSTab1.TabVisible(1) = True

Form2.SSTab1 = 1


End Sub

Private Sub Exit_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
End Sub

Private Sub Form_Load()
    LoadDatabase
    CurrentMoviesArray
    LoadIntialPictures
     Initializations
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
searchtxt.SetFocus
End Sub

Private Sub Help_Click()
MsgBox "The Motion Picture Association of America (MPAA) film rating system is used in the United States and its territories to rate a film's suitability for certain audiences based on its content. ... It is administered by the Classification & Ratings Administration (CARA), an independent division of the MPAA.", vbOKOnly, "About"
End Sub

Private Sub Label2_Click()
    FirstItemClick
End Sub

Private Sub Label3_Click()
    SecondItemClick
End Sub

Private Sub Label4_Click()
    ThirdItemClick
End Sub

Private Sub Picture1_Click()
    FirstItemClick
End Sub

Private Sub Picture2_Click()
    SecondItemClick
End Sub

Private Sub Picture3_Click()
    ThirdItemClick
End Sub

Private Sub RateaMovie_Click()
Form1.Hide
Load Form5
Form5.Show
End Sub

Private Sub searchbttn_Click()
If searchtxt.Text = "" Then
MsgBox "Enter Movie Name First!"
Exit Sub
Else
searchitem
End If
End Sub

Private Sub VScroll1_Change()
    ScrollChange
End Sub
