VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9855
   ClientLeft      =   6210
   ClientTop       =   1605
   ClientWidth     =   13155
   LinkTopic       =   "Form2"
   ScaleHeight     =   9855
   ScaleWidth      =   13155
   Begin VB.Frame Frame1 
      Height          =   10755
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   13275
      Begin TabDlg.SSTab SSTab1 
         Height          =   9555
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   16854
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ADD MOVIE"
         TabPicture(0)   =   "Form2.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label5"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label6"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label7"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label8"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label9"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Text1"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Text2"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Text4"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Text5"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Text6"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Text7"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Text8"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Command1"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Option1"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Option2"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Command2"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Command3"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "CommonDialog1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "DTPicker1"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).ControlCount=   23
         TabCaption(1)   =   "UPDATE MOVIE "
         TabPicture(1)   =   "Form2.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "DTPicker2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Command6"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Command7"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Text9"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Command8"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Text10"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Command9"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "CommonDialog2"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Check1"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Check2"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Check3"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Option3"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Option4"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Check4"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Text11"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "DELETE MOVIE "
         TabPicture(2)   =   "Form2.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Command4"
         Tab(2).Control(1)=   "Command5"
         Tab(2).Control(2)=   "Text3"
         Tab(2).Control(3)=   "Label10"
         Tab(2).ControlCount=   4
         Begin VB.TextBox Text11 
            Height          =   3075
            Left            =   -74520
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   4920
            Width           =   11835
         End
         Begin VB.CheckBox Check4 
            Caption         =   "SYNOPSIS :"
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
            Left            =   -74520
            TabIndex        =   41
            Top             =   4200
            Width           =   2295
         End
         Begin VB.OptionButton Option4 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -69960
            TabIndex        =   40
            Top             =   3480
            Width           =   1305
         End
         Begin VB.OptionButton Option3 
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -70800
            TabIndex        =   39
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "CURRENTLY IN THEATRES : "
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
            Left            =   -74520
            TabIndex        =   38
            Top             =   3360
            Width           =   3735
         End
         Begin VB.CheckBox Check2 
            Caption         =   "RELEASE DATE :"
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
            Left            =   -74520
            TabIndex        =   37
            Top             =   2400
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "IMAGE :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74520
            TabIndex        =   35
            Top             =   1560
            Width           =   2535
         End
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   -66960
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -68280
            TabIndex        =   34
            Top             =   1560
            Width           =   1755
         End
         Begin VB.TextBox Text10 
            Height          =   585
            Left            =   -71640
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1560
            Width           =   3105
         End
         Begin VB.CommandButton Command8 
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
            Left            =   -65760
            TabIndex        =   32
            Top             =   720
            Width           =   3465
         End
         Begin VB.TextBox Text9 
            Height          =   585
            Left            =   -71640
            TabIndex        =   31
            Top             =   840
            Width           =   5115
         End
         Begin VB.CommandButton Command7 
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
            Left            =   -65760
            TabIndex        =   29
            Top             =   1800
            Width           =   3465
         End
         Begin VB.CommandButton Command6 
            Caption         =   "UPDATE"
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
            Left            =   -65760
            TabIndex        =   28
            Top             =   720
            Width           =   3465
         End
         Begin VB.CommandButton Command4 
            Caption         =   "DELETE"
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
            Left            =   -65760
            TabIndex        =   27
            Top             =   720
            Width           =   3465
         End
         Begin VB.CommandButton Command5 
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
            Left            =   -65760
            TabIndex        =   26
            Top             =   1800
            Width           =   3465
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -71880
            TabIndex        =   25
            Top             =   1080
            Width           =   5355
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   585
            Left            =   3990
            TabIndex        =   23
            Top             =   2250
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   1032
            _Version        =   393216
            Format          =   209518593
            CurrentDate     =   43045
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   8400
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command3 
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
            Left            =   9240
            TabIndex        =   22
            Top             =   1800
            Width           =   3465
         End
         Begin VB.CommandButton Command2 
            Caption         =   "ADD  MOVIE"
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
            Left            =   9240
            TabIndex        =   21
            Top             =   720
            Width           =   3465
         End
         Begin VB.OptionButton Option2 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6570
            TabIndex        =   20
            Top             =   6300
            Width           =   1305
         End
         Begin VB.OptionButton Option1 
            Caption         =   "YES"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4710
            TabIndex        =   19
            Top             =   6300
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7110
            TabIndex        =   18
            Top             =   1500
            Width           =   1755
         End
         Begin VB.TextBox Text8 
            Height          =   2355
            Left            =   510
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   7080
            Width           =   11835
         End
         Begin VB.TextBox Text7 
            Height          =   585
            Left            =   3990
            TabIndex        =   16
            Top             =   5430
            Width           =   4845
         End
         Begin VB.TextBox Text6 
            Height          =   585
            Left            =   3990
            TabIndex        =   15
            Top             =   4650
            Width           =   4845
         End
         Begin VB.TextBox Text5 
            Height          =   585
            Left            =   3990
            TabIndex        =   14
            Top             =   3840
            Width           =   4845
         End
         Begin VB.TextBox Text4 
            Height          =   585
            Left            =   3990
            TabIndex        =   13
            Top             =   3060
            Width           =   4845
         End
         Begin VB.TextBox Text2 
            Height          =   585
            Left            =   3990
            TabIndex        =   12
            Top             =   1500
            Width           =   3105
         End
         Begin VB.TextBox Text1 
            Height          =   585
            Left            =   3990
            TabIndex        =   11
            Top             =   750
            Width           =   4875
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   585
            Left            =   -71640
            TabIndex        =   36
            Top             =   2520
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   1032
            _Version        =   393216
            Format          =   209518593
            CurrentDate     =   43045
         End
         Begin VB.Label Label11 
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
            Left            =   -74880
            TabIndex        =   30
            Top             =   960
            Width           =   3435
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "NAME OF MOVIE :"
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
            Left            =   -74640
            TabIndex        =   24
            Top             =   1200
            Width           =   2895
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "IMAGE : "
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
            Left            =   510
            TabIndex        =   10
            Top             =   1530
            Width           =   3435
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "CURRENTLY IN THEATRES : "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   510
            TabIndex        =   9
            Top             =   6240
            Width           =   3435
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "SYNOPSIS : "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   8
            Top             =   6720
            Width           =   3375
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "CAST : "
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
            Left            =   510
            TabIndex        =   7
            Top             =   5460
            Width           =   3435
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "DIRECTOR : "
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
            Left            =   510
            TabIndex        =   6
            Top             =   4680
            Width           =   3435
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "GENRE : "
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
            Left            =   510
            TabIndex        =   5
            Top             =   3870
            Width           =   3435
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "LANGUAGE : "
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
            Left            =   510
            TabIndex        =   4
            Top             =   3060
            Width           =   3435
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "RELEASE DATE : "
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
            Left            =   510
            TabIndex        =   3
            Top             =   2280
            Width           =   3435
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
            Left            =   510
            TabIndex        =   2
            Top             =   780
            Width           =   3435
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public searchupd As String

Private Sub Command1_Click()
    CommonDialog1.Filter = "Apps (*jpg|*.jpg*|All files (*.*)|*.*|*jpeg|*.jpeg*|*png|*.png*|*gif|*.gif*"
    CommonDialog1.DefaultExt = "jpeg"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen
    Text2.Text = CommonDialog1.filename
End Sub

Private Sub Command2_Click()
    AddRecord
End Sub

Private Sub Command3_Click()
    Form2.Hide
    Form1.Show
   
End Sub

Private Sub Command4_Click()
DeleteRecord
End Sub

Private Sub Command5_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Command6_Click()
UpdateRecord
End Sub

Private Sub Command7_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Command8_Click()

searchupdate
End Sub

Private Sub Command9_Click()
CommonDialog1.Filter = "Apps (*jpg|*.jpg*|All files (*.*)"
    CommonDialog2.DefaultExt = "jpeg"
    CommonDialog2.DialogTitle = "Select File"
    CommonDialog2.ShowOpen
    Text10.Locked = False
    Text10.Text = CommonDialog2.filename
    Text10.Locked = True
    
End Sub

Private Sub Form_Load()
Form2.SSTab1.TabVisible(0) = True
Form2.SSTab1.TabVisible(1) = True
Form2.SSTab1.TabVisible(2) = True
Form2.SSTab1 = 0

End Sub

Private Sub Label12_Click()

End Sub

