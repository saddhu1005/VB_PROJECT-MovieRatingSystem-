Attribute VB_Name = "RatingsCalculation"

Public Sub CalculateRating()

'gross rating calculation goes here

End Sub
Public Sub AddRating()
Dim name, reveiw As String
Dim rating As Integer
name = UCase(Form5.Text1.Text)
review = Trim(Form5.Text2.Text)

If Form5.Option1.Value = True Then
rating = 1
ElseIf Form5.Option2.Value = True Then
rating = 2
ElseIf Form5.Option3.Value = True Then
rating = 3
ElseIf Form5.Option4.Value = True Then
rating = 4
ElseIf Form5.Option5.Value = True Then
rating = 5
ElseIf Form5.Option6.Value = True Then
rating = 6
ElseIf Form5.Option7.Value = True Then
rating = 7
ElseIf Form5.Option8.Value = True Then
rating = 8
ElseIf Form5.Option9.Value = True Then
rating = 9
ElseIf Form5.Option2.Value = True Then
rating = 10
Else
MsgBox "Choose Rating First "
Exit Sub
End If
If review = "" Then
MsgBox "Please Write A Review"
Exit Sub
End If


'storing codes goes here
'write codes to store rating and review

CalculateRating    'for calculating and storing gross rating


MsgBox "Succcesfully Rated"

Form5.Text1.Locked = False
Form5.Text1.Text = ""
Form5.Text2.Text = ""
Form5.Command2.Visible = False
Form5.Command3.Visible = True
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
Public Sub updatesearch()
searchupd = UCase(Form5.Text1.Text)
Dim X, loc As Integer

X = 0
loc = 0
 Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim db_path As String
    
    db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from MovieDetails", cn, adOpenStatic, adLockOptimistic
        
    With rs
.MoveFirst
 Do While (Not .EOF) And (Not X = 1)
 If searchupd = .Fields(0) Then
 X = 1
 editshow
 
 End If
 .MoveNext
 Loop
 .Close
 End With
 If X <> 1 Then
 MsgBox "Movie Record Not Found ,Enter correctly"
  Form5.Text1.Text = ""
 searchupd = ""

 Exit Sub
 End If
 searchupd = ""
 
End Sub
Public Sub editshow()
 MsgBox "  Enter  Rating And Review .", vbOKOnly, "Movie Found"
Form5.Command3.Visible = False
 Form5.Command2.Visible = True
 Form5.Text1.Locked = True
 
End Sub

