Attribute VB_Name = "RecordDeleteUpdate"
 Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim deletejob As String
 Dim db_path As String
 Dim X As Integer

Public Sub DeleteRecord()
deletejob = UCase(Form2.Text3.Text)
If deletejob = "" Then
MsgBox "Enter Movie Name First!"
Exit Sub
End If
X = 0
 db_path = App.Path + "\db\MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from MovieDetails", cn, adOpenStatic, adLockOptimistic
        
                
    With rs
.MoveFirst
 Do While (Not .EOF) And (Not X = 1)
 If deletejob = .Fields(0) Then
 X = 1
 .Delete
 MsgBox "Movie Record Deleted Successfully"
  
   
 Unload Form2
   

  Unload Form1
Load Form1
  Load Form2
  LoadDatabase
    CurrentMoviesArray
    LoadIntialPictures
     Initializations
     
       Form2.Hide
 Form1.Show
 Form1.SSTab1 = 2
 End If
 .MoveNext
 Loop
 .Close
 End With
 If X <> 1 Then
MsgBox "No Record Found, Please type correctly !", vbOKOnly, "! Sorry !"
Form2.Text3.Text = ""

Exit Sub

End If
Form2.Text3.Text = ""
deletejob = ""


End Sub

Public Sub searchupdate()
searchupd = UCase(Form2.Text9.Text)
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
 showedit
 
 End If
 .MoveNext
 Loop
 .Close
 End With
 If X <> 1 Then
 MsgBox "Movie Record Not Found ,Enter correctly"
  Form2.Text9.Text = ""
 searchupd = ""

 Exit Sub
 End If
 searchupd = ""
 
End Sub
Public Sub UpdateRecord()

Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim name, image_path, filename, ds As String
    Dim synopsis1, synopsis2, synopsis3, synopsis4, synopsis5, synopsis6, synopsis7, original_synopsis As String
    Dim is_current, no_of_synopsis, extra_length As Integer
    Dim release_date As Date
    Dim f0, f1, f2, f3 As Integer
    f0 = 0
    f1 = 0
    f2 = 0
    f3 = 0

    X = 0
    name = UCase(Form2.Text9.Text)          'Movie Name
      image_folder_path = App.Path + "\db\"
    ds = image_folder_path + "MovieRatingSystem.mdb"
    
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & ds
    Set rs = New ADODB.Recordset
        rs.Open "Select * from MovieDetails where NAME='" & name & " '", cn, adOpenStatic, adLockOptimistic
    
  
                                                      
    
   ' director = Form2.Text6.Text                                                         'Movie Director
                                                    'Movie Cast
   ' genre = Form2.Text5.Text                                                            'Movie Genre
    
   If Form2.Check3.Value = 1 Then
    If Form2.Option3.Value = True Then                                                  'Movie is current
        is_current = 1
        f2 = 1
    End If
    If Form2.Option4.Value = True Then
        is_current = 0
        f2 = 1
    End If
    End If
    If Form2.Check2.Value = 1 Then
    f1 = 1
    release_date = Form2.DTPicker2.Value                                               'Movie Release Date
    End If
    If Form2.Check4.Value = 1 Then
    f3 = 1
    no_of_synopsis = (Len(Form2.Text11.Text) \ 255) + 1  'Breaking the Original Synopsis
    original_synopsis = Trim(Form2.Text11.Text)
    
    If no_of_synopsis = 1 Then
        synopsis1 = original_synopsis
    ElseIf no_of_synopsis = 2 Then
        synopsis1 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis1, "")
        synopsis2 = original_synopsis
    ElseIf no_of_synopsis = 3 Then
        synopsis1 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis1, "")
        synopsis2 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis2, "")
        synopsis3 = original_synopsis
    ElseIf no_of_synopsis = 4 Then
        synopsis1 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis1, "")
        synopsis2 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis2, "")
        synopsis3 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis3, "")
        synopsis4 = original_synopsis
    ElseIf no_of_synopsis = 5 Then
        synopsis1 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis1, "")
        synopsis2 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis2, "")
        synopsis3 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis3, "")
        synopsis4 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis4, "")
        synopsis5 = original_synopsis
    ElseIf no_of_synopsis = 6 Then
        synopsis1 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis1, "")
        synopsis2 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis2, "")
        synopsis3 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis3, "")
        synopsis4 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis4, "")
        synopsis5 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis5, "")
        synopsis6 = original_synopsis
    ElseIf no_of_synopsis = 7 Then
        synopsis1 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis1, "")
        synopsis2 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis2, "")
        synopsis3 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis3, "")
        synopsis4 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis4, "")
        synopsis5 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis5, "")
        synopsis6 = Left(original_synopsis, 255)
        original_synopsis = Replace(original_synopsis, synopsis6, "")
        synopsis7 = original_synopsis
    Else
        f3 = 0
        MsgBox ("Synosis Is Too Long")
    End If
    End If
    If Form2.Check1.Value = 1 Then
     f0 = 1
    
    image_path = Form2.CommonDialog2.filename 'Getting the Image path Ready
    filename = Mid(image_path, InStrRev(image_path, "\") + 1, Len(image_path))
    relative_image_path = "MovieImages\" + filename
    
    If f0 = 1 And image_path = "" Then
    MsgBox ("Choose a Image First")
    Exit Sub
   
    Else
    FileCopy image_path, image_folder_path + relative_image_path
   End If
End If

With rs


'If name = .Fields(0) Then
'X = 1
If .Supports(adUpdate) Then

If f0 = 1 Then
  '.Fields("NAME").Value = name
       .Fields("IMAGE").Value = relative_image_path
        End If
        If f1 = 1 Then
        .Fields("RELEASE DATE").Value = release_date
        End If
       ' .Fields("LANGUAGE").Value = language
       ' .Fields("DIRECTOR").Value = director
      '  .Fields("GENRE").Value = genre
      
        If f2 = 1 Then
        .Fields("CURRENT").Value = is_current
        End If
        If f3 = 1 Then
       
        If (Not (synopsis1 = "")) Then
            .Fields("SYNOPSIS1").Value = synopsis1
        End If
        If (Not (synopsis2 = "")) Then
            .Fields("SYNOPSIS2").Value = synopsis2
        End If
        If (Not (synopsis3 = "")) Then
            .Fields("SYNOPSIS3").Value = synopsis3
        End If
        If (Not (synopsis4 = "")) Then
            .Fields("SYNOPSIS4").Value = synopsis4
        End If
        If (Not (synopsis5 = "")) Then
            .Fields("SYNOPSIS5").Value = synopsis5
        End If
        If (Not (synopsis6 = "")) Then
            .Fields("SYNOPSIS6").Value = synopsis6
        End If
        If (Not (synopsis7 = "")) Then
            .Fields("SYNOPSIS7").Value = synopsis7
        End If
        End If
       
        
        .Update
    
    MsgBox ("Movie Record Successfully Updated")
    
   End If
   .Close
   cn.Close
    End With
      Form2.Text9.Text = ""
    Form2.Text10.Text = ""
      Form2.Text11.Text = ""
        
    Unload Form2
   

  Unload Form1
Load Form1
  Load Form2
  LoadDatabase
    CurrentMoviesArray
    LoadIntialPictures
     Initializations
     
       Form2.Hide
 Form1.Show
 Form1.SSTab1 = 2
 Form2.Command8.Visible = True
 Form2.Command6.Visible = False
 Form2.Text9.Locked = False

End Sub

Public Sub showedit()

MsgBox "Chose &  Enter  Details .", vbOKOnly, "Movie Found"
Form2.Command8.Visible = False
 Form2.Command6.Visible = True
 Form2.Text9.Locked = True
 End Sub
 
