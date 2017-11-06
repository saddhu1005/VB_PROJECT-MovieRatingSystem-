Attribute VB_Name = "RecordAddition"
Public Sub AddRecord()
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim name, director, genre, language, cast, image_path, relative_image_path, image_folder_path, filename, ds As String
    Dim synopsis1, synopsis2, synopsis3, synopsis4, synopsis5, synopsis6, synopsis7, original_synopsis As String
    Dim is_current, no_of_synopsis, extra_length As Integer
    Dim release_date As Date
      image_folder_path = App.Path + "\db\"
    ds = image_folder_path + "MovieRatingSystem.mdb"
    
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & ds
    Set rs = New ADODB.Recordset
        rs.Open "Select * from MovieDetails", cn, adOpenStatic, adLockOptimistic
    
    
    name = UCase(Form2.Text1.Text)                                                      'Movie Name
    
    director = Trim(Form2.Text6.Text)                                                         'Movie Director
    
    cast = Trim(Form2.Text7.Text)                                                             'Movie Cast
    
    genre = Trim(Form2.Text5.Text)                                                           'Movie Genre
    
    language = Trim(Form2.Text4.Text)                                                         'Movie Language
    
    If Form2.Option1.Value = True Then                                                  'Movie is current
        is_current = 1
    End If
    If Form2.Option2.Value = True Then
        is_current = 0
    End If
    
    release_date = Form2.DTPicker1.Value                                               'Movie Release Date
    
    no_of_synopsis = (Len(Form2.Text8.Text) \ 255) + 1                                 'Breaking the Original Synopsis
    original_synopsis = Form2.Text8.Text
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
        MsgBox ("Synosis Is Too Long")
    End If
    
        
    image_path = Form2.CommonDialog1.filename 'Getting the Image path Ready
    filename = Mid(image_path, InStrRev(image_path, "\") + 1, Len(image_path))
    relative_image_path = "MovieImages\" + filename
    If image_path = "" Then
    MsgBox ("Choose a Image First")
    Exit Sub
    ElseIf ((name = "") Or (language = "") Or (director = "") Or (genre = "") Or (cast = "")) Then
    MsgBox "Fill All the Fields ", vbOKOnly, "Error!"
    Exit Sub
    Else
    FileCopy image_path, image_folder_path + relative_image_path
    
    With rs
        .AddNew
        .Fields("NAME").Value = name
        .Fields("IMAGE").Value = relative_image_path
        .Fields("RELEASE DATE").Value = release_date
        .Fields("LANGUAGE").Value = language
        .Fields("DIRECTOR").Value = director
        .Fields("GENRE").Value = genre
        .Fields("CAST").Value = cast
        .Fields("CURRENT").Value = is_current
        If (Not (synopsis1 = "")) Then
            .Fields("SYNOPSIS1") = synopsis1
        End If
        If (Not (synopsis2 = "")) Then
            .Fields("SYNOPSIS2") = synopsis2
        End If
        If (Not (synopsis3 = "")) Then
            .Fields("SYNOPSIS3") = synopsis3
        End If
        If (Not (synopsis4 = "")) Then
            .Fields("SYNOPSIS4") = synopsis4
        End If
        If (Not (synopsis5 = "")) Then
            .Fields("SYNOPSIS5") = synopsis5
        End If
        If (Not (synopsis6 = "")) Then
            .Fields("SYNOPSIS1") = synopsis6
        End If
        If (Not (synopsis7 = "")) Then
            .Fields("SYNOPSIS7") = synopsis7
        End If
        .Update
        .Close
    End With
    
    MsgBox ("Movie Successfully Added")
    cn.Close
    
    Form2.Text1.Text = ""
    Form2.Text2.Text = ""
    Form2.Text4.Text = ""
    Form2.Text5.Text = ""
    Form2.Text6.Text = ""
    Form2.Text7.Text = ""
    Form2.Text8.Text = ""
         
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
End Sub
