Attribute VB_Name = "ClickEvents"

Public Sub ItemClick(index As Integer)
    Form1.SSTab1.TabVisible(0) = True
    Form1.SSTab1.TabVisible(1) = True
    Form1.SSTab1.TabVisible(2) = True
    Form1.SSTab1.Tab = 0
    
    Form1.Picture4.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + current_movies(1, index))
    Form1.Picture4.PaintPicture Pic, 0, 0, Form1.Picture4.ScaleWidth, Form1.Picture4.ScaleHeight
    Set Form1.Picture4.Picture = Form1.Picture4.Image
    
    Form1.Label5.Caption = current_movies(0, index)
        
    If (Not (current_movies(2, index) = "")) Then
        Form1.Label11.Caption = "RATING: " + CStr(current_movies(2, index))
    Else
        Form1.Label11.Caption = "RATING: Not Rated "
    End If
    
    
    Form1.Label6.Caption = "RELEASE DATE: " + CStr(current_movies(3, index))
    
    Dim Synopsis As String
    For k = 8 To 14
        If (Not (current_movies(k, index) = "")) Then
            Synopsis = Synopsis + current_movies(k, index)
        End If
    Next
    Form1.Label13.Caption = Synopsis
    
    Form1.Label7.Caption = "LANGUAGE: " + current_movies(4, index)
    
    Form1.Label8.Caption = "DIRECTOR: " + current_movies(5, index)
    
    Form1.Label9.Caption = "GENRE: " + current_movies(6, index)
    
    Form1.Label10.Caption = "CAST: " + current_movies(7, index)
    
    'RATINGS AND REVIEWS-->

 Form1.Picture5.AutoRedraw = True                                      'SETTING IMAGE
    Set Pic = LoadPicture(image_folder_path + current_movies(1, index))
    Form1.Picture5.PaintPicture Pic, 0, 0, Form1.Picture5.ScaleWidth, Form1.Picture5.ScaleHeight
    Set Form1.Picture5.Picture = Form1.Picture5.Image
    
      Form1.Label14.Caption = current_movies(0, index)  'MOVIE NAME
     
    If (Not (current_movies(2, 0) = "")) Then
        Form1.Label16.Caption = "" + CStr(current_movies(2, index))
    Else
        Form1.Label16.Caption = "Not Rated "
    End If
  
     ' Form1.Label16.Caption = "*&rating   'Rating"
      Form1.Label15.Caption = "Uploading Soon"  'reviews
End Sub
Public Sub FirstItemClick()
    ItemClick index
End Sub
Public Sub SecondItemClick()
    ItemClick index + 1
End Sub
Public Sub ThirdItemClick()
    ItemClick index + 2

End Sub
Public Sub searchitem()
Dim searchjob As String
Dim X, loc As Integer
searchjob = UCase(Form1.searchtxt.Text)
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
 If searchjob = .Fields(0) Then
 X = 1
 
   image_folder_path = App.Path + "\db\"                                                 'Setting the path to Image Folder
     Set db = OpenDatabase(image_folder_path + "MovieRatingSystem.mdb")                    'Loading the Database
     Set cs = db.OpenRecordset("select * from MovieDetails")                               'Loading the RecordSet
 cs.MoveLast
    total_records = cs.RecordCount                                               'Counting total number of current movies
    cs.MoveFirst
    all_movies = cs.GetRows(total_records)                                 'Making array of ALL_movies
Form1.searchtxt.Text = ""
  Form1.SSTab1.TabVisible(0) = True
    Form1.SSTab1.TabVisible(1) = True
    Form1.SSTab1.Tab = 0
     Form1.SSTab1.TabVisible(2) = True
     Form1.Picture4.AutoRedraw = True
    Set Pic = LoadPicture(image_folder_path + all_movies(1, loc))
    Form1.Picture4.PaintPicture Pic, 0, 0, Form1.Picture4.ScaleWidth, Form1.Picture4.ScaleHeight
    Set Form1.Picture4.Picture = Form1.Picture4.Image
    
    
    
    Form1.Label5.Caption = all_movies(0, loc)
        
    If (Not (all_movies(2, loc) = "")) Then
        Form1.Label11.Caption = "RATING: " + CStr(all_movies(2, loc))
    Else
        Form1.Label11.Caption = "RATING: Not Rated"
    End If
    
    
    Form1.Label6.Caption = "RELEASE DATE: " + CStr(all_movies(3, loc))
    
    
    Dim Synopsis As String
    For k = 8 To 14
        If Not (all_movies(k, loc) = "") Then
            Synopsis = Synopsis + all_movies(k, loc)
        End If
    Next
    Form1.Label13.Caption = Synopsis
    
    Form1.Label7.Caption = "LANGUAGE: " + all_movies(4, loc)
    
    Form1.Label8.Caption = "DIRECTOR: " + all_movies(5, loc)
    
    Form1.Label9.Caption = "GENRE: " + all_movies(6, loc)
    
    Form1.Label10.Caption = "CAST: " + all_movies(7, loc)
    
'RATINGS AND REVIEWS-->

 Form1.Picture5.AutoRedraw = True                                      'SETTING IMAGE
    Set Pic = LoadPicture(image_folder_path + all_movies(1, loc))
    Form1.Picture5.PaintPicture Pic, 0, 0, Form1.Picture5.ScaleWidth, Form1.Picture5.ScaleHeight
    Set Form1.Picture5.Picture = Form1.Picture5.Image
    
      Form1.Label14.Caption = all_movies(0, loc)  'MOVIE NAME
      
      If (Not (all_movies(2, loc) = "")) Then
        Form1.Label16.Caption = "" + CStr(all_movies(2, loc))
    Else
        Form1.Label16.Caption = "Not Rated "
    End If
     ' Form1.Label16.Caption = "Not rated "   'Rating
      Form1.Label15.Caption = "Uploading Soon "  'reviews
      
    
End If
loc = loc + 1
.MoveNext
Loop
.Close
End With
If X <> 1 Then
MsgBox "No Record Found, Please type correctly !", vbOKOnly, "! Sorry !"
Form1.searchtxt.Text = ""
Form1.searchtxt.SetFocus
Exit Sub
End If

End Sub
