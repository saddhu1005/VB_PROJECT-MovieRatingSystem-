Attribute VB_Name = "FormLoadActivities"
Public db As Database
Public cs As Recordset
Public current_movies As Variant
Public total_current_records As Integer
Public index As Integer
Public image_folder_path As String
Public all_movies As Variant
Public total_records As Integer
Public Pic As Picture


Public Sub LoadDatabase()
     image_folder_path = App.Path + "\db\"                                                 'Setting the path to Image Folder
     Set db = OpenDatabase(image_folder_path + "MovieRatingSystem.mdb")                    'Loading the Database
     Set cs = db.OpenRecordset("select * from MovieDetails where CURRENT=1")               'Loading the RecordSet
End Sub

Public Sub CurrentMoviesArray()
    cs.MoveLast
    total_current_records = cs.RecordCount                                               'Counting total number of current movies
    cs.MoveFirst
    current_movies = cs.GetRows(total_current_records)                                   'Making array of current_movies
End Sub

Public Sub LoadIntialPictures()
  
    
    If total_current_records >= 3 Then
        
        index = 0                                                                        'Setting Index to zero
        
        Form1.Picture1.AutoRedraw = True
        Set Pic = LoadPicture(image_folder_path + current_movies(1, index))
        Form1.Picture1.PaintPicture Pic, 0, 0, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight
        Set Form1.Picture1.Picture = Form1.Picture1.Image
        Form1.Label2.Caption = current_movies(0, i)
        
        Form1.Picture2.AutoRedraw = True
        Set Pic = LoadPicture(image_folder_path + current_movies(1, index + 1))
        Form1.Picture2.PaintPicture Pic, 0, 0, Form1.Picture2.ScaleWidth, Form1.Picture2.ScaleHeight
        Set Form1.Picture2.Picture = Form1.Picture2.Image
        Form1.Label3.Caption = current_movies(0, i + 1)
        
        Form1.Picture3.AutoRedraw = True
        Set Pic = LoadPicture(image_folder_path + current_movies(1, index + 2))
        Form1.Picture3.PaintPicture Pic, 0, 0, Form1.Picture3.ScaleWidth, Form1.Picture3.ScaleHeight
        Set Form1.Picture3.Picture = Form1.Picture3.Image
        Form1.Label4.Caption = current_movies(0, i + 2)
   
        End If
        
End Sub
Public Sub Initializations()
        
        Form1.VScroll1.Value = 0                                                                 'Setting Values For Vertical scroll bar
        Form1.VScroll1.Max = total_current_records - 3
        Form1.VScroll1.Min = 0
        
        Form1.SSTab1.TabVisible(0) = False                                                        'Hiding tabs 0 and 1 in sstab1
        Form1.SSTab1.TabVisible(1) = False
End Sub
