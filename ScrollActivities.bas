Attribute VB_Name = "ScrollActivities"
Public Sub ScrollChange()
    Dim j As Integer
    Dim Pic As Picture
    j = Form1.VScroll1.Value
    If j > index Then
        Form1.Picture1.Picture = Form1.Picture2.Picture
        Form1.Label2.Caption = Form1.Label3.Caption
        
        Form1.Picture2.Picture = Form1.Picture3.Picture
        Form1.Label3.Caption = Form1.Label4.Caption
        
        Form1.Picture3.AutoRedraw = True
        Set Pic = LoadPicture(image_folder_path + current_movies(1, j + 2))
        Form1.Picture3.PaintPicture Pic, 0, 0, Form1.Picture3.ScaleWidth, Form1.Picture3.ScaleHeight
        Set Form1.Picture3.Picture = Form1.Picture3.Image
        Form1.Label4.Caption = current_movies(0, j + 2)
        index = j
    End If
    If j < index Then
        Form1.Picture3.Picture = Form1.Picture2.Picture
        Form1.Label4.Caption = Form1.Label3.Caption
        
        Form1.Picture2.Picture = Form1.Picture1.Picture
        Form1.Label3.Caption = Form1.Label2.Caption
    
        Form1.Picture1.AutoRedraw = True
        Set Pic = LoadPicture(image_folder_path + current_movies(1, j))
        Form1.Picture1.PaintPicture Pic, 0, 0, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight
        Set Form1.Picture1.Picture = Form1.Picture1.Image
        Form1.Label2.Caption = current_movies(0, j)
        index = j
    End If
End Sub
