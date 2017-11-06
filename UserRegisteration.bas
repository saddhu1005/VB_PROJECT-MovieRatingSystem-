Attribute VB_Name = "UserActivity"

Public Sub RegisterUser()
    Dim first_name, last_name, username, password1, password2, final_password As String
    first_name = Form4.Text1.Text
    last_name = Form4.Text2.Text
    username = Form4.Text3.Text
    password1 = Form4.Text4.Text
    password2 = Form4.Text5.Text
    If (password1 = password2) Then
        final_password = password1
    Else
        Form4.Text4.Text = ""
        Form4.Text5.Text = ""
        MsgBox ("Passwords Do Not Match.Please Renter the Passwords !")
        Exit Sub
    End If
    
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim db_path As String
    If (first_name = "" Or last_name = "") Or (username = "" Or final_password = "") Then
   ' Form4.Text1.Text = ""
    ' Form4.Text2.Text = ""
    ' Form4.Text3.Text = ""
    Form4.Text4.Text = ""
    Form4.Text5.Text = ""
    MsgBox "All Fields are Mandatory ,fill the details again", vbokayonly, "Error!!"
    Exit Sub
    Else
    db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from UserDetails", cn, adOpenStatic, adLockOptimistic
        
    With rs
        .AddNew
        .Fields("FIRST NAME").Value = first_name
        .Fields("LAST NAME").Value = last_name
        .Fields("USERNAME").Value = username
        .Fields("PASSWORD").Value = final_password
        .Update
    End With
    
    MsgBox ("User Successfully Registered !")
    Unload Form4
    Form3.Show
    End If
End Sub


Public Sub UserLogin()
    Dim username, password, db_path, sql As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim msg As String
    msg = "Invalid Credentails.Please Enter Valid Credentials!!"
    username = Form3.Text1.Text
    password = Form3.Text2.Text
    sql = "SELECT * FROM UserDetails WHERE UserName = '" & username & "'"
    If sql = Null Or username = "" Or password = "" Then
     MsgBox (msg)
        Form3.Text1.Text = ""
        Form3.Text2.Text = ""
        Exit Sub
        Else
    db_path = App.Path + "\db\" + "MovieRatingSystem.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open sql, cn, adOpenStatic, adLockOptimistic
        
    If rs.Fields("PASSWORD").Value = password Then
        Form1.Show
       
        Unload Form3
    Else
        MsgBox (msg)
        Form3.Text1.Text = ""
        Form3.Text2.Text = ""
        Exit Sub
    End If
    End If
End Sub
