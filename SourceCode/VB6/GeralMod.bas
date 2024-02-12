Attribute VB_Name = "GeralMod"
Option Explicit

Public conn As Object

Public Sub OpenConnection()

    Dim strConn As String
    
        Set conn = CreateObject("ADODB.Connection")
        strConn = "DSN=odbc;DATABASE=postgres;SERVER=Localhost;PORT=5432;UID=admin;PWD=admin;"
    
    conn.Open strConn

End Sub


