<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<%
Response.ContentType = "application/json"
Response.Charset = "UTF-8"

Dim songId, conn, rs, newPlays
songId = Request.QueryString("id")

If songId = "" Or Not IsNumeric(songId) Then
    Response.Write "{""success"":false, ""error"":""ID de canción inválido""}"
    Response.End
End If

Set conn = GetDBConnection()

' Actualizar contador de reproducciones
conn.Execute "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = " & CLng(songId)

' Consultar nuevo valor
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = " & CLng(songId), conn

If Not rs.EOF Then
    newPlays = rs("PLAYS")
    If IsNull(newPlays) Then newPlays = 0
    Response.Write "{""success"":true, ""plays"":" & newPlays & "}"
Else
    Response.Write "{""success"":false, ""error"":""Canción no encontrada""}"
End If

CloseRecordset rs
CloseDBConnection conn
%>
