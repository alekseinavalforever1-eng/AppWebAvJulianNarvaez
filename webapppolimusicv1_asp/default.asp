<%@ Language="VBScript" %>
<%
' Default redirector
If Not IsEmpty(Session("ID_USER")) And Session("ID_USER") <> "" Then
    Response.Redirect "index.asp"
Else
    Response.Redirect "login.asp"
End If
%>
