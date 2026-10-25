<%@ Language="VBScript" %>
<%
' ============================================================================
' PoliMusic - Logout Handler
' Clears session and cookies, then redirects to login
' ============================================================================
Session.Contents.RemoveAll()
Session.Abandon()
Response.Redirect "login.asp"
%>
