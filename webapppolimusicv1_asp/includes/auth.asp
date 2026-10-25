<%
' ============================================================================
' PoliMusic - Authentication and Role Management
' ============================================================================

Sub RequireAuth()
    If IsEmpty(Session("ID_USER")) Or Session("ID_USER") = "" Then
        Response.Redirect "login.asp"
        Response.End
    End If
End Sub

Sub RequireAdmin()
    RequireAuth()
    If CInt("0" & Session("USER_TYPE")) <> 1 Then
        Response.Write "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Acceso Denegado</title>"
        Response.Write "<link rel='stylesheet' href='css/style.css'></head><body>"
        Response.Write "<div style='display:flex;align-items:center;justify-content:center;height:100vh;flex-direction:column;color:#fff;text-align:center;padding:20px;'>"
        Response.Write "<div class='glass-panel' style='padding:40px;border-radius:16px;max-width:500px;border:1px solid rgba(179,157,219,0.3);'>"
        Response.Write "<h1 style='color:#b39ddb;margin-bottom:15px;font-size:24px;'>Acceso Restringido</h1>"
        Response.Write "<p style='color:#d1c4e9;margin-bottom:25px;'>Se requieren permisos de Administrador (USER_TYPE = 1) para acceder a este módulo de administración.</p>"
        Response.Write "<a href='index.asp' class='btn btn-primary' style='display:inline-block;padding:10px 24px;'>Volver al Inicio</a>"
        Response.Write "</div></div></body></html>"
        Response.End
    End If
End Sub

Function IsAdmin()
    If Not IsEmpty(Session("USER_TYPE")) And CInt("0" & Session("USER_TYPE")) = 1 Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
End Function

Function GetCurrentUsername()
    If Not IsEmpty(Session("USERNAME")) Then
        GetCurrentUsername = Session("USERNAME")
    Else
        GetCurrentUsername = "Invitado"
    End If
End Function

Function GetCurrentUserPhoto()
    If Not IsEmpty(Session("USER_PHOTO")) And Trim(Session("USER_PHOTO")) <> "" Then
        GetCurrentUserPhoto = ResolveMediaUrl(Session("USER_PHOTO"), "userImages/default.jpg")
    Else
        GetCurrentUserPhoto = "userImages/default.jpg"
    End If
End Function
%>
