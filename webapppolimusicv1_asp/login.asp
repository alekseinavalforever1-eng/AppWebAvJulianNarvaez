<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<%
Dim errorMessage, usernameInput
errorMessage = ""
usernameInput = ""

' Si ya está autenticado, redirigir al index
If Not IsEmpty(Session("ID_USER")) And Session("ID_USER") <> "" Then
    Response.Redirect "index.asp"
    Response.End
End If

' Procesar POST de Login
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim postUser, postPass
    postUser = Trim(Request.Form("username"))
    postPass = Trim(Request.Form("password"))
    usernameInput = postUser

    If postUser = "" Or postPass = "" Then
        errorMessage = "Por favor ingrese tanto el usuario como la contraseña."
    Else
        Dim conn, rs, sql
        Set conn = GetDBConnection()
        
        sql = "SELECT ID_USER, USERNAME, PASSWORD, EMAIL, USER_PHOTO, USER_TYPE " & _
              "FROM TBL_USER WHERE USERNAME = '" & SafeSql(postUser) & "' AND PASSWORD = '" & SafeSql(postPass) & "'"
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, conn
        
        If Not rs.EOF Then
            ' Almacenar datos en variables de sesión
            Session("ID_USER") = rs("ID_USER")
            Session("USERNAME") = rs("USERNAME")
            Session("USER_TYPE") = CInt(rs("USER_TYPE"))
            Session("USER_PHOTO") = rs("USER_PHOTO")
            Session("EMAIL") = rs("EMAIL")
            
            CloseRecordset rs
            CloseDBConnection conn
            
            ' Redirigir a la página principal
            Response.Redirect "index.asp"
            Response.End
        Else
            errorMessage = "Credenciales incorrectas. Verifique el usuario y la contraseña."
            CloseRecordset rs
            CloseDBConnection conn
        End If
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Iniciar Sesión</title>
  <link rel="stylesheet" href="css/style.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
    }
  </style>
</head>
<body class="login-body">

  <div class="login-card">
    <div class="login-header">
      <div class="brand-icon">
        <svg width="30" height="30" viewBox="0 0 24 24" fill="#120e1f">
          <path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/>
        </svg>
      </div>
      <h1>PoliMusic</h1>
      <p>Clon de Spotify en ASP Clásico &amp; VBScript</p>
    </div>

    <% If errorMessage <> "" Then %>
    <div class="alert alert-danger">
      <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
      </svg>
      <span><%=Server.HTMLEncode(errorMessage)%></span>
    </div>
    <% End If %>

    <form method="POST" action="login.asp" id="loginForm">
      <div class="form-group" style="margin-bottom: 18px;">
        <label for="username">Usuario</label>
        <input type="text" id="username" name="username" class="form-control" 
               placeholder="Ej. userNormal o userAdmin" 
               value="<%=Server.HTMLEncode(usernameInput)%>" required autofocus>
      </div>

      <div class="form-group" style="margin-bottom: 24px;">
        <label for="password">Contraseña</label>
        <input type="password" id="password" name="password" class="form-control" 
               placeholder="Ingrese su contraseña" required>
      </div>

      <button type="submit" class="btn btn-primary" style="width: 100%; padding: 12px; font-size: 15px;">
        Iniciar Sesión
      </button>
    </form>

    <!-- Credenciales de prueba de la Base de Datos -->
    <div class="demo-credentials-box">
      <h4>Usuarios de prueba (BDD_PoliMusic):</h4>
      <div style="margin-bottom: 8px;">
        <strong>Estudiante (Admin):</strong> <code>Julian.Narvaez</code> / <code>12345678</code>
        <button type="button" class="demo-quick-btn" onclick="fillCredentials('Julian.Narvaez', '12345678')">Autocompletar Julián Narváez</button>
      </div>
      <div style="margin-bottom: 8px;">
        <strong>Admin General:</strong> <code>userAdmin</code> / <code>12345678</code>
        <button type="button" class="demo-quick-btn" onclick="fillCredentials('userAdmin', '12345678')">Autocompletar Admin</button>
      </div>
      <div>
        <strong>Normal:</strong> <code>userNormal</code> / <code>12345678</code>
        <button type="button" class="demo-quick-btn" onclick="fillCredentials('userNormal', '12345678')">Autocompletar Normal</button>
      </div>
    </div>

    <div style="text-align:center; margin-top:20px; font-size:12px; color:var(--text-dim);">
      Desarrollado por <strong>Julián Narváez</strong> &bull; EPN
    </div>
  </div>

  <script>
    function fillCredentials(u, p) {
      document.getElementById('username').value = u;
      document.getElementById('password').value = p;
    }
  </script>
</body>
</html>
