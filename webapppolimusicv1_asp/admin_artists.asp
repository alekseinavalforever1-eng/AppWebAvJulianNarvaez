<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Solo accesible para rol ADMIN (USER_TYPE = 1)
RequireAdmin()

Dim conn, action, msgSuccess, msgError
Set conn = GetDBConnection()

action = LCase(Trim(Request("action")))
msgSuccess = ""
msgError = ""

' Variables de formulario
Dim artistId, artistName, artistLastname, artistImage
artistId = Request("id")
artistName = Trim(Request.Form("artist_name"))
artistLastname = Trim(Request.Form("artist_lastname"))
artistImage = Trim(Request.Form("artist_image"))

' PROCESAR ACCIONES POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "save_new" Then
        If artistName = "" Then
            msgError = "El nombre del artista es obligatorio."
            action = "new"
        Else
            If artistImage = "" Then artistImage = "~/ArtistImages/artist1.jpg"
            
            conn.Execute "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (" & _
                         "'" & SafeSql(artistName) & "', " & _
                         "'" & SafeSql(artistLastname) & "', " & _
                         "'" & SafeSql(artistImage) & "')"
            
            msgSuccess = "¡Artista registrado exitosamente!"
            action = ""
        End If

    ElseIf action = "save_edit" Then
        If artistId = "" Or Not IsNumeric(artistId) Or artistName = "" Then
            msgError = "Datos incompletos para actualizar el artista."
            action = "edit"
        Else
            If artistImage = "" Then artistImage = "~/ArtistImages/artist1.jpg"
            
            conn.Execute "UPDATE TBL_ARTIST SET " & _
                         "ARTIST_NAME = '" & SafeSql(artistName) & "', " & _
                         "ARTIST_LASTNAME = '" & SafeSql(artistLastname) & "', " & _
                         "ARTIST_IMAGE = '" & SafeSql(artistImage) & "' " & _
                         "WHERE ID_ARTIST = " & CLng(artistId)

            msgSuccess = "¡Artista actualizado correctamente!"
            action = ""
        End If
    End If
End If

' PROCESAR ELIMINACIÓN (GET)
If action = "delete" Then
    If artistId <> "" And IsNumeric(artistId) Then
        ' Eliminar referencias en TBL_PLAYER
        conn.Execute "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = " & CLng(artistId)
        ' Eliminar artista
        conn.Execute "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = " & CLng(artistId)
        
        msgSuccess = "¡Artista eliminado satisfactoriamente!"
        action = ""
    End If
End If

' Cargar datos para edición si aplica
Dim editArtName, editArtLast, editArtImg
If action = "edit" Then
    If artistId <> "" And IsNumeric(artistId) Then
        Dim rsEdit
        Set rsEdit = conn.Execute("SELECT ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = " & CLng(artistId))
        If Not rsEdit.EOF Then
            editArtName = rsEdit("ARTIST_NAME")
            editArtLast = rsEdit("ARTIST_LASTNAME")
            editArtImg = rsEdit("ARTIST_IMAGE")
        End If
        CloseRecordset rsEdit
    End If
End If

' Listado de artistas para la tabla
Dim rsArtistsList, searchQ
searchQ = Trim(Request("q"))

Dim sqlList
sqlList = "SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
          "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS TOTAL_SONGS " & _
          "FROM TBL_ARTIST a "

If searchQ <> "" Then
    sqlList = sqlList & " WHERE a.ARTIST_NAME LIKE '%" & SafeSql(searchQ) & "%' OR a.ARTIST_LASTNAME LIKE '%" & SafeSql(searchQ) & "%' "
End If
sqlList = sqlList & " ORDER BY a.ID_ARTIST DESC"

Set rsArtistsList = Server.CreateObject("ADODB.Recordset")
rsArtistsList.Open sqlList, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - CRUD de Artistas (Admin)</title>
  <link rel="stylesheet" href="css/style.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style> body { font-family: 'Plus Jakarta Sans', sans-serif; } </style>
</head>
<body>
  <div class="app-container">
    <!--#include file="includes/sidebar.asp"-->

    <div class="main-wrapper">
      <!--#include file="includes/topbar.asp"-->

      <div class="content-body">
        
        <!-- Encabezado de la página -->
        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-subtitle">Panel de Control de Administrador</span>
            <h1 class="section-title">Gestión de Artistas (CRUD TBL_ARTIST)</h1>
          </div>
          <% If action <> "new" Then %>
          <a href="admin_artists.asp?action=new" class="btn btn-primary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Agregar Nuevo Artista</span>
          </a>
          <% End If %>
        </div>

        <% If msgSuccess <> "" Then %>
        <div class="alert alert-success">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
          <span><%=msgSuccess%></span>
        </div>
        <% End If %>

        <% If msgError <> "" Then %>
        <div class="alert alert-danger">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
          <span><%=msgError%></span>
        </div>
        <% End If %>

        <!-- FORMULARIO DE CREACIÓN / EDICIÓN -->
        <% If action = "new" Or action = "edit" Then %>
        <div class="crud-container">
          <h2 style="font-size:18px; color:var(--primary-light); margin-bottom:16px;">
            <% If action = "new" Then Response.Write "Nuevo Artista" Else Response.Write "Editar Artista #" & artistId %>
          </h2>

          <form method="POST" action="admin_artists.asp?action=<%=IIf(action = "new", "save_new", "save_edit")%>">
            <% If action = "edit" Then %>
            <input type="hidden" name="id" value="<%=artistId%>">
            <% End If %>

            <div class="form-grid">
              <div class="form-group">
                <label for="artist_name">Nombre *</label>
                <input type="text" id="artist_name" name="artist_name" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", artistName, editArtName))%>" 
                       placeholder="Ej. Freddie" required>
              </div>

              <div class="form-group">
                <label for="artist_lastname">Apellido</label>
                <input type="text" id="artist_lastname" name="artist_lastname" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", artistLastname, editArtLast))%>" 
                       placeholder="Ej. Mercury">
              </div>

              <div class="form-group">
                <label for="artist_image">Ruta de la Foto del Artista</label>
                <input type="text" id="artist_image" name="artist_image" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", IIf(artistImage <> "", artistImage, "~/ArtistImages/artist1.jpg"), editArtImg))%>" 
                       placeholder="~/ArtistImages/artist1.jpg">
                <small style="color:var(--text-dim);font-size:11px;">Rutas disponibles: ~/ArtistImages/artist1.jpg a artist5.png</small>
              </div>
            </div>

            <div style="display:flex; gap:12px;">
              <button type="submit" class="btn btn-primary">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
                <span>Guardar Artista</span>
              </button>
              <a href="admin_artists.asp" class="btn btn-secondary">Cancelar</a>
            </div>
          </form>
        </div>
        <% End If %>

        <!-- BUSCADOR RÁPIDO -->
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
          <form method="GET" action="admin_artists.asp" style="display:flex; gap:10px;">
            <input type="text" name="q" value="<%=Server.HTMLEncode(searchQ)%>" placeholder="Buscar por nombre o apellido..." class="form-control" style="width:260px; padding:6px 12px; font-size:13px;">
            <button type="submit" class="btn btn-secondary btn-sm">Buscar</button>
            <% If searchQ <> "" Then %>
              <a href="admin_artists.asp" class="btn btn-sm" style="color:var(--text-dim);">Limpiar</a>
            <% End If %>
          </form>
        </div>

        <!-- TABLA DE LISTADO DE ARTISTAS -->
        <div class="song-table-container">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width:50px;">ID</th>
                <th>Foto</th>
                <th>Nombre Completo</th>
                <th>Canciones Asociadas</th>
                <th>Ruta de Imagen</th>
                <th style="text-align:right;">Acciones</th>
              </tr>
            </thead>
            <tbody>
              <%
              If rsArtistsList.EOF Then
              %>
              <tr>
                <td colspan="6" style="text-align:center;padding:30px;color:var(--text-muted);">
                  No hay artistas registrados.
                </td>
              </tr>
              <%
              Else
                Do While Not rsArtistsList.EOF
                  Dim rowArtId, rowArtName, rowArtLast, rowArtFull, rowArtImg, rowSongCount
                  rowArtId = rsArtistsList("ID_ARTIST")
                  rowArtName = rsArtistsList("ARTIST_NAME")
                  rowArtLast = rsArtistsList("ARTIST_LASTNAME")
                  rowArtFull = FormatArtistFullName(rowArtName, rowArtLast)
                  rowArtImg = ResolveMediaUrl(rsArtistsList("ARTIST_IMAGE"), "ArtistImages/artist1.jpg")
                  rowSongCount = rsArtistsList("TOTAL_SONGS")
              %>
              <tr>
                <td style="color:var(--text-dim);font-weight:600;"><%=rowArtId%></td>
                <td>
                  <img src="<%=rowArtImg%>" alt="<%=Server.HTMLEncode(rowArtFull)%>" style="width:40px;height:40px;border-radius:50%;object-fit:cover;" onerror="this.src='ArtistImages/artist1.jpg'">
                </td>
                <td style="font-weight:600;color:#ffffff;"><%=Server.HTMLEncode(rowArtFull)%></td>
                <td style="color:var(--primary-light);"><%=rowSongCount%> canciones</td>
                <td style="color:var(--text-dim);font-size:12px;font-family:monospace;"><%=Server.HTMLEncode(rsArtistsList("ARTIST_IMAGE"))%></td>
                <td style="text-align:right;white-space:nowrap;">
                  <a href="admin_artists.asp?action=edit&id=<%=rowArtId%>" class="btn btn-secondary btn-sm" style="padding:4px 10px;margin-right:6px;" title="Editar">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                    <span>Editar</span>
                  </a>
                  <a href="admin_artists.asp?action=delete&id=<%=rowArtId%>" 
                     class="btn btn-danger btn-sm" 
                     style="padding:4px 10px;"
                     onclick="return confirm('¿Está seguro de eliminar al artista <%=Server.HTMLEncode(rowArtFull)%>?');"
                     title="Eliminar">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                    <span>Eliminar</span>
                  </a>
                </td>
              </tr>
              <%
                  rsArtistsList.MoveNext
                Loop
              End If
              %>
            </tbody>
          </table>
        </div>

      </div>
    </div>
  </div>

  <!--#include file="includes/player_bar.asp"-->
</body>
</html>
<%
CloseRecordset rsArtistsList
CloseDBConnection conn
%>
