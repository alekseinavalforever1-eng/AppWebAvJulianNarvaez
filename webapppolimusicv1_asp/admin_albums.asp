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
Dim albumId, albumTitle, releaseYear, albumImagePath
albumId = Request("id")
albumTitle = Trim(Request.Form("title_album"))
releaseYear = Trim(Request.Form("release_year"))
albumImagePath = Trim(Request.Form("album_image_path"))

' PROCESAR ACCIONES POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "save_new" Then
        If albumTitle = "" Or releaseYear = "" Or Not IsNumeric(releaseYear) Then
            msgError = "El título del álbum y un año de lanzamiento válido son obligatorios."
            action = "new"
        Else
            If albumImagePath = "" Then albumImagePath = "~/albumImages/album1.png"
            
            conn.Execute "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (" & _
                         "'" & SafeSql(albumTitle) & "', " & _
                         CLng(releaseYear) & ", " & _
                         "'" & SafeSql(albumImagePath) & "')"
            
            msgSuccess = "¡Álbum registrado exitosamente!"
            action = ""
        End If

    ElseIf action = "save_edit" Then
        If albumId = "" Or Not IsNumeric(albumId) Or albumTitle = "" Or releaseYear = "" Or Not IsNumeric(releaseYear) Then
            msgError = "Datos incompletos para actualizar el álbum."
            action = "edit"
        Else
            If albumImagePath = "" Then albumImagePath = "~/albumImages/album1.png"
            
            conn.Execute "UPDATE TBL_ALBUM SET " & _
                         "TITLE_ALBUM = '" & SafeSql(albumTitle) & "', " & _
                         "RELEASE_YEAR = " & CLng(releaseYear) & ", " & _
                         "ALBUM_IMAGE_PATH = '" & SafeSql(albumImagePath) & "' " & _
                         "WHERE ID_ALBUM = " & CLng(albumId)

            msgSuccess = "¡Álbum actualizado correctamente!"
            action = ""
        End If
    End If
End If

' PROCESAR ELIMINACIÓN (GET)
If action = "delete" Then
    If albumId <> "" And IsNumeric(albumId) Then
        ' Eliminar referencias en TBL_TRACKLIST
        conn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & CLng(albumId)
        ' Eliminar álbum
        conn.Execute "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = " & CLng(albumId)
        
        msgSuccess = "¡Álbum eliminado satisfactoriamente!"
        action = ""
    End If
End If

' Cargar datos para edición si aplica
Dim editAlbTitle, editAlbYear, editAlbImg
If action = "edit" Then
    If albumId <> "" And IsNumeric(albumId) Then
        Dim rsEdit
        Set rsEdit = conn.Execute("SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = " & CLng(albumId))
        If Not rsEdit.EOF Then
            editAlbTitle = rsEdit("TITLE_ALBUM")
            editAlbYear = rsEdit("RELEASE_YEAR")
            editAlbImg = rsEdit("ALBUM_IMAGE_PATH")
        End If
        CloseRecordset rsEdit
    End If
End If

' Listado de álbumes para la tabla
Dim rsAlbumsList, searchQ
searchQ = Trim(Request("q"))

Dim sqlList
sqlList = "SELECT a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
          "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS TOTAL_TRACKS " & _
          "FROM TBL_ALBUM a "

If searchQ <> "" Then
    sqlList = sqlList & " WHERE a.TITLE_ALBUM LIKE '%" & SafeSql(searchQ) & "%' "
End If
sqlList = sqlList & " ORDER BY a.ID_ALBUM DESC"

Set rsAlbumsList = Server.CreateObject("ADODB.Recordset")
rsAlbumsList.Open sqlList, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - CRUD de Álbumes (Admin)</title>
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
            <h1 class="section-title">Gestión de Álbumes (CRUD TBL_ALBUM)</h1>
          </div>
          <% If action <> "new" Then %>
          <a href="admin_albums.asp?action=new" class="btn btn-primary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Agregar Nuevo Álbum</span>
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
            <% If action = "new" Then Response.Write "Nuevo Álbum" Else Response.Write "Editar Álbum #" & albumId %>
          </h2>

          <form method="POST" action="admin_albums.asp?action=<%=IIf(action = "new", "save_new", "save_edit")%>">
            <% If action = "edit" Then %>
            <input type="hidden" name="id" value="<%=albumId%>">
            <% End If %>

            <div class="form-grid">
              <div class="form-group">
                <label for="title_album">Título del Álbum *</label>
                <input type="text" id="title_album" name="title_album" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", albumTitle, editAlbTitle))%>" 
                       placeholder="Ej. Abbey Road" required>
              </div>

              <div class="form-group">
                <label for="release_year">Año de Lanzamiento *</label>
                <input type="number" id="release_year" name="release_year" class="form-control" 
                       value="<%=IIf(action = "new", IIf(releaseYear <> "", releaseYear, "2026"), editAlbYear)%>" 
                       min="1900" max="2100" required>
              </div>

              <div class="form-group">
                <label for="album_image_path">Ruta de la Portada</label>
                <input type="text" id="album_image_path" name="album_image_path" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", IIf(albumImagePath <> "", albumImagePath, "~/albumImages/album1.png"), editAlbImg))%>" 
                       placeholder="~/albumImages/album1.png">
                <small style="color:var(--text-dim);font-size:11px;">Rutas disponibles: ~/albumImages/album1.png a album5.png</small>
              </div>
            </div>

            <div style="display:flex; gap:12px;">
              <button type="submit" class="btn btn-primary">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
                <span>Guardar Álbum</span>
              </button>
              <a href="admin_albums.asp" class="btn btn-secondary">Cancelar</a>
            </div>
          </form>
        </div>
        <% End If %>

        <!-- BUSCADOR RÁPIDO -->
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
          <form method="GET" action="admin_albums.asp" style="display:flex; gap:10px;">
            <input type="text" name="q" value="<%=Server.HTMLEncode(searchQ)%>" placeholder="Buscar por título..." class="form-control" style="width:260px; padding:6px 12px; font-size:13px;">
            <button type="submit" class="btn btn-secondary btn-sm">Buscar</button>
            <% If searchQ <> "" Then %>
              <a href="admin_albums.asp" class="btn btn-sm" style="color:var(--text-dim);">Limpiar</a>
            <% End If %>
          </form>
        </div>

        <!-- TABLA DE LISTADO DE ÁLBUMES -->
        <div class="song-table-container">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width:50px;">ID</th>
                <th>Portada</th>
                <th>Título del Álbum</th>
                <th>Año</th>
                <th>Canciones</th>
                <th>Ruta de Portada</th>
                <th style="text-align:right;">Acciones</th>
              </tr>
            </thead>
            <tbody>
              <%
              If rsAlbumsList.EOF Then
              %>
              <tr>
                <td colspan="7" style="text-align:center;padding:30px;color:var(--text-muted);">
                  No hay álbumes registrados.
                </td>
              </tr>
              <%
              Else
                Do While Not rsAlbumsList.EOF
                  Dim rowAlbId, rowAlbTitle, rowAlbYear, rowAlbImg, rowTrackCount
                  rowAlbId = rsAlbumsList("ID_ALBUM")
                  rowAlbTitle = rsAlbumsList("TITLE_ALBUM")
                  rowAlbYear = rsAlbumsList("RELEASE_YEAR")
                  rowAlbImg = ResolveMediaUrl(rsAlbumsList("ALBUM_IMAGE_PATH"), "albumImages/album1.png")
                  rowTrackCount = rsAlbumsList("TOTAL_TRACKS")
              %>
              <tr>
                <td style="color:var(--text-dim);font-weight:600;"><%=rowAlbId%></td>
                <td>
                  <img src="<%=rowAlbImg%>" alt="<%=Server.HTMLEncode(rowAlbTitle)%>" style="width:40px;height:40px;border-radius:4px;object-fit:cover;" onerror="this.src='albumImages/album1.png'">
                </td>
                <td style="font-weight:600;color:#ffffff;"><%=Server.HTMLEncode(rowAlbTitle)%></td>
                <td style="color:var(--text-dim);"><%=rowAlbYear%></td>
                <td style="color:var(--primary-light);"><%=rowTrackCount%> pistas</td>
                <td style="color:var(--text-dim);font-size:12px;font-family:monospace;"><%=Server.HTMLEncode(rsAlbumsList("ALBUM_IMAGE_PATH"))%></td>
                <td style="text-align:right;white-space:nowrap;">
                  <a href="admin_albums.asp?action=edit&id=<%=rowAlbId%>" class="btn btn-secondary btn-sm" style="padding:4px 10px;margin-right:6px;" title="Editar">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                    <span>Editar</span>
                  </a>
                  <a href="admin_albums.asp?action=delete&id=<%=rowAlbId%>" 
                     class="btn btn-danger btn-sm" 
                     style="padding:4px 10px;"
                     onclick="return confirm('¿Está seguro de eliminar el álbum <%=Server.HTMLEncode(rowAlbTitle)%>?');"
                     title="Eliminar">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                    <span>Eliminar</span>
                  </a>
                </td>
              </tr>
              <%
                  rsAlbumsList.MoveNext
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
CloseRecordset rsAlbumsList
CloseDBConnection conn
%>
