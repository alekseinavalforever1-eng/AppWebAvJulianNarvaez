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
Dim songId, songName, songPath, songPlays, artistId, albumId
songId = Request("id")
songName = Trim(Request.Form("song_name"))
songPath = Trim(Request.Form("song_path"))
songPlays = Trim(Request.Form("plays"))
artistId = Trim(Request.Form("artist_id"))
albumId = Trim(Request.Form("album_id"))

' PROCESAR ACCIONES POST
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If action = "save_new" Then
        If songName = "" Or songPath = "" Then
            msgError = "El nombre de la canción y la ruta del archivo son obligatorios."
            action = "new"
        Else
            If Not IsNumeric(songPlays) Then songPlays = 0
            
            ' Insertar Canción
            Dim rsInsert, newId
            Set rsInsert = conn.Execute("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES ('" & SafeSql(songName) & "', '" & SafeSql(songPath) & "', " & CLng(songPlays) & "); SELECT SCOPE_IDENTITY() AS NEW_ID;")
            
            If Not rsInsert.EOF Then
                newId = rsInsert("NEW_ID")
            Else
                Dim rsId
                Set rsId = conn.Execute("SELECT MAX(ID_SONG) FROM TBL_SONG")
                newId = rsId(0)
                CloseRecordset rsId
            End If
            CloseRecordset rsInsert

            ' Asociar Artista en TBL_PLAYER
            If artistId <> "" And IsNumeric(artistId) Then
                conn.Execute "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (" & CLng(newId) & ", " & CLng(artistId) & ")"
            End If

            ' Asociar Álbum en TBL_TRACKLIST
            If albumId <> "" And IsNumeric(albumId) Then
                conn.Execute "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (" & CLng(albumId) & ", " & CLng(newId) & ")"
            End If

            msgSuccess = "¡Canción creada exitosamente!"
            action = ""
        End If

    ElseIf action = "save_edit" Then
        If songId = "" Or Not IsNumeric(songId) Or songName = "" Or songPath = "" Then
            msgError = "Datos incompletos para actualizar la canción."
            action = "edit"
        Else
            If Not IsNumeric(songPlays) Then songPlays = 0
            
            ' Actualizar TBL_SONG
            conn.Execute "UPDATE TBL_SONG SET " & _
                         "SONG_NAME = '" & SafeSql(songName) & "', " & _
                         "SONG_PATH = '" & SafeSql(songPath) & "', " & _
                         "PLAYS = " & CLng(songPlays) & " " & _
                         "WHERE ID_SONG = " & CLng(songId)

            ' Actualizar Artista en TBL_PLAYER
            conn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & CLng(songId)
            If artistId <> "" And IsNumeric(artistId) Then
                conn.Execute "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (" & CLng(songId) & ", " & CLng(artistId) & ")"
            End If

            ' Actualizar Álbum en TBL_TRACKLIST
            conn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & CLng(songId)
            If albumId <> "" And IsNumeric(albumId) Then
                conn.Execute "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (" & CLng(albumId) & ", " & CLng(songId) & ")"
            End If

            msgSuccess = "¡Canción actualizada correctamente!"
            action = ""
        End If
    End If
End If

' PROCESAR ELIMINACIÓN (GET)
If action = "delete" Then
    If songId <> "" And IsNumeric(songId) Then
        ' Eliminar relaciones foráneas
        conn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = " & CLng(songId)
        conn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & CLng(songId)
        conn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & CLng(songId)
        ' Eliminar la canción
        conn.Execute "DELETE FROM TBL_SONG WHERE ID_SONG = " & CLng(songId)
        
        msgSuccess = "¡Canción eliminada satisfactoriamente!"
        action = ""
    End If
End If

' Cargar datos para edición si aplica
Dim editSongName, editSongPath, editSongPlays, editArtistId, editAlbumId
If action = "edit" Then
    If songId <> "" And IsNumeric(songId) Then
        Dim rsEdit
        Set rsEdit = conn.Execute("SELECT SONG_NAME, SONG_PATH, ISNULL(PLAYS, 0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = " & CLng(songId))
        If Not rsEdit.EOF Then
            editSongName = rsEdit("SONG_NAME")
            editSongPath = rsEdit("SONG_PATH")
            editSongPlays = rsEdit("PLAYS")
        End If
        CloseRecordset rsEdit

        ' Obtener artista actual
        Dim rsArtCurr
        Set rsArtCurr = conn.Execute("SELECT TOP 1 ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = " & CLng(songId))
        If Not rsArtCurr.EOF Then editArtistId = rsArtCurr("ID_ARTIST")
        CloseRecordset rsArtCurr

        ' Obtener álbum actual
        Dim rsAlbCurr
        Set rsAlbCurr = conn.Execute("SELECT TOP 1 ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = " & CLng(songId))
        If Not rsAlbCurr.EOF Then editAlbumId = rsAlbCurr("ID_ALBUM")
        CloseRecordset rsAlbCurr
    End If
End If

' Listas desplegables de Artistas y Álbumes
Dim rsAllArtists, rsAllAlbums
Set rsAllArtists = conn.Execute("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME ASC")
Set rsAllAlbums = conn.Execute("SELECT ID_ALBUM, TITLE_ALBUM FROM TBL_ALBUM ORDER BY TITLE_ALBUM ASC")

' Lista de canciones para la tabla principal
Dim rsSongsList, searchQ
searchQ = Trim(Request("q"))

Dim sqlList
sqlList = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
          "(SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON t.ID_ALBUM = al.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG) AS TITLE_ALBUM, " & _
          "(SELECT TOP 1 a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, '') " & _
          " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON p.ID_ARTIST = a.ID_ARTIST " & _
          " WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_FULL_NAME " & _
          "FROM TBL_SONG s "

If searchQ <> "" Then
    sqlList = sqlList & " WHERE s.SONG_NAME LIKE '%" & SafeSql(searchQ) & "%' "
End If
sqlList = sqlList & " ORDER BY s.ID_SONG DESC"

Set rsSongsList = Server.CreateObject("ADODB.Recordset")
rsSongsList.Open sqlList, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - CRUD de Canciones (Admin)</title>
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
            <h1 class="section-title">Gestión de Canciones (CRUD TBL_SONG)</h1>
          </div>
          <% If action <> "new" Then %>
          <a href="admin_songs.asp?action=new" class="btn btn-primary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Agregar Nueva Canción</span>
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
            <% If action = "new" Then Response.Write "Nueva Canción" Else Response.Write "Editar Canción #" & songId %>
          </h2>

          <form method="POST" action="admin_songs.asp?action=<%=IIf(action = "new", "save_new", "save_edit")%>">
            <% If action = "edit" Then %>
            <input type="hidden" name="id" value="<%=songId%>">
            <% End If %>

            <div class="form-grid">
              <div class="form-group">
                <label for="song_name">Nombre de la Canción *</label>
                <input type="text" id="song_name" name="song_name" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", songName, editSongName))%>" 
                       placeholder="Ej. Adventure" required>
              </div>

              <div class="form-group">
                <label for="song_path">Ruta del Archivo de Audio (MP3) *</label>
                <input type="text" id="song_path" name="song_path" class="form-control" 
                       value="<%=Server.HTMLEncode(IIf(action = "new", IIf(songPath <> "", songPath, "~/songFiles/bensound-adventure.mp3"), editSongPath))%>" 
                       placeholder="~/songFiles/bensound-adventure.mp3" required>
                <small style="color:var(--text-dim);font-size:11px;">Rutas disponibles: ~/songFiles/bensound-adventure.mp3, bensound-dreams.mp3, etc.</small>
              </div>

              <div class="form-group">
                <label for="artist_id">Artista (TBL_PLAYER)</label>
                <select id="artist_id" name="artist_id" class="form-control">
                  <option value="">-- Sin Artista Asociado --</option>
                  <%
                  Dim currentArtVal
                  currentArtVal = IIf(action = "new", artistId, editArtistId)
                  If Not rsAllArtists.EOF Then
                    rsAllArtists.MoveFirst
                    Do While Not rsAllArtists.EOF
                      Dim aId, aFull
                      aId = rsAllArtists("ID_ARTIST")
                      aFull = FormatArtistFullName(rsAllArtists("ARTIST_NAME"), rsAllArtists("ARTIST_LASTNAME"))
                  %>
                  <option value="<%=aId%>" <% If CStr(aId) = CStr(currentArtVal) Then Response.Write "selected" %>>
                    <%=Server.HTMLEncode(aFull)%>
                  </option>
                  <%
                      rsAllArtists.MoveNext
                    Loop
                  End If
                  %>
                </select>
              </div>

              <div class="form-group">
                <label for="album_id">Álbum (TBL_TRACKLIST)</label>
                <select id="album_id" name="album_id" class="form-control">
                  <option value="">-- Sin Álbum (Sencillo) --</option>
                  <%
                  Dim currentAlbVal
                  currentAlbVal = IIf(action = "new", albumId, editAlbumId)
                  If Not rsAllAlbums.EOF Then
                    rsAllAlbums.MoveFirst
                    Do While Not rsAllAlbums.EOF
                      Dim alId, alTitle
                      alId = rsAllAlbums("ID_ALBUM")
                      alTitle = rsAllAlbums("TITLE_ALBUM")
                  %>
                  <option value="<%=alId%>" <% If CStr(alId) = CStr(currentAlbVal) Then Response.Write "selected" %>>
                    <%=Server.HTMLEncode(alTitle)%>
                  </option>
                  <%
                      rsAllAlbums.MoveNext
                    Loop
                  End If
                  %>
                </select>
              </div>

              <div class="form-group">
                <label for="plays">Reproducciones Iniciales</label>
                <input type="number" id="plays" name="plays" class="form-control" 
                       value="<%=IIf(action = "new", "0", editSongPlays)%>" min="0">
              </div>
            </div>

            <div style="display:flex; gap:12px;">
              <button type="submit" class="btn btn-primary">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
                <span>Guardar Canción</span>
              </button>
              <a href="admin_songs.asp" class="btn btn-secondary">Cancelar</a>
            </div>
          </form>
        </div>
        <% End If %>

        <!-- BUSCADOR RÁPIDO -->
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
          <form method="GET" action="admin_songs.asp" style="display:flex; gap:10px;">
            <input type="text" name="q" value="<%=Server.HTMLEncode(searchQ)%>" placeholder="Buscar por título..." class="form-control" style="width:240px; padding:6px 12px; font-size:13px;">
            <button type="submit" class="btn btn-secondary btn-sm">Buscar</button>
            <% If searchQ <> "" Then %>
              <a href="admin_songs.asp" class="btn btn-sm" style="color:var(--text-dim);">Limpiar</a>
            <% End If %>
          </form>
        </div>

        <!-- TABLA DE LISTADO PARA EL CRUD -->
        <div class="song-table-container">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width:50px;">ID</th>
                <th>Nombre de la Canción</th>
                <th>Artista</th>
                <th>Álbum</th>
                <th>Reproducciones</th>
                <th>Ruta MP3</th>
                <th style="text-align:right;">Acciones</th>
              </tr>
            </thead>
            <tbody>
              <%
              If rsSongsList.EOF Then
              %>
              <tr>
                <td colspan="7" style="text-align:center;padding:30px;color:var(--text-muted);">
                  No hay canciones registradas.
                </td>
              </tr>
              <%
              Else
                Do While Not rsSongsList.EOF
                  Dim rowId, rowName, rowPath, rowPlays, rowArtist, rowAlbum
                  rowId = rsSongsList("ID_SONG")
                  rowName = rsSongsList("SONG_NAME")
                  rowPath = rsSongsList("SONG_PATH")
                  rowPlays = rsSongsList("PLAYS")
                  rowArtist = rsSongsList("ARTIST_FULL_NAME")
                  If IsNull(rowArtist) Or Trim(rowArtist) = "" Then rowArtist = "Sin asignar"
                  rowAlbum = rsSongsList("TITLE_ALBUM")
                  If IsNull(rowAlbum) Or Trim(rowAlbum) = "" Then rowAlbum = "Sin álbum"
              %>
              <tr>
                <td style="color:var(--text-dim);font-weight:600;"><%=rowId%></td>
                <td style="font-weight:600;color:#ffffff;"><%=Server.HTMLEncode(rowName)%></td>
                <td style="color:var(--primary-light);"><%=Server.HTMLEncode(rowArtist)%></td>
                <td style="color:var(--text-muted);"><%=Server.HTMLEncode(rowAlbum)%></td>
                <td style="color:var(--text-dim);"><%=FormatNumber(rowPlays, 0)%></td>
                <td style="color:var(--text-dim);font-size:12px;font-family:monospace;"><%=Server.HTMLEncode(rowPath)%></td>
                <td style="text-align:right;white-space:nowrap;">
                  <a href="admin_songs.asp?action=edit&id=<%=rowId%>" class="btn btn-secondary btn-sm" style="padding:4px 10px;margin-right:6px;" title="Editar">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                    <span>Editar</span>
                  </a>
                  <a href="admin_songs.asp?action=delete&id=<%=rowId%>" 
                     class="btn btn-danger btn-sm" 
                     style="padding:4px 10px;"
                     onclick="return confirm('¿Está seguro de eliminar permanentemente la canción <%=Server.HTMLEncode(rowName)%>?');"
                     title="Eliminar">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                    <span>Eliminar</span>
                  </a>
                </td>
              </tr>
              <%
                  rsSongsList.MoveNext
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
CloseRecordset rsSongsList
CloseRecordset rsAllArtists
CloseRecordset rsAllAlbums
CloseDBConnection conn
%>
