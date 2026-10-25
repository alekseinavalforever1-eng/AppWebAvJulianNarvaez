<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireAuth()

Dim conn, rsSongs, filterQuery, sortParam
filterQuery = Trim(Request.QueryString("q"))
sortParam = Trim(Request.QueryString("sort"))

Set conn = GetDBConnection()

Dim sql
sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
      "(SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON t.ID_ALBUM = al.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG) AS TITLE_ALBUM, " & _
      "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON t.ID_ALBUM = al.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE_PATH, " & _
      "(SELECT TOP 1 a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, '') " & _
      " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON p.ID_ARTIST = a.ID_ARTIST " & _
      " WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_FULL_NAME " & _
      "FROM TBL_SONG s "

If filterQuery <> "" Then
    sql = sql & " WHERE s.SONG_NAME LIKE '%" & SafeSql(filterQuery) & "%' "
End If

If sortParam = "plays" Then
    sql = sql & " ORDER BY s.PLAYS DESC, s.SONG_NAME ASC"
ElseIf sortParam = "title" Then
    sql = sql & " ORDER BY s.SONG_NAME ASC"
Else
    sql = sql & " ORDER BY s.ID_SONG ASC"
End If

Set rsSongs = Server.CreateObject("ADODB.Recordset")
rsSongs.Open sql, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Listado de Canciones (Song List)</title>
  <link rel="stylesheet" href="css/style.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style> body { font-family: 'Plus Jakarta Sans', sans-serif; } </style>
</head>
<body>
  <div class="app-container">
    <!--#include file="includes/sidebar.asp"-->

    <div class="main-wrapper">
      <!--#include file="includes/topbar.asp"-->

      <div class="content-body">
        
        <!-- Header de la sección -->
        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-subtitle">Explorador Musical</span>
            <h1 class="section-title">Listado de Canciones (Song List)</h1>
          </div>
          <% If IsAdmin() Then %>
          <a href="admin_songs.asp?action=new" class="btn btn-primary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Nueva Canción</span>
          </a>
          <% End If %>
        </div>

        <!-- Filtros y Ordenamiento -->
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px; flex-wrap:wrap; gap:14px;">
          <form method="GET" action="songs.asp" style="display:flex; gap:10px; align-items:center;">
            <input type="text" name="q" value="<%=Server.HTMLEncode(filterQuery)%>" placeholder="Filtrar por título..." class="form-control" style="width:240px; padding:6px 12px; font-size:13px;">
            <button type="submit" class="btn btn-secondary btn-sm">Buscar</button>
            <% If filterQuery <> "" Then %>
              <a href="songs.asp" class="btn btn-sm" style="color:var(--text-dim);">Limpiar</a>
            <% End If %>
          </form>

          <div style="display:flex; gap:8px; align-items:center;">
            <span style="font-size:12px; color:var(--text-dim);">Ordenar por:</span>
            <a href="songs.asp?sort=default<% If filterQuery <> "" Then Response.Write "&q=" & Server.URLEncode(filterQuery) %>" class="btn btn-sm <% If sortParam = "" Or sortParam = "default" Then Response.Write "btn-primary" Else Response.Write "btn-secondary" %>">ID</a>
            <a href="songs.asp?sort=title<% If filterQuery <> "" Then Response.Write "&q=" & Server.URLEncode(filterQuery) %>" class="btn btn-sm <% If sortParam = "title" Then Response.Write "btn-primary" Else Response.Write "btn-secondary" %>">Título</a>
            <a href="songs.asp?sort=plays<% If filterQuery <> "" Then Response.Write "&q=" & Server.URLEncode(filterQuery) %>" class="btn btn-sm <% If sortParam = "plays" Then Response.Write "btn-primary" Else Response.Write "btn-secondary" %>">Más Reproducidas</a>
          </div>
        </div>

        <!-- Tabla Estilizada de Canciones -->
        <div class="song-table-container">
          <table class="song-table">
            <thead>
              <tr>
                <th class="col-num">#</th>
                <th>Título / Artista</th>
                <th>Álbum</th>
                <th>Reproducciones</th>
                <th class="col-actions">Reproducir</th>
              </tr>
            </thead>
            <tbody>
              <%
              Dim songCounter
              songCounter = 1
              
              If rsSongs.EOF Then
              %>
              <tr>
                <td colspan="5" style="text-align:center;padding:40px;color:var(--text-muted);">
                  No se encontraron canciones que coincidan con la búsqueda.
                </td>
              </tr>
              <%
              Else
                Do While Not rsSongs.EOF
                  Dim sId, sName, sPath, sPlays, sAlbum, sArtist, sCover
                  sId = rsSongs("ID_SONG")
                  sName = rsSongs("SONG_NAME")
                  sPath = ResolveMediaUrl(rsSongs("SONG_PATH"), "songFiles/bensound-adventure.mp3")
                  sPlays = rsSongs("PLAYS")
                  
                  sAlbum = rsSongs("TITLE_ALBUM")
                  If IsNull(sAlbum) Or Trim(sAlbum) = "" Then sAlbum = "Sencillo"
                  
                  sArtist = rsSongs("ARTIST_FULL_NAME")
                  If IsNull(sArtist) Or Trim(sArtist) = "" Then sArtist = "Varios Artistas"
                  
                  sCover = ResolveMediaUrl(rsSongs("ALBUM_IMAGE_PATH"), "albumImages/album1.png")
              %>
              <tr class="song-row" 
                  data-song-id="<%=sId%>"
                  data-song-title="<%=Server.HTMLEncode(sName)%>"
                  data-song-artist="<%=Server.HTMLEncode(sArtist)%>"
                  data-song-path="<%=sPath%>"
                  data-song-cover="<%=sCover%>">
                <td class="col-num"><%=songCounter%></td>
                <td>
                  <div class="song-info-cell">
                    <img src="<%=sCover%>" alt="<%=Server.HTMLEncode(sName)%>" class="song-thumb" onerror="this.src='albumImages/album1.png'">
                    <div class="song-details">
                      <span class="song-name"><%=Server.HTMLEncode(sName)%></span>
                      <span class="song-artist-link"><%=Server.HTMLEncode(sArtist)%></span>
                    </div>
                  </div>
                </td>
                <td class="col-album"><%=Server.HTMLEncode(sAlbum)%></td>
                <td class="col-plays"><%=FormatNumber(sPlays, 0)%></td>
                <td class="col-actions">
                  <button type="button" class="row-play-btn" title="Reproducir <%=Server.HTMLEncode(sName)%>">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" class="row-play-icon">
                      <path d="M8 5v14l11-7z"/>
                    </svg>
                  </button>
                </td>
              </tr>
              <%
                  songCounter = songCounter + 1
                  rsSongs.MoveNext
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
CloseRecordset rsSongs
CloseDBConnection conn
%>
