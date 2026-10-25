<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireAuth()

Dim conn, rsPlaylists
Set conn = GetDBConnection()

Dim sql
sql = "SELECT p.ID_PLAYLIST, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, " & _
      "u.USERNAME, " & _
      "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONG_COUNT " & _
      "FROM TBL_PLAYLIST p " & _
      "INNER JOIN TBL_USER u ON p.ID_USER = u.ID_USER " & _
      "ORDER BY p.ID_PLAYLIST ASC"

Set rsPlaylists = Server.CreateObject("ADODB.Recordset")
rsPlaylists.Open sql, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Playlists</title>
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
        
        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-subtitle">Listas de Reproducción</span>
            <h1 class="section-title">Playlists de la Comunidad</h1>
          </div>
        </div>

        <div class="media-grid">
          <%
          If rsPlaylists.EOF Then
          %>
          <div style="grid-column: 1/-1; text-align:center; padding:50px; color:var(--text-muted);">
            No hay playlists disponibles.
          </div>
          <%
          Else
            Do While Not rsPlaylists.EOF
              Dim pId, pTitle, pPhoto, pUser, pCount, pType
              pId = rsPlaylists("ID_PLAYLIST")
              pTitle = rsPlaylists("TITLE")
              pPhoto = ResolveMediaUrl(rsPlaylists("PHOTO"), "playlistImages/default.jpg")
              pUser = rsPlaylists("USERNAME")
              pCount = rsPlaylists("SONG_COUNT")
              pType = rsPlaylists("TYPE")
          %>
          <div class="media-card" onclick="window.location.href='playlist_detail.asp?id=<%=pId%>'">
            <div class="card-image-wrap">
              <img src="<%=pPhoto%>" alt="<%=Server.HTMLEncode(pTitle)%>" class="card-image" onerror="this.src='playlistImages/default.jpg'">
              <div class="card-play-overlay">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              </div>
            </div>
            <div class="card-title"><%=Server.HTMLEncode(pTitle)%></div>
            <div class="card-subtitle">De <%=Server.HTMLEncode(pUser)%> &bull; <%=pCount%> canciones</div>
          </div>
          <%
              rsPlaylists.MoveNext
            Loop
          End If
          %>
        </div>

      </div>
    </div>
  </div>

  <!--#include file="includes/player_bar.asp"-->
</body>
</html>
<%
CloseRecordset rsPlaylists
CloseDBConnection conn
%>
