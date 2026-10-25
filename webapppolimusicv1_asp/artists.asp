<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireAuth()

Dim conn, rsArtists
Set conn = GetDBConnection()

Dim sql
sql = "SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
      "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
      "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME ASC"

Set rsArtists = Server.CreateObject("ADODB.Recordset")
rsArtists.Open sql, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Artistas</title>
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
            <span class="section-subtitle">Creadores de Música</span>
            <h1 class="section-title">Artistas</h1>
          </div>
          <% If IsAdmin() Then %>
          <a href="admin_artists.asp?action=new" class="btn btn-primary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Nuevo Artista</span>
          </a>
          <% End If %>
        </div>

        <div class="media-grid">
          <%
          If rsArtists.EOF Then
          %>
          <div style="grid-column: 1/-1; text-align:center; padding:50px; color:var(--text-muted);">
            No hay artistas registrados en la base de datos.
          </div>
          <%
          Else
            Do While Not rsArtists.EOF
              Dim artId, artName, artLastName, artFull, artImg, songCount
              artId = rsArtists("ID_ARTIST")
              artName = rsArtists("ARTIST_NAME")
              artLastName = rsArtists("ARTIST_LASTNAME")
              artFull = FormatArtistFullName(artName, artLastName)
              artImg = ResolveMediaUrl(rsArtists("ARTIST_IMAGE"), "ArtistImages/artist1.jpg")
              songCount = rsArtists("SONG_COUNT")
          %>
          <div class="media-card artist-card" onclick="window.location.href='artist_detail.asp?id=<%=artId%>'">
            <div class="card-image-wrap">
              <img src="<%=artImg%>" alt="<%=Server.HTMLEncode(artFull)%>" class="card-image" onerror="this.src='ArtistImages/artist1.jpg'">
              <div class="card-play-overlay">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              </div>
            </div>
            <div class="card-title" style="text-align:center;"><%=Server.HTMLEncode(artFull)%></div>
            <div class="card-subtitle" style="text-align:center;"><%=songCount%> canciones</div>
          </div>
          <%
              rsArtists.MoveNext
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
CloseRecordset rsArtists
CloseDBConnection conn
%>
