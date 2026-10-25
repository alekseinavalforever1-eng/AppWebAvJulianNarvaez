<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireAuth()

Dim conn, rsAlbums
Set conn = GetDBConnection()

Dim sql
sql = "SELECT a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
      "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS TRACK_COUNT " & _
      "FROM TBL_ALBUM a ORDER BY a.TITLE_ALBUM ASC"

Set rsAlbums = Server.CreateObject("ADODB.Recordset")
rsAlbums.Open sql, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Álbumes</title>
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
            <span class="section-subtitle">Discografía</span>
            <h1 class="section-title">Colección de Álbumes</h1>
          </div>
          <% If IsAdmin() Then %>
          <a href="admin_albums.asp?action=new" class="btn btn-primary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Nuevo Álbum</span>
          </a>
          <% End If %>
        </div>

        <div class="media-grid">
          <%
          If rsAlbums.EOF Then
          %>
          <div style="grid-column: 1/-1; text-align:center; padding:50px; color:var(--text-muted);">
            No hay álbumes disponibles en la base de datos.
          </div>
          <%
          Else
            Do While Not rsAlbums.EOF
              Dim albId, albTitle, albYear, albImg, trackCount
              albId = rsAlbums("ID_ALBUM")
              albTitle = rsAlbums("TITLE_ALBUM")
              albYear = rsAlbums("RELEASE_YEAR")
              albImg = ResolveMediaUrl(rsAlbums("ALBUM_IMAGE_PATH"), "albumImages/album1.png")
              trackCount = rsAlbums("TRACK_COUNT")
          %>
          <div class="media-card" onclick="window.location.href='album_detail.asp?id=<%=albId%>'">
            <div class="card-image-wrap">
              <img src="<%=albImg%>" alt="<%=Server.HTMLEncode(albTitle)%>" class="card-image" onerror="this.src='albumImages/album1.png'">
              <div class="card-play-overlay">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              </div>
            </div>
            <div class="card-title"><%=Server.HTMLEncode(albTitle)%></div>
            <div class="card-subtitle"><%=albYear%> &bull; <%=trackCount%> canciones</div>
          </div>
          <%
              rsAlbums.MoveNext
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
CloseRecordset rsAlbums
CloseDBConnection conn
%>
