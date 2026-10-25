<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireAuth()

Dim albumId, conn, rsAlbum, rsTracks
albumId = Request.QueryString("id")

If albumId = "" Or Not IsNumeric(albumId) Then
    Response.Redirect "albums.asp"
    Response.End
End If

Set conn = GetDBConnection()

' Obtener datos del Álbum
Set rsAlbum = Server.CreateObject("ADODB.Recordset")
rsAlbum.Open "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = " & CLng(albumId), conn

If rsAlbum.EOF Then
    CloseRecordset rsAlbum
    CloseDBConnection conn
    Response.Redirect "albums.asp"
    Response.End
End If

Dim albTitle, albYear, albImg
albTitle = rsAlbum("TITLE_ALBUM")
albYear = rsAlbum("RELEASE_YEAR")
albImg = ResolveMediaUrl(rsAlbum("ALBUM_IMAGE_PATH"), "albumImages/album1.png")

' Obtener canciones del Álbum
Dim sqlTracks
sqlTracks = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
            "(SELECT TOP 1 a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, '') " & _
            " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON p.ID_ARTIST = a.ID_ARTIST " & _
            " WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_FULL_NAME " & _
            "FROM TBL_TRACKLIST t " & _
            "INNER JOIN TBL_SONG s ON t.ID_SONG = s.ID_SONG " & _
            "WHERE t.ID_ALBUM = " & CLng(albumId) & " " & _
            "ORDER BY s.ID_SONG ASC"

Set rsTracks = Server.CreateObject("ADODB.Recordset")
rsTracks.Open sqlTracks, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - <%=Server.HTMLEncode(albTitle)%></title>
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
        <a href="albums.asp" class="btn btn-secondary btn-sm" style="margin-bottom:20px;">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>
          <span>Volver a Álbumes</span>
        </a>

        <!-- Hero del Álbum -->
        <div class="detail-hero">
          <img src="<%=albImg%>" alt="<%=Server.HTMLEncode(albTitle)%>" class="detail-cover" onerror="this.src='albumImages/album1.png'">
          <div class="detail-meta">
            <span class="detail-type">Álbum Oficial</span>
            <h1 class="detail-title"><%=Server.HTMLEncode(albTitle)%></h1>
            <div class="detail-info">
              <span>Año de Lanzamiento: <strong><%=albYear%></strong></span>
            </div>
          </div>
        </div>

        <div class="song-list-heading-box">
          <div style="display:flex;align-items:center;gap:12px;">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="var(--primary-pastel)"><path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/></svg>
            <h2>Pistas del Álbum</h2>
          </div>
        </div>

        <!-- Lista de Canciones del Álbum -->
        <div class="song-table-container">
          <table class="song-table">
            <thead>
              <tr>
                <th class="col-num">#</th>
                <th>Título / Artista</th>
                <th>Reproducciones</th>
                <th class="col-actions">Reproducir</th>
              </tr>
            </thead>
            <tbody>
              <%
              Dim trackNum
              trackNum = 1

              If rsTracks.EOF Then
              %>
              <tr>
                <td colspan="4" style="text-align:center;padding:30px;color:var(--text-muted);">
                  Este álbum no tiene pistas asociadas actualmente.
                </td>
              </tr>
              <%
              Else
                Do While Not rsTracks.EOF
                  Dim tId, tName, tPath, tPlays, tArtist
                  tId = rsTracks("ID_SONG")
                  tName = rsTracks("SONG_NAME")
                  tPath = ResolveMediaUrl(rsTracks("SONG_PATH"), "songFiles/bensound-adventure.mp3")
                  tPlays = rsTracks("PLAYS")
                  tArtist = rsTracks("ARTIST_FULL_NAME")
                  If IsNull(tArtist) Or Trim(tArtist) = "" Then tArtist = "Varios Artistas"
              %>
              <tr class="song-row"
                  data-song-id="<%=tId%>"
                  data-song-title="<%=Server.HTMLEncode(tName)%>"
                  data-song-artist="<%=Server.HTMLEncode(tArtist)%>"
                  data-song-path="<%=tPath%>"
                  data-song-cover="<%=albImg%>">
                <td class="col-num"><%=trackNum%></td>
                <td>
                  <div class="song-info-cell">
                    <img src="<%=albImg%>" alt="<%=Server.HTMLEncode(tName)%>" class="song-thumb" onerror="this.src='albumImages/album1.png'">
                    <div class="song-details">
                      <span class="song-name"><%=Server.HTMLEncode(tName)%></span>
                      <span class="song-artist-link"><%=Server.HTMLEncode(tArtist)%></span>
                    </div>
                  </div>
                </td>
                <td class="col-plays"><%=FormatNumber(tPlays, 0)%></td>
                <td class="col-actions">
                  <button type="button" class="row-play-btn" title="Reproducir <%=Server.HTMLEncode(tName)%>">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" class="row-play-icon">
                      <path d="M8 5v14l11-7z"/>
                    </svg>
                  </button>
                </td>
              </tr>
              <%
                  trackNum = trackNum + 1
                  rsTracks.MoveNext
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
CloseRecordset rsTracks
CloseRecordset rsAlbum
CloseDBConnection conn
%>
