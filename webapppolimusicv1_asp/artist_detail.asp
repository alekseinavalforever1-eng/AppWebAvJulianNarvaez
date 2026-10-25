<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
RequireAuth()

Dim artistId, conn, rsArtist, rsSongs
artistId = Request.QueryString("id")

If artistId = "" Or Not IsNumeric(artistId) Then
    Response.Redirect "artists.asp"
    Response.End
End If

Set conn = GetDBConnection()

' Obtener datos del Artista
Set rsArtist = Server.CreateObject("ADODB.Recordset")
rsArtist.Open "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = " & CLng(artistId), conn

If rsArtist.EOF Then
    CloseRecordset rsArtist
    CloseDBConnection conn
    Response.Redirect "artists.asp"
    Response.End
End If

Dim artName, artLastName, artFull, artImg
artName = rsArtist("ARTIST_NAME")
artLastName = rsArtist("ARTIST_LASTNAME")
artFull = FormatArtistFullName(artName, artLastName)
artImg = ResolveMediaUrl(rsArtist("ARTIST_IMAGE"), "ArtistImages/artist1.jpg")

' Obtener canciones asociadas al Artista
Dim sqlSongs
sqlSongs = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
           "al.TITLE_ALBUM, al.ALBUM_IMAGE_PATH " & _
           "FROM TBL_PLAYER p " & _
           "INNER JOIN TBL_SONG s ON p.ID_SONG = s.ID_SONG " & _
           "LEFT JOIN TBL_TRACKLIST t ON s.ID_SONG = t.ID_SONG " & _
           "LEFT JOIN TBL_ALBUM al ON t.ID_ALBUM = al.ID_ALBUM " & _
           "WHERE p.ID_ARTIST = " & CLng(artistId) & " " & _
           "ORDER BY s.ID_SONG ASC"

Set rsSongs = Server.CreateObject("ADODB.Recordset")
rsSongs.Open sqlSongs, conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - <%=Server.HTMLEncode(artFull)%></title>
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
        <a href="artists.asp" class="btn btn-secondary btn-sm" style="margin-bottom:20px;">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>
          <span>Volver a Artistas</span>
        </a>

        <!-- Hero del Artista -->
        <div class="detail-hero">
          <img src="<%=artImg%>" alt="<%=Server.HTMLEncode(artFull)%>" class="detail-cover artist-circle" onerror="this.src='ArtistImages/artist1.jpg'">
          <div class="detail-meta">
            <span class="detail-type">Artista Verificado</span>
            <h1 class="detail-title"><%=Server.HTMLEncode(artFull)%></h1>
            <div class="detail-info">
              <span>Música y Discografía Oficial</span>
            </div>
          </div>
        </div>

        <div class="song-list-heading-box">
          <div style="display:flex;align-items:center;gap:12px;">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="var(--primary-pastel)"><path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/></svg>
            <h2>Canciones de <%=Server.HTMLEncode(artFull)%></h2>
          </div>
        </div>

        <!-- Tabla de Canciones del Artista -->
        <div class="song-table-container">
          <table class="song-table">
            <thead>
              <tr>
                <th class="col-num">#</th>
                <th>Título</th>
                <th>Álbum</th>
                <th>Reproducciones</th>
                <th class="col-actions">Reproducir</th>
              </tr>
            </thead>
            <tbody>
              <%
              Dim sNum
              sNum = 1

              If rsSongs.EOF Then
              %>
              <tr>
                <td colspan="5" style="text-align:center;padding:30px;color:var(--text-muted);">
                  No hay canciones registradas para este artista.
                </td>
              </tr>
              <%
              Else
                Do While Not rsSongs.EOF
                  Dim sId, sName, sPath, sPlays, sAlbum, sCover
                  sId = rsSongs("ID_SONG")
                  sName = rsSongs("SONG_NAME")
                  sPath = ResolveMediaUrl(rsSongs("SONG_PATH"), "songFiles/bensound-adventure.mp3")
                  sPlays = rsSongs("PLAYS")
                  sAlbum = rsSongs("TITLE_ALBUM")
                  If IsNull(sAlbum) Or Trim(sAlbum) = "" Then sAlbum = "Sencillo"
                  sCover = ResolveMediaUrl(rsSongs("ALBUM_IMAGE_PATH"), "albumImages/album1.png")
              %>
              <tr class="song-row"
                  data-song-id="<%=sId%>"
                  data-song-title="<%=Server.HTMLEncode(sName)%>"
                  data-song-artist="<%=Server.HTMLEncode(artFull)%>"
                  data-song-path="<%=sPath%>"
                  data-song-cover="<%=sCover%>">
                <td class="col-num"><%=sNum%></td>
                <td>
                  <div class="song-info-cell">
                    <img src="<%=sCover%>" alt="<%=Server.HTMLEncode(sName)%>" class="song-thumb" onerror="this.src='albumImages/album1.png'">
                    <div class="song-details">
                      <span class="song-name"><%=Server.HTMLEncode(sName)%></span>
                      <span class="song-artist-link"><%=Server.HTMLEncode(artFull)%></span>
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
                  sNum = sNum + 1
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
CloseRecordset rsArtist
CloseDBConnection conn
%>
