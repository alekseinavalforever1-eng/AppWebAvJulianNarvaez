<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Verificar autenticación obligatoria
RequireAuth()

Dim conn, rsSongs, rsAlbums, rsArtists, rsStats
Set conn = GetDBConnection()

' Obtener estadísticas rápidas
Dim totalSongs, totalAlbums, totalArtists
totalSongs = 0
totalAlbums = 0
totalArtists = 0

Set rsStats = conn.Execute("SELECT COUNT(*) FROM TBL_SONG")
If Not rsStats.EOF Then totalSongs = rsStats(0)
CloseRecordset rsStats

Set rsStats = conn.Execute("SELECT COUNT(*) FROM TBL_ALBUM")
If Not rsStats.EOF Then totalAlbums = rsStats(0)
CloseRecordset rsStats

Set rsStats = conn.Execute("SELECT COUNT(*) FROM TBL_ARTIST")
If Not rsStats.EOF Then totalArtists = rsStats(0)
CloseRecordset rsStats

' 1. Consulta para el Listado de Canciones (Song List) con Artista y Álbum
Dim sqlSongs
sqlSongs = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
           "(SELECT TOP 1 al.TITLE_ALBUM FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON t.ID_ALBUM = al.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG) AS TITLE_ALBUM, " & _
           "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON t.ID_ALBUM = al.ID_ALBUM WHERE t.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE_PATH, " & _
           "(SELECT TOP 1 a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, '') " & _
           " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON p.ID_ARTIST = a.ID_ARTIST " & _
           " WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_FULL_NAME " & _
           "FROM TBL_SONG s " & _
           "ORDER BY s.ID_SONG ASC"

Set rsSongs = Server.CreateObject("ADODB.Recordset")
rsSongs.Open sqlSongs, conn

' 2. Consulta para Álbumes Destacados
Set rsAlbums = Server.CreateObject("ADODB.Recordset")
rsAlbums.Open "SELECT TOP 5 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY ID_ALBUM ASC", conn

' 3. Consulta para Artistas Destacados
Set rsArtists = Server.CreateObject("ADODB.Recordset")
rsArtists.Open "SELECT TOP 5 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ID_ARTIST ASC", conn
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Inicio &amp; Song List</title>
  <link rel="stylesheet" href="css/style.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Plus Jakarta Sans', sans-serif; }
  </style>
</head>
<body>
  <div class="app-container">
    <!-- Barra Lateral -->
    <!--#include file="includes/sidebar.asp"-->

    <!-- Contenido Principal -->
    <div class="main-wrapper">
      <!-- Barra Superior con Bienvenida y Logout -->
      <!--#include file="includes/topbar.asp"-->

      <div class="content-body">
        
        <!-- Hero Banner con mensaje de bienvenida usando la variable de sesión -->
        <div class="hero-banner">
          <span class="hero-tag">PoliMusic Experience &bull; Color Morado Pastel</span>
          <h1 class="hero-title">
            ¡Bienvenido de vuelta, <span><%=Server.HTMLEncode(Session("USERNAME"))%></span>!
          </h1>
          <p class="hero-sub">
            Disfruta de la mejor música con una interfaz inspirada en Spotify diseñada en ASP Clásico y VBScript con SQL Server 2025.
          </p>
          <div style="display:flex; gap:12px; margin-top:20px;">
            <a href="songs.asp" class="btn btn-primary">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              <span>Explorar Todo el Catálogo</span>
            </a>
            <% If IsAdmin() Then %>
            <a href="admin_songs.asp" class="btn btn-secondary">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
              <span>Administrar Canciones (Admin)</span>
            </a>
            <% End If %>
          </div>
        </div>

        <!-- ====================================================================
             REQUISITO ESPECÍFICO:
             "Debe mostrar un texto que indique que va a presentar el listado de canciones (Song List)"
             "Debe mostrar el listado de canciones que use igualmente estilos css"
             "Debe tener la capacidad de poder reproducir las canciones"
             ==================================================================== -->
        <div class="song-list-heading-box">
          <div style="display:flex;align-items:center;gap:12px;">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="var(--primary-pastel)">
              <path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/>
            </svg>
            <h2>Listado de Canciones (Song List)</h2>
          </div>
          <span class="song-count-badge"><%=totalSongs%> Canciones Disponibles</span>
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
                <td colspan="5" style="text-align:center;padding:30px;color:var(--text-muted);">
                  No hay canciones registradas en la base de datos.
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
                  data-song-path="<%=sCover%>" 
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

        <!-- Sección de Álbumes Destacados -->
        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-subtitle">Colecciones</span>
            <h2 class="section-title">Álbumes Destacados</h2>
          </div>
          <a href="albums.asp" class="btn btn-secondary btn-sm">Ver Todos</a>
        </div>

        <div class="media-grid">
          <%
          If Not rsAlbums.EOF Then
            Do While Not rsAlbums.EOF
              Dim albId, albTitle, albYear, albImg
              albId = rsAlbums("ID_ALBUM")
              albTitle = rsAlbums("TITLE_ALBUM")
              albYear = rsAlbums("RELEASE_YEAR")
              albImg = ResolveMediaUrl(rsAlbums("ALBUM_IMAGE_PATH"), "albumImages/album1.png")
          %>
          <div class="media-card" onclick="window.location.href='album_detail.asp?id=<%=albId%>'">
            <div class="card-image-wrap">
              <img src="<%=albImg%>" alt="<%=Server.HTMLEncode(albTitle)%>" class="card-image" onerror="this.src='albumImages/album1.png'">
              <div class="card-play-overlay">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              </div>
            </div>
            <div class="card-title"><%=Server.HTMLEncode(albTitle)%></div>
            <div class="card-subtitle"><%=albYear%> &bull; Álbum</div>
          </div>
          <%
              rsAlbums.MoveNext
            Loop
          End If
          %>
        </div>

        <!-- Sección de Artistas Destacados -->
        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-subtitle">Músicos &amp; Bandas</span>
            <h2 class="section-title">Artistas Destacados</h2>
          </div>
          <a href="artists.asp" class="btn btn-secondary btn-sm">Ver Todos</a>
        </div>

        <div class="media-grid">
          <%
          If Not rsArtists.EOF Then
            Do While Not rsArtists.EOF
              Dim artId, artName, artLastName, artFull, artImg
              artId = rsArtists("ID_ARTIST")
              artName = rsArtists("ARTIST_NAME")
              artLastName = rsArtists("ARTIST_LASTNAME")
              artFull = FormatArtistFullName(artName, artLastName)
              artImg = ResolveMediaUrl(rsArtists("ARTIST_IMAGE"), "ArtistImages/artist1.jpg")
          %>
          <div class="media-card artist-card" onclick="window.location.href='artist_detail.asp?id=<%=artId%>'">
            <div class="card-image-wrap">
              <img src="<%=artImg%>" alt="<%=Server.HTMLEncode(artFull)%>" class="card-image" onerror="this.src='ArtistImages/artist1.jpg'">
              <div class="card-play-overlay">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              </div>
            </div>
            <div class="card-title" style="text-align:center;"><%=Server.HTMLEncode(artFull)%></div>
            <div class="card-subtitle" style="text-align:center;">Artista Oficial</div>
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

  <!-- Barra de Reproductor Inferior Fijo -->
  <!--#include file="includes/player_bar.asp"-->
</body>
</html>
<%
CloseRecordset rsSongs
CloseRecordset rsAlbums
CloseRecordset rsArtists
CloseDBConnection conn
%>
