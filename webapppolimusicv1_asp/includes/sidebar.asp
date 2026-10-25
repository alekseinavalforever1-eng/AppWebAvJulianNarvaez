<%
' ============================================================================
' PoliMusic - Left Sidebar Navigation
' ============================================================================
Dim currentScript
currentScript = LCase(Request.ServerVariables("SCRIPT_NAME"))
%>
<aside class="sidebar">
  <div class="brand-logo">
    <div class="brand-icon">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
        <path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/>
      </svg>
    </div>
    <div class="brand-name">PoliMusic</div>
  </div>

  <div class="nav-section">
    <div class="nav-section-title">Navegación</div>
    
    <a href="index.asp" class="nav-item <% If InStr(currentScript, "index.asp") > 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg>
      <span>Inicio (Home)</span>
    </a>
    
    <a href="songs.asp" class="nav-item <% If InStr(currentScript, "songs.asp") > 0 And InStr(currentScript, "admin") = 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/></svg>
      <span>Canciones (Song List)</span>
    </a>
    
    <a href="albums.asp" class="nav-item <% If InStr(currentScript, "album") > 0 And InStr(currentScript, "admin") = 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 14.5c-2.49 0-4.5-2.01-4.5-4.5S9.51 7.5 12 7.5s4.5 2.01 4.5 4.5-2.01 4.5-4.5 4.5zm0-5.5c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>
      <span>Álbumes</span>
    </a>
    
    <a href="artists.asp" class="nav-item <% If InStr(currentScript, "artist") > 0 And InStr(currentScript, "admin") = 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>
      <span>Artistas</span>
    </a>
    
    <a href="playlists.asp" class="nav-item <% If InStr(currentScript, "playlist") > 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M4 10h12v2H4zm0-4h12v2H4zm0 8h8v2H4zm10 0v6l5-3z"/></svg>
      <span>Playlists</span>
    </a>
  </div>

  <% If IsAdmin() Then %>
  <div class="nav-section">
    <div class="admin-badge-menu">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.11v8.8z"/></svg>
      <span>MÓDULO ADMIN (CRUDs)</span>
    </div>

    <a href="admin_songs.asp" class="nav-item <% If InStr(currentScript, "admin_songs.asp") > 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
      <span>CRUD Canciones</span>
    </a>

    <a href="admin_artists.asp" class="nav-item <% If InStr(currentScript, "admin_artists.asp") > 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
      <span>CRUD Artistas</span>
    </a>

    <a href="admin_albums.asp" class="nav-item <% If InStr(currentScript, "admin_albums.asp") > 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
      <span>CRUD Álbumes</span>
    </a>

    <a href="env_vars.asp" class="nav-item <% If InStr(currentScript, "env_vars.asp") > 0 Then Response.Write "active" %>">
      <svg viewBox="0 0 24 24"><path d="M9.4 16.6 4.8 12l4.6-4.6L8 6l-6 6 6 6 1.4-1.4zm5.2 0 4.6-4.6-4.6-4.6L16 6l6 6-6 6-1.4-1.4z"/></svg>
      <span>Variables de Entorno</span>
    </a>
  </div>
  <% End If %>

  <div style="margin-top:auto; padding-top:16px; border-top:1px solid var(--border-color);">
    <a href="logout.asp" class="nav-item" style="color:#ef9a9a;">
      <svg viewBox="0 0 24 24"><path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4V5z"/></svg>
      <span>Cerrar Sesión</span>
    </a>
    
    <div style="padding: 12px 16px; margin-top: 10px; background: rgba(187, 134, 252, 0.08); border-radius: 8px; border: 1px solid rgba(187, 134, 252, 0.2);">
      <div style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--primary-pastel); font-weight: 700;">Estudiante</div>
      <div style="font-size: 13px; font-weight: 600; color: #fff;">Julián Narváez</div>
      <div style="font-size: 11px; color: var(--text-dim);">EPN &bull; Web Avanzadas</div>
    </div>
  </div>
</aside>
