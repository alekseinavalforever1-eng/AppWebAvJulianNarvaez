<%
' ============================================================================
' PoliMusic - Top Navigation Bar
' Displays welcome greeting with Session("USERNAME"), search bar, and logout
' ============================================================================
Dim uPhoto, uType, uTypeName
uPhoto = GetCurrentUserPhoto()
uType = CInt("0" & Session("USER_TYPE"))

If uType = 1 Then
    uTypeName = "Administrador"
Else
    uTypeName = "Usuario"
End If
%>
<header class="topbar">
  <div class="topbar-left">
    <div class="search-box">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="var(--text-dim)">
        <path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/>
      </svg>
      <input type="text" id="topSearchInput" placeholder="Buscar canciones, artistas o álbumes...">
    </div>
  </div>

  <div class="topbar-right">
    <div class="user-welcome-chip">
      <img src="<%=uPhoto%>" alt="<%=Server.HTMLEncode(GetCurrentUsername())%>" class="user-avatar" onerror="this.src='userImages/default.jpg'">
      <div class="user-meta">
        <span class="user-greeting">¡Bienvenido, <%=Server.HTMLEncode(GetCurrentUsername())%>!</span>
        <div style="display:flex;align-items:center;gap:6px;">
          <span class="user-display-name"><%=Server.HTMLEncode(GetCurrentUsername())%></span>
          <span class="badge-role <%=IIf(uType = 1, "badge-admin", "badge-normal")%>"><%=uTypeName%></span>
        </div>
      </div>
    </div>

    <a href="logout.asp" class="btn btn-secondary btn-sm" title="Cerrar Sesión">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
        <path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4V5z"/>
      </svg>
      <span>Salir</span>
    </a>
  </div>
</header>
<%
Function IIf(cond, trueVal, falseVal)
    If cond Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function
%>
