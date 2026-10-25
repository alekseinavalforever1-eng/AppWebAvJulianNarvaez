# PoliMusic - Clon de Spotify en ASP Clásico (VBScript)

Sitio web estilo Spotify desarrollado en **ASP Clásico con VBScript (.asp)** y conectado a **SQL Server 2025**, diseñado con una estética original en **color morado pastel**, modo oscuro con glassmorphism y reproductor de música interactivo.

---

## 🎵 Características Principales

1. **Arquitectura ASP Clásico Puro**:
   - Todo el backend está desarrollado en páginas con extensión `.asp` y lenguaje `VBScript`.
   - Conexión nativa con ADODB a SQL Server 2025.
   - Sin dependencias de ASP.NET ni C#.

2. **Esquema de Base de Datos Estricto**:
   - Se apega 100% a las tablas de `BDD_PoliMusic.sql` sin crear tablas adicionales:
     - `TBL_USER`
     - `TBL_SONG`
     - `TBL_ARTIST`
     - `TBL_ALBUM`
     - `TBL_TRACKLIST`
     - `TBL_PLAYER`
     - `TBL_PLAYLIST`
     - `TBL_PLAYLIST_SONG`
     - `TBL_CATALOG`

3. **Control de Acceso y Roles**:
   - **Rol Normal (`USER_TYPE = 0`)**: Puede navegar por el catálogo, explorar álbumes, artistas, playlists y reproducir música.
   - **Rol Administrador (`USER_TYPE = 1`)**: Tiene acceso exclusivo al módulo administrativo con **CRUDS completos** para:
     - **Canciones** (`admin_songs.asp`): Crear, listar, editar y eliminar canciones asociando artista y álbum.
     - **Artistas** (`admin_artists.asp`): Crear, listar, editar y eliminar artistas con soporte para imágenes.
     - **Álbumes** (`admin_albums.asp`): Crear, listar, editar y eliminar álbumes con año y portada.

4. **Autenticación y Sesión**:
   - Pantalla de inicio de sesión (`login.asp`) estilizada en morado pastel con botones de autocompletado rápido para pruebas.
   - Almacenamiento de variables de sesión (`Session("ID_USER")`, `Session("USERNAME")`, `Session("USER_TYPE")`, `Session("USER_PHOTO")`).
   - Mensaje de bienvenida personalizado con el nombre del usuario logueado leído desde la sesión.
   - Opción de cierre de sesión (`logout.asp`) que destruye la sesión y redirige al login.

5. **Listado de Canciones (Song List) y Reproductor**:
   - Encabezado explícito **"Listado de Canciones (Song List)"**.
   - Tabla estilizada con portadas, títulos, artistas, álbumes y contador de reproducciones.
   - Reproductor flotante persistente al pie de página (`js/player.js`) con controles:
     - Play / Pause
     - Canción anterior / siguiente
     - Barra de progreso interactiva (seeker) y tiempo actual / duración total
     - Control de volumen y botón de silenciar (mute)
     - Repetir pista (loop)
     - Animación de ecualizador
     - Incremento automático de reproducciones en base de datos mediante llamada AJAX a `api_plays.asp`.

---

## 👥 Usuarios de Prueba (de la Base de Datos)

| Rol | Usuario | Contraseña | Permisos |
|---|---|---|---|
| **Admin** | `userAdmin` | `12345678` | Acceso total + Módulo CRUD (Canciones, Artistas, Álbumes) |
| **Normal** | `userNormal` | `12345678` | Reproducción y navegación general |
| **Normal** | `username1` | `7C222FB2927D828AF22F592134E8932480637C0D` | Reproducción y navegación general |

---

## 🚀 Cómo Ejecutar la Aplicación

### Opción 1: Con IIS Express (Recomendado y ya configurado)
1. Abrir PowerShell en la carpeta del proyecto y ejecutar:
   ```powershell
   & 'C:\Program Files\IIS Express\iisexpress.exe' /path:$pwd /port:8085
   ```
2. Abrir el navegador en:
   ```
   http://localhost:8085/login.asp
   ```

### Opción 2: Con IIS (Internet Information Services)
1. Asegurarse de tener habilitada la característica **ASP** en *Activar o desactivar las características de Windows -> Internet Information Services -> World Wide Web Services -> Características de desarrollo de aplicaciones -> ASP*.
2. Crear un nuevo Sitio Web en el Administrador de IIS apuntando a la carpeta del proyecto.
3. Asignar un grupo de aplicaciones en modo **Classic** o **Integrated** con identidad con permisos de acceso a SQL Server.
4. Navegar a `http://localhost/login.asp`.

---

## 📂 Estructura de Archivos

- `default.asp`: Redirecciona automáticamente según el estado de la sesión.
- `login.asp`: Pantalla de inicio de sesión con validación de credenciales.
- `logout.asp`: Cierre de sesión y limpieza de variables.
- `index.asp`: Página principal con saludo de bienvenida, Song List, álbumes y artistas.
- `songs.asp`: Catálogo completo de canciones con buscador y ordenamiento.
- `albums.asp`: Vista en cuadrícula de álbumes.
- `album_detail.asp`: Detalle de un álbum con su lista de canciones.
- `artists.asp`: Vista en cuadrícula de artistas.
- `artist_detail.asp`: Detalle de un artista y sus canciones.
- `playlists.asp`: Vista de playlists comunitarias.
- `playlist_detail.asp`: Detalle y canciones de una playlist.
- `admin_songs.asp`: CRUD completo de canciones (Admin).
- `admin_artists.asp`: CRUD completo de artistas (Admin).
- `admin_albums.asp`: CRUD completo de álbumes (Admin).
- `api_plays.asp`: Endpoint para incrementar el contador de reproducciones en tiempo real.
- `includes/`:
  - `db.asp`: Conexión centralizada ADODB a SQL Server 2025 con fallback automático.
  - `auth.asp`: Validación de sesión y protección de rutas según rol.
  - `sidebar.asp`: Menú de navegación lateral con enlace dinámico al módulo de administración.
  - `topbar.asp`: Barra superior con saludo, búsqueda y botón de salir.
  - `player_bar.asp`: Barra del reproductor interactivo.
- `css/style.css`: Hoja de estilos con paleta morado pastel y estética Spotify.
- `js/player.js`: Motor de reproducción de audio HTML5 y sincronización de cola.
- `songFiles/`: Archivos de audio MP3.
- `albumImages/`, `ArtistImages/`, `playlistImages/`, `userImages/`: Recursos visuales.
