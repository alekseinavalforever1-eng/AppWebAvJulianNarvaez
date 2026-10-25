<!-- ==========================================================================
     PoliMusic - Bottom Persistent Audio Player Bar
     ========================================================================== -->
<footer class="player-bar" id="playerBar">
  <!-- Track Info -->
  <div class="player-left">
    <img src="albumImages/album1.png" alt="Portada" id="playerThumb" class="player-thumb" onerror="this.src='playlistImages/default.jpg'">
    <div class="player-track-info">
      <div class="player-track-title" id="playerTitle">Selecciona una canción</div>
      <div class="player-track-artist" id="playerArtist">PoliMusic Player</div>
    </div>
  </div>

  <!-- Center Controls -->
  <div class="player-center">
    <div class="player-controls">
      <button class="control-btn" id="btnLoop" title="Repetir canción">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
          <path d="M7 7h10v3l4-4-4-4v3H5v6h2V7zm10 10H7v-3l-4 4 4 4v-3h12v-6h-2v4z"/>
        </svg>
      </button>

      <button class="control-btn" id="btnPrev" title="Anterior">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
          <path d="M6 6h2v12H6zm3.5 6l8.5 6V6z"/>
        </svg>
      </button>

      <button class="play-main-btn" id="btnPlayPause" title="Reproducir / Pausar">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor" id="mainPlayIcon">
          <path d="M8 5v14l11-7z"/>
        </svg>
      </button>

      <button class="control-btn" id="btnNext" title="Siguiente">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
          <path d="M6 18l8.5-6L6 6v12zM16 6v12h2V6h-2z"/>
        </svg>
      </button>
    </div>

    <div class="player-progress-wrap">
      <span class="time-label" id="currentTimeLabel">0:00</span>
      <div class="progress-bar-container" id="progressBarContainer">
        <div class="progress-bar-fill" id="progressFill"></div>
      </div>
      <span class="time-label" id="totalTimeLabel">0:00</span>
    </div>
  </div>

  <!-- Right Controls -->
  <div class="player-right">
    <div class="equalizer" id="audioEqualizer" title="Reproductor Activo" style="margin-right:12px;">
      <div class="eq-bar"></div>
      <div class="eq-bar"></div>
      <div class="eq-bar"></div>
      <div class="eq-bar"></div>
    </div>

    <div class="volume-container">
      <button class="control-btn" id="btnMute" title="Silenciar">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor" id="volumeIcon">
          <path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z"/>
        </svg>
      </button>
      <input type="range" class="volume-slider" id="volumeSlider" min="0" max="100" value="80" title="Volumen">
    </div>
  </div>
</footer>

<audio id="globalAudio" preload="none"></audio>
<script src="js/player.js"></script>
