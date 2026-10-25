/**
 * PoliMusic - Player & Interactive Audio Engine
 * Classic ASP Spotify Clone
 */

(function () {
  let audio = null;
  let isPlaying = false;
  let isMuted = false;
  let isLooping = false;
  let previousVolume = 0.8;
  let currentSongIndex = -1;
  let songQueue = [];

  document.addEventListener("DOMContentLoaded", function () {
    initAudioPlayer();
    scanSongsOnPage();
    setupSearchFilter();
  });

  function initAudioPlayer() {
    audio = document.getElementById("globalAudio");
    if (!audio) {
      audio = new Audio();
      audio.id = "globalAudio";
      document.body.appendChild(audio);
    }

    const btnPlayPause = document.getElementById("btnPlayPause");
    const btnPrev = document.getElementById("btnPrev");
    const btnNext = document.getElementById("btnNext");
    const btnLoop = document.getElementById("btnLoop");
    const btnMute = document.getElementById("btnMute");
    const volumeSlider = document.getElementById("volumeSlider");
    const progressBar = document.getElementById("progressBarContainer");

    if (btnPlayPause) {
      btnPlayPause.addEventListener("click", togglePlayPause);
    }
    if (btnPrev) {
      btnPrev.addEventListener("click", playPrev);
    }
    if (btnNext) {
      btnNext.addEventListener("click", playNext);
    }
    if (btnLoop) {
      btnLoop.addEventListener("click", toggleLoop);
    }
    if (btnMute) {
      btnMute.addEventListener("click", toggleMute);
    }
    if (volumeSlider) {
      volumeSlider.value = 80;
      audio.volume = 0.8;
      volumeSlider.addEventListener("input", function (e) {
        audio.volume = parseFloat(e.target.value) / 100;
        if (audio.volume === 0) {
          isMuted = true;
          updateMuteIcon();
        } else {
          isMuted = false;
          updateMuteIcon();
        }
      });
    }

    if (progressBar) {
      progressBar.addEventListener("click", function (e) {
        if (!audio.duration || isNaN(audio.duration)) return;
        const rect = progressBar.getBoundingClientRect();
        const clickX = e.clientX - rect.left;
        const width = rect.width;
        const percent = Math.max(0, Math.min(1, clickX / width));
        audio.currentTime = percent * audio.duration;
      });
    }

    // Audio element listeners
    audio.addEventListener("timeupdate", updateProgress);
    audio.addEventListener("loadedmetadata", updateDurationDisplay);
    audio.addEventListener("play", function () {
      isPlaying = true;
      updatePlayPauseButtonUI(true);
    });
    audio.addEventListener("pause", function () {
      isPlaying = false;
      updatePlayPauseButtonUI(false);
    });
    audio.addEventListener("ended", function () {
      if (!isLooping) {
        playNext();
      }
    });
  }

  function scanSongsOnPage() {
    songQueue = [];
    const songElements = document.querySelectorAll("[data-song-id]");
    songElements.forEach((el, index) => {
      const songData = {
        id: el.getAttribute("data-song-id"),
        title: el.getAttribute("data-song-title") || "Desconocido",
        artist: el.getAttribute("data-song-artist") || "Artista",
        path: el.getAttribute("data-song-path"),
        cover: el.getAttribute("data-song-cover") || "albumImages/default.jpg",
        element: el,
        index: index
      };
      songQueue.push(songData);

      // Add click listener
      el.addEventListener("click", function (e) {
        // Prevent trigger if clicking on an interactive link/button inside
        if (e.target.closest("a") && !e.target.closest(".row-play-btn")) return;
        playTrackByIndex(index);
      });

      const playBtn = el.querySelector(".row-play-btn");
      if (playBtn) {
        playBtn.addEventListener("click", function (e) {
          e.stopPropagation();
          if (currentSongIndex === index && isPlaying) {
            togglePlayPause();
          } else {
            playTrackByIndex(index);
          }
        });
      }
    });
  }

  function playTrackByIndex(index) {
    if (index < 0 || index >= songQueue.length) return;
    currentSongIndex = index;
    const song = songQueue[index];

    // Highlight row
    document.querySelectorAll("[data-song-id]").forEach(el => {
      el.classList.remove("playing");
      const icon = el.querySelector(".row-play-icon");
      if (icon) {
        icon.innerHTML = '<path d="M8 5v14l11-7z"/>';
      }
    });

    if (song.element) {
      song.element.classList.add("playing");
      const icon = song.element.querySelector(".row-play-icon");
      if (icon) {
        icon.innerHTML = '<path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/>';
      }
    }

    // Update bottom player info
    const thumbEl = document.getElementById("playerThumb");
    const titleEl = document.getElementById("playerTitle");
    const artistEl = document.getElementById("playerArtist");

    if (thumbEl) thumbEl.src = song.cover;
    if (titleEl) titleEl.textContent = song.title;
    if (artistEl) artistEl.textContent = song.artist;

    // Load and play
    audio.src = song.path;
    audio.load();
    const playPromise = audio.play();
    if (playPromise !== undefined) {
      playPromise.catch(error => {
        console.warn("Audio autoplay blocked or file missing: ", error);
      });
    }

    // Increment play count via AJAX
    recordPlayCount(song.id, song.element);
  }

  function recordPlayCount(songId, songEl) {
    if (!songId) return;
    fetch("api_plays.asp?id=" + encodeURIComponent(songId), {
      method: "POST",
      body: ""
    })
      .then(res => res.json())
      .then(data => {
        if (data && data.success && songEl) {
          const playsCell = songEl.querySelector(".col-plays");
          if (playsCell && data.plays !== undefined) {
            playsCell.textContent = data.plays.toLocaleString();
          }
        }
      })
      .catch(err => console.log("Plays sync notice: ", err));
  }

  function togglePlayPause() {
    if (!audio.src) {
      if (songQueue.length > 0) {
        playTrackByIndex(0);
      }
      return;
    }
    if (audio.paused) {
      audio.play();
    } else {
      audio.pause();
    }
  }

  function playNext() {
    if (songQueue.length === 0) return;
    let nextIndex = currentSongIndex + 1;
    if (nextIndex >= songQueue.length) {
      nextIndex = 0; // Loop queue
    }
    playTrackByIndex(nextIndex);
  }

  function playPrev() {
    if (songQueue.length === 0) return;
    if (audio.currentTime > 3) {
      audio.currentTime = 0;
      return;
    }
    let prevIndex = currentSongIndex - 1;
    if (prevIndex < 0) {
      prevIndex = songQueue.length - 1;
    }
    playTrackByIndex(prevIndex);
  }

  function toggleLoop() {
    isLooping = !isLooping;
    audio.loop = isLooping;
    const btnLoop = document.getElementById("btnLoop");
    if (btnLoop) {
      btnLoop.classList.toggle("active", isLooping);
    }
  }

  function toggleMute() {
    isMuted = !isMuted;
    if (isMuted) {
      previousVolume = audio.volume;
      audio.volume = 0;
      const volumeSlider = document.getElementById("volumeSlider");
      if (volumeSlider) volumeSlider.value = 0;
    } else {
      audio.volume = previousVolume > 0 ? previousVolume : 0.8;
      const volumeSlider = document.getElementById("volumeSlider");
      if (volumeSlider) volumeSlider.value = audio.volume * 100;
    }
    updateMuteIcon();
  }

  function updateMuteIcon() {
    const iconEl = document.getElementById("volumeIcon");
    if (!iconEl) return;
    if (isMuted || audio.volume === 0) {
      iconEl.innerHTML = '<path d="M16.5 12c0-1.77-1.02-3.29-2.5-4.03v2.21l2.45 2.45c.03-.2.05-.41.05-.63zm2.5 0c0 .94-.2 1.82-.54 2.64l1.51 1.51C20.63 14.91 21 13.5 21 12c0-4.28-2.99-7.86-7-8.77v2.06c2.89.86 5 3.54 5 6.71zM4.27 3L3 4.27 7.73 9H3v6h4l5 5v-6.73l4.25 4.25c-.67.52-1.42.93-2.25 1.18v2.06c1.38-.31 2.63-.95 3.69-1.81L19.73 21 21 19.73l-9-9L4.27 3zM12 4L9.91 6.09 12 8.18V4z"/>';
    } else {
      iconEl.innerHTML = '<path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z"/>';
    }
  }

  function updatePlayPauseButtonUI(playing) {
    const mainBtnIcon = document.getElementById("mainPlayIcon");
    if (mainBtnIcon) {
      if (playing) {
        mainBtnIcon.innerHTML = '<path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/>';
      } else {
        mainBtnIcon.innerHTML = '<path d="M8 5v14l11-7z"/>';
      }
    }
    // Update active row play icon
    if (currentSongIndex >= 0 && currentSongIndex < songQueue.length) {
      const activeEl = songQueue[currentSongIndex].element;
      if (activeEl) {
        const icon = activeEl.querySelector(".row-play-icon");
        if (icon) {
          if (playing) {
            icon.innerHTML = '<path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/>';
          } else {
            icon.innerHTML = '<path d="M8 5v14l11-7z"/>';
          }
        }
      }
    }
  }

  function updateProgress() {
    if (!audio.duration || isNaN(audio.duration)) return;
    const percent = (audio.currentTime / audio.duration) * 100;
    const fillEl = document.getElementById("progressFill");
    const currentTimeEl = document.getElementById("currentTimeLabel");

    if (fillEl) fillEl.style.width = percent + "%";
    if (currentTimeEl) currentTimeEl.textContent = formatTime(audio.currentTime);
  }

  function updateDurationDisplay() {
    const totalTimeEl = document.getElementById("totalTimeLabel");
    if (totalTimeEl && audio.duration && !isNaN(audio.duration)) {
      totalTimeEl.textContent = formatTime(audio.duration);
    }
  }

  function formatTime(seconds) {
    if (isNaN(seconds)) return "0:00";
    const mins = Math.floor(seconds / 60);
    const secs = Math.floor(seconds % 60);
    return mins + ":" + (secs < 10 ? "0" : "") + secs;
  }

  function setupSearchFilter() {
    const searchInput = document.getElementById("topSearchInput");
    if (!searchInput) return;

    searchInput.addEventListener("input", function (e) {
      const q = e.target.value.toLowerCase().trim();
      const rows = document.querySelectorAll(".song-row");
      rows.forEach(row => {
        const title = (row.getAttribute("data-song-title") || "").toLowerCase();
        const artist = (row.getAttribute("data-song-artist") || "").toLowerCase();
        if (title.includes(q) || artist.includes(q)) {
          row.style.display = "";
        } else {
          row.style.display = "none";
        }
      });
    });
  }

  // Expose global play helper for cards or other components
  window.PoliPlayer = {
    playByIndex: playTrackByIndex,
    togglePlay: togglePlayPause
  };
})();
