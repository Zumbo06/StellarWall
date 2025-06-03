import sys
import os
import ctypes
from ctypes import wintypes
import random
from datetime import datetime, time as dt_time, timedelta
from collections import deque
import json
import shutil

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QCheckBox, QFrame, QSlider,
    QSpinBox, QComboBox, QListWidget, QListWidgetItem, QStackedWidget,
    QGroupBox, QTimeEdit, QScrollArea, QSystemTrayIcon, QMenu, QTabWidget,
    QAbstractItemView, QGraphicsOpacityEffect
)
from PyQt6.QtGui import QMovie, QPixmap, QColor, QFont, QIcon, QScreen, QAction, QImage, QPainter, QPalette
from PyQt6.QtCore import Qt, QUrl, QSize, QTimer, QTime, QStandardPaths, QPropertyAnimation, QParallelAnimationGroup, QEasingCurve, QEventLoop

from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput, QVideoSink, QVideoFrame, QMediaMetaData
from PyQt6.QtMultimediaWidgets import QVideoWidget




APP_NAME = "StellarWall" 
SETTINGS_FILE_NAME = "settings.json"

try:
    import win32gui, win32con, win32com.client
    PYWIN32_AVAILABLE = True
except ImportError: PYWIN32_AVAILABLE = False; print("-" * 68); print("Warning: pywin32 library not found. Some features disabled."); print("pip install pywin32"); print("-" * 68)
try:
    from PIL import Image, ImageSequence
    PILLOW_AVAILABLE = True
except ImportError: PILLOW_AVAILABLE = False; print("-" * 68); print("Warning: Pillow library not found. GIF previews disabled."); print("pip install Pillow"); print("-" * 68)

user32 = ctypes.WinDLL('user32', use_last_error=True)
shell32 = ctypes.WinDLL('shell32', use_last_error=True)
SMTO_NORMAL = 0x0000; CSIDL_STARTUP = 0x0007


def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller. """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS (one-file mode)
        base_path = sys._MEIPASS
    except Exception:
        # Not running in a PyInstaller one-file bundle
        if getattr(sys, 'frozen', False):
            # Running in PyInstaller one-folder mode or other frozen context
            base_path = os.path.dirname(sys.executable)
        else:
            # Development mode: get path relative to the script file
            base_path = os.path.dirname(os.path.abspath(__file__))
            
    return os.path.join(base_path, relative_path)


def find_workerw():
    progman = user32.FindWindowW("Progman", None)
    if not progman: print("Error: Could not find Progman window."); return None
    result = wintypes.LPVOID(); user32.SendMessageTimeoutW(progman, 0x052C, 0,0,SMTO_NORMAL,1000,ctypes.byref(result))
    workerw_hwnd = None; current_hwnd = user32.FindWindowExW(progman, None, "WorkerW", None)
    while current_hwnd:
        if not user32.FindWindowExW(current_hwnd, None, "SHELLDLL_DefView", None): workerw_hwnd=current_hwnd; break
        current_hwnd = user32.FindWindowExW(progman, current_hwnd, "WorkerW", None)
    if not workerw_hwnd: workerw_hwnd = user32.FindWindowExW(progman, None, "WorkerW", None)
    if not workerw_hwnd:
        hwnd_iter = user32.FindWindowExW(None,None,"WorkerW",None)
        while hwnd_iter:
            if user32.IsWindowVisible(hwnd_iter):
                parent = user32.GetParent(hwnd_iter)
                if (parent == progman or parent == 0) and not user32.FindWindowExW(hwnd_iter,None,"SHELLDLL_DefView",None):
                    workerw_hwnd = hwnd_iter; break
            hwnd_iter = user32.FindWindowExW(None, hwnd_iter, "WorkerW", None)
    return workerw_hwnd

def set_wallpaper_parent(player_hwnd, workerw_hwnd):
    if not player_hwnd or not workerw_hwnd: return False
    res = user32.SetParent(player_hwnd, workerw_hwnd)
    if res == 0 and ctypes.get_last_error() != 0: print(f"Error setting parent: {ctypes.get_last_error()}"); return False
    return True

class WallpaperPlayerWindow(QWidget):
    def __init__(self, main_app_ref):
        super().__init__(); self.main_app = main_app_ref
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False); self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setStyleSheet("background-color: black;"); self.layout = QVBoxLayout(self); self.layout.setContentsMargins(0,0,0,0)
        
        self.gif_label = QLabel(self); self.gif_label.setAlignment(Qt.AlignmentFlag.AlignCenter); self.gif_label.hide(); self.layout.addWidget(self.gif_label)
        self.movie = None

        self.video_widget_a = QVideoWidget(self)
        self.video_widget_b = QVideoWidget(self)
        self.video_widget_a.hide(); self.video_widget_b.hide()
        self.layout.addWidget(self.video_widget_a); self.layout.addWidget(self.video_widget_b)
        
        self.player_a = QMediaPlayer()
        self.player_b = QMediaPlayer()
        
        self.audio_output_a = QAudioOutput()
        self.player_a.setAudioOutput(self.audio_output_a)
        self.audio_output_b = QAudioOutput()
        self.player_b.setAudioOutput(self.audio_output_b)

        self.player_a.setVideoOutput(self.video_widget_a)
        self.player_b.setVideoOutput(self.video_widget_b)
        
        self.active_player = self.player_a; self.active_video_widget = self.video_widget_a
        
        self.sound_enabled_for_current_mp4 = False
        self._loop_mp4_path = None 
        self._swap_initiated = False 
        self._swap_retry_count = 0 
        self.MAX_SWAP_RETRIES = 20 

        self.player_a.mediaStatusChanged.connect(lambda s, p=self.player_a: self._handle_mp4_generic_status(s, p))
        self.player_a.errorOccurred.connect(lambda err, msg, p=self.player_a: self._handle_mp4_error(err, msg, p))
        self.player_b.mediaStatusChanged.connect(lambda s, p=self.player_b: self._handle_mp4_generic_status(s, p))
        self.player_b.errorOccurred.connect(lambda err, msg, p=self.player_b: self._handle_mp4_error(err, msg, p))
        
        self.player_a.positionChanged.connect(lambda pos: self._handle_mp4_loop_position(pos, self.player_a))
        self.player_b.positionChanged.connect(lambda pos: self._handle_mp4_loop_position(pos, self.player_b))

        self.is_paused = False; self.current_file_path = None 
        screen_geometry=QApplication.primaryScreen().geometry(); self.setGeometry(screen_geometry)
        self._initial_play_setup_slot_connected_player = None
        self.content_hidden_by_focus_loss = False 

    def _get_player_id(self, player_instance):
        if player_instance == self.player_a: return "A"
        if player_instance == self.player_b: return "B"
        return "Unknown"

    def play_gif(self, file_path):
        self.clear_content()
        self.current_file_path = file_path
        self._loop_mp4_path = None 

        if not self.gif_label: self.gif_label=QLabel(self); self.gif_label.setAlignment(Qt.AlignmentFlag.AlignCenter); self.layout.insertWidget(0, self.gif_label)
        self.video_widget_a.hide(); self.video_widget_b.hide(); self.gif_label.show()
        self.movie=QMovie(file_path);
        if not self.movie.isValid(): self.gif_label.setText(f"Error loading GIF: {os.path.basename(file_path)}"); print(f"QMovie error GIF ({file_path}): {self.movie.lastErrorString()}"); return
        self.gif_label.setMovie(self.movie); self.movie.setScaledSize(self.size()); self.movie.start(); self.is_paused=False

    def play_mp4(self, file_path, sound_enabled=False):
        self.clear_content() 
        self.current_file_path = os.path.normpath(file_path) if file_path else None
        self.sound_enabled_for_current_mp4 = sound_enabled
        self._loop_mp4_path = os.path.normpath(file_path) if file_path else None

        if self.gif_label: self.gif_label.hide()
        
        self.active_player = self.player_a
        self.active_video_widget = self.video_widget_a
        self.video_widget_b.hide()
        self.video_widget_a.show()

        self._set_player_audio(self.player_a, sound_enabled)
        self._set_player_audio(self.player_b, False) 

        if self._initial_play_setup_slot_connected_player:
            try:
                self._initial_play_setup_slot_connected_player.mediaStatusChanged.disconnect(self._initial_mp4_play_setup_slot)
            except TypeError: pass 
        
        self.player_a.mediaStatusChanged.connect(self._initial_mp4_play_setup_slot)
        self._initial_play_setup_slot_connected_player = self.player_a
        
        if self.current_file_path:
            self.player_a.setSource(QUrl.fromLocalFile(self.current_file_path))
        else:
            print(f"WPW ({id(self)}): play_mp4 called with invalid file_path.")
            self.clear_content()
            return
        self.is_paused = False

    def _initial_mp4_play_setup_slot(self, status: QMediaPlayer.MediaStatus):
        player_instance = self.sender()
        if not player_instance: return
        
        if player_instance != self.active_player:
            try:
                if self._initial_play_setup_slot_connected_player == player_instance:
                    player_instance.mediaStatusChanged.disconnect(self._initial_mp4_play_setup_slot)
                    self._initial_play_setup_slot_connected_player = None
            except TypeError: pass
            return

        if status == QMediaPlayer.MediaStatus.LoadedMedia:
            try:
                player_instance.mediaStatusChanged.disconnect(self._initial_mp4_play_setup_slot)
                self._initial_play_setup_slot_connected_player = None
            except TypeError: pass
            
            player_instance.setPosition(0) 
            player_instance.play()

        elif status == QMediaPlayer.MediaStatus.EndOfMedia :
             try:
                player_instance.mediaStatusChanged.disconnect(self._initial_mp4_play_setup_slot)
                self._initial_play_setup_slot_connected_player = None
             except TypeError: pass
             err_file = self.current_file_path or "Unknown file"
             print(f"MP4 {err_file} initial load resulted in EndOfMedia immediately. Treating as error. Status: {status}")
             if self.main_app and hasattr(self.main_app, 'status_label'): 
                self.main_app.status_label.setText(f"Error (short/empty) MP4: {os.path.basename(err_file)}")

        elif status in [QMediaPlayer.MediaStatus.InvalidMedia, QMediaPlayer.MediaStatus.NoMedia, 
                        QMediaPlayer.MediaStatus.StalledMedia]:
            try:
                player_instance.mediaStatusChanged.disconnect(self._initial_mp4_play_setup_slot)
                self._initial_play_setup_slot_connected_player = None
            except TypeError: pass
            err_file = self.current_file_path or "Unknown file"
            print(f"MP4 {err_file} initial load/play issue: {status}")
            if self.main_app and hasattr(self.main_app, 'status_label'): 
                self.main_app.status_label.setText(f"Error loading MP4: {os.path.basename(err_file)}")

    def _set_player_audio(self, player, sound_on):
        audio_output = player.audioOutput()
        if audio_output: 
            audio_output.setMuted(not sound_on)
            if sound_on: audio_output.setVolume(0.5)
        elif hasattr(player,'setMuted'): 
            player.setMuted(not sound_on)

    def _is_player_visually_ready(self, player: QMediaPlayer) -> bool:
        if not player: return False
        status = player.mediaStatus()
        state = player.playbackState()
        has_video = player.hasVideo()

        if state == QMediaPlayer.PlaybackState.PlayingState and has_video:
            return True
        
        if (status == QMediaPlayer.MediaStatus.LoadedMedia or status == QMediaPlayer.MediaStatus.BufferedMedia) and \
           state == QMediaPlayer.PlaybackState.PausedState and self.is_paused and has_video: 
            return True
            
        if status == QMediaPlayer.MediaStatus.EndOfMedia and has_video: 
            return True
            
        return False

    def _handle_mp4_generic_status(self, status: QMediaPlayer.MediaStatus, player_instance: QMediaPlayer):
        if not player_instance: return

        if status == QMediaPlayer.MediaStatus.LoadedMedia:
            normalized_loop_path = os.path.normpath(self._loop_mp4_path) if self._loop_mp4_path else ""
            normalized_player_source = ""
            if player_instance.source().isValid():
                normalized_player_source = os.path.normpath(player_instance.source().toLocalFile())

            if player_instance != self.active_player and \
               self._loop_mp4_path and \
               player_instance.source().isValid() and \
               normalized_player_source == normalized_loop_path:
                if player_instance.playbackState() != QMediaPlayer.PlaybackState.PlayingState and not self.is_paused :
                    player_instance.play() 
        
        elif status == QMediaPlayer.MediaStatus.EndOfMedia and not self.is_paused:
            if player_instance == self.active_player:
                normalized_loop_path = os.path.normpath(self._loop_mp4_path) if self._loop_mp4_path else ""
                normalized_current_path = os.path.normpath(self.current_file_path) if self.current_file_path else ""

                if self._loop_mp4_path and normalized_loop_path == normalized_current_path:
                    if not self._swap_initiated: 
                        self._prepare_and_swap_mp4_players(force_swap=True)
                else:
                    if self.main_app:
                         self.main_app.play_next_from_playlist_on_media_end()

    def _handle_mp4_loop_position(self, position: int, player_instance: QMediaPlayer):
        if self._swap_initiated: return 

        normalized_loop_path = os.path.normpath(self._loop_mp4_path) if self._loop_mp4_path else ""
        normalized_current_path = os.path.normpath(self.current_file_path) if self.current_file_path else ""

        if player_instance != self.active_player or \
           not self._loop_mp4_path or \
           normalized_loop_path != normalized_current_path or \
           self.is_paused:
            return
        
        duration = player_instance.duration()
        if duration > 1000 and position >= duration - 1000: 
            self._prepare_and_swap_mp4_players(force_swap=False) 

    def _prepare_and_swap_mp4_players(self, force_swap=False):
        if not self._loop_mp4_path:
            self._swap_initiated = False 
            return
        if self._swap_initiated and not force_swap: 
            return
        
        self._swap_initiated = True 
        self._swap_retry_count = 0 

        player_that_was_active_when_swap_initiated = self.active_player
        standby_player = self.player_b if player_that_was_active_when_swap_initiated == self.player_a else self.player_a
        
        normalized_loop_path_for_prep = os.path.normpath(self._loop_mp4_path)
        current_standby_source_path_normalized = ""
        if standby_player.source().isValid():
            current_standby_source_path_normalized = os.path.normpath(standby_player.source().toLocalFile())
        
        source_needs_setting = (current_standby_source_path_normalized != normalized_loop_path_for_prep or
                                standby_player.mediaStatus() == QMediaPlayer.MediaStatus.NoMedia or
                                standby_player.playbackState() == QMediaPlayer.PlaybackState.StoppedState)

        if source_needs_setting:
            if standby_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState: 
                standby_player.stop()
            standby_player.setSource(QUrl.fromLocalFile(self._loop_mp4_path)) 
        
        standby_current_status = standby_player.mediaStatus()
        can_try_play_standby = (
            standby_current_status == QMediaPlayer.MediaStatus.LoadedMedia or
            standby_current_status == QMediaPlayer.MediaStatus.BufferedMedia or
            standby_current_status == QMediaPlayer.MediaStatus.EndOfMedia
        )

        if standby_player.playbackState() != QMediaPlayer.PlaybackState.PlayingState and \
           not self.is_paused and \
           can_try_play_standby: 
            standby_player.play()

        self._set_player_audio(standby_player, False) 

        delay_for_swap = 250 if not force_swap else 150 
        QTimer.singleShot(delay_for_swap, lambda p=player_that_was_active_when_swap_initiated: self._perform_actual_swap(p))

    def _perform_actual_swap(self, expected_player_to_be_active_before_this_swap):
        if not self._swap_initiated or self.active_player != expected_player_to_be_active_before_this_swap:
            if self._swap_initiated: self._swap_initiated = False
            return

        if self._swap_retry_count >= self.MAX_SWAP_RETRIES:
            print(f"WPW ({id(self)}): MAX_SWAP_RETRIES reached for player {self._get_player_id(expected_player_to_be_active_before_this_swap)}. Aborting swap.")
            self._swap_initiated = False
            if expected_player_to_be_active_before_this_swap.mediaStatus() == QMediaPlayer.MediaStatus.EndOfMedia:
                expected_player_to_be_active_before_this_swap.setPosition(0)
                expected_player_to_be_active_before_this_swap.play()
            return

        standby_player = self.player_b if self.active_player == self.player_a else self.player_a
        standby_video_widget = self.video_widget_b if self.active_video_widget == self.video_widget_a else self.video_widget_a
        
        standby_source_local_file = os.path.normpath(standby_player.source().toLocalFile()) if standby_player.source().isValid() else ""
        normalized_loop_path = os.path.normpath(self._loop_mp4_path) if self._loop_mp4_path else ""

        if not standby_player.source().isValid() or standby_source_local_file != normalized_loop_path:
            print(f"WPW ({id(self)}): Swap critical fail: Standby source incorrect.")
            self._swap_initiated = False 
            return

        if self._is_player_visually_ready(standby_player):
            old_active_player = self.active_player
            old_active_video_widget = self.active_video_widget
            
            self.active_player = standby_player
            self.active_video_widget = standby_video_widget
            
            old_active_video_widget.hide()
            self.active_video_widget.show()
            self.active_video_widget.raise_() 

            self._set_player_audio(self.active_player, self.sound_enabled_for_current_mp4)
            if self.active_player.playbackState() != QMediaPlayer.PlaybackState.PlayingState and not self.is_paused:
                self.active_player.play()
            
            QTimer.singleShot(150, lambda p=old_active_player: p.stop() if p and p.playbackState() != QMediaPlayer.PlaybackState.StoppedState else None)
            
            self._swap_initiated = False 
            self._swap_retry_count = 0
            
            QTimer.singleShot(250, self._prepare_next_loop_instance) 
        else:
            self._swap_retry_count += 1
            
            standby_current_status = standby_player.mediaStatus()
            can_try_play_standby_poll = (
                standby_current_status == QMediaPlayer.MediaStatus.LoadedMedia or
                standby_current_status == QMediaPlayer.MediaStatus.BufferedMedia or
                standby_current_status == QMediaPlayer.MediaStatus.EndOfMedia
            )
            if can_try_play_standby_poll and \
                standby_player.playbackState() != QMediaPlayer.PlaybackState.PlayingState and not self.is_paused:
                standby_player.play()

            QTimer.singleShot(150, lambda p=expected_player_to_be_active_before_this_swap: self._perform_actual_swap(p))

    def _prepare_next_loop_instance(self):
        if not self._loop_mp4_path or self.is_paused: return 

        standby_player = self.player_b if self.active_player == self.player_a else self.player_a 
        
        normalized_loop_path_for_prep = os.path.normpath(self._loop_mp4_path)
        current_standby_source_path_normalized = os.path.normpath(standby_player.source().toLocalFile()) if standby_player.source().isValid() else ""

        if current_standby_source_path_normalized != normalized_loop_path_for_prep or \
           standby_player.mediaStatus() == QMediaPlayer.MediaStatus.NoMedia or \
           standby_player.playbackState() == QMediaPlayer.PlaybackState.StoppedState:
            standby_player.setSource(QUrl.fromLocalFile(self._loop_mp4_path)) 
        
        if standby_player.playbackState() != QMediaPlayer.PlaybackState.PlayingState and not self.is_paused:
            standby_player.play() 

        self._set_player_audio(standby_player, False)

    def _handle_mp4_error(self, error, error_string, player_instance: QMediaPlayer):
        player_id = self._get_player_id(player_instance)
        src = player_instance.source().toLocalFile() if player_instance.source().isValid() else "N/A"
        print(f"WPW ({id(self)}) MP Error (Player {player_id}, Source: {src}, Err: {error}): {error_string}")
        player_instance.stop()
        if player_instance == self.active_player and self.main_app and hasattr(self.main_app, 'status_label'):
             self.main_app.status_label.setText(f"MP4 Error: {os.path.basename(src or '')}")
        self._swap_initiated = False 

    def pause_playback(self):
        if self.is_paused: return
        self.is_paused = True
        if self.movie and self.movie.state()==QMovie.MovieState.Running: self.movie.setPaused(True)
        
        if self.player_a and self.player_a.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self.player_a.pause()
        if self.player_b and self.player_b.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self.player_b.pause()

    def resume_playback(self):
        if not self.is_paused: return
        self.is_paused = False
        if self.movie and self.movie.state()==QMovie.MovieState.Paused: self.movie.setPaused(False)
        
        if self.active_player and self.active_player.playbackState() == QMediaPlayer.PlaybackState.PausedState:
            self.active_player.play()

        standby_player = self.player_b if self.active_player == self.player_a else self.player_a
        if standby_player and standby_player.playbackState() == QMediaPlayer.PlaybackState.PausedState:
            normalized_loop_path = os.path.normpath(self._loop_mp4_path) if self._loop_mp4_path else ""
            normalized_standby_source = os.path.normpath(standby_player.source().toLocalFile()) if standby_player.source().isValid() else ""
            
            if standby_player.source().isValid() and \
               self._loop_mp4_path and normalized_standby_source == normalized_loop_path:
                standby_player.play()

    def clear_content(self):
        self._swap_initiated = False 
        self._swap_retry_count = 0

        if self.movie: self.movie.stop(); self.movie.deleteLater(); self.movie=None 
        if hasattr(self, 'gif_label'): self.gif_label.hide()
        
        if self._initial_play_setup_slot_connected_player:
            try:
                self._initial_play_setup_slot_connected_player.mediaStatusChanged.disconnect(self._initial_mp4_play_setup_slot)
            except TypeError: pass
            self._initial_play_setup_slot_connected_player = None

        if self.player_a:
            
            if self.player_a.mediaStatus() != QMediaPlayer.MediaStatus.NoMedia:
                self.player_a.stop(); self.player_a.setSource(QUrl())
            if self.audio_output_a and self.player_a.audioOutput() == self.audio_output_a:
                self.player_a.setAudioOutput(None) 
        if self.player_b:
            
            if self.player_b.mediaStatus() != QMediaPlayer.MediaStatus.NoMedia:
                self.player_b.stop(); self.player_b.setSource(QUrl())
            if self.audio_output_b and self.player_b.audioOutput() == self.audio_output_b:
                self.player_b.setAudioOutput(None) 
        
        if self.video_widget_a: self.video_widget_a.hide()
        if self.video_widget_b: self.video_widget_b.hide()
        
        self.active_player = self.player_a 
        self.active_video_widget = self.video_widget_a

        self._loop_mp4_path = None
        self.current_file_path = None
        self.is_paused = False
        self.content_hidden_by_focus_loss = False

    def hide_content_widgets(self):
        if self.gif_label: self.gif_label.hide()
        if self.video_widget_a: self.video_widget_a.hide()
        if self.video_widget_b: self.video_widget_b.hide()
        self.content_hidden_by_focus_loss = True
        if self.main_app: self.main_app.log_msg("WPW: Content widgets hidden due to focus loss.")


    def show_content_widgets(self):
        if self.current_file_path:
            if self.current_file_path.lower().endswith('.gif') and self.gif_label:
                self.gif_label.show()
            elif self.current_file_path.lower().endswith('.mp4'):
                if self.active_video_widget: 
                    self.active_video_widget.show()
        self.content_hidden_by_focus_loss = False
        if self.main_app: self.main_app.log_msg("WPW: Content widgets shown after regaining focus.")


    def stop_and_clear_playback(self):
        self.clear_content()

    def closeEvent(self, event):
        self.stop_and_clear_playback()
        if self.player_a: self.player_a.deleteLater(); self.player_a = None
        if self.player_b: self.player_b.deleteLater(); self.player_b = None
        if self.audio_output_a: self.audio_output_a.deleteLater(); self.audio_output_a = None
        if self.audio_output_b: self.audio_output_b.deleteLater(); self.audio_output_b = None
        super().closeEvent(event)

    def resizeEvent(self, event): 
        super().resizeEvent(event)
        if self.movie and self.gif_label and self.gif_label.isVisible(): 
            self.movie.setScaledSize(self.size())

class LiveWallpaperApp(QMainWindow):
    DAYS_OF_WEEK = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    MAX_RECENT_WALLPAPERS = 5

    def __init__(self):
        super().__init__()
        self.workerw_hwnd = None
        self.current_wallpaper_path_single_mode_selection = None; self.current_audio_path = None
        self.bg_audio_player = QMediaPlayer(); self.bg_audio_output = QAudioOutput(); self.bg_audio_player.setAudioOutput(self.bg_audio_output); self.bg_audio_output.setVolume(0.5)
        self.wallpaper_playlist = []; self.current_playlist_index = -1; self.playlist_timer = QTimer(self); self.playlist_timer.timeout.connect(self.handle_playlist_timer_tick)
        self.is_playlist_active = False; self.interval_play_order = "Manual Order"
        self.time_of_day_wallpapers = {p:None for p in ["Morning","Afternoon","Evening","Night"]}; self.time_of_day_slots = {"Morning":dt_time(6,0),"Afternoon":dt_time(12,0),"Evening":dt_time(18,0),"Night":dt_time(22,0)}
        self.day_of_week_wallpapers = {d:[] for d in self.DAYS_OF_WEEK}; self.current_day_playlist_indices = {d:-1 for d in self.DAYS_OF_WEEK}; self.last_checked_day_int = -1
        self.recent_wallpapers = deque(maxlen=self.MAX_RECENT_WALLPAPERS); self.tray_engine_pause_resume_action = None; self.app_icon = None
        self.settings_file_path = self._get_settings_file_path(); self.setting_start_with_windows = False; self.setting_pause_on_focus_loss = False
        
        self.setting_video_preview_quality = Qt.TransformationMode.SmoothTransformation 
        self.setting_low_spec_mode_enabled = False 
        self.setting_aggressive_gpu_reduction_on_focus_loss = False

        self.desktop_focus_timer = QTimer(self); self.desktop_focus_timer.timeout.connect(self.check_desktop_focus)
        self.is_desktop_focused = True; self.wallpaper_was_manually_paused = False; self.audio_was_manually_stopped = True; self.audio_was_focus_paused = False
        
        self._preview_player = None; self._preview_sink = None; self._preview_target_label = None; self._preview_file_path_being_processed = None
        
        self.active_player_window = None 
        self.transition_player_window = None
        self.opacity_effect_active = None
        self.opacity_effect_transition = None
        self.current_transition_animation = None

        self.interval_playlist_ui_populated = False
        self.dow_playlists_ui_populated = False

        self._setup_window_properties(); self._setup_main_ui_layout_with_tabs(); self._setup_tray_icon(); self.load_settings()
        
        loaded_mode_idx = getattr(self, 'loaded_mode_index_from_settings', 0) 
        if hasattr(self, 'mode_combo'): 
            original_idx = self.mode_combo.currentIndex()
            self.mode_combo.setCurrentIndex(loaded_mode_idx) 
            if self.mode_combo.currentIndex() == original_idx and self.mode_combo.currentIndex() == loaded_mode_idx :
                 self.update_mode_ui(loaded_mode_idx) 
        if hasattr(self, 'loaded_mode_index_from_settings'): 
            del self.loaded_mode_index_from_settings 

    def log_msg(self, message):
        print(f"{datetime.now().strftime('%H:%M:%S.%f')} {APP_NAME}: {message}")

    def _get_settings_file_path(self): return os.path.join(os.path.dirname(os.path.abspath(__file__)), SETTINGS_FILE_NAME)
    def _setup_window_properties(self):
        self.setWindowTitle(f"{APP_NAME} Live Wallpaper")
        self.setGeometry(150, 150, 780, 850)
        
        self.icon_path_filesystem = get_resource_path("logo.png")
        self.log_msg(f"Attempting to load icon from filesystem path: {self.icon_path_filesystem}")

        if os.path.exists(self.icon_path_filesystem):
            self.app_icon = QIcon(self.icon_path_filesystem)
            if self.app_icon.isNull():
                self.log_msg(f"Warning: QIcon loaded from '{self.icon_path_filesystem}' is null, image might be invalid.")
                style = self.style()
                self.app_icon = QIcon(style.standardIcon(style.StandardPixmap.SP_DesktopIcon))
            else:
                 self.log_msg(f"App icon loaded successfully from: {self.icon_path_filesystem}")
        else: 
            self.log_msg(f"Warning: App icon file '{self.icon_path_filesystem}' not found. Using default.")
            style = self.style()
            self.app_icon = QIcon(style.standardIcon(style.StandardPixmap.SP_DesktopIcon))
        
        self.setWindowIcon(self.app_icon)

    def _setup_main_ui_layout_with_tabs(self):
        self.central_widget = QWidget(); self.setCentralWidget(self.central_widget)
        self.main_application_layout = QVBoxLayout(self.central_widget)
        self.main_application_layout.setContentsMargins(10,10,10,10); self.main_application_layout.setSpacing(10)
        self._apply_stylesheet() 
        
        self.main_tab_widget = QTabWidget(); self.main_application_layout.addWidget(self.main_tab_widget)
        
        self.wallpaper_tab_widget = QWidget()
        self.wallpaper_config_tab_layout = QVBoxLayout(self.wallpaper_tab_widget)
        self.wallpaper_config_tab_layout.setContentsMargins(15,15,15,15); self.wallpaper_config_tab_layout.setSpacing(10)
        self._create_wallpaper_config_tab_content()
        self.main_tab_widget.addTab(self.wallpaper_tab_widget, f"{APP_NAME} Engine")

        self.settings_tab_widget = QWidget()
        self.app_settings_tab_layout = QVBoxLayout(self.settings_tab_widget)
        self.app_settings_tab_layout.setContentsMargins(15,15,15,15); self.app_settings_tab_layout.setSpacing(10)
        self._create_application_settings_tab_ui()
        self.main_tab_widget.addTab(self.settings_tab_widget, "Application Settings")

    def _create_wallpaper_config_tab_content(self):
        title_label = QLabel(f"{APP_NAME} Live Wallpaper")
        title_font = QFont("Orbitron", 26) 
        font_check = QFont(title_font.family()) 
        if not QFont(font_check).exactMatch(): title_font = QFont("Segoe UI", 26, QFont.Weight.Bold) 
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setObjectName("titleLabel")
        self.wallpaper_config_tab_layout.addWidget(title_label)

        mode_group = QGroupBox("Wallpaper Mode")
        mode_layout = QHBoxLayout()
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Single Wallpaper", "Playlist - Interval", "Playlist - Time of Day", "Playlist - Day of Week"])
        self.mode_combo.currentIndexChanged.connect(self.update_mode_ui)
        mode_layout.addWidget(self.mode_combo)
        mode_group.setLayout(mode_layout)
        self.wallpaper_config_tab_layout.addWidget(mode_group)

        self.wallpaper_mode_config_stack = QStackedWidget()
        self.wallpaper_config_tab_layout.addWidget(self.wallpaper_mode_config_stack)

        self._create_single_mode_ui()
        self._create_playlist_interval_mode_ui()
        self._create_playlist_time_of_day_ui()
        self._create_playlist_day_of_week_ui()

        audio_group = QGroupBox("Background Audio")
        audio_main_layout = QVBoxLayout()
        audio_file_layout = QHBoxLayout()
        self.audio_file_label = QLabel("No audio selected")
        self.audio_file_label.setObjectName("fileLabel")
        self.audio_file_label.setMinimumHeight(30) 
        select_audio_button = QPushButton("Select Audio (MP3)")
        select_audio_button.setObjectName("selectButton")
        select_audio_button.clicked.connect(self.select_audio_file)
        audio_file_layout.addWidget(self.audio_file_label, 1) 
        audio_file_layout.addWidget(select_audio_button)
        audio_main_layout.addLayout(audio_file_layout)
        audio_controls_layout = QHBoxLayout()
        self.play_audio_button = QPushButton("Play Audio")
        self.play_audio_button.clicked.connect(self.play_background_audio)
        self.play_audio_button.setEnabled(False)
        self.stop_audio_button = QPushButton("Stop Audio")
        self.stop_audio_button.clicked.connect(self.stop_background_audio)
        self.stop_audio_button.setEnabled(False)
        self.bg_audio_volume_slider = QSlider(Qt.Orientation.Horizontal)
        self.bg_audio_volume_slider.setRange(0, 100)
        self.bg_audio_volume_slider.setValue(50)
        self.bg_audio_volume_slider.valueChanged.connect(self.set_background_audio_volume)
        audio_controls_layout.addWidget(self.play_audio_button)
        audio_controls_layout.addWidget(self.stop_audio_button)
        audio_controls_layout.addWidget(QLabel("Vol:"))
        audio_controls_layout.addWidget(self.bg_audio_volume_slider)
        audio_main_layout.addLayout(audio_controls_layout)
        audio_group.setLayout(audio_main_layout)
        self.wallpaper_config_tab_layout.addWidget(audio_group)

        self.wallpaper_config_tab_layout.addWidget(QFrame(frameShape=QFrame.Shape.HLine, frameShadow=QFrame.Shadow.Sunken))

        button_layout = QHBoxLayout()
        self.apply_button = QPushButton("Apply Wallpaper")
        self.apply_button.setObjectName("applyButton")
        self.apply_button.clicked.connect(self.handle_apply_action)
        self.apply_button.setEnabled(False) 
        self.pause_resume_button = QPushButton("Pause Visual")
        self.pause_resume_button.setObjectName("pauseButton") 
        self.pause_resume_button.clicked.connect(self.toggle_pause_wallpaper_ui_button)
        self.pause_resume_button.setEnabled(False)
        self.stop_button = QPushButton("Stop & Clear Visual")
        self.stop_button.setObjectName("stopButton")
        self.stop_button.clicked.connect(self.stop_clear_wallpaper_external)
        self.stop_button.setEnabled(False)
        button_layout.addStretch()
        button_layout.addWidget(self.apply_button)
        button_layout.addWidget(self.pause_resume_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()
        self.wallpaper_config_tab_layout.addLayout(button_layout)
        
        self.wallpaper_config_tab_layout.addStretch(1) 

        self.status_label = QLabel("Status: Ready")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.wallpaper_config_tab_layout.addWidget(self.status_label)

    def _create_application_settings_tab_ui(self):
        startup_group = QGroupBox("Application Startup")
        startup_layout = QVBoxLayout()
        self.start_with_windows_checkbox = QCheckBox("Start with Windows")
        if not PYWIN32_AVAILABLE: 
            self.start_with_windows_checkbox.setEnabled(False)
            self.start_with_windows_checkbox.setToolTip("pywin32 library not found. This feature is disabled.")
        self.start_with_windows_checkbox.toggled.connect(self.toggle_start_with_windows)
        startup_layout.addWidget(self.start_with_windows_checkbox)
        startup_layout.addStretch() 
        startup_group.setLayout(startup_layout)
        self.app_settings_tab_layout.addWidget(startup_group)

        optimization_group = QGroupBox("Performance Optimization")
        optimization_layout = QVBoxLayout()
        self.pause_on_focus_loss_checkbox = QCheckBox("Pause wallpaper & audio when Desktop is not active")
        if not PYWIN32_AVAILABLE: 
            self.pause_on_focus_loss_checkbox.setEnabled(False)
            self.pause_on_focus_loss_checkbox.setToolTip("pywin32 library not found. This feature is disabled.")
        self.pause_on_focus_loss_checkbox.toggled.connect(self.toggle_pause_on_focus_loss)
        optimization_layout.addWidget(self.pause_on_focus_loss_checkbox)

        self.aggressive_gpu_reduction_checkbox = QCheckBox("Aggressively reduce GPU use when desktop is not active")
        self.aggressive_gpu_reduction_checkbox.setToolTip("Hides wallpaper content when desktop loses focus. May cause flicker on focus change.")
        if not PYWIN32_AVAILABLE: 
             self.aggressive_gpu_reduction_checkbox.setEnabled(False)
             self.aggressive_gpu_reduction_checkbox.setToolTip("pywin32 library not found. This feature is disabled.")
        self.aggressive_gpu_reduction_checkbox.toggled.connect(self.toggle_aggressive_gpu_reduction)
        optimization_layout.addWidget(self.aggressive_gpu_reduction_checkbox)


        self.low_spec_mode_checkbox = QCheckBox("Low Spec PC Mode (Limit MP4s to 1080p)")
        self.low_spec_mode_checkbox.setToolTip("Prevents playback of MP4 videos with height greater than 1080 pixels.")
        self.low_spec_mode_checkbox.toggled.connect(self.toggle_low_spec_mode)
        optimization_layout.addWidget(self.low_spec_mode_checkbox)

        preview_quality_layout = QHBoxLayout()
        preview_quality_label = QLabel("Video Preview Scaling Quality:")
        self.preview_quality_combo = QComboBox()
        self.preview_quality_combo.addItems(["Smooth (Better Quality, Slower)", "Fast (Lower Quality, Faster)"])
        self.preview_quality_combo.currentIndexChanged.connect(self.on_preview_quality_changed)
        preview_quality_layout.addWidget(preview_quality_label)
        preview_quality_layout.addWidget(self.preview_quality_combo)
        optimization_layout.addLayout(preview_quality_layout)
        
        optimization_layout.addStretch() 
        optimization_group.setLayout(optimization_layout)
        self.app_settings_tab_layout.addWidget(optimization_group)
        
        self.app_settings_tab_layout.addStretch() 

    def toggle_aggressive_gpu_reduction(self, checked):
        self.setting_aggressive_gpu_reduction_on_focus_loss = checked
        self.log_msg(f"Aggressive GPU reduction on focus loss {'enabled' if checked else 'disabled'}.")
        if self.setting_pause_on_focus_loss or self.setting_aggressive_gpu_reduction_on_focus_loss:
            if not self.desktop_focus_timer.isActive(): self.desktop_focus_timer.start(1000) 
        else: 
            if self.desktop_focus_timer.isActive(): self.desktop_focus_timer.stop()

        if not checked and self.active_player_window and self.active_player_window.content_hidden_by_focus_loss:
            if not self.active_player_window.is_paused: 
                self.active_player_window.show_content_widgets()
        self.save_settings()

    def toggle_low_spec_mode(self, checked):
        self.setting_low_spec_mode_enabled = checked
        self.log_msg(f"Low Spec PC Mode {'enabled' if checked else 'disabled'}.")
        self.save_settings()
        if checked and self.active_player_window and self.active_player_window.current_file_path:
            file_ext = os.path.splitext(self.active_player_window.current_file_path)[1].lower()
            if file_ext == ".mp4":
                 self.status_label.setText("Low Spec Mode enabled. Will apply to next MP4.")


    def on_preview_quality_changed(self, index):
        if index == 0: 
            self.setting_video_preview_quality = Qt.TransformationMode.SmoothTransformation
        else: 
            self.setting_video_preview_quality = Qt.TransformationMode.FastTransformation
        self.log_msg(f"Preview quality set to: {'Smooth' if index == 0 else 'Fast'}")
        self.save_settings()
        if hasattr(self, 'mode_combo') and self.mode_combo.currentIndex() == 0 and \
           self.current_wallpaper_path_single_mode_selection:
            self._update_single_mode_preview(self.current_wallpaper_path_single_mode_selection)

    def _setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.app_icon) 
        self.tray_icon.setToolTip(f"{APP_NAME} Engine") 

        self.tray_menu = QMenu(self)
        show_action = QAction(f"Show/Hide {APP_NAME}", self) 
        show_action.triggered.connect(self.toggle_main_window_visibility)
        self.tray_menu.addAction(show_action)
        self.tray_menu.addSeparator()

        self.tray_engine_pause_resume_action = QAction("Pause Engine", self)
        self.tray_engine_pause_resume_action.triggered.connect(self.toggle_engine_pause_tray)
        self.tray_engine_pause_resume_action.setEnabled(False) 
        self.tray_menu.addAction(self.tray_engine_pause_resume_action)
        
        self.recent_wallpapers_menu = QMenu("Recent Wallpapers", self)
        self.tray_menu.addMenu(self.recent_wallpapers_menu)
        self.update_recent_wallpapers_tray_menu() 

        self.tray_menu.addSeparator()
        quit_action = QAction(f"Quit {APP_NAME} Engine", self) 
        quit_action.triggered.connect(self.quit_application)
        self.tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.show()
        self.tray_icon.activated.connect(self.handle_tray_activation)

    def handle_tray_activation(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger or \
           reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.toggle_main_window_visibility()

    def toggle_main_window_visibility(self):
        if self.isVisible():
            self.hide()
        else:
            self.showNormal() 
            self.activateWindow() 
            self.raise_() 

    def _create_single_mode_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        wp_section_label = QLabel("Single Visual Wallpaper")
        wp_section_label.setObjectName("sectionHeader")
        layout.addWidget(wp_section_label)
        
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout(preview_group) 
        self.single_mode_preview_label = QLabel("No wallpaper selected for preview.")
        self.single_mode_preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.single_mode_preview_label.setMinimumSize(320, 180) 
        self.single_mode_preview_label.setFrameShape(QFrame.Shape.StyledPanel) 
        self.single_mode_preview_label.setStyleSheet("background-color: #202030; border: 1px solid #404050; color: #808090;")
        preview_layout.addWidget(self.single_mode_preview_label)
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)

        file_layout = QHBoxLayout()
        self.single_file_label = QLabel("No wallpaper selected")
        self.single_file_label.setObjectName("fileLabel")
        self.single_file_label.setMinimumHeight(30)
        select_button = QPushButton("Select Wallpaper File")
        select_button.setObjectName("selectButton")
        select_button.clicked.connect(self.select_single_wallpaper_file)
        file_layout.addWidget(self.single_file_label, 1)
        file_layout.addWidget(select_button)
        layout.addLayout(file_layout)

        options_layout = QHBoxLayout()
        self.sound_checkbox = QCheckBox("Enable Sound (for MP4 Visual)")
        self.sound_checkbox.setObjectName("soundCheckbox") 
        self.sound_checkbox.setChecked(False) 
        options_layout.addWidget(self.sound_checkbox)
        options_layout.addStretch()
        layout.addLayout(options_layout)
        
        layout.addStretch() 
        self.wallpaper_mode_config_stack.addWidget(widget)

    def _create_playlist_interval_mode_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        pl_section_label = QLabel("Playlist - Interval Mode")
        pl_section_label.setObjectName("sectionHeader")
        layout.addWidget(pl_section_label)

        folder_select_layout = QHBoxLayout()
        self.playlist_folder_label = QLabel("No folder selected")
        self.playlist_folder_label.setObjectName("fileLabel")
        select_folder_button = QPushButton("Load Folder")
        select_folder_button.setToolTip("Load all wallpapers from a folder into the list below.")
        select_folder_button.clicked.connect(self.select_playlist_folder_and_populate_list)
        folder_select_layout.addWidget(self.playlist_folder_label, 1)
        folder_select_layout.addWidget(select_folder_button)
        layout.addLayout(folder_select_layout)

        add_remove_layout = QHBoxLayout()
        add_files_button = QPushButton("Add File(s) to List")
        add_files_button.setToolTip("Add individual wallpaper files to the list below.")
        add_files_button.clicked.connect(self.add_files_to_interval_playlist)
        self.remove_selected_button = QPushButton("Remove Selected from List")
        self.remove_selected_button.setToolTip("Remove the highlighted wallpaper(s) from the list.")
        self.remove_selected_button.clicked.connect(self.remove_selected_from_interval_playlist)
        self.remove_selected_button.setEnabled(False) 
        add_remove_layout.addWidget(add_files_button)
        add_remove_layout.addWidget(self.remove_selected_button)
        add_remove_layout.addStretch()
        layout.addLayout(add_remove_layout)


        list_management_layout = QHBoxLayout()
        self.interval_playlist_listwidget = QListWidget()
        self.interval_playlist_listwidget.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.interval_playlist_listwidget.setDefaultDropAction(Qt.DropAction.MoveAction) 
        self.interval_playlist_listwidget.model().rowsMoved.connect(self.sync_wallpaper_playlist_from_listwidget)
        self.interval_playlist_listwidget.itemSelectionChanged.connect(
            lambda: self.remove_selected_button.setEnabled(bool(self.interval_playlist_listwidget.selectedItems()))
        )
        list_management_layout.addWidget(self.interval_playlist_listwidget, 3) 

        reorder_buttons_layout = QVBoxLayout()
        move_up_button = QPushButton("Move Up")
        move_up_button.clicked.connect(self.move_interval_playlist_item_up)
        move_down_button = QPushButton("Move Down")
        move_down_button.clicked.connect(self.move_interval_playlist_item_down)
        clear_list_button = QPushButton("Clear List")
        clear_list_button.clicked.connect(self.clear_interval_playlist)
        reorder_buttons_layout.addWidget(move_up_button)
        reorder_buttons_layout.addWidget(move_down_button)
        reorder_buttons_layout.addSpacing(20) 
        reorder_buttons_layout.addWidget(clear_list_button)
        reorder_buttons_layout.addStretch()
        list_management_layout.addLayout(reorder_buttons_layout, 1) 
        layout.addLayout(list_management_layout)

        settings_row_layout = QHBoxLayout()
        interval_group = QGroupBox("Timing")
        interval_layout = QHBoxLayout(interval_group) 
        interval_layout.addWidget(QLabel("Change every:"))
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setRange(1, 3600) 
        self.interval_spinbox.setValue(30) 
        self.interval_unit_combo = QComboBox()
        self.interval_unit_combo.addItems(["Minutes", "Hours"])
        self.interval_spinbox.valueChanged.connect(self._update_active_interval_timer)
        self.interval_unit_combo.currentIndexChanged.connect(self._update_active_interval_timer)
        interval_layout.addWidget(self.interval_spinbox)
        interval_layout.addWidget(self.interval_unit_combo)
        settings_row_layout.addWidget(interval_group)

        play_order_group = QGroupBox("Order")
        play_order_layout = QHBoxLayout(play_order_group) 
        self.interval_play_order_combo = QComboBox()
        self.interval_play_order_combo.addItems(["Manual Order", "Sequential (Initial Shuffle)", "Shuffle Each Cycle", "Random Pick Each Time"])
        self.interval_play_order_combo.currentIndexChanged.connect(self.set_interval_play_order)
        play_order_layout.addWidget(self.interval_play_order_combo)
        settings_row_layout.addWidget(play_order_group)
        layout.addLayout(settings_row_layout)
        
        layout.addStretch()
        self.wallpaper_mode_config_stack.addWidget(widget)

    def set_interval_play_order(self, index):
        selected_order_text = self.interval_play_order_combo.itemText(index)
        if "Manual Order" in selected_order_text: self.interval_play_order = "Manual Order"
        elif "Sequential" in selected_order_text: self.interval_play_order = "Sequential" 
        elif "Shuffle Each Cycle" in selected_order_text: self.interval_play_order = "Shuffle Cycle" 
        elif "Random Pick" in selected_order_text: self.interval_play_order = "Random Pick" 
        
        if self.mode_combo.currentIndex() == 1 and self.wallpaper_playlist: 
            if self.interval_play_order == "Sequential" or self.interval_play_order == "Shuffle Cycle":
                random.shuffle(self.wallpaper_playlist) 
            self._populate_interval_listwidget_from_playlist() 
            self.interval_playlist_ui_populated = True 
            if self.is_playlist_active:
                self.status_label.setText(f"Play order: {self.interval_play_order}. Restart playlist to apply.")
        self.save_settings()

    def _update_active_interval_timer(self):
        if self.mode_combo.currentIndex() == 1 and self.is_playlist_active : 
            current_timer_is_active = self.playlist_timer.isActive()
            if current_timer_is_active:
                self.playlist_timer.stop()
            
            interval_value = self.interval_spinbox.value()
            unit = self.interval_unit_combo.currentText()
            if unit == "Hours":
                interval_value *= 60 
            
            new_interval_ms = interval_value * 60 * 1000 
            
            if new_interval_ms > 0:
                self.playlist_timer.setInterval(new_interval_ms)
                if current_timer_is_active: 
                    self.playlist_timer.start()
                self.status_label.setText(f"Interval updated to {interval_value} {unit.lower().rstrip('s')}.")
                self.save_settings()
            else:
                self.status_label.setText("Invalid interval. Timer not updated.")
                if current_timer_is_active: self.playlist_timer.stop() 

    def _create_playlist_time_of_day_ui(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        tod_section_label = QLabel("Playlist - Time of Day Mode")
        tod_section_label.setObjectName("sectionHeader")
        layout.addWidget(tod_section_label)

        self.tod_buttons = {}
        self.tod_labels = {}

        for period in self.time_of_day_slots.keys(): 
            row_layout = QHBoxLayout()
            time_display = self.time_of_day_slots[period].strftime("%H:%M")
            row_layout.addWidget(QLabel(f"{period} (~{time_display}):"))
            
            self.tod_labels[period] = QLabel("No wallpaper set")
            self.tod_labels[period].setObjectName("fileLabelMini") 
            
            self.tod_buttons[period] = QPushButton("Set")
            self.tod_buttons[period].setObjectName("setTodButton") 
            self.tod_buttons[period].clicked.connect(lambda checked=False, p=period: self.set_time_of_day_wallpaper(p))
            
            row_layout.addWidget(self.tod_labels[period], 1) 
            row_layout.addWidget(self.tod_buttons[period])
            layout.addLayout(row_layout)
        
        layout.addStretch()
        self.wallpaper_mode_config_stack.addWidget(widget)

    def _create_playlist_day_of_week_ui(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True) 

        content_widget = QWidget() 
        layout = QVBoxLayout(content_widget)

        dow_section_label = QLabel("Playlist - Day of Week Mode")
        dow_section_label.setObjectName("sectionHeader")
        layout.addWidget(dow_section_label)

        self.dow_list_widgets = {}
        self.dow_add_buttons = {}
        self.dow_clear_buttons = {}

        for day in self.DAYS_OF_WEEK:
            day_group = QGroupBox(day)
            day_layout = QVBoxLayout() 

            self.dow_add_buttons[day] = QPushButton(f"Add Wallpaper(s) to {day}")
            self.dow_add_buttons[day].clicked.connect(lambda checked=False, d=day: self.add_wallpapers_to_day(d))
            day_layout.addWidget(self.dow_add_buttons[day])

            self.dow_list_widgets[day] = QListWidget()
            self.dow_list_widgets[day].setFixedHeight(80) 
            day_layout.addWidget(self.dow_list_widgets[day])

            self.dow_clear_buttons[day] = QPushButton(f"Clear {day}'s Wallpapers")
            self.dow_clear_buttons[day].clicked.connect(lambda checked=False, d=day: self.clear_wallpapers_for_day(d))
            day_layout.addWidget(self.dow_clear_buttons[day])
            
            day_group.setLayout(day_layout)
            layout.addWidget(day_group)
        
        layout.addStretch() 
        scroll_area.setWidget(content_widget) 
        self.wallpaper_mode_config_stack.addWidget(scroll_area)

    def update_mode_ui(self, index):
        if hasattr(self, 'wallpaper_mode_config_stack'):
            self.wallpaper_mode_config_stack.setCurrentIndex(index)

        self.is_playlist_active = (index > 0)
        self.stop_clear_wallpaper_external()

        if index != 0: 
            if self._preview_player: 
                if self._preview_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState:
                    self._preview_player.stop()
                self._preview_player.setSource(QUrl())
                if self._preview_sink:
                    self._preview_player.setVideoSink(None) 
            self._preview_file_path_being_processed = None
            if hasattr(self, 'single_mode_preview_label'):
                self.single_mode_preview_label.setText("Preview N/A for this mode.")
                self.single_mode_preview_label.setPixmap(QPixmap())
        else: 
            self._update_single_mode_preview(self.current_wallpaper_path_single_mode_selection)
        
        if index == 1 and not self.interval_playlist_ui_populated:
            self._populate_interval_listwidget_from_playlist()
            self.interval_playlist_ui_populated = True
        elif index == 3 and not self.dow_playlists_ui_populated:
            self._populate_dow_listwidgets_from_data()
            # self.dow_playlists_ui_populated is set by _populate_dow_listwidgets_from_data

        if index == 0:
            self.apply_button.setEnabled(bool(self.current_wallpaper_path_single_mode_selection))
        elif index == 1:
            self.apply_button.setEnabled(len(self.wallpaper_playlist) > 0)
        elif index == 2:
            self.apply_button.setEnabled(any(self.time_of_day_wallpapers.values()))
        elif index == 3:
            self.apply_button.setEnabled(any(len(wps) > 0 for wps in self.day_of_week_wallpapers.values()))

    def _apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background-color: #1e1e2f; color: #d0d0ff; font-size: 9pt; }
            QTabWidget::pane { border-top: 1px solid #33334c; margin-top: -1px;}
            QTabBar::tab { background-color: #2a2a3f; color: #a0a0cc; padding: 8px 15px; border-top-left-radius: 4px; border-top-right-radius: 4px; border: 1px solid #33334c; border-bottom: none; margin-right: 2px;}
            QTabBar::tab:selected { background-color: #1e1e2f; color: #d0d0ff; border-bottom: 1px solid #1e1e2f; } 
            QTabBar::tab:!selected:hover { background-color: #3d3d5c; }
            QLabel#titleLabel { color: #7f7fff; padding-bottom: 10px; }
            QLabel#sectionHeader { color: #9f9fff; font-size: 11pt; font-weight: bold; margin-top:8px; margin-bottom:3px; border-bottom: 1px solid #33334c; padding-bottom: 2px;}
            QGroupBox { font-size: 10pt; font-weight:bold; color: #b0b0ef; border: 1px solid #33334c; border-radius: 5px; margin-top: 0.5em; padding: 0.5em;}
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px 0 3px; }
            QLabel#fileLabel, QLabel#statusLabel { color: #a0a0cc; padding: 5px; border: 1px solid #33334c; border-radius: 4px; background-color: #2a2a3f; min-height:20px; }
            QLabel#fileLabelMini { color: #a0a0cc; padding: 2px 5px; border: 1px solid #33334c; border-radius: 3px; background-color: #2a2a3f; font-size:8pt; }
            QPushButton { background-color: #3d3d5c; color: #e0e0ff; border: 1px solid #505070; padding: 7px 12px; border-radius: 5px; }
            QPushButton:hover { background-color: #505070; border: 1px solid #66668c; }
            QPushButton:pressed { background-color: #2a2a3f; }
            QPushButton:disabled { background-color: #2c2c3f; color: #707080; border-color: #404050; }
            QPushButton#pauseButton { background-color: #c8a032; } QPushButton#pauseButton:hover { background-color: #e0b640; } 
            QPushButton#setTodButton { padding: 3px 8px; font-size: 8pt; } 
            QComboBox, QSpinBox, QListWidget { border: 1px solid #33334c; border-radius: 3px; padding: 3px 5px; background-color: #2a2a3f; min-height:20px; color: #d0d0ff; }
            QComboBox::drop-down { border-left: 1px solid #33334c; } 
            QComboBox QAbstractItemView { background-color: #2a2a3f; color: #d0d0ff; selection-background-color: #505070; }
            QListWidget { padding: 1px;  }
            QCheckBox { spacing: 5px; color: #c0c0e0; }
            QCheckBox::indicator { width:13px; height:13px; border:1px solid #505070; border-radius:3px; }
            QCheckBox::indicator:unchecked { background-color: #2a2a3f; } QCheckBox::indicator:checked { background-color: #7f7fff; }
            QSlider::groove:horizontal { border:1px solid #33334c; background:#2a2a3f; height:6px; border-radius:3px; }
            QSlider::handle:horizontal { background:#7f7fff; border:1px solid #505070; width:14px; margin:-4px 0; border-radius:7px; }
            QFrame[frameShape="4"] { border-top: 1px solid #33334c; margin-top:8px; margin-bottom:8px; } 
            QScrollArea { border: none; background-color: transparent; }
        """)

    def select_single_wallpaper_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Single Wallpaper File", 
                                                 os.path.expanduser("~"), 
                                                 "Media Files (*.gif *.mp4);;All Files (*)")
        if file_path:
            self.current_wallpaper_path_single_mode_selection = file_path
            self.single_file_label.setText(os.path.basename(file_path))
            self.apply_button.setEnabled(True)
            self.status_label.setText(f"Single selected: {os.path.basename(file_path)}")
            self._update_single_mode_preview(file_path)
            self.save_settings()
        else: 
            self.current_wallpaper_path_single_mode_selection = None
            self.single_file_label.setText("No wallpaper selected")
            self.apply_button.setEnabled(False)
            self._update_single_mode_preview(None)
            self.save_settings()

    def _update_single_mode_preview(self, file_path):
        if not hasattr(self, 'single_mode_preview_label'): return

        current_mode_is_single = hasattr(self, 'mode_combo') and self.mode_combo.currentIndex() == 0

        def release_preview_player_resources():
            if self._preview_player:
                if self._preview_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState:
                    self._preview_player.stop()
                self._preview_player.setSource(QUrl()) 
                if self._preview_sink:
                    self._preview_player.setVideoSink(None) 
            self._preview_file_path_being_processed = None


        if not current_mode_is_single or not file_path or not os.path.exists(file_path):
            self.single_mode_preview_label.setText("No wallpaper selected or preview N/A.")
            self.single_mode_preview_label.setPixmap(QPixmap())
            release_preview_player_resources()
            return

        try:
            preview_width = self.single_mode_preview_label.width() - 10
            preview_height = self.single_mode_preview_label.height() - 10

            if file_path.lower().endswith(".gif") and PILLOW_AVAILABLE:
                release_preview_player_resources() 
                pil_image = Image.open(file_path)
                first_frame = pil_image.convert("RGBA")
                img_data = first_frame.tobytes("raw", "RGBA")
                q_image = QImage(img_data, first_frame.width, first_frame.height, QImage.Format.Format_RGBA8888)
                pixmap = QPixmap.fromImage(q_image)
                scaled_pixmap = pixmap.scaled(preview_width, preview_height,
                                              Qt.AspectRatioMode.KeepAspectRatio,
                                              self.setting_video_preview_quality)
                self.single_mode_preview_label.setPixmap(scaled_pixmap)
                self.single_mode_preview_label.setText("")
            elif file_path.lower().endswith(".mp4"):
                self._grab_mp4_frame_for_preview(file_path) 
            else: 
                release_preview_player_resources()
                self.single_mode_preview_label.setText("Unsupported for preview.")
                self.single_mode_preview_label.setPixmap(QPixmap())
        except ImportError:
            release_preview_player_resources()
            self.single_mode_preview_label.setText("Preview: Pillow library missing for GIFs.")
            self.single_mode_preview_label.setPixmap(QPixmap())
        except Exception as e:
            release_preview_player_resources()
            self.single_mode_preview_label.setText(f"Preview Error:\n{str(e)[:100]}")
            self.single_mode_preview_label.setPixmap(QPixmap())
            print(f"Error previewing {file_path}: {e}")

    def _grab_mp4_frame_for_preview(self, file_path):
        if not hasattr(self, 'single_mode_preview_label'): return 
        
        self._preview_target_label = self.single_mode_preview_label 
        self._preview_target_label.setText("Loading MP4 preview...")
        self._preview_target_label.setPixmap(QPixmap()) 
        self._preview_file_path_being_processed = file_path 

        if not self._preview_player: 
            self._preview_player = QMediaPlayer()
            self._preview_sink = QVideoSink()
            self._preview_player.errorOccurred.connect(self._handle_preview_player_error)
            self._preview_player.mediaStatusChanged.connect(
                lambda status, player=self._preview_player: self._handle_preview_media_status_changed_for_player(status, player)
            )
        
        try: self._preview_sink.videoFrameChanged.disconnect(self._handle_preview_frame)
        except TypeError: pass 
        self._preview_sink.videoFrameChanged.connect(self._handle_preview_frame)
        
        if self._preview_player.videoSink() != self._preview_sink: 
            self._preview_player.setVideoSink(self._preview_sink)

        if self._preview_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState:
            self._preview_player.stop() 
        self._preview_player.setSource(QUrl.fromLocalFile(file_path))

    def _handle_preview_media_status_changed_for_player(self, status, player_instance=None): 
        if player_instance != self._preview_player or not self._preview_player or \
           not self._preview_player.source().isValid() or \
           self._preview_player.source().toLocalFile() != self._preview_file_path_being_processed:
            return

        if status == QMediaPlayer.MediaStatus.LoadedMedia:
            self._preview_player.setPosition(1) 
            self._preview_player.pause() 
        elif status in [QMediaPlayer.MediaStatus.EndOfMedia, QMediaPlayer.MediaStatus.InvalidMedia, 
                        QMediaPlayer.MediaStatus.NoMedia, QMediaPlayer.MediaStatus.StalledMedia]:
            if self._preview_target_label and self._preview_file_path_being_processed: 
                self._preview_target_label.setText(f"MP4 Preview Error:\nCould not load video frame.")
            if self._preview_player: self._preview_player.stop()
            self._preview_file_path_being_processed = None 

    def _handle_preview_frame(self, frame: QVideoFrame):
        if not self._preview_target_label or not self._preview_player or \
           (hasattr(self, 'wallpaper_mode_config_stack') and self.wallpaper_mode_config_stack.currentIndex() != 0) or \
           not (self._preview_player.source().isValid() and self._preview_player.source().toLocalFile() == self._preview_file_path_being_processed):
            
            if self._preview_player and self._preview_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState:
                self._preview_player.stop()
            try: 
                if self._preview_sink: self._preview_sink.videoFrameChanged.disconnect(self._handle_preview_frame)
            except TypeError: pass 
            return

        if not frame.isValid(): return

        video_image = frame.toImage()
        if not video_image.isNull():
            pixmap = QPixmap.fromImage(video_image)
            preview_width = self._preview_target_label.width() - 10 
            preview_height = self._preview_target_label.height() - 10
            scaled_pixmap = pixmap.scaled(preview_width, preview_height, 
                                          Qt.AspectRatioMode.KeepAspectRatio, 
                                          self.setting_video_preview_quality)
            self._preview_target_label.setPixmap(scaled_pixmap)
            self._preview_target_label.setText("") 
        
        if self._preview_player and self._preview_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState:
            self._preview_player.stop()
        try: 
            if self._preview_sink: self._preview_sink.videoFrameChanged.disconnect(self._handle_preview_frame)
        except TypeError: pass
        self._preview_file_path_being_processed = None

    def _handle_preview_player_error(self, error, error_string=""):
        current_preview_path = self._preview_file_path_being_processed 
        print(f"Preview Player Error: {error} - {error_string} for {current_preview_path}")
        if self._preview_target_label and current_preview_path: 
            err_msg = error_string if error_string else QMediaPlayer.Error(error).name 
            self._preview_target_label.setText(f"MP4 Preview Error:\n{err_msg[:100]}")
        if self._preview_player: self._preview_player.stop() 
        self._preview_file_path_being_processed = None 

    def _is_mp4_resolution_acceptable(self, file_path):
        self.log_msg(f"Low Spec Mode: Checking resolution for {os.path.basename(file_path)}")
        resolution = QSize() 
        loop = QEventLoop()

        check_player = QMediaPlayer()
        temp_sink = QVideoSink() 
        check_player.setVideoSink(temp_sink)

        connection_status_changed = None
        connection_error_occurred = None

        def on_status_changed(status):
            nonlocal resolution 
            if status == QMediaPlayer.MediaStatus.LoadedMedia:
                meta_res = check_player.metaData().value(QMediaMetaData.Key.Resolution)
                if isinstance(meta_res, QSize) and meta_res.isValid():
                    resolution = meta_res
                    self.log_msg(f"Low Spec Mode (MetaData): Resolution {resolution.width()}x{resolution.height()}")
                    loop.quit()
                    return

                if hasattr(check_player, 'videoTracks') and callable(check_player.videoTracks):
                    video_track_list = check_player.videoTracks()
                    if video_track_list:
                        frame_size = video_track_list[0].frameSize()
                        if frame_size.isValid():
                            resolution = frame_size
                            self.log_msg(f"Low Spec Mode (videoTracks): Resolution {resolution.width()}x{resolution.height()}")
                            loop.quit()
                            return
                
                self.log_msg(f"Low Spec Mode: Resolution not found via MetaData/videoTracks for {os.path.basename(file_path)} on LoadedMedia.")
                loop.quit() 

            elif status in [QMediaPlayer.MediaStatus.InvalidMedia, QMediaPlayer.MediaStatus.NoMedia, QMediaPlayer.MediaStatus.StalledMedia]:
                self.log_msg(f"Low Spec Mode: Media problem ({status}) for {os.path.basename(file_path)} during check.")
                loop.quit()

        def on_error(error_code, error_string): 
            self.log_msg(f"Low Spec Mode: Player error ({error_string}) during check for {os.path.basename(file_path)}.")
            loop.quit()

        connection_status_changed = check_player.mediaStatusChanged.connect(on_status_changed)
        connection_error_occurred = check_player.errorOccurred.connect(on_error)
        check_player.setSource(QUrl.fromLocalFile(file_path))

        timeout_timer = QTimer()
        timeout_timer.setSingleShot(True)
        timeout_timer.timeout.connect(loop.quit) 
        timeout_timer.start(3000) 

        loop.exec() 
        timeout_timer.stop() 
        
        if connection_status_changed:
            try: check_player.mediaStatusChanged.disconnect(connection_status_changed)
            except TypeError: pass 
        if connection_error_occurred:
            try: check_player.errorOccurred.disconnect(connection_error_occurred)
            except TypeError: pass
        
        check_player.stop() 
        check_player.setVideoSink(None) 
        check_player.setSource(QUrl())   
        check_player.deleteLater()
        temp_sink.deleteLater()

        if not resolution.isValid(): 
            self.log_msg(f"Low Spec Mode: Failed to determine resolution for {os.path.basename(file_path)}. Allowing playback.")
            return True 

        if resolution.height() > 1080: 
            self.log_msg(f"Low Spec Mode: Video {os.path.basename(file_path)} resolution ({resolution.width()}x{resolution.height()}) exceeds 1080p height. Blocking.")
            return False
        
        self.log_msg(f"Low Spec Mode: Video {os.path.basename(file_path)} resolution OK ({resolution.width()}x{resolution.height()}).")
        return True

    def select_playlist_folder_and_populate_list(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Wallpaper Folder", os.path.expanduser("~"))
        if folder_path:
            self.wallpaper_playlist = [] 
            valid_extensions = ('.gif', '.mp4') 
            
            try: file_list = sorted(os.listdir(folder_path)) 
            except OSError: self.status_label.setText("Error reading folder contents."); return

            for item in file_list:
                if item.lower().endswith(valid_extensions):
                    self.wallpaper_playlist.append(os.path.join(folder_path, item))
            
            if self.wallpaper_playlist: 
                if self.interval_play_order == "Sequential" or self.interval_play_order == "Shuffle Cycle":
                    random.shuffle(self.wallpaper_playlist) 
            
            self._populate_interval_listwidget_from_playlist()
            self.interval_playlist_ui_populated = True 
            self.playlist_folder_label.setText(os.path.basename(folder_path) if folder_path else "No folder selected")
            self.apply_button.setEnabled(len(self.wallpaper_playlist) > 0)
            self.status_label.setText(f"Loaded {len(self.wallpaper_playlist)} items. Order: {self.interval_play_order}")
            self.save_settings()

    def add_files_to_interval_playlist(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Add Wallpaper(s) to Interval Playlist", 
                                                os.path.expanduser("~"), 
                                                "Media Files (*.gif *.mp4);;All Files (*)")
        if files:
            added_count = 0
            for f_path in files:
                if f_path not in self.wallpaper_playlist: 
                    self.wallpaper_playlist.append(f_path)
                    item = QListWidgetItem(os.path.basename(f_path))
                    item.setData(Qt.ItemDataRole.UserRole, f_path) 
                    self.interval_playlist_listwidget.addItem(item)
                    added_count +=1
            if added_count > 0:
                self.apply_button.setEnabled(True) 
                self.status_label.setText(f"{added_count} file(s) added.")
                self.interval_playlist_ui_populated = True 
                self.save_settings()

    def remove_selected_from_interval_playlist(self):
        selected_items_widgets = self.interval_playlist_listwidget.selectedItems()
        if not selected_items_widgets: return

        rows_to_remove = sorted([self.interval_playlist_listwidget.row(item) for item in selected_items_widgets], reverse=True)
        
        removed_count = 0
        for row in rows_to_remove:
            if 0 <= row < len(self.wallpaper_playlist): 
                removed_path = self.wallpaper_playlist.pop(row) 
                self.interval_playlist_listwidget.takeItem(row) 
                removed_count +=1
        
        if removed_count > 0:
            self.status_label.setText(f"Removed {removed_count} item(s).")
            self.apply_button.setEnabled(len(self.wallpaper_playlist) > 0)
            if not self.wallpaper_playlist and self.is_playlist_active and self.mode_combo.currentIndex() == 1:
                self.stop_clear_wallpaper_external() 
            self.save_settings()

    def move_interval_playlist_item_up(self):
        current_row = self.interval_playlist_listwidget.currentRow()
        if current_row > 0: 
            item_widget = self.interval_playlist_listwidget.takeItem(current_row)
            self.interval_playlist_listwidget.insertItem(current_row - 1, item_widget)
            self.interval_playlist_listwidget.setCurrentRow(current_row - 1) 
            self.sync_wallpaper_playlist_from_listwidget() 

    def move_interval_playlist_item_down(self):
        current_row = self.interval_playlist_listwidget.currentRow()
        if 0 <= current_row < self.interval_playlist_listwidget.count() - 1: 
            item_widget = self.interval_playlist_listwidget.takeItem(current_row)
            self.interval_playlist_listwidget.insertItem(current_row + 1, item_widget)
            self.interval_playlist_listwidget.setCurrentRow(current_row + 1) 
            self.sync_wallpaper_playlist_from_listwidget() 
            
    def clear_interval_playlist(self):
        self.wallpaper_playlist.clear()
        self.interval_playlist_listwidget.clear()
        self.current_playlist_index = -1 
        self.apply_button.setEnabled(False)
        if self.is_playlist_active and self.mode_combo.currentIndex() == 1: 
            self.stop_clear_wallpaper_external()
        self.status_label.setText("Interval playlist cleared.")
        self.interval_playlist_ui_populated = True 
        self.save_settings()

    def sync_wallpaper_playlist_from_listwidget(self):
        updated_playlist = []
        for i in range(self.interval_playlist_listwidget.count()):
            item = self.interval_playlist_listwidget.item(i)
            full_path = item.data(Qt.ItemDataRole.UserRole) 
            if full_path and os.path.exists(full_path):
                updated_playlist.append(full_path)
            elif full_path: 
                print(f"Warning: File '{full_path}' in list no longer exists during sync.")
            else: 
                print(f"Warning: List item '{item.text()}' missing full path data during sync.")

        self.wallpaper_playlist = updated_playlist
        self.current_playlist_index = -1 
        self.status_label.setText("Playlist order updated from UI.")
        self.interval_playlist_ui_populated = True 
        self.save_settings() 

    def _populate_interval_listwidget_from_playlist(self):
        if not hasattr(self, 'interval_playlist_listwidget'): return 
        self.log_msg("Populating interval playlist UI.")
        self.interval_playlist_listwidget.clear()
        temp_valid_playlist = [] 
        for path in self.wallpaper_playlist:
            if os.path.exists(path):
                item = QListWidgetItem(os.path.basename(path))
                item.setData(Qt.ItemDataRole.UserRole, path) 
                self.interval_playlist_listwidget.addItem(item)
                temp_valid_playlist.append(path)
            else:
                print(f"Warning: File '{path}' from saved playlist not found. Skipping from UI list.")
        self.wallpaper_playlist = temp_valid_playlist 
        
    def set_time_of_day_wallpaper(self, period):
        file_path, _ = QFileDialog.getOpenFileName(self, f"Select Wallpaper for {period}", 
                                                 os.path.expanduser("~"), 
                                                 "Media Files (*.gif *.mp4)")
        if file_path:
            self.time_of_day_wallpapers[period] = file_path
            self.tod_labels[period].setText(os.path.basename(file_path))
            self.apply_button.setEnabled(any(self.time_of_day_wallpapers.values()))
            self.status_label.setText(f"{period} WP set: {os.path.basename(file_path)}")
            self.save_settings()

    def add_wallpapers_to_day(self, day_name):
        files, _ = QFileDialog.getOpenFileNames(self, f"Select Wallpaper(s) for {day_name}",
                                                os.path.expanduser("~"),
                                                "Media Files (*.gif *.mp4);;All Files (*)")
        if files:
            if day_name not in self.day_of_week_wallpapers:
                self.day_of_week_wallpapers[day_name] = []
            
            self.day_of_week_wallpapers[day_name].extend(files)
            if day_name in self.dow_list_widgets and self.dow_playlists_ui_populated : 
                self.dow_list_widgets[day_name].clear() 
                for file_path_item in self.day_of_week_wallpapers[day_name]:
                    self.dow_list_widgets[day_name].addItem(os.path.basename(file_path_item))
            
            self.apply_button.setEnabled(True) 
            self.status_label.setText(f"{len(files)} WP(s) added to {day_name}.")
            self.save_settings()

    def clear_wallpapers_for_day(self, day_name):
        if day_name in self.day_of_week_wallpapers:
            self.day_of_week_wallpapers[day_name] = []
            if day_name in self.dow_list_widgets and self.dow_playlists_ui_populated: 
                self.dow_list_widgets[day_name].clear()
            self.current_day_playlist_indices[day_name] = -1 
            
            self.apply_button.setEnabled(any(len(wps) > 0 for wps in self.day_of_week_wallpapers.values()))
            self.status_label.setText(f"Wallpapers cleared for {day_name}.")
            self.save_settings()

    def _populate_dow_listwidgets_from_data(self):
        if not hasattr(self, 'dow_list_widgets'): return
        self.log_msg("Populating Day of Week playlist UIs.")
        for day, paths in self.day_of_week_wallpapers.items():
            if day in self.dow_list_widgets:
                self.dow_list_widgets[day].clear()
                valid_paths_for_ui = []
                for p in paths:
                    if os.path.exists(p):
                        self.dow_list_widgets[day].addItem(os.path.basename(p))
                        valid_paths_for_ui.append(p)
                    else:
                        print(f"Warning: File '{p}' for DOW '{day}' not found. Skipping from UI.")
                self.day_of_week_wallpapers[day] = valid_paths_for_ui 
        self.dow_playlists_ui_populated = True


    def select_audio_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Background Audio File",
                                                 os.path.expanduser("~"), 
                                                 "Audio Files (*.mp3);;All Files (*)") 
        if file_path:
            self.current_audio_path = file_path
            self.audio_file_label.setText(os.path.basename(file_path))
            self.play_audio_button.setEnabled(True)
            self.status_label.setText(f"Audio selected: {os.path.basename(file_path)}")
            self.save_settings()

    def play_background_audio(self):
        if self.current_audio_path and os.path.exists(self.current_audio_path):
            self.bg_audio_player.setSource(QUrl.fromLocalFile(self.current_audio_path))
            self.bg_audio_player.play()
            self.stop_audio_button.setEnabled(True)
            self.play_audio_button.setText("Restart Audio") 
            
            try: self.bg_audio_player.mediaStatusChanged.disconnect(self._handle_bg_audio_status)
            except TypeError: pass 
            self.bg_audio_player.mediaStatusChanged.connect(self._handle_bg_audio_status)

            self.status_label.setText(f"Playing audio: {os.path.basename(self.current_audio_path)}")
            self.audio_was_manually_stopped = False 
            self.audio_was_focus_paused = False 
        elif self.current_audio_path: 
            self.status_label.setText(f"Audio file not found: {os.path.basename(self.current_audio_path)}")
            self.audio_file_label.setText("Audio file missing")
            self.current_audio_path = None 
            self.play_audio_button.setEnabled(False)

    def _handle_bg_audio_status(self, status):
        if status == QMediaPlayer.MediaStatus.EndOfMedia: 
            self.bg_audio_player.setPosition(0)
            self.bg_audio_player.play()

    def stop_background_audio(self):
        if self.bg_audio_player:
            self.bg_audio_player.stop()
        self.stop_audio_button.setEnabled(False)
        self.play_audio_button.setText("Play Audio") 
        self.status_label.setText("Background audio stopped.")
        self.audio_was_manually_stopped = True 
        self.audio_was_focus_paused = False 

    def set_background_audio_volume(self, value):
        if self.bg_audio_output:
            self.bg_audio_output.setVolume(float(value) / 100.0) 
        self.save_settings()

    def handle_apply_action(self):
        mode_index = self.mode_combo.currentIndex()
        self.is_playlist_active = (mode_index > 0) 
        
        path_to_play_if_single = self.current_wallpaper_path_single_mode_selection 

        if self.current_transition_animation and self.current_transition_animation.state() == QPropertyAnimation.State.Running:
            self.current_transition_animation.stop() 
            if self.transition_player_window: 
                self.transition_player_window.close()
                self.transition_player_window.deleteLater()
                self.transition_player_window = None
            if self.active_player_window and self.opacity_effect_active : 
                self.opacity_effect_active.setOpacity(1.0)

        self.playlist_timer.stop() 
        self.status_label.setText("Applying wallpaper...")
        QApplication.processEvents() 

        if not self.workerw_hwnd:
            self.workerw_hwnd = find_workerw()
        if not self.workerw_hwnd:
            self.status_label.setText("Error: Could not find desktop WorkerW.")
            return

        new_wallpaper_path = None
        if mode_index == 0: 
            new_wallpaper_path = path_to_play_if_single
        elif mode_index == 1: 
            if not self.interval_playlist_ui_populated: 
                self._populate_interval_listwidget_from_playlist()
                self.interval_playlist_ui_populated = True
            if self.wallpaper_playlist:
                if self.interval_play_order == "Sequential" or self.interval_play_order == "Shuffle Each Cycle":
                    random.shuffle(self.wallpaper_playlist) 
                    self._populate_interval_listwidget_from_playlist() 
                self.current_playlist_index = 0 
                if self.wallpaper_playlist: 
                    new_wallpaper_path = self.wallpaper_playlist[0]
                else: self.status_label.setText("Interval playlist is empty."); return
            else: self.status_label.setText("Interval playlist is empty."); return
        elif mode_index == 2: 
            new_wallpaper_path = self._get_current_time_of_day_wallpaper_path()
        elif mode_index == 3: 
            if not self.dow_playlists_ui_populated: 
                self._populate_dow_listwidgets_from_data()
            new_wallpaper_path = self._get_current_day_of_week_wallpaper_path(reset_sub_index=True)
        
        if not new_wallpaper_path or not os.path.exists(new_wallpaper_path):
            self.status_label.setText("No valid wallpaper for current mode/selection.")
            if self.active_player_window and self.active_player_window.isVisible():
                self._animate_fade(self.active_player_window, self.active_player_window.windowOpacity(), 0.0, 300, self._cleanup_after_fade_out_active)
            return

        self._transition_to_wallpaper(new_wallpaper_path, mode_index)
        self.save_settings() 

    def _create_and_setup_player_window(self):
        player = WallpaperPlayerWindow(self)
        opacity_effect = QGraphicsOpacityEffect(player) 
        player.setGraphicsEffect(opacity_effect)
        player.setWindowOpacity(0.0) 
        return player, opacity_effect

    def _transition_to_wallpaper(self, new_path, mode_index_for_timer_restart):
        if self.current_transition_animation and self.current_transition_animation.state() == QPropertyAnimation.State.Running:
            self.current_transition_animation.stop()
            if self.transition_player_window:
                self.transition_player_window.close() 
                self.transition_player_window.deleteLater()
                self.transition_player_window = None


        if self.transition_player_window: 
            self.transition_player_window.stop_and_clear_playback()
            self.transition_player_window.close()
            self.transition_player_window.deleteLater()
        
        self.transition_player_window, self.opacity_effect_transition = self._create_and_setup_player_window()
        
        self.transition_player_window.setWindowOpacity(0.0) 
        self.transition_player_window.show() 

        player_hwnd_int_trans = self.transition_player_window.winId()
        if not player_hwnd_int_trans: 
            self.status_label.setText("Error: New WP window has no ID.")
            if self.transition_player_window:
                self.transition_player_window.close(); self.transition_player_window.deleteLater()
            self.transition_player_window = None; return
            
        player_hwnd_trans = wintypes.HWND(int(player_hwnd_int_trans))
        
        if not set_wallpaper_parent(player_hwnd_trans, self.workerw_hwnd):
            self.status_label.setText("Error parenting new WP window.")
            if self.transition_player_window:
                self.transition_player_window.close(); self.transition_player_window.deleteLater()
            self.transition_player_window = None; return

        content_loaded = self._load_content_into_player(self.transition_player_window, new_path)
        if not content_loaded: 
            if self.transition_player_window:
                self.transition_player_window.close(); self.transition_player_window.deleteLater()
            self.transition_player_window = None
            return

        fade_duration = 500 
        self.current_transition_animation = QParallelAnimationGroup(self)

        if self.active_player_window and self.opacity_effect_active and self.active_player_window.windowOpacity() > 0.01 : 
            anim_out = QPropertyAnimation(self.opacity_effect_active, b"opacity")
            anim_out.setDuration(fade_duration)
            anim_out.setStartValue(self.active_player_window.windowOpacity()) 
            anim_out.setEndValue(0.0)
            anim_out.setEasingCurve(QEasingCurve.Type.InOutCubic)
            self.current_transition_animation.addAnimation(anim_out)

        anim_in = QPropertyAnimation(self.opacity_effect_transition, b"opacity")
        anim_in.setDuration(fade_duration)
        anim_in.setStartValue(0.0) 
        anim_in.setEndValue(1.0) 
        anim_in.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self.current_transition_animation.addAnimation(anim_in)
        
        self.current_transition_animation.finished.connect(
            lambda: self._finish_transition(new_path, mode_index_for_timer_restart)
        )
        self.current_transition_animation.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)

    def _load_content_into_player(self, player_window, file_path):
        if not player_window: return False
        if not file_path or not os.path.exists(file_path): 
            player_window.stop_and_clear_playback() 
            return False

        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension == ".mp4" and self.setting_low_spec_mode_enabled:
            if not self._is_mp4_resolution_acceptable(file_path):
                self.status_label.setText(f"Skipped (Low Spec): {os.path.basename(file_path)} >1080p")
                player_window.stop_and_clear_playback() 
                return False 


        if file_extension == ".gif":
            player_window.play_gif(file_path)
        elif file_extension == ".mp4":
            is_sound_enabled_for_new_active = False 
            if self.mode_combo.currentIndex() == 0: 
                 is_sound_enabled_for_new_active = self.sound_checkbox.isChecked()
            player_window.play_mp4(file_path, sound_enabled=is_sound_enabled_for_new_active) 
        else:
            self.status_label.setText(f"Unsupported type: {os.path.basename(file_path)}.")
            player_window.stop_and_clear_playback()
            return False
        
        self._add_to_recent_wallpapers(file_path) 
        self.status_label.setText(f"Preparing: {os.path.basename(file_path)}") 
        return True

    def _finish_transition(self, new_path_that_became_active, mode_index_of_new_path):
        old_active_player_window = self.active_player_window 
        
        self.active_player_window = self.transition_player_window
        self.opacity_effect_active = self.opacity_effect_transition
        
        self.transition_player_window = None 
        self.opacity_effect_transition = None

        if old_active_player_window:
            old_active_player_window.stop_and_clear_playback()
            old_active_player_window.hide() 
            old_active_player_window.close() 
            old_active_player_window.deleteLater() 

        self.stop_button.setEnabled(True)
        self.pause_resume_button.setEnabled(True)
        self.pause_resume_button.setText("Pause Visual") 
        if self.tray_engine_pause_resume_action:
            self.tray_engine_pause_resume_action.setText("Pause Engine")
            self.tray_engine_pause_resume_action.setEnabled(True)
        
        self.apply_button.setText("Restart Playlist" if self.is_playlist_active else "Change Visual")
        self.status_label.setText(f"Active: {os.path.basename(new_path_that_became_active)}")
        self.wallpaper_was_manually_paused = False 

        if self.is_playlist_active:
            timer_interval_ms = 0
            if mode_index_of_new_path == 1: 
                interval = self.interval_spinbox.value()
                unit = self.interval_unit_combo.currentText()
                if unit == "Hours": interval *= 60
                timer_interval_ms = interval * 60 * 1000
            elif mode_index_of_new_path == 2: 
                timer_interval_ms = 60 * 1000 
            elif mode_index_of_new_path == 3: 
                timer_interval_ms = 5 * 60 * 1000 
            
            if timer_interval_ms > 0:
                self.playlist_timer.setInterval(timer_interval_ms)
                self.playlist_timer.start()
        
        self.current_transition_animation = None 

    def _animate_fade(self, player_window_qwidget, start_val, end_val, duration, on_finished_slot=None):
        if not player_window_qwidget: return None
        
        opacity_effect = player_window_qwidget.graphicsEffect()
        if not isinstance(opacity_effect, QGraphicsOpacityEffect): 
            opacity_effect = QGraphicsOpacityEffect(player_window_qwidget)
            player_window_qwidget.setGraphicsEffect(opacity_effect)

        animation = QPropertyAnimation(opacity_effect, b"opacity", self) 
        animation.setDuration(duration)
        animation.setStartValue(float(start_val))
        animation.setEndValue(float(end_val))
        animation.setEasingCurve(QEasingCurve.Type.InOutCubic) 

        if on_finished_slot:
            animation.finished.connect(on_finished_slot)
        
        animation.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped) 
        return animation

    def _cleanup_after_fade_out_active(self):
        if self.active_player_window and self.active_player_window != self.transition_player_window: 
            self.active_player_window.stop_and_clear_playback()
            self.active_player_window.hide()
            self.active_player_window.close()
            self.active_player_window.deleteLater() 
            self.active_player_window = None 
            self.opacity_effect_active = None 

            self.stop_button.setEnabled(False)
            self.pause_resume_button.setEnabled(False)
            self.pause_resume_button.setText("Pause Visual")
            if self.tray_engine_pause_resume_action:
                 self.tray_engine_pause_resume_action.setEnabled(False)
                 self.tray_engine_pause_resume_action.setText("Pause Engine")

    def _get_current_time_of_day_wallpaper_path(self):
        now = datetime.now().time()
        selected_wp = None

        if self.time_of_day_slots["Night"] <= now or now < self.time_of_day_slots["Morning"]:
            selected_wp = self.time_of_day_wallpapers["Night"]
        elif self.time_of_day_slots["Evening"] <= now:
            selected_wp = self.time_of_day_wallpapers["Evening"]
        elif self.time_of_day_slots["Afternoon"] <= now:
            selected_wp = self.time_of_day_wallpapers["Afternoon"]
        elif self.time_of_day_slots["Morning"] <= now:
            selected_wp = self.time_of_day_wallpapers["Morning"]
        
        if not selected_wp: 
            for period_name in ["Morning", "Afternoon", "Evening", "Night"]: 
                if self.time_of_day_wallpapers[period_name]:
                    selected_wp = self.time_of_day_wallpapers[period_name]
                    break 
        return selected_wp

    def _get_current_day_of_week_wallpaper_path(self, reset_sub_index=False):
        today_int = datetime.today().weekday() 
        today_name = self.DAYS_OF_WEEK[today_int]
        
        wallpapers_for_today = self.day_of_week_wallpapers[today_name]
        if not wallpapers_for_today: return None 

        current_playing_file = self.active_player_window.current_file_path if self.active_player_window else None

        if reset_sub_index or today_int != self.last_checked_day_int or \
           (current_playing_file and current_playing_file not in wallpapers_for_today):
            self.current_day_playlist_indices[today_name] = 0 
            self.last_checked_day_int = today_int
        
        idx = self.current_day_playlist_indices[today_name]
        if not (0 <= idx < len(wallpapers_for_today)):
            self.current_day_playlist_indices[today_name] = 0 
            idx = 0
            if not wallpapers_for_today: return None 

        return wallpapers_for_today[idx]

    def handle_playlist_timer_tick(self):
        if self.active_player_window and self.active_player_window.is_paused:
            if self.playlist_timer.isActive(): 
                current_interval = self.playlist_timer.interval()
                self.playlist_timer.stop() 
                self.playlist_timer.start(current_interval) 
            return

        mode_index = self.mode_combo.currentIndex()
        next_wallpaper_path = None

        if mode_index == 1: 
            if not self.wallpaper_playlist: self.playlist_timer.stop(); return 
            
            if self.interval_play_order == "Manual Order" or self.interval_play_order == "Sequential":
                self.current_playlist_index = (self.current_playlist_index + 1) % len(self.wallpaper_playlist)
                next_wallpaper_path = self.wallpaper_playlist[self.current_playlist_index]
            elif self.interval_play_order == "Shuffle Each Cycle":
                self.current_playlist_index += 1
                if self.current_playlist_index >= len(self.wallpaper_playlist) or self.current_playlist_index < 0: 
                    if not self.wallpaper_playlist: self.playlist_timer.stop(); return 
                    random.shuffle(self.wallpaper_playlist)
                    self._populate_interval_listwidget_from_playlist() 
                    self.interval_playlist_ui_populated = True
                    self.current_playlist_index = 0
                if self.wallpaper_playlist: 
                    next_wallpaper_path = self.wallpaper_playlist[self.current_playlist_index]
            elif self.interval_play_order == "Random Pick":
                if self.wallpaper_playlist:
                    next_wallpaper_path = random.choice(self.wallpaper_playlist)
        
        elif mode_index == 2: 
            next_wallpaper_path = self._get_current_time_of_day_wallpaper_path()
        
        elif mode_index == 3: 
            today_int = datetime.today().weekday()
            if today_int != self.last_checked_day_int: 
                next_wallpaper_path = self._get_current_day_of_week_wallpaper_path(reset_sub_index=True)
            else: 
                return 
        
        current_playing_file = self.active_player_window.current_file_path if self.active_player_window else None
        
        if next_wallpaper_path and os.path.exists(next_wallpaper_path) and \
           (not current_playing_file or os.path.normpath(current_playing_file) != os.path.normpath(next_wallpaper_path)):
            self._transition_to_wallpaper(next_wallpaper_path, mode_index)
        elif next_wallpaper_path and not os.path.exists(next_wallpaper_path):
            self.status_label.setText(f"Playlist file missing: {os.path.basename(next_wallpaper_path)}. Skipping.")
            if mode_index == 1 and next_wallpaper_path in self.wallpaper_playlist:
                try:
                    self.wallpaper_playlist.remove(next_wallpaper_path)
                    self._populate_interval_listwidget_from_playlist() 
                    self.interval_playlist_ui_populated = True
                    if self.wallpaper_playlist: 
                        QTimer.singleShot(0, self.handle_playlist_timer_tick) 
                    else: 
                        self.playlist_timer.stop()
                        self.status_label.setText("Interval playlist empty.")
                except ValueError: pass 
        elif not self.wallpaper_playlist and mode_index == 1 : 
             self.playlist_timer.stop()
             self.status_label.setText("Interval playlist is empty.")

    def play_next_from_playlist_on_media_end(self):
        mode_index = self.mode_combo.currentIndex()
        if not self.is_playlist_active: return 

        if mode_index == 3: 
            today_int = datetime.today().weekday()
            today_name = self.DAYS_OF_WEEK[today_int]
            wallpapers_for_today = self.day_of_week_wallpapers[today_name]

            if wallpapers_for_today and len(wallpapers_for_today) > 0:
                current_idx = self.current_day_playlist_indices[today_name]
                current_idx = (current_idx + 1) % len(wallpapers_for_today) 
                self.current_day_playlist_indices[today_name] = current_idx
                
                next_wp_path = wallpapers_for_today[current_idx]
                if os.path.exists(next_wp_path):
                    self._transition_to_wallpaper(next_wp_path, mode_index)
                else:
                    self.status_label.setText(f"DoW file missing on cycle: {os.path.basename(next_wp_path)}")

    def _play_visual_content(self, file_path): 
        pass 

    def _add_to_recent_wallpapers(self, file_path):
        if file_path and os.path.exists(file_path):
            normalized_path = os.path.normpath(file_path)
            if normalized_path in self.recent_wallpapers:
                self.recent_wallpapers.remove(normalized_path)
            self.recent_wallpapers.appendleft(normalized_path) 
            self.update_recent_wallpapers_tray_menu() 

    def update_recent_wallpapers_tray_menu(self):
        if not hasattr(self, 'recent_wallpapers_menu') or not self.recent_wallpapers_menu: return

        self.recent_wallpapers_menu.clear()
        if not self.recent_wallpapers:
            no_recent_action = QAction("No recent wallpapers", self)
            no_recent_action.setEnabled(False)
            self.recent_wallpapers_menu.addAction(no_recent_action)
        else:
            temp_valid_recents = deque(maxlen=self.MAX_RECENT_WALLPAPERS) 
            processed_paths_for_menu = set()

            for wp_path in list(self.recent_wallpapers): 
                if os.path.exists(wp_path):
                    normalized_path = os.path.normpath(wp_path)
                    if normalized_path not in processed_paths_for_menu:
                        temp_valid_recents.append(normalized_path) 
                        processed_paths_for_menu.add(normalized_path)

                        action_text = os.path.basename(normalized_path)
                        max_len = 40 
                        if len(action_text) > max_len:
                            action_text = action_text[:max_len-3] + "..." 
                        
                        recent_action = QAction(action_text, self)
                        recent_action.triggered.connect(lambda checked=False, path=normalized_path: self.play_recent_wallpaper(path))
                        self.recent_wallpapers_menu.addAction(recent_action)
            
            self.recent_wallpapers = temp_valid_recents

    def play_recent_wallpaper(self, file_path):
        if not os.path.exists(file_path):
            self.status_label.setText(f"Error: Recent file '{os.path.basename(file_path)}' no longer exists.")
            try: self.recent_wallpapers.remove(file_path) 
            except ValueError: pass
            self.update_recent_wallpapers_tray_menu() 
            return

        self.status_label.setText(f"Playing recent: {os.path.basename(file_path)}")
        
        self.current_wallpaper_path_single_mode_selection = file_path
        if hasattr(self, 'single_file_label'): self.single_file_label.setText(os.path.basename(file_path))

        if self.mode_combo.currentIndex() != 0:
            self.mode_combo.setCurrentIndex(0) 
            QTimer.singleShot(10, self.handle_apply_action) 
        else: 
            self._update_single_mode_preview(file_path) 
            self.apply_button.setEnabled(True)
            self.handle_apply_action() 

    def toggle_pause_wallpaper_ui_button(self): self._perform_toggle_pause(is_engine_toggle=False)

    def toggle_engine_pause_tray(self): 
        visual_player = self.active_player_window
        visual_is_active_and_playing = visual_player and visual_player.current_file_path and not visual_player.is_paused
        visual_is_active_and_paused = visual_player and visual_player.current_file_path and visual_player.is_paused
        
        audio_is_playing = self.bg_audio_player and self.current_audio_path and not self.audio_was_manually_stopped and \
                           self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState
        audio_can_be_engine_paused = self.bg_audio_player and self.current_audio_path and not self.audio_was_manually_stopped and \
                                     (self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState or self.audio_was_focus_paused)

        if self.tray_engine_pause_resume_action.text() == "Pause Engine": 
            if visual_is_active_and_playing:
                visual_player.pause_playback()
                if self.setting_aggressive_gpu_reduction_on_focus_loss: 
                    visual_player.hide_content_widgets()
                if self.playlist_timer.isActive(): self.playlist_timer.stop()
                if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setText("Resume Visual"); self.pause_resume_button.setEnabled(True)
            
            if audio_is_playing: 
                self.bg_audio_player.pause()

            self.tray_engine_pause_resume_action.setText("Resume Engine")
            self.status_label.setText(f"{APP_NAME} Engine Paused (Tray).")
            self.wallpaper_was_manually_paused = True 
            if audio_can_be_engine_paused and self.bg_audio_player and self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PausedState:
                 self.audio_was_focus_paused = True 

        else: 
            if visual_is_active_and_paused and self.wallpaper_was_manually_paused : 
                if self.setting_aggressive_gpu_reduction_on_focus_loss and visual_player.content_hidden_by_focus_loss:
                     visual_player.show_content_widgets()
                visual_player.resume_playback()
                self._restart_playlist_timer_if_applicable()
                if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setText("Pause Visual")
            
            if self.audio_was_focus_paused and self.current_audio_path and not self.audio_was_manually_stopped and self.bg_audio_player: 
                self.bg_audio_player.play()
                self.audio_was_focus_paused = False

            self.tray_engine_pause_resume_action.setText("Pause Engine")
            self.status_label.setText(f"{APP_NAME} Engine Resumed (Tray).")
            self.wallpaper_was_manually_paused = False 
        
        self.save_settings() 

    def _perform_toggle_pause(self, is_engine_toggle=False, force_resume=False):
        active_player = self.active_player_window
        active_playing_path = active_player.current_file_path if active_player else None

        if not active_player or not active_playing_path: 
            if hasattr(self, 'pause_resume_button'):
                self.pause_resume_button.setEnabled(False)
                self.pause_resume_button.setText("Pause Visual")
            if not is_engine_toggle and self.tray_engine_pause_resume_action: 
                self.tray_engine_pause_resume_action.setEnabled(False)
                self.tray_engine_pause_resume_action.setText("Pause Engine")
            return

        if active_player.is_paused or force_resume: 
            if self.setting_aggressive_gpu_reduction_on_focus_loss and active_player.content_hidden_by_focus_loss:
                active_player.show_content_widgets() 
            active_player.resume_playback()
            if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setText("Pause Visual")
            self._restart_playlist_timer_if_applicable()
            self.status_label.setText(f"Resumed: {os.path.basename(active_playing_path)}")
            if not is_engine_toggle: 
                 self.wallpaper_was_manually_paused = False 

            if not is_engine_toggle and self.tray_engine_pause_resume_action:
                is_audio_ok_for_engine_play = not self.current_audio_path or \
                                              self.audio_was_manually_stopped or \
                                              (self.bg_audio_player and self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState) or \
                                              not self.audio_was_focus_paused 
                if is_audio_ok_for_engine_play :
                    self.tray_engine_pause_resume_action.setText("Pause Engine")
        
        else: 
            active_player.pause_playback()
            if self.setting_aggressive_gpu_reduction_on_focus_loss and not is_engine_toggle: 
                active_player.hide_content_widgets()

            if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setText("Resume Visual")
            if self.playlist_timer.isActive(): self.playlist_timer.stop() 
            self.status_label.setText(f"Paused: {os.path.basename(active_playing_path)}")
            if not is_engine_toggle: 
                self.wallpaper_was_manually_paused = True 

            if not is_engine_toggle and self.tray_engine_pause_resume_action:
                self.tray_engine_pause_resume_action.setText("Resume Engine") 
        
        if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setEnabled(True)
        if not is_engine_toggle and self.tray_engine_pause_resume_action: 
            self.tray_engine_pause_resume_action.setEnabled(True) 

        if not is_engine_toggle: self.save_settings()

    def _restart_playlist_timer_if_applicable(self):
        if self.is_playlist_active and not self.playlist_timer.isActive():
            mode_index = self.mode_combo.currentIndex()
            timer_interval_ms = 0
            if mode_index == 1: 
                interval = self.interval_spinbox.value()
                unit = self.interval_unit_combo.currentText()
                if unit == "Hours": interval *= 60
                timer_interval_ms = interval * 60 * 1000
            elif mode_index == 2: 
                timer_interval_ms = 60 * 1000 
            elif mode_index == 3: 
                timer_interval_ms = 5 * 60 * 1000 
            
            if timer_interval_ms > 0:
                self.playlist_timer.setInterval(timer_interval_ms)
                self.playlist_timer.start()

    def stop_clear_wallpaper_internal(self): 
        self.playlist_timer.stop()

        if self.current_transition_animation and self.current_transition_animation.state() == QPropertyAnimation.State.Running:
            self.current_transition_animation.stop()
        
        if self.active_player_window:
            self.active_player_window.stop_and_clear_playback()
            self.active_player_window.hide()
            self.active_player_window.close()
            self.active_player_window.deleteLater()
            self.active_player_window = None
            self.opacity_effect_active = None
        
        if self.transition_player_window: 
            self.transition_player_window.stop_and_clear_playback()
            self.transition_player_window.hide()
            self.transition_player_window.close()
            self.transition_player_window.deleteLater()
            self.transition_player_window = None
            self.opacity_effect_transition = None

        if self.tray_engine_pause_resume_action: 
            self.tray_engine_pause_resume_action.setEnabled(False)
            self.tray_engine_pause_resume_action.setText("Pause Engine")
        
        self.wallpaper_was_manually_paused = False 

    def stop_clear_wallpaper_external(self): 
        self.status_label.setText("Stopping & Clearing visual wallpaper...")
        self.stop_clear_wallpaper_internal() 

        self.stop_button.setEnabled(False)
        self.pause_resume_button.setEnabled(False)
        self.pause_resume_button.setText("Pause Visual")
        
        self.apply_button.setText("Apply Wallpaper") 
        current_mode_index = self.mode_combo.currentIndex()
        if current_mode_index == 0: 
            self.apply_button.setEnabled(bool(self.current_wallpaper_path_single_mode_selection))
        elif current_mode_index == 1: 
            self.apply_button.setEnabled(len(self.wallpaper_playlist) > 0)
        elif current_mode_index == 2: 
            self.apply_button.setEnabled(any(self.time_of_day_wallpapers.values()))
        elif current_mode_index == 3: 
            self.apply_button.setEnabled(any(len(wps) > 0 for wps in self.day_of_week_wallpapers.values()))
        
        self.status_label.setText("Visual wallpaper stopped & cleared.")
        self.save_settings()

    def save_settings(self):
        interval_val = getattr(self,'interval_spinbox',None) and self.interval_spinbox.value() or 30
        interval_unit_idx = getattr(self,'interval_unit_combo',None) and self.interval_unit_combo.currentIndex() or 0
        single_sound = getattr(self,'sound_checkbox',None) and self.sound_checkbox.isChecked() or False
        
        start_win_cb = getattr(self,'start_with_windows_checkbox',None)
        start_win = start_win_cb.isChecked() if start_win_cb else self.setting_start_with_windows 
        
        pause_focus_cb = getattr(self,'pause_on_focus_loss_checkbox',None)
        pause_focus = pause_focus_cb.isChecked() if pause_focus_cb else self.setting_pause_on_focus_loss

        playlist_folder_display_text = "No folder selected"
        if hasattr(self,'playlist_folder_label') and self.playlist_folder_label:
            playlist_folder_display_text = self.playlist_folder_label.text()
        if playlist_folder_display_text == "No folder selected": playlist_folder_display_text = ""

        current_interval_play_order = self.interval_play_order 

        preview_quality_idx = 0
        if hasattr(self, 'preview_quality_combo'):
            preview_quality_idx = self.preview_quality_combo.currentIndex()
        
        low_spec_mode = getattr(self, 'low_spec_mode_checkbox', None) and self.low_spec_mode_checkbox.isChecked() or self.setting_low_spec_mode_enabled
        agg_gpu_reduction = getattr(self, 'aggressive_gpu_reduction_checkbox', None) and self.aggressive_gpu_reduction_checkbox.isChecked() or self.setting_aggressive_gpu_reduction_on_focus_loss


        settings_data = {
            "version": 1.9, 
            "app_name_for_shortcut": APP_NAME, 
            "last_mode_index": self.mode_combo.currentIndex() if hasattr(self, 'mode_combo') else 0,
            "single_wallpaper_path": self.current_wallpaper_path_single_mode_selection,
            "single_sound_enabled": single_sound,
            "interval_playlist_folder_display": playlist_folder_display_text,
            "interval_playlist_files": self.wallpaper_playlist,
            "interval_value": interval_val,
            "interval_unit_index": interval_unit_idx,
            "interval_play_order": current_interval_play_order,
            "time_of_day_wallpapers": self.time_of_day_wallpapers,
            "day_of_week_wallpapers": self.day_of_week_wallpapers,
            "background_audio_path": self.current_audio_path,
            "background_audio_volume": self.bg_audio_volume_slider.value() if hasattr(self, 'bg_audio_volume_slider') else 50,
            "recent_wallpapers": list(self.recent_wallpapers), 
            "auto_play_on_startup": True, 
            "last_active_wallpaper_path": self.active_player_window.current_file_path if self.active_player_window and self.active_player_window.current_file_path else None,
            "is_last_active_paused": self.wallpaper_was_manually_paused or \
                                     (self.active_player_window.is_paused if self.active_player_window else False), 
            "audio_was_playing_before_exit": not self.audio_was_manually_stopped and \
                                             (self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState if hasattr(self, 'bg_audio_player') and self.bg_audio_player else False),
            "setting_start_with_windows": start_win,
            "setting_pause_on_focus_loss": pause_focus,
            "setting_video_preview_quality_index": preview_quality_idx,
            "setting_low_spec_mode_enabled": low_spec_mode,
            "setting_aggressive_gpu_reduction_on_focus_loss": agg_gpu_reduction 
        }
        try:
            with open(self.settings_file_path, 'w') as f:
                json.dump(settings_data, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")
            if hasattr(self, 'status_label'): self.status_label.setText(f"Error saving settings: {e}")

    def load_settings(self):
        self.interval_playlist_ui_populated = False
        self.dow_playlists_ui_populated = False

        if not os.path.exists(self.settings_file_path):
            self.log_msg("No settings file found. Using defaults.")
            if hasattr(self, 'mode_combo'): self.update_mode_ui(self.mode_combo.currentIndex()) 
            if hasattr(self, 'preview_quality_combo'): 
                 self.preview_quality_combo.setCurrentIndex(0) 
                 self.setting_video_preview_quality = Qt.TransformationMode.SmoothTransformation
            if hasattr(self, 'low_spec_mode_checkbox'):
                self.low_spec_mode_checkbox.setChecked(False) 
                self.setting_low_spec_mode_enabled = False
            if hasattr(self, 'aggressive_gpu_reduction_checkbox'):
                self.aggressive_gpu_reduction_checkbox.setChecked(False) 
                self.setting_aggressive_gpu_reduction_on_focus_loss = False
            return False
        
        try:
            with open(self.settings_file_path, 'r') as f:
                settings_data = json.load(f)
            self.log_msg("Settings loaded.")

            # If app name changed, old shortcuts might still exist. This doesn't auto-clean them.
            # stored_app_name = settings_data.get("app_name_for_shortcut", APP_NAME) # For potential future use

            self.loaded_mode_index_from_settings = settings_data.get("last_mode_index", 0) 

            single_wp = settings_data.get("single_wallpaper_path")
            if single_wp and os.path.exists(single_wp):
                self.current_wallpaper_path_single_mode_selection = single_wp
                if hasattr(self,'single_file_label'): self.single_file_label.setText(os.path.basename(single_wp))
            
            if hasattr(self,'sound_checkbox'): self.sound_checkbox.setChecked(settings_data.get("single_sound_enabled", False))
            
            if hasattr(self,'playlist_folder_label'): 
                folder_display = settings_data.get("interval_playlist_folder_display", "No folder selected")
                self.playlist_folder_label.setText(folder_display if folder_display else "No folder selected")

            self.wallpaper_playlist = [p for p in settings_data.get("interval_playlist_files", []) if os.path.exists(p)]
            
            if hasattr(self,'interval_spinbox'): self.interval_spinbox.setValue(settings_data.get("interval_value", 30))
            if hasattr(self,'interval_unit_combo'): self.interval_unit_combo.setCurrentIndex(settings_data.get("interval_unit_index", 0))
            
            loaded_play_order_text = settings_data.get("interval_play_order", "Manual Order")
            self.set_interval_play_order_from_text(loaded_play_order_text) 

            loaded_tod = settings_data.get("time_of_day_wallpapers", {})
            for period, path in loaded_tod.items():
                if path and os.path.exists(path):
                    self.time_of_day_wallpapers[period] = path
                    if period in self.tod_labels and hasattr(self.tod_labels[period], 'setText'): 
                        self.tod_labels[period].setText(os.path.basename(path))
                else: self.time_of_day_wallpapers[period] = None 
            
            loaded_dow = settings_data.get("day_of_week_wallpapers", {})
            for day, paths in loaded_dow.items():
                if day in self.day_of_week_wallpapers: 
                    self.day_of_week_wallpapers[day] = [p for p in paths if os.path.exists(p)]
            
            audio_p = settings_data.get("background_audio_path")
            if audio_p and os.path.exists(audio_p):
                self.current_audio_path = audio_p
                if hasattr(self,'audio_file_label'): self.audio_file_label.setText(os.path.basename(audio_p))
                if hasattr(self,'play_audio_button'): self.play_audio_button.setEnabled(True)
            
            if hasattr(self, 'bg_audio_volume_slider') and self.bg_audio_volume_slider:
                self.bg_audio_volume_slider.setValue(settings_data.get("background_audio_volume", 50))
                if self.bg_audio_output: self.bg_audio_output.setVolume(float(self.bg_audio_volume_slider.value())/100.0)

            loaded_recents = settings_data.get("recent_wallpapers", [])
            self.recent_wallpapers.clear() 
            for path in loaded_recents: 
                if os.path.exists(path): self.recent_wallpapers.append(path) 
            self.update_recent_wallpapers_tray_menu()

            self.setting_start_with_windows = settings_data.get("setting_start_with_windows", False)
            if hasattr(self,'start_with_windows_checkbox'): self.start_with_windows_checkbox.setChecked(self.setting_start_with_windows)
            
            self.setting_pause_on_focus_loss = settings_data.get("setting_pause_on_focus_loss", False)
            if hasattr(self,'pause_on_focus_loss_checkbox'): self.pause_on_focus_loss_checkbox.setChecked(self.setting_pause_on_focus_loss)
            
            self.setting_aggressive_gpu_reduction_on_focus_loss = settings_data.get("setting_aggressive_gpu_reduction_on_focus_loss", False)
            if hasattr(self, 'aggressive_gpu_reduction_checkbox'):
                self.aggressive_gpu_reduction_checkbox.setChecked(self.setting_aggressive_gpu_reduction_on_focus_loss)

            self.toggle_pause_on_focus_loss(self.setting_pause_on_focus_loss, from_load=True) 
            self.toggle_aggressive_gpu_reduction(self.setting_aggressive_gpu_reduction_on_focus_loss) 


            preview_quality_idx = settings_data.get("setting_video_preview_quality_index", 0) 
            if hasattr(self, 'preview_quality_combo'):
                current_pq_idx = self.preview_quality_combo.currentIndex()
                self.preview_quality_combo.setCurrentIndex(preview_quality_idx) 
                if self.preview_quality_combo.currentIndex() == current_pq_idx and self.preview_quality_combo.currentIndex() != preview_quality_idx: 
                     self.setting_video_preview_quality = Qt.TransformationMode.SmoothTransformation if preview_quality_idx == 0 else Qt.TransformationMode.FastTransformation
            
            self.setting_low_spec_mode_enabled = settings_data.get("setting_low_spec_mode_enabled", False)
            if hasattr(self, 'low_spec_mode_checkbox'):
                self.low_spec_mode_checkbox.setChecked(self.setting_low_spec_mode_enabled)

            auto_play_enabled = settings_data.get("auto_play_on_startup", True)
            last_active_wp_on_exit = settings_data.get("last_active_wallpaper_path")
            was_paused_on_exit = settings_data.get("is_last_active_paused", False)
            audio_was_active_on_exit = settings_data.get("audio_was_playing_before_exit", False)

            def deferred_auto_play_and_state_restore():
                current_mode_now = self.mode_combo.currentIndex()
                visual_to_start = None

                if current_mode_now == 0:
                    if last_active_wp_on_exit and os.path.exists(last_active_wp_on_exit):
                        self.current_wallpaper_path_single_mode_selection = last_active_wp_on_exit
                        if hasattr(self,'single_file_label'): self.single_file_label.setText(os.path.basename(last_active_wp_on_exit))
                        self._update_single_mode_preview(last_active_wp_on_exit) 
                        self.apply_button.setEnabled(True)
                        visual_to_start = True
                elif current_mode_now == 1:
                    if not self.interval_playlist_ui_populated: self._populate_interval_listwidget_from_playlist(); self.interval_playlist_ui_populated = True
                    if self.wallpaper_playlist: visual_to_start = True
                elif current_mode_now == 2 and any(self.time_of_day_wallpapers.values()): visual_to_start = True
                elif current_mode_now == 3:
                    if not self.dow_playlists_ui_populated: self._populate_dow_listwidgets_from_data()
                    if any(len(wps) > 0 for wps in self.day_of_week_wallpapers.values()): visual_to_start = True
                
                if auto_play_enabled and visual_to_start:
                    self.handle_apply_action()
                    if was_paused_on_exit:
                        QTimer.singleShot(750, lambda: self._perform_toggle_pause(is_engine_toggle=True)) 
                    
                if audio_was_active_on_exit and self.current_audio_path and os.path.exists(self.current_audio_path):
                    self.play_background_audio()
                    if was_paused_on_exit and not self.audio_was_manually_stopped:
                        QTimer.singleShot(800, lambda: self.bg_audio_player.pause() if self.bg_audio_player and self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState else None)
                        self.audio_was_focus_paused = True

            if hasattr(self, 'mode_combo'):
                original_mode_idx_before_load = self.mode_combo.currentIndex()
                self.mode_combo.setCurrentIndex(self.loaded_mode_index_from_settings)
                if self.mode_combo.currentIndex() == original_mode_idx_before_load and \
                   self.mode_combo.currentIndex() == self.loaded_mode_index_from_settings :
                    self.update_mode_ui(self.loaded_mode_index_from_settings) 
            
            QTimer.singleShot(300, deferred_auto_play_and_state_restore)
            return True

        except Exception as e:
            print(f"Error loading settings: {e} (Line: {e.__traceback__.tb_lineno if e.__traceback__ else 'N/A'})")
            if hasattr(self, 'status_label'): self.status_label.setText(f"Error loading settings: {e}")
            # Fallback to defaults for all settings if load fails badly
            if hasattr(self, 'mode_combo'): self.update_mode_ui(self.mode_combo.currentIndex())
            if hasattr(self, 'preview_quality_combo'):
                 self.preview_quality_combo.setCurrentIndex(0)
                 self.setting_video_preview_quality = Qt.TransformationMode.SmoothTransformation
            if hasattr(self, 'low_spec_mode_checkbox'):
                self.low_spec_mode_checkbox.setChecked(False)
                self.setting_low_spec_mode_enabled = False
            if hasattr(self, 'aggressive_gpu_reduction_checkbox'):
                self.aggressive_gpu_reduction_checkbox.setChecked(False)
                self.setting_aggressive_gpu_reduction_on_focus_loss = False
            return False

    def set_interval_play_order_from_text(self, order_text_from_settings):
        if hasattr(self, 'interval_play_order_combo'):
            for i in range(self.interval_play_order_combo.count()):
                combo_item_text = self.interval_play_order_combo.itemText(i)
                if (order_text_from_settings == "Manual Order" and "Manual Order" in combo_item_text) or \
                   (order_text_from_settings == "Sequential" and "Sequential (Initial Shuffle)" in combo_item_text) or \
                   (order_text_from_settings == "Shuffle Cycle" and "Shuffle Each Cycle" in combo_item_text) or \
                   (order_text_from_settings == "Random Pick" and "Random Pick Each Time" in combo_item_text):
                    self.interval_play_order_combo.setCurrentIndex(i)
                    return
            self.interval_play_order_combo.setCurrentIndex(0) 
        else: 
            self.interval_play_order = order_text_from_settings

    def toggle_start_with_windows(self, checked):
        if not PYWIN32_AVAILABLE:
            self.status_label.setText("Start with Windows: pywin32 library missing.")
            if hasattr(self, 'start_with_windows_checkbox'): self.start_with_windows_checkbox.setChecked(False) 
            return

        self.setting_start_with_windows = checked
        app_name_shortcut = APP_NAME + ".lnk" 
        current_script_path = os.path.abspath(sys.argv[0])
        
        target_path = sys.executable
        target_args = "" # No args if target is the .exe
        if not getattr(sys, 'frozen', False): # If running from Python script
            target_args = f'"{current_script_path}"'
            # Prefer pythonw.exe for silent startup if running from source
            if target_path.lower().endswith("python.exe"):
                pythonw_path = target_path.lower().replace("python.exe", "pythonw.exe")
                if os.path.exists(pythonw_path):
                    target_path = pythonw_path
        
        startup_path_bytes = ctypes.create_unicode_buffer(260) 
        if shell32.SHGetFolderPathW(None, CSIDL_STARTUP, None, 0, startup_path_bytes) == 0: 
            startup_folder = startup_path_bytes.value
            shortcut_path = os.path.join(startup_folder, app_name_shortcut)

            if checked: 
                try:
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.TargetPath = target_path
                    shortcut.Arguments = target_args
                    shortcut.WindowStyle = 7 
                    # Icon for shortcut should ideally come from the .exe itself
                    # This is set when PyInstaller builds the .exe (using --icon option)
                    shortcut.IconLocation = target_path 
                    shortcut.Description = f"{APP_NAME} - Live Wallpaper Engine" 
                    shortcut.WorkingDirectory = os.path.dirname(target_path if getattr(sys, 'frozen', False) else current_script_path)
                    shortcut.Save()
                    self.status_label.setText("Added to Windows startup.")
                except Exception as e:
                    self.status_label.setText(f"Error adding to startup: {e}")
                    print(f"Shortcut creation error: {e}")
                    if hasattr(self, 'start_with_windows_checkbox'): self.start_with_windows_checkbox.setChecked(False) 
            else: 
                if os.path.exists(shortcut_path):
                    try:
                        os.remove(shortcut_path)
                        self.status_label.setText("Removed from Windows startup.")
                    except Exception as e:
                        self.status_label.setText(f"Error removing from startup: {e}")
                        print(f"Shortcut removal error: {e}")
                        if hasattr(self, 'start_with_windows_checkbox'): self.start_with_windows_checkbox.setChecked(True) 
        else:
            self.status_label.setText("Could not find Startup folder.")
            if hasattr(self, 'start_with_windows_checkbox'): self.start_with_windows_checkbox.setChecked(False) 
        
        self.save_settings()

    def toggle_pause_on_focus_loss(self, checked, from_load=False):
        if not PYWIN32_AVAILABLE:
            self.status_label.setText("Pause on focus loss: pywin32 library missing.")
            if hasattr(self, 'pause_on_focus_loss_checkbox'): self.pause_on_focus_loss_checkbox.setChecked(False) 
            return

        self.setting_pause_on_focus_loss = checked
        
        if self.setting_pause_on_focus_loss or self.setting_aggressive_gpu_reduction_on_focus_loss:
            if not self.desktop_focus_timer.isActive(): self.desktop_focus_timer.start(1000) 
        else: 
            if self.desktop_focus_timer.isActive(): self.desktop_focus_timer.stop()
        
        if not checked:
            if self.active_player_window and self.active_player_window.is_paused and \
               not self.wallpaper_was_manually_paused and \
               (not self.setting_aggressive_gpu_reduction_on_focus_loss or not self.active_player_window.content_hidden_by_focus_loss):
                self._perform_toggle_pause(force_resume=True) 
            
            if self.bg_audio_player and self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PausedState and \
               self.audio_was_focus_paused and not self.wallpaper_was_manually_paused:
                self.bg_audio_player.play()
                self.audio_was_focus_paused = False
        
        if not from_load: 
            self.save_settings()

    def check_desktop_focus(self):
        if not PYWIN32_AVAILABLE: return
        if not self.setting_pause_on_focus_loss and not self.setting_aggressive_gpu_reduction_on_focus_loss:
            if self.desktop_focus_timer.isActive(): self.desktop_focus_timer.stop()
            return
        if not self.desktop_focus_timer.isActive(): self.desktop_focus_timer.start(1000)


        try:
            fg_window = win32gui.GetForegroundWindow()
            desktop_class_names = ["Progman", "WorkerW"] 
            fg_window_class = win32gui.GetClassName(fg_window)
            
            is_desktop_now = fg_window_class in desktop_class_names or \
                             fg_window == self.winId() or \
                             (self.active_player_window and fg_window == self.active_player_window.winId())

            if self.active_player_window and self.active_player_window.current_file_path and not self.wallpaper_was_manually_paused: 
                if not is_desktop_now and self.is_desktop_focused: 
                    if not self.active_player_window.is_paused: 
                        self.active_player_window.pause_playback()
                        if self.setting_aggressive_gpu_reduction_on_focus_loss:
                            self.active_player_window.hide_content_widgets()
                        if self.playlist_timer.isActive(): self.playlist_timer.stop()
                        if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setText("Resume Visual") 
                        if self.tray_engine_pause_resume_action and self.tray_engine_pause_resume_action.text() == "Pause Engine":
                            self.tray_engine_pause_resume_action.setText("Resume Engine") 

                elif is_desktop_now and not self.is_desktop_focused: 
                    if self.active_player_window.is_paused : 
                        if self.setting_aggressive_gpu_reduction_on_focus_loss and self.active_player_window.content_hidden_by_focus_loss:
                             self.active_player_window.show_content_widgets()
                        self.active_player_window.resume_playback()
                        self._restart_playlist_timer_if_applicable()
                        if hasattr(self, 'pause_resume_button'): self.pause_resume_button.setText("Pause Visual")
                        if self.tray_engine_pause_resume_action: 
                            is_audio_ok_for_engine_play = not self.current_audio_path or \
                                                          self.audio_was_manually_stopped or \
                                                          (self.bg_audio_player and self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState) or \
                                                          not self.audio_was_focus_paused
                            if is_audio_ok_for_engine_play:
                                self.tray_engine_pause_resume_action.setText("Pause Engine")

            if self.setting_pause_on_focus_loss: 
                if self.current_audio_path and not self.audio_was_manually_stopped and self.bg_audio_player: 
                    if not is_desktop_now and self.is_desktop_focused: 
                        if self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
                            self.bg_audio_player.pause()
                            self.audio_was_focus_paused = True 
                            if self.tray_engine_pause_resume_action and self.tray_engine_pause_resume_action.text() == "Pause Engine":
                                self.tray_engine_pause_resume_action.setText("Resume Engine") 

                    elif is_desktop_now and not self.is_desktop_focused: 
                        if self.bg_audio_player.playbackState() == QMediaPlayer.PlaybackState.PausedState and self.audio_was_focus_paused: 
                            self.bg_audio_player.play()
                            self.audio_was_focus_paused = False
                            if self.tray_engine_pause_resume_action: 
                                is_visual_ok_for_engine_play = not (self.active_player_window and self.active_player_window.current_file_path) or \
                                                                not self.active_player_window.is_paused or \
                                                                self.wallpaper_was_manually_paused
                                if is_visual_ok_for_engine_play:
                                    self.tray_engine_pause_resume_action.setText("Pause Engine")
            
            self.is_desktop_focused = is_desktop_now 
        except Exception as e:
            print(f"Error checking desktop focus: {e}")
            self.is_desktop_focused = True 

    def closeEvent(self, event):
        self.save_settings() 
        event.ignore() 
        self.hide()    
        self.tray_icon.showMessage(f"{APP_NAME} Engine", 
                                   "Minimized to system tray.",
                                   QSystemTrayIcon.MessageIcon.Information,
                                   2000) 

    def quit_application(self):
        self.save_settings() 
        if hasattr(self, 'status_label'): self.status_label.setText(f"Quitting {APP_NAME} Engine...") 
        
        self.stop_clear_wallpaper_internal() 

        if self.bg_audio_player:
            self.bg_audio_player.stop()
            self.bg_audio_player.setSource(QUrl())
            self.bg_audio_player.setAudioOutput(None)
            self.bg_audio_player.deleteLater()
            self.bg_audio_player = None
        if self.bg_audio_output:
            self.bg_audio_output.deleteLater() 
            self.bg_audio_output = None

        self.desktop_focus_timer.stop()      

        if self._preview_player:
            if self._preview_player.playbackState() != QMediaPlayer.PlaybackState.StoppedState:
                self._preview_player.stop()
            self._preview_player.setSource(QUrl())
            if self._preview_sink:
                try: self._preview_sink.videoFrameChanged.disconnect(self._handle_preview_frame)
                except TypeError: pass
                self._preview_player.setVideoSink(None)
                self._preview_sink.deleteLater() 
                self._preview_sink = None
            self._preview_player.deleteLater() 
            self._preview_player = None
        
        if hasattr(self, 'wallpaper_playlist'): self.wallpaper_playlist.clear()
        if hasattr(self, 'time_of_day_wallpapers'): self.time_of_day_wallpapers.clear()
        if hasattr(self, 'day_of_week_wallpapers'):
            for day_list in self.day_of_week_wallpapers.values(): day_list.clear()
            self.day_of_week_wallpapers.clear()
        if hasattr(self, 'recent_wallpapers'): self.recent_wallpapers.clear()

        if hasattr(self, 'tray_icon'): self.tray_icon.hide()                
        QApplication.instance().quit()       

if __name__ == "__main__":
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setQuitOnLastWindowClosed(False) 

    app = QApplication(sys.argv)
    main_window = LiveWallpaperApp()
    main_window.show()
    sys.exit(app.exec())