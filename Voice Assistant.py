import sys
import json
import os
import subprocess
import platform
import re
import time
import webbrowser
import requests
import wikipedia
import pygame
import pyautogui
import psutil
import pywhatkit
import pyperclip
from datetime import datetime as dt
from urllib.parse import quote
import smtplib
from email.mime.text import MIMEText
from PIL import Image, ImageGrab
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit,
    QLabel, QWidget, QComboBox, QSlider, QComboBox, QFileDialog, QMessageBox,
    QTabWidget, QDockWidget, QToolBar, QToolButton, QMenu, QAction, QStatusBar,
    QInputDialog, QProgressBar, QSplitter, QFileSystemModel, QTreeView, QDialog,
    QSystemTrayIcon
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer, QSize, QJsonDocument, QDir
from PyQt5.QtGui import QIcon, QTextCursor, QPixmap, QFont, QPalette, QColor, QPainter, QPen
import speech_recognition as sr
import pyttsx3
from plyer import notification
import keyboard
import importlib.util
import threading
import numpy as np

# Suppress Wikipedia parser warning
wikipedia.wikipedia.GuessedAtParserWarning = lambda *args, **kwargs: None

# Windows-specific volume control
try:
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    volume_control_available = True
except ImportError:
    volume_control_available = False

class VoiceThread(QThread):
    """Thread for handling voice recognition."""
    finished_signal = pyqtSignal(str)
    listening_signal = pyqtSignal(bool)
    error_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.is_listening = False

    def run(self):
        self.is_listening = True
        self.listening_signal.emit(True)
        try:
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source)
                audio = self.recognizer.listen(source, timeout=5)
            self.is_listening = False
            self.listening_signal.emit(False)
            try:
                text = self.recognizer.recognize_google(audio)
                self.finished_signal.emit(text)
            except sr.UnknownValueError:
                self.error_signal.emit("Could not understand audio")
            except sr.RequestError as e:
                self.error_signal.emit(f"Could not request results; {e}")
        except Exception as e:
            self.error_signal.emit(f"Microphone error: {str(e)}")

class AnimatedVoiceIndicator(QLabel):
    """Animated waveform for voice activity."""
    def __init__(self):
        super().__init__()
        self.setFixedSize(40, 20)
        self.wave_timer = QTimer(self)
        self.wave_timer.timeout.connect(self.update)
        self.phase = 0
        self.is_listening = False

    def start_animation(self):
        self.is_listening = True
        self.wave_timer.start(50)

    def stop_animation(self):
        self.is_listening = False
        self.wave_timer.stop()
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        if self.is_listening:
            painter.setPen(QPen(QColor("#4CAF50"), 2))
            width = self.width() / 6
            for i in range(5):
                height = 10 * abs(np.sin(self.phase + i * 0.5))
                painter.drawLine(
                    int(width * i + width / 2), int(self.height() - height),
                    int(width * i + width / 2), self.height()
                )
            self.phase = (self.phase + 0.2) % (2 * np.pi)
        else:
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor("transparent"))
            painter.drawRect(self.rect())

class FilePreviewDialog(QDialog):
    """Dialog with file explorer and preview pane."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("File Manager with Preview")
        self.setGeometry(200, 200, 800, 600)
        layout = QHBoxLayout(self)

        # File explorer
        self.splitter = QSplitter(Qt.Horizontal)
        self.file_model = QFileSystemModel()
        self.file_model.setRootPath(QDir.homePath())
        self.file_tree = QTreeView()
        self.file_tree.setModel(self.file_model)
        self.file_tree.setRootIndex(self.file_model.index(QDir.homePath()))
        self.file_tree.selectionModel().selectionChanged.connect(self.show_preview)
        self.splitter.addWidget(self.file_tree)

        # Preview pane
        self.preview_widget = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_widget)
        self.preview_label = QLabel("Select a file to preview")
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_layout.addWidget(self.preview_label)
        self.preview_layout.addWidget(self.preview_text)
        self.preview_text.hide()
        self.splitter.addWidget(self.preview_widget)

        layout.addWidget(self.splitter)

        # Buttons
        button_layout = QHBoxLayout()
        self.open_btn = QPushButton("Open")
        self.open_btn.clicked.connect(self.open_selected)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.open_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

    def show_preview(self):
        """Show preview of selected file."""
        indexes = self.file_tree.selectedIndexes()
        if not indexes:
            return
        file_path = self.file_model.filePath(indexes[0])
        self.preview_label.setText(f"Preview: {os.path.basename(file_path)}")
        self.preview_text.hide()
        self.preview_label.setPixmap(QPixmap())
        if os.path.isfile(file_path):
            if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                pixmap = QPixmap(file_path).scaled(300, 300, Qt.KeepAspectRatio)
                self.preview_label.setPixmap(pixmap)
            elif file_path.lower().endswith(('.txt', '.py', '.json')):
                try:
                    with open(file_path, 'r') as f:
                        content = f.read()[:1000]  # Limit preview size
                    self.preview_text.setPlainText(content)
                    self.preview_text.show()
                except:
                    self.preview_text.setPlainText("Cannot preview this file")
                    self.preview_text.show()

    def open_selected(self):
        """Open the selected file."""
        indexes = self.file_tree.selectedIndexes()
        if indexes:
            file_path = self.file_model.filePath(indexes[0])
            self.accept()
            self.parent().open_file_or_folder(file_path, folder=os.path.dirname(file_path))

class VoiceAssistantGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=4096)
        self.reminders = []
        self.alarms = []
        self.tasks = []
        self.task_history = []  # For undo/redo
        self.timers = []
        self.music_playing = False
        self.music_file = None
        self.current_radio_station = None
        self.aliases = {}
        self.bookmarks = {}
        self.notes_file = os.path.join(os.path.expanduser("~"), "Documents", "assistant_notes.txt")
        self.config_file = os.path.join(os.path.expanduser("~"), "Documents", "assistant_config.json")
        self.current_theme = "Light"
        self.sidebar_position = "Left"
        self.plugins = {}
        self.command_history = []
        self.tray_icon = None
        self.minimized_to_tray = False

        self.initUI()
        self.init_tts()
        self.init_tray_icon()
        self.load_config()
        self.load_plugins()

        self.check_timer = QTimer(self)
        self.check_timer.timeout.connect(self.check_scheduled_events)
        self.check_timer.start(1000)

        # Register global hotkey
        try:
            keyboard.add_hotkey('ctrl+alt+v', self.start_listening)
        except:
            self.append_to_log("Failed to register hotkey", "Warning")

    def init_tts(self):
        """Initialize text-to-speech engine."""
        try:
            self.engine = pyttsx3.init(driverName=None)
            self.engine.setProperty('rate', 150)
            self.engine.setProperty('volume', 1.0)
            voices = self.engine.getProperty('voices')
            if voices:
                female_voices = [v for v in voices if "female" in v.name.lower()]
                self.engine.setProperty('voice', female_voices[0].id if female_voices else voices[0].id)
            self.speak("Voice assistant initialized. How can I help you?")
        except Exception as e:
            self.append_to_log(f"Failed to initialize TTS: {str(e)}", "Error")
            print("Error: Text-to-speech failed to initialize.")

    def init_tray_icon(self):
        
            if not QSystemTrayIcon.isSystemTrayAvailable():
                return
                
            self.tray_icon = QSystemTrayIcon(self)
            self.tray_icon.setIcon(self.get_icon('assistant.png'))
            
            # Create tray menu
            tray_menu = QMenu()
            
            show_action = QAction("Show", self)
            show_action.triggered.connect(self.show_normal)
            tray_menu.addAction(show_action)
            
            listen_action = QAction("Listen", self)
            listen_action.triggered.connect(self.start_listening)
            tray_menu.addAction(listen_action)
            
            exit_action = QAction("Exit", self)
            exit_action.triggered.connect(self.close)
            tray_menu.addAction(exit_action)
            
            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.show()
            
            # Handle tray icon clicks
            self.tray_icon.activated.connect(self.tray_icon_clicked)

    def tray_icon_clicked(self, reason):
        
            if reason == QSystemTrayIcon.Trigger:  # Single click
                self.show_normal()
    
    def show_normal(self):
        """Restore the window from minimized state."""
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.activateWindow()
        self.minimized_to_tray = False

    def initUI(self):
        """Initialize the enhanced user interface."""
        self.setWindowTitle('Voice Assistant Pro')
        self.setWindowIcon(self.get_icon('assistant.png'))
        self.setGeometry(100, 100, 1200, 800)

        # Central widget with tabbed interface
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Toolbar for main actions
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.toolbar.setIconSize(QSize(24, 24))
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        # Toolbar actions
        self.listen_action = QAction(self.get_icon('mic.png'), "Listen (Ctrl+L)", self)
        self.listen_action.setShortcut("Ctrl+L")
        self.listen_action.triggered.connect(self.start_listening)
        self.toolbar.addAction(self.listen_action)

        self.speak_action = QAction(self.get_icon('speak.png'), "Speak Text (Ctrl+S)", self)
        self.speak_action.setShortcut("Ctrl+S")
        self.speak_action.triggered.connect(self.speak_selected_text)
        self.toolbar.addAction(self.speak_action)

        self.stop_action = QAction(self.get_icon('stop.png'), "Stop (Ctrl+T)", self)
        self.stop_action.setShortcut("Ctrl+T")
        self.stop_action.triggered.connect(self.stop_media)
        self.toolbar.addAction(self.stop_action)

        self.toolbar.addSeparator()

        self.theme_action = QAction(self.get_icon('theme.png'), "Toggle Theme (Ctrl+M)", self)
        self.theme_action.setShortcut("Ctrl+M")
        self.theme_action.triggered.connect(self.toggle_theme)
        self.toolbar.addAction(self.theme_action)

        self.settings_action = QAction(self.get_icon('settings.png'), "Settings (Ctrl+P)", self)
        self.settings_action.setShortcut("Ctrl+P")
        self.settings_action.triggered.connect(self.show_settings)
        self.toolbar.addAction(self.settings_action)

        # Sidebar for quick commands
        self.sidebar = QDockWidget("Quick Commands", self)
        self.sidebar.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.sidebar.setMinimumWidth(150)
        self.sidebar_widget = QWidget()
        self.sidebar_layout = QVBoxLayout(self.sidebar_widget)
        self.sidebar_position = self.sidebar_position  # Will be loaded from config
        self.addDockWidget(Qt.LeftDockWidgetArea if self.sidebar_position == "Left" else Qt.RightDockWidgetArea, self.sidebar)

        # Quick command buttons
        quick_commands = [
            ("Time", 'time.png', self.get_time),
            ("Date", 'date.png', self.get_date),
            ("Weather", 'weather.png', lambda: self.process_command("what's the weather")),
            ("Joke", 'joke.png', lambda: self.process_command("tell me a joke")),
            ("Tasks", 'tasks.png', lambda: self.process_command("list tasks")),
            ("Notes", 'notes.png', self.read_notes),
            ("Screenshot", 'screenshot.png', self.take_screenshot),
            ("Bookmarks", 'bookmark.png', lambda: self.process_command("list bookmarks")),
            ("Files", 'files.png', self.show_file_manager)
        ]

        for text, icon, func in quick_commands:
            btn = QToolButton()
            btn.setText(text)
            btn.setIcon(self.get_icon(icon))
            btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
            btn.setFixedHeight(40)
            btn.clicked.connect(func)
            btn.setToolTip(text)
            btn.setAccessibleName(text)
            self.sidebar_layout.addWidget(btn)

        self.sidebar_layout.addStretch()
        self.sidebar.setWidget(self.sidebar_widget)

        # Command input panel
        self.command_panel = QWidget()
        self.command_layout = QHBoxLayout(self.command_panel)
        self.command_input = QComboBox()
        self.command_input.setEditable(True)
        self.command_input.setPlaceholderText("Type command or select from history...")
        self.command_input.lineEdit().returnPressed.connect(self.process_text_command)
        self.command_input.setFont(QFont("Arial", 12))
        self.command_input.setCompleter(None)
        self.command_input.lineEdit().textEdited.connect(self.suggest_commands)
        self.command_layout.addWidget(self.command_input)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setIcon(self.get_icon('clear.png'))
        self.clear_btn.clicked.connect(lambda: self.command_input.lineEdit().clear())
        self.command_layout.addWidget(self.clear_btn)

        self.toggle_command_btn = QPushButton("Hide")
        self.toggle_command_btn.setIcon(self.get_icon('toggle.png'))
        self.toggle_command_btn.clicked.connect(self.toggle_command_panel)
        self.command_layout.addWidget(self.toggle_command_btn)

        self.main_layout.addWidget(self.command_panel)

        # Tabbed interface
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        # Log tab
        self.log_widget = QWidget()
        self.log_layout = QVBoxLayout(self.log_widget)
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setFont(QFont("Consolas", 11))
        self.log_display.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log_display.customContextMenuRequested.connect(self.show_log_context_menu)
        self.log_layout.addWidget(self.log_display)
        self.tabs.addTab(self.log_widget, self.get_icon('log.png'), "Log")

        # Tasks tab
        self.tasks_widget = QWidget()
        self.tasks_layout = QVBoxLayout(self.tasks_widget)
        self.tasks_display = QTextEdit()
        self.tasks_display.setReadOnly(True)
        self.tasks_display.setFont(QFont("Consolas", 11))
        self.tasks_layout.addWidget(self.tasks_display)
        self.tasks_buttons = QHBoxLayout()
        self.undo_task_btn = QPushButton("Undo")
        self.undo_task_btn.setIcon(self.get_icon('undo.png'))
        self.undo_task_btn.clicked.connect(self.undo_task)
        self.redo_task_btn = QPushButton("Redo")
        self.redo_task_btn.setIcon(self.get_icon('redo.png'))
        self.redo_task_btn.clicked.connect(self.redo_task)
        self.tasks_buttons.addWidget(self.undo_task_btn)
        self.tasks_buttons.addWidget(self.redo_task_btn)
        self.tasks_layout.addLayout(self.tasks_buttons)
        self.tabs.addTab(self.tasks_widget, self.get_icon('tasks.png'), "Tasks")

        # Bookmarks tab
        self.bookmarks_widget = QWidget()
        self.bookmarks_layout = QVBoxLayout(self.bookmarks_widget)
        self.bookmarks_display = QTextEdit()
        self.bookmarks_display.setReadOnly(True)
        self.bookmarks_display.setFont(QFont("Consolas", 11))
        self.bookmarks_layout.addWidget(self.bookmarks_display)
        self.tabs.addTab(self.bookmarks_widget, self.get_icon('bookmark.png'), "Bookmarks")

        # Status bar
        self.status_bar = QStatusBar()
        self.status_label = QLabel("Ready")
        self.status_bar.addWidget(self.status_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(100)
        self.progress_bar.setVisible(False)
        self.status_bar.addWidget(self.progress_bar)
        self.system_info_label = QLabel("CPU: 0% | Mem: 0%")
        self.status_bar.addPermanentWidget(self.system_info_label)
        self.setStatusBar(self.status_bar)

        # Animated voice indicator
        self.voice_indicator = AnimatedVoiceIndicator()
        self.status_bar.addPermanentWidget(self.voice_indicator)

        # Timer for updating system info
        self.system_info_timer = QTimer(self)
        self.system_info_timer.timeout.connect(self.update_system_info)
        self.system_info_timer.start(5000)

        self.apply_styles()

    def get_icon(self, icon_name):
        """Return QIcon with fallback if icon file is missing."""
        icon_path = icon_name
        if os.path.exists(icon_path):
            return QIcon(icon_path)
        return QIcon()

    def apply_styles(self):
        """Apply modern UI styling with theme support."""
        if self.current_theme == "Dark":
            palette = QPalette()
            palette.setColor(QPalette.Window, QColor(30, 30, 30))
            palette.setColor(QPalette.WindowText, Qt.white)
            palette.setColor(QPalette.Base, QColor(40, 40, 40))
            palette.setColor(QPalette.AlternateBase, QColor(50, 50, 50))
            palette.setColor(QPalette.Text, Qt.white)
            palette.setColor(QPalette.Button, QColor(60, 60, 60))
            palette.setColor(QPalette.ButtonText, Qt.white)
            self.setPalette(palette)
            self.setStyleSheet("""
                QMainWindow { background-color: #1E1E1E; }
                QDockWidget { background-color: #252525; border: 1px solid #333; }
                QToolBar { background-color: #252525; border-bottom: 1px solid #333; }
                QToolButton {
                    background-color: #333333; color: white; border-radius: 4px;
                    padding: 8px; margin: 4px;
                }
                QToolButton:hover { background-color: #444444; }
                QTextEdit {
                    background-color: #2A2A2A; color: #FFFFFF; border: 1px solid #444;
                    font-family: Consolas; font-size: 11pt; padding: 4px;
                }
                QComboBox {
                    background-color: #2A2A2A; color: #FFFFFF; border: 1px solid #444;
                    border-radius: 4px; padding: 6px; font-size: 12pt;
                }
                QComboBox QAbstractItemView { background-color: #2A2A2A; color: #FFFFFF; }
                QPushButton {
                    background-color: #3A3A3A; color: white; border-radius: 4px;
                    padding: 6px 12px; margin: 4px;
                }
                QPushButton:hover { background-color: #4A4A4A; }
                QTabWidget::pane { border: 1px solid #333; }
                QTabBar::tab {
                    background-color: #252525; color: white; padding: 8px 16px;
                }
                QTabBar::tab:selected { background-color: #3A3A3A; }
                QStatusBar {
                    background-color: #252525; color: #BBBBBB; border-top: 1px solid #333;
                }
                QProgressBar {
                    border: 1px solid #444; background-color: #2A2A2A; color: #FFFFFF;
                    text-align: center; border-radius: 4px;
                }
                QProgressBar::chunk {
                    background-color: #4CAF50; border-radius: 4px;
                }
            """)
        else:
            palette = QPalette()
            palette.setColor(QPalette.Window, QColor(245, 245, 245))
            palette.setColor(QPalette.WindowText, Qt.black)
            palette.setColor(QPalette.Base, Qt.white)
            palette.setColor(QPalette.AlternateBase, QColor(230, 230, 230))
            palette.setColor(QPalette.Text, Qt.black)
            palette.setColor(QPalette.Button, QColor(220, 220, 220))
            palette.setColor(QPalette.ButtonText, Qt.black)
            self.setPalette(palette)
            self.setStyleSheet("""
                QMainWindow { background-color: #F5F5F5; }
                QDockWidget { background-color: #FFFFFF; border: 1px solid #DDD; }
                QToolBar { background-color: #FFFFFF; border-bottom: 1px solid #DDD; }
                QToolButton {
                    background-color: #E0E0E0; color: black; border-radius: 4px;
                    padding: 8px; margin: 4px;
                }
                QToolButton:hover { background-color: #D0D0D0; }
                QTextEdit {
                    background-color: #FFFFFF; color: #000000; border: 1px solid #BDBDBD;
                    font-family: Consolas; font-size: 11pt; padding: 4px;
                }
                QComboBox {
                    background-color: #FFFFFF; color: #000000; border: 1px solid #BDBDBD;
                    border-radius: 4px; padding: 6px; font-size: 12pt;
                }
                QComboBox QAbstractItemView { background-color: #FFFFFF; color: #000000; }
                QPushButton {
                    background-color: #E0E0E0; color: black; border-radius: 4px;
                    padding: 6px 12px; margin: 4px;
                }
                QPushButton:hover { background-color: #D0D0D0; }
                QTabWidget::pane { border: 1px solid #DDD; }
                QTabBar::tab {
                    background-color: #E0E0E0; color: black; padding: 8px 16px;
                }
                QTabBar::tab:selected { background-color: #FFFFFF; }
                QStatusBar {
                    background-color: #ECEFF1; color: #666666; border-top: 1px solid #BDBDBD;
                }
                QProgressBar {
                    border: 1px solid #BDBDBD; background-color: #FFFFFF; color: #000000;
                    text-align: center; border-radius: 4px;
                }
                QProgressBar::chunk {
                    background-color: #4CAF50; border-radius: 4px;
                }
            """)

    def toggle_theme(self):
        """Toggle between light and dark themes."""
        self.current_theme = "Dark" if self.current_theme == "Light" else "Light"
        self.apply_styles()
        self.append_to_log(f"Switched to {self.current_theme} theme", "System")
        self.speak(f"Switched to {self.current_theme} theme")

    def toggle_command_panel(self):
        """Show or hide the command input panel."""
        if self.command_panel.isVisible():
            self.command_panel.hide()
            self.toggle_command_btn.setText("Show")
            self.toggle_command_btn.setIcon(self.get_icon('toggle_show.png'))
        else:
            self.command_panel.show()
            self.toggle_command_btn.setText("Hide")
            self.toggle_command_btn.setIcon(self.get_icon('toggle_hide.png'))

    def update_system_info(self):
        """Update system resource info in the status bar."""
        try:
            cpu_usage = psutil.cpu_percent(interval=0.1)
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            self.system_info_label.setText(f"CPU: {cpu_usage:.1f}% | Mem: {memory_percent:.1f}%")
        except Exception as e:
            self.append_to_log(f"Failed to update system info: {str(e)}", "Error")

    def load_config(self):
        """Load configuration from file."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    config = json.load(f)
                    self.aliases = config.get("aliases", {})
                    self.bookmarks = config.get("bookmarks", {})
                    self.current_theme = config.get("theme", "Light")
                    self.sidebar_position = config.get("sidebar_position", "Left")
                    self.command_history = config.get("command_history", [])[:20]
                    if hasattr(self, 'engine') and self.engine:
                        self.engine.setProperty('rate', config.get("tts_rate", 150))
                        self.engine.setProperty('volume', config.get("tts_volume", 1.0))
                        if "tts_voice" in config:
                            voices = self.engine.getProperty('voices')
                            for voice in voices:
                                if voice.id == config["tts_voice"]:
                                    self.engine.setProperty('voice', voice.id)
                                    break
                self.apply_styles()
                self.update_command_history()
                self.update_sidebar_position()
        except Exception as e:
            self.append_to_log(f"Failed to load config: {str(e)}", "Error")

    def save_config(self):
        """Save configuration to file."""
        config = {
            "aliases": self.aliases,
            "bookmarks": self.bookmarks,
            "theme": self.current_theme,
            "sidebar_position": self.sidebar_position,
            "command_history": self.command_history[-20:],
            "tts_rate": self.engine.getProperty('rate') if hasattr(self, 'engine') and self.engine else 150,
            "tts_volume": self.engine.getProperty('volume') if hasattr(self, 'engine') and self.engine else 1.0,
            "tts_voice": self.engine.getProperty('voice') if hasattr(self, 'engine') and self.engine else ""
        }
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, "w") as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            self.append_to_log(f"Failed to save config: {str(e)}", "Error")

    def export_config(self):
        """Export configuration to a file."""
        file_path, _ = QFileDialog.getSaveFileName(self, "Export Config", "", "JSON Files (*.json);;All Files (*)")
        if file_path:
            try:
                with open(file_path, "w") as f:
                    json.dump({
                        "aliases": self.aliases,
                        "bookmarks": self.bookmarks,
                        "command_history": self.command_history
                    }, f, indent=4)
                self.append_to_log(f"Configuration exported to {file_path}", "System")
                self.speak("Configuration exported successfully")
            except Exception as e:
                self.append_to_log(f"Failed to export config: {str(e)}", "Error")
                self.speak("Failed to export configuration")

    def import_config(self):
        """Import configuration from a file."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Import Config", "", "JSON Files (*.json);;All Files (*)")
        if file_path:
            try:
                with open(file_path, "r") as f:
                    config = json.load(f)
                    self.aliases.update(config.get("aliases", {}))
                    self.bookmarks.update(config.get("bookmarks", {}))
                    self.command_history.extend(config.get("command_history", []))
                    self.command_history = self.command_history[-20:]
                    self.save_config()
                    self.update_command_history()
                    self.list_bookmarks()
                    self.append_to_log(f"Configuration imported from {file_path}", "System")
                    self.speak("Configuration imported successfully")
            except Exception as e:
                self.append_to_log(f"Failed to import config: {str(e)}", "Error")
                self.speak("Failed to import configuration")

    def update_command_history(self):
        """Update the command input dropdown with history."""
        self.command_input.clear()
        self.command_input.addItems(self.command_history)

    def suggest_commands(self, text):
        """Suggest commands based on partial input."""
        if not text:
            return
        suggestions = [cmd for cmd in self.command_history if text.lower() in cmd.lower()]
        self.command_input.clear()
        self.command_input.addItems(suggestions)
        self.command_input.lineEdit().setText(text)

    def append_to_log(self, text, speaker="System"):
        """Add text to log display with timestamp."""
        timestamp = dt.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {speaker}: {text}"
        self.log_display.append(log_entry)
        self.log_display.moveCursor(QTextCursor.End)
        if speaker == "Error":
            with open("error_log.txt", "a") as f:
                f.write(f"{log_entry}\n")

    def speak(self, text):
        """Convert text to speech and log it."""
        if not hasattr(self, 'engine') or not self.engine:
            self.append_to_log("TTS engine not available", "Error")
            return
        try:
            self.append_to_log(text, "Assistant")
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            self.append_to_log(f"Speech error: {str(e)}", "Error")

    def start_listening(self):
        """Start listening for voice commands."""
        if hasattr(self, 'voice_thread') and self.voice_thread and self.voice_thread.is_listening:
            return
        self.voice_thread = VoiceThread()
        self.voice_thread.finished_signal.connect(self.process_voice_command)
        self.voice_thread.listening_signal.connect(self.update_listening_status)
        self.voice_thread.error_signal.connect(self.handle_voice_error)
        self.voice_thread.start()

    def update_listening_status(self, is_listening):
        """Update UI based on listening state with animated indicator."""
        if is_listening:
            self.status_label.setText("Listening... Speak now")
            self.listen_action.setEnabled(False)
            self.voice_indicator.start_animation()
            self.append_to_log("Listening for voice command...", "System")
        else:
            self.status_label.setText("Ready")
            self.listen_action.setEnabled(True)
            self.voice_indicator.stop_animation()

    def handle_voice_error(self, error):
        """Handle errors from voice thread."""
        self.append_to_log(error, "Error")
        self.speak(error)
        self.update_listening_status(False)

    def process_voice_command(self, command):
        """Process a command from voice input."""
        self.append_to_log(command, "You")
        self.command_input.lineEdit().setText(command)
        self.command_history.append(command)
        self.command_history = self.command_history[-20:]
        self.update_command_history()
        self.process_command(command)

    def process_text_command(self):
        """Process a command from text input."""
        command = self.command_input.currentText().strip()
        if command:
            self.append_to_log(command, "You")
            self.command_history.append(command)
            self.command_history = self.command_history[-20:]
            self.update_command_history()
            self.process_command(command)
            self.command_input.lineEdit().clear()

    def speak_selected_text(self):
        """Speak the currently selected text."""
        selected_text = self.log_display.textCursor().selectedText()
        if selected_text:
            self.speak(selected_text)
        else:
            self.speak("No text selected")

    def stop_media(self):
        """Stop any currently playing media."""
        if pygame.mixer.music.get_busy():
            pygame.mixer.music.stop()
            self.music_playing = False
            self.append_to_log("Media stopped", "System")
            self.speak("Media stopped")
        else:
            self.speak("No media is currently playing")

    def show_log_context_menu(self, position):
        """Show context menu for log display with search option."""
        menu = self.log_display.createStandardContextMenu()
        clear_action = menu.addAction("Clear Log")
        clear_action.triggered.connect(self.log_display.clear)
        save_action = menu.addAction("Save Log...")
        save_action.triggered.connect(self.save_log_to_file)
        search_action = menu.addAction("Search Log...")
        search_action.triggered.connect(self.search_log)
        menu.exec_(self.log_display.mapToGlobal(position))

    def save_log_to_file(self):
        """Save log content to a file."""
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Log", "", "Text Files (*.txt);;All Files (*)")
        if file_path:
            try:
                with open(file_path, "w") as f:
                    f.write(self.log_display.toPlainText())
                self.append_to_log(f"Log saved to {file_path}", "System")
            except Exception as e:
                self.append_to_log(f"Failed to save log: {str(e)}", "Error")

    def search_log(self):
        """Search through the log display."""
        search_dialog = QInputDialog(self)
        search_dialog.setWindowTitle("Search Log")
        search_dialog.setLabelText("Enter search term:")
        if search_dialog.exec_():
            term = search_dialog.textValue()
            if term:
                self.log_display.find(term, QJsonDocument.FindCaseSensitively)
                self.append_to_log(f"Searched log for: {term}", "System")

    def get_time(self):
        """Get current time."""
        current_time = dt.now().strftime("%I:%M %p")
        self.append_to_log(f"The current time is {current_time}", "Assistant")
        self.speak(f"The current time is {current_time}")

    def get_date(self):
        """Get current date."""
        current_date = dt.now().strftime("%B %d, %Y")
        self.append_to_log(f"Today's date is {current_date}", "Assistant")
        self.speak(f"Today's date is {current_date}")

    def show_file_manager(self):
        """Show a file manager dialog with preview."""
        dialog = FilePreviewDialog(self)
        dialog.exec_()

    def create_file(self, filename, content=None, folder="Desktop"):
        """Create a file in a separate thread."""
        def perform_operation():
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                base_folder = os.path.join(os.path.expanduser("~"), folder.capitalize())
                os.makedirs(base_folder, exist_ok=True)
                filepath = os.path.join(base_folder, filename)
                if os.path.exists(filepath):
                    self.speak(f"File {filename} already exists.")
                    return False
                for i in range(1, 101):
                    time.sleep(0.01)  # Simulate work
                    self.progress_bar.setValue(i)
                with open(filepath, "w") as f:
                    f.write(content or "")
                self.speak(f"File {filename} created successfully in {folder}.")
                return True
            except Exception as e:
                self.append_to_log(f"Failed to create file: {str(e)}", "Error")
                self.speak("Failed to create file. See log for details.")
                return False
            finally:
                self.progress_bar.setVisible(False)

        threading.Thread(target=perform_operation, daemon=True).start()

    def create_folder(self, foldername, folder="Desktop"):
        """Create a folder with proper error handling."""
        try:
            base_folder = os.path.join(os.path.expanduser("~"), folder.capitalize())
            os.makedirs(base_folder, exist_ok=True)
            folderpath = os.path.join(base_folder, foldername)
            if os.path.exists(folderpath):
                self.speak(f"Folder {foldername} already exists.")
                return False
            os.makedirs(folderpath)
            self.speak(f"Folder {foldername} created successfully in {folder}.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to create folder: {str(e)}", "Error")
            self.speak("Failed to create folder. See log for details.")
            return False

    def delete_file(self, filename, folder="Desktop"):
        """Delete a file with confirmation."""
        try:
            base_folder = os.path.join(os.path.expanduser("~"), folder.capitalize())
            filepath = os.path.join(base_folder, filename)
            if not os.path.exists(filepath):
                self.speak(f"File {filename} not found in {folder}.")
                return False
            confirm = QMessageBox.question(
                self, "Confirm Delete",
                f"Are you sure you want to delete {filename}?",
                QMessageBox.Yes | QMessageBox.No
            )
            if confirm == QMessageBox.Yes:
                os.remove(filepath)
                self.speak(f"File {filename} deleted from {folder}.")
                return True
            self.speak("File deletion canceled.")
            return False
        except Exception as e:
            self.append_to_log(f"Failed to delete file: {str(e)}", "Error")
            self.speak("Failed to delete file. See log for details.")
            return False

    def delete_folder(self, foldername, folder="Desktop"):
        """Delete a folder with confirmation."""
        try:
            base_folder = os.path.join(os.path.expanduser("~"), folder.capitalize())
            folderpath = os.path.join(base_folder, foldername)
            if not os.path.exists(folderpath):
                self.speak(f"Folder {foldername} not found in {folder}.")
                return False
            confirm = QMessageBox.question(
                self, "Confirm Delete",
                f"Are you sure you want to delete {foldername}?",
                QMessageBox.Yes | QMessageBox.No
            )
            if confirm == QMessageBox.Yes:
                import shutil
                shutil.rmtree(folderpath)
                self.speak(f"Folder {foldername} deleted from {folder}.")
                return True
            self.speak("Folder deletion canceled.")
            return False
        except Exception as e:
            self.append_to_log(f"Failed to delete folder: {str(e)}", "Error")
            self.speak("Failed to delete folder. See log for details.")
            return False

    def open_file_or_folder(self, name, folder="Desktop"):
        """Open a file or folder."""
        try:
            base_folder = folder if os.path.isabs(folder) else os.path.join(os.path.expanduser("~"), folder.capitalize())
            path = os.path.join(base_folder, name) if not os.path.isabs(name) else name
            if not os.path.exists(path):
                self.speak(f"{name} not found in {folder}.")
                return False
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
            self.speak(f"Opening {name}.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to open {name}: {str(e)}", "Error")
            self.speak(f"Failed to open {name}. See log for details.")
            return False

    def open_application(self, app_name):
        """Open a desktop application."""
        app_commands = {
            "notepad": "notepad" if platform.system() == "Windows" else "gedit",
            "calculator": "calc" if platform.system() == "Windows" else "gnome-calculator",
            "browser": "start chrome" if platform.system() == "Windows" else "firefox",
            "terminal": "cmd" if platform.system() == "Windows" else "gnome-terminal"
        }
        if app_name in app_commands:
            try:
                if platform.system() == "Windows":
                    subprocess.run(app_commands[app_name], shell=True)
                else:
                    subprocess.run([app_commands[app_name]])
                self.speak(f"Opening {app_name}.")
                return True
            except Exception as e:
                self.append_to_log(f"Failed to open {app_name}: {str(e)}", "Error")
                self.speak(f"Failed to open {app_name}. See log for details.")
                return False
        self.speak(f"Application {app_name} not recognized. Try notepad, calculator, browser, or terminal.")
        return False

    def take_screenshot(self):
        """Take and save a screenshot."""
        try:
            screenshots_dir = os.path.join(os.path.expanduser("~"), "Pictures", "Screenshots")
            os.makedirs(screenshots_dir, exist_ok=True)
            timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(screenshots_dir, f"screenshot_{timestamp}.png")
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)
            self.speak(f"Screenshot saved as screenshot_{timestamp}.png in your Pictures folder.")
            notification.notify(
                title="Screenshot Taken",
                message=f"Saved as screenshot_{timestamp}.png",
                timeout=5
            )
            return True
        except Exception as e:
            self.append_to_log(f"Failed to take screenshot: {str(e)}", "Error")
            self.speak("Failed to take screenshot. See log for details.")
            return False

    def get_system_resources(self):
        """Get detailed system resource information."""
        try:
            cpu_usage = psutil.cpu_percent(interval=1)
            cpu_count = psutil.cpu_count()
            cpu_freq = psutil.cpu_freq().current if hasattr(psutil.cpu_freq(), 'current') else "N/A"
            memory = psutil.virtual_memory()
            memory_total = memory.total / (1024**3)
            memory_used = memory.used / (1024**3)
            memory_percent = memory.percent
            disk = psutil.disk_usage('/')
            disk_total = disk.total / (1024**3)
            disk_used = disk.used / (1024**3)
            disk_percent = disk.percent
            battery = psutil.sensors_battery()
            battery_info = ""
            if battery:
                battery_percent = battery.percent
                battery_plugged = "plugged in" if battery.power_plugged else "not plugged in"
                battery_info = f"Battery: {battery_percent}% ({battery_plugged}). "
            message = (
                f"System resources: CPU: {cpu_usage}% on {cpu_count} cores ({cpu_freq} MHz). "
                f"Memory: {memory_used:.1f}/{memory_total:.1f} GB ({memory_percent}%). "
                f"Disk: {disk_used:.1f}/{disk_total:.1f} GB ({disk_percent}%). {battery_info}"
            )
            self.speak(message)
            return True
        except Exception as e:
            self.append_to_log(f"Failed to get system resources: {str(e)}", "Error")
            self.speak("Failed to get system resources. See log for details.")
            return False

    def manage_clipboard(self, action, content=None):
        """Read or set clipboard content."""
        try:
            if action == "read":
                content = pyperclip.paste()
                self.speak(f"Clipboard content: {content}")
                return True
            elif action == "set" and content:
                pyperclip.copy(content)
                self.speak(f"Clipboard set to: {content}")
                return True
            self.speak("Please specify clipboard content to set.")
            return False
        except Exception as e:
            self.append_to_log(f"Clipboard error: {str(e)}", "Error")
            self.speak("Failed to manage clipboard. See log for details.")
            return False

    def add_bookmark(self, name, url):
        """Add a bookmark with validation."""
        def perform_operation():
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                if not url.startswith(('http://', 'https://')):
                    url = 'https://' + url
                response = requests.get(url, timeout=5)
                for i in range(1, 101):
                    time.sleep(0.01)
                    self.progress_bar.setValue(i)
                if response.status_code >= 400:
                    self.speak(f"Warning: Website returned status code {response.status_code}")
                self.bookmarks[name.lower()] = url
                self.save_config()
                self.speak(f"Bookmark '{name}' added for {url}.")
                self.list_bookmarks()
                return True
            except requests.RequestException:
                self.speak("Warning: Could not verify website availability")
                self.bookmarks[name.lower()] = url
                self.save_config()
                self.speak(f"Bookmark '{name}' added for {url}.")
                self.list_bookmarks()
                return True
            except Exception as e:
                self.append_to_log(f"Failed to add bookmark: {str(e)}", "Error")
                self.speak("Failed to add bookmark. See log for details.")
                return False
            finally:
                self.progress_bar.setVisible(False)

        threading.Thread(target=perform_operation, daemon=True).start()

    def open_bookmark(self, name):
        """Open a bookmark with fuzzy matching."""
        name = name.lower()
        if name in self.bookmarks:
            webbrowser.open(self.bookmarks[name])
            self.speak(f"Opening bookmark {name}.")
            return True
        matches = [k for k in self.bookmarks.keys() if name in k]
        if matches:
            webbrowser.open(self.bookmarks[matches[0]])
            self.speak(f"Opening bookmark {matches[0]}.")
            return True
        self.speak(f"Bookmark {name} not found.")
        return False

    def list_bookmarks(self):
        """List all bookmarks in the bookmarks tab."""
        if not self.bookmarks:
            self.bookmarks_display.setText("No bookmarks found.")
            self.speak("No bookmarks found.")
        else:
            bookmark_list = "\n".join(f"{name}: {url}" for name, url in self.bookmarks.items())
            self.bookmarks_display.setText(f"Bookmarks:\n{bookmark_list}")
            self.speak(f"You have {len(self.bookmarks)} bookmarks.")

    def open_new_tab(self, url=None):
        """Open a new tab in the default browser."""
        try:
            url = url or "https://www.google.com"
            webbrowser.open_new_tab(url)
            self.speak(f"Opened new tab with {url}.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to open new tab: {str(e)}", "Error")
            self.speak("Failed to open new tab. See log for details.")
            return False

    def close_browser_tab(self):
        """Close the current browser tab."""
        try:
            pyautogui.hotkey("ctrl", "w")
            self.speak("Closed current browser tab.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to close tab: {str(e)}", "Error")
            self.speak("Failed to close tab. See log for details.")
            return False

    def open_incognito_mode(self, url=None):
        """Open browser in incognito mode."""
        try:
            url = url or "https://www.google.com"
            if platform.system() == "Windows":
                subprocess.run(["start", "chrome", "--incognito", url], shell=True)
            elif platform.system() == "Linux":
                subprocess.run(["firefox", "--private-window", url])
            elif platform.system() == "Darwin":
                subprocess.run(["open", "-a", "Safari", "--args", "--private", url])
            self.speak(f"Opened browser in incognito mode with {url}.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to open incognito mode: {str(e)}", "Error")
            self.speak("Failed to open incognito mode. See log for details.")
            return False

    def scrape_website(self, url):
        """Scrape basic information from a website in a thread."""
        def perform_operation():
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                headers = {'User-Agent': 'Mozilla/5.0'}
                response = requests.get(url, headers=headers)
                for i in range(1, 101):
                    time.sleep(0.01)
                    self.progress_bar.setValue(i)
                soup = BeautifulSoup(response.text, 'html.parser')
                title = soup.title.string if soup.title else "No title found"
                self.speak(f"The title of the website is: {title}")
                return True
            except Exception as e:
                self.append_to_log(f"Failed to scrape website: {str(e)}", "Error")
                self.speak("Failed to scrape website. See log for details.")
                return False
            finally:
                self.progress_bar.setVisible(False)

        threading.Thread(target=perform_operation, daemon=True).start()

    def autofill_form(self, website, query):
        """Simulate filling out a search form on a website."""
        try:
            webbrowser.open(website)
            time.sleep(3)
            pyautogui.write(query)
            pyautogui.press("enter")
            self.speak(f"Filled search form on {website} with {query}.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to autofill form: {str(e)}", "Error")
            self.speak("Failed to autofill form. See log for details.")
            return False

    def show_settings(self):
        """Show settings dialog with enhanced options."""
        settings_dialog = QMainWindow(self)
        settings_dialog.setWindowTitle("Assistant Settings")
        settings_dialog.setGeometry(200, 200, 500, 500)
        settings_widget = QWidget()
        settings_layout = QVBoxLayout()
        settings_widget.setLayout(settings_layout)

        # TTS Settings
        tts_group = QWidget()
        tts_layout = QVBoxLayout()
        tts_group.setLayout(tts_layout)
        tts_label = QLabel("Text-to-Speech Settings")
        tts_label.setStyleSheet("font-weight: bold; font-size: 14pt;")
        tts_layout.addWidget(tts_label)
        rate_label = QLabel("Speech Rate:")
        rate_slider = QSlider(Qt.Horizontal)
        rate_slider.setRange(100, 250)
        rate_slider.setValue(self.engine.getProperty('rate'))
        rate_slider.valueChanged.connect(lambda v: self.engine.setProperty('rate', v))
        tts_layout.addWidget(rate_label)
        tts_layout.addWidget(rate_slider)
        volume_label = QLabel("Speech Volume:")
        volume_slider = QSlider(Qt.Horizontal)
        volume_slider.setRange(0, 100)
        volume_slider.setValue(int(self.engine.getProperty('volume') * 100))
        volume_slider.valueChanged.connect(lambda v: self.engine.setProperty('volume', v/100))
        tts_layout.addWidget(volume_label)
        tts_layout.addWidget(volume_slider)
        voice_label = QLabel("Voice:")
        voice_combo = QComboBox()
        voices = self.engine.getProperty('voices')
        current_voice = self.engine.getProperty('voice')
        for i, voice in enumerate(voices):
            voice_combo.addItem(voice.name)
            if voice.id == current_voice:
                voice_combo.setCurrentIndex(i)
        voice_combo.currentIndexChanged.connect(lambda i: self.engine.setProperty('voice', voices[i].id))
        tts_layout.addWidget(voice_label)
        tts_layout.addWidget(voice_combo)
        language_label = QLabel("TTS Language:")
        language_combo = QComboBox()
        language_combo.addItems(["en", "es", "fr", "de"])
        language_combo.currentTextChanged.connect(self.set_tts_language)
        tts_layout.addWidget(language_label)
        tts_layout.addWidget(language_combo)
        settings_layout.addWidget(tts_group)

        # Font Size
        font_group = QWidget()
        font_layout = QVBoxLayout()
        font_group.setLayout(font_layout)
        font_label = QLabel("Interface Font Size")
        font_label.setStyleSheet("font-weight: bold; font-size: 14pt;")
        font_layout.addWidget(font_label)
        font_size_combo = QComboBox()
        font_size_combo.addItems(["10", "12", "14", "16"])
        font_size_combo.setCurrentText("12")
        font_size_combo.currentTextChanged.connect(self.change_font_size)
        font_layout.addWidget(font_size_combo)
        settings_layout.addWidget(font_group)

        # Sidebar Position
        sidebar_group = QWidget()
        sidebar_layout = QVBoxLayout()
        sidebar_group.setLayout(sidebar_layout)
        sidebar_label = QLabel("Sidebar Position")
        sidebar_label.setStyleSheet("font-weight: bold; font-size: 14pt;")
        sidebar_combo = QComboBox()
        sidebar_combo.addItems(["Left", "Right"])
        sidebar_combo.setCurrentText(self.sidebar_position)
        sidebar_combo.currentTextChanged.connect(self.update_sidebar_position)
        sidebar_layout.addWidget(sidebar_label)
        sidebar_layout.addWidget(sidebar_combo)
        settings_layout.addWidget(sidebar_group)

        # Export/Import
        config_group = QWidget()
        config_layout = QHBoxLayout()
        config_group.setLayout(config_layout)
        export_btn = QPushButton("Export Config")
        export_btn.setIcon(self.get_icon('export.png'))
        export_btn.clicked.connect(self.export_config)
        import_btn = QPushButton("Import Config")
        import_btn.setIcon(self.get_icon('import.png'))
        import_btn.clicked.connect(self.import_config)
        config_layout.addWidget(export_btn)
        config_layout.addWidget(import_btn)
        settings_layout.addWidget(config_group)

        # Save button
        save_btn = QPushButton("Save Settings")
        save_btn.setIcon(self.get_icon('save.png'))
        save_btn.clicked.connect(self.save_config)
        save_btn.clicked.connect(settings_dialog.close)
        settings_layout.addWidget(save_btn)

        settings_dialog.setCentralWidget(settings_widget)
        settings_dialog.show()

    def set_tts_language(self, lang):
        """Set TTS language (mock implementation)."""
        self.append_to_log(f"Set TTS language to {lang}", "System")
        self.speak(f"Text-to-speech language set to {lang}. Note: Actual language switching requires additional TTS engine support.")

    def update_sidebar_position(self, position=None):
        """Update sidebar docking position."""
        if position:
            self.sidebar_position = position
        self.removeDockWidget(self.sidebar)
        self.addDockWidget(Qt.LeftDockWidgetArea if self.sidebar_position == "Left" else Qt.RightDockWidgetArea, self.sidebar)
        self.append_to_log(f"Sidebar moved to {self.sidebar_position}", "System")

    def change_font_size(self, size):
        """Change the font size of the UI."""
        font = QFont("Arial", int(size))
        self.command_input.setFont(font)
        self.log_display.setFont(QFont("Consolas", int(size)))
        self.tasks_display.setFont(QFont("Consolas", int(size)))
        self.bookmarks_display.setFont(QFont("Consolas", int(size)))
        self.append_to_log(f"Changed font size to {size}pt", "System")

    def load_plugins(self):
        """Load plugins from the plugins directory."""
        plugins_dir = os.path.join(os.path.dirname(__file__), "plugins")
        if not os.path.exists(plugins_dir):
            return
        for filename in os.listdir(plugins_dir):
            if filename.endswith(".py"):
                plugin_path = os.path.join(plugins_dir, filename)
                spec = importlib.util.spec_from_file_location(filename[:-3], plugin_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                if hasattr(module, "register_plugin"):
                    try:
                        self.plugins[filename[:-3]] = module.register_plugin(self)
                        self.append_to_log(f"Loaded plugin: {filename}", "System")
                    except Exception as e:
                        self.append_to_log(f"Failed to load plugin {filename}: {str(e)}", "Error")

    def load_reminders_and_tasks(self):
        """Load reminders and tasks from JSON file."""
        try:
            with open("reminders_tasks.json", "r") as f:
                data = json.load(f)
                self.reminders = [
                    (dt.strptime(r["time"], "%Y-%m-%d %H:%M"), r["message"])
                    for r in data.get("reminders", [])
                ]
                self.tasks = data.get("tasks", [])
        except FileNotFoundError:
            self.reminders = []
            self.tasks = []

    def save_reminders_and_tasks(self):
        """Save reminders and tasks to JSON file."""
        data = {
            "reminders": [
                {"time": r[0].strftime("%Y-%m-%d %H:%M"), "message": r[1]}
                for r in self.reminders
            ],
            "tasks": self.tasks
        }
        try:
            with open("reminders_tasks.json", "w") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            self.append_to_log(f"Failed to save reminders/tasks: {str(e)}", "Error")

    def add_reminder(self, time_str, message):
        """Add a reminder."""
        try:
            reminder_time = dt.strptime(time_str, "%H:%M %d-%m-%Y")
            self.reminders.append((reminder_time, message))
            self.save_reminders_and_tasks()
            self.speak(f"Reminder set for {message} at {time_str}")
            notification.notify(
                title="Reminder Set",
                message=f"{message} at {time_str}",
                timeout=5
            )
            return True
        except ValueError:
            self.speak("Invalid time format. Use HH:MM DD-MM-YYYY")
            return False

    def add_task(self, task):
        """Add a task with undo support."""
        try:
            self.tasks.append(task)
            self.task_history.append(("add", task))
            self.save_reminders_and_tasks()
            self.speak(f"Task added: {task}")
            self.list_tasks()
            return True
        except Exception as e:
            self.append_to_log(f"Failed to add task: {str(e)}", "Error")
            return False

    def undo_task(self):
        """Undo the last task operation."""
        if not self.task_history:
            self.speak("No task operations to undo")
            return
        operation, task = self.task_history.pop()
        if operation == "add":
            self.tasks.remove(task)
            self.save_reminders_and_tasks()
            self.speak(f"Undid adding task: {task}")
            self.list_tasks()
        self.task_history.append(("remove", task))

    def redo_task(self):
        """Redo the last undone task operation."""
        if not self.task_history or self.task_history[-1][0] != "remove":
            self.speak("No task operations to redo")
            return
        operation, task = self.task_history.pop()
        self.tasks.append(task)
        self.save_reminders_and_tasks()
        self.speak(f"Redid adding task: {task}")
        self.list_tasks()
        self.task_history.append(("add", task))

    def list_tasks(self):
        """List all tasks in the tasks tab."""
        if not self.tasks:
            self.tasks_display.setText("No tasks set.")
            self.speak("No tasks set.")
        else:
            task_list = "\n".join(f"{i+1}. {task}" for i, task in enumerate(self.tasks))
            self.tasks_display.setText(f"Tasks:\n{task_list}")
            self.speak(f"You have {len(self.tasks)} tasks.")

    def get_weather(self, city=None):
        """Get weather for a city or default location."""
        api_key = "YOUR_OPENWEATHERMAP_API_KEY"
        if not city:
            city = "London"
        try:
            url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
            response = requests.get(url).json()
            if response.get("cod") == 200:
                weather = response["weather"][0]["description"]
                temp = response["main"]["temp"]
                self.speak(f"In {city}, it's {weather} with a temperature of {temp} degrees Celsius.")
                return True
            self.speak("Could not fetch weather data.")
            return False
        except Exception as e:
            self.append_to_log(f"Weather error: {str(e)}", "Error")
            self.speak("Failed to fetch weather.")
            return False

    def send_email(self, recipient, subject, body):
        """Send an email."""
        sender = "your_email@gmail.com"
        password = "your_app_password"
        try:
            msg = MIMEText(body)
            msg["Subject"] = subject
            msg["From"] = sender
            msg["To"] = recipient
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(sender, password)
                server.sendmail(sender, recipient, msg.as_string())
            self.speak(f"Email sent to {recipient}")
            return True
        except Exception as e:
            self.append_to_log(f"Email error: {str(e)}", "Error")
            self.speak("Failed to send email.")
            return False

    def save_note(self, note):
        """Save a note to a file."""
        try:
            timestamp = dt.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.notes_file, "a") as f:
                f.write(f"[{timestamp}] {note}\n")
            self.speak("Note saved.")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to save note: {str(e)}", "Error")
            self.speak("Failed to save note.")
            return False

    def read_notes(self):
        """Read all notes."""
        try:
            with open(self.notes_file, "r") as f:
                notes = f.read()
            if notes:
                self.append_to_log(f"Notes:\n{notes}", "Assistant")
                self.speak("Here are your notes.")
            else:
                self.speak("No notes found.")
            return True
        except FileNotFoundError:
            self.speak("No notes found.")
            return False
        except Exception as e:
            self.append_to_log(f"Failed to read notes: {str(e)}", "Error")
            self.speak("Failed to read notes.")
            return False

    def add_calendar_event(self, title, date_str):
        """Add a calendar event (mock implementation)."""
        try:
            event_date = dt.strptime(date_str, "%d-%m-%Y %H:%M")
            self.speak(f"Event '{title}' scheduled for {event_date}. (Note: Full Google Calendar integration requires API setup.)")
            return True
        except ValueError:
            self.speak("Invalid date format. Use DD-MM-YYYY HH:MM")
            return False

    def add_alias(self, alias, command):
        """Add a command alias."""
        try:
            self.aliases[alias.lower()] = command
            self.save_config()
            self.speak(f"Alias '{alias}' set for command '{command}'")
            return True
        except Exception as e:
            self.append_to_log(f"Failed to add alias: {str(e)}", "Error")
            self.speak("Failed to add alias.")
            return False

    def get_battery_status(self):
        """Get battery status."""
        try:
            battery = psutil.sensors_battery()
            if battery:
                percent = battery.percent
                plugged = "plugged in" if battery.power_plugged else "not plugged in"
                self.speak(f"Battery is at {percent}% and {plugged}.")
                return True
            self.speak("Battery status not available.")
            return False
        except Exception as e:
            self.append_to_log(f"Battery status error: {str(e)}", "Error")
            self.speak("Failed to get battery status.")
            return False

    def check_scheduled_events(self):
        """Check alarms, timers, and reminders."""
        current_time = dt.now()
        for reminder in self.reminders[:]:
            if reminder[0] <= current_time:
                self.append_to_log(f"Reminder: {reminder[1]}", "System")
                self.speak(f"Reminder: {reminder[1]}")
                notification.notify(
                    title="Reminder",
                    message=reminder[1],
                    timeout=10
                )
                self.reminders.remove(reminder)
                self.save_reminders_and_tasks()
        self.check_alarms()
        self.check_timers()

    def check_alarms(self):
        """Check if any alarms match the current time."""
        current_time = dt.now().strftime("%H:%M")
        for alarm in self.alarms[:]:
            alarm_str = alarm.strftime("%H:%M")
            if alarm_str == current_time:
                self.append_to_log("Alarm! It's time!", "System")
                self.speak("Alarm! It's time!")
                notification.notify(
                    title="Alarm",
                    message="It's time!",
                    timeout=10
                )
                self.alarms.remove(alarm)

    def check_timers(self):
        """Check if any timers have expired."""
        current_time = time.time()
        for timer in self.timers[:]:
            end_time, duration = timer
            if current_time >= end_time:
                self.append_to_log(f"Timer for {duration} is up!", "System")
                self.speak(f"Timer for {duration} is up!")
                notification.notify(
                    title="Timer",
                    message=f"Timer for {duration} is up!",
                    timeout=10
                )
                self.timers.remove(timer)

    def process_command(self, command):
        """Main command processing method."""
        if not command:
            return

        command = command.lower().strip()

        if command in self.aliases:
            command = self.aliases[command]
            self.append_to_log(f"Using alias: {command}", "System")

        # Check for plugin commands
        for plugin_name, plugin in self.plugins.items():
            if command.startswith(plugin_name):
                try:
                    plugin["execute"](command[len(plugin_name):].strip())
                    return
                except Exception as e:
                    self.append_to_log(f"Plugin {plugin_name} error: {str(e)}", "Error")
                    return

        # File/folder commands
        file_cmd_match = re.match(
            r"(create|delete|open) (file|folder) (?:named|called)? ?'?([^']+)'? ?(?:in|on)? ?(desktop|documents)?",
            command
        )
        if file_cmd_match:
            action, obj_type, name, location = file_cmd_match.groups()
            location = location or "desktop"
            if action == "create":
                if obj_type == "file":
                    self.create_file(name, folder=location)
                else:
                    self.create_folder(name, folder=location)
            elif action == "delete":
                if obj_type == "file":
                    self.delete_file(name, folder=location)
                else:
                    self.delete_folder(name, folder=location)
            elif action == "open":
                self.open_file_or_folder(name, folder=location)
            return

        # Bookmark commands
        bookmark_cmd_match = re.match(
            r"(add|open|list) bookmarks?(?: named)? ?'?([^']*)?'?(?: for)? ?(.*)?",
            command
        )
        if bookmark_cmd_match:
            action, name, url = bookmark_cmd_match.groups()
            if action == "add" and name and url:
                self.add_bookmark(name, url)
            elif action == "open" and name:
                self.open_bookmark(name)
            elif action == "list":
                self.list_bookmarks()
            return

        # Other commands
        if command == "exit" or command == "quit":
            self.speak("Goodbye")
            self.close()
        elif "time" in command:
            self.get_time()
        elif "date" in command:
            self.get_date()
        elif "hello" in command or "hi" in command:
            self.speak("Hello there! How can I help you today?")
        elif "system resources" in command or "system info" in command:
            self.get_system_resources()
        elif "take screenshot" in command:
            self.take_screenshot()
        elif "show file manager" in command:
            self.show_file_manager()
        elif "open application" in command:
            app_name = command.replace("open application", "").strip()
            if app_name:
                self.open_application(app_name)
            else:
                self.speak("Please specify an application name.")
        elif "clipboard" in command:
            if "read clipboard" in command:
                self.manage_clipboard("read")
            elif "set clipboard" in command:
                content = command.replace("set clipboard", "").strip()
                if content:
                    self.manage_clipboard("set", content)
                else:
                    self.speak("Please specify content for the clipboard.")
        elif "set reminder" in command:
            match = re.search(r"set reminder (.+) at (\d{2}:\d{2} \d{2}-\d{2}-\d{4})", command)
            if match:
                message, time_str = match.groups()
                self.add_reminder(time_str, message)
            else:
                self.speak("Please say: set reminder [message] at HH:MM DD-MM-YYYY")
        elif "add task" in command:
            task = command.replace("add task", "").strip()
            if task:
                self.add_task(task)
            else:
                self.speak("Please specify a task")
        elif "list tasks" in command:
            self.list_tasks()
        elif "weather" in command:
            city_match = re.search(r"weather in (\w+)", command)
            city = city_match.group(1) if city_match else None
            self.get_weather(city)
        elif "send email" in command:
            match = re.search(r"send email to (\S+) subject (.+) body (.+)", command)
            if match:
                recipient, subject, body = match.groups()
                self.send_email(recipient, subject, body)
            else:
                self.speak("Please say: send email to [address] subject [subject] body [message]")
        elif "take note" in command:
            note = command.replace("take note", "").strip()
            if note:
                self.save_note(note)
            else:
                self.speak("Please specify a note")
        elif "read notes" in command:
            self.read_notes()
        elif "schedule event" in command:
            match = re.search(r"schedule event (.+) on (\d{2}-\d{2}-\d{4} \d{2}:\d{2})", command)
            if match:
                title, date_str = match.groups()
                self.add_calendar_event(title, date_str)
            else:
                self.speak("Please say: schedule event [title] on DD-MM-YYYY HH:MM")
        elif "set alias" in command:
            match = re.search(r"set alias (\w+) for (.+)", command)
            if match:
                alias, cmd = match.groups()
                self.add_alias(alias, cmd)
            else:
                self.speak("Please say: set alias [name] for [command]")
        elif "battery status" in command:
            self.get_battery_status()
        elif "open new tab" in command:
            url = command.replace("open new tab", "").strip()
            self.open_new_tab(url if url else None)
        elif "close tab" in command:
            self.close_browser_tab()
        elif "open incognito" in command:
            url = command.replace("open incognito", "").strip()
            self.open_incognito_mode(url if url else None)
        elif "scrape website" in command:
            url = command.replace("scrape website", "").strip()
            if url:
                self.scrape_website(url)
            else:
                self.speak("Please specify a website URL.")
        elif "fill form" in command:
            match = re.search(r"fill form on (.+) with (.+)", command)
            if match:
                website, query = match.groups()
                self.autofill_form(website, query)
            else:
                self.speak("Please say: fill form on [website] with [query]")
        elif "play" in command:
            self.handle_play_command(command)
        elif "pause" in command and ("music" in command or "song" in command):
            self.pause_music()
        elif "stop" in command and ("music" in command or "song" in command):
            self.stop_media()
        elif "resume" in command and ("music" in command or "song" in command):
            self.resume_music()
        elif "volume" in command:
            self.handle_volume_command(command)
        elif "search" in command and "on" in command:
            self.handle_search_command(command)
        elif "open" in command and ("google" in command or "youtube" in command):
            site = "google" if "google" in command else "youtube"
            webbrowser.open(f"https://www.{site}.com")
            self.speak(f"Opening {site}")
        else:
            self.handle_unknown_command(command)

    def handle_play_command(self, command):
        """Handle play commands for different media types."""
        if "song" in command or "music" in command:
            query = command.replace("play", "").replace("song", "").replace("music", "").strip()
            if query:
                self.play_music(query)
            else:
                self.speak("Please specify a song or artist to play.")
        elif "radio" in command:
            station = command.replace("play radio", "").strip()
            if station:
                self.play_radio(station)
            else:
                self.speak("Please specify a radio station.")
        elif "youtube" in command:
            query = command.replace("play", "").replace("on youtube", "").strip()
            if query:
                self.play_youtube(query)
            else:
                self.speak("Please specify a video to play on YouTube.")
        else:
            self.speak("Please specify what to play, like a song, radio, or YouTube video.")

    def play_music(self, query):
        """Play music using a file or online search in a thread."""
        def perform_operation():
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                # Check if query is a local file
                music_dir = os.path.join(os.path.expanduser("~"), "Music")
                file_path = os.path.join(music_dir, query)
                if os.path.exists(file_path) and file_path.lower().endswith(('.mp3', '.wav')):
                    pygame.mixer.music.load(file_path)
                    pygame.mixer.music.play()
                    self.music_playing = True
                    self.music_file = file_path
                    self.speak(f"Playing {query} from local music.")
                else:
                    # Use pywhatkit to play music online
                    pywhatkit.playonyt(query)
                    self.music_playing = True
                    self.music_file = None
                    self.speak(f"Playing {query} on YouTube.")
                for i in range(1, 101):
                    time.sleep(0.01)
                    self.progress_bar.setValue(i)
                notification.notify(
                    title="Music Playing",
                    message=f"Now playing: {query}",
                    timeout=5
                )
            except Exception as e:
                self.append_to_log(f"Failed to play music: {str(e)}", "Error")
                self.speak("Failed to play music. See log for details.")
            finally:
                self.progress_bar.setVisible(False)

        threading.Thread(target=perform_operation, daemon=True).start()

    def play_radio(self, station):
        """Play a radio station (mock implementation)."""
        try:
            # In a real implementation, this would use a radio streaming API or URL
            self.current_radio_station = station
            self.speak(f"Playing radio station {station}. (Note: Actual radio streaming requires additional setup.)")
            notification.notify(
                title="Radio Playing",
                message=f"Station: {station}",
                timeout=5
            )
        except Exception as e:
            self.append_to_log(f"Failed to play radio: {str(e)}", "Error")
            self.speak("Failed to play radio.")

    def play_youtube(self, query):
        """Play a YouTube video using pywhatkit."""
        def perform_operation():
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                pywhatkit.playonyt(query)
                for i in range(1, 101):
                    time.sleep(0.01)
                    self.progress_bar.setValue(i)
                self.speak(f"Playing {query} on YouTube.")
                notification.notify(
                    title="YouTube Playing",
                    message=f"Now playing: {query}",
                    timeout=5
                )
            except Exception as e:
                self.append_to_log(f"Failed to play YouTube video: {str(e)}", "Error")
                self.speak("Failed to play YouTube video.")
            finally:
                self.progress_bar.setVisible(False)

        threading.Thread(target=perform_operation, daemon=True).start()

    def pause_music(self):
        """Pause currently playing music."""
        if self.music_playing and pygame.mixer.music.get_busy():
            pygame.mixer.music.pause()
            self.speak("Music paused.")
            self.append_to_log("Music paused", "System")
            notification.notify(
                title="Music Paused",
                message="Playback paused",
                timeout=5
            )
        else:
            self.speak("No music is currently playing.")

    def resume_music(self):
        """Resume paused music."""
        if self.music_playing and not pygame.mixer.music.get_busy():
            pygame.mixer.music.unpause()
            self.speak("Music resumed.")
            self.append_to_log("Music resumed", "System")
            notification.notify(
                title="Music Resumed",
                message="Playback resumed",
                timeout=5
            )
        else:
            self.speak("No music is paused or playing.")

    def handle_volume_command(self, command):
        """Handle volume control commands."""
        if volume_control_available and platform.system() == "Windows":
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            
            if "volume up" in command:
                current_volume = volume.GetMasterVolumeLevelScalar()
                volume.SetMasterVolumeLevelScalar(min(current_volume + 0.1, 1.0), None)
                self.speak("Volume increased.")
            elif "volume down" in command:
                current_volume = volume.GetMasterVolumeLevelScalar()
                volume.SetMasterVolumeLevelScalar(max(current_volume - 0.1, 0.0), None)
                self.speak("Volume decreased.")
            elif "mute" in command:
                volume.SetMute(True, None)
                self.speak("Volume muted.")
            elif "unmute" in command:
                volume.SetMute(False, None)
                self.speak("Volume unmuted.")
            elif "set volume" in command:
                match = re.search(r"set volume to (\d+)", command)
                if match:
                    level = int(match.group(1)) / 100
                    if 0 <= level <= 1:
                        volume.SetMasterVolumeLevelScalar(level, None)
                        self.speak(f"Volume set to {int(level * 100)} percent.")
                    else:
                        self.speak("Please specify a volume level between 0 and 100.")
                else:
                    self.speak("Please specify a volume level, e.g., set volume to 50.")
            else:
                self.speak("Please specify a volume command like up, down, mute, unmute, or set volume to [number].")
        else:
            self.speak("Volume control is only available on Windows with pycaw installed.")

    def handle_search_command(self, command):
        """Handle search commands for different platforms."""
        def perform_operation():
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                query = command.replace("search", "").replace("on", "").strip()
                if "google" in command:
                    search_url = f"https://www.google.com/search?q={quote(query)}"
                    webbrowser.open(search_url)
                    self.speak(f"Searching Google for {query}.")
                elif "youtube" in command:
                    search_url = f"https://www.youtube.com/results?search_query={quote(query)}"
                    webbrowser.open(search_url)
                    self.speak(f"Searching YouTube for {query}.")
                elif "wikipedia" in command:
                    try:
                        summary = wikipedia.summary(query, sentences=2)
                        self.speak(summary)
                        self.append_to_log(f"Wikipedia summary for {query}: {summary}", "Assistant")
                    except wikipedia.exceptions.DisambiguationError as e:
                        self.speak(f"Multiple results found for {query}. Please be more specific.")
                    except wikipedia.exceptions.PageError:
                        self.speak(f"No Wikipedia page found for {query}.")
                else:
                    self.speak("Please specify where to search, e.g., Google, YouTube, or Wikipedia.")
                for i in range(1, 101):
                    time.sleep(0.01)
                    self.progress_bar.setValue(i)
            except Exception as e:
                self.append_to_log(f"Search error: {str(e)}", "Error")
                self.speak("Failed to perform search.")
            finally:
                self.progress_bar.setVisible(False)

        threading.Thread(target=perform_operation, daemon=True).start()

    def handle_unknown_command(self, command):
        """Handle unrecognized commands with suggestions or plugin fallback."""
        # Try fuzzy matching with known commands
        known_commands = [
            "time", "date", "weather", "set reminder", "add task", "list tasks",
            "take screenshot", "open application", "play music", "volume up",
            "search on google", "open bookmark", "add bookmark", "list bookmarks",
            "create file", "delete folder", "open file", "take note", "read notes",
            "send email", "schedule event", "set alias", "battery status"
        ]
        matches = [cmd for cmd in known_commands if command in cmd or cmd in command]
        if matches:
            self.speak(f"Did you mean '{matches[0]}'? Please try again.")
            return

        # Try plugins for custom commands
        for plugin_name, plugin in self.plugins.items():
            if plugin.get("handles_unknown", False):
                try:
                    plugin["execute"](command)
                    return
                except Exception as e:
                    self.append_to_log(f"Plugin {plugin_name} error on unknown command: {str(e)}", "Error")

        # Default to Wikipedia search as a fallback
        try:
            summary = wikipedia.summary(command, sentences=1)
            self.speak(f"I didn't understand the command, but here's a brief info: {summary}")
            self.append_to_log(f"Wikipedia fallback for {command}: {summary}", "Assistant")
        except wikipedia.exceptions.DisambiguationError:
            self.speak(f"I'm not sure what you mean by '{command}'. Could you clarify or specify a command?")
        except wikipedia.exceptions.PageError:
            self.speak(f"Sorry, I don't understand '{command}'. Try commands like 'time', 'weather', or 'set reminder'.")
        except Exception as e:
            self.append_to_log(f"Unknown command error: {str(e)}", "Error")
            self.speak("I didn't understand that command. Please try again or use a different command.")

    def closeEvent(self, event):
        """Handle window close event."""
        if self.tray_icon and self.tray_icon.isVisible():
            # Just minimize to tray instead of closing
            event.ignore()
            self.hide()
            self.minimized_to_tray = True
            self.tray_icon.showMessage(
                "Voice Assistant",
                "The assistant is still running in the background",
                QSystemTrayIcon.Information,
                2000
            )
        else:
            # Actually close the application
            self.save_config()
            self.save_reminders_and_tasks()
            pygame.mixer.quit()
            try:
                keyboard.unhook_all()
            except:
                pass
            if self.tray_icon:
                self.tray_icon.hide()
            event.accept()

    def changeEvent(self, event):
        """Handle window state changes (minimizing)."""
        if event.type() == event.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                if self.tray_icon and self.tray_icon.isVisible():
                    # Minimize to tray instead of taskbar
                    self.hide()
                    self.minimized_to_tray = True
                    self.tray_icon.showMessage(
                        "Voice Assistant",
                        "The assistant is still running in the background",
                        QSystemTrayIcon.Information,
                        2000
                    )
        super().changeEvent(event)

    # ... [rest of the existing methods remain the same] ...

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Ensure the application doesn't quit when last window is closed
    app.setQuitOnLastWindowClosed(False)
    
    assistant = VoiceAssistantGUI()
    assistant.show()
    sys.exit(app.exec_())
