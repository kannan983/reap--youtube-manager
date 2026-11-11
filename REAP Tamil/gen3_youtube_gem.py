import sys
import os
import json
import logging
import re
from datetime import datetime, timezone
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTabWidget,
    QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QProgressBar, QMessageBox, QTextEdit, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont

# --- Google API Imports ---
import google.oauth2.credentials
import google_auth_oauthlib.flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# --- Constants ---
SCOPES = ["https://www.googleapis.com/auth/youtube.force-ssl"]
API_SERVICE_NAME = "youtube"
API_VERSION = "v3"
OAUTH_CLIENT_FILE = "oauth_client.json"
API_KEY_FILE = "api.txt"
TOKEN_FILE = "token.json"
LOGS_FOLDER = "logs"
EXCEL_FOLDER = "Excel sheet"
APP_TITLE = "Reap Tamil YouTube Channel Manager"
COURSE_INTRO_SUFFIX = " Course Introduction"
# Regex to check if a title IS in the desired format (Chapter X[Y] - Topic)
TITLE_FORMAT_REGEX = re.compile(r'^Chapter \d+[A-Z]? - .*', re.IGNORECASE)
# Regex for parsing chapters (more detailed)
CHAPTER_REGEX = re.compile(r'^Chapter (\d+)([A-Z]?) - .*', re.IGNORECASE)
BASE_CHAPTER_REGEX = re.compile(r'^Chapter (\d+) - .*', re.IGNORECASE)

# --- Logging Setup ---
def setup_logging(log_dir):
    if not os.path.exists(log_dir): os.makedirs(log_dir)
    log_filename = datetime.now().strftime("%Y-%m-%d_%H-%M-%S.log")
    log_filepath = os.path.join(log_dir, log_filename)
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]: root_logger.removeHandler(handler)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.FileHandler(log_filepath), logging.StreamHandler(sys.stdout)]
    )
    logging.info("Logging initialized.")

# --- Natural Sort & Excel Logic Helpers ---
def excel_sort_key(video_title):
    """ Custom key function for Excel sorting logic. """
    title = video_title.strip()
    if title.lower().endswith(COURSE_INTRO_SUFFIX.lower()): return (0, 0, '', title)
    base_match = BASE_CHAPTER_REGEX.match(title)
    sub_match = CHAPTER_REGEX.match(title)
    if base_match and not (sub_match and sub_match.group(2)): return (1, int(base_match.group(1)), '', title)
    if sub_match: return (1, int(sub_match.group(1)), sub_match.group(2).upper(), title)
    return (2, 0, '', title)

def sanitize_filename(text, replace_space_with='_', max_length=150):
    """Sanitizes text for filenames."""
    text = ' '.join(text.split()).replace(' ', replace_space_with)
    sanitized = re.sub(r'[\\/*?:"<>|]', "", text)
    return sanitized[:max_length].strip()

# --- LED Label (Unchanged) ---
class LedIndicator(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(16, 16); self.set_state('grey')
    def set_state(self, color):
        styles = {
            'green': "background-color: qlineargradient(spread:pad, x1:0.145, y1:0.16, x2:1, y2:1, stop:0 rgba(20, 255, 20, 255), stop:1 rgba(0, 150, 0, 255));",
            'red': "background-color: qlineargradient(spread:pad, x1:0.145, y1:0.16, x2:1, y2:1, stop:0 rgba(255, 20, 20, 255), stop:1 rgba(150, 0, 0, 255));",
            'grey': "background-color: qlineargradient(spread:pad, x1:0.145, y1:0.16, x2:1, y2:1, stop:0 rgba(150, 150, 150, 255), stop:1 rgba(80, 80, 80, 255));"
        }
        self.setStyleSheet(f"QLabel {{ border: 1px solid #555; border-radius: 8px; {styles.get(color, styles['grey'])} }}")

# --- Worker Threads (Unchanged) ---
class AuthWorker(QThread): # (Unchanged)
    auth_finished = pyqtSignal(object, str)
    def __init__(self, mode, auth_folder_path): super().__init__(); self.mode = mode; self.auth_folder_path = auth_folder_path; self.client_secrets_file = os.path.join(self.auth_folder_path, OAUTH_CLIENT_FILE); self.token_file = os.path.join(self.auth_folder_path, TOKEN_FILE); self._is_running = True
    def run(self):
        logging.info(f"Starting auth (mode: {self.mode}).")
        credentials = None
        if os.path.exists(self.token_file):
            try:
                credentials = Credentials.from_authorized_user_file(self.token_file, SCOPES)
                if credentials and credentials.valid: self.auth_finished.emit(credentials, "Token loaded."); return
                if credentials and credentials.expired and credentials.refresh_token:
                    credentials.refresh(Request())
                    with open(self.token_file, 'w') as token: token.write(credentials.to_json())
                    self.auth_finished.emit(credentials, "Token refreshed."); return
                credentials = None
            except Exception as e: logging.error(f"Token error: {e}"); credentials = None
        if self.mode == 'token': self.auth_finished.emit(None, "Token invalid/missing."); return
        if not os.path.exists(self.client_secrets_file): self.auth_finished.emit(None, f"{OAUTH_CLIENT_FILE} not found."); return
        try:
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(self.client_secrets_file, SCOPES)
            credentials = flow.run_local_server(port=0)
            if not self._is_running: self.auth_finished.emit(None, "Aborted."); return
            with open(self.token_file, 'w') as token: token.write(credentials.to_json())
            self.auth_finished.emit(credentials, "Auth successful.")
        except Exception as e: logging.error(f"Browser flow error: {e}"); self.auth_finished.emit(None, f"Auth failed: {e}")
    def stop(self): self._is_running = False

class PlaylistWorker(QThread): # (Unchanged)
    playlists_ready = pyqtSignal(list, str)
    def __init__(self, youtube_service_oauth): super().__init__(); self.youtube_service = youtube_service_oauth
    def run(self):
        logging.info("Fetching playlists...")
        if not self.youtube_service: self.playlists_ready.emit([], "Error: Not authenticated."); return
        playlists, next_page_token = [], None
        try:
            while True:
                r = self.youtube_service.playlists().list(part="snippet,contentDetails", mine=True, maxResults=50, pageToken=next_page_token).execute()
                for i in r.get("items", []): playlists.append({"id": i["id"], "title": i["snippet"]["title"], "description": i["snippet"].get("description", ""), "video_count": i["contentDetails"]["itemCount"]})
                next_page_token = r.get("nextPageToken")
                if not next_page_token: break
            self.playlists_ready.emit(playlists, "Playlists loaded.")
        except Exception as e: logging.error(f"Error fetching playlists: {e}"); self.playlists_ready.emit([], f"Error: {e}")

class VideoListWorker(QThread): # (Unchanged)
    videos_ready = pyqtSignal(list, str)
    def __init__(self, youtube_service_api, playlist_id): super().__init__(); self.youtube_service = youtube_service_api; self.playlist_id = playlist_id
    def run(self):
        logging.info(f"Fetching videos for {self.playlist_id}...")
        if not self.youtube_service or not self.playlist_id: self.videos_ready.emit([], "Error: Service/ID missing."); return
        videos, next_page_token = [], None
        try:
            while True:
                r = self.youtube_service.playlistItems().list(part="snippet", playlistId=self.playlist_id, maxResults=50, pageToken=next_page_token).execute()
                for i in r.get("items", []): s = i["snippet"]; videos.append({"id": i["id"], "videoId": s["resourceId"]["videoId"], "title": s["title"], "description": s.get("description", ""), "position": s["position"]})
                next_page_token = r.get("nextPageToken")
                if not next_page_token: break
            self.videos_ready.emit(videos, "Videos fetched for processing.")
        except Exception as e: logging.error(f"Error fetching videos: {e}"); self.videos_ready.emit([], f"Error: {e}")

class UpdateWorker(QThread): # (Unchanged)
    update_progress = pyqtSignal(int); update_finished = pyqtSignal(str)
    def __init__(self, youtube_service_oauth, videos_to_update): super().__init__(); self.youtube_service = youtube_service_oauth; self.videos_to_update = videos_to_update
    def run(self):
        logging.info(f"Starting update: {len(self.videos_to_update)} videos.")
        if not self.youtube_service: self.update_finished.emit("Error: Not authenticated."); return
        total, ok, fail = len(self.videos_to_update), 0, 0
        for i, v_data in enumerate(self.videos_to_update):
            v_id, title, desc = v_data["videoId"], v_data["new_title"], v_data["new_description"]
            try:
                get_r = self.youtube_service.videos().list(part="snippet", id=v_id).execute()
                if not get_r.get("items"): logging.error(f"Video not found {v_id}."); fail += 1; continue
                cat_id = get_r["items"][0]["snippet"]["categoryId"]
                body = {"id": v_id, "snippet": {"title": title, "description": desc, "categoryId": cat_id}}
                self.youtube_service.videos().update(part="snippet", body=body).execute()
                ok += 1; logging.info(f"Updated {v_id}")
            except Exception as e: logging.error(f"Error updating {v_id}: {e}"); fail += 1
            self.update_progress.emit(int(((i + 1) / total) * 100))
        self.update_finished.emit(f"Update complete. OK: {ok}, Failed: {fail}.")

# --- Main Application Window ---
class YoutubeManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # ... (Auth variables) ...
        self.auth_folder = None; self.api_key = None; self.credentials = None
        self.youtube_service_oauth = None; self.youtube_service_api = None
        self.auth_worker = None
        self.videos_marked_for_update = set()
        self.current_playlist_data = {}
        self.processed_excel_data = None

        self.setWindowTitle(APP_TITLE)
        self.setGeometry(100, 100, 800, 700)
        font = QFont(); font.setPointSize(10); QApplication.setFont(font)
        self.central_widget = QWidget(); self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.tabs = QTabWidget()
        self.tab1 = QWidget(); self.tab2 = QWidget(); self.tab3 = QWidget()
        self.tabs.addTab(self.tab1, "1. Authentication")
        self.tabs.addTab(self.tab2, "2. Rename & Describe")
        self.tabs.addTab(self.tab3, "3. Excel Export")
        self.statusBar = self.statusBar()
        self.led_indicator = LedIndicator()
        self.status_label = QLabel("Welcome!")
        self.statusBar.addPermanentWidget(self.led_indicator, 0)
        self.statusBar.addPermanentWidget(self.status_label, 1)
        self.init_tab1(); self.init_tab2(); self.init_tab3()
        self.layout.addWidget(self.tabs)
        self.set_auth_state(False)

    def init_tab1(self): # (Unchanged)
        layout = QVBoxLayout(self.tab1); layout.setAlignment(Qt.AlignTop)
        folder_layout = QHBoxLayout(); folder_layout.addWidget(QLabel("Auth Folder:"))
        self.folder_path_edit = QLineEdit(); self.folder_path_edit.setReadOnly(True)
        self.browse_button = QPushButton("Browse..."); self.browse_button.clicked.connect(self.browse_folder)
        folder_layout.addWidget(self.folder_path_edit); folder_layout.addWidget(self.browse_button)
        layout.addLayout(folder_layout)
        button_layout = QHBoxLayout()
        self.token_auth_button = QPushButton("Authorize with Token"); self.token_auth_button.setEnabled(False); self.token_auth_button.clicked.connect(self.token_auth_clicked)
        self.auth_button = QPushButton("Authorize (Browser)"); self.auth_button.setEnabled(False); self.auth_button.clicked.connect(self.start_auth_browser)
        self.abort_button = QPushButton("Abort / Reset"); self.abort_button.setEnabled(False); self.abort_button.clicked.connect(self.abort_auth_clicked)
        button_layout.addWidget(self.token_auth_button); button_layout.addWidget(self.auth_button); button_layout.addWidget(self.abort_button)
        layout.addLayout(button_layout); layout.addItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))
        self.auth_status_label = QLabel("Status: Awaiting folder."); self.auth_status_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.auth_status_label)
        instructions = QTextEdit(); instructions.setReadOnly(True)
        instructions.setHtml("<h3>Instructions:</h3><ol><li>Select folder.</li><li>Use 'Token' or 'Browser'.</li><li>Status/LED will update.</li></ol>")
        layout.addWidget(instructions); layout.addStretch()

    def init_tab2(self): # (Unchanged)
        layout = QVBoxLayout(self.tab2)
        pl_layout = QHBoxLayout(); pl_layout.addWidget(QLabel("Select Playlist:"))
        self.playlist_combo_rename = QComboBox(); self.playlist_combo_rename.setEnabled(False)
        pl_layout.addWidget(self.playlist_combo_rename, 1)
        self.load_videos_button = QPushButton("Load Videos"); self.load_videos_button.setEnabled(False); self.load_videos_button.clicked.connect(self.load_videos_rename_clicked)
        pl_layout.addWidget(self.load_videos_button); layout.addLayout(pl_layout)
        self.video_status_label = QLabel("Status: Select playlist & load.")
        layout.addWidget(self.video_status_label)
        self.video_table = QTableWidget(); self.video_table.setColumnCount(4)
        self.video_table.setHorizontalHeaderLabels(["Original", "Corrected", "Description", "ID"]); self.video_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.video_table.setColumnHidden(3, True); layout.addWidget(self.video_table)
        bottom_layout = QHBoxLayout()
        self.rename_count_label = QLabel("0 videos to rename."); self.rename_count_label.setStyleSheet("font-weight: bold; color: #d35400;")
        bottom_layout.addWidget(self.rename_count_label, 1)
        self.apply_button = QPushButton("Apply Changes"); self.apply_button.setEnabled(False); self.apply_button.clicked.connect(self.apply_changes_clicked)
        self.progress_bar = QProgressBar(); self.progress_bar.setVisible(False); self.progress_bar.setMaximumWidth(200)
        bottom_layout.addWidget(self.apply_button); bottom_layout.addWidget(self.progress_bar); layout.addLayout(bottom_layout)

    def init_tab3(self): # (Unchanged)
        layout = QVBoxLayout(self.tab3)
        pl_layout = QHBoxLayout()
        pl_layout.addWidget(QLabel("Select Playlist:"))
        self.playlist_combo_excel = QComboBox(); self.playlist_combo_excel.setEnabled(False)
        pl_layout.addWidget(self.playlist_combo_excel, 1)
        self.load_excel_button = QPushButton("Load Video Info"); self.load_excel_button.setEnabled(False)
        self.load_excel_button.clicked.connect(self.load_excel_preview_clicked)
        pl_layout.addWidget(self.load_excel_button); layout.addLayout(pl_layout)
        self.excel_status_label = QLabel("Status: Select playlist & load for preview.")
        layout.addWidget(self.excel_status_label)
        self.excel_table = QTableWidget(); self.excel_table.setColumnCount(7)
        self.excel_table.setHorizontalHeaderLabels(["Course Code", "Chapter Name", "YouTube URL", "Video Title", "Video Description", "OrderNo", "Language Code"])
        self.excel_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); self.excel_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.excel_table)
        self.generate_excel_button = QPushButton("Generate Excel Sheet"); self.generate_excel_button.setEnabled(False)
        self.generate_excel_button.clicked.connect(self.generate_excel_clicked)
        layout.addWidget(self.generate_excel_button, 0, Qt.AlignRight)

    # --- Auth Methods (Unchanged) ---
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Auth Folder")
        if folder: self.auth_folder = folder; self.folder_path_edit.setText(folder); setup_logging(os.path.join(folder, LOGS_FOLDER)); self.update_auth_button_states()
    def update_auth_button_states(self):
        oauth_ok, api_ok, token_exists = False, False, False
        if self.auth_folder:
            oauth_path = os.path.join(self.auth_folder, OAUTH_CLIENT_FILE); api_path = os.path.join(self.auth_folder, API_KEY_FILE); token_path = os.path.join(self.auth_folder, TOKEN_FILE)
            oauth_ok = os.path.exists(oauth_path); token_exists = os.path.exists(token_path)
            try:
                with open(api_path, 'r') as f: content = f.read().strip(); match = re.search(r'"API Key"\s*:\s*"(.*?)"', content)
                if match: self.api_key = match.group(1); api_ok = True
                else: self.api_key = None; api_ok = False
            except Exception: self.api_key = None; api_ok = False
        files_ok = oauth_ok and api_ok
        self.auth_button.setEnabled(files_ok); self.token_auth_button.setEnabled(files_ok and token_exists); self.abort_button.setEnabled(bool(self.auth_folder))
        if not self.auth_folder: self.auth_status_label.setText("Status: Awaiting folder.")
        elif not files_ok: self.auth_status_label.setText("Status: Error! oauth_client.json/api.txt missing.")
        else: self.auth_status_label.setText("Status: Ready. Use Token or Browser.")
    def token_auth_clicked(self): self.start_auth('token')
    def start_auth_browser(self): self.start_auth('browser')
    def start_auth(self, mode):
        if not self.auth_folder: return
        self.auth_button.setEnabled(False); self.token_auth_button.setEnabled(False); self.abort_button.setEnabled(True)
        self.auth_status_label.setText(f"Status: Starting '{mode}' auth..."); self.status_label.setText("Auth in progress...")
        self.auth_worker = AuthWorker(mode, self.auth_folder); self.auth_worker.auth_finished.connect(self.auth_finished); self.auth_worker.start()
    def abort_auth_clicked(self):
        if self.auth_worker and self.auth_worker.isRunning(): self.auth_worker.stop()
        self.auth_worker = None; self.set_auth_state(False); self.update_auth_button_states()
    def auth_finished(self, credentials, message):
        self.auth_worker = None; self.auth_status_label.setText(f"Status: {message}")
        if credentials:
            self.credentials = credentials; self.set_auth_state(True, message)
            try:
                self.youtube_service_oauth = build(API_SERVICE_NAME, API_VERSION, credentials=self.credentials)
                self.youtube_service_api = build(API_SERVICE_NAME, API_VERSION, developerKey=self.api_key)
                self.load_playlists()
            except Exception as e: self.show_message("Error", f"Service build failed: {e}"); self.set_auth_state(False, f"Service fail: {e}")
        else: self.show_message("Auth Failed", message); self.set_auth_state(False, message)
    def set_auth_state(self, is_authenticated, message=""):
        if is_authenticated and self.credentials:
            self.led_indicator.set_state('green'); self.tabs.setTabEnabled(1, True); self.tabs.setTabEnabled(2, True)
            self.auth_button.setEnabled(False); self.token_auth_button.setEnabled(False); self.abort_button.setEnabled(True)
            try:
                now_utc = datetime.now(timezone.utc); expiry_utc = self.credentials.expiry.replace(tzinfo=timezone.utc)
                days_left = (expiry_utc - now_utc).days
                status_msg = f"Authenticated! Expires in {days_left} days. ({message})" if days_left >= 0 else f"Authenticated! (Expired). ({message})"
                self.status_label.setText(status_msg); logging.info(status_msg)
            except Exception as e: self.status_label.setText(f"Authenticated! (Expiry ?). ({message})"); logging.error(f"Expiry check: {e}")
        else:
            self.credentials = None; self.youtube_service_oauth = None; self.youtube_service_api = None
            self.led_indicator.set_state('red'); self.tabs.setTabEnabled(1, False); self.tabs.setTabEnabled(2, False)
            self.playlist_combo_rename.clear(); self.playlist_combo_excel.clear()
            self.video_table.setRowCount(0); self.excel_table.setRowCount(0); self.status_label.setText(f"Not Authenticated. ({message})")
            self.update_auth_button_states()

    # --- Playlist Loading ---
    def load_playlists(self):
        if not self.youtube_service_oauth: self.show_message("Error", "Auth first."); return
        self.status_label.setText("Loading playlists...");
        self.playlist_combo_rename.clear(); self.playlist_combo_rename.setEnabled(False)
        self.playlist_combo_excel.clear(); self.playlist_combo_excel.setEnabled(False)
        self.load_videos_button.setEnabled(False); self.load_excel_button.setEnabled(False)
        self.playlist_worker = PlaylistWorker(self.youtube_service_oauth)
        self.playlist_worker.playlists_ready.connect(self.playlists_loaded)
        self.playlist_worker.start()

    def playlists_loaded(self, playlists, message):
        self.status_label.setText(message); self.current_playlist_data = {}
        if playlists:
            self.playlist_combo_rename.setEnabled(True); self.load_videos_button.setEnabled(True)
            self.playlist_combo_excel.setEnabled(True); self.load_excel_button.setEnabled(True)
            for pl in playlists:
                d = (pl['description'][:30] + '...') if len(pl['description']) > 30 else pl['description']
                item_text = f"{pl['title']} - {d} - {pl['video_count']} videos"
                self.current_playlist_data[pl['id']] = pl # Store full data using ID as key
                self.playlist_combo_rename.addItem(item_text, pl['id'])
                self.playlist_combo_excel.addItem(item_text, pl['id'])
        else: self.show_message("Info", "No playlists found/error.")

    # --- Tab 2 Methods (Unchanged) ---
    def load_videos_rename_clicked(self):
        idx = self.playlist_combo_rename.currentIndex();
        if idx < 0: return
        pl_id = self.playlist_combo_rename.itemData(idx); self.status_label.setText(f"Loading rename videos..."); self.video_table.setRowCount(0); self.load_videos_button.setEnabled(False); self.apply_button.setEnabled(False)
        self.video_worker_rename = VideoListWorker(self.youtube_service_api, pl_id); self.video_worker_rename.videos_ready.connect(self.videos_loaded_rename); self.video_worker_rename.start()
    def videos_loaded_rename(self, videos, message):
        self.status_label.setText(message); self.load_videos_button.setEnabled(True)
        if not videos: self.video_status_label.setText(f"Status: No videos / error: {message}"); return
        self.video_table.setRowCount(len(videos)); self.videos_marked_for_update.clear(); videos_to_rename_count = 0; bold_font = QFont(); bold_font.setBold(True)
        for row, video in enumerate(videos):
            original_title = video['title']; video_id = video['videoId']; needs_correction = False
            original_item = QTableWidgetItem(original_title); original_item.setFlags(original_item.flags() & ~Qt.ItemIsEditable)
            if original_title.lower().endswith(COURSE_INTRO_SUFFIX.lower()): corrected_title = original_title; new_description = original_title; needs_correction = False
            elif TITLE_FORMAT_REGEX.match(original_title): parts = original_title.split(" - ", 1); corrected_title = original_title; new_description = parts[1].strip() if len(parts) > 1 else ""; needs_correction = False
            else: needs_correction = True; match = re.match(r'(.*? \d+[A-Z]?)\s+(.*)', original_title); corrected_title, new_description = (f"{match.group(1)} - {match.group(2)}", match.group(2).strip()) if match else (f"{original_title} - ", "")
            if needs_correction: original_item.setFont(bold_font); videos_to_rename_count += 1; self.videos_marked_for_update.add(video_id)
            corrected_item = QTableWidgetItem(corrected_title); desc_item = QTableWidgetItem(new_description); id_item = QTableWidgetItem(video_id); id_item.setFlags(id_item.flags() & ~Qt.ItemIsEditable)
            self.video_table.setItem(row, 0, original_item); self.video_table.setItem(row, 1, corrected_item); self.video_table.setItem(row, 2, desc_item); self.video_table.setItem(row, 3, id_item)
        self.video_status_label.setText(f"Loaded {len(videos)} videos."); self.rename_count_label.setText(f"{videos_to_rename_count} to rename."); self.apply_button.setEnabled(videos_to_rename_count > 0)
    def apply_changes_clicked(self):
        if not self.videos_marked_for_update: self.show_message("Info", "No videos marked."); return
        videos_to_process = []
        for row in range(self.video_table.rowCount()):
            if self.video_table.item(row, 3).text() in self.videos_marked_for_update: videos_to_process.append({"videoId": self.video_table.item(row, 3).text(), "new_title": self.video_table.item(row, 1).text(), "new_description": self.video_table.item(row, 2).text()})
        if not videos_to_process: self.show_message("Info", "No videos to process."); return
        if QMessageBox.question(self, 'Confirm', f"Update {len(videos_to_process)} videos?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.No: return
        self.apply_button.setEnabled(False); self.progress_bar.setValue(0); self.progress_bar.setVisible(True)
        self.update_worker = UpdateWorker(self.youtube_service_oauth, videos_to_process); self.update_worker.update_progress.connect(self.update_progress); self.update_worker.update_finished.connect(self.update_finished); self.update_worker.start()
    def update_progress(self, value): self.progress_bar.setValue(value); self.status_label.setText(f"Updating... {value}%")
    def update_finished(self, message): self.progress_bar.setVisible(False); self.status_label.setText(message); self.show_message("Update Complete", message); self.load_videos_rename_clicked()


    # --- Tab 3 Methods ---
    def load_excel_preview_clicked(self):
        idx = self.playlist_combo_excel.currentIndex()
        if idx < 0: return
        pl_id = self.playlist_combo_excel.itemData(idx)
        self.excel_status_label.setText(f"Status: Fetching videos for {pl_id}...")
        self.excel_table.setRowCount(0); self.load_excel_button.setEnabled(False); self.generate_excel_button.setEnabled(False); self.processed_excel_data = None
        self.video_worker_excel = VideoListWorker(self.youtube_service_api, pl_id)
        self.video_worker_excel.videos_ready.connect(self.process_and_display_excel_data)
        self.video_worker_excel.start()

    def process_and_display_excel_data(self, videos, message):
        """Processes video data according to Excel logic and displays it."""
        self.load_excel_button.setEnabled(True)
        if not videos: self.excel_status_label.setText(f"Status: Failed - {message}"); return

        idx = self.playlist_combo_excel.currentIndex()
        pl_id = self.playlist_combo_excel.itemData(idx)
        pl_data = self.current_playlist_data.get(pl_id, {})
        pl_title_raw = pl_data.get("title", "Unknown_Playlist") # <-- Use raw title
        pl_description_raw = pl_data.get("description", "No_Description") # <-- Use raw description

        logging.info(f"Processing {len(videos)} videos for Excel (Playlist Title: {pl_title_raw}).")

        # 1. Extract Course Code & Language Code from Playlist Title
        parts = pl_title_raw.split('_')
        course_code = parts[0] if parts else pl_title_raw
        lang_code = parts[-1] if len(parts) > 1 and len(parts[-1]) == 2 else "N/A" # Check if last part looks like lang code

        # 2. Sort Videos
        sorted_videos = sorted(videos, key=lambda v: excel_sort_key(v['title']))

        # 3. Build Base Chapter Map
        base_chapter_titles = {}
        for video in sorted_videos:
            title = video['title']
            base_match = BASE_CHAPTER_REGEX.match(title)
            # Ensure it's a base chapter (no letter part)
            if base_match:
                sub_match = CHAPTER_REGEX.match(title)
                if sub_match and not sub_match.group(2):
                    base_chapter_titles[int(base_match.group(1))] = title

        # 4. Process and build data rows
        processed_rows = []
        current_base_chapter_num = -1
        sub_chapter_count = 1
        for video in sorted_videos:
            title = video['title']
            video_id = video['videoId']
            description = video['description']
            url = f"https://youtu.be/{video_id}" # As requested
            chapter_name = ""
            order_no = ""

            is_intro = title.lower().endswith(COURSE_INTRO_SUFFIX.lower())
            base_match = BASE_CHAPTER_REGEX.match(title)
            sub_match = CHAPTER_REGEX.match(title)

            if is_intro:
                chapter_name = ""
                order_no = ""
                current_base_chapter_num = -1
            elif base_match and not (sub_match and sub_match.group(2)): # Base Chapter
                chapter_name = title
                order_no = 0
                current_base_chapter_num = int(base_match.group(1))
                sub_chapter_count = 1
            elif sub_match: # Sub Chapter
                chap_num = int(sub_match.group(1))
                chapter_name = base_chapter_titles.get(chap_num, title) # <-- Use base title, fallback to self
                if chap_num == current_base_chapter_num:
                    order_no = sub_chapter_count
                    sub_chapter_count += 1
                else: # Sub-chapter starting new sequence (base missing?)
                    order_no = 1
                    current_base_chapter_num = chap_num
                    sub_chapter_count = 2
            else: # Unexpected format
                chapter_name = title
                order_no = ""
                current_base_chapter_num = -1

            processed_rows.append({
                "Course Code": course_code,
                "Chapter Name": chapter_name,
                "YouTube URL": url,
                "Video Title": title,
                "Video Description": description,
                "OrderNo in Chapter": str(order_no),
                "Language Code": lang_code
            })

        self.processed_excel_data = processed_rows
        self.populate_excel_table(processed_rows)
        self.excel_status_label.setText(f"Status: {len(processed_rows)} videos processed. Ready to generate Excel.")
        self.generate_excel_button.setEnabled(True)

    def populate_excel_table(self, data):
        # (Unchanged)
        self.excel_table.setRowCount(len(data))
        headers = ["Course Code", "Chapter Name", "YouTube URL", "Video Title", "Video Description", "OrderNo in Chapter", "Language Code"]
        for row_idx, row_data in enumerate(data):
            for col_idx, header in enumerate(headers):
                item_text = str(row_data.get(header, ""))
                self.excel_table.setItem(row_idx, col_idx, QTableWidgetItem(item_text))
        self.excel_table.resizeColumnsToContents()

    def generate_excel_clicked(self):
        """Generates and saves the Excel file AUTOMATICALLY."""
        if not self.processed_excel_data or not self.auth_folder:
            self.show_message("Error", "No data loaded or auth folder not set.", QMessageBox.Warning)
            return

        idx = self.playlist_combo_excel.currentIndex()
        pl_id = self.playlist_combo_excel.itemData(idx)
        pl_data = self.current_playlist_data.get(pl_id, {})
        pl_title_raw = pl_data.get("title", "Unknown_Playlist")
        pl_description_raw = pl_data.get("description", "No_Description")

        # Extract course code from title
        course_code = pl_title_raw.split('_')[0] if pl_title_raw else pl_title_raw

        # Sanitize description and build filename
        sanitized_desc = sanitize_filename(pl_description_raw)
        filename_base = f"{sanitized_desc}_{course_code}.xlsx"

        excel_dir = os.path.join(self.auth_folder, EXCEL_FOLDER)
        if not os.path.exists(excel_dir):
            try:
                os.makedirs(excel_dir)
                logging.info(f"Created Excel folder: {excel_dir}")
            except Exception as e:
                 self.show_message("Error", f"Could not create Excel folder: {e}", QMessageBox.Critical)
                 logging.error(f"Failed to create Excel folder: {e}")
                 return

        full_path = os.path.join(excel_dir, filename_base)

        try:
            logging.info(f"Generating Excel file automatically: {full_path}")
            df = pd.DataFrame(self.processed_excel_data)
            # Ensure columns are in the exact order
            df = df[["Course Code", "Chapter Name", "YouTube URL", "Video Title", "Video Description", "OrderNo in Chapter", "Language Code"]]
            df.to_excel(full_path, index=False, engine='openpyxl')
            self.show_message("Success", f"Excel file saved successfully to:\n{full_path}")
            logging.info("Excel file generation successful.")
            # Optionally, open the folder:
            # os.startfile(excel_dir) # Windows only
        except Exception as e:
            self.show_message("Error", f"Failed to save Excel file: {e}", QMessageBox.Critical)
            logging.error(f"Excel save failed: {e}", exc_info=True)


    # --- Helper & Closing Methods ---
    def show_message(self, title, message, icon=QMessageBox.Information):
        msg = QMessageBox(self); msg.setIcon(icon); msg.setText(message); msg.setWindowTitle(title); msg.setStandardButtons(QMessageBox.Ok); msg.exec_()
    def closeEvent(self, event):
        if self.auth_worker and self.auth_worker.isRunning(): self.auth_worker.stop()
        logging.info("Application closing."); event.accept()

# --- Main Execution ---
if __name__ == '__main__':
    print("Ensure you have installed pandas and openpyxl: pip install pandas openpyxl")
    app = QApplication(sys.argv)
    window = YoutubeManagerApp()
    window.show()
    sys.exit(app.exec_())