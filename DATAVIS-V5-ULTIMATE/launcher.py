import sys
import subprocess
import os
import webbrowser
import urllib.request
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QProgressBar, QGraphicsDropShadowEffect, QFrame
)
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QColor, QFont


STREAMLIT_URL = "http://localhost:8501"


class LauncherApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DATAVIS v5 Launcher")
        self.setFixedSize(380, 460)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.process = None
        self.check_timer = None       # polls whether the Streamlit server is up
        self.pulse_timer = None       # animates the status dot while running
        self._pulse_state = True

        self._build_ui()
        self._apply_styles()
        self._fade_in()

    # ------------------------------------------------------------
    # UI CONSTRUCTION
    # ------------------------------------------------------------
    def _build_ui(self):
        outer = QVBoxLayout()
        outer.setContentsMargins(0, 0, 0, 0)

        # main card (gives us the rounded gradient panel + drop shadow)
        self.card = QFrame(objectName="card")
        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(28, 26, 28, 26)
        card_layout.setSpacing(14)

        shadow = QGraphicsDropShadowEffect(blurRadius=40, xOffset=0, yOffset=8)
        shadow.setColor(QColor(0, 0, 0, 160))
        self.card.setGraphicsEffect(shadow)

        # custom title bar (frameless window needs its own close/drag handling)
        title_row = QHBoxLayout()
        title_label = QLabel("⚡ DATAVIS v5🖥️")
        title_label.setObjectName("titleLabel")
        close_btn = QPushButton("✕")
        close_btn.setObjectName("closeBtn")
        close_btn.setFixedSize(26, 26)
        close_btn.clicked.connect(self.close)
        title_row.addWidget(title_label)
        title_row.addStretch()
        title_row.addWidget(close_btn)
        card_layout.addLayout(title_row)

        subtitle = QLabel("Project Launcher & Control Panel")
        subtitle.setObjectName("subtitle")
        card_layout.addWidget(subtitle)

        card_layout.addSpacing(6)

        # status row: pulsing dot + text
        status_row = QHBoxLayout()
        self.status_dot = QLabel("●")
        self.status_dot.setObjectName("statusDotOff")
        self.status_label = QLabel("Not running")
        self.status_label.setObjectName("statusText")
        status_row.addWidget(self.status_dot)
        status_row.addWidget(self.status_label)
        status_row.addStretch()
        card_layout.addLayout(status_row)

        # progress bar — indeterminate while the server boots up
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)   # indeterminate ("busy") mode
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(6)
        self.progress.hide()
        card_layout.addWidget(self.progress)

        card_layout.addSpacing(10)

        # primary action button
        self.btn_run = QPushButton("▶  Run Data Visualizer")
        self.btn_run.setObjectName("primaryBtn")
        self.btn_run.setCursor(Qt.PointingHandCursor)
        self.btn_run.clicked.connect(self.launch_visualizer)
        card_layout.addWidget(self.btn_run)

        # secondary / stop button
        self.btn_stop = QPushButton("■  Stop")
        self.btn_stop.setObjectName("stopBtn")
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.clicked.connect(self.stop_visualizer)
        self.btn_stop.setEnabled(False)
        card_layout.addWidget(self.btn_stop)

        card_layout.addSpacing(6)

        # utility: open the running app in the default browser
        self.btn_open = QPushButton("🌐  Open in Browser")
        self.btn_open.setObjectName("stopBtn")
        self.btn_open.setCursor(Qt.PointingHandCursor)
        self.btn_open.clicked.connect(self.open_in_browser)
        self.btn_open.setEnabled(False)
        card_layout.addWidget(self.btn_open)

        card_layout.addSpacing(10)

        future_label = QLabel("v5 ULTIMATE — Dashboard · Cleaning · Excel Export")
        future_label.setObjectName("futureLabel")
        future_label.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(future_label)

        card_layout.addStretch()
        outer.addWidget(self.card)
        self.setLayout(outer)

        # allow dragging the frameless window by its title bar
        self._drag_pos = None
        title_label.mousePressEvent = self._start_drag
        title_label.mouseMoveEvent = self._do_drag

    def _start_drag(self, event):
        self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()

    def _do_drag(self, event):
        if self._drag_pos is not None and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self._drag_pos)

    # ------------------------------------------------------------
    # STYLING (QSS — Qt's CSS-like stylesheet language)
    # ------------------------------------------------------------
    def _apply_styles(self):
        self.setStyleSheet("""
            #card {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #1e1b2e, stop:0.5 #2a2140, stop:1 #1a2438
                );
                border-radius: 18px;
                border: 1px solid rgba(255,255,255,0.06);
            }
            #titleLabel {
                color: #ffffff;
                font-size: 17px;
                font-weight: 700;
            }
            #subtitle {
                color: #9d95b8;
                font-size: 12px;
            }
            #closeBtn {
                background: transparent;
                color: #9d95b8;
                border: none;
                border-radius: 13px;
                font-size: 13px;
            }
            #closeBtn:hover {
                background: #ff5f57;
                color: white;
            }
            #statusText {
                color: #e6e1f5;
                font-size: 13px;
                font-weight: 500;
            }
            #statusDotOff {
                color: #6b6480;
                font-size: 14px;
            }
            #statusDotOn {
                color: #38f5b0;
                font-size: 14px;
            }
            #futureLabel {
                color: #635c7a;
                font-size: 11px;
                font-style: italic;
            }
            QProgressBar {
                background: rgba(255,255,255,0.06);
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #7C4DFF, stop:0.5 #18C6C6, stop:1 #FF6EC7
                );
                border-radius: 3px;
            }
            #primaryBtn {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #7C4DFF, stop:1 #18C6C6
                );
                color: white;
                font-size: 14px;
                font-weight: 600;
                border: none;
                border-radius: 10px;
                padding: 12px;
            }
            #primaryBtn:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #8d63ff, stop:1 #2fd9d9
                );
            }
            #primaryBtn:disabled {
                background: #3a3550;
                color: #7a7390;
            }
            #stopBtn {
                background: rgba(255,255,255,0.05);
                color: #e6e1f5;
                font-size: 13px;
                font-weight: 600;
                border: 1px solid rgba(255,255,255,0.12);
                border-radius: 10px;
                padding: 10px;
            }
            #stopBtn:hover:enabled {
                background: rgba(255, 95, 87, 0.15);
                border: 1px solid #ff5f57;
                color: #ff8a84;
            }
            #stopBtn:disabled {
                color: #4a4560;
            }
        """)

    # ------------------------------------------------------------
    # ANIMATIONS
    # ------------------------------------------------------------
    def _fade_in(self):
        self.opacity_anim = QPropertyAnimation(self, b"windowOpacity")
        self.opacity_anim.setDuration(400)
        self.opacity_anim.setStartValue(0)
        self.opacity_anim.setEndValue(1)
        self.opacity_anim.setEasingCurve(QEasingCurve.OutCubic)
        self.opacity_anim.start()

    def _start_pulse(self):
        self.pulse_timer = QTimer(self)
        self.pulse_timer.timeout.connect(self._toggle_pulse)
        self.pulse_timer.start(600)

    def _stop_pulse(self):
        if self.pulse_timer:
            self.pulse_timer.stop()
            self.pulse_timer = None
        self.status_dot.setObjectName("statusDotOff")
        self._refresh_dot_style()

    def _toggle_pulse(self):
        self._pulse_state = not self._pulse_state
        self.status_dot.setStyleSheet(
            "color: #38f5b0; font-size: 14px;" if self._pulse_state
            else "color: rgba(56,245,176,80); font-size: 14px;"
        )

    def _refresh_dot_style(self):
        # forces the stylesheet to re-apply after changing objectName
        self.status_dot.style().unpolish(self.status_dot)
        self.status_dot.style().polish(self.status_dot)

    # ------------------------------------------------------------
    # CORE LOGIC
    # ------------------------------------------------------------
    def launch_visualizer(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        app_path = os.path.join(base_dir, "app.py")

        self.process = subprocess.Popen(["streamlit", "run", app_path])

        self.status_label.setText("Starting…")
        self.progress.show()
        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)

        # poll localhost:8501 every 500ms until Streamlit actually responds,
        # instead of just assuming it's up right after Popen returns
        self.check_timer = QTimer(self)
        self.check_timer.timeout.connect(self._check_server_ready)
        self.check_timer.start(500)

    def _check_server_ready(self):
        try:
            urllib.request.urlopen(STREAMLIT_URL, timeout=0.5)
            self._on_server_ready()
        except Exception:
            pass  # not up yet — keep polling

    def _on_server_ready(self):
        self.check_timer.stop()
        self.check_timer = None
        self.progress.hide()
        self.status_label.setText("Running")
        self.status_dot.setObjectName("statusDotOn")
        self._refresh_dot_style()
        self._start_pulse()
        self.btn_open.setEnabled(True)
        self.open_in_browser()

    def open_in_browser(self):
        webbrowser.open(STREAMLIT_URL)

    def stop_visualizer(self):
        if self.check_timer:
            self.check_timer.stop()
            self.check_timer = None

        if self.process is not None:
            self.process.terminate()
            self.process = None

        self._stop_pulse()
        self.progress.hide()
        self.status_label.setText("Not running")
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_open.setEnabled(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LauncherApp()
    window.show()
    sys.exit(app.exec_())