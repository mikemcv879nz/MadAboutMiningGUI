# miner_gui.py
# Dynamic Miner GUI (PySide6) with:
# - Settings stored in AppData (persists across PyInstaller rebuilds)
# - Editable Miners table (manual typing works)
# - Browse button per miner (starts in ./miners, filters BAT/EXE)
# - XMRig auto-config option (writes config.generated.json next to xmrig.exe)
# - ANSI + keyword-based coloring for miner output
# - App name + Theme (accent color picker / QSS theme file)
# - Border color override (layers on top of theme)
# - Windows AppUserModelID + Qt app/window icon support
# - Robust logging to AppData + improved QProcess start/error reporting

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple
import time
import math

import sys
import os
import json
import datetime
import shlex
import re
import traceback
import subprocess
from pathlib import Path
import html

import psutil

# Windows-only: process suspend/resume helpers (used for XMRig pause)
if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

from PySide6.QtGui import QDesktopServices, QTextCursor, QIcon, QColor, QPainter, QPixmap, QTransform, QFont, QAction
from PySide6.QtCore import QProcess, QRectF, QSize, QStandardPaths, QTimer, QUrl, Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QTabWidget,
    QFrame, QMessageBox,
    QFormLayout, QCheckBox, QFileDialog, QToolButton,
    QDialog, QDialogButtonBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QLineEdit, QComboBox, QAbstractItemView,
    QSlider,
    QListWidget, QListWidgetItem, QSizePolicy, QInputDialog,
    QColorDialog, QSystemTrayIcon, QMenu)

# ---------------- Base directory (python run + PyInstaller onedir/onefile exe) ----------------
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent


# ---------------- Portable mode ----------------
def is_portable_mode() -> bool:
    """Portable mode stores settings/logs next to the EXE.

    Enabled when any of these are true:
      - A marker file exists next to the EXE: portable.flag / portable.txt / portable_mode
      - Environment variable MINERGUI_PORTABLE is 1/true/yes
      - A settings file already exists next to the EXE (miner_gui.settings.json)
    """
    try:
        env = (os.environ.get("MINERGUI_PORTABLE") or "").strip().lower()
        if env in ("1", "true", "yes", "y", "on"):
            return True
    except Exception:
        pass

    try:
        for name in ("portable.flag", "portable.txt", "portable_mode"):
            if (APP_DIR / name).exists():
                return True
    except Exception:
        pass

    try:
        if (APP_DIR / "miner_gui.settings.json").exists():
            return True
    except Exception:
        pass

    return False


def _dir_writable(folder: Path) -> bool:
    try:
        folder.mkdir(parents=True, exist_ok=True)
        test = folder / ".write_test"
        test.write_text("ok", encoding="utf-8")
        test.unlink(missing_ok=True)
        return True
    except Exception:
        return False



def resource_path(rel: str) -> str:
    """
    Absolute path to bundled resources for dev + PyInstaller.
    Works for onedir and onefile (sys._MEIPASS exists in onefile).
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return str(Path(base) / rel)
    return str((APP_DIR / rel).resolve())


# ---------------- Windows AppUserModelID ----------------
def set_windows_appusermodel_id(app_id: str) -> None:
    """
    Set Windows taskbar grouping/app identity. Must run BEFORE QApplication is created.
    """
    if sys.platform != "win32":
        return
    app_id = (app_id or "").strip()
    if not app_id:
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


# ---------------- Settings stored in AppData (persists across rebuilds) ----------------

def get_settings_path() -> Path:
    # Portable mode: keep settings beside the EXE/folder (useful for zipping/sharing)
    if is_portable_mode():
        if _dir_writable(APP_DIR):
            return APP_DIR / "miner_gui.settings.json"
    # Default: per-user AppData
    base = Path(QStandardPaths.writableLocation(QStandardPaths.AppDataLocation))
    base.mkdir(parents=True, exist_ok=True)
    return base / "miner_gui.settings.json"


SETTINGS_PATH = get_settings_path()


# ---------------- Logging to AppData ----------------

def get_log_path() -> Path:
    # Portable mode: keep logs beside the EXE/folder (useful for zipping/sharing)
    if is_portable_mode():
        if _dir_writable(APP_DIR):
            return APP_DIR / "miner_gui.log"
    # Default: per-user AppData
    base = Path(QStandardPaths.writableLocation(QStandardPaths.AppDataLocation))
    base.mkdir(parents=True, exist_ok=True)
    return base / "miner_gui.log"


LOG_PATH = get_log_path()


def log_to_file(msg: str) -> None:
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        LOG_PATH.write_text("", encoding="utf-8") if not LOG_PATH.exists() else None
        with LOG_PATH.open("a", encoding="utf-8", errors="replace") as f:
            f.write(f"{ts} {msg}\n")
    except Exception:
        # Never let logging crash the GUI
        pass


# ---------------- App name helpers ----------------
def effective_app_name(settings: dict) -> str:
    app_cfg = settings.get("app", {}) or {}
    if bool(app_cfg.get("use_custom_name", False)):
        nm = (app_cfg.get("custom_name") or "").strip()
        if nm:
            return nm
    return (app_cfg.get("display_name") or "Miner GUI").strip() or "Miner GUI"
# --------------------------- Data Models ---------------------------

# ---------------- Icon support ----------------
def get_icon_path() -> str | None:
    """
    Find a usable .ico file from likely locations.
    """
    candidates = [
        APP_DIR / "MinerGUI.ico",
        APP_DIR / "assets" / "MinerGUI.ico",
        APP_DIR / "assets" / "app.ico",
        Path(resource_path("MinerGUI.ico")),
        Path(resource_path("assets/MinerGUI.ico")),
        Path(resource_path("assets/app.ico")),
    ]
    for p in candidates:
        try:
            if p.exists():
                return str(p)
        except Exception:
            pass
    return None


def apply_icons(app: QApplication, win: QMainWindow) -> None:
    """
    Set BOTH application icon and window icon (titlebar + taskbar).
    Note: Windows taskbar icon also depends on the EXE icon at build time.
    """
    icon_path = get_icon_path()
    if not icon_path:
        log_to_file("Icon not found in expected locations (MinerGUI.ico / assets/MinerGUI.ico).")
        return

    icon = QIcon(icon_path)
    if icon.isNull():
        log_to_file(f"Icon file exists but failed to load (is it a valid .ico?): {icon_path}")
        return

    app.setWindowIcon(icon)
    win.setWindowIcon(icon)



# ---------------- Windows Defender exclusion helpers ----------------
def _norm_folder(p: str) -> str:
    try:
        pp = Path(p).expanduser()
        # Keep Windows drive casing stable; resolve if possible.
        if pp.exists():
            pp = pp.resolve()
        return str(pp)
    except Exception:
        return str(p or "").strip()


def add_defender_exclusion_admin(folder: str) -> tuple[bool, str]:
    """Add a Windows Defender exclusion for a folder.
    Requires admin. Shows a UAC prompt. Returns (ok, message).
    """
    if sys.platform != "win32":
        return False, "Windows Defender exclusions are only supported on Windows."
    folder = _norm_folder(folder)
    if not folder:
        return False, "No folder path provided."
    try:
        fp = Path(folder)
        if not fp.exists() or not fp.is_dir():
            return False, "Folder does not exist."
    except Exception:
        return False, "Invalid folder path."

    # Elevate PowerShell to run Add-MpPreference. We wait and propagate the ExitCode.
    # NOTE: This intentionally requires explicit user consent via UAC.
    ps_script = (
        "$p = \"" + folder.replace('"', '""') + "\"; "
        "$proc = Start-Process -FilePath PowerShell -Verb RunAs -Wait -PassThru "
        "-ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-Command', "
        "\"Add-MpPreference -ExclusionPath `\"$p`\"\" ); "
        "exit $proc.ExitCode"
    )
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if r.returncode == 0:
            return True, "Exclusion added."
        return False, (r.stderr or r.stdout or f"PowerShell exited with code {r.returncode}.").strip()
    except subprocess.TimeoutExpired:
        return False, "Timed out waiting for the elevated PowerShell process."
    except Exception as e:
        return False, f"Failed to run PowerShell: {e}"

# ---------------- Theme helpers ----------------
def read_text_file(path: str) -> str:
    try:
        return Path(path).read_text(encoding="utf-8")
    except Exception:
        return ""


def builtin_qss(mode: str, accent: str) -> str:
    accent = (accent or "#44ff44").strip()

    if mode == "builtin_light":
        bg = "#f3f3f3"
        panel = "#ffffff"
        fg = "#111111"
        border = "#cfcfcf"
    else:  # builtin_dark
        bg = "#111111"
        panel = "#1a1a1a"
        fg = "#dddddd"
        border = "#2a2a2a"

    return f"""
QWidget {{
    background: {bg};
    color: {fg};
    font-size: 10pt;
}}

QTextEdit {{
    background: {panel};
    color: {fg};
    border: 1px solid {border};
    border-radius: 6px;
}}

QLineEdit, QComboBox {{
    background: {panel};
    color: {fg};
    border: 1px solid {border};
    border-radius: 6px;
    padding: 4px;
}}

QPushButton {{
    background: {panel};
    color: {fg};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 6px 10px;
}}
QPushButton:hover {{
    border: 1px solid {accent};
}}
QPushButton:pressed {{
    border: 1px solid {accent};
}}

QTabWidget::pane {{
    border: 1px solid {border};
}}
QTabBar::tab {{
    background: {panel};
    border: 1px solid {border};
    padding: 6px 10px;
    margin-right: 2px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}}
QTabBar::tab:selected {{
    border-bottom: 2px solid {accent};
}}
"""


def apply_theme(app: QApplication, settings: dict) -> None:
    th = settings.get("theme", {}) or {}
    mode = (th.get("mode") or "builtin_dark").strip()
    accent = (th.get("accent") or "#44ff44").strip()
    qss_file = (th.get("qss_file") or "").strip()

    if mode == "custom_qss" and qss_file:
        qss = read_text_file(qss_file)
        qss = qss.replace("{ACCENT}", accent)
        if qss.strip():
            app.setStyleSheet(qss)
            return

    app.setStyleSheet(builtin_qss(mode, accent))


def apply_border_color(app: QApplication, settings: dict) -> None:
    """
    Apply ONLY border color styling (layers on top of the current theme).
    """
    bc = (settings.get("ui", {}).get("border_color") or "#44ff44").strip()

    qss = f"""
    /* Inputs */
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
        border: 1px solid {bc};
        border-radius: 6px;
        padding: 4px;
    }}

    /* Buttons */
    QPushButton {{
        border: 1px solid {bc};
        border-radius: 8px;
        padding: 6px 10px;
    }}
    QPushButton:hover {{
        border: 2px solid {bc};
    }}

    /* Tabs frame + selected tab */
    QTabWidget::pane {{
        border: 1px solid {bc};
        border-radius: 8px;
        top: -1px;
    }}
    QTabBar::tab:selected {{
        border-bottom: 2px solid {bc};
    }}

    /* Tables */
    QTableWidget {{
        border: 1px solid {bc};
        border-radius: 8px;
    }}

    /* Text areas */
    QTextEdit {{
        border: 1px solid {bc};
        border-radius: 8px;
    }}
    """

    app.setStyleSheet((app.styleSheet() or "") + qss)


# ---------------- ANSI + Color Support ----------------
ANSI_ESCAPE_RE = re.compile(r"\x1B\[[0-?]*[ -/]*[@-~]")
ANSI_SGR_RE = re.compile(r"\x1b\[([0-9;]*)m")

ANSI_30_37 = {
    30: "#000000",
    31: "#cc0000",
    32: "#00aa00",
    33: "#aa8800",
    34: "#0000cc",
    35: "#aa00aa",
    36: "#00aaaa",
    37: "#cccccc",
}
ANSI_90_97 = {
    90: "#666666",
    91: "#ff4444",
    92: "#44ff44",
    93: "#ffff44",
    94: "#4444ff",
    95: "#ff44ff",
    96: "#44ffff",
    97: "#ffffff",
}


def strip_ansi(s: str) -> str:
    return ANSI_ESCAPE_RE.sub("", s).replace("\x1b", "")


def ansi_to_html_line(s: str) -> str:
    """
    Convert a single line containing ANSI SGR codes into safe HTML.
    Supports:
      - reset (0)
      - 30-37 / 90-97 foreground colors
      - 38;5;N (256-color) foreground
    """
    def color_256(n: int) -> str:
        system = {
            0: "#000000", 1: "#800000", 2: "#008000", 3: "#808000",
            4: "#000080", 5: "#800080", 6: "#008080", 7: "#c0c0c0",
            8: "#808080", 9: "#ff0000", 10: "#00ff00", 11: "#ffff00",
            12: "#0000ff", 13: "#ff00ff", 14: "#00ffff", 15: "#ffffff"
        }
        if 0 <= n <= 15:
            return system[n]
        if 16 <= n <= 231:
            n -= 16
            r = (n // 36) % 6
            g = (n // 6) % 6
            b = n % 6

            def v(x: int) -> int:
                return 55 + x * 40 if x > 0 else 0

            return f"#{v(r):02x}{v(g):02x}{v(b):02x}"
        if 232 <= n <= 255:
            gray = 8 + (n - 232) * 10
            return f"#{gray:02x}{gray:02x}{gray:02x}"
        return "#cccccc"

    out: list[str] = []
    fg: str | None = None
    parts = ANSI_SGR_RE.split(s)  # text, codes, text, codes...

    for i in range(0, len(parts), 2):
        text = parts[i]
        if text:
            escaped = html.escape(text)
            out.append(f'<span style="color:{fg}">{escaped}</span>' if fg else escaped)

        if i + 1 < len(parts):
            code_str = parts[i + 1].strip()
            codes = [c for c in code_str.split(";") if c != ""] or ["0"]

            idx = 0
            while idx < len(codes):
                try:
                    c = int(codes[idx])
                except ValueError:
                    idx += 1
                    continue

                if c == 0:
                    fg = None
                elif c in ANSI_30_37:
                    fg = ANSI_30_37[c]
                elif c in ANSI_90_97:
                    fg = ANSI_90_97[c]
                elif c == 39:
                    fg = None
                elif c == 38:
                    if idx + 2 < len(codes) and codes[idx + 1] == "5":
                        try:
                            n = int(codes[idx + 2])
                            fg = color_256(n)
                        except ValueError:
                            pass
                        idx += 2
                idx += 1

    return f'<span style="font-family:Consolas, monospace; white-space: pre;">{"".join(out)}</span>'


XMRIG_RULES = [
    (re.compile(r"\baccepted\b", re.IGNORECASE), "#44ff44"),
    (re.compile(r"\bnew job\b", re.IGNORECASE), "#ff44ff"),
    (re.compile(r"\brejected\b", re.IGNORECASE), "#ff4444"),
    (re.compile(r"\berror\b|\bfailed\b|\bexception\b", re.IGNORECASE), "#ff4444"),
    (re.compile(r"\bdifficulty\b|\bhashrate\b|\bshare\b", re.IGNORECASE), "#ffff44"),
]


def line_to_html_with_rules(line: str, miner_id: str) -> str:
    if "\x1b[" in line:
        return ansi_to_html_line(line)

    clean = html.escape(line)
    if miner_id.lower() == "xmrig":
        for rx, color in XMRIG_RULES:
            if rx.search(line):
                return f'<span style="font-family:Consolas, monospace; white-space: pre; color:{color}">{clean}</span>'

    return f'<span style="font-family:Consolas, monospace; white-space: pre;">{clean}</span>'


def append_html(log_widget: QTextEdit, html_snippet: str) -> None:
    cursor = log_widget.textCursor()
    cursor.movePosition(QTextCursor.End)
    cursor.insertHtml(html_snippet)
    cursor.insertBlock()
    log_widget.setTextCursor(cursor)
    log_widget.ensureCursorVisible()


def now_local_12h() -> str:
    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%I:%M:%S %p").lstrip("0")
    return f"{date_str}  {time_str}"


def decode_process_bytes(b: bytes) -> str:
    """
    Best-effort decoding for miner output on Windows.
    """
    if not b:
        return ""
    # Try UTF-8 first
    try:
        return b.decode("utf-8")
    except UnicodeDecodeError:
        pass
    # Windows fallback
    if sys.platform == "win32":
        try:
            return b.decode("mbcs", errors="replace")
        except Exception:
            return b.decode(errors="replace")
    return b.decode(errors="replace")



# ---------------- Coin Rain (Promo Panel) ----------------
import random


def app_base_dir() -> Path:
    """Compatibility wrapper for older code paths."""
    return APP_DIR



@dataclass
class CoinParticle:
    pm: QPixmap
    x: float
    y: float
    vx: float
    vy: float
    angle: float
    vangle: float
    scale: float


class CoinRainWidget(QWidget):
    """
    Lightweight 'falling coins' animation for the right-side promo panel.
    Loads PNGs from APP_DIR / "assets/coins".
    """
    def __init__(self, pixmaps: list[QPixmap], parent: QWidget | None = None):
        super().__init__(parent)
        self._pixmaps = [p for p in pixmaps if not p.isNull()]
        self._parts: list[CoinParticle] = []

        self._size_scale = 0.5   # multiplied against per-particle random scale
        self._speed_scale = 1.0

        self._max_particles = 28
        self._spawn_prob = 0.35

        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(16)

        self.setMinimumHeight(360)

    def set_size_pct(self, pct: int) -> None:
        pct = max(5, min(150, int(pct)))
        self._size_scale = pct / 100.0

    def set_speed_pct(self, pct: int) -> None:
        pct = max(25, min(300, int(pct)))
        self._speed_scale = pct / 100.0

    def set_low_distraction(self, low: bool) -> None:
        if low:
            self._max_particles = 10
            self._spawn_prob = 0.12
        else:
            self._max_particles = 28
            self._spawn_prob = 0.35

    def _spawn(self) -> None:
        if not self._pixmaps:
            return
        pm = random.choice(self._pixmaps)
        w = max(1, self.width())
        x = random.random() * max(1, w - 32)
        y = -80 - random.random() * 120
        base_scale = 0.20 + random.random() * 0.55
        scale = base_scale * self._size_scale
        vy = (1.2 + random.random() * 3.8) * self._speed_scale
        vx = (-0.6 + random.random() * 1.2) * self._speed_scale
        angle = random.random() * 360.0
        vangle = (-2.0 + random.random() * 4.0) * self._speed_scale
        self._parts.append(CoinParticle(pm, x, y, vx, vy, angle, vangle, scale))

    def _tick(self) -> None:
        # Pause animation when not visible / minimized
        if not self.isVisible():
            return
        if len(self._parts) < self._max_particles and random.random() < self._spawn_prob:
            self._spawn()

        h = max(1, self.height())
        w = max(1, self.width())
        alive: list[CoinParticle] = []
        for p in self._parts:
            p.x += p.vx
            p.y += p.vy
            p.angle = (p.angle + p.vangle) % 360.0
            if p.y < h + 120:
                if p.x < -120:
                    p.x = w + 20
                elif p.x > w + 120:
                    p.x = -20
                alive.append(p)
        self._parts = alive
        self.update()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        try:
            painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform, True)
            painter.setOpacity(0.08)
            painter.fillRect(self.rect(), Qt.black)
            painter.setOpacity(1.0)

            if not self._pixmaps:
                painter.setOpacity(0.65)
                painter.drawText(self.rect(), Qt.AlignCenter, "Add PNGs to assets/coins")
                return

            for p in self._parts:
                pm = p.pm
                tw = pm.width() * p.scale
                th = pm.height() * p.scale
                if tw < 2 or th < 2:
                    continue
                cx = p.x + tw / 2.0
                cy = p.y + th / 2.0
                tr = QTransform()
                tr.translate(cx, cy)
                tr.rotate(p.angle)
                tr.translate(-cx, -cy)
                painter.setTransform(tr)
                painter.drawPixmap(QRectF(p.x, p.y, tw, th), pm, QRectF(pm.rect()))
                painter.resetTransform()
        finally:
            painter.end()



class MarqueeLabel(QWidget):
    """
    Simple single-line marquee that scrolls text horizontally.
    Direction: left-to-right by default (as requested).
    """
    def __init__(self, text: str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self._text = text
        self._offset = 0
        self._speed_px = 2  # pixels per tick
        self._direction = 1  # 1 = left->right, -1 = right->left

        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(30)

        self.setMinimumHeight(24)

    def setText(self, text: str) -> None:
        self._text = text or ""
        self._offset = 0
        self.update()

    def text(self) -> str:
        return self._text

    def setSpeed(self, px_per_tick: int) -> None:
        self._speed_px = max(1, int(px_per_tick))

    def setDirectionLeftToRight(self, enabled: bool = True) -> None:
        self._direction = 1 if enabled else -1

    def _tick(self) -> None:
        if not self._text:
            return
        self._offset += self._speed_px * self._direction
        # wrap around: when text fully leaves, reset
        fm = self.fontMetrics()
        text_w = fm.horizontalAdvance(self._text) + 40  # padding between repeats
        w = max(1, self.width())
        if self._direction == 1:
            if self._offset > w:
                self._offset = -text_w
        else:
            if self._offset < -text_w:
                self._offset = w
        self.update()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        try:
            painter.setRenderHint(QPainter.TextAntialiasing, True)
            painter.fillRect(self.rect(), Qt.transparent)

            if not self._text:
                return

            fm = self.fontMetrics()
            text_w = fm.horizontalAdvance(self._text)
            y = int((self.height() + fm.ascent() - fm.descent()) / 2)

            # draw repeated to ensure continuous look
            gap = 40
            total = text_w + gap
            x = self._offset
            while x < self.width() + total:
                painter.drawText(int(x), y, self._text)
                x += total
        finally:
            painter.end()



def load_coin_pixmaps():
    coin_dir = Path(resource_path("assets/coins"))
    pixmaps = []

    if not coin_dir.exists():
        return pixmaps

    for p in sorted(coin_dir.glob("*.png")):
        pm = QPixmap(str(p))
        if pm.isNull():
            continue
        pixmaps.append(pm)

    return pixmaps


# ---------------- Process Kill Helpers ----------------
def kill_process_tree(pid: int, include_parent: bool = True) -> None:
    try:
        parent = psutil.Process(pid)
    except psutil.Error:
        return

    for c in parent.children(recursive=True):
        try:
            c.kill()
        except psutil.Error:
            pass

    if include_parent:
        try:
            parent.kill()
        except psutil.Error:
            pass




# ---------------- Windows process suspend/resume (for Pause toggle) ----------------
def _open_process_for_suspend_resume(pid: int):
    if sys.platform != "win32":
        raise RuntimeError("Suspend/Resume only supported on Windows.")
    PROCESS_SUSPEND_RESUME = 0x0800
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    access = PROCESS_SUSPEND_RESUME | PROCESS_QUERY_LIMITED_INFORMATION
    handle = ctypes.windll.kernel32.OpenProcess(access, False, pid)
    if not handle:
        raise OSError(f"OpenProcess failed for PID {pid}")
    return handle

def suspend_process(pid: int) -> None:
    if sys.platform != "win32":
        raise RuntimeError("Suspend only supported on Windows.")
    handle = _open_process_for_suspend_resume(pid)
    try:
        ntdll = ctypes.windll.ntdll
        res = ntdll.NtSuspendProcess(handle)
        if res != 0:
            raise OSError(f"NtSuspendProcess failed (status {res})")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)

def resume_process(pid: int) -> None:
    if sys.platform != "win32":
        raise RuntimeError("Resume only supported on Windows.")
    handle = _open_process_for_suspend_resume(pid)
    try:
        ntdll = ctypes.windll.ntdll
        res = ntdll.NtResumeProcess(handle)
        if res != 0:
            raise OSError(f"NtResumeProcess failed (status {res})")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)

SAFE_EXE_NAME_RE = re.compile(r"^[A-Za-z0-9_.-]+\.exe$", re.IGNORECASE)


def kill_by_name(exe_name: str) -> int:
    exe_name = (exe_name or "").strip()
    if not SAFE_EXE_NAME_RE.match(exe_name):
        return 0
    exe_name_l = exe_name.lower()

    killed = 0
    for p in psutil.process_iter(["name"]):
        try:
            name = (p.info.get("name") or "").lower()
            if name == exe_name_l:
                p.kill()
                killed += 1
        except psutil.Error:
            pass
    return killed


# ---------------- Settings Schema ----------------
def _default_settings() -> dict:
    return {
        "ui": {
            "single_active": False,
            "border_color": "#44ff44",
            "coin_scale_pct": 50,
            "coin_speed_pct": 100,
            "low_distraction_while_mining": True,
            "promo_shrink_pct": 50,
            "promo_is_shrunk": False,
            "scripts_panel_collapsed": False,
            "auto_collapse_scripts_panel": True,
            "auto_expand_scripts_panel": True,
            "tray_enabled": True,
            "tray_show_hashrate": True,
            "tray_minimize_on_close": False,

        },
        "miners": [
            {
                "id": "xmrig",
                "name": "XMRig",
                "type": "EXE",
                "path": str((APP_DIR / "miners/xmrig/xmrig.exe").resolve()),
                "args": "--config {XMRIG_CONFIG_NAME}",
                "workdir": "",
                "kill_names": ["xmrig.exe"],
                "enabled": True,
                "scripts_dir": "",
                "active_scripts": []
            },
            {
                "id": "wildrig",
                "name": "WildRig",
                "type": "BAT",
                "path": "",
                "args": "",
                "workdir": "",
                "kill_names": ["wildrig.exe"],
                "enabled": True,
                "scripts_dir": "",
                "active_scripts": []
            },
            {
                "id": "rigel",
                "name": "Rigel",
                "type": "BAT",
                "path": "",
                "args": "",
                "workdir": "",
                "kill_names": ["rigel.exe"],
                "enabled": True,
                "scripts_dir": "",
                "active_scripts": []
            },
        ],
        "xmrig": {
            "auto_config": True,
            "pool": "",
            "wallet": "",
            "worker": "",
            "pass": "x",
            "extra": "",
        },
        "app": {
            "display_name": "Miner GUI",
            "use_custom_name": False,
            "custom_name": "",
            "app_id": "MinerGUI.Michael.Mining"
        },
        "theme": {
            "mode": "builtin_dark",
            "accent": "#44ff44",
            "qss_file": ""
        },
        "security": {
            "prompt_defender_exclusions": True,
            "defender_exclusions": []
        },
    }


def _deep_merge(base: dict, overlay: dict) -> dict:
    for k, v in overlay.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            base[k] = _deep_merge(base[k], v)
        else:
            base[k] = v
    return base


def migrate_settings(settings: dict) -> dict:
    """Back-compat + script-folder support for BAT miners.

    New keys (per miner):
      - scripts_dir: folder containing .bat files (used for BAT miners)
      - active_scripts: list[str] of checked .bat filenames (relative: just filename)
    """
    try:
        miners = settings.get("miners", []) or []
        for m in miners:
            mtype = (m.get("type") or "EXE").upper().strip()
            if mtype != "BAT":
                # Keep EXE miners untouched (but ensure keys exist for future)
                m.setdefault("scripts_dir", "")
                m.setdefault("active_scripts", [])
                continue

            # Legacy path could be: a BAT file OR a folder
            raw_path = (m.get("scripts_dir") or m.get("path") or "").strip()

            scripts_dir = ""
            active: list[str] = list(m.get("active_scripts") or [])

            try:
                p = Path(raw_path) if raw_path else None
                if p and p.exists():
                    if p.is_file():
                        scripts_dir = str(p.parent.resolve())
                        if not active:
                            active = [p.name]
                    elif p.is_dir():
                        scripts_dir = str(p.resolve())
                else:
                    # If empty/missing, try default ./miners/<id>
                    mid = str(m.get("id") or "").strip()
                    if mid:
                        cand = (APP_DIR / "miners" / mid)
                        if cand.exists() and cand.is_dir():
                            scripts_dir = str(cand.resolve())
            except Exception:
                pass

            m["scripts_dir"] = scripts_dir
            m["active_scripts"] = active

            # For UI consistency, keep "path" as the scripts folder for BAT miners
            if scripts_dir:
                m["path"] = scripts_dir
    except Exception:
        pass
    return settings



def discover_miners_from_folder(miners_root: Path) -> list[dict]:
    """Optional auto-discovery: miners/<id>/miner.json -> miner definition.
    This is additive (won't overwrite existing miner IDs).
    miner.json example:
      {
        "id": "srbminer",
        "name": "SRBMiner",
        "type": "BAT",
        "scripts_dir": "miners/srbminer",
        "workdir": "miners/srbminer",
        "kill_names": ["SRBMiner-MULTI.exe"],
        "enabled": true
      }
    Paths may be absolute or relative to APP_DIR.
    """
    out: list[dict] = []
    try:
        if not miners_root.exists():
            return out
        for d in sorted([p for p in miners_root.iterdir() if p.is_dir()]):
            cfg = d / "miner.json"
            if not cfg.exists():
                continue
            try:
                data = json.loads(cfg.read_text(encoding="utf-8"))
            except Exception:
                continue

            mid = str(data.get("id") or d.name).strip()
            if not mid:
                continue
            mtype = str(data.get("type") or "BAT").upper().strip()
            name = str(data.get("name") or mid).strip() or mid

            def _norm_path(v: str) -> str:
                v = (v or "").strip()
                if not v:
                    return ""
                p = Path(v)
                if not p.is_absolute():
                    p = (APP_DIR / p).resolve()
                return str(p)

            miner: dict = {
                "id": mid,
                "name": name,
                "type": ("EXE" if mtype == "EXE" else "BAT"),
                "path": _norm_path(str(data.get("path") or "")),
                "args": str(data.get("args") or ""),
                "workdir": _norm_path(str(data.get("workdir") or "")),
                "kill_names": list(data.get("kill_names") or []),
                "enabled": bool(data.get("enabled", True)),
                "scripts_dir": _norm_path(str(data.get("scripts_dir") or "")),
                "active_scripts": list(data.get("active_scripts") or []),
            }

            # Sensible defaults:
            if miner["type"] == "BAT":
                if not miner["scripts_dir"]:
                    miner["scripts_dir"] = str(d.resolve())
                miner["path"] = miner["scripts_dir"]  # UI/back-compat
                if not miner["workdir"]:
                    miner["workdir"] = miner["scripts_dir"]
            else:
                if not miner["path"]:
                    # try miner exe in folder
                    exes = sorted(d.glob("*.exe"))
                    if exes:
                        miner["path"] = str(exes[0].resolve())
                if not miner["workdir"] and miner["path"]:
                    miner["workdir"] = str(Path(miner["path"]).parent.resolve())

            out.append(miner)
    except Exception:
        pass
    return out


def merge_discovered_miners(settings: dict) -> dict:
    miners = settings.get("miners", []) or []
    existing_ids = {str(m.get("id","")).strip() for m in miners if str(m.get("id","")).strip()}
    discovered = discover_miners_from_folder(APP_DIR / "miners")
    for m in discovered:
        if str(m.get("id","")).strip() not in existing_ids:
            miners.append(m)
    settings["miners"] = miners
    return settings

def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        return merge_discovered_miners(migrate_settings(_default_settings()))
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        base = _default_settings()
        merged = _deep_merge(base, data)
        return merge_discovered_miners(migrate_settings(merged))
    except Exception:
        return merge_discovered_miners(migrate_settings(_default_settings()))


def save_settings(settings: dict) -> None:
    SETTINGS_PATH.write_text(json.dumps(settings, indent=2), encoding="utf-8")


# ---------------- Argument Expansion ----------------
def expand_placeholders(s: str, xmrig_config_path: Path | None) -> str:
    """
    Supported placeholders:
      {APP_DIR} {DATE} {TIME} {DATETIME}
      {XMRIG_CONFIG} full path
      {XMRIG_CONFIG_NAME} filename only
    Also expands environment variables like %USERPROFILE%.
    """
    now = datetime.datetime.now()
    repl = {
        "{APP_DIR}": str(APP_DIR),
        "{DATE}": now.strftime("%Y-%m-%d"),
        "{TIME}": now.strftime("%H:%M:%S"),
        "{DATETIME}": now.strftime("%Y-%m-%d_%H-%M-%S"),
        "{XMRIG_CONFIG}": str(xmrig_config_path) if xmrig_config_path else "",
        "{XMRIG_CONFIG_NAME}": xmrig_config_path.name if xmrig_config_path else "config.generated.json",
    }

    out = s or ""
    for k, v in repl.items():
        out = out.replace(k, v)

    return os.path.expandvars(out)


def split_args(arg_str: str) -> list[str]:
    arg_str = (arg_str or "").strip()
    if not arg_str:
        return []
    return shlex.split(arg_str, posix=False)


# ---------------- Settings Dialog ----------------
class SettingsDialog(QDialog):
    COL_ENABLED = 0
    COL_ID = 1
    COL_NAME = 2
    COL_TYPE = 3
    COL_PATH = 4
    COL_BROWSE = 5
    COL_ARGS = 6
    COL_WORKDIR = 7
    COL_KILL = 8

    def __init__(self, parent: QWidget, settings: dict):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.settings = json.loads(json.dumps(settings))  # deep copy
        self.setMinimumWidth(1020)

        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)

        self._build_miners_tab()
        self._build_xmrig_tab()
        self._build_app_theme_tab()

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _build_miners_tab(self):
        tab = QWidget()
        v = QVBoxLayout(tab)

        help_lbl = QLabel(
            "Add/edit miners here. Each miner can be EXE or BAT.\n"
            "Browse starts in the ./miners folder by default.\n"
            "Args supports: {DATE} {TIME} {DATETIME} {APP_DIR}\n"
            "Kill names are optional (comma-separated exe names) for fallback stopping."
        )
        help_lbl.setWordWrap(True)
        v.addWidget(help_lbl)

        self.single_active = QCheckBox("Single-active mode (starting one miner stops others)")
        self.single_active.setChecked(bool(self.settings.get("ui", {}).get("single_active", False)))
        v.addWidget(self.single_active)

        ui = self.settings.get("ui", {})
        self.border_color_edit = QLineEdit(str(ui.get("border_color", "#44ff44")))
        self.border_color_btn = QPushButton("Pick…")
        self.border_color_btn.clicked.connect(self._pick_border_color)

        row = QHBoxLayout()
        row.addWidget(self.border_color_edit, 1)
        row.addWidget(self.border_color_btn)

        v.addWidget(QLabel("Border color (hex):"))
        v.addLayout(row)

        self.miner_table = QTableWidget(0, 9)
        self.miner_table.setHorizontalHeaderLabels([
            "Enabled", "ID", "Name", "Type", "Path (EXE) / Scripts Folder (BAT)", "Browse", "Args", "Workdir (optional)", "Kill names (csv)"
        ])
        self.miner_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.miner_table.horizontalHeader().setSectionResizeMode(self.COL_ENABLED, QHeaderView.ResizeToContents)
        self.miner_table.horizontalHeader().setSectionResizeMode(self.COL_ID, QHeaderView.ResizeToContents)
        self.miner_table.horizontalHeader().setSectionResizeMode(self.COL_TYPE, QHeaderView.ResizeToContents)
        self.miner_table.horizontalHeader().setSectionResizeMode(self.COL_BROWSE, QHeaderView.ResizeToContents)

        self.miner_table.setEditTriggers(
            QAbstractItemView.DoubleClicked |
            QAbstractItemView.SelectedClicked |
            QAbstractItemView.EditKeyPressed |
            QAbstractItemView.AnyKeyPressed
        )
        self.miner_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.miner_table.setSelectionMode(QAbstractItemView.SingleSelection)

        v.addWidget(self.miner_table, 1)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Miner")
        del_btn = QPushButton("Delete Miner")
        up_btn = QPushButton("Move Up")
        dn_btn = QPushButton("Move Down")
        btn_row.addWidget(add_btn)
        btn_row.addWidget(del_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(up_btn)
        btn_row.addWidget(dn_btn)
        v.addLayout(btn_row)


        add_btn.clicked.connect(self._miners_add)
        del_btn.clicked.connect(self._miners_del)
        up_btn.clicked.connect(lambda: self._miners_move(-1))
        dn_btn.clicked.connect(lambda: self._miners_move(1))

        self._miners_rebuild_table(self.settings.get("miners", []) or [])
        self.tabs.addTab(tab, "Miners")

    def _pick_border_color(self) -> None:
        seed = QColor(self.border_color_edit.text().strip() or "#44ff44")
        col = QColorDialog.getColor(seed, self, "Pick border color")
        if col.isValid():
            self.border_color_edit.setText(col.name())

    def _build_xmrig_tab(self):
        tab = QWidget()
        form = QFormLayout(tab)

        xm = self.settings.get("xmrig", {}) or {}
        self.xm_auto = QCheckBox("Auto-generate config.generated.json from fields below")
        self.xm_auto.setChecked(bool(xm.get("auto_config", True)))

        self.x_pool = QLineEdit(str(xm.get("pool", "")))
        self.x_wallet = QLineEdit(str(xm.get("wallet", "")))
        self.x_worker = QLineEdit(str(xm.get("worker", "")))
        self.x_pass = QLineEdit(str(xm.get("pass", "x") or "x"))
        self.x_extra = QLineEdit(str(xm.get("extra", "")))
        self.x_extra.setPlaceholderText("Extra args appended to the miner's Args (optional)")

        form.addRow("", self.xm_auto)
        form.addRow("Pool:", self.x_pool)
        form.addRow("Wallet/User:", self.x_wallet)
        form.addRow("Worker:", self.x_worker)
        form.addRow("Password:", self.x_pass)
        form.addRow("Extra args:", self.x_extra)

        self.tabs.addTab(tab, "XMRig")

    def _build_app_theme_tab(self):
        tab = QWidget()
        form = QFormLayout(tab)

        app_cfg = self.settings.get("app", {}) or {}
        theme_cfg = self.settings.get("theme", {}) or {}

        self.app_display_name = QLineEdit(str(app_cfg.get("display_name", "Miner GUI")))
        self.app_use_custom = QCheckBox("Use custom app name (overrides default)")
        self.app_use_custom.setChecked(bool(app_cfg.get("use_custom_name", False)))
        self.app_custom_name = QLineEdit(str(app_cfg.get("custom_name", "")))

        def _toggle_custom_name():
            self.app_custom_name.setEnabled(self.app_use_custom.isChecked())

        self.app_use_custom.stateChanged.connect(_toggle_custom_name)
        _toggle_custom_name()

        form.addRow("App name (default):", self.app_display_name)
        form.addRow("", self.app_use_custom)
        form.addRow("Custom app name:", self.app_custom_name)

        # Coin animation
        ui = self.settings.get("ui", {}) or {}
        self.coin_scale = QSlider(Qt.Horizontal)
        self.coin_scale.setRange(5, 150)
        self.coin_scale.setValue(int(ui.get("coin_scale_pct", 50)))
        self.coin_speed = QSlider(Qt.Horizontal)
        self.coin_speed.setRange(25, 300)
        self.coin_speed.setValue(int(ui.get("coin_speed_pct", 100)))
        self.low_distraction = QCheckBox("Low distraction while mining (fewer coins)")
        self.low_distraction.setChecked(bool(ui.get("low_distraction_while_mining", True)))

        form.addRow(QLabel("Coin animation"), QLabel(""))
        form.addRow("Coin size (%):", self.coin_scale)
        form.addRow("Coin drop speed (%):", self.coin_speed)
        form.addRow("", self.low_distraction)

        # Promo panel width shrink
        self.promo_shrink = QSlider(Qt.Horizontal)
        self.promo_shrink.setRange(20, 100)
        self.promo_shrink.setValue(int(ui.get("promo_shrink_pct", 50)))
        self.promo_shrink.setToolTip("When the arrow is clicked, the promo panel width shrinks to this % of normal.")
        form.addRow("Promo panel width when shrunk (%):", self.promo_shrink)


        self.theme_mode = QComboBox()
        self.theme_mode.addItems(["builtin_dark", "builtin_light", "custom_qss"])
        self.theme_mode.setCurrentText(str(theme_cfg.get("mode", "builtin_dark")))

        self.theme_accent = QLineEdit(str(theme_cfg.get("accent", "#44ff44")))
        pick_btn = QPushButton("Pick…")
        pick_btn.clicked.connect(self._pick_accent_color)

        accent_row = QHBoxLayout()
        accent_row.addWidget(self.theme_accent, 1)
        accent_row.addWidget(pick_btn)
        accent_wrap = QWidget()
        accent_wrap.setLayout(accent_row)

        self.theme_qss_file = QLineEdit(str(theme_cfg.get("qss_file", "")))
        browse_btn = QPushButton("Browse…")
        browse_btn.clicked.connect(self._browse_qss_file)

        qss_row = QHBoxLayout()
        qss_row.addWidget(self.theme_qss_file, 1)
        qss_row.addWidget(browse_btn)
        qss_wrap = QWidget()
        qss_wrap.setLayout(qss_row)

        form.addRow("Theme mode:", self.theme_mode)
        form.addRow("Accent color:", accent_wrap)
        form.addRow("Custom QSS file:", qss_wrap)

        help_lbl = QLabel(
            "Notes:\n"
            "- builtin_dark / builtin_light use Accent color for highlights.\n"
            "- custom_qss applies your .qss file. You may use {ACCENT} inside the file.\n"
            "- Theme and app name changes apply after you click Save."
        )
        help_lbl.setWordWrap(True)
        form.addRow(help_lbl)

        self.tabs.addTab(tab, "App + Theme")

    def _pick_accent_color(self):
        current = (self.theme_accent.text() or "").strip()
        seed = QColor(current) if current else QColor("#44ff44")
        col = QColorDialog.getColor(seed, self, "Pick accent color")
        if col.isValid():
            self.theme_accent.setText(col.name())

    def _browse_qss_file(self):
        chosen, _ = QFileDialog.getOpenFileName(
            self,
            "Select QSS theme file",
            str(APP_DIR),
            "QSS files (*.qss);;All files (*.*)"
        )
        if chosen:
            self.theme_qss_file.setText(chosen)

    def _miners_rebuild_table(self, miners: list[dict]) -> None:
        self.miner_table.setRowCount(0)
        for m in miners:
            self._miners_add_row(m)
    def _maybe_prompt_defender_exclusion(self, chosen_path: str, miner_type: str) -> None:
        """Prompt the user to add a Windows Defender exclusion for a miner folder.

        - For EXE selection: excludes the EXE's parent folder.
        - For BAT selection: excludes the chosen scripts folder.
        Requires admin and user approval (UAC prompt).
        """
        if sys.platform != "win32":
            return

        sec = self.settings.get("security", {}) or {}
        if not bool(sec.get("prompt_defender_exclusions", True)):
            return

        miner_type = (miner_type or "").upper().strip()
        target = ""
        try:
            p = Path(chosen_path)
            target = str((p if (miner_type == "BAT") else p.parent).resolve())
        except Exception:
            return

        if not target:
            return

        excluded = set([_norm_folder(x) for x in (sec.get("defender_exclusions", []) or [])])
        if _norm_folder(target) in excluded:
            return

        msg = (
            "Windows Defender sometimes quarantines miners downloaded from the internet.\n\n"
            f"Add a Defender exclusion for this folder?\n\n{target}\n\n"
            "This requires Administrator approval (UAC prompt)."
        )
        ans = QMessageBox.question(self, "Add Defender exclusion?", msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans != QMessageBox.Yes:
            return

        ok, detail = add_defender_exclusion_admin(target)
        if ok:
            sec.setdefault("defender_exclusions", [])
            sec["defender_exclusions"].append(_norm_folder(target))
            self.settings["security"] = sec
            QMessageBox.information(self, "Defender exclusion", "Exclusion added successfully.")
        else:
            QMessageBox.warning(self, "Defender exclusion", f"Failed to add exclusion.\n\n{detail}")


    def _miners_add_row(self, miner: dict | None = None):
        miner = miner or {
            "id": "",
            "name": "",
            "type": "EXE",
            "path": "",
            "args": "",
            "workdir": "",
            "kill_names": [],
            "enabled": True
        }

        row = self.miner_table.rowCount()
        self.miner_table.insertRow(row)

        enabled_cb = QCheckBox()
        enabled_cb.setChecked(bool(miner.get("enabled", True)))
        self.miner_table.setCellWidget(row, self.COL_ENABLED, enabled_cb)

        id_item = QTableWidgetItem(str(miner.get("id", "")))
        name_item = QTableWidgetItem(str(miner.get("name", "")))

        type_combo = QComboBox()
        type_combo.addItems(["EXE", "BAT"])
        t = (miner.get("type") or "EXE").upper()
        type_combo.setCurrentText("BAT" if t == "BAT" else "EXE")

        path_item = QTableWidgetItem(str(miner.get("path", "")))
        args_item = QTableWidgetItem(str(miner.get("args", "")))
        wd_item = QTableWidgetItem(str(miner.get("workdir", "")))
        kills = miner.get("kill_names", []) or []
        kill_item = QTableWidgetItem(", ".join([str(x) for x in kills]))

        self.miner_table.setItem(row, self.COL_ID, id_item)
        self.miner_table.setItem(row, self.COL_NAME, name_item)
        self.miner_table.setCellWidget(row, self.COL_TYPE, type_combo)
        self.miner_table.setItem(row, self.COL_PATH, path_item)

        browse_btn = QPushButton("Browse…")
        browse_btn.clicked.connect(self._browse_clicked)
        self.miner_table.setCellWidget(row, self.COL_BROWSE, browse_btn)

        self.miner_table.setItem(row, self.COL_ARGS, args_item)
        self.miner_table.setItem(row, self.COL_WORKDIR, wd_item)
        self.miner_table.setItem(row, self.COL_KILL, kill_item)

    def _browse_clicked(self):
        btn = self.sender()
        if btn is None:
            return
        for r in range(self.miner_table.rowCount()):
            if self.miner_table.cellWidget(r, self.COL_BROWSE) is btn:
                self._browse_miner_path(r)
                return

    def _browse_miner_path(self, row: int) -> None:
        type_widget = self.miner_table.cellWidget(row, self.COL_TYPE)
        mtype = type_widget.currentText().strip().upper() if isinstance(type_widget, QComboBox) else "EXE"

        path_item = self.miner_table.item(row, self.COL_PATH)
        current_path = (path_item.text().strip() if path_item else "")

        miners_root = (APP_DIR / "miners")
        downloads = Path(QStandardPaths.writableLocation(QStandardPaths.DownloadLocation))

        # Default browse start:
        # 1) current path folder (if set and exists)
        # 2) ./miners (if exists)
        # 3) Downloads (if exists) else Home
        start_dir = (
            miners_root if miners_root.exists()
            else (downloads if downloads.exists() else Path.home())
        )

        try:
            if current_path:
                cp = Path(current_path)
                if cp.exists():
                    start_dir = cp.parent if cp.is_file() else cp
        except Exception:
            pass

        chosen = ""
        if mtype == "BAT":
            # For BAT miners we select a FOLDER containing scripts, not a single .bat file.
            chosen = QFileDialog.getExistingDirectory(self, "Select scripts folder", str(start_dir))
        else:
            flt = "Executable files (*.exe);;All files (*.*)"
            chosen, _ = QFileDialog.getOpenFileName(self, "Select EXE file", str(start_dir), flt)

        if not chosen:
            return

        if not path_item:
            path_item = QTableWidgetItem("")
            self.miner_table.setItem(row, self.COL_PATH, path_item)
        path_item.setText(chosen)

        # Default workdir to the chosen folder (or file's parent) if blank
        wd_item = self.miner_table.item(row, self.COL_WORKDIR)
        wd_text = wd_item.text().strip() if wd_item else ""
        if not wd_text:
            if not wd_item:
                wd_item = QTableWidgetItem("")
                self.miner_table.setItem(row, self.COL_WORKDIR, wd_item)
            p = Path(chosen)
            wd_item.setText(str((p if p.is_dir() else p.parent).resolve()))

        # Optional: prompt to add Windows Defender exclusion for miner folder
        try:
            self._maybe_prompt_defender_exclusion(chosen, mtype)
        except Exception:
            pass


    def _miners_add(self):
        self._miners_add_row(None)

    def _miners_del(self):
        row = self.miner_table.currentRow()
        if row >= 0:
            self.miner_table.removeRow(row)

    def _miners_move(self, direction: int):
        row = self.miner_table.currentRow()
        if row < 0:
            return
        new_row = row + direction
        if new_row < 0 or new_row >= self.miner_table.rowCount():
            return

        miners = self._collect_miners_preview(include_blank_rows=True)
        miners[row], miners[new_row] = miners[new_row], miners[row]
        self._miners_rebuild_table(miners)
        self.miner_table.setCurrentCell(new_row, self.COL_NAME)

    def _collect_miners_preview(self, include_blank_rows: bool = False) -> list[dict]:
        miners: list[dict] = []
        for r in range(self.miner_table.rowCount()):
            enabled_widget = self.miner_table.cellWidget(r, self.COL_ENABLED)
            enabled = enabled_widget.isChecked() if isinstance(enabled_widget, QCheckBox) else True

            mid = (self.miner_table.item(r, self.COL_ID).text() if self.miner_table.item(r, self.COL_ID) else "").strip()
            name = (self.miner_table.item(r, self.COL_NAME).text() if self.miner_table.item(r, self.COL_NAME) else "").strip()

            type_widget = self.miner_table.cellWidget(r, self.COL_TYPE)
            mtype = type_widget.currentText().strip() if isinstance(type_widget, QComboBox) else "EXE"

            path = (self.miner_table.item(r, self.COL_PATH).text() if self.miner_table.item(r, self.COL_PATH) else "").strip()
            args = (self.miner_table.item(r, self.COL_ARGS).text() if self.miner_table.item(r, self.COL_ARGS) else "").strip()
            wd = (self.miner_table.item(r, self.COL_WORKDIR).text() if self.miner_table.item(r, self.COL_WORKDIR) else "").strip()
            kills = (self.miner_table.item(r, self.COL_KILL).text() if self.miner_table.item(r, self.COL_KILL) else "").strip()
            kill_names = [x.strip() for x in kills.split(",") if x.strip()]

            if not include_blank_rows and (not name and not mid and not path):
                continue

            prev = None
            try:
                for pm in (self.settings.get("miners", []) or []):
                    if str(pm.get("id", "")).strip() and str(pm.get("id", "")).strip() == mid:
                        prev = pm
                        break
            except Exception:
                prev = None

            scripts_dir = str((prev or {}).get("scripts_dir", "") or "")
            active_scripts = list((prev or {}).get("active_scripts", []) or [])

            # For BAT miners we store the scripts folder in path as well (UI consistency)
            if mtype.upper().strip() == "BAT":
                scripts_dir = path

            miners.append({
                "enabled": enabled,
                "id": mid,
                "name": name,
                "type": mtype,
                "path": path,
                "args": args,
                "workdir": wd,
                "kill_names": kill_names,
                "scripts_dir": scripts_dir,
                "active_scripts": active_scripts
            })
        return miners
    def _miners_get_selected_script_path(self) -> Path | None:
        try:
            row = self.miner_table.currentRow()
            if row < 0:
                return None
            item = self.miner_table.item(row, self.COL_PATH)
            if item is None:
                return None
            raw = (item.text() or "").strip()
            if not raw:
                return None
            p = Path(raw)
            if not p.is_absolute():
                p = (APP_DIR / p).resolve()
            return p
        except Exception:
            return None

    def _miners_open_selected(self) -> None:
        p = self._miners_get_selected_script_path()
        if p is None or not p.exists():
            QMessageBox.warning(self, "Open selected script", "No valid miner script selected.")
            return
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(p)))
        except Exception:
            QMessageBox.warning(self, "Open selected script", "Failed to open the selected file.")

    def _miners_edit_selected(self) -> None:
        p = self._miners_get_selected_script_path()
        if p is None or not p.exists():
            QMessageBox.warning(self, "Edit selected script", "No valid miner script selected.")
            return
        if p.suffix.lower() not in (".bat", ".cmd", ".ps1", ".sh", ".json", ".txt"):
            QMessageBox.warning(self, "Edit selected script", "Selected file type is not supported for editing.")
            return
        try:
            subprocess.Popen(["notepad.exe", str(p)])
        except Exception:
            QMessageBox.warning(self, "Edit selected script", "Failed to open Notepad.")



    def collect_settings(self) -> dict:
        out = json.loads(json.dumps(self.settings))

        # UI
        out["ui"] = out.get("ui", {})
        out["ui"]["single_active"] = self.single_active.isChecked()
        out["ui"]["border_color"] = (self.border_color_edit.text().strip() or "#44ff44")
        out["ui"]["coin_scale_pct"] = int(self.coin_scale.value())
        out["ui"]["coin_speed_pct"] = int(self.coin_speed.value())
        out["ui"]["low_distraction_while_mining"] = bool(self.low_distraction.isChecked())
        out["ui"]["promo_shrink_pct"] = int(self.promo_shrink.value())

        # XMRig
        out["xmrig"] = {
            "auto_config": self.xm_auto.isChecked(),
            "pool": self.x_pool.text().strip(),
            "wallet": self.x_wallet.text().strip(),
            "worker": self.x_worker.text().strip(),
            "pass": (self.x_pass.text().strip() or "x"),
            "extra": self.x_extra.text().strip(),
        }

        # App
        out["app"] = out.get("app", {})
        out["app"]["display_name"] = (self.app_display_name.text().strip() or "Miner GUI")
        out["app"]["use_custom_name"] = bool(self.app_use_custom.isChecked())
        out["app"]["custom_name"] = self.app_custom_name.text().strip()

        # Theme
        out["theme"] = out.get("theme", {})
        out["theme"]["mode"] = self.theme_mode.currentText().strip()
        out["theme"]["accent"] = (self.theme_accent.text().strip() or "#44ff44")
        out["theme"]["qss_file"] = self.theme_qss_file.text().strip()

        # Miners
        miners = self._collect_miners_preview(include_blank_rows=False)

        # Ensure unique IDs
        used = set()
        for m in miners:
            mid = (m.get("id") or "").strip()
            if not mid:
                base = re.sub(r"[^a-z0-9]+", "", (m.get("name") or "").lower()) or "miner"
                cand = base
                i = 2
                while cand in used:
                    cand = f"{base}{i}"
                    i += 1
                mid = cand
            if mid in used:
                base = mid
                i = 2
                cand = f"{base}{i}"
                while cand in used:
                    i += 1
                    cand = f"{base}{i}"
                mid = cand
            used.add(mid)
            m["id"] = mid

        out["miners"] = miners
        return out


# ---------------- Main App ----------------
class MinerGUI(QMainWindow):
    def __init__(self, settings: dict):
        super().__init__()
        self.profit_engine = None  # profit switching disabled

        icon_path = Path(resource_path("assets/app.ico"))
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.settings = settings
        self.setWindowTitle(effective_app_name(self.settings))

        self.procs: dict[str, QProcess | None] = {}
        self.proc_pids: dict[str, int | None] = {}

        self.proc_paused: dict[str, bool] = {}
        self.proc_last_error: dict[str, str | None] = {}
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        top = QHBoxLayout()
        root.addLayout(top)

        self.time_label = QLabel("Time: --")
        top.addWidget(self.time_label)
        top.addStretch(1)

        self.settings_btn = QToolButton()
        self.settings_btn.setText("⚙")
        self.settings_btn.setToolTip("Settings")
        self.settings_btn.clicked.connect(self.open_settings)
        top.addWidget(self.settings_btn)

        main = QHBoxLayout()
        root.addLayout(main, 1)

        self.tabs = QTabWidget()
        main.addWidget(self.tabs, 1)

        # Middle panel: script selector (BAT folders + checkbox activation)
        self.scripts_panel = self._build_scripts_panel()
        main.addWidget(self.scripts_panel, 0)

        # Right-side promo panel (coin rain)
        self.promo_panel = self._build_promo_panel()
        main.addWidget(self.promo_panel, 0)
        self._promo_apply_width()

        self.miner_tabs: dict[str, dict] = {}
        self.global_log = QTextEdit()
        self.global_log.setReadOnly(True)

        self._rebuild_miner_tabs()

        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self._tick_clock)
        self.clock_timer.start(1000)
        self._tick_clock()
        self._apply_coin_settings()
        self._promo_apply_width()

        self._update_all_buttons()
        self._init_tray()

    def _miners_list(self) -> list[dict]:
        return self.settings.get("miners", []) or []

    def _rebuild_miner_tabs(self):
        self.tabs.clear()
        self.miner_tabs.clear()

        enabled_miners = [m for m in self._miners_list() if bool(m.get("enabled", True))]

        for m in enabled_miners:
            mid = str(m.get("id", "")).strip() or "miner"
            self.procs.setdefault(mid, None)
            self.proc_pids.setdefault(mid, None)

            self.proc_paused.setdefault(mid, False)
            self.proc_last_error.setdefault(mid, None)
        for m in enabled_miners:
            mid = str(m.get("id", "")).strip()
            name = str(m.get("name", mid)).strip() or mid
            tab, widgets = self._build_miner_tab(mid, name)
            self.miner_tabs[mid] = widgets
            self.tabs.addTab(tab, name)

        self.tabs.addTab(self.global_log, "Global Logs")

        # Keep scripts panel miner list in sync
        try:
            self._scripts_refresh_miner_combo()
        except Exception:
            pass

    def _build_miner_tab(self, miner_id: str, display_name: str):
        w = QWidget()
        w.setProperty("miner_id", miner_id)
        layout = QVBoxLayout(w)

        btn_row = QHBoxLayout()
        start = QPushButton(f"Start {display_name}")
        stop = QPushButton(f"Stop {display_name}")
        kill = QPushButton(f"Force Kill {display_name}")

        pause = None
        if miner_id.lower() == "xmrig":
            pause = QPushButton(f"Pause {display_name}")
            btn_row.addWidget(pause)

        btn_row.addWidget(start)
        btn_row.addWidget(stop)
        btn_row.addWidget(kill)
        layout.addLayout(btn_row)

        # Status line
        status_row = QHBoxLayout()
        status_lbl = QLabel('Status: Idle')
        status_lbl.setStyleSheet('font-weight: 600;')
        detail_lbl = QLabel('')
        detail_lbl.setStyleSheet('opacity: 0.75;')
        detail_lbl.setWordWrap(True)
        status_row.addWidget(status_lbl)
        status_row.addStretch(1)
        status_row.addWidget(detail_lbl, 1)
        layout.addLayout(status_row)

        log = QTextEdit()
        log.setReadOnly(True)
        layout.addWidget(log, 1)

        start.clicked.connect(lambda: self.start_miner(miner_id))
        stop.clicked.connect(lambda: self.stop_miner(miner_id))
        kill.clicked.connect(lambda: self.kill_miner(miner_id))
        if pause is not None:
            pause.clicked.connect(lambda: self.toggle_pause_miner(miner_id))

        widgets = {"tab": w, "start": start, "stop": stop, "kill": kill, "log": log, "status": status_lbl, "detail": detail_lbl}
        if pause is not None:
            widgets["pause"] = pause
        return w, widgets

    def open_settings(self) -> None:
        dlg = SettingsDialog(self, self.settings)
        if dlg.exec() != QDialog.Accepted:
            return

        self.settings = dlg.collect_settings()
        save_settings(self.settings)

        app = QApplication.instance()
        if app is not None:
            apply_theme(app, self.settings)
            apply_border_color(app, self.settings)

        self.setWindowTitle(effective_app_name(self.settings))
        if any(p is not None and p.state() != QProcess.NotRunning for p in self.procs.values()):
            # Do not rebuild tabs while miners are running (keeps logs/controls visible)
            self._log_global("Settings applied (miners running; tab rebuild deferred).")
        else:
            self._rebuild_miner_tabs()
            self._update_all_buttons()
        # Refresh scripts panel after settings changes
        try:
            self._scripts_refresh_miner_combo()
            self._scripts_reload()
        except Exception:
            pass

        self._apply_coin_settings()


    def _any_miner_running(self) -> bool:
        for mid, proc in self.procs.items():
            if proc is not None and proc.state() != QProcess.NotRunning:
                return True
        return False

    def _build_promo_panel(self) -> QWidget:
        """Build the right-side promo/animation panel."""
        panel = QFrame()
        panel.setFrameShape(QFrame.StyledPanel)
        self._promo_full_width = 300
        panel.setFixedWidth(self._promo_full_width)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        header = QHBoxLayout()
        title = QLabel("Mad About Mining GUI")
        title.setStyleSheet("font-weight: 600;")
        header.addWidget(title, 1)

        self.promo_toggle_btn = QToolButton()
        self.promo_toggle_btn.setText("⤢")  # shrink/expand
        self.promo_toggle_btn.setToolTip("Shrink/expand this panel")
        self.promo_toggle_btn.clicked.connect(self._promo_toggle_shrink)
        header.addWidget(self.promo_toggle_btn, 0)

        layout.addLayout(header)

        support = QLabel("If you would like to support me\nLTC: LQMPsC7w75zaYMNJUmHboL6XfwNFAoPkqW")
        support.setWordWrap(True)
        support.setStyleSheet("opacity: 0.85;")
        layout.addWidget(support)

        try:
            coin_pixmaps = load_coin_pixmaps()
        except Exception:
            coin_pixmaps = []

        self.coin_anim = CoinRainWidget(coin_pixmaps, panel)
        self.coin_anim.setMinimumHeight(420)
        layout.addWidget(self.coin_anim, 1)

        marquee = MarqueeLabel("Mad About Mining GUI", panel)
        marquee.setStyleSheet("opacity: 0.85; font-weight: 600;")
        marquee.setDirectionLeftToRight(True)
        marquee.setSpeed(2)
        layout.addWidget(marquee)

        return panel




    # ---------------- Promo Panel Shrink/Expand ----------------
    def _promo_is_shrunk(self) -> bool:
        try:
            return bool(self.settings.get("ui", {}).get("promo_is_shrunk", False))
        except Exception:
            return False

    def _promo_shrink_pct(self) -> int:
        try:
            v = int(self.settings.get("ui", {}).get("promo_shrink_pct", 50))
            return max(20, min(100, v))
        except Exception:
            return 50

    def _promo_apply_width(self) -> None:
        try:
            panel = getattr(self, "promo_panel", None)
            if panel is None:
                return
            full_w = int(getattr(self, "_promo_full_width", 300))
            if self._promo_is_shrunk():
                pct = self._promo_shrink_pct()
                w = max(120, int(full_w * (pct / 100.0)))
            else:
                w = full_w
            panel.setFixedWidth(w)

            btn = getattr(self, "promo_toggle_btn", None)
            if btn is not None:
                btn.setText("⤢" if not self._promo_is_shrunk() else "⤡")
        except Exception:
            pass

    def _promo_set_shrunk(self, shrunk: bool) -> None:
        try:
            self.settings.setdefault("ui", {})
            self.settings["ui"]["promo_is_shrunk"] = bool(shrunk)
            save_settings(self.settings)
        except Exception:
            pass
        self._promo_apply_width()

    def _promo_toggle_shrink(self) -> None:
        self._promo_set_shrunk(not self._promo_is_shrunk())

    # ---------------- Scripts Panel (BAT selection) ----------------
    def _build_scripts_panel(self) -> QWidget:
        """Build the middle panel: per-miner BAT folder + checkbox activation.

        This panel is collapsible to reclaim horizontal space.
        """
        panel = QFrame()
        panel.setFrameShape(QFrame.StyledPanel)

        expanded_w = 340
        collapsed_w = 26  # thin handle with arrow

        # store for toggle logic
        self._scripts_expanded_w = expanded_w
        self._scripts_collapsed_w = collapsed_w
        self._scripts_is_collapsed = False
        self._scripts_panel_ref = panel

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        # ---- Header (always visible) ----
        header = QHBoxLayout()
        header.setContentsMargins(0, 0, 0, 0)

        self.scripts_toggle_btn = QToolButton()
        self.scripts_toggle_btn.setText("◀")
        self.scripts_toggle_btn.setToolTip("Hide Scripts panel")
        self.scripts_toggle_btn.setAutoRaise(True)
        self.scripts_toggle_btn.clicked.connect(self._scripts_toggle_collapsed)
        header.addWidget(self.scripts_toggle_btn)

        self.scripts_title_lbl = QLabel("Scripts")
        self.scripts_title_lbl.setStyleSheet("font-weight: 600;")
        header.addWidget(self.scripts_title_lbl)
        header.addStretch(1)

        layout.addLayout(header)

        # ---- Content (hidden when collapsed) ----
        self.scripts_content = QWidget(panel)
        content = QVBoxLayout(self.scripts_content)
        content.setContentsMargins(4, 2, 4, 4)
        content.setSpacing(8)

        self.scripts_miner_combo = QComboBox()
        self.scripts_miner_combo.currentIndexChanged.connect(self._scripts_on_miner_changed)
        content.addWidget(self.scripts_miner_combo)

        self.scripts_folder_label = QLabel("Folder: (not set)")
        self.scripts_folder_label.setWordWrap(True)
        self.scripts_folder_label.setStyleSheet("opacity: 0.85;")
        content.addWidget(self.scripts_folder_label)

        folder_btn_row = QHBoxLayout()
        self.scripts_browse_btn = QPushButton("Set Folder…")
        self.scripts_open_folder_btn = QPushButton("Open Folder")
        folder_btn_row.addWidget(self.scripts_browse_btn, 1)
        folder_btn_row.addWidget(self.scripts_open_folder_btn, 1)
        content.addLayout(folder_btn_row)

        self.scripts_browse_btn.clicked.connect(self._scripts_browse_folder)
        self.scripts_open_folder_btn.clicked.connect(self._scripts_open_folder)

        # Bulk selection controls
        bulk_row = QHBoxLayout()
        self.scripts_all_btn = QPushButton("All")
        self.scripts_none_btn = QPushButton("None")
        self.scripts_invert_btn = QPushButton("Invert")
        bulk_row.addWidget(self.scripts_all_btn)
        bulk_row.addWidget(self.scripts_none_btn)
        bulk_row.addWidget(self.scripts_invert_btn)
        bulk_row.addStretch(1)
        content.addLayout(bulk_row)

        self.scripts_all_btn.clicked.connect(lambda: self._scripts_bulk_check("all"))
        self.scripts_none_btn.clicked.connect(lambda: self._scripts_bulk_check("none"))
        self.scripts_invert_btn.clicked.connect(lambda: self._scripts_bulk_check("invert"))

        self.scripts_list = QListWidget()
        self.scripts_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.scripts_list.itemChanged.connect(self._scripts_on_item_changed)
        content.addWidget(self.scripts_list, 1)

        act_row = QHBoxLayout()
        self.scripts_reload_btn = QPushButton("Reload")
        self.scripts_open_btn = QPushButton("Open")
        self.scripts_edit_btn = QPushButton("Edit")
        act_row.addWidget(self.scripts_reload_btn)
        act_row.addWidget(self.scripts_open_btn)
        act_row.addWidget(self.scripts_edit_btn)
        content.addLayout(act_row)

        self.scripts_reload_btn.clicked.connect(self._scripts_reload)
        self.scripts_open_btn.clicked.connect(self._scripts_open_selected)
        self.scripts_edit_btn.clicked.connect(self._scripts_edit_selected)

        self.scripts_hint_lbl = QLabel(
            "Tip: tick one or more BAT files to mark them as active for that miner.\n"
            "When you press Start on a BAT miner, the app will use the active script."
        )
        self.scripts_hint_lbl.setWordWrap(True)
        self.scripts_hint_lbl.setStyleSheet("opacity: 0.70;")
        content.addWidget(self.scripts_hint_lbl)

        layout.addWidget(self.scripts_content, 1)

        self._scripts_refresh_miner_combo()
        # Sync with currently selected tab
        self.tabs.currentChanged.connect(self._scripts_sync_to_current_tab)
        QTimer.singleShot(0, self._scripts_sync_to_current_tab)

        # Auto-expand on interaction when collapsed
        try:
            panel.installEventFilter(self)
            self.scripts_content.installEventFilter(self)
            self.scripts_list.installEventFilter(self)
            self.scripts_miner_combo.installEventFilter(self)
        except Exception:
            pass

        # Apply initial collapsed state from settings
        collapsed = bool(self.settings.get("ui", {}).get("scripts_panel_collapsed", False))
        self._scripts_set_collapsed(collapsed, panel=panel, expanded_w=expanded_w, collapsed_w=collapsed_w)

        return panel


        
    def _scripts_set_collapsed(self, collapsed: bool, panel: QWidget | None = None,
                              expanded_w: int | None = None, collapsed_w: int | None = None) -> None:
        """Collapse/expand the Scripts panel."""
        try:
            ui = self.settings.get("ui", {}) or {}
            panel = panel or getattr(self, "_scripts_panel_ref", None) or getattr(self, "scripts_panel", None)
            if panel is None:
                return
            expanded_w = int(expanded_w or getattr(self, "_scripts_expanded_w", 340))
            collapsed_w = int(collapsed_w or getattr(self, "_scripts_collapsed_w", 26))

            self._scripts_is_collapsed = bool(collapsed)

            # Hide/show content
            try:
                self.scripts_content.setVisible(not collapsed)
            except Exception:
                pass

            # Adjust width
            try:
                panel.setFixedWidth(collapsed_w if collapsed else expanded_w)
            except Exception:
                pass

            # Toggle button arrow/tooltip
            try:
                if collapsed:
                    self.scripts_toggle_btn.setText("▶")
                    self.scripts_toggle_btn.setToolTip("Show Scripts panel")
                else:
                    self.scripts_toggle_btn.setText("◀")
                    self.scripts_toggle_btn.setToolTip("Hide Scripts panel")
            except Exception:
                pass

            # Persist state
            ui["scripts_panel_collapsed"] = bool(collapsed)
            self.settings["ui"] = ui
            save_settings(self.settings)
        except Exception:
            pass

    def _scripts_toggle_collapsed(self) -> None:
        """User clicked the arrow handle."""
        collapsed = not bool(getattr(self, "_scripts_is_collapsed", False))
        self._scripts_set_collapsed(collapsed)

    def eventFilter(self, obj, event):
        """Auto-expand scripts panel on click/focus when collapsed (optional)."""
        try:
            from PySide6.QtCore import QEvent
            if bool(getattr(self, "_scripts_is_collapsed", False)):
                ui = self.settings.get("ui", {}) or {}
                if bool(ui.get("auto_expand_scripts_panel", True)):
                    # If user interacts with scripts panel/children, expand
                    if obj in (getattr(self, "_scripts_panel_ref", None), getattr(self, "scripts_content", None),
                               getattr(self, "scripts_list", None), getattr(self, "scripts_miner_combo", None)):
                        if event.type() in (QEvent.MouseButtonPress, QEvent.FocusIn):
                            self._scripts_set_collapsed(False)
        except Exception:
            pass
        return super().eventFilter(obj, event)


    def _scripts_refresh_miner_combo(self) -> None:
        self.scripts_miner_combo.blockSignals(True)
        self.scripts_miner_combo.clear()
        for m in (self._miners_list() or []):
            if not bool(m.get("enabled", True)):
                continue
            mid = str(m.get("id", "")).strip()
            name = str(m.get("name", mid)).strip() or mid
            if mid:
                self.scripts_miner_combo.addItem(name, mid)
        self.scripts_miner_combo.blockSignals(False)

    def _scripts_sync_to_current_tab(self) -> None:
        try:
            idx = int(self.tabs.currentIndex())
        except Exception:
            return
        mid = None
        try:
            mid = self.tabs.widget(idx).property("miner_id")
        except Exception:
            mid = None
        if not mid:
            # Global logs or unknown
            return
        # set combo to this miner
        for i in range(self.scripts_miner_combo.count()):
            if self.scripts_miner_combo.itemData(i) == mid:
                self.scripts_miner_combo.setCurrentIndex(i)
                return

    def _scripts_current_miner_id(self) -> str | None:
        try:
            return str(self.scripts_miner_combo.currentData()).strip() or None
        except Exception:
            return None

    def _scripts_on_miner_changed(self) -> None:
        self._scripts_reload()

    def _scripts_get_dir(self, miner: dict) -> str:
        # Prefer scripts_dir for BAT miners; fall back to path for older configs
        return (miner.get("scripts_dir") or miner.get("path") or "").strip()

    def _scripts_list_bats(self, folder: str) -> list[Path]:
        try:
            d = Path(folder)
            if not d.exists() or not d.is_dir():
                return []
            return sorted(d.glob("*.bat"))
        except Exception:
            return []

    
    def _scripts_select_all(self) -> None:
        self.scripts_list.blockSignals(True)
        for i in range(self.scripts_list.count()):
            it = self.scripts_list.item(i)
            it.setCheckState(Qt.Checked)
        self.scripts_list.blockSignals(False)
        # Persist
        cur = self.scripts_list.currentItem()
        if cur is not None:
            self._scripts_on_item_changed(cur)

    def _scripts_select_none(self) -> None:
        # Keep at least one checked if there are any scripts (prevents "start does nothing")
        if self.scripts_list.count() <= 0:
            return
        self.scripts_list.blockSignals(True)
        for i in range(self.scripts_list.count()):
            it = self.scripts_list.item(i)
            it.setCheckState(Qt.Unchecked)
        # Re-check first item for safety
        self.scripts_list.item(0).setCheckState(Qt.Checked)
        self.scripts_list.blockSignals(False)
        self._scripts_on_item_changed(self.scripts_list.item(0))

    def _scripts_invert_selection(self) -> None:
        if self.scripts_list.count() <= 0:
            return
        self.scripts_list.blockSignals(True)
        for i in range(self.scripts_list.count()):
            it = self.scripts_list.item(i)
            it.setCheckState(Qt.Unchecked if it.checkState() == Qt.Checked else Qt.Checked)
        # Ensure at least one checked
        any_checked = any(self.scripts_list.item(i).checkState() == Qt.Checked for i in range(self.scripts_list.count()))
        if not any_checked and self.scripts_list.count() > 0:
            self.scripts_list.item(0).setCheckState(Qt.Checked)
        self.scripts_list.blockSignals(False)
        self._scripts_on_item_changed(self.scripts_list.item(0))
    def _scripts_reload(self) -> None:
        mid = self._scripts_current_miner_id()
        miner = self._find_miner(mid) if mid else None

        self.scripts_list.blockSignals(True)
        self.scripts_list.clear()

        if not miner:
            self.scripts_folder_label.setText("Folder: (no miner selected)")
            self.scripts_browse_btn.setEnabled(False)
            self.scripts_open_folder_btn.setEnabled(False)
            self.scripts_list.blockSignals(False)
            return

        mtype = (miner.get("type") or "EXE").upper().strip()
        folder = self._scripts_get_dir(miner)

        is_bat = (mtype == "BAT")
        self.scripts_browse_btn.setEnabled(is_bat)
        self.scripts_open_folder_btn.setEnabled(bool(folder))

        if not is_bat:
            self.scripts_folder_label.setText("Folder: (this miner is EXE-based)")
            self.scripts_list.blockSignals(False)
            return

        self.scripts_folder_label.setText(f"Folder: {folder or '(not set)'}")

        active = set([str(x) for x in (miner.get('active_scripts') or [])])

        for p in self._scripts_list_bats(folder):
            item = QListWidgetItem(p.name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            item.setCheckState(Qt.Checked if p.name in active else Qt.Unchecked)
            item.setData(Qt.UserRole, str(p))
            self.scripts_list.addItem(item)

        # Empty state
        if folder and self.scripts_list.count() == 0:
            it = QListWidgetItem("(no .bat files found in this folder)")
            it.setFlags(Qt.ItemIsEnabled)
            self.scripts_list.addItem(it)

        self.scripts_list.blockSignals(False)

        # If folder has BATs but none are active, auto-activate first (nice UX)
        if folder and self.scripts_list.count() > 0 and not active:
            first = self.scripts_list.item(0)
            first.setCheckState(Qt.Checked)

    def _scripts_on_item_changed(self, item: QListWidgetItem) -> None:
        mid = self._scripts_current_miner_id()
        miner = self._find_miner(mid) if mid else None
        if not miner:
            return

        # Persist checked items as active_scripts (filenames)
        active: list[str] = []
        for i in range(self.scripts_list.count()):
            it = self.scripts_list.item(i)
            # ignore placeholder rows
            if not it.flags() & Qt.ItemIsUserCheckable:
                continue
            if it.checkState() == Qt.Checked:
                active.append(it.text().strip())

        # Safety: keep at least one active when scripts exist
        if not active:
            for i in range(self.scripts_list.count()):
                it = self.scripts_list.item(i)
                if it.flags() & Qt.ItemIsUserCheckable:
                    it.setCheckState(Qt.Checked)
                    active = [it.text().strip()]
                    break
        miner["active_scripts"] = active
        miner["scripts_dir"] = self._scripts_get_dir(miner)
        # Keep path in sync for BAT miners (back-compat)
        if (miner.get("type") or "").upper().strip() == "BAT":
            miner["path"] = miner.get("scripts_dir") or miner.get("path") or ""

        save_settings(self.settings)

    def _scripts_browse_folder(self) -> None:
        mid = self._scripts_current_miner_id()
        miner = self._find_miner(mid) if mid else None
        if not miner:
            return

        start_dir = self._scripts_get_dir(miner)
        if not start_dir:
            # sensible default
            cand = APP_DIR / "miners" / (miner.get("id") or "")
            start_dir = str(cand.resolve()) if cand.exists() else str((APP_DIR / "miners").resolve())

        chosen = QFileDialog.getExistingDirectory(self, "Select scripts folder", start_dir)
        if not chosen:
            return

        # Optional: prompt to add Windows Defender exclusion for this scripts folder
        try:
            self._maybe_prompt_defender_exclusion_folder(chosen)
        except Exception:
            pass

        miner["scripts_dir"] = chosen
        miner["path"] = chosen  # keep UI/back-compat consistent for BAT miners
        if not miner.get("workdir"):
            miner["workdir"] = chosen

        save_settings(self.settings)
        self._scripts_reload()

    def _maybe_prompt_defender_exclusion_folder(self, folder: str) -> None:
        if sys.platform != "win32":
            return
        sec = self.settings.get("security", {}) or {}
        if not bool(sec.get("prompt_defender_exclusions", True)):
            return

        target = _norm_folder(folder)
        if not target:
            return

        excluded = set([_norm_folder(x) for x in (sec.get("defender_exclusions", []) or [])])
        if target in excluded:
            return

        msg = (
            "Windows Defender sometimes quarantines miners downloaded from the internet.\n\n"
            f"Add a Defender exclusion for this folder?\n\n{target}\n\n"
            "This requires Administrator approval (UAC prompt)."
        )
        ans = QMessageBox.question(self, "Add Defender exclusion?", msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ans != QMessageBox.Yes:
            return

        ok, detail = add_defender_exclusion_admin(target)
        if ok:
            sec.setdefault("defender_exclusions", [])
            sec["defender_exclusions"].append(target)
            self.settings["security"] = sec
            save_settings(self.settings)
            self._log_global(f"Defender exclusion added: {target}")
        else:
            QMessageBox.warning(self, "Defender exclusion", f"Failed to add exclusion.\n\n{detail}")

    def _scripts_open_folder(self) -> None:
        mid = self._scripts_current_miner_id()
        miner = self._find_miner(mid) if mid else None
        if not miner:
            return
        folder = self._scripts_get_dir(miner)
        if not folder:
            return
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(folder).resolve())))
        except Exception:
            pass

    def _scripts_selected_script_path(self) -> Path | None:
        try:
            item = self.scripts_list.currentItem()
            if item is None:
                return None
            raw = str(item.data(Qt.UserRole) or "").strip()
            if raw:
                p = Path(raw)
                return p if p.exists() else None
            # fallback: folder + filename
            mid = self._scripts_current_miner_id()
            miner = self._find_miner(mid) if mid else None
            folder = self._scripts_get_dir(miner or {})
            p = Path(folder) / item.text().strip()
            return p if p.exists() else None
        except Exception:
            return None

    def _scripts_open_selected(self) -> None:
        p = self._scripts_selected_script_path()
        if p is None:
            QMessageBox.warning(self, "Open script", "Select a BAT file first.")
            return
        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(p)))
        except Exception:
            QMessageBox.warning(self, "Open script", "Failed to open the selected file.")

    def _scripts_edit_selected(self) -> None:
        p = self._scripts_selected_script_path()
        if p is None:
            QMessageBox.warning(self, "Edit script", "Select a BAT file first.")
            return
        if p.suffix.lower() not in (".bat", ".cmd", ".ps1", ".sh", ".json", ".txt"):
            QMessageBox.warning(self, "Edit script", "Selected file type is not supported for editing.")
            return
        try:
            subprocess.Popen(["notepad.exe", str(p)])
        except Exception:
            QMessageBox.warning(self, "Edit script", "Failed to open Notepad.")
    def _apply_coin_settings(self) -> None:
        ui = self.settings.get("ui", {}) or {}
        try:
            self.coin_anim.set_size_pct(int(ui.get("coin_scale_pct", 50)))
            self.coin_anim.set_speed_pct(int(ui.get("coin_speed_pct", 100)))
            low = bool(ui.get("low_distraction_while_mining", True))
            self.coin_anim.set_low_distraction(low and self._any_miner_running())
        except Exception:
            pass

    def _tick_clock(self):
        self.time_label.setText(f"Time: {now_local_12h()}")

    def _log_global(self, msg: str):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.global_log.append(f"{ts} {msg}")
        log_to_file(msg)

    def _find_miner(self, miner_id: str) -> dict | None:
        for m in self._miners_list():
            if str(m.get("id", "")).strip() == miner_id:
                return m
        return None

    def _xmrig_generated_config_path(self, xmrig_exe_path: str) -> Path:
        return Path(xmrig_exe_path).parent / "config.generated.json"

    def _write_xmrig_config(self, xmrig_exe_path: str) -> Path | None:
        xm = self.settings.get("xmrig", {}) or {}
        if not bool(xm.get("auto_config", True)):
            return None

        pool = (xm.get("pool") or "").strip()
        wallet = (xm.get("wallet") or "").strip()
        worker = (xm.get("worker") or "").strip()
        password = (xm.get("pass") or "x").strip() or "x"

        if not pool or not wallet:
            QMessageBox.warning(
                self,
                "XMRig settings missing",
                "XMRig auto-config is enabled.\n"
                "Please set Pool and Wallet/User in Settings → XMRig, or disable auto-config."
            )
            return None

        user = wallet if not worker else f"{wallet}.{worker}"
        cfg = {
            "autosave": True,
            "cpu": True,
            "opencl": False,
            "cuda": False,
            "pools": [{"url": pool, "user": user, "pass": password, "keepalive": True, "tls": False}]
        }

        out_path = self._xmrig_generated_config_path(xmrig_exe_path)
        out_path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        return out_path

    def is_running(self, miner_id: str) -> bool:
        proc = self.procs.get(miner_id)
        return proc is not None and proc.state() != QProcess.NotRunning

    def is_paused(self, miner_id: str) -> bool:
        return bool(self.proc_paused.get(miner_id, False))

    def pause_miner(self, miner_id: str, silent: bool = False) -> None:
        """Pause (suspend) a running miner process.

        Implementation: on Windows only, suspend the process via NtSuspendProcess.
        """
        if sys.platform != "win32":
            if not silent:
                QMessageBox.information(self, "Pause not supported", "Pause/Resume is only supported on Windows.")
            return
        if not self.is_running(miner_id):
            if not silent:
                self._log_global(f"{miner_id}: Pause requested but not running.")
            return
        if self.is_paused(miner_id):
            return

        pid = self.proc_pids.get(miner_id)
        if not pid:
            if not silent:
                QMessageBox.warning(self, "Pause failed", "Could not determine the miner process PID.")
            return

        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)

        try:
            suspend_process(int(pid))
            self.proc_paused[miner_id] = True
            if not silent:
                self._log_global(f"{name}: Paused (process suspended).")
        except Exception as e:
            if not silent:
                QMessageBox.critical(self, "Pause failed", f"Failed to pause {name}.\n\n{e}")

        self._update_all_buttons()
        try:
            self._apply_coin_settings()
        except Exception:
            pass

    def resume_miner(self, miner_id: str, silent: bool = False) -> None:
        if sys.platform != "win32":
            if not silent:
                QMessageBox.information(self, "Resume not supported", "Pause/Resume is only supported on Windows.")
            return
        if not self.is_running(miner_id):
            self.proc_paused[miner_id] = False
            self._update_all_buttons()
            return
        if not self.is_paused(miner_id):
            return

        pid = self.proc_pids.get(miner_id)
        if not pid:
            self.proc_paused[miner_id] = False
            self._update_all_buttons()
            return

        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)

        try:
            resume_process(int(pid))
            self.proc_paused[miner_id] = False
            if not silent:
                self._log_global(f"{name}: Resumed.")
        except Exception as e:
            if not silent:
                QMessageBox.critical(self, "Resume failed", f"Failed to resume {name}.\n\n{e}")

        self._update_all_buttons()
        try:
            self._apply_coin_settings()
        except Exception:
            pass

    def toggle_pause_miner(self, miner_id: str) -> None:
        if self.is_paused(miner_id):
            self.resume_miner(miner_id)
        else:
            self.pause_miner(miner_id)



    def _bat_folder_and_active(self, miner: dict) -> tuple[str, list[str]]:
        folder = (miner.get("scripts_dir") or miner.get("path") or "").strip()
        active = [str(x).strip() for x in (miner.get("active_scripts") or []) if str(x).strip()]
        return folder, active

    def _choose_bat_script_path(self, miner_id: str) -> str | None:
        """Resolve which BAT file to run for a BAT miner.

        Rules:
          - If exactly 1 active script: use it.
          - If >1 active scripts: prompt user to pick one.
          - If none active:
              - If exactly 1 BAT exists in folder: use it (and mark active).
              - Else prompt user to pick from all BATs in folder.
        """
        miner = self._find_miner(miner_id)
        if not miner:
            return None
        folder, active = self._bat_folder_and_active(miner)
        if not folder:
            return None

        bats = self._scripts_list_bats(folder)
        if not bats:
            return None

        # Validate active against folder
        available = {p.name: str(p) for p in bats}
        active = [a for a in active if a in available]

        def _persist_active(names: list[str]) -> None:
            miner["active_scripts"] = names
            miner["scripts_dir"] = folder
            miner["path"] = folder  # keep back-compat for BAT miners
            save_settings(self.settings)
            try:
                self._scripts_reload()
            except Exception:
                pass

        if len(active) == 1:
            return available[active[0]]

        if len(active) > 1:
            choice, ok = QInputDialog.getItem(
                self,
                "Select script",
                f"Multiple active scripts are checked for {miner.get('name', miner_id)}.\nChoose which one to run:",
                active,
                0,
                False
            )
            if ok and choice in available:
                return available[choice]
            return None

        # No active scripts
        if len(bats) == 1:
            _persist_active([bats[0].name])
            return str(bats[0])

        all_names = [p.name for p in bats]
        choice, ok = QInputDialog.getItem(
            self,
            "Select script",
            f"No active script is selected for {miner.get('name', miner_id)}.\nChoose a script to run:",
            all_names,
            0,
            False
        )
        if ok and choice in available:
            _persist_active([choice])
            return available[choice]
        return None

    def start_miner(self, miner_id: str):
        miner = self._find_miner(miner_id)
        if not miner:
            QMessageBox.critical(self, "Missing miner", f"Miner not found: {miner_id}")
            return

        if self.is_running(miner_id):
            QMessageBox.warning(self, "Running", f"{miner.get('name', miner_id)} is already running.")
            return

        if bool(self.settings.get("ui", {}).get("single_active", False)):
            for mid in list(self.miner_tabs.keys()):
                if mid != miner_id and self.is_running(mid):
                    self.stop_miner(mid)

        mtype = (miner.get("type") or "EXE").upper().strip()
        path = (miner.get("path") or "").strip()
        args_str = (miner.get("args") or "").strip()
        workdir = (miner.get("workdir") or "").strip()

        if not path:
            QMessageBox.warning(self, "Missing path", f"Set Path for {miner.get('name', miner_id)} in Settings.")
            return

        exe = ""
        args: list[str] = []
        forced_workdir: str | None = None
        xmrig_cfg_path: Path | None = None

        if mtype == "BAT":
            bat_path = self._choose_bat_script_path(miner_id)
            if not bat_path:
                QMessageBox.warning(
                    self,
                    "No script selected",
                    f"Select a scripts folder and tick at least one .bat for {miner.get('name', miner_id)}\n"
                    f"using the Scripts panel."
                )
                return
            if not os.path.exists(bat_path):
                QMessageBox.critical(self, "BAT not found", f"BAT file not found:\n{bat_path}")
                return

            comspec = os.environ.get("ComSpec")
            if comspec and os.path.exists(comspec):
                exe = comspec
            else:
                exe = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "cmd.exe")

            args = ["/c", bat_path]
            forced_workdir = workdir or str(Path(bat_path).parent.resolve())
        else:
            exe_path = Path(path)
            if not exe_path.is_absolute():
                exe_path = (APP_DIR / exe_path).resolve()
            exe = str(exe_path)

            if not os.path.exists(exe):
                QMessageBox.critical(self, "EXE not found", f"Executable not found:\n{exe}")
                return

            if miner_id.lower() == "xmrig":
                xmrig_cfg_path = self._write_xmrig_config(exe)
                extra = str(self.settings.get("xmrig", {}).get("extra") or "").strip()
                if extra:
                    args_str = (args_str + " " + extra).strip()

            expanded_args_str = expand_placeholders(args_str, xmrig_cfg_path)
            args = split_args(expanded_args_str)
            forced_workdir = workdir or str(Path(exe).parent.resolve())

        if forced_workdir:
            forced_workdir = os.path.expandvars(forced_workdir)

        self._start_process(miner_id, exe, args, forced_workdir)
        try:
            self._apply_coin_settings()
        except Exception:
            pass


    def _start_process(self, miner_id: str, exe: str, args: list[str], workdir: str | None):
        widgets = self.miner_tabs.get(miner_id)
        if not widgets:
            return
        logw: QTextEdit = widgets["log"]

        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)

        wdir = workdir or str(Path(exe).parent.resolve())

        self._log_global(f"Starting {name}: {exe}")
        self._log_global(f"{name} args: {' '.join(args) if args else '(none)'}")
        self._log_global(f"{name} workdir: {wdir}")

        logw.append(f"--- Starting {name} ---")
        logw.append(f"Exe: {exe}")
        logw.append(f"Args: {' '.join(args) if args else '(none)'}")
        logw.append(f"Workdir: {wdir}")

        # Clear last error on (re)start
        self.proc_last_error[miner_id] = None

        proc = QProcess(self)
        proc.setProcessChannelMode(QProcess.MergedChannels)
        proc.setWorkingDirectory(wdir)

        # Windows: prevent console apps (e.g., miners) from spawning an external console window
        if os.name == 'nt':
            try:
                CREATE_NO_WINDOW = 0x08000000

                def _cpam(args):
                    # QProcess.CreateProcessArguments (Qt6): allow setting Win32 creation flags
                    try:
                        args.flags |= CREATE_NO_WINDOW
                    except Exception:
                        pass

                if hasattr(proc, 'setCreateProcessArgumentsModifier'):
                    proc.setCreateProcessArgumentsModifier(_cpam)
            except Exception:
                pass

        proc.readyReadStandardOutput.connect(lambda mid=miner_id, p=proc, lw=logw: self._on_ready_read(mid, p, lw))
        proc.errorOccurred.connect(lambda err, mid=miner_id, p=proc: self._on_proc_error(mid, err, p))
        proc.finished.connect(lambda code, status, mid=miner_id: self._on_proc_finished(mid, code))

        proc.start(exe, args)

        if not proc.waitForStarted(5000):
            err = proc.errorString() if hasattr(proc, "errorString") else "Unknown error"
            self._log_global(f"{name}: Failed to start. QProcess error: {err}")
            logw.append(f"Failed to start. QProcess error: {err}")
            QMessageBox.critical(self, "Start failed", f"{name} did not start.\n\n{err}\n\nSee log:\n{LOG_PATH}")
            return

        self.procs[miner_id] = proc
        try:
            self.proc_pids[miner_id] = int(proc.processId())
        except Exception:
            self.proc_pids[miner_id] = None

        self.proc_paused[miner_id] = False
        self._update_all_buttons()

        # Auto-collapse Scripts panel when mining starts (optional)
        try:
            ui = self.settings.get("ui", {}) or {}
            if bool(ui.get("auto_collapse_scripts_panel", True)):
                self._scripts_set_collapsed(True)
        except Exception:
            pass


    # ---------------- Tray Icon ----------------
    def _tray_enabled(self) -> bool:
        try:
            return bool(self.settings.get("ui", {}).get("tray_enabled", True))
        except Exception:
            return True

    def _tray_show_hashrate(self) -> bool:
        try:
            return bool(self.settings.get("ui", {}).get("tray_show_hashrate", True))
        except Exception:
            return True

    def _init_tray(self) -> None:
        if not self._tray_enabled():
            return

        # Ensure tray timer exists
        if not hasattr(self, "_tray_timer") or self._tray_timer is None:
            self._tray_timer = QTimer(self)
            self._tray_timer.timeout.connect(self._tray_update_tooltip)

        try:
            icon_path = get_icon_path()
            icon = QIcon(icon_path) if icon_path else QIcon(resource_path("assets/app.ico"))
        except Exception:
            icon = QIcon()

        self.tray = QSystemTrayIcon(icon, self)
        self._tray_menu = QMenu()
        self._tray_menu.aboutToShow.connect(self._tray_rebuild_menu)
        self.tray.setContextMenu(self._tray_menu)
        self.tray.setToolTip("Miner GUI")
        self.tray.show()

        self._tray_timer.start(1000)
        self._tray_update_tooltip()

        def _on_activated(reason):
            if reason == QSystemTrayIcon.Trigger:
                if self.isVisible():
                    self.hide()
                else:
                    self.showNormal()
                    self.raise_()
                    self.activateWindow()

        self.tray.activated.connect(_on_activated)

    def _tray_rebuild_menu(self) -> None:
        if not self._tray_menu:
            return
        m = self._tray_menu
        m.clear()

        if self._tray_show_hashrate():
            xm = getattr(self, 'last_hashrate', {}).get('xmrig')
            a = QAction(f"XMRig: {xm}" if xm else "XMRig: (no hashrate yet)", self)
            a.setEnabled(False)
            m.addAction(a)
            m.addSeparator()

        running_miners = []
        for miner in (self._miners_list() or []):
            mid = str(miner.get("id", "")).strip()
            if mid and self.is_running(mid):
                running_miners.append(miner)

        if not running_miners:
            a = QAction("No miners running", self)
            a.setEnabled(False)
            m.addAction(a)
        else:
            for miner in running_miners:
                mid = str(miner.get("id", "")).strip()
                name = str(miner.get("name", mid)).strip() or mid
                sub = QMenu(name, m)

                act_stop = QAction("Stop", self)
                act_stop.triggered.connect(lambda _=False, x=mid: self.stop_miner(x))
                sub.addAction(act_stop)

                if mid.lower() == "xmrig":
                    paused = self.is_paused(mid)
                    act_pause = QAction("Resume" if paused else "Pause", self)
                    act_pause.triggered.connect(lambda _=False, x=mid: self.toggle_pause_miner(x))
                    sub.addAction(act_pause)

                hr = getattr(self, 'last_hashrate', {}).get(mid.lower())
                if hr:
                    sub.addSeparator()
                    act_hr = QAction(f"Hashrate: {hr}", self)
                    act_hr.setEnabled(False)
                    sub.addAction(act_hr)

                m.addMenu(sub)

        m.addSeparator()

        act_settings = QAction("Settings…", self)
        act_settings.triggered.connect(self.open_settings)
        m.addAction(act_settings)

        act_show = QAction("Show window", self)
        act_show.triggered.connect(lambda: (self.showNormal(), self.raise_(), self.activateWindow()))
        m.addAction(act_show)

        act_quit = QAction("Quit", self)
        act_quit.triggered.connect(QApplication.instance().quit)
        m.addAction(act_quit)

    def _tray_update_tooltip(self) -> None:
        if not getattr(self, "tray", None):
            return

        lines = []
        if self._tray_show_hashrate():
            xm = getattr(self, 'last_hashrate', {}).get('xmrig')
            if xm:
                lines.append(f"XMRig: {xm}")

        running_names = []
        for miner in (self._miners_list() or []):
            mid = str(miner.get("id", "")).strip()
            if mid and self.is_running(mid):
                running_names.append(str(miner.get("name", mid)).strip() or mid)
        lines.append("Running: " + ", ".join(running_names) if running_names else "No miners running")

        self.tray.setToolTip("\n".join(lines))

    # ---------------- Hashrate capture ----------------
    _XMRIG_SPEED_RE = re.compile(r"speed\s+(?:10s/60s/15m\s+)?([0-9.]+)\s*([kMGT]?H/s)", re.IGNORECASE)
    _XMRIG_HASHRATE_RE = re.compile(r"hashrate\s*[:=]?\s*([0-9.]+)\s*([kMGT]?H/s)", re.IGNORECASE)

    def _maybe_capture_hashrate(self, miner_id: str, line: str) -> None:
        mid = (miner_id or "").lower()
        if mid != "xmrig":
            return
        s = strip_ansi(line)
        m = self._XMRIG_SPEED_RE.search(s) or self._XMRIG_HASHRATE_RE.search(s)
        if not m:
            return
        val, unit = m.group(1).strip(), m.group(2).strip()
        getattr(self, 'last_hashrate', {}).setdefault('xmrig','')
        self.last_hashrate["xmrig"] = f"{val} {unit}"



    def closeEvent(self, event) -> None:
        try:
            if self._tray_enabled() and bool(self.settings.get("ui", {}).get("tray_minimize_on_close", False)):
                event.ignore()
                self.hide()
                if getattr(self, "tray", None):
                    self.tray.showMessage("Miner GUI", "Still running in the tray.", QSystemTrayIcon.Information, 1500)
                return
        except Exception:
            pass
        super().closeEvent(event)

    def _on_ready_read(self, miner_id: str, proc: QProcess, log_widget: QTextEdit):
        data = bytes(proc.readAllStandardOutput())
        text = decode_process_bytes(data)
        if not text:
            return
        for line in text.splitlines():
            try:
                self._maybe_capture_hashrate(miner_id, line)
            except Exception:
                pass
            append_html(log_widget, line_to_html_with_rules(line, miner_id))
            self._log_global(f"[{miner_id}] {strip_ansi(line)}")

    def stop_miner(self, miner_id: str):
        proc = self.procs.get(miner_id)
        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)

        if not proc or proc.state() == QProcess.NotRunning:
            self._log_global(f"{name}: Stop requested but not running.")
            return

        if self.is_paused(miner_id):
            self.resume_miner(miner_id, silent=True)

        self._log_global(f"{name}: Stop requested (graceful terminate).")
        proc.terminate()

        t = QTimer(self)
        t.setSingleShot(True)
        t.timeout.connect(lambda mid=miner_id: self._escalate_kill_if_running(mid))
        t.start(5000)

        self._update_all_buttons()

    def _escalate_kill_if_running(self, miner_id: str):
        if self.is_running(miner_id):
            miner = self._find_miner(miner_id) or {}
            name = miner.get("name", miner_id)
            self._log_global(f"{name}: Did not stop gracefully; escalating to kill.")
            self.kill_miner(miner_id)

    def kill_miner(self, miner_id: str):
        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)

        if self.is_paused(miner_id):
            self.resume_miner(miner_id, silent=True)

        proc = self.procs.get(miner_id)
        if proc and proc.state() != QProcess.NotRunning:
            self._log_global(f"{name}: Force kill requested.")
            try:
                proc.kill()
                proc.waitForFinished(2000)
            except Exception:
                pass

        pid = self.proc_pids.get(miner_id)
        if pid:
            kill_process_tree(pid, include_parent=True)

        kills = miner.get("kill_names", []) or []
        killed_total = 0
        for kname in kills:
            killed_total += kill_by_name(str(kname).strip())

        if killed_total:
            self._log_global(f"{name}: kill-by-name fallback killed {killed_total} process(es).")

        self.procs[miner_id] = None
        self.proc_pids[miner_id] = None
        self.proc_paused[miner_id] = False
        self.proc_paused[miner_id] = False
        self._update_all_buttons()

    def _on_proc_error(self, miner_id: str, err, proc: QProcess):
        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)
        err_str = proc.errorString() if hasattr(proc, "errorString") else str(err)
        msg = f"{name}: process error {err} ({err_str})"
        self.proc_last_error[miner_id] = msg
        self._log_global(msg)
        self._update_all_buttons()

    def _on_proc_finished(self, miner_id: str, exit_code: int):
        miner = self._find_miner(miner_id) or {}
        name = miner.get("name", miner_id)
        self._log_global(f"{name}: exited (code {exit_code}).")
        if int(exit_code) != 0:
            self.proc_last_error[miner_id] = f"{name}: exited with code {exit_code}"
        else:
            # Successful exit clears prior errors
            self.proc_last_error[miner_id] = None

        self.procs[miner_id] = None
        self.proc_pids[miner_id] = None
        self.proc_paused[miner_id] = False
        self._update_all_buttons()


    def _update_all_buttons(self):
        for mid, w in self.miner_tabs.items():
            running = self.is_running(mid)
            w["start"].setEnabled(not running)
            w["stop"].setEnabled(running)
            w["kill"].setEnabled(running)

            if "pause" in w:
                paused = self.is_paused(mid)
                w["pause"].setEnabled(running)
                # Keep the label accurate
                base_name = (self._find_miner(mid) or {}).get("name", mid)
                w["pause"].setText(f"{'Resume' if paused else 'Pause'} {base_name}")

            # Status display
            if "status" in w:
                paused = self.is_paused(mid)
                err = self.proc_last_error.get(mid)
                if running and paused:
                    w["status"].setText("Status: Paused")
                elif running:
                    w["status"].setText("Status: Running")
                elif err:
                    w["status"].setText("Status: Error")
                else:
                    w["status"].setText("Status: Idle")

                detail = ""
                miner = self._find_miner(mid) or {}
                if err and not running:
                    detail = err
                else:
                    if (miner.get("type") or "").upper().strip() == "BAT":
                        active = miner.get("active_scripts") or []
                        detail = ("Active: " + ", ".join([str(x) for x in active])) if active else "Active: (none selected)"
                w["detail"].setText(detail)

    def start_miner_overrides(
        self,
        miner_id: str,
        args_override: str | None = None,
        workdir_override: str | None = None
    ) -> None:
        """
        Start a miner using the configured entry, but allow overriding args and/or workdir.

        This is intended for profit-switching (inject pool/wallet/worker per target).
        """
        miner = self._find_miner(miner_id)
        if not miner:
            QMessageBox.critical(self, "Missing miner", f"Miner not found: {miner_id}")
            return

        if self.is_running(miner_id):
            QMessageBox.warning(self, "Running", f"{miner.get('name', miner_id)} is already running.")
            return

        if bool(self.settings.get("ui", {}).get("single_active", False)):
            for mid in list(self.miner_tabs.keys()):
                if mid != miner_id and self.is_running(mid):
                    self.stop_miner(mid)

        mtype = (miner.get("type") or "EXE").upper().strip()
        path = (miner.get("path") or "").strip()

        args_str = (args_override if args_override is not None else (miner.get("args") or "")).strip()
        workdir = (workdir_override if workdir_override is not None else (miner.get("workdir") or "")).strip()

        if not path:
            QMessageBox.warning(self, "Missing path", f"Set Path for {miner.get('name', miner_id)} in Settings.")
            return

        exe = ""
        args: list[str] = []
        forced_workdir: str | None = None
        xmrig_cfg_path: Path | None = None

        if mtype == "BAT":
            bat_path = self._choose_bat_script_path(miner_id)
            if not bat_path:
                QMessageBox.warning(
                    self,
                    "No script selected",
                    f"Select a scripts folder and tick at least one .bat for {miner.get('name', miner_id)}\n"
                    f"using the Scripts panel."
                )
                return
            if not os.path.exists(bat_path):
                QMessageBox.critical(self, "BAT not found", f"BAT file not found:\n{bat_path}")
                return

            comspec = os.environ.get("ComSpec")
            if comspec and os.path.exists(comspec):
                exe = comspec
            else:
                exe = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "cmd.exe")

            args = ["/c", bat_path]
            forced_workdir = workdir or str(Path(bat_path).parent.resolve())
        else:
            exe_path = Path(path)
            if not exe_path.is_absolute():
                exe_path = (APP_DIR / exe_path).resolve()
            exe = str(exe_path)

            if not os.path.exists(exe):
                QMessageBox.critical(self, "EXE not found", f"Executable not found:\n{exe}")
                return

            if miner_id.lower() == "xmrig":
                xmrig_cfg_path = self._write_xmrig_config(exe)
                extra = str(self.settings.get("xmrig", {}).get("extra") or "").strip()
                if extra:
                    args_str = (args_str + " " + extra).strip()

            expanded_args_str = expand_placeholders(args_str, xmrig_cfg_path)
            args = split_args(expanded_args_str)
            forced_workdir = workdir or str(Path(exe).parent.resolve())

        if forced_workdir:
            forced_workdir = os.path.expandvars(forced_workdir)

        self._start_process(miner_id, exe, args, forced_workdir)

# ---------------- Entry ----------------
def main() -> int:
    try:
        settings = load_settings()

        # AppUserModelID must be set BEFORE QApplication is created (Windows)
        set_windows_appusermodel_id(settings.get("app", {}).get("app_id", "MinerGUI.Michael.Mining"))

        app = QApplication(sys.argv)
        app.setApplicationName(effective_app_name(settings))

        icon_path = Path(__file__).resolve().parent / "assets" / "app.ico"
        if icon_path.exists():
            app.setWindowIcon(QIcon(str(icon_path)))

        apply_theme(app, settings)
        apply_border_color(app, settings)

        win = MinerGUI(settings)
        apply_icons(app, win)

        # Startup sanity check: miners folder (helps for frozen builds)
        miners_root = APP_DIR / "miners"
        if not miners_root.exists():
            log_to_file(f"Warning: miners folder not found at {miners_root}")
            QMessageBox.warning(
                win,
                "Miners folder missing",
                f"Expected miners folder:\n{miners_root}\n\n"
                f"If you are running a built EXE, ensure miners/ is shipped alongside it."
            )

        win.resize(1100, 760)
        win.show()
        return app.exec()

    except Exception:
        tb = traceback.format_exc()
        log_to_file("FATAL EXCEPTION:\n" + tb)
        raise


if __name__ == "__main__":
    raise SystemExit(main())