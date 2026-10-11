"""
快捷方式清理工具 v1.0
扫描桌面 / 开始菜单 / 快速启动等位置的 .lnk 与 .url 快捷方式，
检测目标是否失效，并支持删除（移入回收站）。
依赖：PyQt5, pywin32（requests 可选，用于 .url 校验）
"""
import os
import re
import sys
import shutil
import logging
import winreg
import ctypes
from ctypes import wintypes
from logging.handlers import RotatingFileHandler

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QCheckBox, QComboBox,
    QLabel, QLineEdit, QProgressBar, QMessageBox, QAbstractItemView,
    QHeaderView, QFileDialog, QMenu, QAction, QFrame
)

try:
    import pythoncom
    import win32com.client
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

APP_TITLE = "快捷方式清理工具 v1.0"

# ---------------- 日志 ----------------
# 打包成 exe 后使用 exe 所在目录，脚本模式使用脚本所在目录
_BASE_DIR = (os.path.dirname(os.path.abspath(sys.executable))
             if getattr(sys, "frozen", False)
             else os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = os.path.join(_BASE_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "shortcut_cleaner.log")


def _setup_logger():
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        lg = logging.getLogger("shortcut_cleaner")
        lg.setLevel(logging.INFO)
        if not lg.handlers:
            h = RotatingFileHandler(LOG_FILE, maxBytes=1024 * 1024,
                                    backupCount=3, encoding="utf-8")
            h.setFormatter(logging.Formatter(
                "%(asctime)s [%(levelname)s] %(message)s"))
            lg.addHandler(h)
        return lg
    except OSError:
        return logging.getLogger("shortcut_cleaner")  # 无文件时仅内存


logger = _setup_logger()

# ---------------- 系统目录 ----------------
CSIDL_DESKTOPDIRECTORY = 0x10        # 当前用户桌面
CSIDL_COMMON_DESKTOPDIRECTORY = 0x19 # 公共桌面
CSIDL_PROGRAMS = 0x02                # 用户开始菜单
CSIDL_COMMON_PROGRAMS = 0x17         # 公共开始菜单
CSIDL_APPDATA = 0x1A


def get_shell_folder(csidl):
    buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    if ctypes.windll.shell32.SHGetFolderPathW(None, csidl, None, 0, buf) == 0:
        return buf.value
    return ""


def quick_launch_dir():
    appdata = get_shell_folder(CSIDL_APPDATA)
    return os.path.join(appdata, "Microsoft", "Internet Explorer", "Quick Launch") if appdata else ""


# ---------------- 状态定义 ----------------
STATUS_META = {
    "ok":        ("目标存在",       QColor("#2e7d32")),
    "broken":    ("目标丢失",       QColor("#c62828")),
    "nodrive":   ("驱动器不存在",   QColor("#c62828")),
    "network":   ("网络目标(未验证)", QColor("#e65100")),
    "notarget":  ("无目标路径",     QColor("#616161")),
    "url_ok":    ("可访问",         QColor("#2e7d32")),
    "url_bad":   ("无法访问",       QColor("#c62828")),
    "url_unknown": ("无法验证",     QColor("#e65100")),
    "url_protocol_registered": ("本地协议已注册", QColor("#2e7d32")),
    "url_protocol_missing": ("协议未注册",     QColor("#c62828")),
    "url_steam_not_installed": ("Steam 游戏未安装", QColor("#c62828")),
}
# 视为“失效、可清理”的状态
BROKEN_STATES = {"broken", "nodrive", "url_protocol_missing", "url_steam_not_installed"}
SUSPECT_STATES = BROKEN_STATES | {"url_bad"}
# 所有“异常”状态（含无法验证）
ABNORMAL_STATES = SUSPECT_STATES | {"url_unknown"}
# .url 文件的所有异常状态（由下拉框统一控制处理方式）
URL_ABNORMAL_STATES = {"url_unknown", "url_bad",
                       "url_protocol_missing", "url_steam_not_installed"}

UNC_RE = re.compile(r"^\\\\[^\\]+\\[^\\]+")
STEAM_RUN_RE = re.compile(r"^steam://rungameid/(\d+)", re.IGNORECASE)

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

# 关闭 requests 的 SSL 校验警告
if HAS_REQUESTS:
    from requests.packages.urllib3.exceptions import InsecureRequestWarning
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

    def _http_session():
        s = requests.Session()
        s.mount("http://", requests.adapters.HTTPAdapter(max_retries=2))
        s.mount("https://", requests.adapters.HTTPAdapter(max_retries=2))
        return s


def steam_app_installed(appid):
    """查询注册表判断 Steam 游戏是否已安装"""
    try:
        appid_int = int(appid)
    except ValueError:
        return None  # 未知
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Valve\Steam\Apps\%d" % appid_int) as k:
            val, _ = winreg.QueryValueEx(k, "Installed")
            return val == 1
    except FileNotFoundError:
        return False
    except OSError:
        return None


def protocol_registered(scheme):
    """检测 Windows 是否注册了 scheme 协议处理器（如 steam / magnet / mailto）"""
    scheme = scheme.lower()
    if scheme in ("http", "https"):
        return True
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, scheme) as k:
            try:
                with winreg.OpenKey(k, r"shell\open\command") as _:
                    return True
            except FileNotFoundError:
                # 协议键存在但没有 open\command，仍视为可能有效
                return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


class ShortcutItem(object):
    __slots__ = ("path", "name", "dir", "target", "status", "detail")

    def __init__(self, path, target, status, detail=""):
        self.path = path
        self.name = os.path.basename(path)
        self.dir = os.path.dirname(path)
        self.target = target
        self.status = status
        self.detail = detail


# ---------------- 扫描线程 ----------------
class ScanThread(QThread):
    sig_one = pyqtSignal(object)      # ShortcutItem
    sig_progress = pyqtSignal(str)    # 当前扫描目录
    sig_done = pyqtSignal(int, int)   # 总数, 失效数

    def __init__(self, roots, check_url=False):
        super().__init__()
        self.roots = roots
        self.check_url = check_url
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        if HAS_PYWIN32:
            pythoncom.CoInitialize()
        total = broken = 0
        try:
            for root in self.roots:
                if self._stop:
                    break
                self.sig_progress.emit(root)
                for dirpath, dirnames, filenames in os.walk(root):
                    if self._stop:
                        break
                    for fn in filenames:
                        if self._stop:
                            break
                        ext = os.path.splitext(fn)[1].lower()
                        full = os.path.join(dirpath, fn)
                        if ext == ".lnk":
                            item = self._check_lnk(full)
                        elif ext == ".url":
                            item = self._check_url_file(full)
                        else:
                            continue
                        if item:
                            total += 1
                            if item.status in SUSPECT_STATES:
                                broken += 1
                            self.sig_one.emit(item)
        except Exception:
            logger.exception("扫描线程异常")
        finally:
            if HAS_PYWIN32:
                pythoncom.CoUninitialize()
        self.sig_done.emit(total, broken)

    # ---- .lnk 检测 ----
    def _check_lnk(self, path):
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            sc = shell.CreateShortCut(path)
            target = (sc.TargetPath or "").strip()
            args = (sc.Arguments or "").strip()
        except Exception as e:
            logger.warning("解析 .lnk 失败 %s: %s", path, e)
            return ShortcutItem(path, "", "notarget", "解析失败: %s" % e)

        if not target:
            # 无目标：可能是命令型快捷方式或 UWP 应用入口
            hint = args if args else "可能是 UWP 应用或特殊目标"
            return ShortcutItem(path, hint, "notarget")

        expanded = os.path.expandvars(target)
        if os.path.exists(expanded):
            return ShortcutItem(path, expanded, "ok")

        # 网络路径（UNC）：默认不探测，避免长时间阻塞
        if UNC_RE.match(expanded) or expanded.startswith("\\\\"):
            return ShortcutItem(path, expanded, "network")

        # 检查盘符是否存在（常见于 U 盘 / 已卸载分区）
        drive, _ = os.path.splitdrive(expanded)
        if drive and not os.path.exists(drive + "\\"):
            return ShortcutItem(path, expanded, "nodrive")

        return ShortcutItem(path, expanded, "broken")

    # ---- .url 检测 ----
    def _check_url_file(self, path):
        url = ""
        try:
            with open(path, "r", encoding="utf-8-sig", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line.lower().startswith("url="):
                        url = line[4:].strip()
                        break
        except OSError:
            pass

        if not url:
            return ShortcutItem(path, "", "url_unknown", "未找到 URL 字段")

        scheme = url.split(":", 1)[0].lower() if ":" in url else ""
        if scheme not in ("http", "https"):
            if not protocol_registered(scheme):
                return ShortcutItem(
                    path, url, "url_protocol_missing",
                    "%s:// 协议未注册，打开会失败" % scheme)

            # Steam 游戏快捷方式：进一步校验游戏是否已安装
            m = STEAM_RUN_RE.match(url)
            if m:
                appid = m.group(1)
                installed = steam_app_installed(appid)
                if installed is True:
                    return ShortcutItem(
                        path, url, "url_protocol_registered",
                        "Steam AppID %s 已安装" % appid)
                elif installed is False:
                    return ShortcutItem(
                        path, url, "url_steam_not_installed",
                        "Steam AppID %s 未安装（需下载）" % appid)
                else:
                    return ShortcutItem(
                        path, url, "url_unknown",
                        "无法确认 Steam AppID %s 安装状态" % appid)

            return ShortcutItem(
                path, url, "url_protocol_registered",
                "%s:// 已注册，内容需到对应客户端内确认" % scheme)

        if not (self.check_url and HAS_REQUESTS):
            return ShortcutItem(path, url, "url_unknown",
                                "未启用在线校验" if not self.check_url else "未安装 requests")

        # 先用 HEAD 探测；若被 403/405 等拒绝，回退到 GET 重试
        s = _http_session()
        try:
            r = s.head(url, timeout=8, allow_redirects=True,
                       headers=HTTP_HEADERS, verify=False)
            code = r.status_code
            method = "HEAD"
        except requests.RequestException as e:
            # HEAD 直接失败（超时/SSL/连接），尝试 GET
            try:
                r = s.get(
                    url, timeout=10, allow_redirects=True,
                    headers=HTTP_HEADERS, stream=True, verify=False)
                with r:
                    code = r.status_code
                method = "GET"
            except requests.RequestException as e2:
                detail = "连接失败: %s" % (str(e2) if str(e2) else str(e))
                logger.info("URL 连接失败 %s: %s", url, detail)
                return ShortcutItem(path, url, "url_bad", detail)

        if code < 400:
            return ShortcutItem(path, url, "url_ok", "%s HTTP %d" % (method, code))

        if code in (403, 405):
            # 部分站点只拒绝 HEAD，改用 GET 再试一次
            try:
                r2 = s.get(
                    url, timeout=10, allow_redirects=True,
                    headers=HTTP_HEADERS, stream=True, verify=False)
                with r2:
                    code2 = r2.status_code
                if code2 < 400:
                    return ShortcutItem(path, url, "url_ok", "GET HTTP %d" % code2)
                if code2 in (401, 403, 405, 418, 429, 451):
                    return ShortcutItem(
                        path, url, "url_unknown",
                        "HTTP %d (站点启用防护)" % code2)
                return ShortcutItem(path, url, "url_bad", "GET HTTP %d" % code2)
            except requests.RequestException as e:
                logger.info("URL GET 失败 %s (HEAD %d): %s", url, code, e)
                return ShortcutItem(
                    path, url, "url_unknown",
                    "HEAD %d; GET 连接失败: %s" % (code, str(e)))

        if code in (401, 403, 405, 418, 429, 451):
            return ShortcutItem(
                path, url, "url_unknown",
                "%s HTTP %d (站点启用防护)" % (method, code))

        return ShortcutItem(path, url, "url_bad", "%s HTTP %d" % (method, code))


# ---------------- 回收站删除 ----------------
FO_DELETE = 3
FOF_SILENT = 0x4
FOF_NOCONFIRMATION = 0x10
FOF_ALLOWUNDO = 0x40
FOF_NOERRORUI = 0x400


class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("wFunc", ctypes.c_uint),
        ("pFrom", ctypes.c_wchar_p),
        ("pTo", ctypes.c_wchar_p),
        ("fFlags", ctypes.c_ushort),
        ("fAnyOperationsAborted", wintypes.BOOL),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", ctypes.c_wchar_p),
    ]


def delete_to_recycle_bin(paths):
    """把多个文件移入回收站，返回 (成功数, 失败数)"""
    if not paths:
        return 0, 0
    buf = "\x00".join(paths) + "\x00\x00"
    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = buf
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if res == 0 and not op.fAnyOperationsAborted:
        return len(paths), 0
    return 0, len(paths)


# ---------------- 主窗口 ----------------
COL_COUNT = 5  # 勾选 / 名称 / 所在位置 / 目标 / 状态


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1000, 620)
        self.items = []          # ShortcutItem 列表
        self.thread = None
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # --- 扫描范围 ---
        range_box = QFrame()
        range_box.setFrameShape(QFrame.StyledPanel)
        rl = QVBoxLayout(range_box)
        rl.setContentsMargins(10, 8, 10, 8)

        title = QLabel("扫描范围")
        title.setStyleSheet("font-weight:bold;")
        rl.addWidget(title)

        cb_layout = QHBoxLayout()
        self.cb_desktop = QCheckBox("用户桌面")
        self.cb_common_desktop = QCheckBox("公共桌面")
        self.cb_start = QCheckBox("开始菜单")
        self.cb_common_start = QCheckBox("公共开始菜单")
        self.cb_quick = QCheckBox("快速启动栏")
        # 扫描范围默认全部不勾选
        for cb in (self.cb_desktop, self.cb_common_desktop,
                   self.cb_start, self.cb_common_start):
            cb_layout.addWidget(cb)
        self.cb_quick.setChecked(False)
        cb_layout.addWidget(self.cb_quick)
        self.cb_subdirs = QCheckBox("包含子目录")
        self.cb_subdirs.setChecked(True)
        cb_layout.addWidget(self.cb_subdirs)
        self.cb_url_check = QCheckBox("在线校验 .url 链接")
        self.cb_url_check.setChecked(True)
        if not HAS_REQUESTS:
            self.cb_url_check.setEnabled(False)
            self.cb_url_check.setToolTip("未安装 requests 库")
        cb_layout.addWidget(self.cb_url_check)
        cb_layout.addStretch()
        rl.addLayout(cb_layout)

        custom_layout = QHBoxLayout()
        self.edit_custom = QLineEdit()
        self.edit_custom.setPlaceholderText("自定义扫描目录（可选）")
        btn_browse = QPushButton("浏览…")
        btn_browse.clicked.connect(self._pick_dir)
        custom_layout.addWidget(QLabel("自定义目录:"))
        custom_layout.addWidget(self.edit_custom, 1)
        custom_layout.addWidget(btn_browse)
        rl.addLayout(custom_layout)
        root.addWidget(range_box)

        # --- 操作栏 ---
        op_layout = QHBoxLayout()
        self.btn_scan = QPushButton("开始扫描")
        self.btn_scan.setStyleSheet(
            "QPushButton{background:#1976d2;color:white;border:none;"
            "padding:6px 18px;border-radius:4px;font-weight:bold;}"
            "QPushButton:hover{background:#1565c0;}"
            "QPushButton:disabled{background:#90a4ae;}")
        self.btn_scan.clicked.connect(self._start_scan)

        self.btn_stop = QPushButton("停止")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop_scan)

        self.cmb_filter = QComboBox()
        self.cmb_filter.addItems(["全部", "仅失效项", "仅有效项"])
        self.cmb_filter.currentIndexChanged.connect(self._refresh_table)

        self.btn_select_broken = QPushButton("全选异常项")
        self.btn_select_broken.clicked.connect(self._select_broken)

        self.btn_delete = QPushButton("删除所选（移入回收站）")
        self.btn_delete.setStyleSheet(
            "QPushButton{background:#c62828;color:white;border:none;"
            "padding:6px 18px;border-radius:4px;font-weight:bold;}"
            "QPushButton:hover{background:#b71c1c;}")
        self.btn_delete.clicked.connect(self._delete_selected)

        for w in (self.btn_scan, self.btn_stop, self.cmb_filter,
                  self.btn_select_broken, self.btn_delete):
            op_layout.addWidget(w)
        op_layout.addStretch()
        root.addLayout(op_layout)

        # --- 异常 .url 处理 ---
        unv_layout = QHBoxLayout()
        unv_layout.addWidget(QLabel("异常 .url 的处理:"))
        self.cmb_unverified = QComboBox()
        self.cmb_unverified.addItems(["保留", "按失效项一起删除", "移动到指定文件夹"])
        self.cmb_unverified.setCurrentIndex(0)
        self.cmb_unverified.currentIndexChanged.connect(self._on_unverified_changed)
        unv_layout.addWidget(self.cmb_unverified)

        self.edit_unverified_dir = QLineEdit("E:\\Software\\Shortcut\\下载类网站")
        self.edit_unverified_dir.setPlaceholderText("目标文件夹")
        self.edit_unverified_dir.setVisible(False)
        unv_layout.addWidget(self.edit_unverified_dir, 1)

        self.btn_browse_unverified = QPushButton("浏览…")
        self.btn_browse_unverified.setVisible(False)
        self.btn_browse_unverified.clicked.connect(self._pick_unverified_dir)
        unv_layout.addWidget(self.btn_browse_unverified)

        self.btn_move_unverified = QPushButton("移动所选异常 .url 项")
        self.btn_move_unverified.setVisible(False)
        self.btn_move_unverified.setStyleSheet(
            "QPushButton{background:#f57c00;color:white;border:none;"
            "padding:6px 18px;border-radius:4px;font-weight:bold;}"
            "QPushButton:hover{background:#ef6c00;}")
        self.btn_move_unverified.clicked.connect(self._move_unverified)
        unv_layout.addWidget(self.btn_move_unverified)
        unv_layout.addStretch()
        root.addLayout(unv_layout)

        # --- 进度 ---
        prog_layout = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(6)
        self.lbl_status = QLabel("就绪")
        self.btn_log = QPushButton("日志")
        self.btn_log.setFlat(True)
        self.btn_log.setToolTip("打开日志文件 %s" % LOG_FILE)
        self.btn_log.clicked.connect(self._open_log)
        prog_layout.addWidget(self.progress, 1)
        prog_layout.addWidget(self.lbl_status)
        prog_layout.addWidget(self.btn_log)
        root.addLayout(prog_layout)

        # --- 结果表格 ---
        self.table = QTableWidget(0, COL_COUNT)
        self.table.setHorizontalHeaderLabels(["", "名称", "所在位置", "目标", "状态"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._context_menu)
        self.table.setSortingEnabled(True)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.setColumnWidth(1, 200)
        self.table.setColumnWidth(2, 280)
        root.addWidget(self.table, 1)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setStyleSheet("color:#616161;")
        root.addWidget(self.lbl_summary)

        if not HAS_PYWIN32:
            QMessageBox.critical(
                self, "缺少依赖",
                "未检测到 pywin32，无法解析快捷方式。\n\n请执行：\npip install pywin32")
            self.btn_scan.setEnabled(False)

    # ---------- 扫描 ----------
    def _open_log(self):
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            if not os.path.exists(LOG_FILE):
                with open(LOG_FILE, "a", encoding="utf-8"):
                    pass
            os.startfile(LOG_FILE)
        except OSError as e:
            QMessageBox.critical(self, "错误", "无法打开日志：\n%s" % e)

    def _pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择扫描目录")
        if d:
            self.edit_custom.setText(d)

    def _collect_roots(self):
        roots = []
        if self.cb_desktop.isChecked():
            roots.append(get_shell_folder(CSIDL_DESKTOPDIRECTORY))
        if self.cb_common_desktop.isChecked():
            roots.append(get_shell_folder(CSIDL_COMMON_DESKTOPDIRECTORY))
        if self.cb_start.isChecked():
            roots.append(get_shell_folder(CSIDL_PROGRAMS))
        if self.cb_common_start.isChecked():
            roots.append(get_shell_folder(CSIDL_COMMON_PROGRAMS))
        if self.cb_quick.isChecked():
            q = quick_launch_dir()
            if q:
                roots.append(q)
        custom = self.edit_custom.text().strip().strip('"')
        if custom and os.path.isdir(custom):
            roots.append(custom)

        if not self.cb_subdirs.isChecked():
            roots = [(r, False) for r in roots]
        else:
            roots = [(r, True) for r in roots]
        return [r for r, _ in roots]

    def _start_scan(self):
        roots = self._collect_roots()
        if not roots:
            QMessageBox.information(self, "提示", "请至少选择一个扫描范围。")
            return

        if self.thread and self.thread.isRunning():
            return

        self.items = []
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self.progress.setRange(0, 0)  # 忙碌指示
        self.btn_scan.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.lbl_status.setText("正在扫描…")

        self.thread = ScanThread(roots, self.cb_url_check.isChecked())
        self.thread.sig_one.connect(self._on_item)
        self.thread.sig_progress.connect(lambda s: self.lbl_status.setText("正在扫描: " + s))
        self.thread.sig_done.connect(self._on_done)
        self.thread.start()

    def _stop_scan(self):
        if self.thread:
            self.thread.stop()
            self.lbl_status.setText("正在停止…")

    def _on_item(self, item):
        self.items.append(item)
        # 扫描过程中仅追加“失效项”，避免大量有效项拖慢 UI
        if self._is_suspect(item.status) and self.cmb_filter.currentIndex() != 2:
            self._append_row(item)

    def _on_done(self, total, broken):
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.btn_scan.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.lbl_status.setText("扫描完成")
        self.lbl_summary.setText("共扫描 %d 个快捷方式，其中失效 %d 个。" % (total, broken))
        self._refresh_table()

    # ---------- 表格 ----------
    def _is_suspect(self, status):
        return status in ABNORMAL_STATES

    def _visible(self, item):
        f = self.cmb_filter.currentIndex()
        if f == 1:
            return self._is_suspect(item.status)
        if f == 2:
            return not self._is_suspect(item.status)
        return True

    def _append_row(self, item):
        r = self.table.rowCount()
        self.table.insertRow(r)

        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        chk.setCheckState(Qt.Unchecked)
        self.table.setItem(r, 0, chk)

        self._set_cell(r, 1, item.name, tooltip=item.path)
        self._set_cell(r, 2, item.dir, tooltip=item.dir)
        self._set_cell(r, 3, item.target or item.detail or "-", tooltip=item.target or item.detail)

        label, color = STATUS_META.get(item.status, ("未知", QColor("#333")))
        st = QTableWidgetItem(label)
        st.setData(Qt.UserRole, item.status)
        if item.detail:
            st.setToolTip(item.detail)
        if item.status in BROKEN_STATES:
            st.setForeground(color)
            f = QFont()
            f.setBold(True)
            st.setFont(f)
        else:
            st.setForeground(color)
        self.table.setItem(r, 4, st)

    def _set_cell(self, row, col, text, tooltip=""):
        it = QTableWidgetItem(text)
        if tooltip:
            it.setToolTip(tooltip)
        self.table.setItem(row, col, it)

    def _refresh_table(self):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        for item in self.items:
            if self._visible(item):
                self._append_row(item)
        self.table.setSortingEnabled(True)
        # 修正状态列颜色（排序不影响颜色，颜色在 _append_row 已设置）
        self.lbl_status.setText("显示 %d / %d 项" % (self.table.rowCount(), len(self.items)))

    def _select_broken(self):
        for r in range(self.table.rowCount()):
            st = self.table.item(r, 4)
            chk = self.table.item(r, 0)
            if st is None or chk is None:
                continue
            abnormal = st.data(Qt.UserRole) in ABNORMAL_STATES
            chk.setCheckState(Qt.Checked if abnormal else Qt.Unchecked)

    def _on_unverified_changed(self, idx):
        move_mode = (idx == 2)
        self.edit_unverified_dir.setVisible(move_mode)
        self.btn_browse_unverified.setVisible(move_mode)
        self.btn_move_unverified.setVisible(move_mode)

    def _pick_unverified_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择移动目标文件夹",
                                             self.edit_unverified_dir.text().strip())
        if d:
            self.edit_unverified_dir.setText(d)

    def _move_unverified(self):
        target_dir = self.edit_unverified_dir.text().strip().strip('"')
        if not target_dir:
            QMessageBox.information(self, "提示", "请先选择目标文件夹。")
            return
        try:
            os.makedirs(target_dir, exist_ok=True)
        except OSError as e:
            QMessageBox.critical(self, "错误", "无法创建目标文件夹：\n%s" % e)
            return

        moved_paths = []
        moved_names = []
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk is None or chk.checkState() != Qt.Checked:
                continue
            st = self.table.item(r, 4)
            if st is None or st.data(Qt.UserRole) not in URL_ABNORMAL_STATES:
                continue
            name_item = self.table.item(r, 1)
            if not name_item or not name_item.toolTip():
                continue
            src = name_item.toolTip()
            if not os.path.exists(src):
                continue
            dst = os.path.join(target_dir, os.path.basename(src))
            base, ext = os.path.splitext(dst)
            n = 1
            while os.path.exists(dst):
                dst = "%s_%d%s" % (base, n, ext)
                n += 1
            try:
                shutil.move(src, dst)
                moved_paths.append(src)
                moved_names.append(os.path.basename(src))
            except OSError as e:
                logger.error("移动失败 %s: %s", src, e)

        if not moved_paths:
            QMessageBox.information(self, "提示",
                                    "没有选中的异常 .url 项被移动。\n请勾选目标行后重试。")
            return

        self.items = [i for i in self.items if i.path not in moved_paths]
        self._refresh_table()
        QMessageBox.information(
            self, "完成",
            "已将 %d 个异常 .url 文件移动到：\n%s" % (
                len(moved_paths), target_dir))

    # ---------- 操作 ----------
    def _delete_selected(self):
        # 第 1 列 tooltip 存有完整路径
        delete_paths = []      # 普通失效项（待删除）
        unknown_paths = []     # 异常 .url 项（按用户选择处理）
        for r in range(self.table.rowCount()):
            chk = self.table.item(r, 0)
            if chk is None or chk.checkState() != Qt.Checked:
                continue
            st = self.table.item(r, 4)
            name_item = self.table.item(r, 1)
            if not name_item or not name_item.toolTip():
                continue
            path = name_item.toolTip()
            if st is not None and st.data(Qt.UserRole) in URL_ABNORMAL_STATES:
                unknown_paths.append(path)
            else:
                delete_paths.append(path)

        if not delete_paths and not unknown_paths:
            QMessageBox.information(self, "提示", "请先勾选要处理的快捷方式。")
            return

        mode = self.cmb_unverified.currentIndex()
        moved_paths = []

        # “移动到指定文件夹”模式：先准备好目标目录
        if mode == 2 and unknown_paths:
            target_dir = self.edit_unverified_dir.text().strip().strip('"')
            if not target_dir:
                QMessageBox.information(self, "提示", "请先填写移动目标文件夹。")
                return
            try:
                os.makedirs(target_dir, exist_ok=True)
            except OSError as e:
                QMessageBox.critical(self, "错误", "无法创建目标文件夹：\n%s" % e)
                return

        # 组装确认信息
        msg = []
        if delete_paths:
            msg.append("删除 %d 个失效项（移入回收站）" % len(delete_paths))
        if unknown_paths:
            if mode == 0:
                msg.append("跳过 %d 个异常 .url 项（当前选择：保留）" % len(unknown_paths))
            elif mode == 1:
                msg.append("连同 %d 个异常 .url 项一起删除" % len(unknown_paths))
            else:
                msg.append("移动 %d 个异常 .url 项到：\n%s"
                           % (len(unknown_paths),
                              self.edit_unverified_dir.text().strip()))
        ret = QMessageBox.question(
            self, "确认操作", "\n".join(msg) + "\n\n是否继续？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret != QMessageBox.Yes:
            return

        # 执行删除
        to_recycle = list(delete_paths)
        if mode == 1:
            to_recycle.extend(unknown_paths)
        ok = fail = 0
        if to_recycle:
            ok, fail = delete_to_recycle_bin(to_recycle)
            if fail:
                logger.error("移入回收站失败 %d 个，涉及文件见下：", fail)
                for p in to_recycle:
                    logger.error("  回收站操作对象: %s", p)

        # 执行移动
        move_fail = 0
        if mode == 2:
            target_dir = self.edit_unverified_dir.text().strip().strip('"')
            for src in unknown_paths:
                if not os.path.exists(src):
                    logger.warning("移动时源文件不存在: %s", src)
                    move_fail += 1
                    continue
                dst = os.path.join(target_dir, os.path.basename(src))
                base, ext = os.path.splitext(dst)
                n = 1
                while os.path.exists(dst):
                    dst = "%s_%d%s" % (base, n, ext)
                    n += 1
                try:
                    shutil.move(src, dst)
                    moved_paths.append(src)
                except OSError as e:
                    logger.error("移动失败 %s: %s", src, e)
                    move_fail += 1

        # 汇总结果
        parts = []
        if to_recycle:
            if fail:
                parts.append("移入回收站：成功 %d，失败 %d（详见日志）" % (ok, fail))
            else:
                parts.append("已将 %d 个移入回收站" % ok)
        if mode == 2:
            parts.append("已移动 %d 个异常 .url 项" % len(moved_paths))
            if move_fail:
                parts.append("移动失败 %d 个（详见日志）" % move_fail)
        if mode == 0 and unknown_paths:
            parts.append("已跳过 %d 个异常 .url 项" % len(unknown_paths))
        QMessageBox.information(self, "完成", "\n".join(parts) if parts else "没有可处理的项。")

        # 更新列表：删除的 + 移动的 都从结果中移除
        removed = set(to_recycle) | set(moved_paths)
        self.items = [i for i in self.items if i.path not in removed]
        self._refresh_table()

    def _context_menu(self, pos):
        r = self.table.rowAt(pos.y())
        if r < 0:
            return
        item = self.table.item(r, 1)
        path = item.toolTip() if item else ""
        if not path:
            return

        target = ""
        for it in self.items:
            if it.path == path:
                target = it.target
                break

        menu = QMenu(self)
        act_open_dir = QAction("打开所在位置", menu)
        act_open_dir.triggered.connect(lambda: self._explore_select(path))
        menu.addAction(act_open_dir)

        if target and os.path.exists(target):
            act_open_target = QAction("打开目标位置", menu)
            act_open_target.triggered.connect(lambda: self._explore_select(target))
            menu.addAction(act_open_target)

        act_check = QAction("重新检测该项", menu)
        act_check.triggered.connect(lambda: self._recheck(path))
        menu.addAction(act_check)

        menu.exec_(self.table.viewport().mapToGlobal(pos))

    @staticmethod
    def _explore_select(path):
        path = os.path.abspath(path)
        if not os.path.exists(path):
            os.startfile(os.path.dirname(path))
            return

        ptr = ctypes.c_void_p()
        hr = ctypes.windll.shell32.SHParseDisplayName(
            path, None, ctypes.byref(ptr), 0, None)
        if hr != 0 or not ptr.value:
            os.startfile(os.path.dirname(path))
            return

        try:
            ctypes.windll.shell32.SHOpenFolderAndSelectItems(ptr, 0, None, 0)
        finally:
            ctypes.windll.ole32.CoTaskMemFree(ptr)

    def _recheck(self, path):
        if not path.lower().endswith(".lnk") or not HAS_PYWIN32:
            return
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass  # Qt 主线程通常已完成 OLE 初始化
        t = ScanThread([], False)
        item = t._check_lnk(path)
        for i, it in enumerate(self.items):
            if it.path == path:
                self.items[i] = item
                break
        else:
            self.items.append(item)
        self._refresh_table()


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 9))
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
