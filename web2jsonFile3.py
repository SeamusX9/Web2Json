import sys
import json
import time
import os
import threading
# ========== 全局屏蔽所有警告 ==========
import warnings
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore')

# 统一导入，避免分散导入导致的初始化问题
from PyQt5.QtGui import QIcon, QCursor, QTextCursor, QFont
from PyQt5.QtCore import (
    QUrl, QDateTime, Qt, pyqtSignal, QObject, QTimer,
    QTranslator, QLocale, QLibraryInfo,
)
from PyQt5.QtWidgets import (QApplication, QMainWindow, QToolBar, QLineEdit, 
                            QAction, QPushButton, QVBoxLayout, QWidget, 
                            QFileDialog, QMessageBox, QProgressBar, QHBoxLayout,
                            QTextEdit, QDockWidget, QTabWidget, QGroupBox,
                            QLabel, QDialog, QDialogButtonBox, QDoubleSpinBox, QCheckBox,
                            QMenu)
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEnginePage, QWebEngineProfile
from PyQt5.QtWebEngineCore import QWebEngineUrlRequestInterceptor, QWebEngineUrlRequestInfo

from json_to_excel_dialog import JsonToExcelDialog

# ========== 解决pynput导入问题（移到程序入口，避免QApplication前调用UI组件） ==========
def check_pynput():
    try:
        from pynput import mouse, keyboard
        return mouse, keyboard
    except ImportError:
        return None, None

class DebugSignalEmitter(QObject):
    debug_signal = pyqtSignal(str)

class RequestInterceptor(QWebEngineUrlRequestInterceptor):
    def __init__(self, parent=None):
        super(RequestInterceptor, self).__init__(parent)
        self.is_recording = False
        self.debug_emitter = DebugSignalEmitter()
        self.debug_signal = self.debug_emitter.debug_signal
        self.save_folder = ""
        
    def interceptRequest(self, info):
        if self.is_recording:
            url = info.requestUrl().toString()
            method = "GET" if info.requestMethod() == b"GET" else "POST"
            self.debug_signal.emit(f"[{method}] 请求: {url}")

# ========== 自动化点击配置窗口 ==========
class AutoClickConfigDialog(QDialog):
    # 子线程（pynput）→ 主线程填入坐标
    fill_coord_signal = pyqtSignal(int, int)

    def __init__(self, parent=None):
        super(AutoClickConfigDialog, self).__init__(parent)
        self.parent_win = parent
        self.setWindowTitle("自动化点击配置")
        self.setMinimumSize(460, 420)
        self.resize(480, 440)
        if parent and not parent.windowIcon().isNull():
            self.setWindowIcon(parent.windowIcon())

        self._ctrl_pressed = False
        self._pick_key_listener = None
        self._pick_mouse_listener = None

        self.fill_coord_signal.connect(self._on_fill_coord)
        self.init_ui()
        self._apply_dialog_style()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_current_mouse_pos)
        self.timer.start(80)

        self._start_coord_pick_listeners()

    def _apply_dialog_style(self):
        self.setStyleSheet("""
            QDialog {
                background: #f8fafc;
            }
            QGroupBox {
                font-size: 9.5pt;
                font-weight: 600;
                color: #1e293b;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                margin-top: 12px;
                padding: 14px 12px 10px 12px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #334155;
            }
            QLineEdit, QDoubleSpinBox {
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px 10px;
                background: #f8fafc;
                font-size: 10pt;
                min-height: 22px;
                selection-background-color: #3b82f6;
            }
            QLineEdit:focus, QDoubleSpinBox:focus {
                border-color: #3b82f6;
                background: #ffffff;
            }
            QCheckBox {
                font-size: 9.5pt;
                color: #334155;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 16px; height: 16px;
            }
            QPushButton {
                border: none;
                border-radius: 6px;
                padding: 7px 16px;
                font-size: 9.5pt;
                font-weight: 600;
            }
        """)

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 14)
        main_layout.setSpacing(10)

        # 顶部说明：Ctrl+左键拾取
        tip = QLabel(
            "按住 <b>Ctrl</b> 再点击鼠标左键，自动将当前坐标填入下方配置"
        )
        tip.setWordWrap(True)
        tip.setAlignment(Qt.AlignCenter)
        tip.setStyleSheet(
            "QLabel {"
            "  background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe;"
            "  border-radius: 8px; padding: 10px 12px; font-size: 9.5pt;"
            "}"
        )
        main_layout.addWidget(tip)

        self.pick_status = QLabel("拾取待命 · 移动鼠标到目标位置后 Ctrl + 左键")
        self.pick_status.setAlignment(Qt.AlignCenter)
        self.pick_status.setStyleSheet("color: #64748b; font-size: 9pt; padding: 2px;")
        main_layout.addWidget(self.pick_status)

        # 实时坐标
        current_pos_group = QGroupBox("当前鼠标实时坐标")
        current_pos_layout = QHBoxLayout(current_pos_group)
        current_pos_layout.addWidget(QLabel("X"))
        self.current_x_label = QLabel("0")
        self.current_x_label.setMinimumWidth(72)
        self.current_x_label.setStyleSheet(
            "color: #2563eb; font-weight: 700; font-size: 12pt; font-family: Consolas, monospace;"
        )
        current_pos_layout.addWidget(self.current_x_label)
        current_pos_layout.addSpacing(24)
        current_pos_layout.addWidget(QLabel("Y"))
        self.current_y_label = QLabel("0")
        self.current_y_label.setMinimumWidth(72)
        self.current_y_label.setStyleSheet(
            "color: #2563eb; font-weight: 700; font-size: 12pt; font-family: Consolas, monospace;"
        )
        current_pos_layout.addWidget(self.current_y_label)
        current_pos_layout.addStretch()
        main_layout.addWidget(current_pos_group)

        # 配置坐标
        config_pos_group = QGroupBox("点击坐标配置")
        config_pos_layout = QHBoxLayout(config_pos_group)
        curr_x, curr_y = self.parent_win.get_current_mouse_pos()
        config_pos_layout.addWidget(QLabel("X"))
        self.config_x_edit = QLineEdit(str(curr_x))
        self.config_x_edit.setMaximumWidth(120)
        config_pos_layout.addWidget(self.config_x_edit)
        config_pos_layout.addSpacing(20)
        config_pos_layout.addWidget(QLabel("Y"))
        self.config_y_edit = QLineEdit(str(curr_y))
        self.config_y_edit.setMaximumWidth(120)
        config_pos_layout.addWidget(self.config_y_edit)
        config_pos_layout.addStretch()
        main_layout.addWidget(config_pos_group)

        # 间隔
        interval_group = QGroupBox("点击间隔")
        interval_layout = QHBoxLayout(interval_group)
        self.interval_spin = QDoubleSpinBox()
        self.interval_spin.setRange(0.1, 999.9)
        self.interval_spin.setValue(1.0)
        self.interval_spin.setSingleStep(0.1)
        self.interval_spin.setSuffix(" 秒")
        self.interval_spin.setMaximumWidth(140)
        interval_layout.addWidget(self.interval_spin)
        esc_hint = QLabel("运行中按 ESC 可停止自动点击")
        esc_hint.setStyleSheet("color: #64748b; font-size: 9pt; font-weight: normal;")
        interval_layout.addWidget(esc_hint)
        interval_layout.addStretch()
        main_layout.addWidget(interval_group)

        self.ctrl_end_check = QCheckBox("每次点击前执行 Ctrl+End，跳转到页面底部")
        main_layout.addWidget(self.ctrl_end_check)

        main_layout.addStretch()

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("确定")
        btn_box.button(QDialogButtonBox.Cancel).setText("取消")
        btn_box.button(QDialogButtonBox.Ok).setStyleSheet(
            "QPushButton { background: #2563eb; color: white; min-width: 88px; }"
            "QPushButton:hover { background: #1d4ed8; }"
        )
        btn_box.button(QDialogButtonBox.Cancel).setStyleSheet(
            "QPushButton { background: #e2e8f0; color: #475569; min-width: 88px; }"
            "QPushButton:hover { background: #cbd5e1; }"
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        main_layout.addWidget(btn_box)

    def _start_coord_pick_listeners(self):
        """全局监听 Ctrl + 左键，用于拾取屏幕坐标"""
        kb = self.parent_win.keyboard
        ms = self.parent_win.mouse

        def on_press(key):
            try:
                if key in (kb.Key.ctrl, kb.Key.ctrl_l, kb.Key.ctrl_r):
                    self._ctrl_pressed = True
            except Exception:
                pass

        def on_release(key):
            try:
                if key in (kb.Key.ctrl, kb.Key.ctrl_l, kb.Key.ctrl_r):
                    self._ctrl_pressed = False
            except Exception:
                pass

        def on_click(x, y, button, pressed):
            if not pressed:
                return
            try:
                if button == ms.Button.left and self._ctrl_pressed:
                    self.fill_coord_signal.emit(int(x), int(y))
            except Exception:
                pass

        self._pick_key_listener = kb.Listener(on_press=on_press, on_release=on_release)
        self._pick_mouse_listener = ms.Listener(on_click=on_click)
        self._pick_key_listener.daemon = True
        self._pick_mouse_listener.daemon = True
        self._pick_key_listener.start()
        self._pick_mouse_listener.start()

    def _stop_coord_pick_listeners(self):
        for listener in (self._pick_key_listener, self._pick_mouse_listener):
            if listener is not None:
                try:
                    listener.stop()
                except Exception:
                    pass
        self._pick_key_listener = None
        self._pick_mouse_listener = None

    def _on_fill_coord(self, x, y):
        self.config_x_edit.setText(str(x))
        self.config_y_edit.setText(str(y))
        self.pick_status.setText(f"已拾取坐标  X={x}  Y={y}")
        self.pick_status.setStyleSheet(
            "color: #059669; font-size: 9.5pt; font-weight: 600; padding: 2px;"
        )

    def refresh_current_mouse_pos(self):
        curr_x, curr_y = self.parent_win.get_current_mouse_pos()
        self.current_x_label.setText(str(curr_x))
        self.current_y_label.setText(str(curr_y))

    def done(self, result):
        self.timer.stop()
        self._stop_coord_pick_listeners()
        super(AutoClickConfigDialog, self).done(result)

    def get_config(self):
        try:
            x = int(self.config_x_edit.text().strip())
            y = int(self.config_y_edit.text().strip())
        except (ValueError, TypeError):
            x, y = self.parent_win.get_current_mouse_pos()
            self.parent_win.append_debug_info_safe(
                f"⚠️ 坐标输入非数字，已使用当前鼠标坐标 X={x} Y={y}"
            )
        return {
            "x": x,
            "y": y,
            "interval": self.interval_spin.value(),
            "ctrl_end": self.ctrl_end_check.isChecked(),
        }

# ========== 核心：自定义WebView 【精简版 无右键代码】 ==========
class CustomWebView(QWebEngineView):
    open_link_in_new_tab = pyqtSignal(QUrl)

    def __init__(self, parent=None, target_url=None):
        super(CustomWebView, self).__init__(parent)
        self.hovered_url = None
        self.main_win = parent
        self.target_url = target_url
        
        # 初始化拦截器
        self.interceptor = RequestInterceptor()
        self.interceptor.debug_signal.connect(parent.append_debug_info_safe)
        
        # 初始化自定义页面
        self.page_obj = CustomWebPage(self.interceptor, self)
        self.setPage(self.page_obj)
        self.page_obj.linkHovered.connect(self.on_link_hovered)
        
        # 如果有目标URL，直接加载
        if self.target_url and self.target_url.isValid():
            self.load(self.target_url)

    def on_link_hovered(self, url):
        self.hovered_url = QUrl(url) if url else None

    def append_debug_info(self, text):
        self.main_win.append_debug_info_safe(text)

    def contextMenuEvent(self, event):
        """自定义中文右键菜单（替代 WebEngine 默认英文菜单）"""
        menu = QMenu(self)
        page = self.page()

        def add_web_action(web_action, text):
            act = page.action(web_action)
            if act is None:
                return
            act.setText(text)
            menu.addAction(act)

        add_web_action(QWebEnginePage.Back, "后退")
        add_web_action(QWebEnginePage.Forward, "前进")
        add_web_action(QWebEnginePage.Reload, "重新加载")
        add_web_action(QWebEnginePage.Stop, "停止")
        menu.addSeparator()

        if self.hovered_url and self.hovered_url.isValid():
            open_tab = menu.addAction("在新标签页中打开链接")
            url = QUrl(self.hovered_url)
            open_tab.triggered.connect(lambda checked=False, u=url: self.open_link_in_new_tab.emit(u))
            add_web_action(QWebEnginePage.CopyLinkToClipboard, "复制链接地址")
            menu.addSeparator()

        add_web_action(QWebEnginePage.Copy, "复制")
        add_web_action(QWebEnginePage.Cut, "剪切")
        add_web_action(QWebEnginePage.Paste, "粘贴")
        add_web_action(QWebEnginePage.SelectAll, "全选")
        menu.addSeparator()
        add_web_action(QWebEnginePage.CopyImageToClipboard, "复制图像")
        add_web_action(QWebEnginePage.CopyImageUrlToClipboard, "复制图像地址")
        add_web_action(QWebEnginePage.DownloadImageToDisk, "图片另存为…")
        add_web_action(QWebEnginePage.DownloadLinkToDisk, "链接另存为…")
        menu.addSeparator()
        add_web_action(QWebEnginePage.ViewSource, "查看页面源代码")

        # 去掉不可用项
        for act in list(menu.actions()):
            if not act.isSeparator() and not act.isEnabled():
                menu.removeAction(act)

        if any(not a.isSeparator() for a in menu.actions()):
            menu.exec_(event.globalPos())

# ========== 核心：自定义WebPage 【修复：current_new_url初始化】 ==========
class CustomWebPage(QWebEnginePage):
    send_webview_to_main = pyqtSignal(object)

    def __init__(self, interceptor, parent=None):
        super(CustomWebPage, self).__init__(parent)
        self.interceptor = interceptor
        self.parent_webview = parent
        self.main_win = parent.main_win
        self.current_new_url = None  # 修复：显式初始化，避免属性不存在错误
        
        self.profile().setUrlRequestInterceptor(self.interceptor)
        self.loadFinished.connect(self.on_load_finished)

    def createWindow(self, window_type):
        # 创建新标签页 正常打开链接
        target_url = self.current_new_url if self.current_new_url else QUrl("about:blank")
        new_webview = CustomWebView(self.main_win, target_url)
        self.send_webview_to_main.emit(new_webview)
        self.current_new_url = None
        return new_webview.page()

    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        if not self.interceptor.is_recording:
            return
        self.main_win.append_debug_info_safe(f"控制台: {message}")
        try:
            if (message.startswith('{') and message.endswith('}')) or (message.startswith('[') and message.endswith(']')):
                json_data = json.loads(message)
                self._save_single_json_file(json_data, 'console')
            elif message.startswith('FETCH_JSON:') or message.startswith('XHR_JSON:'):
                json_str = message.split(':', 1)[1]
                json_data = json.loads(json_str)
                source = 'fetch' if message.startswith('FETCH_JSON:') else 'xhr'
                self._save_single_json_file(json_data, source)
        except Exception as e:
            self.main_win.append_debug_info_safe(f"非JSON数据: {str(e)}")
    
    def _save_single_json_file(self, json_data, source):
        interceptor = self.interceptor
        if not interceptor.save_folder:
            interceptor.save_folder = QFileDialog.getExistingDirectory(None, "选择JSON文件保存目录", os.path.expanduser("~"))
            if not interceptor.save_folder:
                self.main_win.append_debug_info_safe("⚠️ 未选择保存目录，本次JSON数据未保存！")
                self.interceptor.is_recording = False
                return
        current_time = QDateTime.currentDateTime()
        time_str = current_time.toString("yyyyMMdd_HHmmss_zzz")
        timestamp = current_time.toString("yyyy-MM-dd HH:mm:ss.zzz")
        file_name = f"{time_str}_{source}.json"
        file_path = os.path.join(interceptor.save_folder, file_name)
        save_data = {'timestamp': timestamp,'source': source,'data': json_data}
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
            self.main_win.append_debug_info_safe(f"✅ 成功保存JSON: {file_name}")
        except Exception as e:
            self.main_win.append_debug_info_safe(f"❌ 保存失败: {str(e)}")
        
    def on_load_finished(self, ok):
        if ok and self.interceptor.is_recording:
            self.runJavaScript("""
                (function() {
                    const originalFetch = window.fetch;
                    window.fetch = async function(url, options) {
                        const response = await originalFetch(url, options);
                        const clonedResponse = response.clone();
                        try {const data = await clonedResponse.json();console.log(`FETCH_JSON:${JSON.stringify(data)}`);} catch (e) {}
                        return response;
                    };
                })();
                (function() {
                    const originalXhrOpen = XMLHttpRequest.prototype.open;
                    XMLHttpRequest.prototype.open = function(method, url) {
                        this.addEventListener('load', function() {
                            try {
                                if (this.getResponseHeader('Content-Type') && this.getResponseHeader('Content-Type').includes('application/json')) {
                                    const data = JSON.parse(this.responseText);console.log(`XHR_JSON:${JSON.stringify(data)}`);
                                }
                            } catch (e) {}
                        });
                        return originalXhrOpen.apply(this, arguments);
                    };
                })();
                window.open = function(url) {
                    console.log(`WINDOW_OPEN:${url}`);
                    const a = document.createElement('a');a.href=url;a.target='_blank';a.click();
                };
            """)
            self.main_win.append_debug_info_safe("✅ 页面加载完成，已注入JSON捕获脚本")

# ========== 主窗口 【修复：1.重复中央部件 2.优雅关闭键盘监听器】 ==========
class BrowserWindow(QMainWindow):
    # 新增：线程安全的信号，用于子线程向主线程发送日志文本
    signal_send_log = pyqtSignal(str)
    
    window_count = 0
    def __init__(self, mouse, keyboard):
        super(BrowserWindow, self).__init__()
        BrowserWindow.window_count += 1
        # 传入pynput的mouse/keyboard，避免全局导入
        self.mouse = mouse
        self.keyboard = keyboard
        
        # 绑定信号到槽函数，子线程发信号，主线程执行UI操作 - 根治线程UI警告的核心
        self.signal_send_log.connect(self.append_debug_info_ui)
        
        # 自动化点击变量
        self.auto_click_running = False
        self.auto_click_thread = None
        self.auto_click_config = None
        self.mouse_ctrl = self.mouse.Controller()
        self.keyboard_ctrl = self.keyboard.Controller()
        self.key_listener = self.keyboard.Listener(on_press=self.on_key_press)
        self.key_listener.start()

        self.init_ui()
        self.init_debug_window()
        self.init_menu()
        # 图标设置可能写调试日志，必须在 debug_text 创建之后
        self.set_window_icon_safe()
        self.init_tab_widget()

    # ========== 获取当前鼠标坐标方法 ==========
    def get_current_mouse_pos(self):
        pos = QCursor.pos()
        return pos.x(), pos.y()

    # ========== 关键：线程安全的日志打印方法（所有地方都调用这个） ==========
    def append_debug_info_safe(self, text):
        timestamp = time.strftime("%H:%M:%S")
        log_text = f"[{timestamp}] {text}"
        # 发送信号，由主线程执行UI更新，彻底避免线程问题，无任何警告
        self.signal_send_log.emit(log_text)
    
    # ========== 槽函数：主线程执行的UI更新操作 ==========
    def append_debug_info_ui(self, log_text):
        # 调试窗口未就绪时跳过，避免 PyQt 槽内 AttributeError 在 Py3.14 上原生崩溃
        if not hasattr(self, 'debug_text') or self.debug_text is None:
            return
        self.debug_text.append(log_text)
        self.debug_text.moveCursor(QTextCursor.End)

    def init_ui(self):
        self.setWindowTitle('海洋浏览器  JSON采集专用')
        self.setGeometry(100, 100, 1400, 900)
        self.setStyleSheet("""
            QMainWindow { background: #f1f5f9; }
            QMenuBar {
                background: #ffffff; border-bottom: 1px solid #e2e8f0;
                padding: 2px 6px; font-size: 9.5pt;
            }
            QMenuBar::item { padding: 5px 10px; border-radius: 4px; }
            QMenuBar::item:selected { background: #eff6ff; color: #1d4ed8; }
            QMenu {
                background: #ffffff; border: 1px solid #e2e8f0;
                padding: 4px; font-size: 9.5pt;
            }
            QMenu::item { padding: 6px 28px 6px 12px; border-radius: 4px; }
            QMenu::item:selected { background: #eff6ff; color: #1d4ed8; }
            QTabWidget::pane {
                border: 1px solid #e2e8f0;
                background: #ffffff; top: -1px;
            }
            QTabBar::tab {
                background: #e2e8f0; color: #475569; padding: 7px 14px;
                margin-right: 2px; border-top-left-radius: 6px;
                border-top-right-radius: 6px; font-size: 9pt;
            }
            QTabBar::tab:selected {
                background: #ffffff; color: #1e293b; font-weight: 600;
            }
            QDockWidget::title {
                background: #e2e8f0; padding: 6px 8px;
            }
        """)
        self.create_toolbar()

        # 修复：删除重复的中央部件创建，仅保留一次
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumHeight(3)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(
            "QProgressBar { background: #e2e8f0; border: none; }"
            "QProgressBar::chunk { background: #3b82f6; }"
        )
        self.progress_bar.hide()
        self.main_layout.addWidget(self.progress_bar)

    # 修复：安全设置窗口图标，兼容源码/打包两种运行方式
    def set_window_icon_safe(self):
        # 打包后资源在 _MEIPASS；源码运行则取脚本同目录
        if getattr(sys, '_MEIPASS', None):
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        ico_path = os.path.join(base_dir, "icon.ico")
        if os.path.exists(ico_path):
            self.setWindowIcon(QIcon(ico_path))
        else:
            # 兜底：EXE 旁的 icon.ico
            beside_exe = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "icon.ico")
            if os.path.exists(beside_exe):
                self.setWindowIcon(QIcon(beside_exe))
            else:
                self.append_debug_info_safe("⚠️ 未找到 icon.ico，使用默认图标")

    def init_tab_widget(self):
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.tabBarDoubleClicked.connect(lambda x: self.add_new_tab())
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        self.add_new_tab(QUrl("https://www.baidu.com"))
        self.main_layout.addWidget(self.tab_widget)

    def init_debug_window(self):
        self.debug_dock = QDockWidget("调试信息", self)
        self.debug_dock.setObjectName("DebugDock")
        self.debug_dock.setAllowedAreas(Qt.BottomDockWidgetArea | Qt.RightDockWidgetArea)
        # 关闭按钮只隐藏，不销毁，便于从设置菜单重新打开
        self.debug_dock.setAttribute(Qt.WA_DeleteOnClose, False)
        self.debug_dock.setFeatures(
            QDockWidget.DockWidgetClosable
            | QDockWidget.DockWidgetMovable
            | QDockWidget.DockWidgetFloatable
        )
        self.debug_text = QTextEdit()
        self.debug_text.setReadOnly(True)
        self.debug_text.setLineWrapMode(QTextEdit.WidgetWidth)
        self.debug_text.setStyleSheet("QTextEdit { font-size: 9pt; font-family: Consolas, 'Microsoft YaHei UI', monospace; }")
        self.debug_dock.setWidget(self.debug_text)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.debug_dock)
        self.debug_dock.setMinimumHeight(160)
        self.debug_dock.hide()  # 默认不显示调试面板
        self.debug_dock.visibilityChanged.connect(self._on_debug_dock_visibility_changed)

    def init_menu(self):
        """菜单栏：设置 / 帮助"""
        menubar = self.menuBar()
        settings_menu = menubar.addMenu("设置(&S)")

        # 不使用 checkable，避免与普通菜单项左侧对不齐
        self.show_debug_action = QAction("显示调试面板(&D)", self)
        self.show_debug_action.setShortcut("Ctrl+Shift+D")
        self.show_debug_action.setStatusTip("显示或隐藏底部调试信息面板")
        self.show_debug_action.triggered.connect(self.toggle_debug_panel)
        settings_menu.addAction(self.show_debug_action)

        json_excel_action = QAction("JSON 转 Excel(&E)…", self)
        json_excel_action.setStatusTip("按模板勾选字段，批量将捕获的 JSON 导出为 Excel")
        json_excel_action.triggered.connect(self.open_json_to_excel)
        settings_menu.addAction(json_excel_action)

        help_menu = menubar.addMenu("帮助(&H)")
        guide_action = QAction("简易操作指南(&G)…", self)
        guide_action.triggered.connect(self.show_user_guide)
        help_menu.addAction(guide_action)
        help_menu.addSeparator()
        about_action = QAction("关于(&A)…", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def open_json_to_excel(self):
        dialog = JsonToExcelDialog(self)
        dialog.exec_()

    def toggle_debug_panel(self):
        """显示 / 隐藏调试面板"""
        if self.debug_dock.isVisible():
            self.debug_dock.hide()
        else:
            self.debug_dock.show()
            self.debug_dock.raise_()

    def _on_debug_dock_visibility_changed(self, visible):
        """面板显隐时同步菜单文案"""
        if hasattr(self, 'show_debug_action'):
            self.show_debug_action.setText(
                "隐藏调试面板(&D)" if visible else "显示调试面板(&D)"
            )

    def show_user_guide(self):
        QMessageBox.information(
            self,
            "简易操作指南",
            "【浏览】\n"
            "· 地址栏输入网址后回车访问；工具栏可后退 / 前进 / 刷新\n"
            "· 标签栏双击可新建标签页；＋ 可新建窗口\n\n"
            "【采集 JSON】\n"
            "1. 打开目标页面后点击「开始监听」\n"
            "2. 在页面中操作以触发接口请求\n"
            "3. 首次捕获时选择保存目录，JSON 将自动落盘\n"
            "4. 可点「打开目录」查看文件；「重置」清空监听状态\n\n"
            "【自动点击】\n"
            "1. 点「配置」→ 按住 Ctrl 再左键，拾取屏幕坐标\n"
            "2. 设置间隔后「启动点击」；按 ESC 停止\n"
            "3. 可与「开始监听」配合，边点击边采集\n\n"
            "【JSON 转 Excel】\n"
            "设置 → JSON 转 Excel：加载模板并勾选字段，批量导出 xlsx\n\n"
            "【调试面板】\n"
            "设置 → 显示调试面板（Ctrl+Shift+D）\n\n"
            "——\n"
            "请合法使用本工具，禁止用于非法用途；\n"
            "不允许非法获取、使用或存储数据。"
        )

    def show_about(self):
        QMessageBox.information(
            self,
            "关于",
            "海洋浏览器 · JSON 采集专用\n"
            "版本 1.0\n\n"
            "用于在公开网页浏览过程中捕获接口返回的 JSON，\n"
            "并支持自动点击辅助与批量导出 Excel。\n\n"
            "——\n"
            "请合法使用本工具，禁止用于非法用途；\n"
            "不允许非法获取、使用或存储数据。\n"
            "使用者应自行遵守相关法律法规及网站服务条款，\n"
            "因违规使用所产生的后果由使用者自行承担。"
        )

    def _toolbar_btn_style(self, kind="default"):
        """工具栏按钮统一样式；listen_idle / listen_active 为监听开关两态；字号用 pt 适配 DPI"""
        base = (
            "QPushButton {"
            "  border: none; border-radius: 6px; padding: 6px 14px;"
            "  font-size: 10pt; font-weight: 600; min-height: 28px;"
            "}"
            "QPushButton:disabled { color: #9ca3af; background: #e5e7eb; }"
        )
        styles = {
            "default": (
                "QPushButton { background: #f3f4f6; color: #374151; }"
                "QPushButton:hover { background: #e5e7eb; }"
                "QPushButton:pressed { background: #d1d5db; }"
            ),
            "listen_idle": (
                "QPushButton { background: #10b981; color: white; }"
                "QPushButton:hover { background: #059669; }"
                "QPushButton:pressed { background: #047857; }"
            ),
            "listen_active": (
                "QPushButton { background: #ef4444; color: white; }"
                "QPushButton:hover { background: #dc2626; }"
                "QPushButton:pressed { background: #b91c1c; }"
            ),
            "accent": (
                "QPushButton { background: #3b82f6; color: white; }"
                "QPushButton:hover { background: #2563eb; }"
                "QPushButton:pressed { background: #1d4ed8; }"
            ),
            "warn": (
                "QPushButton { background: #f59e0b; color: white; }"
                "QPushButton:hover { background: #d97706; }"
                "QPushButton:pressed { background: #b45309; }"
            ),
        }
        return base + styles.get(kind, styles["default"])

    def _set_listen_btn_state(self, recording):
        """监听按钮：未监听=开始，监听中=停止"""
        if recording:
            self.listen_btn.setText("■ 停止监听")
            self.listen_btn.setToolTip("点击停止 JSON 监听")
            self.listen_btn.setStyleSheet(self._toolbar_btn_style("listen_active"))
            self.listen_btn.setProperty("recording", True)
        else:
            self.listen_btn.setText("▶ 开始监听")
            self.listen_btn.setToolTip("点击开始 JSON 监听")
            self.listen_btn.setStyleSheet(self._toolbar_btn_style("listen_idle"))
            self.listen_btn.setProperty("recording", False)

    def create_toolbar(self):
        toolbar = QToolBar()
        toolbar.setMovable(False)
        toolbar.setStyleSheet(
            "QToolBar { background: #ffffff; border-bottom: 1px solid #e5e7eb; spacing: 6px; padding: 4px 8px; }"
            "QToolButton { padding: 4px 8px; border-radius: 4px; font-size: 12pt; }"
            "QToolButton:hover { background: #f3f4f6; }"
            "QLineEdit { border: 1px solid #d1d5db; border-radius: 6px; padding: 5px 10px; "
            "  background: #f9fafb; selection-background-color: #3b82f6; min-height: 22px; font-size: 10pt; }"
            "QLineEdit:focus { border-color: #3b82f6; background: #ffffff; }"
        )
        self.addToolBar(toolbar)
        back_action = QAction('◀', self)
        back_action.setToolTip('返回')
        back_action.triggered.connect(lambda: self.get_current_webview().back())
        toolbar.addAction(back_action)
        forward_action = QAction('▶', self)
        forward_action.setToolTip('前进')
        forward_action.triggered.connect(lambda: self.get_current_webview().forward())
        toolbar.addAction(forward_action)
        reload_action = QAction('↻', self)
        reload_action.setToolTip('刷新')
        reload_action.triggered.connect(lambda: self.get_current_webview().reload())
        toolbar.addAction(reload_action)
        new_win_action = QAction('＋', self)
        new_win_action.setToolTip('新建窗口')
        new_win_action.triggered.connect(self.create_new_window)
        toolbar.addAction(new_win_action)
        
        self.url_edit = QLineEdit()
        self.url_edit.setMinimumWidth(300)
        self.url_edit.setPlaceholderText("输入网址后回车访问…")
        self.url_edit.returnPressed.connect(self.load_url)
        toolbar.addWidget(self.url_edit)
        
        control_layout = QHBoxLayout()
        control_layout.setContentsMargins(8, 0, 4, 0)
        control_layout.setSpacing(8)
        control_widget = QWidget()
        control_widget.setLayout(control_layout)
        
        # 监听：同一按钮两态切换
        self.listen_btn = QPushButton()
        self.listen_btn.setCursor(Qt.PointingHandCursor)
        self.listen_btn.clicked.connect(self.toggle_recording)
        self._set_listen_btn_state(False)
        control_layout.addWidget(self.listen_btn)

        self.open_folder_btn = QPushButton('📂 打开目录')
        self.open_folder_btn.setCursor(Qt.PointingHandCursor)
        self.open_folder_btn.setStyleSheet(self._toolbar_btn_style("default"))
        self.open_folder_btn.clicked.connect(self.open_save_folder)
        self.open_folder_btn.setEnabled(False)
        control_layout.addWidget(self.open_folder_btn)

        self.clear_btn = QPushButton('↺ 重置')
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setStyleSheet(self._toolbar_btn_style("default"))
        self.clear_btn.clicked.connect(self.reset_status)
        control_layout.addWidget(self.clear_btn)

        control_layout.addSpacing(16)

        self.config_click_btn = QPushButton('⚙ 配置')
        self.config_click_btn.setCursor(Qt.PointingHandCursor)
        self.config_click_btn.setToolTip('配置自动点击坐标与间隔')
        self.config_click_btn.setStyleSheet(self._toolbar_btn_style("accent"))
        self.config_click_btn.clicked.connect(self.open_click_config)
        control_layout.addWidget(self.config_click_btn)

        self.start_click_btn = QPushButton('🖱 启动点击')
        self.start_click_btn.setCursor(Qt.PointingHandCursor)
        self.start_click_btn.setToolTip('按配置开始自动点击（ESC 停止）')
        self.start_click_btn.setStyleSheet(self._toolbar_btn_style("warn"))
        self.start_click_btn.clicked.connect(self.start_auto_click)
        self.start_click_btn.setEnabled(False)
        control_layout.addWidget(self.start_click_btn)
        
        toolbar.addWidget(control_widget)

    def add_new_tab(self, target_url=None):
        web_view = CustomWebView(self, target_url)
        web_view.open_link_in_new_tab.connect(self.add_new_tab)
        web_view.page_obj.send_webview_to_main.connect(self.add_new_tab_from_webview)
        
        tab_title = target_url.host() or "新页面" if (target_url and target_url.isValid()) else "新页面"
        tab_idx = self.tab_widget.addTab(web_view, tab_title)
        web_view.page().titleChanged.connect(lambda t, i=tab_idx: self.tab_widget.setTabText(i, t if t else tab_title))
        web_view.urlChanged.connect(lambda u: self.url_edit.setText(u.toString()))
        web_view.loadProgress.connect(self.progress_bar.setValue)
        web_view.loadStarted.connect(self.progress_bar.show)
        web_view.loadFinished.connect(lambda ok: self.progress_bar.hide())
        
        self.tab_widget.setCurrentIndex(tab_idx)
        self.append_debug_info_safe(f"📄 新增标签页: {tab_title}")

    def add_new_tab_from_webview(self, web_view):
        if isinstance(web_view, CustomWebView):
            web_view.open_link_in_new_tab.connect(self.add_new_tab)
            web_view.page_obj.send_webview_to_main.connect(self.add_new_tab_from_webview)
            
            tab_title = web_view.url().host() or "新页面"
            tab_idx = self.tab_widget.addTab(web_view, tab_title)
            web_view.page().titleChanged.connect(lambda t, i=tab_idx: self.tab_widget.setTabText(i, t if t else tab_title))
            web_view.urlChanged.connect(lambda u: self.url_edit.setText(u.toString()))
            
            self.tab_widget.setCurrentIndex(tab_idx)
            self.append_debug_info_safe(f"🌐 新窗口已转为标签页: {tab_title}")

    # ========== 打开配置窗口 ==========
    def open_click_config(self):
        dialog = AutoClickConfigDialog(self)
        if dialog.exec_():
            self.auto_click_config = dialog.get_config()
            self.start_click_btn.setEnabled(True)
            self.append_debug_info_safe(f"✅ 点击配置完成: X={self.auto_click_config['x']} Y={self.auto_click_config['y']} 间隔={self.auto_click_config['interval']}秒")

    def start_auto_click(self):
        if self.auto_click_running or not self.auto_click_config:
            return
        self.auto_click_running = True
        self.start_click_btn.setEnabled(False)
        self.config_click_btn.setEnabled(False)
        self.auto_click_thread = threading.Thread(target=self.auto_click_task, daemon=True)
        self.auto_click_thread.start()
        self.append_debug_info_safe("▶️ 自动化点击已启动 | ESC键停止")

    def auto_click_task(self):
        while self.auto_click_running:
            try:
                target_x = self.auto_click_config['x']
                target_y = self.auto_click_config['y']

                if self.auto_click_config['ctrl_end']:
                    self.keyboard_ctrl.press(self.keyboard.Key.ctrl)
                    self.keyboard_ctrl.press(self.keyboard.Key.end)
                    time.sleep(0.1)
                    self.keyboard_ctrl.release(self.keyboard.Key.end)
                    self.keyboard_ctrl.release(self.keyboard.Key.ctrl)
                    time.sleep(0.2)

                self.mouse_ctrl.position = (target_x, target_y)
                self.mouse_ctrl.click(self.mouse.Button.left, 1)
                self.append_debug_info_safe(f"🔵 执行点击: X={target_x} Y={target_y}")
                time.sleep(self.auto_click_config['interval'])
            except Exception as e:
                self.append_debug_info_safe(f"❌ 点击异常: {str(e)}")
                time.sleep(self.auto_click_config['interval'])

    def stop_auto_click(self):
        self.auto_click_running = False
        self.start_click_btn.setEnabled(True)
        self.config_click_btn.setEnabled(True)
        self.append_debug_info_safe("⏹️ 自动化点击已停止")

    def on_key_press(self, key):
        if key == self.keyboard.Key.esc and self.auto_click_running:
            self.stop_auto_click()

    # ========== 基础方法 ==========
    def close_tab(self, index):
        if self.tab_widget.count() > 1:
            self.tab_widget.removeTab(index)
        else:
            self.tab_widget.widget(index).load(QUrl("https://www.baidu.com"))
            self.tab_widget.setTabText(index, "百度首页")

    def get_current_webview(self):
        return self.tab_widget.currentWidget()

    def get_current_interceptor(self):
        return self.get_current_webview().interceptor

    def on_tab_changed(self):
        curr_view = self.get_current_webview()
        if curr_view:
            self.url_edit.setText(curr_view.url().toString())

    def load_url(self):
        url_text = self.url_edit.text().strip()
        url = QUrl(url_text)
        if not url.scheme():
            url = QUrl(f"https://{url_text}")
        self.get_current_webview().load(url)

    def create_new_window(self):
        new_window = BrowserWindow(self.mouse, self.keyboard)
        new_window.show()
        new_window.append_debug_info_safe(f"✨ 新建窗口[{BrowserWindow.window_count}]")
        self.append_debug_info_safe(f"✨ 已创建新窗口[{BrowserWindow.window_count}]")

    # 修复：优雅关闭键盘监听器，避免进程残留
    def closeEvent(self, event):
        BrowserWindow.window_count -= 1
        # 停止自动化点击
        self.auto_click_running = False
        # 优雅关闭键盘监听器
        self.key_listener.stop()
        self.key_listener.join(timeout=1)  # 等待1秒，确保线程结束
        # 打印关闭日志
        self.append_debug_info_safe(f"🔚 窗口关闭，剩余窗口: {BrowserWindow.window_count}")
        if BrowserWindow.window_count <= 0:
            QApplication.quit()
        event.accept()
        
    def toggle_recording(self):
        if self.listen_btn.property("recording"):
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        interceptor = self.get_current_interceptor()
        interceptor.is_recording = True
        interceptor.save_folder = ""
        self._set_listen_btn_state(True)
        self.open_folder_btn.setEnabled(False)
        self.append_debug_info_safe('▶️ 开始监听JSON数据...')
        self.get_current_webview().page_obj.on_load_finished(True)
        
    def stop_recording(self):
        interceptor = self.get_current_interceptor()
        interceptor.is_recording = False
        self._set_listen_btn_state(False)
        self.open_folder_btn.setEnabled(len(interceptor.save_folder) > 0)
        self.append_debug_info_safe('⏹️ 停止监听JSON数据')
        
    def open_save_folder(self):
        interceptor = self.get_current_interceptor()
        if interceptor.save_folder and os.path.exists(interceptor.save_folder):
            if sys.platform == "win32":
                os.startfile(interceptor.save_folder)
            elif sys.platform == "darwin":
                os.system(f"open '{interceptor.save_folder}'")
            else:
                os.system(f"xdg-open '{interceptor.save_folder}'")
        else:
            QMessageBox.warning(self, "提示", "保存目录不存在！")
            
    def reset_status(self):
        interceptor = self.get_current_interceptor()
        interceptor.is_recording = False
        interceptor.save_folder = ""
        self._set_listen_btn_state(False)
        self.open_folder_btn.setEnabled(False)
        self.append_debug_info_safe('🔄 已重置监听状态')
        QMessageBox.information(self, '提示', '状态已重置！')

def _enable_windows_dpi_awareness():
    """在创建 QApplication 之前声明 DPI 感知，避免高分屏字体/控件错位"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        # Per-Monitor DPI Aware V2（Win10 1703+）
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)
        ctypes.windll.user32.SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        return
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ========== 程序入口 【核心修复：1.QApplication最先创建 2.统一依赖检测】 ==========
if __name__ == '__main__':
    # 高 DPI：环境变量 + 属性必须在 QApplication 创建之前
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")
    _enable_windows_dpi_awareness()

    # WebEngine 要求：在创建 QApplication 之前设置共享 OpenGL 上下文
    QApplication.setAttribute(Qt.AA_ShareOpenGLContexts)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # 第一步：创建QApplication，必须是第一个UI操作（根治打包闪退核心）
    app = QApplication(sys.argv)
    app.setApplicationName("海洋浏览器-JSON采集专用")
    app.setApplicationVersion("1.0")
    # 加载 Qt 中文翻译，使标准对话框 / 部分右键菜单显示中文
    _qt_translator = QTranslator()
    _trans_path = QLibraryInfo.location(QLibraryInfo.TranslationsPath)
    if _qt_translator.load(QLocale(QLocale.Chinese, QLocale.China), "qtbase", "_", _trans_path):
        app.installTranslator(_qt_translator)
    _qt_translator2 = QTranslator()
    if _qt_translator2.load(QLocale(QLocale.Chinese, QLocale.China), "qt", "_", _trans_path):
        app.installTranslator(_qt_translator2)
    # 非整数缩放（如 125%/150%）时按实际比例，减少字体发糊/过大
    try:
        app.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except AttributeError:
        pass
    # 使用 pt 字号的系统 UI 字体，随 DPI 正常缩放
    _ui_font = QFont("Microsoft YaHei UI", 9)
    if not QFont("Microsoft YaHei UI").exactMatch():
        _ui_font = QFont("Segoe UI", 9)
    app.setFont(_ui_font)
    # 应用级图标（任务栏/关于窗口等），打包时与 EXE icon.ico 一致
    _ico = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))), "icon.ico")
    if os.path.exists(_ico):
        app.setWindowIcon(QIcon(_ico))

    # 第二步：检测PyQtWebEngine依赖（模块顶部已导入，此处仅兜底）
    try:
        from PyQt5.QtWebEngineWidgets import QWebEngineView  # noqa: F401
    except ImportError:
        QMessageBox.critical(None, "错误", "请安装PyQtWebEngine: pip install PyQtWebEngine")
        sys.exit(1)

    # 第三步：检测pynput依赖（移到此处，避免QApplication前调用UI）
    mouse_module, keyboard_module = check_pynput()
    if mouse_module is None or keyboard_module is None:
        QMessageBox.critical(None, "错误", "请安装pynput: pip install pynput")
        sys.exit(1)

    # 第四步：创建主窗口，传入pynput模块，避免全局导入
    window = BrowserWindow(mouse_module, keyboard_module)
    window.show()

    # 第五步：运行应用主循环
    sys.exit(app.exec_())