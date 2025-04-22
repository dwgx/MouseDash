import json
import time
import os
import logging
import psutil
import win32gui
import win32process
import random
from typing import Dict, List, Any

from PyQt5 import QtCore
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QEvent, QObject
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QMessageBox, QVBoxLayout, QWidget, QHBoxLayout, QLabel, QSpinBox,
    QTableWidget, QTableWidgetItem, QPushButton, QInputDialog, QDoubleSpinBox, QKeySequenceEdit,
    QCheckBox, QMenu, QTextBrowser
)
from PyQt5.QtGui import QPixmap, QKeySequence
from PyQt5.QtNetwork import QNetworkAccessManager, QNetworkRequest, QNetworkReply
from pynput import mouse, keyboard
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController, Listener as KeyboardListener
from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, FluentIcon, setTheme, InfoBar, InfoBarPosition,
    SubtitleLabel, BodyLabel, Theme, PrimaryPushButton, PushButton, ComboBox, SpinBox,
    HyperlinkButton, CardWidget
)

# Directory paths
MACRO_DIR = "./macros/"
CONFIG_DIR = "./config/"

# Ensure directories exist
os.makedirs(MACRO_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

# Configure logging
logging.basicConfig(
    filename=os.path.join(MACRO_DIR, 'macro_controller.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)


class PlaybackThread(QThread):
    finished = pyqtSignal()
    interrupted = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self._stop_flag = False

    def run(self):
        try:
            speed_factor = self.parent.speed_spin.value()
            for _ in range(self.parent.play_count_spin.value()):
                if self._stop_flag:
                    break
                self.parent.start_time = time.time()
                for i, event in enumerate(self.parent.events):
                    if self._stop_flag:
                        break
                    event_time = event["time"] / speed_factor
                    elapsed = time.time() - self.parent.start_time
                    if self.parent.prevent_background_detection:
                        event_time += self.get_adaptive_delay(i, len(self.parent.events))
                    time.sleep(max(0, event_time - elapsed))
                    self.parent.start_time = time.time()
                    if event["type"] == "move":
                        self.parent.mouse_controller.position = (event["x"], event["y"])
                    elif event["type"] == "click":
                        button = Button.left if event["button"] == "Button.left" else Button.right
                        if event["pressed"]:
                            self.parent.mouse_controller.press(button)
                        else:
                            self.parent.mouse_controller.release(button)
                        if self.parent.prevent_background_detection:
                            time.sleep(self.get_click_delay())
                    elif event["type"] == "scroll":
                        self.parent.mouse_controller.scroll(event["dx"], event["dy"])
                    elif event["type"] == "key_press":
                        key = event["key"]
                        try:
                            if hasattr(Key, key.replace("Key.", "")):
                                self.parent.keyboard_controller.press(getattr(Key, key.replace("Key.", "")))
                            else:
                                self.parent.keyboard_controller.press(key.strip("'"))
                            if self.parent.prevent_background_detection:
                                time.sleep(self.get_key_delay())
                        except Exception as e:
                            logging.warning(f"无效的按键 {key}: {str(e)}")
                    elif event["type"] == "key_release":
                        key = event["key"]
                        try:
                            if hasattr(Key, key.replace("Key.", "")):
                                self.parent.keyboard_controller.release(getattr(Key, key.replace("Key.", "")))
                            else:
                                self.parent.keyboard_controller.release(key.strip("'"))
                        except Exception as e:
                            logging.warning(f"无效的释放键 {key}: {str(e)}")
                if self.parent.prevent_background_detection and self.parent.play_count_spin.value() > 1:
                    time.sleep(self.get_loop_delay())
            if not self._stop_flag:
                self.finished.emit()
            else:
                self.interrupted.emit()
        except Exception as e:
            self.error.emit(str(e))

    def get_adaptive_delay(self, index: int, total: int) -> float:
        base_delay = 0.02 if index % 5 == 0 else 0.01
        variation = random.uniform(-0.005, 0.005)
        return max(0, base_delay + variation)

    def get_click_delay(self) -> float:
        return random.uniform(0.03, 0.07)

    def get_key_delay(self) -> float:
        return random.uniform(0.02, 0.05)

    def get_loop_delay(self) -> float:
        return random.uniform(0.8, 1.2)

    def stop(self):
        self._stop_flag = True


class DisclaimerWidget(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setObjectName("DisclaimerWidget")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setSpacing(15)
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 15, 15, 15)
        self.title_label = SubtitleLabel("关于 & 免责声明")
        self.author_label = BodyLabel("超级宏 - 宏录制与播放工具")
        self.disclaimer_label = BodyLabel("本软件仅限个人使用。请勿用于非法用途。")
        self.disclaimer_label.setWordWrap(True)
        self.github_link = HyperlinkButton(url="https://github.com/dwgx", text="GitHub 仓库")
        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.author_label)
        card_layout.addWidget(self.disclaimer_label)
        card_layout.addWidget(self.github_link)
        self.avatar_label = QLabel()
        self.avatar_label.setFixedSize(100, 100)
        self.avatar_label.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(self.avatar_label)
        layout.addWidget(card)
        layout.addStretch()


class TutorialWidget(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setObjectName("TutorialWidget")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setSpacing(10)
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 15, 15, 15)
        self.title_label = SubtitleLabel("使用教程")
        self.description_label = BodyLabel("了解如何使用超级宏录制和播放宏")
        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.description_label)
        self.tutorial_browser = QTextBrowser()
        self.tutorial_browser.setOpenExternalLinks(True)
        self.tutorial_browser.setStyleSheet("""
            QTextBrowser { 
                background-color: palette(base); 
                color: palette(text); 
                border: 1px solid palette(mid); 
                padding: 10px;
            }
        """)
        self.tutorial_browser.setHtml("""
            <h2>超级宏使用教程</h2>
            <p><b>超级宏</b> 是一个用于录制和播放鼠标和键盘操作的工具，适用于自动化任务。以下是使用步骤：</p>
            <h3>1. 录制宏</h3>
            <ul>
                <li>在主页选择录制模式（仅鼠标、仅键盘或两者）。</li>
                <li>点击“开始录制”或按快捷键（默认 F1）开始录制。</li>
                <li>执行您想要录制的鼠标和键盘操作。</li>
                <li>点击“停止录制”或按快捷键（默认 F2）结束录制。</li>
                <li>点击“保存录制”保存宏，或“取消录制”丢弃。</li>
            </ul>
            <h3>2. 播放宏</h3>
            <ul>
                <li>确保已录制或加载宏（点击“加载宏”选择保存的 .json 文件）。</li>
                <li>设置播放次数和速度（默认 1 次，1x 速度）。</li>
                <li>点击“播放宏”或按快捷键（默认 F4）开始播放。</li>
                <li>点击“停止播放”或按快捷键（默认 F5）中断播放。</li>
                <li>支持开关模式（默认，按一次播放，按停止键停止）和释放模式（按住播放键播放，松开停止）。</li>
            </ul>
            <h3>3. 管理宏</h3>
            <ul>
                <li>点击“管理宏”进入宏管理页面。</li>
                <li>查看已保存的宏列表，可使用、编辑或删除宏。</li>
            </ul>
            <h3>4. 编辑宏</h3>
            <ul>
                <li>在宏管理页面选择“编辑”或直接进入“编辑宏”页面。</li>
                <li>查看和修改宏事件（鼠标移动、点击、键盘按键等）。</li>
                <li>添加新事件或删除选中的事件，保存更改。</li>
            </ul>
            <h3>5. 设置</h3>
            <ul>
                <li>在设置页面修改快捷键（开始/停止录制、播放/停止播放）。</li>
                <li>选择监控进程（默认全局，或指定某程序）。</li>
                <li>切换播放模式（开关模式或释放模式）。</li>
                <li>启用反检测以模拟人类操作，避免被检测。</li>
                <li>选择界面主题（系统模式、亮色模式、暗色模式）。</li>
            </ul>
            <h3>注意事项</h3>
            <ul>
                <li>建议以管理员权限运行以确保所有功能正常。</li>
                <li>录制时间最长 5 分钟，超时自动停止。</li>
                <li>宏文件保存为 JSON 格式，存储在 ./macros/ 目录下。</li>
                <li>请勿将本软件用于非法用途。</li>
            </ul>
            <p>如有问题，请访问 <a href="https://github.com/dwgx">GitHub 仓库</a> 获取帮助。</p>
        """)
        card_layout.addWidget(self.tutorial_browser)
        layout.addWidget(card)
        layout.addStretch()


class MacroEditor(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setSpacing(10)
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 15, 15, 15)
        self.title_label = SubtitleLabel("宏编辑器")
        self.description_label = BodyLabel("查看和编辑录制的宏事件")
        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.description_label)
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["类型", "时间 (秒)", "X坐标", "Y坐标", "详情"])
        self.table.setEditTriggers(QTableWidget.DoubleClicked)
        self.table.setStyleSheet("""
            QTableWidget { background-color: palette(base); color: palette(text); gridline-color: palette(mid); }
            QTableWidget::item:selected { background-color: palette(highlight); color: palette(highlightedText); }
        """)
        card_layout.addWidget(self.table)
        button_layout = QHBoxLayout()
        self.add_button = PushButton("添加事件", icon=FluentIcon.ADD)
        self.delete_button = PushButton("删除选中", icon=FluentIcon.DELETE)
        self.save_button = PrimaryPushButton("保存更改", icon=FluentIcon.SAVE)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.save_button)
        card_layout.addLayout(button_layout)
        layout.addWidget(card)
        layout.addStretch()
        self.add_button.clicked.connect(self.add_event)
        self.delete_button.clicked.connect(self.delete_event)
        self.save_button.clicked.connect(self.save_changes)
        self.load_events()

    def load_events(self):
        self.table.setRowCount(0)
        try:
            events = getattr(self.parent, 'events', [])
            for event in events:
                row = self.table.rowCount()
                self.table.insertRow(row)
                self.table.setItem(row, 0, QTableWidgetItem(event["type"]))
                self.table.setItem(row, 1, QTableWidgetItem(f"{event['time']:.2f}"))
                x = event.get("x", "-")
                y = event.get("y", "-")
                self.table.setItem(row, 2, QTableWidgetItem(str(x)))
                self.table.setItem(row, 3, QTableWidgetItem(str(y)))
                details = self.get_event_details(event)
                self.table.setItem(row, 4, QTableWidgetItem(details))
        except Exception as e:
            self.parent.show_warning("加载失败", f"加载失败: {str(e)}")
            logging.error(f"加载宏事件失败: {str(e)}")

    def get_event_details(self, event: Dict[str, Any]) -> str:
        if event["type"] == "click":
            return f"按钮: {event['button']}, 按下: {event['pressed']}"
        elif event["type"] == "scroll":
            return f"DX: {event['dx']}, DY: {event['dy']}"
        elif event["type"] in ["key_press", "key_release"]:
            return f"按键: {event['key']}"
        return "-"

    def add_event(self):
        event_type, ok = QInputDialog.getItem(
            self, "添加事件", "选择事件类型",
            ["move", "click", "scroll", "key_press", "key_release"], 0, False
        )
        if ok:
            event = {"type": event_type, "time": 0.0}
            if event_type in ["move", "click", "scroll"]:
                event["x"] = 0
                event["y"] = 0
            if event_type == "click":
                event["button"] = "Button.left"
                event["pressed"] = True
            elif event_type == "scroll":
                event["dx"] = 0
                event["dy"] = 0
            elif event_type in ["key_press", "key_release"]:
                event["key"] = "'a'"
            self.parent.events.append(event)
            self.load_events()
            logging.info(f"添加新事件: {event_type}")

    def delete_event(self):
        selected_rows = set(index.row() for index in self.table.selectedIndexes())
        if not selected_rows:
            self.parent.show_warning("无选中", "请先选择要删除的事件")
            return
        for row in sorted(selected_rows, reverse=True):
            self.parent.events.pop(row)
        self.load_events()
        logging.info(f"删除事件: {selected_rows}")

    def save_changes(self):
        try:
            for row in range(self.table.rowCount()):
                event = self.parent.events[row]
                event["type"] = self.table.item(row, 0).text()
                try:
                    event["time"] = float(self.table.item(row, 1).text())
                except ValueError:
                    raise ValueError("时间必须是数字")
                if event["type"] in ["move", "click", "scroll"]:
                    try:
                        event["x"] = float(self.table.item(row, 2).text())
                        event["y"] = float(self.table.item(row, 3).text())
                    except ValueError:
                        raise ValueError("X/Y坐标必须是数字")
                if event["type"] == "click":
                    details = self.table.item(row, 4).text()
                    event["button"] = details.split(",")[0].split(":")[1].strip()
                    event["pressed"] = "True" in details
                elif event["type"] == "scroll":
                    details = self.table.item(row, 4).text()
                    try:
                        event["dx"] = int(details.split(",")[0].split(":")[1].strip())
                        event["dy"] = int(details.split(",")[1].split(":")[1].strip())
                    except ValueError:
                        raise ValueError("DX/DY必须是整数")
                elif event["type"] in ["key_press", "key_release"]:
                    event["key"] = self.table.item(row, 4).text().split(":")[1].strip()
            self.parent.show_success("成功", "宏事件更新成功")
            logging.info("宏事件已修改")
        except Exception as e:
            self.parent.show_error("保存失败", f"保存失败: {str(e)}")
            logging.error(f"保存宏事件失败: {str(e)}")


class MacroManager(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setSpacing(10)
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 15, 15, 15)
        self.title_label = SubtitleLabel("宏管理")
        self.description_label = BodyLabel("管理已保存的宏")
        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.description_label)
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["文件名", "事件数", "操作"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setColumnWidth(0, 300)
        self.table.setColumnWidth(1, 100)
        self.table.setColumnWidth(2, 100)
        card_layout.addWidget(self.table)
        button_layout = QHBoxLayout()
        self.refresh_button = PrimaryPushButton("刷新列表", icon=FluentIcon.SYNC)
        button_layout.addWidget(self.refresh_button)
        card_layout.addLayout(button_layout)
        layout.addWidget(card)
        layout.addStretch()
        self.refresh_button.clicked.connect(self.load_macros)
        self.load_macros()

    def load_macros(self):
        self.table.setRowCount(0)
        try:
            for file_name in os.listdir(MACRO_DIR):
                if file_name.endswith(".json"):
                    file_path = os.path.join(MACRO_DIR, file_name)
                    with open(file_path, 'r', encoding='utf-8') as f:
                        events = json.load(f)
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    self.table.setItem(row, 0, QTableWidgetItem(file_name))
                    self.table.setItem(row, 1, QTableWidgetItem(str(len(events))))
                    action_button = PushButton("操作", icon=FluentIcon.MORE)
                    action_button.setMinimumWidth(80)
                    self.table.setCellWidget(row, 2, action_button)
                    action_button.clicked.connect(lambda _, fn=file_name: self.show_action_menu(fn, action_button))
                    logging.info(f"创建操作按钮: {file_name}")
        except Exception as e:
            self.parent.show_error("加载失败", f"加载失败: {str(e)}")
            logging.error(f"加载宏列表失败: {str(e)}")

    def show_action_menu(self, file_name: str, button: QPushButton):
        menu = QMenu(self)
        use_action = menu.addAction("使用")
        edit_action = menu.addAction("编辑")
        delete_action = menu.addAction("删除")
        use_action.triggered.connect(lambda: self.use_macro(file_name))
        edit_action.triggered.connect(lambda: self.edit_macro(file_name))
        delete_action.triggered.connect(lambda: self.delete_macro(file_name))
        button_pos = button.mapToGlobal(QtCore.QPoint(0, button.height()))
        menu.exec_(button_pos)

    def use_macro(self, file_name: str):
        try:
            file_path = os.path.join(MACRO_DIR, file_name)
            with open(file_path, 'r', encoding='utf-8') as f:
                self.parent.events = json.load(f)
            self.parent.show_success("成功", f"已加载宏: {file_name}")
            logging.info(f"加载宏: {file_name}")
        except Exception as e:
            self.parent.show_error("加载失败", f"加载失败: {str(e)}")
            logging.error(f"加载宏 {file_name} 失败: {str(e)}")

    def edit_macro(self, file_name: str):
        try:
            file_path = os.path.join(MACRO_DIR, file_name)
            with open(file_path, 'r', encoding='utf-8') as f:
                self.parent.events = json.load(f)
            self.parent.editor_widget.load_events()
            self.parent.stackedWidget.setCurrentWidget(self.parent.editor_widget)
            self.parent.show_success("成功", f"已加载宏: {file_name}")
            logging.info(f"编辑宏: {file_name}")
        except Exception as e:
            self.parent.show_error("编辑失败", f"编辑失败 {file_name}: {str(e)}")
            logging.error(f"编辑宏 {file_name} 失败: {str(e)}")

    def delete_macro(self, file_name: str):
        reply = QMessageBox.question(
            self, f"确认删除: {file_name}",
            f"确认删除: {file_name}",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                os.remove(os.path.join(MACRO_DIR, file_name))
                self.load_macros()
                self.parent.show_success("成功", f"已删除宏: {file_name}")
                logging.info(f"删除宏: {file_name}")
            except Exception as e:
                self.parent.show_error("删除失败", f"删除失败 {file_name}: {str(e)}")
                logging.error(f"删除宏 {file_name} 失败: {str(e)}")


class SettingsWidget(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setSpacing(10)
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 15, 15, 15)
        self.title_label = SubtitleLabel("设置")
        self.description_label = BodyLabel("配置应用设置")
        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.description_label)

        # Shortcuts
        shortcut_layout = QVBoxLayout()
        start_layout = QHBoxLayout()
        self.start_key_label = BodyLabel("开始录制快捷键")
        self.start_key_edit = QKeySequenceEdit(QKeySequence("F1"))
        start_layout.addWidget(self.start_key_label)
        start_layout.addWidget(self.start_key_edit)
        shortcut_layout.addLayout(start_layout)

        stop_layout = QHBoxLayout()
        self.stop_key_label = BodyLabel("停止录制快捷键")
        self.stop_key_edit = QKeySequenceEdit(QKeySequence("F2"))
        stop_layout.addWidget(self.stop_key_label)
        stop_layout.addWidget(self.stop_key_edit)
        shortcut_layout.addLayout(stop_layout)

        play_layout = QHBoxLayout()
        self.play_key_label = BodyLabel("播放宏快捷键")
        self.play_key_edit = QKeySequenceEdit(QKeySequence("F4"))
        play_layout.addWidget(self.play_key_label)
        play_layout.addWidget(self.play_key_edit)
        shortcut_layout.addLayout(play_layout)

        stop_play_layout = QHBoxLayout()
        self.stop_play_key_label = BodyLabel("停止播放快捷键")
        self.stop_play_key_edit = QKeySequenceEdit(QKeySequence("F5"))
        stop_play_layout.addWidget(self.stop_play_key_label)
        stop_play_layout.addWidget(self.stop_play_key_edit)
        shortcut_layout.addLayout(stop_play_layout)
        card_layout.addLayout(shortcut_layout)

        # Process Monitoring
        process_layout = QHBoxLayout()
        self.process_label = BodyLabel("监控进程")
        self.process_combo = ComboBox()
        self.process_combo.addItem("全局")
        self.populate_processes()
        self.process_combo.setCurrentIndex(0)
        process_layout.addWidget(self.process_label)
        process_layout.addWidget(self.process_combo)
        card_layout.addLayout(process_layout)

        # Playback Mode
        playback_mode_layout = QHBoxLayout()
        self.playback_mode_label = BodyLabel("播放模式")
        self.playback_mode_combo = ComboBox()
        self.playback_mode_combo.addItems(["开关模式", "释放模式"])
        self.playback_mode_combo.setCurrentIndex(0)
        playback_mode_layout.addWidget(self.playback_mode_label)
        playback_mode_layout.addWidget(self.playback_mode_combo)
        card_layout.addLayout(playback_mode_layout)

        # Anti-Detection
        anti_detection_layout = QVBoxLayout()
        self.prevent_background_detection_checkbox = QCheckBox("启用反检测")
        self.prevent_background_detection_checkbox.setChecked(True)
        anti_detection_layout.addWidget(self.prevent_background_detection_checkbox)
        card_layout.addLayout(anti_detection_layout)

        # Theme
        theme_layout = QHBoxLayout()
        self.theme_label = BodyLabel("主题")
        self.theme_combo = ComboBox()
        self.theme_combo.addItems(["系统模式", "亮色模式", "暗色模式"])
        self.theme_combo.setCurrentIndex(0)
        self.theme_combo.currentIndexChanged.connect(self.apply_theme)
        theme_layout.addWidget(self.theme_label)
        theme_layout.addWidget(self.theme_combo)
        card_layout.addLayout(theme_layout)

        self.apply_button = PrimaryPushButton("应用设置", icon=FluentIcon.SAVE)
        self.apply_button.clicked.connect(self.apply_settings)
        card_layout.addWidget(self.apply_button)
        layout.addWidget(card)
        layout.addStretch()

    def populate_processes(self):
        try:
            processes = set()
            for proc in psutil.process_iter(['name']):
                try:
                    name = proc.info['name'].lower()
                    if name.endswith('.exe'):
                        processes.add(name)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            sorted_processes = sorted(processes)
            self.process_combo.addItems(sorted_processes)
        except Exception as e:
            logging.error(f"无法列举进程: {str(e)}")

    def apply_settings(self):
        try:
            start_seq = self.start_key_edit.keySequence().toString()
            stop_seq = self.stop_key_edit.keySequence().toString()
            play_seq = self.play_key_edit.keySequence().toString()
            stop_play_seq = self.stop_play_key_edit.keySequence().toString()
            if not all([start_seq, stop_seq, play_seq, stop_play_seq]):
                raise ValueError("快捷键不能为空")
            sequences = {start_seq, stop_seq, play_seq, stop_play_seq}
            if len(sequences) != 4:
                raise ValueError("快捷键必须唯一")
            process_name = self.process_combo.currentText()
            playback_mode = self.playback_mode_combo.currentText() == "释放模式"
            self.parent.prevent_background_detection = self.prevent_background_detection_checkbox.isChecked()
            self.parent.update_settings(start_seq, stop_seq, play_seq, stop_play_seq, process_name, playback_mode)
            self.parent.show_success("成功", "设置更新成功")
            logging.info(
                f"设置更新: 开始={start_seq}, 停止={stop_seq}, 播放={play_seq}, 停止播放={stop_play_seq}, 进程={process_name}, 反检测={self.parent.prevent_background_detection}, 播放模式={'释放' if playback_mode else '开关'}")
        except Exception as e:
            self.parent.show_error("错误", f"错误: {str(e)}")
            logging.error(f"应用设置失败: {str(e)}")

    def apply_theme(self, index: int):
        themes = ['auto', 'light', 'dark']
        self.parent.set_theme(themes[index])
        self.parent.save_config()


class PynputSignals(QObject):
    move_signal = pyqtSignal(int, int)
    click_signal = pyqtSignal(int, int, str, bool)
    scroll_signal = pyqtSignal(int, int, int, int)
    press_signal = pyqtSignal(str)
    release_signal = pyqtSignal(str)
    shortcut_signal = pyqtSignal(str)


class ListenerThread(QThread):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.hotkey_listener = None
        self.is_running = False

    def run(self):
        self.is_running = True
        try:
            while self.is_running:
                # Keep the thread alive; hotkeys are handled by GlobalHotKeys
                time.sleep(0.1)
        except Exception as e:
            self.parent.logger.error(f"监听器线程异常: {str(e)}")
            self.parent.show_error("监听器错误", f"监听器错误: {str(e)}")

    def update_hotkeys(self, shortcuts: Dict[str, str]):
        """Update the global hotkeys with the provided shortcuts."""
        try:
            # Stop the existing listener if it exists
            self.stop_listener()

            # Convert shortcuts to pynput format and create handlers
            hotkey_map = {}
            for seq, action in shortcuts.items():
                # Convert PyQt5 QKeySequence format to pynput format
                pynput_seq = self.convert_to_pynput_format(seq)
                if pynput_seq:
                    hotkey_map[pynput_seq] = lambda act=action: self.parent.pynput_signals.shortcut_signal.emit(act)

            if not hotkey_map:
                self.parent.logger.warning("未设置有效的快捷键")
                return

            # Start a new GlobalHotKeys listener
            self.hotkey_listener = keyboard.GlobalHotKeys(hotkey_map)
            self.hotkey_listener.start()
            self.parent.logger.info(f"已更新全局快捷键: {list(hotkey_map.keys())}")
        except Exception as e:
            self.parent.logger.error(f"更新快捷键失败: {str(e)}")
            self.parent.show_error("快捷键错误", f"更新快捷键失败: {str(e)}")

    def convert_to_pynput_format(self, seq: str) -> str:
        """Convert a PyQt5 QKeySequence string to pynput hotkey format."""
        try:
            seq = seq.lower()
            parts = seq.split('+')
            formatted_parts = []

            for part in parts:
                part = part.strip()
                # Handle modifiers
                if part == 'ctrl':
                    formatted_parts.append('<ctrl>')
                elif part == 'shift':
                    formatted_parts.append('<shift>')
                elif part == 'alt':
                    formatted_parts.append('<alt>')
                # Handle function keys
                elif part.startswith('f') and part[1:].isdigit():
                    formatted_parts.append(f'<{part}>')
                # Handle regular keys
                else:
                    formatted_parts.append(part)

            return '+'.join(formatted_parts)
        except Exception as e:
            self.parent.logger.error(f"转换快捷键格式失败 ({seq}): {str(e)}")
            return ""

    def stop_listener(self):
        """Stop the hotkey listener if it's running."""
        if self.hotkey_listener:
            try:
                self.hotkey_listener.stop()
                self.hotkey_listener = None
                self.parent.logger.info("全局快捷键监听器已停止")
            except Exception as e:
                self.parent.logger.error(f"停止全局快捷键监听器失败: {str(e)}")

    def stop(self):
        """Stop the thread and the listener."""
        self.is_running = False
        self.stop_listener()
        self.quit()
        self.wait()


class MacroController(FluentWindow):
    def __init__(self):
        super().__init__()
        # Initialize attributes
        self.theme = 'auto'
        self.events: List[Dict[str, Any]] = []
        self.playback_thread = None
        self.listener_thread = None
        self.start_shortcut = "F1"
        self.stop_shortcut = "F2"
        self.play_shortcut = "F4"
        self.stop_play_shortcut = "F5"
        self.target_process = "全局"
        self.prevent_background_detection = True
        self.playback_mode = False  # False: Toggle mode, True: Hold mode
        self.recording = False
        self.start_time = None
        self.mouse_listener = None
        self.keyboard_listener = None
        self.last_move_time = 0.0
        self.is_play_key_pressed = False
        self.pynput_signals = PynputSignals()
        self.pynput_signals.move_signal.connect(self.handle_move)
        self.pynput_signals.click_signal.connect(self.handle_click)
        self.pynput_signals.scroll_signal.connect(self.handle_scroll)
        self.pynput_signals.press_signal.connect(self.handle_press)
        self.pynput_signals.release_signal.connect(self.handle_release)
        self.pynput_signals.shortcut_signal.connect(self.handle_shortcut)
        self.network_manager = QNetworkAccessManager()
        self.pressed_keys = set()
        self.mouse_controller = MouseController()
        self.keyboard_controller = KeyboardController()
        self.timer = QTimer()
        self.main_widget = None
        self.manager_widget = None
        self.editor_widget = None
        self.settings_widget = None
        self.disclaimer_widget = None
        self.tutorial_widget = None
        self.start_button = None
        self.stop_button = None
        self.save_button = None
        self.cancel_button = None
        self.play_button = None
        self.stop_playback_button = None
        self.load_button = None
        self.manager_button = None
        self.mode_combo = None
        self.play_count_spin = None
        self.speed_spin = None
        self.time_label = None
        self.event_label = None

        self.init_logger()
        self.init_ui()
        self.init_recorder()
        self.load_config()
        self.start_global_listener()
        self.load_avatar()
        self.logger.info("应用启动")
        if not self.is_admin():
            self.show_warning("权限警告", "请以管理员权限运行以确保所有功能正常")

    def is_admin(self) -> bool:
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except Exception:
            return False

    def init_logger(self):
        self.logger = logging.getLogger(__name__)
        self.logger.info("日志初始化")

    def save_config(self):
        config_file = os.path.join(CONFIG_DIR, 'config.json')
        try:
            config = {
                'theme': self.theme,
                'speed': self.speed_spin.value(),
                'start_shortcut': self.start_shortcut,
                'stop_shortcut': self.stop_shortcut,
                'play_shortcut': self.play_shortcut,
                'stop_play_shortcut': self.stop_play_shortcut,
                'target_process': self.target_process,
                'playback_mode': self.playback_mode,
                'prevent_background_detection': self.prevent_background_detection
            }
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            self.logger.error(f"保存配置失败: {str(e)}")
            self.show_error("配置错误", f"配置错误: {str(e)}")

    def load_config(self):
        config_file = os.path.join(CONFIG_DIR, 'config.json')
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.theme = config.get('theme', 'auto')
                    speed = config.get('speed', 100)
                    self.start_shortcut = config.get('start_shortcut', 'F1')
                    self.stop_shortcut = config.get('stop_shortcut', 'F2')
                    self.play_shortcut = config.get('play_shortcut', 'F4')
                    self.stop_play_shortcut = config.get('stop_play_shortcut', 'F5')
                    self.target_process = config.get('target_process', '全局')
                    self.playback_mode = config.get('playback_mode', False)
                    self.prevent_background_detection = config.get('prevent_background_detection', True)
                    self.speed_spin.setValue(speed)
                    self.update_settings(
                        self.start_shortcut,
                        self.stop_shortcut,
                        self.play_shortcut,
                        self.stop_play_shortcut,
                        self.target_process,
                        self.playback_mode
                    )
                self.set_theme(self.theme)
        except Exception as e:
            self.logger.error(f"加载配置失败: {str(e)}")
            self.show_error("配置错误", f"配置错误: {str(e)}")

    def set_theme(self, theme: str):
        self.theme = theme
        if theme == 'auto':
            setTheme(Theme.AUTO)
        elif theme == 'light':
            setTheme(Theme.LIGHT)
        elif theme == 'dark':
            setTheme(Theme.DARK)
        for widget in [self.main_widget, self.manager_widget, self.editor_widget,
                       self.settings_widget, self.disclaimer_widget, self.tutorial_widget]:
            if widget:
                widget.setStyleSheet("")

    def load_avatar(self):
        url = "http://q2.qlogo.cn/headimg_dl?dst_uin=136666451&spec=100"
        request = QNetworkRequest(QtCore.QUrl(url))
        reply = self.network_manager.get(request)
        reply.finished.connect(lambda: self.on_avatar_loaded(reply))

    def on_avatar_loaded(self, reply: QNetworkReply):
        try:
            data = reply.readAll()
            pixmap = QPixmap()
            pixmap.loadFromData(data)
            if not pixmap.isNull() and self.disclaimer_widget:
                self.disclaimer_widget.avatar_label.setPixmap(
                    pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                self.logger.warning("头像加载失败: 数据无效或关于页面未初始化")
        except Exception as e:
            self.logger.error(f"头像加载失败: {str(e)}")
        reply.deleteLater()

    def init_ui(self):
        self.setWindowTitle("老鼠大师")
        self.resize(900, 600)
        self.main_widget = QWidget()
        self.main_widget.setObjectName("MainInterface")
        main_layout = QVBoxLayout(self.main_widget)
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.setSpacing(15)

        # Status Card
        status_card = CardWidget()
        status_layout = QVBoxLayout(status_card)
        status_layout.setContentsMargins(15, 15, 15, 15)
        self.status_card_title = SubtitleLabel("状态")
        self.author_label = BodyLabel("作者: dwgx")
        status_info_layout = QHBoxLayout()
        self.time_label = BodyLabel("录制时间: 0.0s")
        self.event_label = BodyLabel("事件数: 0")
        status_info_layout.addWidget(self.time_label)
        status_info_layout.addWidget(self.event_label)
        status_layout.addWidget(self.status_card_title)
        status_layout.addWidget(self.author_label)
        status_layout.addLayout(status_info_layout)
        main_layout.addWidget(status_card)

        # Recording Control Card
        recording_card = CardWidget()
        recording_layout = QVBoxLayout(recording_card)
        recording_layout.setContentsMargins(15, 15, 15, 15)
        self.recording_card_title = SubtitleLabel("录制控制")
        recording_layout.addWidget(self.recording_card_title)

        mode_layout = QHBoxLayout()
        self.mode_label = BodyLabel("录制模式")
        self.mode_combo = ComboBox()
        self.mode_combo.addItems(["仅鼠标", "仅键盘", "鼠标和键盘"])
        self.mode_combo.setCurrentIndex(2)
        mode_layout.addWidget(self.mode_label)
        mode_layout.addWidget(self.mode_combo)
        recording_layout.addLayout(mode_layout)

        control_layout = QHBoxLayout()
        self.start_button = PrimaryPushButton(f"开始录制 ({self.start_shortcut})", icon=FluentIcon.PLAY)
        self.stop_button = PushButton(f"停止录制 ({self.stop_shortcut})", icon=FluentIcon.CANCEL)
        self.save_button = PushButton("保存录制", icon=FluentIcon.SAVE)
        self.cancel_button = PushButton("取消录制", icon=FluentIcon.CANCEL)
        control_layout.addWidget(self.start_button)
        control_layout.addWidget(self.stop_button)
        control_layout.addWidget(self.save_button)
        control_layout.addWidget(self.cancel_button)
        recording_layout.addLayout(control_layout)
        main_layout.addWidget(recording_card)

        # Playback Control Card
        playback_card = CardWidget()
        playback_layout = QVBoxLayout(playback_card)
        playback_layout.setContentsMargins(15, 15, 15, 15)
        self.playback_card_title = SubtitleLabel("播放控制")
        playback_layout.addWidget(self.playback_card_title)

        play_control_layout = QHBoxLayout()
        self.play_button = PushButton(f"播放宏 ({self.play_shortcut})", icon=FluentIcon.PLAY)
        self.stop_playback_button = PushButton(f"停止播放 ({self.stop_play_shortcut})", icon=FluentIcon.CANCEL)
        self.load_button = PushButton("加载宏", icon=FluentIcon.FOLDER)
        play_control_layout.addWidget(self.play_button)
        play_control_layout.addWidget(self.stop_playback_button)
        play_control_layout.addWidget(self.load_button)
        playback_layout.addLayout(play_control_layout)

        settings_layout = QHBoxLayout()
        self.play_count_label = BodyLabel("播放次数")
        self.play_count_spin = SpinBox()
        self.play_count_spin.setRange(1, 100)
        self.play_count_spin.setValue(1)
        self.speed_label = BodyLabel("播放速度 (x)")
        self.speed_spin = QDoubleSpinBox()
        self.speed_spin.setRange(1, 99999999999)
        self.speed_spin.setValue(100)
        self.speed_spin.setSingleStep(50)
        self.speed_spin.setKeyboardTracking(True)
        self.speed_spin.valueChanged.connect(self.validate_speed)
        settings_layout.addWidget(self.play_count_label)
        settings_layout.addWidget(self.play_count_spin)
        settings_layout.addWidget(self.speed_label)
        settings_layout.addWidget(self.speed_spin)
        playback_layout.addLayout(settings_layout)
        main_layout.addWidget(playback_card)

        # Manage Macros Button
        self.manager_button = PrimaryPushButton("管理宏", icon=FluentIcon.FOLDER)
        self.manager_button.clicked.connect(self.show_macro_manager)
        main_layout.addWidget(self.manager_button)
        main_layout.addStretch()

        # Initialize Other Pages
        self.disclaimer_widget = DisclaimerWidget(self)
        self.disclaimer_widget.setObjectName("DisclaimerInterface")
        self.tutorial_widget = TutorialWidget(self)
        self.tutorial_widget.setObjectName("TutorialWidget")
        self.settings_widget = SettingsWidget(self)
        self.settings_widget.setObjectName("SettingsInterface")
        self.editor_widget = MacroEditor(self)
        self.editor_widget.setObjectName("EditorInterface")
        self.manager_widget = MacroManager(self)
        self.manager_widget.setObjectName("ManagerInterface")

        # Set Initial Button States
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.cancel_button.setEnabled(False)
        self.stop_playback_button.setEnabled(False)

        # Connect Signals
        self.start_button.clicked.connect(self.start_recording)
        self.stop_button.clicked.connect(self.stop_recording)
        self.save_button.clicked.connect(self.save_recording)
        self.cancel_button.clicked.connect(self.cancel_recording)
        self.play_button.clicked.connect(self.start_playback)
        self.stop_playback_button.clicked.connect(self.stop_playback)
        self.load_button.clicked.connect(self.load_macro)

        # Navigation Bar
        self.addSubInterface(self.main_widget, FluentIcon.HOME, "主页")
        self.addSubInterface(self.manager_widget, FluentIcon.FOLDER, "宏管理")
        self.addSubInterface(self.editor_widget, FluentIcon.EDIT, "编辑宏")
        self.addSubInterface(self.tutorial_widget, FluentIcon.BOOK_SHELF, "使用教程",
                             position=NavigationItemPosition.BOTTOM)
        self.addSubInterface(self.settings_widget, FluentIcon.SETTING, "设置", position=NavigationItemPosition.BOTTOM)
        self.addSubInterface(self.disclaimer_widget, FluentIcon.INFO, "关于", position=NavigationItemPosition.BOTTOM)

    def validate_speed(self, value: float):
        if value <= 0:
            self.speed_spin.setValue(0.01)
            self.show_warning("无效速度", "播放速度必须为正数")

    def init_recorder(self):
        self.recording = False
        self.start_time = None
        self.mouse_controller = MouseController()
        self.keyboard_controller = KeyboardController()
        self.mouse_listener = None
        self.keyboard_listener = None
        self.timer.timeout.connect(self.update_status)
        self.last_move_time = 0

    def start_global_listener(self):
        self.listener_thread = ListenerThread(self)
        shortcuts = {
            self.start_shortcut: "start_record",
            self.stop_shortcut: "stop_record",
            self.play_shortcut: "play_macro",
            self.stop_play_shortcut: "stop_play"
        }
        self.listener_thread.start()
        self.listener_thread.update_hotkeys(shortcuts)
        self.logger.info("全局键盘监听器启动")

    def handle_shortcut(self, action: str):
        if not self.is_target_process_active():
            return
        if action == "start_record" and not self.recording:
            QApplication.postEvent(self, QEvent(QEvent.User))
            self.logger.info("触发开始录制快捷键")
        elif action == "stop_record" and self.recording:
            QApplication.postEvent(self, QEvent(QEvent.User + 1))
            self.logger.info("触发停止录制快捷键")
        elif action == "play_macro" and not self.recording:
            if self.playback_mode:
                if not self.is_play_key_pressed:
                    self.is_play_key_pressed = True
                    QApplication.postEvent(self, QEvent(QEvent.User + 2))
                    self.logger.info("触发播放宏快捷键 (释放模式)")
            else:
                QApplication.postEvent(self, QEvent(QEvent.User + 2))
                self.logger.info("触发播放宏快捷键 (开关模式)")
        elif action == "stop_play":
            if self.playback_mode and self.is_play_key_pressed:
                self.is_play_key_pressed = False
                if self.playback_thread and self.playback_thread.isRunning():
                    QApplication.postEvent(self, QEvent(QEvent.User + 3))
                    self.logger.info("触发停止播放快捷键 (释放模式)")
            elif not self.playback_mode and self.playback_thread:
                QApplication.postEvent(self, QEvent(QEvent.User + 3))
                self.logger.info("触发停止播放快捷键 (开关模式)")

    def is_target_process_active(self) -> bool:
        if self.target_process == "全局":
            return True
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return False
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name().lower() == self.target_process.lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied, Exception) as e:
            self.logger.debug(f"检查进程失败: {str(e)}")
            return False

    def update_settings(self, start_seq: str, stop_seq: str, play_seq: str, stop_play_seq: str,
                        process_name: str, playback_mode: bool):
        try:
            self.start_shortcut = start_seq
            self.stop_shortcut = stop_seq
            self.play_shortcut = play_seq
            self.stop_play_shortcut = stop_play_seq
            self.target_process = process_name
            self.playback_mode = playback_mode
            shortcuts = {
                start_seq: "start_record",
                stop_seq: "stop_record",
                play_seq: "play_macro",
                stop_play_seq: "stop_play"
            }
            if self.listener_thread:
                self.listener_thread.update_hotkeys(shortcuts)
            self.start_button.setText(f"开始录制 ({self.start_shortcut})")
            self.stop_button.setText(f"停止录制 ({self.stop_shortcut})")
            self.play_button.setText(f"播放宏 ({self.play_shortcut})")
            self.stop_playback_button.setText(f"停止播放 ({self.stop_play_shortcut})")
            self.logger.info(
                f"更新设置: 开始={start_seq}, 停止={stop_seq}, 播放={play_seq}, 停止播放={stop_play_seq}, 进程={process_name}, 播放模式={'释放' if playback_mode else '开关'}")
        except Exception as e:
            self.logger.error(f"更新设置失败: {str(e)}")
            raise ValueError(f"错误: {str(e)}")

    def customEvent(self, event: QEvent):
        if event.type() == QEvent.User:
            self.start_recording()
        elif event.type() == QEvent.User + 1:
            self.stop_recording()
        elif event.type() == QEvent.User + 2:
            self.start_playback()
        elif event.type() == QEvent.User + 3:
            self.stop_playback()

    def update_status(self):
        if self.recording:
            elapsed = time.time() - self.start_time
            self.time_label.setText(f"录制时间: {elapsed:.1f}s")
            self.event_label.setText(f"事件数: {len(self.events)}")
            if elapsed > 300:
                self.stop_recording()
                self.show_info("超时", "达到最大录制时间 (5分钟)")

    def start_recording(self):
        if self.recording:
            return
        self.recording = True
        self.events = []
        self.start_time = time.time()
        self.pressed_keys.clear()
        self.timer.start(100)
        mode = self.mode_combo.currentIndex()
        try:
            if mode in [0, 2]:  # Mouse or Both
                self.mouse_listener = mouse.Listener(
                    on_move=lambda x, y: self.pynput_signals.move_signal.emit(x, y),
                    on_click=lambda x, y, button, pressed: self.pynput_signals.click_signal.emit(x, y, str(button), pressed),
                    on_scroll=lambda x, y, dx, dy: self.pynput_signals.scroll_signal.emit(x, y, dx, dy)
                )
                self.mouse_listener.start()
                self.logger.info("鼠标监听器启动")
            if mode in [1, 2]:  # Keyboard or Both
                self.keyboard_listener = KeyboardListener(
                    on_press=lambda key: self.pynput_signals.press_signal.emit(str(key)),
                    on_release=lambda key: self.pynput_signals.release_signal.emit(str(key))
                )
                self.keyboard_listener.start()
                self.logger.info("键盘监听器启动")
        except Exception as e:
            self.recording = False
            self.show_error("监听器错误", f"监听器错误: {str(e)}")
            self.logger.error(f"监听器启动失败: {str(e)}")
            return
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.save_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.play_button.setEnabled(False)
        self.load_button.setEnabled(False)
        self.manager_button.setEnabled(False)
        mode_text = ["仅鼠标", "仅键盘", "鼠标和键盘"][mode]
        self.show_success("开始录制", f"以 {mode_text} 模式开始录制")
        self.logger.info(f"开始录制，模式: {mode_text}")

    def stop_recording(self):
        if not self.recording:
            return
        self.recording = False
        self.timer.stop()
        try:
            if self.mouse_listener:
                self.mouse_listener.stop()
                self.mouse_listener = None
                self.logger.info("鼠标监听器停止")
        except Exception as e:
            self.logger.error(f"停止鼠标监听器失败: {str(e)}")
        try:
            if self.keyboard_listener:
                self.keyboard_listener.stop()
                self.keyboard_listener = None
                self.logger.info("键盘监听器停止")
        except Exception as e:
            self.logger.error(f"停止键盘监听器失败: {str(e)}")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        self.play_button.setEnabled(True)
        self.load_button.setEnabled(True)
        self.manager_button.setEnabled(True)
        self.show_info("录制停止", "请保存或取消录制")
        self.logger.info("录制停止")

    def handle_move(self, x: int, y: int):
        if self.recording and self.is_target_process_active():
            current_time = time.time()
            if current_time - self.last_move_time >= 0.05:
                self.events.append({
                    "type": "move",
                    "x": x,
                    "y": y,
                    "time": current_time - self.start_time
                })
                self.last_move_time = current_time

    def handle_click(self, x: int, y: int, button: str, pressed: bool):
        if self.recording and self.is_target_process_active():
            self.events.append({
                "type": "click",
                "x": x,
                "y": y,
                "button": button,
                "pressed": pressed,
                "time": time.time() - self.start_time
            })

    def handle_scroll(self, x: int, y: int, dx: int, dy: int):
        if self.recording and self.is_target_process_active():
            self.events.append({
                "type": "scroll",
                "x": x,
                "y": y,
                "dx": dx,
                "dy": dy,
                "time": time.time() - self.start_time
            })

    def handle_press(self, key: str):
        if self.recording and self.is_target_process_active():
            key_str = key.lower().replace("key.", "").replace("'", "")
            if key_str in ['ctrl_l', 'ctrl_r']:
                key_str = 'ctrl'
            elif key_str in ['shift_l', 'shift_r']:
                key_str = 'shift'
            elif key_str in ['alt_l', 'alt_r']:
                key_str = 'alt'
            # Avoid recording the shortcut keys themselves
            shortcut_keys = {self.start_shortcut.lower(), self.stop_shortcut.lower(),
                             self.play_shortcut.lower(), self.stop_play_shortcut.lower()}
            if key_str not in shortcut_keys and key_str not in self.pressed_keys:
                self.pressed_keys.add(key_str)
                self.events.append({
                    "type": "key_press",
                    "key": key,
                    "time": time.time() - self.start_time
                })

    def handle_release(self, key: str):
        if self.recording and self.is_target_process_active():
            key_str = key.lower().replace("key.", "").replace("'", "")
            if key_str in ['ctrl_l', 'ctrl_r']:
                key_str = 'ctrl'
            elif key_str in ['shift_l', 'shift_r']:
                key_str = 'shift'
            elif key_str in ['alt_l', 'alt_r']:
                key_str = 'alt'
            shortcut_keys = {self.start_shortcut.lower(), self.stop_shortcut.lower(),
                             self.play_shortcut.lower(), self.stop_play_shortcut.lower()}
            if key_str not in shortcut_keys:
                self.pressed_keys.discard(key_str)
                self.events.append({
                    "type": "key_release",
                    "key": key,
                    "time": time.time() - self.start_time
                })

    def save_recording(self):
        if not self.events:
            self.show_warning("无内容", "未录制任何操作")
            return
        file_name, _ = QFileDialog.getSaveFileName(self, "保存录制", MACRO_DIR, "JSON Files (*.json)")
        if file_name:
            try:
                with open(file_name, 'w', encoding='utf-8') as f:
                    json.dump(self.events, f, indent=4, ensure_ascii=False)
                self.show_success("成功", f"宏已保存至 {file_name}")
                self.logger.info(f"宏保存至 {file_name}")
                self.reset_ui()
            except Exception as e:
                self.show_error("保存失败", f"保存失败: {str(e)}")
                self.logger.error(f"保存宏失败: {str(e)}")

    def cancel_recording(self):
        self.events = []
        self.reset_ui()
        self.show_info("已取消", "录制已丢弃")
        self.logger.info("录制取消")

    def reset_ui(self):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.cancel_button.setEnabled(False)
        self.play_button.setEnabled(True)
        self.load_button.setEnabled(True)
        self.manager_button.setEnabled(True)
        self.time_label.setText("录制时间: 0.0s")
        self.event_label.setText("事件数: 0")

    def load_macro(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "加载宏", MACRO_DIR, "JSON Files (*.json)")
        if file_name:
            try:
                with open(file_name, 'r', encoding='utf-8') as f:
                    self.events = json.load(f)
                self.show_success("成功", f"已从 {file_name} 加载宏")
                self.logger.info(f"从 {file_name} 加载宏")
                self.editor_widget.load_events()
            except Exception as e:
                self.show_error("加载失败", f"加载失败: {str(e)}")
                self.logger.error(f"加载宏失败: {str(e)}")

    def start_playback(self):
        if not self.events:
            self.show_warning("无宏", "未加载任何宏")
            return
        if self.playback_thread and self.playback_thread.isRunning():
            self.show_warning("正在播放", "播放正在进行")
            return
        self.playback_thread = PlaybackThread(self)
        self.playback_thread.finished.connect(self.on_playback_finished)
        self.playback_thread.interrupted.connect(self.on_playback_interrupted)
        self.playback_thread.error.connect(self.on_playback_error)
        self.playback_thread.start()
        self.play_button.setEnabled(False)
        self.stop_playback_button.setEnabled(not self.playback_mode)
        self.manager_button.setEnabled(False)
        self.show_info("开始播放", f"正在以 {self.speed_spin.value()}x 速度播放宏")
        self.logger.info(
            f"开始播放，重复: {self.play_count_spin.value()} 次，速度: {self.speed_spin.value()}x，模式: {'释放' if self.playback_mode else '开关'}")

    def stop_playback(self):
        if self.playback_thread and self.playback_thread.isRunning():
            self.playback_thread.stop()
            self.playback_thread.wait()
            self.on_playback_interrupted()

    def on_playback_finished(self):
        self.play_button.setEnabled(True)
        self.stop_playback_button.setEnabled(False)
        self.manager_button.setEnabled(True)
        self.playback_thread = None
        self.is_play_key_pressed = False
        self.show_success("播放完成", "播放成功完成")
        self.logger.info("播放完成")

    def on_playback_interrupted(self):
        self.play_button.setEnabled(True)
        self.stop_playback_button.setEnabled(False)
        self.manager_button.setEnabled(True)
        self.playback_thread = None
        self.is_play_key_pressed = False
        self.show_warning("播放中断", "用户中断了播放")
        self.logger.info("播放中断")

    def on_playback_error(self, error_msg: str):
        self.play_button.setEnabled(True)
        self.stop_playback_button.setEnabled(False)
        self.manager_button.setEnabled(True)
        self.playback_thread = None
        self.is_play_key_pressed = False
        self.show_error("播放错误", f"播放错误: {error_msg}")
        self.logger.error(f"播放错误: {error_msg}")

    def show_macro_manager(self):
        self.stackedWidget.setCurrentWidget(self.manager_widget)
        self.manager_widget.load_macros()

    def show_success(self, title: str, content: str):
        InfoBar.success(title, content, position=InfoBarPosition.TOP_RIGHT, duration=3000, parent=self)

    def show_info(self, title: str, content: str):
        InfoBar.info(title, content, position=InfoBarPosition.TOP_RIGHT, duration=3000, parent=self)

    def show_warning(self, title: str, content: str):
        InfoBar.warning(title, content, position=InfoBarPosition.TOP_RIGHT, duration=3000, parent=self)

    def show_error(self, title: str, content: str):
        InfoBar.error(title, content, position=InfoBarPosition.TOP_RIGHT, duration=3000, parent=self)

    def closeEvent(self, event: QEvent):
        if self.recording:
            reply = QMessageBox.question(
                self, "确认退出",
                "正在录制，是否退出？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.stop_recording()
            else:
                event.ignore()
                return
        if self.playback_thread and self.playback_thread.isRunning():
            reply = QMessageBox.question(
                self, "确认退出",
                "正在播放，是否退出？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.stop_playback()
            else:
                event.ignore()
                return
        try:
            if self.listener_thread:
                self.listener_thread.stop()
                self.logger.info("全局键盘监听器停止")
        except Exception as e:
            self.logger.error(f"停止全局监听器失败: {str(e)}")
        event.accept()
        self.logger.info("应用关闭")


if __name__ == "__main__":
    app = QApplication([])
    window = MacroController()
    window.show()
    app.exec_()