# MouseDash / 老鼠大师

> Windows 宏录制与播放工具 — 录制鼠标键盘操作，一键重放，专治重复劳动。
> A Windows macro recorder & player — capture mouse and keyboard input, replay on demand.

面向游戏玩家、程序员、办公摸鱼人士。录制你的操作，保存成宏，想放几遍放几遍。带反检测、进程锁定、全局快捷键，该有的都有。

Built for gamers, developers, and anyone stuck doing the same clicks a thousand times. Record your actions, save them, replay at will. Comes with anti-detection, process targeting, global hotkeys — the works.

## Features / 功能

- **宏录制 / Recording** — 录制鼠标移动、点击、滚轮、键盘按键。可选仅鼠标、仅键盘或全部录制。
- **宏播放 / Playback** — 可调速度与重复次数。支持开关模式（按一次开始/停止）和释放模式（按住播放、松开停止）。
- **宏管理 / Management** — JSON 格式保存，支持加载、编辑、删除。
- **宏编辑 / Editing** — GUI 中逐事件查看修改，可增删单个动作。
- **自定义快捷键 / Hotkeys** — 录制与播放均可配置快捷键（默认 F1/F2/F4/F5）。
- **进程锁定 / Process targeting** — 全局监听，或仅在指定进程激活时响应（基于 psutil）。
- **反后台检测 / Anti-detection** — 事件间插入自适应随机延迟，模拟人类操作节奏。
- **Windows 通知 / Notifications** — 系统级通知提示录制/播放状态（基于 winotify）。
- **主题切换 / Theming** — 系统 / 亮色 / 暗色，基于 qfluentwidgets。

## Tech Stack / 技术栈

| Component | Library |
|-----------|---------|
| Language | Python 3.9+ |
| GUI | PyQt5 >= 5.15.11, PyQt-Fluent-Widgets >= 1.11.2 |
| Input | pynput >= 1.8.2 |
| Process | psutil >= 7.2.2 |
| Notifications | winotify >= 1.1.0 |
| Windows API | pywin32 >= 312 |

完整依赖见 [`requirements.txt`](requirements.txt)。

## Project Structure / 项目结构

```
MouseDash/
├── main.py            # 单文件主程序：GUI、录制、播放、设置、监听线程
├── requirements.txt   # Python 依赖
├── config/
│   └── config.json    # 默认配置（主题、快捷键、播放模式等）
└── LICENSE            # MIT
```

宏保存在运行时创建的 `./macros/` 目录，配置在 `./config/`，首次启动自动生成。

## Getting Started / 快速开始

### Prerequisites / 环境要求

- Windows
- Python 3.9+（推荐 3.10-3.12）

### Install / 安装

```bash
git clone https://github.com/dwgx/MouseDash.git
cd MouseDash

python -m venv .venv
.\.venv\Scripts\activate

pip install --upgrade pip
pip install -r requirements.txt
```

### Run / 运行

```bash
python main.py
```

快捷键或全局监听不生效时，以管理员权限运行。

## Usage / 使用方法

1. **录制** — 选录制模式，按 F1 开始录制，操作完按 F2 停止，保存为 JSON。
2. **播放** — 加载宏，设重复次数和速度，按 F4 播放，F5 停止。
3. **管理** — 「管理宏」页面查看已有宏，可使用、编辑或删除。
4. **编辑** — 「编辑宏」页面逐事件修改，支持增删事件。
5. **设置** — 调整快捷键、监控进程、播放模式、反检测、通知、主题。

单次录制上限约 300 秒（5 分钟），超时自动停止。

## Configuration / 配置

默认配置 [`config/config.json`](config/config.json)，可通过设置界面修改：

| Key | Description | Default |
|-----|-------------|---------|
| `theme` | 主题 system/light/dark | `dark` |
| `speed` | 播放速度 | `200.0` |
| `start_shortcut` | 开始录制 | `F1` |
| `stop_shortcut` | 停止录制 | `F2` |
| `play_shortcut` | 开始播放 | `F4` |
| `stop_play_shortcut` | 停止播放 | `F5` |
| `target_process` | 监控进程 | `全局` |
| `playback_mode` | false=开关, true=释放 | `false` |
| `prevent_background_detection` | 反检测随机延迟 | `true` |
| `windows_notification_enabled` | Windows 通知 | `true` |

## Status / 状态

v1.0.0 已发布（2025-04-21）。单文件个人项目，功能完整。

Released v1.0.0. Single-file personal project, fully functional.

## License / 许可证

[MIT](LICENSE) - Copyright (c) 2026 dwgx
