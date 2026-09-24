#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""实验室综合管理系统 v2.1 · 主程序

- UI + 串口监听（扫码枪 / 人脸识别模块）
- 数据读写见 db.py（openpyxl 直接操作，带内存缓存与自动备份）
"""

import glob
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog, ttk

try:
    import serial
    from serial.tools import list_ports
except ModuleNotFoundError as exc:
    print(f"缺少 Python 依赖: {exc.name}")
    print("请先在本目录执行: python3 -m pip install -r requirements.txt")
    sys.exit(1)

try:
    import db
except ModuleNotFoundError as exc:
    print(f"缺少数据层模块: {exc.name}（db.py 需要与 main.py 放在同一目录）")
    sys.exit(1)

try:
    import fcntl  # Linux 文件锁（单实例保护）
except ImportError:
    fcntl = None

try:
    import ttkbootstrap as tb   # 现代化界面主题（可选依赖）
    USE_TB = True
except ModuleNotFoundError:
    tb = None
    USE_TB = False

# ================= 配置区 =================
BASE_DIR = Path(__file__).resolve().parent

# Ubuntu 下串口通常是 /dev/ttyUSB*、/dev/ttyACM* 或 /dev/serial/by-id/*。
# 可用环境变量预设端口:
# SCANNER_PORT=/dev/ttyUSB0 FACE_PORT=/dev/ttyUSB1 python3 main.py
DEFAULT_SCANNER_PORT = os.environ.get("SCANNER_PORT", "")
DEFAULT_FACE_PORT = os.environ.get("FACE_PORT", "")

BAUD_RATE = 115200
SCAN_BAUD = 9600

WAREHOUSE_EXCEL = BASE_DIR / 'warehouse_db.xlsx'   # 仓库数据库
CHECKIN_EXCEL = BASE_DIR / '实验室签到表.xlsx'      # 签到记录表
BACKUP_WAREHOUSE = True   # 每次写库前自动备份为 warehouse_db.bak
FACE_DEDUP_SECONDS = 15   # 同一人 N 秒内重复刷脸不重复签到（0 = 关闭）
ADMIN_PASSWORD = "admin888"   # 管理员面板密码（请自行修改）
RECONNECT_INTERVAL = 2.0      # 串口断线后自动重连间隔（秒）
LOG_TO_FILE = True            # 运行日志写入 运行日志.log
MAX_LOG_BYTES = 1_000_000     # 日志文件超过 1MB 自动转存 运行日志.old.log
CONFIG_FILE = BASE_DIR / "config.json"   # 界面配置（主题等）
# ==========================================

# 初始化主窗口
def _load_app_config():
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

APP_CONFIG = _load_app_config()
THEMES = ["cosmo", "litera", "minty", "pulse", "flatly", "darkly", "solar", "cyborg", "superhero"]

def current_theme():
    theme = APP_CONFIG.get("theme", "cosmo")
    return theme if theme in THEMES else "cosmo"

if USE_TB:
    root = tb.Window(themename=current_theme())
else:
    root = tk.Tk()
root.title("实验室综合管理系统")
root.geometry("660x860")
root.minsize(660, 760)
root.attributes("-topmost", True)

# 高 DPI 显示器上限制 Tk 缩放系数，避免控件被放大到溢出窗口
try:
    _scaling = float(root.tk.call("tk", "scaling"))
    if _scaling > 1.6:
        root.tk.call("tk", "scaling", 1.5)
except Exception:
    pass

# 固定配色（卡片式布局）
BG = "#f4f6f9"        # 页面背景
CARD = "#ffffff"      # 卡片背景
BORDER = "#e3e8ef"    # 卡片描边
PRIMARY = "#2f6feb"   # 主色
TEXT = "#1f2937"      # 正文
MUTED = "#6b7280"     # 次要文字
GREEN = "#16a34a"
ORANGE = "#ea580c"
RED = "#dc2626"
BLUE = "#2563eb"
LOG_BG = "#1c2530"    # 日志区深色底
LOG_FG = "#c9d6e3"
_MODE_COLORS = {"IN": "green", "OUT": "blue", "STANDBY": "red"}

_FALLBACK_BG = {"success": "#dcfce7", "danger": "#fee2e2", "warning": "#fef3c7",
                "info": "#dbeafe", "primary": "#dbeafe", "secondary": "#eef2f7"}

import tkinter.font as tkfont
FAM = "Noto Sans CJK SC" if "Noto Sans CJK SC" in tkfont.families(root) else "Sans"
F_TITLE = (FAM, 17, "bold")
F_SECTION = (FAM, 12, "bold")
F_BODY = (FAM, 10)
F_BODY_BOLD = (FAM, 10, "bold")
F_SMALL = (FAM, 9)

if USE_TB:
    try:
        root.style.configure("Treeview", rowheight=26)
    except Exception:
        pass
else:
    _style = ttk.Style()
    try:
        _style.theme_use("clam")
    except Exception:
        pass
    _style.configure("Treeview", rowheight=26)


def mk_button(parent, text, command, style="secondary", width=None, height=None, state="normal"):
    """统一按钮工厂：有 ttkbootstrap 用彩色主题按钮，否则用浅色底按钮。"""
    if USE_TB:
        kwargs = {"bootstyle": style or "secondary"}
    else:
        bg = _FALLBACK_BG.get(style, "#eef2f7")
        kwargs = {"bg": bg, "activebackground": bg, "relief": "flat", "font": F_BODY}
        if height:
            kwargs["height"] = height
    if state != "normal":
        kwargs["state"] = state
    if width:
        kwargs["width"] = width
    cls = tb.Button if USE_TB else tk.Button
    return cls(parent, text=text, command=command, **kwargs)


def mk_check(parent, text, var, command):
    """统一开关工厂：有 ttkbootstrap 用圆角开关样式。"""
    if USE_TB:
        return tb.Checkbutton(parent, text=text, variable=var, command=command, bootstyle="round-toggle")
    return tk.Checkbutton(parent, text=text, variable=var, command=command,
                          bg=CARD, activebackground=CARD, font=F_BODY, anchor="w")

current_user = "未登录"
current_role = db.ROLE_MEMBER
current_mode = "STANDBY"
ser_scan = None
ser_face = None
serial_running = False
warehouse_db = None
checkin_log = None
inventory_win = None
inventory_text_widget = None
log_text = None
last_face_time = {}          # 姓名 -> 最近一次签到时间戳（去重用）
face_buffer = ""
face_buffer_since = 0.0
scan_buffer = b""
scan_buffer_since = 0.0
face_log_cache = {}          # 异常帧日志限流
_poll_generation = 0         # 串口轮询代际，防止 stop→start 后残留双循环
_flash_job = None
scan_target = None           # 启动时选定的扫码枪端口（断线自动重连用）
face_target = None
_last_reconnect_check = 0.0
admin_win = None             # 管理员面板窗口（单例）


def pad_display(text, width):
    """按显示宽度补空格：中日韩全角字符按 2 列计。"""
    s = str(text)
    w = sum(2 if ord(ch) > 0x2E7F else 1 for ch in s)
    return s + " " * max(0, width - w)


def log(msg):
    """带时间戳写入窗口日志区、终端与日志文件。"""
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    if log_text is not None:
        log_text.config(state="normal")
        log_text.insert("end", line + "\n")
        log_text.see("end")
        log_text.config(state="disabled")
    if LOG_TO_FILE:
        try:
            log_path = BASE_DIR / "运行日志.log"
            if log_path.exists() and log_path.stat().st_size > MAX_LOG_BYTES:
                try:
                    os.replace(log_path, BASE_DIR / "运行日志.old.log")
                except OSError:
                    pass
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass   # 日志文件写失败不影响主流程


def _on_tk_error(exc_type, exc_value, exc_tb):
    """把 Tk 回调里的未捕获异常记入日志（便于排查）。"""
    log(f"[异常] {exc_type.__name__}: {exc_value}")


root.report_callback_exception = _on_tk_error


def flash_status(msg, color="#2e7d32"):
    """状态条显示非模态提示，5 秒后自动消失，不阻塞扫码。"""
    global _flash_job
    if _flash_job is not None:
        root.after_cancel(_flash_job)
    status_hint.config(text=msg, fg=color)
    _flash_job = root.after(5000, lambda: status_hint.config(text="", fg="gray"))


# ================= 串口管理 =================
def discover_ports():
    """发现 Ubuntu/Linux 下可用的串口设备。"""
    ports = []
    try:
        ports.extend(port.device for port in list_ports.comports())
    except Exception:
        pass
    patterns = [
        "/dev/serial/by-id/*",
        "/dev/ttyUSB*",
        "/dev/ttyACM*",
        "/dev/ttyS*",
        "/dev/rfcomm*",
    ]
    for pattern in patterns:
        ports.extend(glob.glob(pattern))
    return sorted(dict.fromkeys(ports))


def set_combo_value(combo, var, preferred, ports):
    if preferred and preferred in ports:
        var.set(preferred)
        return
    if var.get() and var.get() in ports:
        return
    var.set(ports[0] if ports else "")


def refresh_ports():
    ports = discover_ports()
    combo_values = [""] + ports
    scanner_combo["values"] = combo_values
    face_combo["values"] = combo_values

    set_combo_value(scanner_combo, scanner_port_var, DEFAULT_SCANNER_PORT, ports)
    preferred_face = DEFAULT_FACE_PORT
    if not preferred_face and len(ports) > 1:
        preferred_face = ports[1]
    if preferred_face:
        set_combo_value(face_combo, face_port_var, preferred_face, ports)
    elif face_port_var.get() not in ports:
        face_port_var.set("")

    port_status = "未发现串口设备" if not ports else f"发现 {len(ports)} 个串口设备"
    serial_status.config(text=f"串口状态: {port_status}")
    log(f"刷新串口：{port_status}" + (f"（{'、'.join(ports)}）" if ports else ""))


def close_serial():
    global ser_scan, ser_face
    for ser in (ser_scan, ser_face):
        if ser and ser.is_open:
            try:
                ser.close()
            except Exception:
                pass
    ser_scan = None
    ser_face = None


def start_serial():
    global ser_scan, ser_face, serial_running, _poll_generation, scan_target, face_target

    scan_port = scanner_port_var.get().strip()
    face_port = face_port_var.get().strip()

    if not scan_port and not face_port:
        messagebox.showwarning("串口未选择", "请至少选择扫码枪或人脸识别模块的串口。")
        return
    if scan_port and face_port and scan_port == face_port:
        messagebox.showwarning("串口冲突", "扫码枪和人脸识别模块不能使用同一个串口。")
        return

    close_serial()
    opened = []
    errors = []

    if scan_port:
        try:
            ser_scan = serial.Serial(scan_port, SCAN_BAUD, timeout=0.1)
            opened.append(f"扫码枪: {scan_port}")
        except serial.SerialException as exc:
            ser_scan = None
            errors.append(f"扫码枪 {scan_port}: {exc}")

    if face_port:
        try:
            ser_face = serial.Serial(face_port, BAUD_RATE, timeout=0.1)
            opened.append(f"人脸模块: {face_port}")
        except serial.SerialException as exc:
            ser_face = None
            errors.append(f"人脸模块 {face_port}: {exc}")

    if not opened:
        serial_status.config(text="串口状态: 打开失败", fg=RED)
        log("串口启动失败：" + "；".join(errors))
        messagebox.showerror(
            "串口打开失败",
            "没有任何串口成功打开。\n\n"
            + "\n".join(errors)
            + "\n\nUbuntu 上请确认设备路径正确，并且当前用户在 dialout 组。",
        )
        return

    scan_target = scan_port or None
    face_target = face_port or None
    _poll_generation += 1   # 使旧轮询循环失效，防止双循环
    serial_running = True
    start_btn.config(state=tk.DISABLED)
    stop_btn.config(state=tk.NORMAL)
    scanner_combo.config(state=tk.DISABLED)
    face_combo.config(state=tk.DISABLED)
    update_serial_status()
    log("串口已启动：" + "；".join(opened) + "；断线自动重连已开启")

    if errors:
        log("[警告] 部分串口打开失败：" + "；".join(errors))
        messagebox.showwarning("部分串口打开失败", "\n".join(errors))

    poll_serial()


def stop_serial():
    global serial_running, _poll_generation, face_buffer, face_buffer_since, scan_buffer, scan_buffer_since, scan_target, face_target
    _poll_generation += 1   # 使旧轮询循环失效
    serial_running = False
    close_serial()
    scan_target = None
    face_target = None
    face_buffer = ""
    face_buffer_since = 0.0
    scan_buffer = b""
    scan_buffer_since = 0.0
    start_btn.config(state=tk.NORMAL)
    stop_btn.config(state=tk.DISABLED)
    scanner_combo.config(state="readonly")
    face_combo.config(state="readonly")
    serial_status.config(text="串口状态: 已停止", fg=MUTED)
    log("串口已停止")


def on_close():
    stop_serial()
    root.destroy()


def update_serial_status():
    """按各端口实时连接状态刷新状态栏（全部正常=绿色，有断线=橙色）。"""
    if not serial_running:
        return
    parts = []
    if scan_target:
        state = "已连接" if (ser_scan and ser_scan.is_open) else "断线重连中"
        parts.append(f"扫码枪 {state}")
    if face_target:
        state = "已连接" if (ser_face and ser_face.is_open) else "断线重连中"
        parts.append(f"人脸模块 {state}")
    all_ok = all(
        ser is not None and ser.is_open
        for ser, target in ((ser_scan, scan_target), (ser_face, face_target))
        if target
    )
    serial_status.config(
        text="串口状态: 监听中 ｜ " + " ｜ ".join(parts),
        fg=GREEN if all_ok else ORANGE,
    )


def _drop_face():
    global ser_face, face_buffer, face_buffer_since
    if ser_face:
        try:
            ser_face.close()
        except Exception:
            pass
    ser_face = None
    face_buffer = ""
    face_buffer_since = 0.0


def _drop_scan():
    global ser_scan, scan_buffer, scan_buffer_since
    if ser_scan:
        try:
            ser_scan.close()
        except Exception:
            pass
    ser_scan = None
    scan_buffer = b""
    scan_buffer_since = 0.0


def maybe_reconnect():
    """周期性尝试重连断线的串口（USB 拔插后自动恢复）。"""
    global ser_scan, ser_face, _last_reconnect_check
    global scan_buffer, scan_buffer_since, face_buffer, face_buffer_since
    now = time.monotonic()
    if now - _last_reconnect_check < RECONNECT_INTERVAL:
        return
    _last_reconnect_check = now
    changed = False
    if scan_target and (ser_scan is None or not ser_scan.is_open):
        if os.path.exists(scan_target):
            try:
                ser_scan = serial.Serial(scan_target, SCAN_BAUD, timeout=0.1)
                scan_buffer = b""
                scan_buffer_since = 0.0
                log(f"[串口] 扫码枪已重新连接：{scan_target}")
                changed = True
            except serial.SerialException:
                pass
    if face_target and (ser_face is None or not ser_face.is_open):
        if os.path.exists(face_target):
            try:
                ser_face = serial.Serial(face_target, BAUD_RATE, timeout=0.1)
                face_buffer = ""
                face_buffer_since = 0.0
                log(f"[串口] 人脸模块已重新连接：{face_target}")
                changed = True
            except serial.SerialException:
                pass
    if changed:
        update_serial_status()


_lock_handle = None


def acquire_single_instance():
    """防止重复启动多个实例同时写同一个数据库。"""
    global _lock_handle
    if fcntl is None:
        return True
    try:
        _lock_handle = open(BASE_DIR / ".app.lock", "w")
        fcntl.flock(_lock_handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        _lock_handle.write(str(os.getpid()))
        _lock_handle.flush()
        return True
    except OSError:
        return False


# ================= 核心功能逻辑 =================
def refresh_inventory_text():
    """把当前库存刷新到库存窗口的表格里。"""
    if inventory_win is None or not inventory_win.winfo_exists():
        return
    try:
        items = warehouse_db.items
        # 已出库的排前面，再按柜号升序（与旧版显示习惯一致）
        items_sorted = sorted(
            items,
            key=lambda d: (d.get(db.COL_STATUS) == db.STATUS_IN, d.get(db.COL_CABINET, "")),
        )
        n_in = sum(1 for d in items if d.get(db.COL_STATUS) == db.STATUS_IN)
        inventory_stats.config(
            text=f"共 {len(items)} 件 · 在库 {n_in} · 已出库 {len(items) - n_in} · 操作人：{current_user}")
        inventory_tree.delete(*inventory_tree.get_children())
        for d in items_sorted:
            status = d.get(db.COL_STATUS, "")
            tag = "in" if status == db.STATUS_IN else "out"
            inventory_tree.insert(
                "", "end",
                values=(status, d.get(db.COL_CABINET, ""), d.get(db.COL_NAME, "")),
                tags=(tag,),
            )
    except Exception as exc:
        messagebox.showerror("清单加载失败", str(exc))


def show_inventory():
    """显示/刷新库存清单窗口（单例：重复调用只刷新，不叠加弹窗）。"""
    global inventory_win, inventory_tree, inventory_stats, inventory_tree_wrap, inventory_btn_row
    if inventory_win is not None and inventory_win.winfo_exists():
        inventory_win.deiconify()
        inventory_win.lift()
        refresh_inventory_text()
        return
    inventory_win = tk.Toplevel(root)
    inventory_win.title("实时库存")
    inventory_win.geometry("560x440")
    inventory_win.attributes("-topmost", True)
    inventory_stats = tk.Label(inventory_win, text="", font=F_BODY_BOLD)
    inventory_stats.pack(padx=14, pady=(14, 6), anchor="w")
    inventory_tree_wrap = tk.Frame(inventory_win)
    inventory_tree_wrap.pack(fill="both", expand=True, padx=14)
    inventory_tree = ttk.Treeview(inventory_tree_wrap, columns=("status", "cabinet", "name"),
                                  show="headings", height=14)
    inventory_tree.heading("status", text="状态")
    inventory_tree.heading("cabinet", text="柜号")
    inventory_tree.heading("name", text="物品名称")
    inventory_tree.column("status", width=90, anchor="center")
    inventory_tree.column("cabinet", width=70, anchor="center")
    inventory_tree.column("name", width=280, anchor="w")
    tree_scroll = tk.Scrollbar(inventory_tree_wrap, command=inventory_tree.yview)
    inventory_tree.configure(yscrollcommand=tree_scroll.set)
    tree_scroll.pack(side=tk.RIGHT, fill="y")
    inventory_tree.pack(side=tk.LEFT, fill="both", expand=True)
    inventory_btn_row = tk.Frame(inventory_win)
    inventory_btn_row.pack(fill="x", padx=14, pady=12)
    mk_button(inventory_btn_row, "刷新", refresh_inventory_text, "secondary", width=10).pack(side=tk.RIGHT)
    inventory_win.protocol("WM_DELETE_WINDOW", inventory_win.destroy)
    apply_inventory_colors()
    refresh_inventory_text()


def set_mode_in():
    global current_mode
    current_mode = "IN"
    mode_display.config(text="当前模式: 入库模式", fg="green")
    log(f"模式切换：入库 (操作人: {current_user})")


def set_mode_out():
    global current_mode
    current_mode = "OUT"
    mode_display.config(text="当前模式: 出库模式", fg="blue")
    log(f"模式切换：出库 (操作人: {current_user})")


def handle_face_line(line):
    """处理人脸模块的一帧数据，如「张三:SUCCESS」。"""
    global current_user, current_role
    line = line.strip()
    if not line:
        return
    if ":SUCCESS" not in line:
        # 失败/非识别帧：相同内容 5 秒内只记一次日志，避免刷屏
        now = time.time()
        if len(face_log_cache) > 200:
            face_log_cache.clear()
        if face_log_cache.get(line, 0) + 5 < now:
            face_log_cache[line] = now
            log(f"[人脸] 未识别帧：{line!r}")
        return
    name = line.split(":", 1)[0].strip()
    if not name:
        log(f"[人脸] 帧缺少姓名：{line!r}")
        return
    current_user = name
    current_role = warehouse_db.find_role_by_name(name)
    update_admin_btn()
    user_display.config(text=f"当前用户: {current_user}")

    now = time.time()
    if FACE_DEDUP_SECONDS > 0 and last_face_time.get(name, 0) + FACE_DEDUP_SECONDS > now:
        log(f"[人脸] {name} 在 {FACE_DEDUP_SECONDS} 秒内重复签到，已跳过")
        flash_status(f"{name} 已签到，{FACE_DEDUP_SECONDS} 秒内不重复记录", color="#ef6c00")
        return
    last_face_time[name] = now

    try:
        checkin_log.append(name)
    except PermissionError as exc:
        messagebox.showerror("保存失败", str(exc))
        return
    except Exception as exc:
        messagebox.showerror("保存失败", f"写入签到表失败：{exc}")
        return
    log(f"[人脸签到成功] {name}（{current_role}）")
    flash_status(f"签到成功：{name}")
    show_inventory()


def process_barcode(barcode):
    """处理扫码枪输入：人员码登录 / 物品码出入库。"""
    global current_user, current_role
    barcode = str(barcode).strip()
    if not barcode:
        return

    try:
        user = warehouse_db.find_user(barcode)
        if user:
            if user != current_user:
                current_user = user
                current_role = warehouse_db.user_role(barcode)
                update_admin_btn()
                user_display.config(text=f"当前用户: {current_user}")
            log(f"[人员码] 登录：{user}（{current_role}）")
            flash_status(f"欢迎，{user}！")
            show_inventory()
            return
        item = warehouse_db.find_item(barcode)
    except Exception as exc:
        messagebox.showerror("读取失败", f"读取仓库数据库失败：{exc}")
        return

    if item is None:
        log(f"[扫码] 未匹配到条码：{barcode}")
        flash_status(f"未匹配到条码：{barcode}", color="#c62828")
        return

    if current_user == "未登录":
        log("[扫码] 未登录，物品码已忽略（请先刷脸或扫人员码）")
        flash_status("请先刷脸或扫人员码登录！", color="#c62828")
        return
    if current_mode == "STANDBY":
        log("[扫码] 未选择出入库模式，物品码已忽略")
        flash_status("请先点击选择【入库】或【出库】！", color="#ef6c00")
        return

    status = db.STATUS_IN if current_mode == "IN" else db.STATUS_OUT
    action = "入库" if current_mode == "IN" else "出库"

    if item.get(db.COL_STATUS) == status:
        log(f"[扫码] {item.get(db.COL_NAME)} 已是「{status}」，重复扫描已忽略")
        flash_status(f"{item.get(db.COL_NAME)} 已{action}，无需重复操作", color="#ef6c00")
        return

    try:
        updated = warehouse_db.update_item(barcode, status, current_user)
    except PermissionError as exc:
        messagebox.showerror("保存失败", str(exc))
        return
    except Exception as exc:
        messagebox.showerror("保存失败", f"写入仓库数据库失败：{exc}")
        return

    log(f"[{action}成功] {updated[db.COL_NAME]}（柜号 {updated[db.COL_CABINET]}），操作人：{current_user}")
    flash_status(f"{action}成功：{updated[db.COL_NAME]}（柜号 {updated[db.COL_CABINET]}）")


def poll_serial():
    """串口轮询：人脸数据按帧缓冲处理，扫码数据按行处理；单端口断线不影响另一个。"""
    global face_buffer, face_buffer_since, scan_buffer, scan_buffer_since
    if not serial_running:
        return
    generation = _poll_generation

    maybe_reconnect()

    if ser_face is not None:
        try:
            if ser_face.is_open and ser_face.in_waiting > 0:
                face_buffer += ser_face.read_all().decode("utf-8", errors="ignore")
                if face_buffer_since == 0.0:
                    face_buffer_since = time.monotonic()
        except (serial.SerialException, OSError) as exc:
            log(f"[串口] 人脸模块断线：{exc}")
            _drop_face()
            update_serial_status()
        if ser_face is not None and face_buffer:
            # 以换行结尾，或缓冲停留超过 200ms，都视为完整帧
            complete = face_buffer.endswith(("\n", "\r")) or (time.monotonic() - face_buffer_since) > 0.2
            if complete:
                for line in face_buffer.splitlines():
                    handle_face_line(line)
                face_buffer = ""
                face_buffer_since = 0.0

    if ser_scan is not None:
        try:
            if ser_scan.is_open and ser_scan.in_waiting > 0:
                scan_buffer += ser_scan.read_all()
                if scan_buffer_since == 0.0:
                    scan_buffer_since = time.monotonic()
        except (serial.SerialException, OSError) as exc:
            log(f"[串口] 扫码枪断线：{exc}")
            _drop_scan()
            update_serial_status()
        if ser_scan is not None and scan_buffer:
            # 以换行/回车结尾，或停顿超过 200ms，视为一条完整扫码
            complete = (b"\n" in scan_buffer or b"\r" in scan_buffer
                        or (time.monotonic() - scan_buffer_since) > 0.2)
            if complete:
                text = scan_buffer.decode("utf-8", errors="ignore")
                for line in re.split(r"[\r\n]+", text):
                    line = line.strip()
                    if line:
                        process_barcode(line)
                scan_buffer = b""
                scan_buffer_since = 0.0

    if serial_running and generation == _poll_generation:
        root.after(10, poll_serial)


def submit_manual():
    """手动输入条码（无扫码枪时测试 / 备用录入）。"""
    code = manual_var.get().strip()
    if not code:
        return
    manual_var.set("")
    log(f"[手动] 输入条码：{code}")
    process_barcode(code)


AUTOSTART_DIR = Path.home() / ".config" / "autostart"
AUTOSTART_FILE = AUTOSTART_DIR / "实验室综合管理系统.desktop"


def _autostart_content():
    return f"""[Desktop Entry]
Type=Application
Name=实验室综合管理系统
Comment=实验室综合管理系统开机自启
Exec=bash -lc 'cd "{BASE_DIR}" && "{BASE_DIR}/.venv/bin/python" main.py --minimized'
Terminal=false
X-GNOME-Autostart-enabled=true
"""


def toggle_autostart():
    """「开机自动启动」开关：创建/删除自启文件。"""
    if autostart_var.get():
        try:
            AUTOSTART_DIR.mkdir(parents=True, exist_ok=True)
            AUTOSTART_FILE.write_text(_autostart_content(), encoding="utf-8")
            log("已开启开机自启（登录后自动最小化启动）")
        except OSError as exc:
            autostart_var.set(False)
            messagebox.showerror("设置失败", f"无法写入自启文件：{exc}")
    else:
        try:
            if AUTOSTART_FILE.exists():
                AUTOSTART_FILE.unlink()
            log("已关闭开机自启")
        except OSError as exc:
            messagebox.showerror("设置失败", str(exc))


def update_admin_btn():
    """管理员登录后才启用「管理」按钮。"""
    admin_btn.config(state=tk.NORMAL if current_role == db.ROLE_ADMIN else tk.DISABLED)


def logout_current():
    """当前用户被删除或降级时强制登出。"""
    global current_user, current_role
    current_user = "未登录"
    current_role = db.ROLE_MEMBER
    user_display.config(text=f"当前用户: {current_user}")
    update_admin_btn()
    log("当前用户已被移除，已自动登出")


# ---- 管理员面板 ----
def open_admin():
    """管理员面板入口：角色校验 + 密码验证。"""
    global admin_win
    if current_role != db.ROLE_ADMIN:
        messagebox.showwarning("权限不足", "请先以管理员身份刷脸或扫人员码登录。")
        return
    password = simpledialog.askstring("管理员验证", "请输入管理员密码：", show="*", parent=root)
    if password is None:
        return
    if password != ADMIN_PASSWORD:
        messagebox.showerror("密码错误", "管理员密码不正确。")
        return
    if admin_win is not None and admin_win.winfo_exists():
        admin_win.deiconify()
        admin_win.lift()
        refresh_admin_tables()
        return
    build_admin_win()
    refresh_admin_tables()


def build_admin_win():
    global admin_win, user_tree, item_tree
    global user_name_var, user_code_var, role_var, item_code_var, item_name_var, item_cabinet_var
    admin_win = tk.Toplevel(root)
    admin_win.title("管理员面板")
    admin_win.geometry("600x560")
    admin_win.attributes("-topmost", True)
    nb = ttk.Notebook(admin_win)
    nb.pack(fill="both", expand=True, padx=8, pady=8)

    # 人员管理
    user_tab = ttk.Frame(nb)
    nb.add(user_tab, text=" 人员管理 ")
    user_tree = ttk.Treeview(user_tab, columns=("name", "barcode", "role"), show="headings", height=12)
    user_tree.heading("name", text="姓名")
    user_tree.heading("barcode", text="条码")
    user_tree.heading("role", text="角色")
    user_tree.column("name", width=110)
    user_tree.column("barcode", width=240)
    user_tree.column("role", width=70)
    user_tree.pack(fill="both", expand=True, padx=8, pady=(8, 4))
    form = ttk.Frame(user_tab)
    form.pack(fill="x", padx=8, pady=4)
    ttk.Label(form, text="姓名:").grid(row=0, column=0)
    user_name_var = tk.StringVar()
    ttk.Entry(form, textvariable=user_name_var, width=12).grid(row=0, column=1, padx=(0, 8))
    ttk.Label(form, text="条码:").grid(row=0, column=2)
    user_code_var = tk.StringVar()
    ttk.Entry(form, textvariable=user_code_var, width=24).grid(row=0, column=3, padx=(0, 8))
    ttk.Label(form, text="角色:").grid(row=0, column=4)
    role_var = tk.StringVar(value=db.ROLE_MEMBER)
    ttk.Combobox(form, textvariable=role_var, values=[db.ROLE_MEMBER, db.ROLE_ADMIN], state="readonly", width=7).grid(row=0, column=5)
    ubtns = ttk.Frame(user_tab)
    ubtns.pack(fill="x", padx=8, pady=(0, 8))
    ttk.Button(ubtns, text="添加人员", command=admin_add_user).pack(side="left", padx=(0, 6))
    ttk.Button(ubtns, text="删除选中", command=admin_del_user).pack(side="left", padx=(0, 6))
    ttk.Button(ubtns, text="切换角色", command=admin_toggle_role).pack(side="left", padx=(0, 6))
    ttk.Button(ubtns, text="刷新", command=refresh_admin_tables).pack(side="left")

    # 物品管理
    item_tab = ttk.Frame(nb)
    nb.add(item_tab, text=" 物品管理 ")
    item_tree = ttk.Treeview(item_tab, columns=("code", "name", "cabinet", "status"), show="headings", height=12)
    item_tree.heading("code", text="条码")
    item_tree.heading("name", text="物品名称")
    item_tree.heading("cabinet", text="柜号")
    item_tree.heading("status", text="状态")
    item_tree.column("code", width=170)
    item_tree.column("name", width=140)
    item_tree.column("cabinet", width=60)
    item_tree.column("status", width=70)
    item_tree.pack(fill="both", expand=True, padx=8, pady=(8, 4))
    iform = ttk.Frame(item_tab)
    iform.pack(fill="x", padx=8, pady=4)
    ttk.Label(iform, text="条码:").grid(row=0, column=0)
    item_code_var = tk.StringVar()
    ttk.Entry(iform, textvariable=item_code_var, width=20).grid(row=0, column=1, padx=(0, 8))
    ttk.Label(iform, text="名称:").grid(row=0, column=2)
    item_name_var = tk.StringVar()
    ttk.Entry(iform, textvariable=item_name_var, width=14).grid(row=0, column=3, padx=(0, 8))
    ttk.Label(iform, text="柜号:").grid(row=0, column=4)
    item_cabinet_var = tk.StringVar()
    ttk.Entry(iform, textvariable=item_cabinet_var, width=8).grid(row=0, column=5)
    ibtns = ttk.Frame(item_tab)
    ibtns.pack(fill="x", padx=8, pady=(0, 8))
    ttk.Button(ibtns, text="添加物品", command=admin_add_item).pack(side="left", padx=(0, 6))
    ttk.Button(ibtns, text="删除选中", command=admin_del_item).pack(side="left", padx=(0, 6))
    ttk.Button(ibtns, text="刷新", command=refresh_admin_tables).pack(side="left")


def refresh_admin_tables():
    if admin_win is None or not admin_win.winfo_exists():
        return
    user_tree.delete(*user_tree.get_children())
    for u in warehouse_db.user_rows:
        user_tree.insert("", "end", values=(u["姓名"], u["条码"], u["角色"]))
    item_tree.delete(*item_tree.get_children())
    for it in warehouse_db.items:
        item_tree.insert("", "end", values=(
            it.get(db.COL_BARCODE, ""), it.get(db.COL_NAME, ""),
            it.get(db.COL_CABINET, ""), it.get(db.COL_STATUS, "")))


def _selected(tree):
    sel = tree.selection()
    if not sel:
        messagebox.showwarning("提示", "请先在列表中选择一行。")
        return None
    return tree.item(sel[0], "values")


def admin_add_user():
    name = user_name_var.get().strip()
    code = user_code_var.get().strip()
    if not name or not code:
        messagebox.showwarning("提示", "姓名和条码都不能为空。")
        return
    role = role_var.get()
    try:
        warehouse_db.add_user(code, name, role)
    except Exception as exc:
        messagebox.showerror("添加失败", str(exc))
        return
    user_name_var.set("")
    user_code_var.set("")
    log(f"[管理] 添加人员：{name}（{code}，{role}）")
    refresh_admin_tables()


def admin_del_user():
    values = _selected(user_tree)
    if values is None:
        return
    name, code = values[0], values[1]
    if not messagebox.askyesno("确认删除", f"确定删除人员「{name}」（{code}）吗？"):
        return
    try:
        warehouse_db.remove_user(code)
    except Exception as exc:
        messagebox.showerror("删除失败", str(exc))
        return
    log(f"[管理] 删除人员：{name}（{code}）")
    if name == current_user:
        logout_current()
    refresh_admin_tables()


def admin_toggle_role():
    global current_role
    values = _selected(user_tree)
    if values is None:
        return
    code = values[1]
    try:
        new_role = db.ROLE_MEMBER if warehouse_db.user_role(code) == db.ROLE_ADMIN else db.ROLE_ADMIN
        warehouse_db.set_user_role(code, new_role)
    except Exception as exc:
        messagebox.showerror("修改失败", str(exc))
        return
    log(f"[管理] 人员 {code} 角色改为：{new_role}")
    if values[0] == current_user:
        current_role = new_role
        update_admin_btn()
    refresh_admin_tables()


def admin_add_item():
    code = item_code_var.get().strip()
    name = item_name_var.get().strip()
    cabinet = item_cabinet_var.get().strip()
    if not code or not name:
        messagebox.showwarning("提示", "条码和物品名称不能为空。")
        return
    try:
        warehouse_db.add_item(code, name, cabinet)
    except Exception as exc:
        messagebox.showerror("添加失败", str(exc))
        return
    item_code_var.set("")
    item_name_var.set("")
    item_cabinet_var.set("")
    log(f"[管理] 添加物品：{name}（{code}，柜号 {cabinet}）")
    refresh_admin_tables()
    refresh_inventory_text()


def admin_del_item():
    values = _selected(item_tree)
    if values is None:
        return
    code, name = values[0], values[1]
    if not messagebox.askyesno("确认删除", f"确定删除物品「{name}」（{code}）吗？"):
        return
    try:
        warehouse_db.remove_item(code)
    except Exception as exc:
        messagebox.showerror("删除失败", str(exc))
        return
    log(f"[管理] 删除物品：{name}（{code}）")
    refresh_admin_tables()
    refresh_inventory_text()


def change_theme(event=None):
    """切换外观主题并保存到 config.json。"""
    theme = theme_var.get()
    if USE_TB:
        try:
            root.style.theme_use(theme)
            root.style.configure("Treeview", rowheight=26)
        except Exception as exc:
            log(f"[主题] 切换失败：{exc}")
            return
    apply_design_colors()
    apply_inventory_colors()
    APP_CONFIG["theme"] = theme
    try:
        CONFIG_FILE.write_text(json.dumps(APP_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass
    log(f"外观主题已切换：{theme}")


def clear_log():
    log_text.config(state="normal")
    log_text.delete("1.0", "end")
    log_text.config(state="disabled")


def init_database():
    """启动时加载数据库，失败给出明确提示。"""
    global warehouse_db, checkin_log
    existed = CHECKIN_EXCEL.exists()
    try:
        warehouse_db = db.WarehouseDB(WAREHOUSE_EXCEL, backup=BACKUP_WAREHOUSE)
    except FileNotFoundError as exc:
        messagebox.showerror("数据文件缺失", str(exc))
        sys.exit(1)
    except ValueError as exc:
        messagebox.showerror("数据表结构错误", str(exc))
        sys.exit(1)
    try:
        checkin_log = db.CheckinLog(CHECKIN_EXCEL)
    except Exception as exc:
        messagebox.showerror("签到表初始化失败", str(exc))
        sys.exit(1)
    log(f"系统启动：已加载 {len(warehouse_db.items)} 件物品、{len(warehouse_db.users)} 名人员。")
    for warning in warehouse_db.warnings:
        log(f"[警告] {warning}")
    if not existed:
        log(f"已自动创建签到表：{CHECKIN_EXCEL.name}")


# ================= 界面 UI 设计 =================
root.configure(bg=BG)

# 0. 顶部横幅
header = tk.Frame(root, bg=PRIMARY, padx=20, pady=12)
header.pack(fill="x")
header_title = tk.Label(header, text="实验室综合管理系统", bg=PRIMARY, fg="#ffffff", font=F_TITLE)
header_title.pack(anchor="w")
header_sub = tk.Label(header, text="人脸识别 · 扫码出入库 · 断线自动重连 · 数据自动备份",
                      bg=PRIMARY, fg="#cfe0ff", font=F_SMALL)
header_sub.pack(anchor="w", pady=(2, 0))

# 1. 使用说明卡片
guide_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
guide_card.pack(padx=16, pady=(12, 6), fill="x")
guide_title = tk.Label(guide_card, text="使用说明", bg=CARD, fg=PRIMARY, font=F_SECTION)
guide_title.pack(anchor="w", padx=14, pady=(10, 2))
guide_text = ("1. 选择串口并点击【启动监听】\n"
              "2. 刷脸或扫人员码登录\n"
              "3. 选择【入库】或【出库】模式\n"
              "4. 扫描物品条码完成登记\n"
              "5. 重复扫描不会重复登记，结果见下方日志\n"
              "6. 管理员登录后可点【管理】维护数据")
guide_body = tk.Label(guide_card, text=guide_text, justify="left", bg=CARD, fg=MUTED, font=F_BODY)
guide_body.pack(anchor="w", padx=14, pady=(0, 10))

# 2. 串口设置卡片
serial_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
serial_card.pack(padx=16, pady=6, fill="x")
serial_title = tk.Label(serial_card, text="串口设置", bg=CARD, fg=PRIMARY, font=F_SECTION)
serial_title.pack(anchor="w", padx=14, pady=(10, 2))

serial_grid = tk.Frame(serial_card, bg=CARD)
serial_grid.pack(fill="x", padx=14)
scanner_port_var = tk.StringVar(value=DEFAULT_SCANNER_PORT)
face_port_var = tk.StringVar(value=DEFAULT_FACE_PORT)
scan_label = tk.Label(serial_grid, text="扫码枪", bg=CARD, fg=TEXT, font=F_BODY, width=7, anchor="w")
scan_label.grid(row=0, column=0, sticky="w", pady=4)
scanner_combo = ttk.Combobox(serial_grid, textvariable=scanner_port_var, state="readonly", width=28, font=F_BODY)
scanner_combo.grid(row=0, column=1, sticky="ew", pady=4)
face_label = tk.Label(serial_grid, text="人脸模块", bg=CARD, fg=TEXT, font=F_BODY, width=7, anchor="w")
face_label.grid(row=1, column=0, sticky="w", pady=4)
face_combo = ttk.Combobox(serial_grid, textvariable=face_port_var, state="readonly", width=28, font=F_BODY)
face_combo.grid(row=1, column=1, sticky="ew", pady=4)
serial_grid.columnconfigure(1, weight=1)

serial_btn_row = tk.Frame(serial_card, bg=CARD)
serial_btn_row.pack(fill="x", padx=14, pady=(8, 0))
mk_button(serial_btn_row, "刷新串口", refresh_ports, "secondary").pack(side=tk.LEFT, padx=(0, 8))
start_btn = mk_button(serial_btn_row, "▶ 启动监听", start_serial, "success")
start_btn.pack(side=tk.LEFT, padx=(0, 8))
stop_btn = mk_button(serial_btn_row, "■ 停止监听", stop_serial, "danger", state="disabled")
stop_btn.pack(side=tk.LEFT)

serial_status = tk.Label(serial_card, text="串口状态: 未启动", anchor="w", bg=CARD, fg=MUTED, font=F_BODY)
serial_status.pack(fill="x", padx=14, pady=(8, 0))

switch_row = tk.Frame(serial_card, bg=CARD)
switch_row.pack(fill="x", padx=14, pady=(2, 12))
topmost_var = tk.BooleanVar(value=True)
mk_check(switch_row, "窗口置顶", topmost_var, lambda: root.attributes("-topmost", topmost_var.get())).pack(side=tk.LEFT)
autostart_var = tk.BooleanVar(value=AUTOSTART_FILE.exists())
mk_check(switch_row, "开机自启", autostart_var, toggle_autostart).pack(side=tk.LEFT, padx=(20, 0))
theme_label = tk.Label(switch_row, text="外观主题", bg=CARD, fg=MUTED, font=F_SMALL)
theme_label.pack(side=tk.RIGHT)
theme_var = tk.StringVar(value=current_theme())
theme_combo = ttk.Combobox(switch_row, textvariable=theme_var, values=THEMES, state="readonly", width=11, font=F_SMALL)
theme_combo.pack(side=tk.RIGHT, padx=(6, 0))
theme_combo.bind("<<ComboboxSelected>>", change_theme)

# 3. 状态卡片
status_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
status_card.pack(padx=16, pady=6, fill="x")
status_row = tk.Frame(status_card, bg=CARD)
status_row.pack(fill="x", padx=14, pady=(10, 0))
user_display = tk.Label(status_row, text=f"当前用户: {current_user}", bg=CARD, fg=BLUE, font=F_BODY_BOLD)
user_display.pack(side=tk.LEFT)
mode_display = tk.Label(status_row, text=f"当前模式: 等待选择", bg=CARD, fg=RED, font=F_BODY_BOLD)
mode_display.pack(side=tk.LEFT, padx=(24, 0))
status_hint = tk.Label(status_card, text="", bg=CARD, fg=GREEN, font=F_BODY, anchor="w")
status_hint.pack(fill="x", padx=14, pady=(4, 10))

# 4. 出入库操作卡片
action_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
action_card.pack(padx=16, pady=6, fill="x")
action_title = tk.Label(action_card, text="出入库操作", bg=CARD, fg=PRIMARY, font=F_SECTION)
action_title.pack(anchor="w", padx=14, pady=(10, 2))
action_row = tk.Frame(action_card, bg=CARD)
action_row.pack(fill="x", padx=14, pady=(0, 12))
mk_button(action_row, "入库模式", set_mode_in, "primary", height=2).grid(row=0, column=0, padx=(0, 8), sticky="ew")
mk_button(action_row, "出库模式", set_mode_out, "warning", height=2).grid(row=0, column=1, padx=(0, 8), sticky="ew")
mk_button(action_row, "查看库存", show_inventory, "info", height=2).grid(row=0, column=2, padx=(0, 8), sticky="ew")
admin_btn = mk_button(action_row, "管理", open_admin, "danger", height=2, state="disabled")
admin_btn.grid(row=0, column=3, sticky="ew")
for col in range(4):
    action_row.columnconfigure(col, weight=1, uniform="action")

# 5. 手动输入卡片
manual_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
manual_card.pack(padx=16, pady=6, fill="x")
manual_row = tk.Frame(manual_card, bg=CARD)
manual_row.pack(fill="x", padx=14, pady=10)
manual_label = tk.Label(manual_row, text="手动输入条码", bg=CARD, fg=TEXT, font=F_BODY, width=9, anchor="w")
manual_label.pack(side=tk.LEFT)
manual_var = tk.StringVar()
manual_entry = tk.Entry(manual_row, textvariable=manual_var, font=F_BODY, relief="solid", bd=1)
manual_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8))
manual_entry.bind("<Return>", lambda event: submit_manual())
mk_button(manual_row, "提交", submit_manual, "secondary", width=8).pack(side=tk.LEFT)

# 6. 运行日志卡片（深色终端风）
log_card = tk.Frame(root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
log_card.pack(padx=16, pady=(6, 14), fill="both", expand=True)
log_head = tk.Frame(log_card, bg=CARD)
log_head.pack(fill="x", padx=14, pady=(8, 2))
log_title = tk.Label(log_head, text="运行日志", bg=CARD, fg=PRIMARY, font=F_SECTION)
log_title.pack(side=tk.LEFT)
clear_btn = tk.Button(log_head, text="清空", command=clear_log, font=F_SMALL, relief="flat", bd=0,
                      bg=CARD, fg=PRIMARY, activebackground=CARD, activeforeground=PRIMARY,
                      cursor="hand2")
clear_btn.pack(side=tk.RIGHT)
log_box = tk.Frame(log_card, bg=LOG_BG)
log_box.pack(fill="both", expand=True, padx=14, pady=(0, 12))
log_text = tk.Text(log_box, height=8, bg=LOG_BG, fg=LOG_FG, insertbackground="#ffffff",
                   font=("DejaVu Sans Mono", 9), state="disabled", relief="flat", bd=0,
                   padx=8, pady=6, wrap="word")
log_scroll = tk.Scrollbar(log_box, command=log_text.yview)
log_text.configure(yscrollcommand=log_scroll.set)
log_scroll.pack(side=tk.RIGHT, fill="y")
log_text.pack(side=tk.LEFT, fill="both", expand=True)


def apply_design_colors():
    """统一应用固定配色。ttkbootstrap 创建/切换主题时会重置经典 tk 控件颜色，这里重新设置。"""
    root.configure(bg=BG)
    header.config(bg=PRIMARY)
    header_title.config(bg=PRIMARY, fg="#ffffff")
    header_sub.config(bg=PRIMARY, fg="#cfe0ff")
    for card in (guide_card, serial_card, status_card, action_card, manual_card, log_card):
        card.config(bg=CARD, highlightbackground=BORDER, highlightthickness=1)
    for panel in (guide_card, serial_card, status_card, action_card, manual_card):
        for w in panel.winfo_children():
            if isinstance(w, (tk.Frame, tk.Label, tk.Entry)):
                w.config(bg=CARD)
    guide_title.config(fg=PRIMARY)
    guide_body.config(fg=MUTED)
    serial_title.config(fg=PRIMARY)
    scan_label.config(fg=TEXT)
    face_label.config(fg=TEXT)
    theme_label.config(fg=MUTED)
    serial_status.config(bg=CARD, fg=MUTED if not serial_running else serial_status.cget("fg"))
    user_display.config(bg=CARD, fg=BLUE)
    mode_display.config(bg=CARD, fg=_MODE_COLORS.get(current_mode, RED))
    status_hint.config(bg=CARD)
    action_title.config(fg=PRIMARY)
    manual_label.config(fg=TEXT)
    manual_entry.config(bg="#ffffff", fg=TEXT, relief="solid", bd=1)
    log_head.config(bg=CARD)
    log_title.config(bg=CARD, fg=PRIMARY)
    clear_btn.config(bg=CARD, fg=PRIMARY, activebackground=CARD, activeforeground=PRIMARY)
    log_box.config(bg=LOG_BG)
    log_text.config(bg=LOG_BG, fg=LOG_FG, insertbackground="#ffffff")
    log_scroll.config(bg="#2b3644", troughcolor=LOG_BG, activebackground="#3d4c5f", relief="flat", bd=0)


def apply_inventory_colors():
    """库存窗口配色（经典控件，需在主题切换后重新应用）。"""
    if inventory_win is None or not inventory_win.winfo_exists():
        return
    inventory_win.config(bg=CARD)
    inventory_stats.config(bg=CARD, fg=TEXT)
    inventory_tree_wrap.config(bg=CARD)
    inventory_btn_row.config(bg=CARD)
    inventory_tree.tag_configure("in", foreground=GREEN)
    inventory_tree.tag_configure("out", foreground=MUTED)


apply_design_colors()

# ================= 启动 =================
START_MINIMIZED = "--minimized" in sys.argv   # 开机自启时最小化启动
root.protocol("WM_DELETE_WINDOW", on_close)
if not acquire_single_instance():
    messagebox.showerror("程序已在运行", "检测到系统已在运行中。\n请使用已经打开的窗口，不要重复启动。")
    sys.exit(1)
refresh_ports()
init_database()
if START_MINIMIZED:
    root.after(200, root.iconify)
    log("已最小化启动（开机自启模式）")


def _fit_window():
    """按内容自适应窗口尺寸（不小于默认值，不超过屏幕范围）。"""
    try:
        root.update_idletasks()
        need_w = root.winfo_reqwidth() + 24
        need_h = root.winfo_reqheight() + 24
        max_w = root.winfo_screenwidth() - 120
        max_h = root.winfo_screenheight() - 80
        w = max(660, min(need_w, max_w))
        h = max(860, min(need_h, max_h))
        root.geometry(f"{w}x{h}")
        log(f"窗口已按内容调整：{w}x{h}")
    except Exception:
        pass


root.after(80, _fit_window)
root.mainloop()
