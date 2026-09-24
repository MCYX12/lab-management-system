#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

echo "检查 Ubuntu 系统依赖..."
missing=()
for cmd in python3 pip3; do
  if ! command -v "$cmd" >/dev/null 2>&1; then
    missing+=("$cmd")
  fi
done

if ! python3 -c "import tkinter" >/dev/null 2>&1; then
  missing+=("python3-tk")
fi

if ! python3 -m venv --help >/dev/null 2>&1; then
  missing+=("python3-venv")
fi

if [ "${#missing[@]}" -gt 0 ]; then
  echo "需要先安装系统组件:"
  echo "sudo apt update"
  echo "sudo apt install -y python3 python3-pip python3-venv python3-tk"
  exit 1
fi

echo "创建/更新项目虚拟环境..."
if [ ! -x ".venv/bin/python" ]; then
  python3 -m venv .venv
fi

.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt

chmod +x run.sh

if ! groups "$USER" | grep -qw dialout; then
  echo
  echo "当前用户还不在 dialout 组，串口可能无法打开。"
  echo "请执行:"
  echo "sudo usermod -aG dialout \"$USER\""
  echo "然后注销并重新登录。"
else
  echo "串口权限组 dialout 已具备。"
fi

echo "创建桌面快捷方式..."
DESKTOP_DIR="$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/Desktop")"
APP_DIR="$(pwd)"
DESKTOP_FILE="$DESKTOP_DIR/实验室综合管理系统.desktop"
cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Name=实验室综合管理系统
Comment=人脸识别和扫码出入库
Exec=bash -lc 'cd "$APP_DIR" && ./run.sh'
Terminal=false
Categories=Utility;
EOF
chmod +x "$DESKTOP_FILE"
if command -v gio >/dev/null 2>&1; then
  gio set "$DESKTOP_FILE" metadata::trusted true 2>/dev/null || true
fi
echo "桌面快捷方式已创建：$DESKTOP_FILE"

echo "开启开机自启（登录后自动最小化启动）..."
AUTOSTART_DIR="$HOME/.config/autostart"
mkdir -p "$AUTOSTART_DIR"
AUTOSTART_FILE="$AUTOSTART_DIR/实验室综合管理系统.desktop"
cat > "$AUTOSTART_FILE" <<EOF
[Desktop Entry]
Type=Application
Name=实验室综合管理系统
Comment=实验室综合管理系统开机自启
Exec=bash -lc 'cd "$APP_DIR" && "$APP_DIR/.venv/bin/python" main.py --minimized'
Terminal=false
X-GNOME-Autostart-enabled=true
EOF
echo "已开启开机自启（可在程序窗口里用「开机自动启动」开关关闭）"

echo
echo "安装完成。以后运行:"
echo "cd \"$(pwd)\" && ./run.sh"
