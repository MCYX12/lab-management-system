# Ubuntu 运行说明

## 从 GitHub 仓库获取

仓库出于隐私考虑不包含数据文件（`warehouse_db.xlsx`、签到表）。克隆后：

```bash
git clone https://github.com/MCYX12/lab-management-system.git
cd lab-management-system
cp warehouse_db.example.xlsx warehouse_db.xlsx
chmod +x install_ubuntu.sh run.sh
./install_ubuntu.sh
```

然后用 Excel/WPS 把 `warehouse_db.xlsx` 里的人员/物品改成真实数据。

## 换到新 Ubuntu 设备

把整个 `人脸+扫码` 文件夹复制到新设备，保留里面的 `main.py`、`db.py`、`warehouse_db.xlsx`、`requirements.txt`、`install_ubuntu.sh` 和 `run.sh`。

不要复制旧设备上的 `.venv` 文件夹；新设备运行安装脚本时会自动重新创建。

## 首次安装

在本目录执行：

```bash
cd ~/下载/人脸+扫码
chmod +x install_ubuntu.sh run.sh
./install_ubuntu.sh
```

如果脚本提示缺系统组件，先执行它显示的 `sudo apt install ...` 命令，再重新运行 `./install_ubuntu.sh`。

如果脚本提示当前用户不在 `dialout` 组，执行：

```bash
sudo usermod -aG dialout "$USER"
```

执行后需要注销并重新登录，否则串口权限不会生效。

## 日常运行

```bash
cd ~/下载/人脸+扫码
./run.sh
```

打开窗口后点击“刷新串口”，分别选择扫码枪和人脸识别模块对应的 `/dev/ttyUSB*`、`/dev/ttyACM*` 或 `/dev/serial/by-id/*`，再点击“启动监听”。

也可以用环境变量预设端口：

```bash
SCANNER_PORT=/dev/ttyUSB0 FACE_PORT=/dev/ttyUSB1 ./run.sh
```

## 桌面快捷方式

运行 `./install_ubuntu.sh` 时会自动在桌面生成「实验室综合管理系统」快捷方式（路径按实际安装位置写死，并自动标记为可信任）。

如果桌面没有自动生成，也可以手动把 `实验室综合管理系统.desktop` 复制到桌面，打开文件把 `Exec=` 一行里的路径改成实际文件夹位置，再右键 → 允许启动。

## 更新说明（v2.1 / v2.2）

- 数据库读写改为 openpyxl 直接操作（不再依赖 pandas），启动更快、扫码响应更快；
- 每次写库前自动备份 `warehouse_db.xlsx` 为 `warehouse_db.bak`，误操作可以恢复；
- 同一物品重复扫描不会重复登记，窗口底部日志区会提示；
- 同一人 15 秒内重复刷脸不重复签到（可在 `main.py` 顶部 `FACE_DEDUP_SECONDS` 调整，设为 0 关闭）；
- 扫码成功不再弹阻塞窗口，改为状态栏提示 + 日志记录，连续扫码不卡顿；
- 运行日志显示在窗口底部，串口异常、未匹配条码等一目了然；
- 程序会自动创建缺失的《实验室签到表.xlsx》，旧版生成的签到表可直接继续使用；
- 注意：`物品表`/`人员表` 的条码列请保持文本格式（右键单元格 → 设置单元格格式 → 文本），避免长条码被 Excel 转成科学计数法。

### v2.2 新增

- 写入改为「先写临时文件再原子替换」，断电/崩溃也不会损坏数据库；若数据库损坏或丢失，启动时自动用 `warehouse_db.bak` 恢复；
- 扫码枪输入带缓冲处理，兼容回车、换行、无结束符等各类扫码枪；
- 窗口新增「手动输入条码」框，没有硬件时也能测试出入库流程；
- 库存清单顶部显示 总数/在库/已出库 统计；
- 防止重复启动：同一文件夹下只允许运行一个程序实例（多开会互相覆盖数据）；
- 未登录或未选择出入库模式时的提示改为状态栏提示 + 日志，不再弹窗打断连续扫码；
- `install_ubuntu.sh` 会自动在桌面生成快捷方式，不再需要手动复制。

### v2.3 新增

- **串口断线自动重连**：扫码枪/人脸模块 USB 松了会自动重连，状态栏实时显示各端口状态，无需手动重启监听；
- **运行日志落盘**：日志自动写入 `运行日志.log`（超过 1MB 转存 `运行日志.old.log`），连未捕获的界面异常也会记录，排查问题更方便；
- **开机自启**：安装脚本默认开启，登录后自动最小化启动；也可在窗口里用「开机自动启动」开关随时开关；
- **多操作员权限**：`人员表` 新增「角色」列（管理员/成员）。管理员登录后【管理】按钮解锁（默认密码 `admin888`，可在 `main.py` 顶部 `ADMIN_PASSWORD` 修改），可在面板里增删人员、切换角色、增删物品；普通成员只能扫码签到和出入库；
- **旧数据兼容**：旧版 `人员表` 没有「角色」列时，现有人员默认全部按管理员处理（首次在面板里添加/修改人员时会自动创建角色列并保留管理员身份），请管理员及时为其他成员设置角色。

### v2.4 新增（界面美化）

- 全新卡片式界面：顶部蓝色横幅 + 白色卡片分区，按钮/输入框/表格/选项卡全部使用现代化主题（ttkbootstrap，Bootstrap 风格）；
- 运行日志改为深色终端风格，串口/扫码/错误信息一目了然；
- 库存清单从纯文本升级为表格视图，「已入库」绿色、「已出库」灰色，顶部带统计栏；
- 内置 9 款外观主题（浅色：cosmo/litera/minty/pulse/flatly；深色：darkly/solar/cyborg/superhero），在「外观主题」下拉框一键切换，选择保存在 `config.json`；
- 窗口尺寸按内容自适应，高分屏（200%+ 缩放）下也不会出现控件溢出；
- 新增依赖 `ttkbootstrap`，`install_ubuntu.sh` 会自动安装。
