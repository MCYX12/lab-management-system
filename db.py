#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""实验室综合管理系统 · 数据层

直接用 openpyxl 读写 Excel（不依赖 pandas）：

- 启动时把「物品表 / 人员表」读入内存缓存，写库成功后即时更新缓存；
- 检测文件修改时间（mtime）：外部用 Excel/WPS 改过文件后自动重载，避免旧缓存覆盖新数据；
- 物品状态更新为「定点改单元格」，保留表格样式、其他工作表与列宽；
- 签到表按「追加一行」写入，行数再多也很快；
- 写库前自动备份 warehouse_db.xlsx -> warehouse_db.bak（可关闭）。
"""

import os
import shutil
import time
from datetime import datetime
from pathlib import Path

import openpyxl

# ---- 表结构约定 ----
ITEM_SHEET = "物品表"
USER_SHEET = "人员表"
CHECKIN_SHEET = "签到表"

COL_BARCODE = "条码"
COL_NAME = "物品名称"
COL_CABINET = "柜号"
COL_STATUS = "状态"
COL_OP_TIME = "最后操作时间"
COL_OPERATOR = "最近操作人"
REQUIRED_ITEM_COLS = [COL_BARCODE, COL_NAME, COL_CABINET, COL_STATUS, COL_OP_TIME, COL_OPERATOR]

STATUS_IN = "已入库"
STATUS_OUT = "已出库"

ROLE_ADMIN = "管理员"
ROLE_MEMBER = "成员"
COL_ROLE = "角色"

TIME_FMT = "%Y-%m-%d %H:%M:%S"


def norm(value):
    """单元格值 -> 规范字符串：None -> ""；浮点整数去尾巴（123.0 -> "123"）；去首尾空白。"""
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    return str(value).strip()


def _save_with_retry(wb, path, retries=1, delay=0.3):
    """先写临时文件再原子替换，断电/崩溃也不会损坏 xlsx；被占用时等待并重试一次。"""
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    last_exc = None
    for _ in range(retries + 1):
        try:
            wb.save(tmp)
            os.replace(tmp, path)   # 同目录原子替换
            return
        except PermissionError as exc:
            last_exc = exc
            time.sleep(delay)
    try:
        if tmp.exists():
            tmp.unlink()
    except OSError:
        pass
    raise PermissionError(
        f"无法写入 {path.name}：文件可能正被 Excel/WPS 打开，请关闭后重试。"
    ) from last_exc


class WarehouseDB:
    """仓库数据库：warehouse_db.xlsx 的「物品表」「人员表」。"""

    def __init__(self, path, backup=True):
        self.path = Path(path)
        self.backup = backup
        self.warnings = []       # 加载时的警告（如重复条码）
        self._mtime = None
        self._items = []         # 物品行列表（含 "_row" 行号）
        self._index = {}         # 条码 -> 物品 dict
        self._users = {}         # 人员条码 -> 姓名
        self.load()

    # ---------- 读取 ----------
    @property
    def items(self):
        self.reload_if_changed()
        return list(self._items)

    @property
    def users(self):
        self.reload_if_changed()
        return dict(self._users)

    def reload_if_changed(self):
        """文件被外部修改后自动重载。"""
        if self.path.exists() and self.path.stat().st_mtime != self._mtime:
            self.load()

    @property
    def _bak_path(self):
        return self.path.with_name(self.path.stem + ".bak")

    def load(self):
        """读入数据库；主文件缺失或损坏时自动从 .bak 备份恢复。"""
        recovery = []
        if not self.path.exists():
            if self._bak_path.exists():
                shutil.copy2(self._bak_path, self.path)
                recovery.append(f"{self.path.name} 缺失，已自动从 {self._bak_path.name} 恢复")
            else:
                raise FileNotFoundError(f"找不到仓库数据库：{self.path}")
        try:
            self._read()
        except Exception as exc:
            if self._bak_path.exists():
                shutil.copy2(self._bak_path, self.path)
                recovery.append(f"{self.path.name} 无法读取（{exc}），已自动从 {self._bak_path.name} 恢复")
                try:
                    self._read()
                except Exception as exc2:
                    raise ValueError(
                        f"{self.path.name} 与其备份 {self._bak_path.name} 均无法读取：{exc2}") from exc2
            else:
                raise ValueError(f"{self.path.name} 无法读取，且没有可用备份：{exc}") from exc
        if recovery:
            self.warnings = recovery + self.warnings
        self._mtime = self.path.stat().st_mtime

    def _read(self):
        """读取并解析工作簿（load 的内部实现）。"""
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=False)
        try:
            if ITEM_SHEET not in wb.sheetnames or USER_SHEET not in wb.sheetnames:
                raise ValueError(f"{self.path.name} 缺少工作表「{ITEM_SHEET}」或「{USER_SHEET}」")
            ws_items = wb[ITEM_SHEET]
            ws_users = wb[USER_SHEET]

            header = [norm(c.value) for c in next(ws_items.iter_rows(min_row=1, max_row=1))]
            missing = [n for n in REQUIRED_ITEM_COLS if n not in header]
            if missing:
                raise ValueError(f"「{ITEM_SHEET}」缺少列：{'、'.join(missing)}")
            self._item_header = header

            items, warnings = [], []
            for row_no, row in enumerate(ws_items.iter_rows(min_row=2, values_only=True), start=2):
                if not row or all(v is None or not str(v).strip() for v in row):
                    continue
                item = {name: norm(v) for name, v in zip(header, row)}
                if not item.get(COL_BARCODE):
                    continue
                item["_row"] = row_no
                items.append(item)
            index = {}
            for item in items:
                barcode = item[COL_BARCODE]
                if barcode in index:
                    warnings.append(f"「{ITEM_SHEET}」条码重复：{barcode}（只更新第一行）")
                    continue
                index[barcode] = item
            self._items, self._index, self.warnings = items, index, warnings

            uheader = [norm(c.value) for c in next(ws_users.iter_rows(min_row=1, max_row=1))]
            if COL_BARCODE not in uheader or "姓名" not in uheader:
                raise ValueError(f"「{USER_SHEET}」缺少列：条码 / 姓名")
            bcol, ncol = uheader.index(COL_BARCODE), uheader.index("姓名")
            rcol = uheader.index(COL_ROLE) if COL_ROLE in uheader else None
            users = {}
            for row in ws_users.iter_rows(min_row=2, values_only=True):
                barcode = norm(row[bcol] if row and len(row) > bcol else None)
                if not barcode:
                    continue
                if rcol is None:
                    role = ROLE_ADMIN   # 还没有「角色」列：旧数据默认全部管理员，避免被锁在门外
                else:
                    role = norm(row[rcol] if len(row) > rcol else None)
                users[barcode] = {
                    "name": norm(row[ncol] if len(row) > ncol else None),
                    "role": role if role == ROLE_ADMIN else ROLE_MEMBER,
                }
            self._users = users
        finally:
            wb.close()

    # ---------- 查询 ----------
    def find_user(self, barcode):
        """按条码查人员姓名，找不到返回 None。"""
        self.reload_if_changed()
        info = self._users.get(str(barcode).strip())
        return info["name"] if info else None

    def user_role(self, barcode):
        """按条码查角色（管理员/成员），找不到返回「成员」。"""
        self.reload_if_changed()
        info = self._users.get(str(barcode).strip())
        return info["role"] if info else ROLE_MEMBER

    def find_role_by_name(self, name):
        """按姓名查角色（人脸签到用），找不到返回「成员」。"""
        self.reload_if_changed()
        for info in self._users.values():
            if info["name"] == name:
                return info["role"]
        return ROLE_MEMBER

    def find_item(self, barcode):
        """按条码查物品信息（含 "_row" 行号），找不到返回 None。"""
        self.reload_if_changed()
        return self._index.get(str(barcode).strip())

    @property
    def user_rows(self):
        """人员表完整行列表（管理员面板用）。"""
        self.reload_if_changed()
        return [{"条码": bc, "姓名": u["name"], "角色": u["role"]} for bc, u in self._users.items()]

    # ---------- 写入 ----------
    def _backup(self):
        """写库前备份当前文件（失败不阻断主流程）。"""
        if self.backup:
            try:
                shutil.copy2(self.path, self._bak_path)
            except OSError:
                pass

    def update_item(self, barcode, status, operator):
        """把某物品置为「已入库/已出库」并记录操作人与时间。返回更新后的物品信息。"""
        if status not in (STATUS_IN, STATUS_OUT):
            raise ValueError(f"非法的状态值：{status}")
        self.reload_if_changed()
        item = self._index.get(str(barcode).strip())
        if item is None:
            return None
        self._backup()
        now = datetime.now().strftime(TIME_FMT)
        wb = openpyxl.load_workbook(self.path)
        try:
            ws = wb[ITEM_SHEET]
            header_row = [norm(c.value) for c in ws[1]]
            cols = {name: i + 1 for i, name in enumerate(header_row)}
            # 按条码重新定位行：即使外部插行/删行导致行号偏移也能写对
            row_no = self._locate_row(ws, cols[COL_BARCODE], barcode)
            if row_no is None:
                raise ValueError(f"「{ITEM_SHEET}」中找不到条码 {barcode} 所在行")
            ws.cell(row=row_no, column=cols[COL_STATUS]).value = status
            ws.cell(row=row_no, column=cols[COL_OP_TIME]).value = now
            ws.cell(row=row_no, column=cols[COL_OPERATOR]).value = operator
            _save_with_retry(wb, self.path)
        finally:
            wb.close()
        # 同步内存缓存，并把 mtime 对齐到刚写的文件
        item[COL_STATUS] = status
        item[COL_OP_TIME] = now
        item[COL_OPERATOR] = operator
        item["_row"] = row_no
        self._mtime = self.path.stat().st_mtime
        return dict(item)

    @staticmethod
    def _locate_row(ws, barcode_col, barcode):
        for row_no in range(2, ws.max_row + 1):
            if norm(ws.cell(row=row_no, column=barcode_col).value) == barcode:
                return row_no
        return None

    # ---------- 管理操作（管理员面板） ----------
    def _manage(self, mutate):
        """管理操作通用流程：备份 → 修改 → 原子保存 → 重载缓存。"""
        self._backup()
        wb = openpyxl.load_workbook(self.path)
        try:
            mutate(wb)
            _save_with_retry(wb, self.path)
        finally:
            wb.close()
        self.load()

    @staticmethod
    def _ensure_role_col(ws):
        """确保人员表有「角色」列；首次创建时把现有人员填充为管理员。"""
        header = [norm(c.value) for c in ws[1]]
        if COL_ROLE in header:
            return header.index(COL_ROLE) + 1
        col = len(header) + 1
        ws.cell(row=1, column=col).value = COL_ROLE
        for row_no in range(2, ws.max_row + 1):
            if ws.cell(row=row_no, column=col).value in (None, ""):
                ws.cell(row=row_no, column=col).value = ROLE_ADMIN
        return col

    def add_item(self, barcode, name, cabinet):
        """新增物品（默认已入库）。"""
        barcode, name, cabinet = str(barcode).strip(), str(name).strip(), str(cabinet).strip()
        if not barcode or not name:
            raise ValueError("条码和物品名称不能为空")
        if self.find_item(barcode) is not None:
            raise ValueError(f"条码 {barcode} 已存在")
        self._manage(lambda wb: wb[ITEM_SHEET].append([barcode, name, cabinet, STATUS_IN, "", ""]))

    def remove_item(self, barcode):
        """删除物品。"""
        barcode = str(barcode).strip()
        if self.find_item(barcode) is None:
            raise ValueError(f"物品条码 {barcode} 不存在")

        def mutate(wb):
            ws = wb[ITEM_SHEET]
            cols = {norm(c.value): i + 1 for i, c in enumerate(ws[1])}
            ws.delete_rows(self._locate_row(ws, cols[COL_BARCODE], barcode))
        self._manage(mutate)

    def add_user(self, barcode, name, role=ROLE_MEMBER):
        """新增人员（自动确保「角色」列存在）。"""
        barcode, name = str(barcode).strip(), str(name).strip()
        if not barcode or not name:
            raise ValueError("条码和姓名不能为空")
        if self.find_user(barcode) is not None:
            raise ValueError(f"人员条码 {barcode} 已存在")
        if role not in (ROLE_ADMIN, ROLE_MEMBER):
            role = ROLE_MEMBER

        def mutate(wb):
            ws = wb[USER_SHEET]
            self._ensure_role_col(ws)
            ws.append([barcode, name, role])
        self._manage(mutate)

    def remove_user(self, barcode):
        """删除人员。"""
        barcode = str(barcode).strip()
        if self.find_user(barcode) is None:
            raise ValueError(f"人员条码 {barcode} 不存在")

        def mutate(wb):
            ws = wb[USER_SHEET]
            cols = {norm(c.value): i + 1 for i, c in enumerate(ws[1])}
            ws.delete_rows(self._locate_row(ws, cols[COL_BARCODE], barcode))
        self._manage(mutate)

    def set_user_role(self, barcode, role):
        """修改人员角色。"""
        barcode = str(barcode).strip()
        if role not in (ROLE_ADMIN, ROLE_MEMBER):
            raise ValueError(f"非法角色：{role}")
        if self.find_user(barcode) is None:
            raise ValueError(f"人员条码 {barcode} 不存在")

        def mutate(wb):
            ws = wb[USER_SHEET]
            cols = {norm(c.value): i + 1 for i, c in enumerate(ws[1])}
            rcol = self._ensure_role_col(ws)
            ws.cell(row=self._locate_row(ws, cols[COL_BARCODE], barcode), column=rcol).value = role
        self._manage(mutate)


class CheckinLog:
    """签到表：只追加写入，文件不存在时自动创建（带表头）。"""

    def __init__(self, path):
        self.path = Path(path)
        self._ensure()

    def _ensure(self):
        if not self.path.exists():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = CHECKIN_SHEET
            ws.append(["姓名", "签到方式", "时间"])
            _save_with_retry(wb, self.path)

    def append(self, name, method="人脸识别"):
        self._ensure()
        wb = openpyxl.load_workbook(self.path)
        try:
            ws = wb[CHECKIN_SHEET] if CHECKIN_SHEET in wb.sheetnames else wb.active
            ws.append([name, method, datetime.now().strftime(TIME_FMT)])
            _save_with_retry(wb, self.path)
        finally:
            wb.close()
