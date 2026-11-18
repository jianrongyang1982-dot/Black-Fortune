# -*- coding: utf-8 -*-
"""
WorkPrice Excel 导入工具
=======================
功能：
  1. 通过对话框选择 Excel 文件；
  2. 读取 ownerid, cateid, catename, price1, price2, price3, workmode 字段；
  3. 以 (ownerid, cateid, workmode) 为业务键增量写入 westlake_WorkPrice：
       - 键已存在 → 只更新 price1/price2/price3（其它字段一律保持原值）；
       - 键不存在 → 插入新行。
     不会删除表中任何现有数据。

  注意：该表没有主键/唯一索引，不能用 INSERT ... ON DUPLICATE KEY UPDATE，
  因此先一次性读出表内已有业务键集合，再逐行 UPDATE 或 INSERT，包在同一个事务里。

表字段映射（其它字段自动/固定写入）：
  - id            : 代码端生成自增 id（该表无 AUTO_INCREMENT）
  - ownerid       : 来自 Excel
  - cateid        : 来自 Excel
  - catename      : 来自 Excel
  - price1/2/3    : 来自 Excel
  - workmode      : 来自 Excel，varchar(255) NOT NULL
                    1=Picking  2=Scanning  3=Replenishment
                    Excel 无该列或单元格为空时回退默认值 '1'（Picking）
  - dbsplitcode   : 固定 1
  - creator       : 固定 'sys'
  - modifier      : 固定 'sys'
  - createTime    : 当前日期时间
  - modifytime    : 当前日期时间

数据库配置优先读取同目录下的 db_config.json，不存在则使用下方 DEFAULT_DB_CONFIG，
并在首次运行时自动生成 db_config.json 方便修改（无需重新编译）。
"""

import os
import sys
import json
import math
from decimal import Decimal
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime
from threading import Thread

import openpyxl
import pymysql

# ---------------------------------------------------------------------------
# 默认数据库配置（首次运行会写入 db_config.json，可直接改文件）
# ---------------------------------------------------------------------------
DEFAULT_DB_CONFIG = {
    "host": "80.238.234.89",
    "port": 6883,
    "user": "wms@roc#rocob",
    "password": "VxY1)eWms#|:@Oh4Hn%k",
    "database": "iscewmsdb",
    "charset": "utf8mb4",
}

TABLE_NAME = "westlake_WorkPrice"
REQUIRED_COLUMNS = ["ownerid", "cateid", "catename", "price1", "price2", "price3"]

# workmode：表中为 varchar(255) NOT NULL，注释为 1/2/3 三种作业模式。
# 设为可选列以兼容旧 Excel；缺列或空值时使用默认值（与库中现有数据一致）。
WORKMODE_COLUMN = "workmode"
DEFAULT_WORKMODE = "1"
WORKMODE_OPTIONS = {"1": "Picking", "2": "Scanning", "3": "Replenishment"}
WORKMODE_ALIASES = {
    "1": "1", "picking": "1", "拣货": "1",
    "2": "2", "scanning": "2", "扫码": "2", "扫描": "2",
    "3": "3", "replenishment": "3", "补货": "3",
}

# 增量导入的业务键与更新范围（表无唯一索引，靠代码比对）
UPSERT_KEY_COLUMNS = ("ownerid", "cateid", "workmode")
UPDATE_COLUMNS = ("price1", "price2", "price3")

FIXED_DBSPLITCODE = 1
FIXED_CREATOR = "sys"
FIXED_MODIFIER = "sys"


# ---------------------------------------------------------------------------
# 配置加载
# ---------------------------------------------------------------------------
def app_base_dir():
    """返回 exe 或脚本所在目录，用于定位 db_config.json。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_db_config():
    cfg_path = os.path.join(app_base_dir(), "db_config.json")
    cfg = dict(DEFAULT_DB_CONFIG)
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            print("Failed to read db_config.json, using defaults:", e)
    else:
        # 首次运行：写出模板，方便后续修改
        try:
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_DB_CONFIG, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("Failed to create db_config.json:", e)
    return cfg


# ---------------------------------------------------------------------------
# 数据清洗
# ---------------------------------------------------------------------------
def to_num(val):
    """将单元格值规范为 int / float / str / None。"""
    if val is None:
        return None
    if isinstance(val, float):
        if math.isnan(val):
            return None
        return int(val) if val.is_integer() else val
    if isinstance(val, int):
        return val
    s = str(val).strip()
    if s == "":
        return None
    # 尝试按数字解析（ownerid/cateid/price 可能是数值型字符串）
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return s


def to_str(val):
    if val is None:
        return None
    if isinstance(val, float) and math.isnan(val):
        return None
    s = str(val).strip()
    return s if s else None


def to_workmode(val):
    """将 Excel 单元格规范为 workmode 字符串（varchar 列）。

    返回 (value, warning)：
      - 数值 1 / 1.0 / 文本 '1' / 'Picking' 等统一映射为 '1'；
      - 空值返回 (None, None)，由调用方回退默认值；
      - 不在 1/2/3 范围内的值保留原字符串，但给出警告。
    """
    if val is None:
        return None, None
    if isinstance(val, float):
        if math.isnan(val):
            return None, None
        s = str(int(val)) if val.is_integer() else str(val)
    else:
        s = str(val).strip()
    if not s:
        return None, None
    normalized = WORKMODE_ALIASES.get(s.lower(), s)
    if normalized not in WORKMODE_OPTIONS:
        return normalized, "workmode='%s' is not one of 1/2/3; value kept as-is" % s
    return normalized, None


def key_value(val):
    """业务键字段统一成字符串比较（Excel 数值 101 与库里 varchar '101' 对齐）。"""
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def make_upsert_key(rec):
    """由记录生成 (ownerid, cateid, workmode) 业务键。"""
    return tuple(key_value(rec[c]) for c in UPSERT_KEY_COLUMNS)


def to_int(val):
    """将数值规范为 int（bigint 列）。返回 (value, warning)。"""
    if val is None:
        return None, None
    if isinstance(val, float):
        if math.isnan(val):
            return None, None
        if not val.is_integer():
            return int(round(val)), "Decimal value %s rounded to %d" % (val, int(round(val)))
        return int(val), None
    if isinstance(val, int):
        return val, None
    s = str(val).strip()
    if s == "":
        return None, None
    try:
        f = float(s)
        if f.is_integer():
            return int(f), None
        return int(round(f)), "Decimal value %s rounded to %d" % (s, int(round(f)))
    except ValueError:
        raise ValueError(
            "Cannot convert '%s' to an integer (price columns must be integers)" % s
        )


def to_decimal(val):
    """将数值规范为 Decimal（price1 为 decimal(12,5) 列，保留小数）。返回 (value, warning)。"""
    if val is None:
        return None, None
    if isinstance(val, float):
        if math.isnan(val):
            return None, None
        return Decimal(str(val)), None
    if isinstance(val, int):
        return Decimal(val), None
    s = str(val).strip()
    if s == "":
        return None, None
    try:
        return Decimal(s), None
    except Exception:
        raise ValueError(
            "Cannot convert '%s' to a number (price1 must be numeric)" % s
        )


# ---------------------------------------------------------------------------
# Excel 读取
# ---------------------------------------------------------------------------
def read_excel(path):
    """读取 Excel，返回 (records, has_workmode)。

    workmode 为可选列：Excel 没有该列时记录 None，由 sanitize_records 回退默认值。
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    try:
        header = next(rows_iter)
    except StopIteration:
        raise ValueError("Excel file is empty: no header row found.")

    col_index = {}
    for i, h in enumerate(header):
        if h is None:
            continue
        col_index[str(h).strip().lower()] = i

    missing = [c for c in REQUIRED_COLUMNS if c not in col_index]
    if missing:
        raise ValueError("Excel is missing required column(s): " + ", ".join(missing))

    idx = {c: col_index[c] for c in REQUIRED_COLUMNS}
    has_workmode = WORKMODE_COLUMN in col_index
    wm_idx = col_index.get(WORKMODE_COLUMN)
    records = []
    for r in rows_iter:
        records.append(
            {
                "ownerid": to_num(r[idx["ownerid"]]),
                "cateid": to_num(r[idx["cateid"]]),
                "catename": to_str(r[idx["catename"]]),
                "price1": to_num(r[idx["price1"]]),
                "price2": to_num(r[idx["price2"]]),
                "price3": to_num(r[idx["price3"]]),
                "workmode": r[wm_idx] if wm_idx is not None and wm_idx < len(r) else None,
            }
        )
    wb.close()
    return records, has_workmode


# ---------------------------------------------------------------------------
# 数据清洗 / 校验（根据 westlake_WorkPrice 真实表结构）
#   - ownerid varchar(25) NOT NULL            : 为空则跳过该行
#   - cateid  varchar(3)  NULL               : 为空写 NULL（注意最多 3 字符）
#   - catename varchar(255) NULL             : 为空写 NULL
#   - price1/2/3 decimal(12,5) NOT NULL default 0 : 保留小数；空写 0
#   - workmode varchar(255) NOT NULL         : 规范为 1/2/3；空写默认 '1'
#   - dbsplitcode / creator / modifier / 时间 : 固定值见常量
# ---------------------------------------------------------------------------
def sanitize_records(records):
    clean = []
    warnings = []
    skipped = 0
    key_rows = {}
    for i, rec in enumerate(records, start=2):  # 第 2 行起为数据
        ownerid = to_str(rec["ownerid"]) if isinstance(rec["ownerid"], str) else rec["ownerid"]
        # ownerid 非空校验（NOT NULL）
        if ownerid is None or (isinstance(ownerid, str) and ownerid == ""):
            warnings.append("Row %d: ownerid is empty, skipped" % i)
            skipped += 1
            continue

        p1, _ = to_decimal(rec["price1"])  # price1 为 decimal 列，保留小数
        p2, _ = to_decimal(rec["price2"])  # price2 为 decimal 列，保留小数
        p3, _ = to_decimal(rec["price3"])  # price3 为 decimal 列，保留小数
        # price 为空时写 0（NOT NULL 且有 default 0）
        p1 = 0 if p1 is None else p1
        p2 = 0 if p2 is None else p2
        p3 = 0 if p3 is None else p3

        # cateid / catename 为空写 NULL
        cateid = rec["cateid"]
        if isinstance(cateid, str) and cateid == "":
            cateid = None
        catename = to_str(rec["catename"])

        # workmode 为 NOT NULL 且无默认值：空值回退默认 '1'（Picking）
        wm, wm_warning = to_workmode(rec.get("workmode"))
        if wm is None:
            wm = DEFAULT_WORKMODE
            warnings.append(
                "Row %d: workmode is empty, using default %s (%s)"
                % (i, wm, WORKMODE_OPTIONS.get(wm, ""))
            )
        elif wm_warning:
            warnings.append("Row %d: %s" % (i, wm_warning))

        # 同一业务键在 Excel 中重复时，后一行的价格会覆盖前一行
        key = make_upsert_key({"ownerid": ownerid, "cateid": cateid, "workmode": wm})
        if key in key_rows:
            warnings.append(
                "Row %d duplicates (ownerid, cateid, workmode) from row %d; "
                "the later price overwrites the earlier one"
                % (i, key_rows[key])
            )
        else:
            key_rows[key] = i

        clean.append(
            {
                "ownerid": ownerid,
                "cateid": cateid,
                "catename": catename,
                "price1": p1,
                "price2": p2,
                "price3": p3,
                "workmode": wm,
            }
        )
    return clean, warnings, skipped


# ---------------------------------------------------------------------------
# 数据库写入
# ---------------------------------------------------------------------------
def write_to_db(cfg, records):
    """按 (ownerid, cateid, workmode) 增量写入：存在只更新价格，不存在则插入。

    表无唯一索引，故先一次性读入表内已有业务键集合，再逐行判断；
    全程在同一个事务里，任何异常都会回滚。不会删除任何数据。
    返回 (inserted, updated, existing_rows)。
    """
    conn = pymysql.connect(
        host=cfg["host"],
        port=int(cfg["port"]),
        user=cfg["user"],
        password=cfg["password"],
        database=cfg["database"],
        charset=cfg.get("charset", "utf8mb4"),
        connect_timeout=15,
        autocommit=False,  # 客户端态；服务端会话仍为 1，事务在下方显式开启
    )
    now = datetime.now()
    insert_sql = (
        "INSERT INTO {table} "
        "(id, ownerid, cateid, catename, price1, price2, price3, "
        "dbsplitcode, creator, modifier, createTime, modifytime, workmode) "
        "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    ).format(table=TABLE_NAME)
    # 只更新价格列：catename / modifier / modifytime / createTime 等一律不动
    set_clause = ", ".join("%s=%%s" % c for c in UPDATE_COLUMNS)
    update_sql = "UPDATE {table} SET {set_clause} WHERE {where}"

    inserted = 0
    updated = 0
    existing_rows = 0
    try:
        with conn.cursor() as cur:
            # OceanBase 会话默认 autocommit=1，而 PyMySQL 的 autocommit(False) 在此
            # 不会下发 SET（驱动误判状态），故显式关自动提交并开启事务，保证整批可回滚。
            cur.execute("SET autocommit = 0")  # 仅作用于本连接
            conn.begin()

            # 1. 读入表内已有的业务键
            cur.execute(
                "SELECT ownerid, cateid, workmode FROM {table}".format(table=TABLE_NAME)
            )
            existing = cur.fetchall()
            existing_rows = len(existing)
            existing_keys = {
                (key_value(r[0]), key_value(r[1]), key_value(r[2])) for r in existing
            }

            # 2. id 列没有 AUTO_INCREMENT，新行由代码端生成
            cur.execute(
                "SELECT COALESCE(MAX(id), 0) FROM {table}".format(table=TABLE_NAME)
            )
            next_id = cur.fetchone()[0] + 1

            for rec in records:
                key = make_upsert_key(rec)
                if key in existing_keys:
                    # cateid 允许为 NULL，用 IS NULL 避免 NULL = NULL 匹配不到
                    if rec["cateid"] is None:
                        where = "ownerid=%s AND cateid IS NULL AND workmode=%s"
                        key_params = (
                            key_value(rec["ownerid"]),
                            key_value(rec["workmode"]),
                        )
                    else:
                        where = "ownerid=%s AND cateid=%s AND workmode=%s"
                        # 键统一按字符串比较，与 varchar 列一致，避免隐式类型转换误匹配
                        key_params = (
                            key_value(rec["ownerid"]),
                            key_value(rec["cateid"]),
                            key_value(rec["workmode"]),
                        )
                    cur.execute(
                        update_sql.format(
                            table=TABLE_NAME, set_clause=set_clause, where=where
                        ),
                        tuple(rec[c] for c in UPDATE_COLUMNS) + key_params,
                    )
                    updated += 1
                else:
                    cur.execute(
                        insert_sql,
                        (
                            next_id,
                            rec["ownerid"],
                            rec["cateid"],
                            rec["catename"],
                            rec["price1"],
                            rec["price2"],
                            rec["price3"],
                            FIXED_DBSPLITCODE,
                            FIXED_CREATOR,
                            FIXED_MODIFIER,
                            now,
                            now,
                            rec["workmode"],
                        ),
                    )
                    next_id += 1
                    inserted += 1
                    existing_keys.add(key)  # 同一 Excel 内重复键的后续行走更新
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()
    return inserted, updated, existing_rows


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class ImporterApp:
    def __init__(self, root):
        self.root = root
        self.cfg = load_db_config()
        self.selected_file = None

        root.title("WorkPrice Excel Import Tool")
        root.geometry("640x480")
        root.resizable(True, True)

        # 连接信息（不显示用户名/密码）
        info = "Target: {host}:{port} / {database}".format(
            host=self.cfg["host"],
            port=self.cfg["port"],
            database=self.cfg["database"],
        )
        tk.Label(root, text=info, fg="#555", anchor="w").pack(
            fill="x", padx=12, pady=(10, 0)
        )

        # 按钮区
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=12, pady=10)
        self.btn_select = tk.Button(
            btn_frame,
            text="1. Select Excel File & Import",
            command=self.on_select,
            height=2,
            bg="#2d7ff9",
            fg="white",
        )
        self.btn_select.pack(side="left", padx=(0, 10))
        tk.Button(
            btn_frame, text="Clear Log", command=self.clear_log, height=2
        ).pack(side="left")

        # 已选文件
        self.file_var = tk.StringVar(value="No file selected")
        tk.Label(root, textvariable=self.file_var, anchor="w", fg="#333").pack(
            fill="x", padx=12
        )

        # 日志
        self.log = scrolledtext.ScrolledText(root, wrap="word", height=18)
        self.log.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        self.log.insert(
            "end",
            "Ready. Click the button above to select an Excel file.\n"
            "Required columns: %s\n"
            "Optional column: workmode (1=Picking 2=Scanning 3=Replenishment; default %s)\n"
            "Mode: match on (ownerid, cateid, workmode) - update prices if found, "
            "insert if not. No rows are deleted.\n\n"
            % (", ".join(REQUIRED_COLUMNS), DEFAULT_WORKMODE),
        )

        # 状态栏
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(
            root, textvariable=self.status_var, anchor="w", relief="sunken"
        ).pack(fill="x", side="bottom")

    def log_line(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def clear_log(self):
        self.log.delete("1.0", "end")

    def on_select(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx"), ("All Files (*.*)", "*.*")],
        )
        if not path:
            return
        # 增量导入前确认写入策略（不再清空表）
        ok = messagebox.askyesno(
            "Confirm Incremental Import",
            "Rows are matched on (ownerid, cateid, workmode):\n"
            "  - Existing -> update price1/price2/price3 only\n"
            "  - Not found -> insert a new row\n\n"
            "No existing data in table %s will be deleted.\n\nContinue?" % TABLE_NAME,
        )
        if not ok:
            self.log_line("Cancelled (no changes made).")
            return
        self.selected_file = path
        self.file_var.set("Selected: " + path)
        self.log_line("Selected file: " + path)
        self.btn_select.config(state="disabled")
        self.status_var.set("Working...")
        Thread(target=self.run_import, args=(path,), daemon=True).start()

    def run_import(self, path):
        try:
            self.log_line("Reading Excel ...")
            records, has_workmode = read_excel(path)
            self.log_line("Read %d row(s)." % len(records))
            if has_workmode:
                self.log_line(
                    "workmode column found (1=Picking 2=Scanning 3=Replenishment)."
                )
            else:
                self.log_line(
                    "workmode column not found; using default %s (%s) for all rows."
                    % (DEFAULT_WORKMODE, WORKMODE_OPTIONS.get(DEFAULT_WORKMODE, ""))
                )
            if not records:
                self.log_line("No rows to import.")
                self.finish("Done: 0 rows (no data)")
                return

            self.log_line("Validating and cleaning data ...")
            clean, warnings, skipped = sanitize_records(records)
            for w in warnings:
                self.log_line("  [WARN] " + w)
            self.log_line(
                "%d valid row(s), %d skipped, %d warning(s)."
                % (len(clean), skipped, len(warnings))
            )

            if not clean:
                self.log_line("No rows to import.")
                self.finish("Done: 0 rows (no valid data)")
                return

            self.log_line("Reading existing rows and matching business keys ...")
            inserted, updated, existing_rows = write_to_db(self.cfg, clean)
            self.log_line(
                "Table %s had %d row(s); %d inserted, %d updated (prices only). "
                "No rows deleted."
                % (TABLE_NAME, existing_rows, inserted, updated)
            )
            self.finish(
                "Import complete: %d inserted / %d updated (%d skipped)"
                % (inserted, updated, skipped)
            )
        except Exception as e:
            self.log_line("Error: " + str(e))
            self.root.after(0, lambda: messagebox.showerror("Import Failed", str(e)))
            self.finish("Failed")

    def finish(self, status):
        self.status_var.set(status)
        self.btn_select.config(state="normal")


def main():
    root = tk.Tk()
    ImporterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
