from __future__ import annotations

import csv
import json
import re
import shutil
import sqlite3
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import pandas as pd
from PySide6.QtCore import QDate, Qt, Signal, QSettings
from PySide6.QtGui import QAction, QColor
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QCheckBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QStatusBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QProgressBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "production.db"
SETTINGS_PATH = DATA_DIR / "settings.json"

STATUSES = ["未手配", "発注済", "納入済", "入庫済", "出庫済"]
ORDER_STATES = ["計画", "進行中", "完了", "保留"]

BOM_HEADER_ALIASES = {
    "production_no": ["製番", "製造番号", "工番", "jobno", "jobnumber", "job_no"],
    "part_no": ["品番", "部品番号", "部品コード", "partno", "partnumber", "part_no"],
    "drawing_no": ["図番", "図面番号", "図面No", "drawingno", "drawingnumber", "drawing_no"],
    "material": ["材質", "材料", "素材", "material"],
    "notes": ["備考", "注記", "メモ", "notes", "remarks"],
    "part_name": ["品名", "部品名", "部品名称", "description", "partname"],
    "assembly_process": ["組立工程", "組立工程名", "組立工程番号", "組立工程No", "組立工程NO", "組立No", "組立NO", "組立番号", "工程", "工程名", "工程番号", "工程No", "工程NO", "組立ライン", "assemblyprocess", "assembly_process", "process"],
    "quantity": ["数量", "個数", "必要数", "qty", "quantity"],
    "vendor": ["発注先", "仕入先", "ws", "vendor", "supplier"],
    "due_date": ["納期", "回答納期", "部品納期", "duedate", "due_date", "delivery"],
    "issue_date": ["出庫予定日", "払出予定日", "issuedate", "issue_date"],
    "status": ["状態", "ステータス", "status", "progress"],
}


PROCUREMENT_HEADER_ALIASES = {
    "production_no": ["製番", "製造番号", "工番", "jobno", "jobnumber", "job_no"],
    "assembly_process": ["組立工程", "工程", "組立ライン", "assemblyprocess"],
    "part_no": ["品番", "部品番号", "partno", "partnumber", "part_no"],
    "process_order": ["工程順", "加工順", "順番", "processorder", "sequence", "seq"],
    "procurement_type": ["手配区分", "区分", "手配種別", "type", "procurementtype"],
    "processing_details": ["加工内容", "工程内容", "処理内容", "加工", "processingdetails", "operation"],
    "vendor": ["発注先", "仕入先", "加工先", "ws", "vendor", "supplier"],
    "due_date": ["納期", "回答納期", "工程納期", "duedate", "due_date", "delivery"],
    "issue_date": ["出庫予定日", "払出予定日", "issuedate", "issue_date"],
    "actual_received_date": ["実績入荷日", "入荷日", "受入日", "receiveddate", "actualreceiveddate"],
    "status": ["状態", "ステータス", "status", "progress"],
    "notes": ["備考", "注記", "メモ", "notes", "remarks"],
}

VENDOR_MASTER_HEADER_ALIASES = {
    "vendor_code": ["発注先コード", "仕入先コード", "加工先コード", "vendorcode", "suppliercode", "code"],
    "vendor_name": ["発注先名称", "発注先名", "仕入先名称", "加工先名称", "vendorname", "suppliername", "name"],
    "notes": ["備考", "注記", "メモ", "notes", "remarks"],
}

ROUTE_MASTER_HEADER_ALIASES = {
    "part_no": ["品番", "部品番号", "partno", "partnumber", "part_no"],
    "route_name": ["工程マスタ名", "持ち回り名", "ルート名", "route", "routename"],
    "process_order": ["工程順", "加工順", "順番", "processorder", "sequence", "seq"],
    "procurement_type": ["手配区分", "区分", "手配種別", "type", "procurementtype"],
    "processing_details": ["加工内容", "工程内容", "処理内容", "加工", "processingdetails", "operation"],
    "vendor": ["標準発注先", "発注先", "仕入先", "加工先", "vendor", "supplier"],
    "lead_time_days": ["標準LT", "標準リードタイム", "リードタイム日数", "leadtime", "lead_time_days"],
    "notes": ["備考", "注記", "メモ", "notes", "remarks"],
}

ORDER_HEADER_ALIASES = {
    "series": ["シリーズ", "製品シリーズ", "機種", "機種名", "series", "model"],
    "production_no": ["製番", "製造番号", "工番", "jobno", "jobnumber", "job_no"],
    "delivery_user": ["納入先ユーザー", "納入先", "ユーザー", "客先", "顧客名", "エンドユーザー", "deliveryuser", "customer", "enduser"],
    "due_date": ["顧客納期", "納期", "客先納期", "customerduedate", "due_date"],
    "assembly_process": ["組立工程", "工程", "組立ライン", "assemblyprocess", "process"],
    "parts_completion_date": ["部品完成日", "部品終結日", "部品揃い日", "partscompletiondate"],
    "parts_issue_start_date": ["部品出庫開始日", "出庫開始日", "払出開始日", "partsissuestartdate"],
    "parts_issue_end_date": ["部品出庫完了日", "出庫完了日", "払出完了日", "partsissueenddate"],
    "assembly_start_date": ["組立開始日", "組立着手日", "assemblystartdate"],
    "witness_date": ["立会日", "立会予定日", "検収立会日", "witnessdate"],
    "shipping_date": ["出荷予定日", "出荷日", "shippingdate", "shipdate"],
    "priority": ["優先度", "優先順位", "priority"],
    "state": ["状態", "進捗状態", "ステータス", "state", "status"],
    "notes": ["備考", "注記", "メモ", "notes", "remark", "remarks"],
}


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKC", text).strip().lower()
    text = re.sub(r"[\s_\-／/\\:：\(\)（）\[\]【】]+", "", text)
    return text


def normalize_assembly_process(value: Any) -> str:
    """組立工程表記を照合しやすい形へ正規化する。例: No.1 / NO1 / 1 → no1。"""
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKC", text).strip().lower()
    text = text.replace("№", "no")
    text = re.sub(r"[\s_\-./／\\:：()（）\[\]【】]+", "", text)
    # Excelで数値セルとして読まれた 1.0 なども同じ工程として扱う。
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    text = re.sub(r"^(組立工程|組立|工程)", "", text)
    text = re.sub(r"^(no)", "", text)
    return text


def parse_date(value: Any) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text:
        return None
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


EXCEL_DATE_COLUMNS = {
    "顧客納期", "部品完成日", "部品出庫開始日", "部品出庫完了日",
    "組立開始日", "立会日", "出荷予定日", "納期", "出庫予定日", "実績入荷日",
}


def _load_app_settings() -> dict[str, Any]:
    try:
        if SETTINGS_PATH.exists():
            data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {}


def _save_app_settings(settings: dict[str, Any]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_sql_dataframe(conn: sqlite3.Connection, sql: str, params: tuple[Any, ...] = ()) -> pd.DataFrame:
    return pd.read_sql_query(sql, conn, params=params)


def _progress_master(conn: sqlite3.Connection) -> dict[str, int]:
    rows = conn.execute(
        "SELECT value, COALESCE(progress_percent, 0) FROM choice_masters WHERE category='status'"
    ).fetchall()
    result = {str(row[0]): int(row[1] or 0) for row in rows}
    return result or {"未手配": 0, "発注済": 25, "納入済": 65, "入庫済": 85, "出庫済": 100}


def _prepare_excel_frames(db: "Database") -> dict[str, pd.DataFrame]:
    conn = db.conn
    progress_map = _progress_master(conn)

    orders = _read_sql_dataframe(conn, """
        SELECT
            po.series AS 製品シリーズ,
            po.production_no AS 製番,
            po.delivery_user AS 納入先ユーザー,
            po.due_date AS 顧客納期,
            po.assembly_process AS 組立工程,
            po.parts_completion_date AS 部品完成日,
            po.parts_issue_start_date AS 部品出庫開始日,
            po.parts_issue_end_date AS 部品出庫完了日,
            po.assembly_start_date AS 組立開始日,
            po.witness_date AS 立会日,
            po.shipping_date AS 出荷予定日,
            po.priority AS 優先度,
            po.state AS 状態,
            po.notes AS 備考,
            COUNT(DISTINCT pt.id) AS 部品点数,
            COUNT(pr.id) AS 手配工程数
        FROM production_orders po
        LEFT JOIN parts pt ON pt.production_order_id=po.id
        LEFT JOIN procurements pr ON pr.part_id=pt.id
        GROUP BY po.id
        ORDER BY po.production_no, po.assembly_process
    """)

    parts = _read_sql_dataframe(conn, """
        SELECT
            po.series AS 製品シリーズ,
            po.production_no AS 製番,
            po.delivery_user AS 納入先ユーザー,
            po.assembly_process AS 組立工程,
            pt.part_no AS 品番,
            pt.part_name AS 品名,
            pt.quantity AS 数量,
            pt.drawing_no AS 図番,
            pt.material AS 材質,
            pt.notes AS 備考,
            COUNT(pr.id) AS 手配工程数,
            SUM(CASE WHEN pr.status='出庫済' THEN 1 ELSE 0 END) AS 完了工程数
        FROM parts pt
        JOIN production_orders po ON po.id=pt.production_order_id
        LEFT JOIN procurements pr ON pr.part_id=pt.id
        GROUP BY pt.id
        ORDER BY po.production_no, po.assembly_process, pt.part_no
    """)

    procurements = _read_sql_dataframe(conn, """
        SELECT
            po.series AS 製品シリーズ,
            po.production_no AS 製番,
            po.delivery_user AS 納入先ユーザー,
            po.assembly_process AS 組立工程,
            pt.part_no AS 品番,
            pt.part_name AS 品名,
            pt.quantity AS 手配個数,
            pt.drawing_no AS 図番,
            pt.material AS 材質,
            pr.process_order AS 工程順,
            pr.procurement_type AS 手配区分,
            pr.processing_details AS 加工内容,
            pr.vendor AS 発注先,
            pr.due_date AS 納期,
            pr.issue_date AS 出庫予定日,
            pr.actual_received_date AS 実績入荷日,
            pr.status AS ステータス,
            po.parts_completion_date AS 部品完成日,
            po.parts_issue_start_date AS 部品出庫開始日,
            po.parts_issue_end_date AS 部品出庫完了日,
            pr.notes AS 備考
        FROM procurements pr
        JOIN parts pt ON pt.id=pr.part_id
        JOIN production_orders po ON po.id=pt.production_order_id
        ORDER BY po.production_no, po.assembly_process, pt.part_no, pr.process_order
    """)

    today = pd.Timestamp(date.today())
    if not procurements.empty:
        procurements["進捗率"] = procurements["ステータス"].map(progress_map).fillna(0).astype(int)
        judgements: list[str] = []
        for _, row in procurements.iterrows():
            status = str(row.get("ステータス") or "")
            due = pd.to_datetime(row.get("納期"), errors="coerce")
            targets = [
                pd.to_datetime(row.get("部品完成日"), errors="coerce"),
                pd.to_datetime(row.get("部品出庫開始日"), errors="coerce"),
            ]
            targets = [v for v in targets if not pd.isna(v)]
            target = min(targets) if targets else pd.NaT
            if progress_map.get(status, 0) >= 100 or status == "出庫済":
                judgement = "完了"
            elif status == "未手配":
                judgement = "未手配"
            elif pd.isna(due):
                judgement = "納期未設定"
            elif due < today:
                judgement = "納期超過"
            elif not pd.isna(target) and due > target:
                judgement = "計画目標超過"
            else:
                judgement = "正常"
            judgements.append(judgement)
        procurements["納期判定"] = judgements
        cols = list(procurements.columns)
        # 判定と進捗をステータスの直後へ配置する。
        for name in ("進捗率", "納期判定"):
            cols.remove(name)
        status_pos = cols.index("ステータス") + 1
        cols[status_pos:status_pos] = ["進捗率", "納期判定"]
        procurements = procurements[cols]

    route_master = _read_sql_dataframe(conn, """
        SELECT
            rm.part_no AS 品番,
            rm.route_name AS 工程マスタ名,
            rms.process_order AS 工程順,
            rms.procurement_type AS 手配区分,
            rms.processing_details AS 加工内容,
            rms.vendor AS 標準発注先,
            rms.lead_time_days AS 標準リードタイム日数,
            rms.notes AS 工程備考,
            rm.notes AS マスタ備考
        FROM route_masters rm
        LEFT JOIN route_master_steps rms ON rms.route_master_id=rm.id
        ORDER BY rm.part_no, rm.route_name, rms.process_order
    """)
    vendors = _read_sql_dataframe(conn, """
        SELECT vendor_code AS 発注先コード, vendor_name AS 発注先名称, notes AS 備考
        FROM vendor_masters ORDER BY vendor_code
    """)
    processing = _read_sql_dataframe(conn, """
        SELECT value AS 加工内容, sort_order AS 表示順
        FROM choice_masters WHERE category='processing_details' ORDER BY sort_order, value
    """)
    procurement_types = _read_sql_dataframe(conn, """
        SELECT value AS 手配区分, sort_order AS 表示順
        FROM choice_masters WHERE category='procurement_type' ORDER BY sort_order, value
    """)
    statuses = _read_sql_dataframe(conn, """
        SELECT value AS ステータス, sort_order AS 表示順, progress_percent AS 進捗率
        FROM choice_masters WHERE category='status' ORDER BY sort_order, value
    """)

    attention = procurements[
        procurements.get("納期判定", pd.Series(dtype=str)).isin(
            ["未手配", "納期未設定", "納期超過", "計画目標超過"]
        )
    ].copy() if not procurements.empty else procurements.copy()

    return {
        "生産計画": orders,
        "部品表": parts,
        "手配": procurements,
        "要対応一覧": attention,
        "品番工程マスタ": route_master,
        "発注先マスタ": vendors,
        "加工内容マスタ": processing,
        "手配区分マスタ": procurement_types,
        "ステータスマスタ": statuses,
    }


def _safe_table_name(index: int) -> str:
    return f"PMExportTable{index}"


def _style_export_workbook(writer: pd.ExcelWriter, frames: dict[str, pd.DataFrame], exported_at: datetime) -> None:
    wb = writer.book
    navy = "1F4E78"
    pale_blue = "D9EAF7"
    white = "FFFFFF"
    border_color = "D9E2F3"
    thin = Side(style="thin", color=border_color)
    header_fill = PatternFill("solid", fgColor=navy)
    header_font = Font(color=white, bold=True)

    # サマリー
    ws = wb["サマリー"]
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"
    ws["A1"] = "生産管理 共有スナップショット"
    ws["A1"].font = Font(size=18, bold=True, color=navy)
    ws.merge_cells("A1:F1")
    ws["A2"] = "出力日時"
    ws["B2"] = exported_at
    ws["B2"].number_format = "yyyy/mm/dd hh:mm:ss"
    ws["A3"] = "データ基準"
    ws["B3"] = "出力時点のSQLiteデータ（参照用）"
    ws["A4"] = "注意"
    ws["B4"] = "このExcelを編集してもアプリのデータには反映されません。"
    ws.merge_cells("B3:F3")
    ws.merge_cells("B4:F4")
    for cell in ("A2", "A3", "A4"):
        ws[cell].font = Font(bold=True, color=navy)

    orders = frames["生産計画"]
    procurements = frames["手配"]
    attention = frames["要対応一覧"]
    kpis = [
        ("生産計画", len(orders)),
        ("部品", len(frames["部品表"])),
        ("手配工程", len(procurements)),
        ("要対応", len(attention)),
    ]
    for col, (label, value) in enumerate(kpis, start=1):
        c1 = ws.cell(row=6, column=col)
        c2 = ws.cell(row=7, column=col)
        c1.value = label
        c2.value = value
        c1.fill = header_fill; c1.font = header_font; c1.alignment = Alignment(horizontal="center")
        c2.fill = PatternFill("solid", fgColor=pale_blue); c2.font = Font(size=16, bold=True, color=navy)
        c2.alignment = Alignment(horizontal="center")
        c1.border = c2.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A10"] = "収録シート"
    ws["A10"].fill = header_fill; ws["A10"].font = header_font
    for r, name in enumerate(frames.keys(), start=11):
        ws.cell(r, 1, name)
        ws.cell(r, 1).hyperlink = f"#'{name}'!A1"
        ws.cell(r, 1).style = "Hyperlink"
        ws.cell(r, 2, f"{len(frames[name]):,} 件")

    # ステータス件数表とグラフ
    if not procurements.empty and "ステータス" in procurements:
        counts = procurements["ステータス"].fillna("(空欄)").value_counts()
        ws["H2"] = "手配ステータス"
        ws["I2"] = "件数"
        for cell in ws[2][7:9]:
            cell.fill = header_fill; cell.font = header_font
        for r, (status, count) in enumerate(counts.items(), start=3):
            ws.cell(r, 8, status); ws.cell(r, 9, int(count))
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "手配ステータス件数"
        chart.y_axis.title = "件数"
        data = Reference(ws, min_col=9, min_row=2, max_row=2 + len(counts))
        cats = Reference(ws, min_col=8, min_row=3, max_row=2 + len(counts))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 7.5
        chart.width = 12
        ws.add_chart(chart, "H10")

    ws.column_dimensions["A"].width = 22
    for col in range(2, 7): ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions["H"].width = 18; ws.column_dimensions["I"].width = 12

    # 一覧シート
    for table_index, (sheet_name, df) in enumerate(frames.items(), start=1):
        ws = wb[sheet_name]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        if ws.max_column > 0:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            ws.row_dimensions[1].height = 28
        if len(df) > 0 and ws.max_column > 0:
            ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            tab = Table(displayName=_safe_table_name(table_index), ref=ref)
            tab.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
                showRowStripes=True, showColumnStripes=False
            )
            ws.add_table(tab)
        for col_idx, column_name in enumerate(df.columns, start=1):
            values = [str(column_name)] + ["" if pd.isna(v) else str(v) for v in df[column_name].head(300)]
            max_len = max((len(v) for v in values), default=8)
            width = min(max(max_len + 2, 10), 34)
            ws.column_dimensions[get_column_letter(col_idx)].width = width
            if column_name in EXCEL_DATE_COLUMNS:
                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row_idx, col_idx).number_format = "yyyy/mm/dd"
            if column_name in {"備考", "工程備考", "マスタ備考"}:
                ws.column_dimensions[get_column_letter(col_idx)].width = 34
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=False)
                cell.border = Border(bottom=thin)
        headers = {cell.value: cell.column for cell in ws[1]}
        if "納期判定" in headers and ws.max_row >= 2:
            col = get_column_letter(headers["納期判定"])
            rng = f"{col}2:{col}{ws.max_row}"
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{col}2="納期超過"'], fill=PatternFill("solid", fgColor="F4CCCC")))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{col}2="未手配"'], fill=PatternFill("solid", fgColor="FCE5CD")))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{col}2="計画目標超過"'], fill=PatternFill("solid", fgColor="FFF2CC")))
        if "進捗率" in headers:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row_idx, headers["進捗率"]).number_format = '0"%"'


def export_excel_snapshot(parent: QWidget, db: "Database") -> None:
    settings = _load_app_settings()
    initial = settings.get("excel_export_folder") or str(APP_DIR)
    folder = QFileDialog.getExistingDirectory(parent, "Excel共有ファイルの出力先を選択", initial)
    if not folder:
        return

    exported_at = datetime.now()
    filename = f"生産管理_{exported_at.strftime('%Y%m%d_%H%M%S')}.xlsx"
    destination = Path(folder) / filename
    suffix = 1
    while destination.exists():
        destination = Path(folder) / f"生産管理_{exported_at.strftime('%Y%m%d_%H%M%S')}_{suffix}.xlsx"
        suffix += 1

    work_dir = DATA_DIR / "excel_export_work"
    work_dir.mkdir(parents=True, exist_ok=True)
    temp_path = work_dir / destination.name

    try:
        frames = _prepare_excel_frames(db)
        # 日付はExcelの日付型で出力する。
        for df in frames.values():
            for column in set(df.columns).intersection(EXCEL_DATE_COLUMNS):
                df[column] = pd.to_datetime(df[column], errors="coerce")

        with pd.ExcelWriter(temp_path, engine="openpyxl", datetime_format="yyyy/mm/dd") as writer:
            pd.DataFrame().to_excel(writer, sheet_name="サマリー", index=False)
            for sheet_name, df in frames.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            _style_export_workbook(writer, frames, exported_at)

        # 共有フォルダには完成済みファイルだけをコピーする。
        shutil.copy2(temp_path, destination)
        temp_path.unlink(missing_ok=True)
        settings["excel_export_folder"] = str(Path(folder))
        _save_app_settings(settings)
        QMessageBox.information(
            parent,
            "Excel共有出力完了",
            f"参照用Excelを出力しました。\n\n{destination}\n\n"
            "タイムスタンプ付きの新規ファイルとして保存しています。",
        )
    except Exception as exc:
        QMessageBox.critical(
            parent,
            "Excel共有出力エラー",
            f"Excelファイルを出力できませんでした。\n\n{exc}\n\n"
            f"作業用ファイルが残っている場合: {temp_path}",
        )


class Database:
    def __init__(self, path: Path):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        is_new_database = not path.exists()
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.create_schema()
        # 一括クリア後の再起動でサンプルデータが復活しないよう、初回作成時だけ投入する。
        if is_new_database:
            self.seed_sample_data()

    def create_schema(self) -> None:
        self.conn.executescript(
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS import_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                imported_at TEXT NOT NULL,
                source_file TEXT NOT NULL,
                row_count INTEGER NOT NULL,
                notes TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS archive_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                scope TEXT NOT NULL,
                label TEXT NOT NULL,
                archived_at TEXT NOT NULL,
                record_count INTEGER NOT NULL DEFAULT 0,
                payload_json TEXT NOT NULL,
                restored_at TEXT
            );
            """
        )
        self.conn.commit()
        self.migrate_phase1_schema()
        self.ensure_order_delivery_user_column()

    def ensure_order_delivery_user_column(self) -> None:
        """既存DBへ納入先ユーザー列を安全に追加する。"""
        columns = {row[1] for row in self.conn.execute("PRAGMA table_info(production_orders)")}
        if "delivery_user" not in columns:
            self.conn.execute(
                "ALTER TABLE production_orders ADD COLUMN delivery_user TEXT NOT NULL DEFAULT ''"
            )
            self.conn.commit()

    def migrate_phase1_schema(self) -> None:
        """Ver.1.0α Phase 1: 生産計画→部品→手配の三階層へ移行する。"""
        tables = {
            row[0] for row in self.conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }

        # 初回作成
        if "production_orders" not in tables:
            self._create_phase1_tables()
            return

        order_sql_row = self.conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='production_orders'"
        ).fetchone()
        order_sql = (order_sql_row[0] or "") if order_sql_row else ""
        needs_order_migration = "UNIQUE(production_no, assembly_process)" not in order_sql.replace("\n", " ")
        needs_part_migration = "parts" not in tables or "procurements" not in tables

        if not needs_order_migration and not needs_part_migration:
            self._create_phase1_tables()
            return

        self.conn.execute("PRAGMA foreign_keys = OFF")
        try:
            self.conn.execute("BEGIN")

            # 旧テーブル名を一時退避
            if "bom_items" in tables:
                self.conn.execute("ALTER TABLE bom_items RENAME TO bom_items_phase1_old")
            if needs_order_migration:
                self.conn.execute("ALTER TABLE production_orders RENAME TO production_orders_phase1_old")

            self._create_phase1_tables(commit=False)

            if needs_order_migration:
                self.conn.execute(
                    """
                    INSERT INTO production_orders (
                        id, series, production_no, delivery_user, due_date, assembly_process,
                        parts_completion_date, parts_issue_start_date, parts_issue_end_date,
                        assembly_start_date, witness_date, shipping_date,
                        priority, state, notes, created_at, updated_at
                    )
                    SELECT id, series, production_no, '', due_date,
                           COALESCE(NULLIF(assembly_process, ''), 'No.1'),
                           parts_completion_date, parts_issue_start_date, parts_issue_end_date,
                           assembly_start_date, witness_date, shipping_date,
                           priority, state, notes, created_at, updated_at
                    FROM production_orders_phase1_old
                    """
                )

            # 旧BOMの1行を、部品1件＋初期手配1件として移行
            old_exists = self.conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='bom_items_phase1_old'"
            ).fetchone()
            if old_exists:
                now = datetime.now().isoformat(timespec="seconds")
                old_rows = self.conn.execute("SELECT * FROM bom_items_phase1_old").fetchall()
                for row in old_rows:
                    order = self.conn.execute(
                        "SELECT id FROM production_orders WHERE production_no=? AND assembly_process=?",
                        (row["production_no"], row["assembly_process"] or "No.1"),
                    ).fetchone()
                    if not order:
                        continue
                    cur = self.conn.execute(
                        """
                        INSERT OR IGNORE INTO parts (
                            production_order_id, part_no, part_name, quantity,
                            drawing_no, material, notes, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, '', '', '', ?, ?)
                        """,
                        (order["id"], row["part_no"], row["part_name"],
                         int(round(row["quantity"] or 1)), now, now),
                    )
                    part = self.conn.execute(
                        "SELECT id FROM parts WHERE production_order_id=? AND part_no=?",
                        (order["id"], row["part_no"]),
                    ).fetchone()
                    if part:
                        self.conn.execute(
                            """
                            INSERT INTO procurements (
                                part_id, process_order, procurement_type, processing_details,
                                vendor, due_date, issue_date, status, notes, created_at, updated_at
                            ) VALUES (?, 10, '加工', '従来手配', ?, ?, ?, ?, '', ?, ?)
                            """,
                            (part["id"], row["vendor"], row["due_date"], row["issue_date"],
                             row["status"], now, now),
                        )

            for table in ("bom_items_phase1_old", "production_orders_phase1_old"):
                if self.conn.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)
                ).fetchone():
                    self.conn.execute(f"DROP TABLE {table}")

            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise
        finally:
            self.conn.execute("PRAGMA foreign_keys = ON")

    def _create_phase1_tables(self, commit: bool = True) -> None:
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS production_orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                series TEXT NOT NULL,
                production_no TEXT NOT NULL,
                delivery_user TEXT NOT NULL DEFAULT '',
                due_date TEXT,
                assembly_process TEXT NOT NULL DEFAULT 'No.1',
                parts_completion_date TEXT,
                parts_issue_start_date TEXT,
                parts_issue_end_date TEXT,
                assembly_start_date TEXT,
                witness_date TEXT,
                shipping_date TEXT,
                priority INTEGER NOT NULL DEFAULT 3,
                state TEXT NOT NULL DEFAULT '計画',
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(production_no, assembly_process)
            );

            CREATE TABLE IF NOT EXISTS parts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                production_order_id INTEGER NOT NULL,
                part_no TEXT NOT NULL,
                part_name TEXT NOT NULL DEFAULT '',
                quantity INTEGER NOT NULL DEFAULT 1 CHECK(quantity >= 0),
                drawing_no TEXT NOT NULL DEFAULT '',
                material TEXT NOT NULL DEFAULT '',
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(production_order_id, part_no),
                FOREIGN KEY(production_order_id)
                    REFERENCES production_orders(id)
                    ON UPDATE CASCADE ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS procurements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_id INTEGER NOT NULL,
                process_order INTEGER NOT NULL DEFAULT 10,
                procurement_type TEXT NOT NULL DEFAULT '加工',
                processing_details TEXT NOT NULL DEFAULT '',
                vendor TEXT NOT NULL DEFAULT '',
                due_date TEXT,
                issue_date TEXT,
                actual_received_date TEXT,
                status TEXT NOT NULL DEFAULT '未手配',
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(part_id, process_order),
                FOREIGN KEY(part_id)
                    REFERENCES parts(id)
                    ON UPDATE CASCADE ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_orders_prod_process
                ON production_orders(production_no, assembly_process);
            CREATE INDEX IF NOT EXISTS idx_parts_order
                ON parts(production_order_id);
            CREATE INDEX IF NOT EXISTS idx_procurements_part_order
                ON procurements(part_id, process_order);
            CREATE INDEX IF NOT EXISTS idx_procurements_due
                ON procurements(due_date, status);

            CREATE TABLE IF NOT EXISTS route_masters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_no TEXT NOT NULL,
                route_name TEXT NOT NULL DEFAULT '標準工程',
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(part_no, route_name)
            );

            CREATE TABLE IF NOT EXISTS route_master_steps (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                route_master_id INTEGER NOT NULL,
                process_order INTEGER NOT NULL DEFAULT 10,
                procurement_type TEXT NOT NULL DEFAULT '加工',
                processing_details TEXT NOT NULL DEFAULT '',
                vendor TEXT NOT NULL DEFAULT '',
                lead_time_days INTEGER NOT NULL DEFAULT 0 CHECK(lead_time_days >= 0),
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(route_master_id, process_order),
                FOREIGN KEY(route_master_id)
                    REFERENCES route_masters(id)
                    ON UPDATE CASCADE ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_route_master_part
                ON route_masters(part_no, route_name);
            CREATE INDEX IF NOT EXISTS idx_route_steps_order
                ON route_master_steps(route_master_id, process_order);

            CREATE TABLE IF NOT EXISTS vendor_masters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vendor_code TEXT NOT NULL COLLATE NOCASE UNIQUE,
                vendor_name TEXT NOT NULL COLLATE NOCASE UNIQUE,
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS choice_masters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT NOT NULL,
                value TEXT NOT NULL COLLATE NOCASE,
                sort_order INTEGER NOT NULL DEFAULT 100,
                progress_percent INTEGER,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(category, value)
            );
            CREATE INDEX IF NOT EXISTS idx_choice_masters_category ON choice_masters(category, sort_order, value);
            """
        )
        self._seed_editable_masters()
        if commit:
            self.conn.commit()

    def _seed_editable_masters(self) -> None:
        now=datetime.now().isoformat(timespec="seconds")
        defaults={
            "procurement_type":[("材料",10,None),("加工",20,None),("購入品",30,None),("社内加工",40,None),("検査",50,None),("納品",60,None)],
            "status":[("未手配",10,0),("発注済",20,25),("納入済",30,65),("入庫済",40,85),("出庫済",50,100)],
        }
        for category,rows in defaults.items():
            for value,order,progress in rows:self.conn.execute("INSERT OR IGNORE INTO choice_masters(category,value,sort_order,progress_percent,created_at,updated_at) VALUES(?,?,?,?,?,?)",(category,value,order,progress,now,now))
        for category,sql in {
            "processing_details":"SELECT DISTINCT processing_details FROM procurements WHERE TRIM(processing_details)<>''",
            "procurement_type":"SELECT DISTINCT procurement_type FROM procurements WHERE TRIM(procurement_type)<>''",
            "status":"SELECT DISTINCT status FROM procurements WHERE TRIM(status)<>''",
        }.items():
            for row in self.conn.execute(sql):self.conn.execute("INSERT OR IGNORE INTO choice_masters(category,value,sort_order,progress_percent,created_at,updated_at) VALUES(?,?,100,NULL,?,?)",(category,row[0],now,now))
        vendor_rows=self.conn.execute("SELECT vendor FROM procurements WHERE TRIM(vendor)<>'' UNION SELECT vendor FROM route_master_steps WHERE TRIM(vendor)<>'' ORDER BY vendor").fetchall()
        for index,row in enumerate(vendor_rows,start=1):self.conn.execute("INSERT OR IGNORE INTO vendor_masters(vendor_code,vendor_name,notes,created_at,updated_at) VALUES(?,?,?,?,?)",(f"V{index:04d}",str(row[0]).strip(),"既存データから自動登録",now,now))

    def seed_sample_data(self) -> None:
        count = self.conn.execute("SELECT COUNT(*) FROM production_orders").fetchone()[0]
        if count:
            return
        today = datetime.now().isoformat(timespec="seconds")
        rows = [
            ("SeriesA", "1001", "東都工業", "2026-08-25", "No.1", "2026-08-10", "2026-08-11", "2026-08-13", "2026-08-14", "2026-08-20", "2026-08-24", 2, "進行中", ""),
            ("SeriesA", "1002", "北星製作所", "2026-09-05", "No.1", "2026-08-20", "2026-08-21", "2026-08-24", "2026-08-25", "2026-09-01", "2026-09-04", 3, "計画", ""),
            ("SeriesA", "1003", "東都工業", "2026-09-18", "No.1", "2026-09-01", "2026-09-02", "2026-09-04", "2026-09-07", "2026-09-14", "2026-09-17", 3, "計画", ""),
            ("SeriesB", "2001", "中央機械", "2026-08-20", "No.1", "2026-08-06", "2026-08-07", "2026-08-09", "2026-08-10", "2026-08-17", "2026-08-19", 1, "進行中", ""),
            ("SeriesB", "2002", "西日本設備", "2026-09-10", "No.1", "2026-08-25", "2026-08-26", "2026-08-28", "2026-08-31", "2026-09-07", "2026-09-09", 3, "計画", ""),
            ("SeriesB", "2003", "中央機械", "2026-09-25", "No.1", "2026-09-09", "2026-09-10", "2026-09-12", "2026-09-14", "2026-09-21", "2026-09-24", 3, "計画", ""),
        ]
        self.conn.executemany(
            """
            INSERT INTO production_orders (
                series, production_no, delivery_user, due_date, assembly_process,
                parts_completion_date, parts_issue_start_date, parts_issue_end_date,
                assembly_start_date, witness_date, shipping_date,
                priority, state, notes, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [r + (today, today) for r in rows],
        )
        self.conn.commit()

    ARCHIVE_SCOPES = {
        "orders": "生産計画（部品・手配を含む）",
        "parts": "部品表（配下の手配を含む）",
        "procurements": "手配",
        "route_masters": "品番別 持ち回り工程マスタ",
        "vendors": "発注先マスタ",
        "processing_details": "加工内容マスタ",
        "procurement_type": "手配区分マスタ",
        "status": "ステータスマスタ",
    }

    def _rows_as_dicts(self, sql: str, params: tuple = ()) -> list[dict]:
        return [dict(row) for row in self.conn.execute(sql, params).fetchall()]

    def snapshot_scope(self, scope: str) -> dict[str, list[dict]]:
        if scope == "orders":
            return {
                "production_orders": self._rows_as_dicts("SELECT * FROM production_orders ORDER BY id"),
                "parts": self._rows_as_dicts("SELECT * FROM parts ORDER BY id"),
                "procurements": self._rows_as_dicts("SELECT * FROM procurements ORDER BY id"),
            }
        if scope == "parts":
            return {
                "parts": self._rows_as_dicts("SELECT * FROM parts ORDER BY id"),
                "procurements": self._rows_as_dicts("SELECT * FROM procurements ORDER BY id"),
            }
        if scope == "procurements":
            return {"procurements": self._rows_as_dicts("SELECT * FROM procurements ORDER BY id")}
        if scope == "route_masters":
            return {
                "route_masters": self._rows_as_dicts("SELECT * FROM route_masters ORDER BY id"),
                "route_master_steps": self._rows_as_dicts("SELECT * FROM route_master_steps ORDER BY id"),
            }
        if scope == "vendors":
            return {"vendor_masters": self._rows_as_dicts("SELECT * FROM vendor_masters ORDER BY id")}
        if scope in {"processing_details", "procurement_type", "status"}:
            return {
                "choice_masters": self._rows_as_dicts(
                    "SELECT * FROM choice_masters WHERE category=? ORDER BY id", (scope,)
                )
            }
        raise ValueError(f"未対応の対象です: {scope}")

    @staticmethod
    def snapshot_count(payload: dict[str, list[dict]]) -> int:
        return sum(len(rows) for rows in payload.values())

    def clear_scope(self, scope: str) -> int:
        payload = self.snapshot_scope(scope)
        count = self.snapshot_count(payload)
        with self.conn:
            if scope == "orders":
                self.conn.execute("DELETE FROM production_orders")
            elif scope == "parts":
                self.conn.execute("DELETE FROM parts")
            elif scope == "procurements":
                self.conn.execute("DELETE FROM procurements")
            elif scope == "route_masters":
                self.conn.execute("DELETE FROM route_masters")
            elif scope == "vendors":
                self.conn.execute("DELETE FROM vendor_masters")
            elif scope in {"processing_details", "procurement_type", "status"}:
                self.conn.execute("DELETE FROM choice_masters WHERE category=?", (scope,))
            else:
                raise ValueError(f"未対応の対象です: {scope}")
        return count

    def archive_scope(self, scope: str, label: str = "") -> tuple[int, int]:
        payload = self.snapshot_scope(scope)
        count = self.snapshot_count(payload)
        if count == 0:
            return 0, 0
        now = datetime.now().isoformat(timespec="seconds")
        display = self.ARCHIVE_SCOPES.get(scope, scope)
        with self.conn:
            cur = self.conn.execute(
                """
                INSERT INTO archive_batches(scope,label,archived_at,record_count,payload_json)
                VALUES(?,?,?,?,?)
                """,
                (scope, label.strip() or display, now, count,
                 json.dumps(payload, ensure_ascii=False, separators=(",", ":"))),
            )
            archive_id = int(cur.lastrowid)
            # スナップショット保存後に対象を空にする。
            if scope == "orders": self.conn.execute("DELETE FROM production_orders")
            elif scope == "parts": self.conn.execute("DELETE FROM parts")
            elif scope == "procurements": self.conn.execute("DELETE FROM procurements")
            elif scope == "route_masters": self.conn.execute("DELETE FROM route_masters")
            elif scope == "vendors": self.conn.execute("DELETE FROM vendor_masters")
            elif scope in {"processing_details", "procurement_type", "status"}:
                self.conn.execute("DELETE FROM choice_masters WHERE category=?", (scope,))
        return archive_id, count

    def list_archives(self, scope: str | None = None) -> list[sqlite3.Row]:
        if scope:
            return self.conn.execute(
                "SELECT * FROM archive_batches WHERE scope=? ORDER BY id DESC", (scope,)
            ).fetchall()
        return self.conn.execute("SELECT * FROM archive_batches ORDER BY id DESC").fetchall()

    def _insert_snapshot_rows(self, table: str, rows: list[dict]) -> None:
        if not rows:
            return
        columns = list(rows[0].keys())
        placeholders = ",".join("?" for _ in columns)
        sql = f"INSERT INTO {table} ({','.join(columns)}) VALUES ({placeholders})"
        self.conn.executemany(sql, [tuple(row.get(c) for c in columns) for row in rows])

    def restore_archive(self, archive_id: int) -> tuple[str, int]:
        archive = self.conn.execute("SELECT * FROM archive_batches WHERE id=?", (archive_id,)).fetchone()
        if not archive:
            raise ValueError("指定されたアーカイブが見つかりません。")
        scope = archive["scope"]
        payload = json.loads(archive["payload_json"])
        # 復帰は対象範囲をアーカイブ時点の内容へ置換する。IDと親子関係も復元する。
        self.conn.execute("PRAGMA foreign_keys = OFF")
        try:
            self.conn.execute("BEGIN")
            if scope == "orders":
                self.conn.execute("DELETE FROM procurements")
                self.conn.execute("DELETE FROM parts")
                self.conn.execute("DELETE FROM production_orders")
                self._insert_snapshot_rows("production_orders", payload.get("production_orders", []))
                self._insert_snapshot_rows("parts", payload.get("parts", []))
                self._insert_snapshot_rows("procurements", payload.get("procurements", []))
            elif scope == "parts":
                self.conn.execute("DELETE FROM procurements")
                self.conn.execute("DELETE FROM parts")
                self._insert_snapshot_rows("parts", payload.get("parts", []))
                self._insert_snapshot_rows("procurements", payload.get("procurements", []))
            elif scope == "procurements":
                self.conn.execute("DELETE FROM procurements")
                self._insert_snapshot_rows("procurements", payload.get("procurements", []))
            elif scope == "route_masters":
                self.conn.execute("DELETE FROM route_master_steps")
                self.conn.execute("DELETE FROM route_masters")
                self._insert_snapshot_rows("route_masters", payload.get("route_masters", []))
                self._insert_snapshot_rows("route_master_steps", payload.get("route_master_steps", []))
            elif scope == "vendors":
                self.conn.execute("DELETE FROM vendor_masters")
                self._insert_snapshot_rows("vendor_masters", payload.get("vendor_masters", []))
            elif scope in {"processing_details", "procurement_type", "status"}:
                self.conn.execute("DELETE FROM choice_masters WHERE category=?", (scope,))
                self._insert_snapshot_rows("choice_masters", payload.get("choice_masters", []))
            else:
                raise ValueError(f"未対応の対象です: {scope}")
            self.conn.execute(
                "UPDATE archive_batches SET restored_at=? WHERE id=?",
                (datetime.now().isoformat(timespec="seconds"), archive_id),
            )
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise
        finally:
            self.conn.execute("PRAGMA foreign_keys = ON")
        return scope, int(archive["record_count"])

    def list_orders(self) -> list[sqlite3.Row]:
        return self.conn.execute(
            """
            SELECT p.*,
                COUNT(DISTINCT pt.id) AS part_count,
                SUM(CASE WHEN pr.status='未手配' THEN 1 ELSE 0 END) AS unarranged_count,
                SUM(CASE WHEN pr.status='出庫済' THEN 1 ELSE 0 END) AS issued_count,
                SUM(CASE
                    WHEN pr.due_date IS NOT NULL
                     AND p.parts_completion_date IS NOT NULL
                     AND pr.due_date > p.parts_completion_date
                    THEN 1 ELSE 0 END) AS late_part_count
            FROM production_orders p
            LEFT JOIN parts pt ON pt.production_order_id = p.id
            LEFT JOIN procurements pr ON pr.part_id = pt.id
            GROUP BY p.id
            ORDER BY p.due_date, p.production_no, p.assembly_process
            """
        ).fetchall()

    def get_order(self, order_id: int) -> sqlite3.Row | None:
        return self.conn.execute(
            "SELECT * FROM production_orders WHERE id=?", (order_id,)
        ).fetchone()

    def save_order(self, data: dict[str, Any], order_id: int | None = None) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        fields = [
            "series", "production_no", "delivery_user", "due_date", "assembly_process",
            "parts_completion_date", "parts_issue_start_date", "parts_issue_end_date",
            "assembly_start_date", "witness_date", "shipping_date",
            "priority", "state", "notes"
        ]
        if order_id is None:
            sql = f"""
                INSERT INTO production_orders ({",".join(fields)}, created_at, updated_at)
                VALUES ({",".join(["?"] * len(fields))}, ?, ?)
            """
            self.conn.execute(sql, [data[f] for f in fields] + [now, now])
        else:
            assignments = ",".join(f"{f}=?" for f in fields)
            self.conn.execute(
                f"UPDATE production_orders SET {assignments}, updated_at=? WHERE id=?",
                [data[f] for f in fields] + [now, order_id],
            )
        self.conn.commit()

    def upsert_order_rows(
        self, rows: list[dict[str, Any]], source_file: str
    ) -> tuple[int, int, list[str]]:
        now = datetime.now().isoformat(timespec="seconds")
        added = 0
        updated = 0
        errors: list[str] = []
        date_fields = [
            "due_date", "parts_completion_date", "parts_issue_start_date",
            "parts_issue_end_date", "assembly_start_date", "witness_date",
            "shipping_date",
        ]
        for i, row in enumerate(rows, start=1):
            series = str(row.get("series") or "").strip()
            production_no = str(row.get("production_no") or "").strip()
            if not series or not production_no:
                errors.append(f"{i}行目: 製品シリーズまたは製番が空です")
                continue

            priority_value = row.get("priority")
            try:
                priority = int(float(priority_value)) if priority_value not in (None, "") else 3
            except (TypeError, ValueError):
                errors.append(f"{i}行目: 優先度を数値として解釈できないため3を使用しました")
                priority = 3
            if priority < 1 or priority > 5:
                errors.append(f"{i}行目: 優先度が範囲外のため1～5に補正しました")
                priority = min(5, max(1, priority))

            state = str(row.get("state") or "計画").strip()
            if state not in ORDER_STATES:
                errors.append(f"{i}行目: 状態 '{state}' は未定義のため「計画」を使用しました")
                state = "計画"

            values = {
                "series": series,
                "production_no": production_no,
                "delivery_user": str(row.get("delivery_user") or "").strip(),
                "assembly_process": str(row.get("assembly_process") or "").strip(),
                "priority": priority,
                "state": state,
                "notes": str(row.get("notes") or "").strip(),
            }
            for field in date_fields:
                raw_value = row.get(field)
                parsed = parse_date(raw_value)
                if raw_value not in (None, "") and parsed is None:
                    errors.append(f"{i}行目: {field} の日付を解釈できないため未設定にしました")
                values[field] = parsed

            exists = self.conn.execute(
                "SELECT 1 FROM production_orders WHERE production_no=? AND assembly_process=?",
                (production_no, values["assembly_process"] or "No.1"),
            ).fetchone()
            self.conn.execute(
                """
                INSERT INTO production_orders (
                    series, production_no, delivery_user, due_date, assembly_process,
                    parts_completion_date, parts_issue_start_date, parts_issue_end_date,
                    assembly_start_date, witness_date, shipping_date,
                    priority, state, notes, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(production_no, assembly_process) DO UPDATE SET
                    series=excluded.series,
                    delivery_user=excluded.delivery_user,
                    due_date=excluded.due_date,
                    assembly_process=excluded.assembly_process,
                    parts_completion_date=excluded.parts_completion_date,
                    parts_issue_start_date=excluded.parts_issue_start_date,
                    parts_issue_end_date=excluded.parts_issue_end_date,
                    assembly_start_date=excluded.assembly_start_date,
                    witness_date=excluded.witness_date,
                    shipping_date=excluded.shipping_date,
                    priority=excluded.priority,
                    state=excluded.state,
                    notes=excluded.notes,
                    updated_at=excluded.updated_at
                """,
                (
                    values["series"], values["production_no"], values["delivery_user"], values["due_date"],
                    values["assembly_process"], values["parts_completion_date"],
                    values["parts_issue_start_date"], values["parts_issue_end_date"],
                    values["assembly_start_date"], values["witness_date"],
                    values["shipping_date"], values["priority"], values["state"],
                    values["notes"], now, now,
                ),
            )
            if exists:
                updated += 1
            else:
                added += 1

        self.conn.execute(
            "INSERT INTO import_history(imported_at, source_file, row_count, notes) VALUES (?, ?, ?, ?)",
            (now, source_file, added + updated, "生産計画取り込み\n" + "\n".join(errors[:20])),
        )
        self.conn.commit()
        return added, updated, errors

    def delete_order(self, order_id: int) -> None:
        self.conn.execute("DELETE FROM production_orders WHERE id=?", (order_id,))
        self.conn.commit()

    def list_bom(self, production_no: str | None = None) -> list[sqlite3.Row]:
        # Phase 1の互換表示: 各部品について工程順が最も若い手配を一覧へ表示する。
        sql = """
            SELECT pt.id,
                   po.production_no,
                   po.assembly_process,
                   pt.part_no,
                   pt.part_name,
                   pt.quantity,
                   COALESCE(pr.vendor, '') AS vendor,
                   pr.due_date,
                   pr.issue_date,
                   COALESCE(pr.status, '未手配') AS status,
                   po.parts_completion_date AS target_completion_date,
                   po.parts_issue_start_date AS target_issue_start_date,
                   CASE
                       WHEN pr.id IS NULL THEN '手配未登録'
                       WHEN pr.due_date IS NULL THEN '納期未設定'
                       WHEN po.parts_completion_date IS NOT NULL
                        AND pr.due_date > po.parts_completion_date THEN '部品集結に遅延'
                       WHEN po.parts_issue_start_date IS NOT NULL
                        AND pr.due_date > po.parts_issue_start_date THEN '出庫開始に遅延'
                       ELSE '正常'
                   END AS due_judgement
            FROM parts pt
            JOIN production_orders po ON po.id = pt.production_order_id
            LEFT JOIN procurements pr ON pr.id = (
                SELECT p2.id FROM procurements p2
                WHERE p2.part_id = pt.id
                ORDER BY p2.process_order, p2.id LIMIT 1
            )
        """
        params: tuple[Any, ...] = ()
        if production_no:
            sql += " WHERE po.production_no=?"
            params = (production_no,)
        sql += " ORDER BY po.production_no, po.assembly_process, pt.part_no"
        return self.conn.execute(sql, params).fetchall()

    def upsert_bom_rows(self, rows: list[dict[str, Any]], source_file: str) -> tuple[int, list[str]]:
        """従来BOM取込を部品＋初期手配へ振り分ける互換インポート。"""
        now = datetime.now().isoformat(timespec="seconds")
        count = 0
        errors: list[str] = []
        for i, row in enumerate(rows, start=1):
            prod = str(row.get("production_no", "")).strip()
            process = str(row.get("assembly_process", "")).strip()
            part_no = str(row.get("part_no", "")).strip()
            if not prod or not process or not part_no:
                errors.append(f"{i}行目: 製番・組立工程・品番のいずれかが空です")
                continue
            order = self.conn.execute(
                "SELECT id, assembly_process FROM production_orders WHERE production_no=? AND assembly_process=?",
                (prod, process),
            ).fetchone()
            # 表記揺れを吸収する。例: 1 / No1 / No.1 / 組立No.1。
            if not order:
                candidates = self.conn.execute(
                    "SELECT id, assembly_process FROM production_orders WHERE production_no=?",
                    (prod,),
                ).fetchall()
                process_key = normalize_assembly_process(process)
                matched = [r for r in candidates if normalize_assembly_process(r["assembly_process"]) == process_key]
                if len(matched) == 1:
                    order = matched[0]
                elif len(matched) > 1:
                    names = ", ".join(str(r["assembly_process"]) for r in matched)
                    errors.append(
                        f"{i}行目: 製番 {prod} の組立工程 {process} は複数候補に一致します（{names}）"
                    )
                    continue
            if not order:
                registered = self.conn.execute(
                    "SELECT assembly_process FROM production_orders WHERE production_no=? ORDER BY assembly_process",
                    (prod,),
                ).fetchall()
                choices = ", ".join(str(r["assembly_process"]) for r in registered) or "登録なし"
                errors.append(
                    f"{i}行目: 製番 {prod} × 組立工程 {process} が生産計画にありません。登録済み工程: {choices}"
                )
                continue
            try:
                quantity = int(float(row.get("quantity") or 1))
            except (TypeError, ValueError):
                errors.append(f"{i}行目: 数量を整数として解釈できません")
                continue
            if quantity < 0:
                errors.append(f"{i}行目: 数量は0以上で入力してください")
                continue

            # BOM再取り込みで組立工程が変更された場合は、既存部品を新しい
            # 生産計画（製番 × 組立工程）へ移動する。部品IDを維持するため、
            # 配下の持ち回り手配工程もそのまま引き継がれる。
            part = self.conn.execute(
                "SELECT id FROM parts WHERE production_order_id=? AND part_no=?",
                (order["id"], part_no),
            ).fetchone()
            if not part:
                existing_parts = self.conn.execute(
                    """
                    SELECT pt.id, po.assembly_process
                    FROM parts pt
                    JOIN production_orders po ON po.id = pt.production_order_id
                    WHERE po.production_no=? AND pt.part_no=?
                    ORDER BY po.assembly_process, pt.id
                    """,
                    (prod, part_no),
                ).fetchall()
                if len(existing_parts) == 1:
                    part = existing_parts[0]
                    self.conn.execute(
                        "UPDATE parts SET production_order_id=?, updated_at=? WHERE id=?",
                        (order["id"], now, part["id"]),
                    )
                elif len(existing_parts) > 1:
                    processes = ", ".join(str(r["assembly_process"]) for r in existing_parts)
                    errors.append(
                        f"{i}行目: 製番 {prod}・品番 {part_no} が複数の組立工程に登録済みのため、"
                        f"更新先を特定できません（{processes}）"
                    )
                    continue

            if part:
                self.conn.execute(
                    """
                    UPDATE parts SET
                        part_name=?, quantity=?, drawing_no=?, material=?, notes=?, updated_at=?
                    WHERE id=?
                    """,
                    (str(row.get("part_name") or "").strip(), quantity,
                     str(row.get("drawing_no") or "").strip(),
                     str(row.get("material") or "").strip(),
                     str(row.get("notes") or "").strip(), now, part["id"]),
                )
            else:
                cur = self.conn.execute(
                    """
                    INSERT INTO parts (
                        production_order_id, part_no, part_name, quantity,
                        drawing_no, material, notes, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (order["id"], part_no, str(row.get("part_name") or "").strip(),
                     quantity, str(row.get("drawing_no") or "").strip(),
                     str(row.get("material") or "").strip(),
                     str(row.get("notes") or "").strip(), now, now),
                )
                part = {"id": cur.lastrowid}

            status = str(row.get("status") or "未手配").strip()
            if status not in STATUSES:
                status = "未手配"
            # 従来BOMの発注情報は工程順10の「初期手配」として保持する。
            self.conn.execute(
                """
                INSERT INTO procurements (
                    part_id, process_order, procurement_type, processing_details,
                    vendor, due_date, issue_date, status, notes, created_at, updated_at
                ) VALUES (?, 10, '加工', '初期手配', ?, ?, ?, ?, '', ?, ?)
                ON CONFLICT(part_id, process_order) DO UPDATE SET
                    vendor=excluded.vendor,
                    due_date=excluded.due_date,
                    issue_date=excluded.issue_date,
                    status=excluded.status,
                    updated_at=excluded.updated_at
                """,
                (part["id"], str(row.get("vendor") or "").strip(),
                 parse_date(row.get("due_date")), parse_date(row.get("issue_date")),
                 status, now, now),
            )
            count += 1

        self.conn.execute(
            "INSERT INTO import_history(imported_at, source_file, row_count, notes) VALUES (?, ?, ?, ?)",
            (now, source_file, count, "Phase 1 BOM互換取り込み\n" + "\n".join(errors[:20])),
        )
        self.conn.commit()
        return count, errors

    def list_procurements(self, part_id: int | None = None) -> list[sqlite3.Row]:
        sql = """
            SELECT pr.*, pt.part_no, pt.part_name, pt.quantity,
                   po.production_no, po.assembly_process,
                   po.parts_completion_date AS target_completion_date,
                   po.parts_issue_start_date AS target_issue_start_date
            FROM procurements pr
            JOIN parts pt ON pt.id = pr.part_id
            JOIN production_orders po ON po.id = pt.production_order_id
        """
        params: tuple[Any, ...] = ()
        if part_id is not None:
            sql += " WHERE pr.part_id=?"
            params = (part_id,)
        sql += " ORDER BY po.production_no, po.assembly_process, pt.part_no, pr.process_order"
        return self.conn.execute(sql, params).fetchall()


    def list_parts(self, production_no: str | None = None) -> list[sqlite3.Row]:
        sql = """
            SELECT pt.id, po.production_no, po.assembly_process,
                   pt.part_no, pt.part_name, pt.quantity,
                   pt.drawing_no, pt.material, pt.notes,
                   COUNT(pr.id) AS procurement_count,
                   SUM(CASE WHEN pr.status='出庫済' THEN 1 ELSE 0 END) AS completed_count
            FROM parts pt
            JOIN production_orders po ON po.id=pt.production_order_id
            LEFT JOIN procurements pr ON pr.part_id=pt.id
        """
        params: tuple[Any, ...] = ()
        if production_no:
            sql += " WHERE po.production_no=?"
            params = (production_no,)
        sql += " GROUP BY pt.id ORDER BY po.production_no, po.assembly_process, pt.part_no"
        return self.conn.execute(sql, params).fetchall()

    def get_procurement(self, procurement_id: int) -> sqlite3.Row | None:
        return self.conn.execute(
            "SELECT * FROM procurements WHERE id=?", (procurement_id,)
        ).fetchone()

    def save_procurement(self, data: dict[str, Any], procurement_id: int | None = None) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        fields = [
            "part_id", "process_order", "procurement_type", "processing_details",
            "vendor", "due_date", "issue_date", "actual_received_date",
            "status", "notes"
        ]
        values = [data.get(f) for f in fields]
        if procurement_id is None:
            self.conn.execute(
                f"INSERT INTO procurements ({','.join(fields)}, created_at, updated_at) "
                f"VALUES ({','.join(['?'] * len(fields))}, ?, ?)",
                values + [now, now],
            )
        else:
            assignments = ",".join(f"{f}=?" for f in fields)
            self.conn.execute(
                f"UPDATE procurements SET {assignments}, updated_at=? WHERE id=?",
                values + [now, procurement_id],
            )
        self.conn.commit()

    def delete_procurement(self, procurement_id: int) -> None:
        self.conn.execute("DELETE FROM procurements WHERE id=?", (procurement_id,))
        self.conn.commit()

    def part_choices(self) -> list[sqlite3.Row]:
        return self.conn.execute(
            """
            SELECT pt.id, po.production_no, po.assembly_process, pt.part_no, pt.part_name
            FROM parts pt JOIN production_orders po ON po.id=pt.production_order_id
            ORDER BY po.production_no, po.assembly_process, pt.part_no
            """
        ).fetchall()

    def list_vendors(self) -> list[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM vendor_masters ORDER BY vendor_code COLLATE NOCASE, vendor_name COLLATE NOCASE").fetchall()

    def vendor_display_values(self) -> list[str]:
        return [f'{r["vendor_code"]} {r["vendor_name"]}' for r in self.list_vendors()]

    def save_vendor(self,vendor_code:str,vendor_name:str,notes:str="",vendor_id:int|None=None)->int:
        code=vendor_code.strip(); name=vendor_name.strip()
        if not code or not name:raise ValueError("発注先コードと発注先名称は必須です。")
        now=datetime.now().isoformat(timespec="seconds")
        if vendor_id is None:
            cur=self.conn.execute("INSERT INTO vendor_masters(vendor_code,vendor_name,notes,created_at,updated_at) VALUES(?,?,?,?,?)",(code,name,notes.strip(),now,now)); vendor_id=cur.lastrowid
        else:self.conn.execute("UPDATE vendor_masters SET vendor_code=?,vendor_name=?,notes=?,updated_at=? WHERE id=?",(code,name,notes.strip(),now,vendor_id))
        self.conn.commit(); return int(vendor_id)

    def delete_vendor(self,vendor_id:int)->None:self.conn.execute("DELETE FROM vendor_masters WHERE id=?",(vendor_id,)); self.conn.commit()

    def upsert_vendor_rows(self,rows:list[dict[str,Any]],source_file:str)->tuple[int,int,list[str]]:
        now=datetime.now().isoformat(timespec="seconds"); added=updated=0; errors=[]
        for i,row in enumerate(rows,start=1):
            code=str(row.get("vendor_code") or "").strip(); name=str(row.get("vendor_name") or "").strip(); notes=str(row.get("notes") or "").strip()
            if not code or not name:errors.append(f"{i}行目: 発注先コードまたは発注先名称が空です"); continue
            by_code=self.conn.execute("SELECT id FROM vendor_masters WHERE vendor_code=? COLLATE NOCASE",(code,)).fetchone(); by_name=self.conn.execute("SELECT id FROM vendor_masters WHERE vendor_name=? COLLATE NOCASE",(name,)).fetchone()
            if by_code and by_name and by_code["id"]!=by_name["id"]:errors.append(f"{i}行目: コード「{code}」と名称「{name}」が別々の既存発注先に一致します"); continue
            target=by_code or by_name
            if target:self.conn.execute("UPDATE vendor_masters SET vendor_code=?,vendor_name=?,notes=?,updated_at=? WHERE id=?",(code,name,notes,now,target["id"])); updated+=1
            else:self.conn.execute("INSERT INTO vendor_masters(vendor_code,vendor_name,notes,created_at,updated_at) VALUES(?,?,?,?,?)",(code,name,notes,now,now)); added+=1
        self.conn.execute("INSERT INTO import_history(imported_at,source_file,row_count,notes) VALUES(?,?,?,?)",(now,source_file,added+updated,"発注先マスタ取り込み\n"+"\n".join(errors[:20]))); self.conn.commit(); return added,updated,errors

    def list_choice_master(self,category:str)->list[sqlite3.Row]:return self.conn.execute("SELECT * FROM choice_masters WHERE category=? ORDER BY sort_order,value COLLATE NOCASE",(category,)).fetchall()
    def master_values(self,category:str)->list[str]:return [r["value"] for r in self.list_choice_master(category)]
    def save_choice_master(self,category:str,value:str,sort_order:int=100,progress_percent:int|None=None,item_id:int|None=None)->int:
        value=value.strip()
        if not value:raise ValueError("名称は必須です。")
        now=datetime.now().isoformat(timespec="seconds")
        if item_id is None:
            cur=self.conn.execute("INSERT INTO choice_masters(category,value,sort_order,progress_percent,created_at,updated_at) VALUES(?,?,?,?,?,?)",(category,value,sort_order,progress_percent,now,now)); item_id=cur.lastrowid
        else:self.conn.execute("UPDATE choice_masters SET value=?,sort_order=?,progress_percent=?,updated_at=? WHERE id=?",(value,sort_order,progress_percent,now,item_id))
        self.conn.commit(); return int(item_id)
    def delete_choice_master(self,item_id:int)->None:self.conn.execute("DELETE FROM choice_masters WHERE id=?",(item_id,)); self.conn.commit()
    def master_summary(self)->dict[str,list[str]]:return {"発注先":self.vendor_display_values(),"加工内容":self.master_values("processing_details"),"手配区分":self.master_values("procurement_type"),"ステータス":self.master_values("status")}


    def upsert_procurement_rows(self, rows: list[dict[str, Any]], source_file: str) -> tuple[int, int, list[str]]:
        """製番×組立工程×品番×工程順で手配工程を追加・更新する。"""
        now = datetime.now().isoformat(timespec="seconds")
        added = updated = 0
        errors: list[str] = []
        for i, row in enumerate(rows, start=1):
            prod = str(row.get("production_no") or "").strip()
            assembly = str(row.get("assembly_process") or "").strip()
            part_no = str(row.get("part_no") or "").strip()
            try:
                process_order = int(float(row.get("process_order")))
            except (TypeError, ValueError):
                errors.append(f"{i}行目: 工程順を整数として解釈できません")
                continue
            if not prod or not assembly or not part_no:
                errors.append(f"{i}行目: 製番・組立工程・品番のいずれかが空です")
                continue
            part = self.conn.execute(
                """SELECT pt.id FROM parts pt JOIN production_orders po ON po.id=pt.production_order_id
                   WHERE po.production_no=? AND po.assembly_process=? AND pt.part_no=?""",
                (prod, assembly, part_no),
            ).fetchone()
            if not part:
                errors.append(f"{i}行目: {prod} × {assembly} × {part_no} が部品表にありません")
                continue
            status = str(row.get("status") or "未手配").strip()
            valid_statuses=self.master_values("status") or STATUSES
            if status not in valid_statuses:
                errors.append(f"{i}行目: 状態 '{status}' は未定義のため「未手配」を使用しました")
                status = "未手配"
            exists = self.conn.execute(
                "SELECT id FROM procurements WHERE part_id=? AND process_order=?",
                (part["id"], process_order),
            ).fetchone()
            values = (
                str(row.get("procurement_type") or "加工").strip(),
                str(row.get("processing_details") or "").strip(),
                str(row.get("vendor") or "").strip(),
                parse_date(row.get("due_date")), parse_date(row.get("issue_date")),
                parse_date(row.get("actual_received_date")), status,
                str(row.get("notes") or "").strip(), now,
            )
            if exists:
                self.conn.execute(
                    """UPDATE procurements SET procurement_type=?, processing_details=?, vendor=?,
                       due_date=?, issue_date=?, actual_received_date=?, status=?, notes=?, updated_at=?
                       WHERE id=?""", values + (exists["id"],)
                )
                updated += 1
            else:
                self.conn.execute(
                    """INSERT INTO procurements(part_id, process_order, procurement_type, processing_details,
                       vendor, due_date, issue_date, actual_received_date, status, notes, created_at, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (part["id"], process_order) + values[:-1] + (now, now),
                )
                added += 1
        self.conn.execute(
            "INSERT INTO import_history(imported_at, source_file, row_count, notes) VALUES (?, ?, ?, ?)",
            (now, source_file, added + updated, "手配工程取り込み\n" + "\n".join(errors[:20])),
        )
        self.conn.commit()
        return added, updated, errors

    def list_route_masters(self) -> list[sqlite3.Row]:
        return self.conn.execute(
            """SELECT rm.*, COUNT(rs.id) AS step_count
               FROM route_masters rm LEFT JOIN route_master_steps rs ON rs.route_master_id=rm.id
               GROUP BY rm.id ORDER BY rm.part_no, rm.route_name"""
        ).fetchall()

    def get_route_master(self, master_id: int) -> sqlite3.Row | None:
        return self.conn.execute("SELECT * FROM route_masters WHERE id=?", (master_id,)).fetchone()

    def save_route_master(self, part_no: str, route_name: str, notes: str = "", master_id: int | None = None) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        if master_id is None:
            cur = self.conn.execute(
                "INSERT INTO route_masters(part_no,route_name,notes,created_at,updated_at) VALUES(?,?,?,?,?)",
                (part_no.strip(), route_name.strip() or "標準工程", notes.strip(), now, now),
            )
            master_id = cur.lastrowid
        else:
            self.conn.execute(
                "UPDATE route_masters SET part_no=?,route_name=?,notes=?,updated_at=? WHERE id=?",
                (part_no.strip(), route_name.strip() or "標準工程", notes.strip(), now, master_id),
            )
        self.conn.commit()
        return int(master_id)

    def delete_route_master(self, master_id: int) -> None:
        self.conn.execute("DELETE FROM route_masters WHERE id=?", (master_id,)); self.conn.commit()

    def list_route_steps(self, master_id: int) -> list[sqlite3.Row]:
        return self.conn.execute(
            "SELECT * FROM route_master_steps WHERE route_master_id=? ORDER BY process_order,id", (master_id,)
        ).fetchall()

    def save_route_step(self, data: dict[str, Any], step_id: int | None = None) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        fields = ["route_master_id","process_order","procurement_type","processing_details","vendor","lead_time_days","notes"]
        vals = [data.get(f) for f in fields]
        if step_id is None:
            self.conn.execute(
                f"INSERT INTO route_master_steps({','.join(fields)},created_at,updated_at) VALUES({','.join(['?']*len(fields))},?,?)",
                vals + [now, now],
            )
        else:
            self.conn.execute(
                f"UPDATE route_master_steps SET {','.join(f'{f}=?' for f in fields)},updated_at=? WHERE id=?",
                vals + [now, step_id],
            )
        self.conn.commit()

    def delete_route_step(self, step_id: int) -> None:
        self.conn.execute("DELETE FROM route_master_steps WHERE id=?", (step_id,)); self.conn.commit()

    def upsert_route_master_rows(self, rows: list[dict[str, Any]], source_file: str) -> tuple[int, list[str]]:
        now = datetime.now().isoformat(timespec="seconds")
        count = 0; errors: list[str] = []
        for i,row in enumerate(rows,start=1):
            part_no=str(row.get("part_no") or "").strip(); route_name=str(row.get("route_name") or "標準工程").strip()
            try: process_order=int(float(row.get("process_order")))
            except (TypeError,ValueError): errors.append(f"{i}行目: 工程順を整数として解釈できません"); continue
            if not part_no: errors.append(f"{i}行目: 品番が空です"); continue
            rm=self.conn.execute("SELECT id FROM route_masters WHERE part_no=? AND route_name=?",(part_no,route_name)).fetchone()
            if not rm:
                cur=self.conn.execute("INSERT INTO route_masters(part_no,route_name,notes,created_at,updated_at) VALUES(?,?,?,?,?)",(part_no,route_name,"",now,now)); master_id=cur.lastrowid
            else: master_id=rm["id"]
            try: lead=max(0,int(float(row.get("lead_time_days") or 0)))
            except (TypeError,ValueError): lead=0; errors.append(f"{i}行目: 標準LTを0日として取り込みました")
            self.conn.execute(
                """INSERT INTO route_master_steps(route_master_id,process_order,procurement_type,processing_details,vendor,lead_time_days,notes,created_at,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(route_master_id,process_order) DO UPDATE SET procurement_type=excluded.procurement_type,
                   processing_details=excluded.processing_details,vendor=excluded.vendor,lead_time_days=excluded.lead_time_days,
                   notes=excluded.notes,updated_at=excluded.updated_at""",
                (master_id,process_order,str(row.get("procurement_type") or "加工").strip(),str(row.get("processing_details") or "").strip(),str(row.get("vendor") or "").strip(),lead,str(row.get("notes") or "").strip(),now,now),
            ); count+=1
        self.conn.execute("INSERT INTO import_history(imported_at,source_file,row_count,notes) VALUES(?,?,?,?)",(now,source_file,count,"工程マスタ取り込み\n"+"\n".join(errors[:20])))
        self.conn.commit(); return count,errors

    def matching_route_masters(self, part_no: str) -> list[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM route_masters WHERE part_no=? ORDER BY route_name", (part_no,)).fetchall()

    def apply_route_master(self, part_id: int, master_id: int, replace_existing: bool = False) -> tuple[int,int]:
        now=datetime.now().isoformat(timespec="seconds")
        if replace_existing: self.conn.execute("DELETE FROM procurements WHERE part_id=?",(part_id,))
        added=skipped=0
        for step in self.list_route_steps(master_id):
            exists=self.conn.execute("SELECT 1 FROM procurements WHERE part_id=? AND process_order=?",(part_id,step["process_order"])).fetchone()
            if exists: skipped+=1; continue
            self.conn.execute("""INSERT INTO procurements(part_id,process_order,procurement_type,processing_details,vendor,status,notes,created_at,updated_at)
                VALUES(?,?,?,?,?,'未手配',?,?,?)""",(part_id,step["process_order"],step["procurement_type"],step["processing_details"],step["vendor"],step["notes"],now,now)); added+=1
        self.conn.commit(); return added,skipped

    def progress_tree_rows(self) -> list[sqlite3.Row]:
        return self.conn.execute(
            """
            SELECT po.id AS order_id, po.series, po.production_no, po.delivery_user, po.assembly_process,
                   po.due_date AS order_due_date, po.parts_completion_date,
                   po.parts_issue_start_date, po.state AS order_state,
                   pt.id AS part_id, pt.part_no, pt.part_name, pt.quantity,
                   pr.id AS procurement_id, pr.process_order, pr.procurement_type,
                   pr.processing_details, pr.vendor, pr.due_date AS procurement_due_date,
                   pr.issue_date, pr.actual_received_date, pr.status AS procurement_status
            FROM production_orders po
            LEFT JOIN parts pt ON pt.production_order_id = po.id
            LEFT JOIN procurements pr ON pr.part_id = pt.id
            ORDER BY po.due_date, po.production_no, po.assembly_process, pt.part_no, pr.process_order
            """
        ).fetchall()


class DateField(QDateEdit):
    def __init__(self):
        super().__init__()
        self.setCalendarPopup(True)
        self.setDisplayFormat("yyyy/MM/dd")
        self.setDate(QDate.currentDate())
        self.setSpecialValueText("未設定")
        self.setMinimumDate(QDate(2000, 1, 1))

    def set_iso(self, value: str | None) -> None:
        if value:
            self.setDate(QDate.fromString(value, "yyyy-MM-dd"))
        else:
            self.setDate(self.minimumDate())

    def iso(self) -> str | None:
        if self.date() == self.minimumDate():
            return None
        return self.date().toString("yyyy-MM-dd")


class OrderDialog(QDialog):
    def __init__(self, parent=None, row: sqlite3.Row | None = None):
        super().__init__(parent)
        self.setWindowTitle("生産計画")
        self.resize(520, 620)
        layout = QFormLayout(self)

        self.series = QLineEdit()
        self.production_no = QLineEdit()
        self.delivery_user = QLineEdit()
        self.due_date = DateField()
        self.assembly_process = QLineEdit("No.1")
        self.parts_completion_date = DateField()
        self.parts_issue_start_date = DateField()
        self.parts_issue_end_date = DateField()
        self.assembly_start_date = DateField()
        self.witness_date = DateField()
        self.shipping_date = DateField()
        self.priority = QSpinBox()
        self.priority.setRange(1, 5)
        self.priority.setValue(3)
        self.state = QComboBox()
        self.state.addItems(ORDER_STATES)
        self.notes = QTextEdit()

        widgets = [
            ("製品シリーズ", self.series),
            ("製番", self.production_no),
            ("納入先ユーザー", self.delivery_user),
            ("納期", self.due_date),
            ("組立工程", self.assembly_process),
            ("部品終結日", self.parts_completion_date),
            ("部品出庫開始日", self.parts_issue_start_date),
            ("部品出庫完了日", self.parts_issue_end_date),
            ("組立開始日", self.assembly_start_date),
            ("立会日", self.witness_date),
            ("出荷予定日", self.shipping_date),
            ("優先度（1=最優先）", self.priority),
            ("状態", self.state),
            ("備考", self.notes),
        ]
        for label, widget in widgets:
            layout.addRow(label, widget)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        if row:
            self.series.setText(row["series"])
            self.production_no.setText(row["production_no"])
            self.delivery_user.setText(row["delivery_user"] or "")
            self.due_date.set_iso(row["due_date"])
            self.assembly_process.setText(row["assembly_process"] or "")
            self.parts_completion_date.set_iso(row["parts_completion_date"])
            self.parts_issue_start_date.set_iso(row["parts_issue_start_date"])
            self.parts_issue_end_date.set_iso(row["parts_issue_end_date"])
            self.assembly_start_date.set_iso(row["assembly_start_date"])
            self.witness_date.set_iso(row["witness_date"])
            self.shipping_date.set_iso(row["shipping_date"])
            self.priority.setValue(row["priority"])
            self.state.setCurrentText(row["state"])
            self.notes.setPlainText(row["notes"] or "")

    def validate_and_accept(self):
        if not self.series.text().strip() or not self.production_no.text().strip():
            QMessageBox.warning(self, "入力不足", "製品シリーズと製番は必須です。")
            return
        self.accept()

    def data(self) -> dict[str, Any]:
        return {
            "series": self.series.text().strip(),
            "production_no": self.production_no.text().strip(),
            "delivery_user": self.delivery_user.text().strip(),
            "due_date": self.due_date.iso(),
            "assembly_process": self.assembly_process.text().strip(),
            "parts_completion_date": self.parts_completion_date.iso(),
            "parts_issue_start_date": self.parts_issue_start_date.iso(),
            "parts_issue_end_date": self.parts_issue_end_date.iso(),
            "assembly_start_date": self.assembly_start_date.iso(),
            "witness_date": self.witness_date.iso(),
            "shipping_date": self.shipping_date.iso(),
            "priority": self.priority.value(),
            "state": self.state.currentText(),
            "notes": self.notes.toPlainText().strip(),
        }


@dataclass
class ImportData:
    frame: pd.DataFrame
    header_row: int
    source_file: str


def read_tabular_file(path: str, header_aliases: dict[str, list[str]]) -> ImportData:
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        raw = pd.read_csv(path, header=None, dtype=object, encoding_errors="replace")
    elif suffix == ".xls":
        raw = pd.read_excel(path, header=None, dtype=object, engine="xlrd")
    else:
        raw = pd.read_excel(path, header=None, dtype=object, engine="openpyxl")

    best_row = 0
    best_score = -1
    alias_values = {normalize_header(a) for values in header_aliases.values() for a in values}
    for idx in range(min(20, len(raw))):
        values = [normalize_header(v) for v in raw.iloc[idx].tolist()]
        score = sum(1 for v in values if v in alias_values)
        if score > best_score:
            best_score = score
            best_row = idx

    if suffix == ".csv":
        frame = pd.read_csv(path, header=best_row, dtype=object, encoding_errors="replace")
    elif suffix == ".xls":
        frame = pd.read_excel(path, header=best_row, dtype=object, engine="xlrd")
    else:
        frame = pd.read_excel(path, header=best_row, dtype=object, engine="openpyxl")
    frame = frame.dropna(how="all").reset_index(drop=True)
    return ImportData(frame=frame, header_row=best_row, source_file=path)


class ImportDialog(QDialog):
    def __init__(
        self, import_data: ImportData, target_labels: dict[str, str],
        header_aliases: dict[str, list[str]], required_targets: set[str],
        title: str, parent=None
    ):
        super().__init__(parent)
        self.import_data = import_data
        self.target_labels = {**target_labels, "ignore": "取り込まない"}
        self.header_aliases = header_aliases
        self.required_targets = required_targets
        self.setWindowTitle(title)
        self.resize(980, 650)
        root = QVBoxLayout(self)

        root.addWidget(QLabel(
            f"検出ヘッダー行: {import_data.header_row + 1}行目 / "
            f"データ件数: {len(import_data.frame)}件"
        ))

        self.mapping_table = QTableWidget(len(import_data.frame.columns), 3)
        self.mapping_table.setHorizontalHeaderLabels(["元データ列", "取り込み先", "自動判定"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.combos: list[QComboBox] = []

        used_targets = set()
        for row, col in enumerate(import_data.frame.columns):
            col_text = str(col)
            normalized = normalize_header(col_text)
            guessed = "ignore"
            for target, aliases in self.header_aliases.items():
                if normalized in {normalize_header(a) for a in aliases}:
                    guessed = target
                    break
            if guessed in used_targets and guessed not in ("ignore",):
                guessed = "ignore"
            used_targets.add(guessed)

            self.mapping_table.setItem(row, 0, QTableWidgetItem(col_text))
            combo = QComboBox()
            for key, label in self.target_labels.items():
                combo.addItem(label, key)
            combo.setCurrentIndex(combo.findData(guessed))
            self.mapping_table.setCellWidget(row, 1, combo)
            self.mapping_table.setItem(
                row, 2, QTableWidgetItem("自動" if guessed != "ignore" else "未設定")
            )
            self.combos.append(combo)

        root.addWidget(self.mapping_table)

        preview = QTableWidget(min(10, len(import_data.frame)), len(import_data.frame.columns))
        preview.setHorizontalHeaderLabels([str(c) for c in import_data.frame.columns])
        preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        for r in range(preview.rowCount()):
            for c in range(preview.columnCount()):
                value = import_data.frame.iloc[r, c]
                preview.setItem(r, c, QTableWidgetItem("" if pd.isna(value) else str(value)))
        group = QGroupBox("先頭10件プレビュー")
        gl = QVBoxLayout(group)
        gl.addWidget(preview)
        root.addWidget(group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_mapping)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def validate_mapping(self):
        targets = [combo.currentData() for combo in self.combos]
        missing = self.required_targets - set(targets)
        if missing:
            missing_labels = "、".join(self.target_labels.get(x, x) for x in sorted(missing))
            QMessageBox.warning(self, "列対応不足", f"必須列（{missing_labels}）を割り当ててください。")
            return
        duplicates = {x for x in targets if x != "ignore" and targets.count(x) > 1}
        if duplicates:
            QMessageBox.warning(self, "列対応重複", "同じ取り込み先が複数列に割り当てられています。")
            return
        self.accept()

    def mapped_rows(self) -> list[dict[str, Any]]:
        mappings = {}
        for index, combo in enumerate(self.combos):
            target = combo.currentData()
            if target != "ignore":
                mappings[target] = self.import_data.frame.columns[index]
        rows = []
        for _, source_row in self.import_data.frame.iterrows():
            item = {}
            for target, source_col in mappings.items():
                value = source_row[source_col]
                item[target] = None if pd.isna(value) else value
            rows.append(item)
        return rows


class TargetMappingDialog(QDialog):
    """取り込み先項目を固定表示し、各項目へ元データ列を割り当てる列対応画面。

    BOMのように必須項目を見落としやすい取り込みで、
    「組立工程」を含む全取り込み先を常に画面上へ表示する。
    """
    def __init__(
        self, import_data: ImportData, target_labels: dict[str, str],
        header_aliases: dict[str, list[str]], required_targets: set[str],
        title: str, parent=None
    ):
        super().__init__(parent)
        self.import_data = import_data
        self.target_labels = target_labels
        self.header_aliases = header_aliases
        self.required_targets = required_targets
        self.setWindowTitle(title)
        self.resize(980, 720)
        root = QVBoxLayout(self)

        required_text = "、".join(target_labels[k] for k in target_labels if k in required_targets)
        root.addWidget(QLabel(
            f"検出ヘッダー行: {import_data.header_row + 1}行目 / "
            f"データ件数: {len(import_data.frame)}件\n"
            f"必須項目: {required_text}"
        ))

        self.mapping_table = QTableWidget(len(target_labels), 3)
        self.mapping_table.setHorizontalHeaderLabels(["取り込み先項目", "元データ列", "必須"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.combos: dict[str, QComboBox] = {}

        normalized_columns = {normalize_header(str(c)): c for c in import_data.frame.columns}
        used_columns = set()
        for row, (target, label) in enumerate(target_labels.items()):
            label_item = QTableWidgetItem(label)
            if target in required_targets:
                font = label_item.font(); font.setBold(True); label_item.setFont(font)
            self.mapping_table.setItem(row, 0, label_item)

            combo = QComboBox()
            combo.addItem("（取り込まない）", None)
            for col in import_data.frame.columns:
                combo.addItem(str(col), col)

            guessed_column = None
            for alias in header_aliases.get(target, []):
                candidate = normalized_columns.get(normalize_header(alias))
                if candidate is not None and candidate not in used_columns:
                    guessed_column = candidate
                    break
            if guessed_column is not None:
                combo.setCurrentIndex(combo.findData(guessed_column))
                used_columns.add(guessed_column)

            self.mapping_table.setCellWidget(row, 1, combo)
            self.mapping_table.setItem(row, 2, QTableWidgetItem("必須" if target in required_targets else "任意"))
            self.combos[target] = combo

        root.addWidget(self.mapping_table)

        preview = QTableWidget(min(10, len(import_data.frame)), len(import_data.frame.columns))
        preview.setHorizontalHeaderLabels([str(c) for c in import_data.frame.columns])
        preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        for r in range(preview.rowCount()):
            for c in range(preview.columnCount()):
                value = import_data.frame.iloc[r, c]
                preview.setItem(r, c, QTableWidgetItem("" if pd.isna(value) else str(value)))
        group = QGroupBox("先頭10件プレビュー")
        gl = QVBoxLayout(group); gl.addWidget(preview); root.addWidget(group)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_mapping)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def validate_mapping(self):
        selected = {target: combo.currentData() for target, combo in self.combos.items()}
        missing = [target for target in self.required_targets if selected.get(target) is None]
        if missing:
            labels = "、".join(self.target_labels.get(x, x) for x in missing)
            QMessageBox.warning(self, "列対応不足", f"必須列（{labels}）へ元データ列を割り当ててください。")
            return
        columns = [column for column in selected.values() if column is not None]
        if len(columns) != len(set(columns)):
            QMessageBox.warning(self, "列対応重複", "同じ元データ列が複数の取り込み先へ割り当てられています。")
            return
        self.accept()

    def mapped_rows(self) -> list[dict[str, Any]]:
        mappings = {target: combo.currentData() for target, combo in self.combos.items() if combo.currentData() is not None}
        rows = []
        for _, source_row in self.import_data.frame.iterrows():
            item = {}
            for target, source_col in mappings.items():
                value = source_row[source_col]
                item[target] = None if pd.isna(value) else value
            rows.append(item)
        return rows


class DashboardWidget(QWidget):
    """Phase 5: 手配実績を基に進捗・遅延原因・今日の要対応を見える化する。"""

    procurement_requested = Signal(int)  # 要対応行から進捗ツリーへ移動するための手配工程ID
    bottleneck_requested = Signal(str)  # 加工内容別の要対応工程一覧を開く

    STATUS_PROGRESS = {
        "未手配": 0,
        "発注済": 25,
        "納入済": 65,
        "入庫済": 85,
        "出庫済": 100,
    }

    def __init__(self, db: Database):
        super().__init__()
        self.db = db
        root = QVBoxLayout(self)

        filter_row = QHBoxLayout()
        self.series_filter = QComboBox()
        self.series_filter.addItem("すべて")
        self.state_filter = QComboBox()
        self.state_filter.addItems(["すべて", "正常", "要注意", "遅延"])
        self.search = QLineEdit()
        self.search.setPlaceholderText("製番・工程・品番・加工内容・発注先を検索")
        refresh_btn = QPushButton("更新")
        filter_row.addWidget(QLabel("機種"))
        filter_row.addWidget(self.series_filter)
        filter_row.addWidget(QLabel("判定"))
        filter_row.addWidget(self.state_filter)
        filter_row.addWidget(self.search, 1)
        filter_row.addWidget(refresh_btn)
        root.addLayout(filter_row)

        cards = QHBoxLayout()
        self.card_labels = {}
        for key, title in [
            ("orders", "登録計画"),
            ("today_due", "本日納期"),
            ("today_issue", "本日出庫予定"),
            ("warning", "要注意計画"),
            ("late", "遅延計画"),
            ("unarranged", "未手配工程"),
        ]:
            box = QGroupBox(title)
            box_layout = QVBoxLayout(box)
            label = QLabel("0")
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("font-size: 24px; font-weight: bold;")
            box_layout.addWidget(label)
            cards.addWidget(box)
            self.card_labels[key] = label
        root.addLayout(cards)

        # 上段左右／上中下の境界を独立してドラッグ可能にする。
        # 分割位置はWindowsユーザーごとのQSettingsへ保存し、DBの仕様は変更しない。
        self.dashboard_layout_settings = QSettings("YAMA AYATO", "production_manager")
        self.dashboard_vertical_splitter = QSplitter(Qt.Vertical)
        self.dashboard_vertical_splitter.setHandleWidth(8)
        self.dashboard_vertical_splitter.setChildrenCollapsible(False)
        self.dashboard_horizontal_splitter = QSplitter(Qt.Horizontal)
        self.dashboard_horizontal_splitter.setHandleWidth(8)
        self.dashboard_horizontal_splitter.setChildrenCollapsible(False)
        self.series_table = QTableWidget()
        self.series_table.setColumnCount(6)
        self.series_table.setHorizontalHeaderLabels(
            ["機種", "計画数", "正常", "要注意", "遅延", "手配進捗"]
        )
        self.series_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.series_table.setAlternatingRowColors(True)
        self.series_table.setMinimumHeight(95)

        self.bottleneck_table = QTableWidget()
        self.bottleneck_table.setColumnCount(5)
        self.bottleneck_table.setHorizontalHeaderLabels(
            ["加工内容", "要対応", "未手配", "納期超過", "主な発注先"]
        )
        self.bottleneck_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.bottleneck_table.setAlternatingRowColors(True)
        self.bottleneck_table.setMinimumHeight(95)
        self.bottleneck_table.cellDoubleClicked.connect(self.open_bottleneck_details)

        self.dashboard_horizontal_splitter.addWidget(self.series_table)
        self.dashboard_horizontal_splitter.addWidget(self.bottleneck_table)
        self.dashboard_horizontal_splitter.setStretchFactor(0, 1)
        self.dashboard_horizontal_splitter.setStretchFactor(1, 1)
        self.dashboard_vertical_splitter.addWidget(self.dashboard_horizontal_splitter)

        self.order_table = QTableWidget()
        self.order_table.setColumnCount(11)
        self.order_table.setHorizontalHeaderLabels(
            ["機種", "製番", "納入先ユーザー", "組立工程", "顧客納期", "部品集結", "部品数", "手配工程", "進捗", "判定", "主な停止原因"]
        )
        self.order_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.order_table.setAlternatingRowColors(True)
        self.order_table.setSortingEnabled(True)
        self.order_table.setMinimumHeight(95)
        self.dashboard_vertical_splitter.addWidget(self.order_table)

        self.action_table = QTableWidget()
        self.action_table.setColumnCount(9)
        self.action_table.setHorizontalHeaderLabels(
            ["優先", "製番", "組立工程", "品番", "加工内容", "発注先", "手配納期", "状態", "判定"]
        )
        self.action_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.action_table.setAlternatingRowColors(True)
        self.action_table.setSortingEnabled(True)
        self.action_table.setMinimumHeight(95)
        self.action_table.cellDoubleClicked.connect(self.open_action_procurement)
        action_panel = QWidget()
        action_layout = QVBoxLayout(action_panel)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.addWidget(QLabel("今日の要対応（未手配・納期超過・計画目標超過） ※ダブルクリックで進捗ツリーへ"))
        action_layout.addWidget(self.action_table, 1)
        self.dashboard_vertical_splitter.addWidget(action_panel)

        for section in range(3):
            self.dashboard_vertical_splitter.setStretchFactor(section, 1)
        # 最低限の可読領域を確保し、境界のドラッグで表が消えないようにする。
        self.dashboard_horizontal_splitter.setSizes([1, 1])
        self.dashboard_vertical_splitter.setSizes([2, 3, 3])
        horizontal_state = self.dashboard_layout_settings.value("dashboard/horizontal_splitter")
        vertical_state = self.dashboard_layout_settings.value("dashboard/vertical_splitter")
        if horizontal_state is not None:
            self.dashboard_horizontal_splitter.restoreState(horizontal_state)
        if vertical_state is not None:
            self.dashboard_vertical_splitter.restoreState(vertical_state)
        self.dashboard_horizontal_splitter.splitterMoved.connect(self.save_dashboard_splitters)
        self.dashboard_vertical_splitter.splitterMoved.connect(self.save_dashboard_splitters)
        root.addWidget(self.dashboard_vertical_splitter, 1)

        self.summary = QLabel()
        self.summary.setStyleSheet("font-weight: bold; padding: 4px;")
        root.addWidget(self.summary)

        self.series_filter.currentTextChanged.connect(self.refresh)
        self.state_filter.currentTextChanged.connect(self.refresh)
        self.search.textChanged.connect(self.refresh)
        refresh_btn.clicked.connect(self.refresh)

    def save_dashboard_splitters(self, _position: int, _index: int) -> None:
        """操作した分割位置を保存し、次回起動時にも復元する。"""
        self.dashboard_layout_settings.setValue(
            "dashboard/horizontal_splitter", self.dashboard_horizontal_splitter.saveState()
        )
        self.dashboard_layout_settings.setValue(
            "dashboard/vertical_splitter", self.dashboard_vertical_splitter.saveState()
        )

    def open_action_procurement(self, row_index: int, _column: int) -> None:
        """表示行の順序が変わっても、内部IDで移動先を確定する。"""
        item = self.action_table.item(row_index, 0)
        if item is None:
            return
        procurement_id = item.data(Qt.UserRole)
        if procurement_id is not None:
            self.procurement_requested.emit(int(procurement_id))

    def open_bottleneck_details(self, row_index: int, _column: int) -> None:
        """ソート後も表示行に保持した加工内容で絞り込みを要求する。"""
        item = self.bottleneck_table.item(row_index, 0)
        if item is not None and item.data(Qt.UserRole) is not None:
            self.bottleneck_requested.emit(str(item.data(Qt.UserRole)))

    @classmethod
    def procurement_progress(cls, status):
        return cls.STATUS_PROGRESS.get(status or "未手配", 0)

    @staticmethod
    def make_progress_bar(value):
        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setValue(value)
        bar.setFormat(f"{value}%")
        bar.setMaximumHeight(20)
        return bar

    @staticmethod
    def procurement_judgement(row):
        if not row["procurement_id"]:
            return "手配なし", "darkorange", True
        if row["procurement_status"] == "出庫済":
            return "完了", "green", False
        today = date.today().isoformat()
        due = row["procurement_due_date"]
        target = row["parts_completion_date"] or row["parts_issue_start_date"]
        if not due:
            return "納期未設定", "darkorange", True
        if due < today:
            return "納期超過", "red", True
        if target and due > target:
            return "目標超過", "red", True
        if row["procurement_status"] == "未手配":
            return "未手配", "darkorange", True
        return "進行中", "blue", False

    @staticmethod
    def order_assessment(order_info):
        if order_info["late"] > 0:
            return "遅延", "red"
        if order_info["problems"] > 0:
            return "要注意", "darkorange"
        return "正常", "green"

    def refresh(self):
        order_rows = self.db.list_orders()
        tree_rows = self.db.progress_tree_rows()

        series_values = sorted({r["series"] for r in order_rows})
        current = self.series_filter.currentText()
        self.series_filter.blockSignals(True)
        self.series_filter.clear()
        self.series_filter.addItem("すべて")
        self.series_filter.addItems(series_values)
        self.series_filter.setCurrentText(current if current in series_values else "すべて")
        self.series_filter.blockSignals(False)

        orders = {}
        bottlenecks = {}
        action_rows = []
        today = date.today().isoformat()

        for row in tree_rows:
            oid = row["order_id"]
            info = orders.setdefault(oid, {
                "series": row["series"], "production_no": row["production_no"],
                "delivery_user": row["delivery_user"] or "", "assembly_process": row["assembly_process"], "due_date": row["order_due_date"],
                "parts_completion_date": row["parts_completion_date"],
                "part_ids": set(), "progress": [], "proc_count": 0,
                "problems": 0, "late": 0, "causes": {},
            })
            if row["part_id"] is not None:
                info["part_ids"].add(row["part_id"])
            if row["procurement_id"] is None:
                continue
            info["proc_count"] += 1
            progress = self.procurement_progress(row["procurement_status"])
            info["progress"].append(progress)
            judgement, _color, problem = self.procurement_judgement(row)
            if not problem:
                continue
            info["problems"] += 1
            if judgement in ("納期超過", "目標超過"):
                info["late"] += 1
            detail = row["processing_details"] or "（加工内容未設定）"
            info["causes"][detail] = info["causes"].get(detail, 0) + 1
            b = bottlenecks.setdefault(detail, {"total": 0, "unarranged": 0, "late": 0, "vendors": {}})
            b["total"] += 1
            if judgement in ("未手配", "手配なし", "納期未設定"):
                b["unarranged"] += 1
            if judgement in ("納期超過", "目標超過"):
                b["late"] += 1
            vendor = row["vendor"] or "未設定"
            b["vendors"][vendor] = b["vendors"].get(vendor, 0) + 1
            priority = 1 if judgement == "納期超過" else 2 if judgement == "目標超過" else 3
            action_rows.append((priority, row, judgement))

        for info in orders.values():
            info["progress_value"] = round(sum(info["progress"]) / len(info["progress"])) if info["progress"] else 0
            info["assessment"], info["assessment_color"] = self.order_assessment(info)
            info["main_cause"] = max(info["causes"], key=info["causes"].get) if info["causes"] else "－"

        series_selected = self.series_filter.currentText()
        state_selected = self.state_filter.currentText()
        term = normalize_header(self.search.text())
        filtered_orders = []
        for info in orders.values():
            if series_selected != "すべて" and info["series"] != series_selected:
                continue
            if state_selected != "すべて" and info["assessment"] != state_selected:
                continue
            searchable = " ".join(str(info[k] or "") for k in ("series", "production_no", "assembly_process", "main_cause"))
            if term and term not in normalize_header(searchable):
                # 手配側に検索語がある場合は残す。
                matching = any(
                    r["order_id"] in orders and r["order_id"] == next((oid for oid, x in orders.items() if x is info), -1)
                    and term in normalize_header(" ".join(str(r[k] or "") for k in (
                        "part_no", "part_name", "processing_details", "vendor", "procurement_status"
                    ))) for r in tree_rows
                )
                if not matching:
                    continue
            filtered_orders.append(info)

        # カードは全体状態を表示する。
        all_infos = list(orders.values())
        self.card_labels["orders"].setText(str(len(all_infos)))
        self.card_labels["today_due"].setText(str(sum(1 for x in all_infos if x["due_date"] == today)))
        self.card_labels["today_issue"].setText(str(sum(
            1 for r in tree_rows if r["issue_date"] == today and r["procurement_status"] != "出庫済"
        )))
        self.card_labels["warning"].setText(str(sum(1 for x in all_infos if x["assessment"] == "要注意")))
        self.card_labels["late"].setText(str(sum(1 for x in all_infos if x["assessment"] == "遅延")))
        self.card_labels["unarranged"].setText(str(sum(
            1 for r in tree_rows if r["procurement_id"] and r["procurement_status"] == "未手配"
        )))

        # 機種別進捗。
        series_summary = {}
        for info in all_infos:
            s = series_summary.setdefault(info["series"], {"count": 0, "正常": 0, "要注意": 0, "遅延": 0, "progress": []})
            s["count"] += 1
            s[info["assessment"]] += 1
            s["progress"].append(info["progress_value"])
        self.series_table.setRowCount(len(series_summary))
        for i, (series, s) in enumerate(sorted(series_summary.items())):
            vals = [series, s["count"], s["正常"], s["要注意"], s["遅延"]]
            for c, value in enumerate(vals):
                self.series_table.setItem(i, c, QTableWidgetItem(str(value)))
            progress = round(sum(s["progress"]) / len(s["progress"])) if s["progress"] else 0
            self.series_table.setCellWidget(i, 5, self.make_progress_bar(progress))

        # 遅延原因上位。
        sorted_bottlenecks = sorted(bottlenecks.items(), key=lambda x: (-x[1]["total"], x[0]))[:10]
        self.bottleneck_table.setSortingEnabled(False)
        self.bottleneck_table.setRowCount(len(sorted_bottlenecks))
        for i, (detail, b) in enumerate(sorted_bottlenecks):
            main_vendor = max(b["vendors"], key=b["vendors"].get) if b["vendors"] else "－"
            vals = [detail, b["total"], b["unarranged"], b["late"], main_vendor]
            for c, value in enumerate(vals):
                item = QTableWidgetItem(str(value))
                if c in (1, 2, 3) and int(value) > 0:
                    item.setForeground(QColor("red" if c == 3 else "darkorange"))
                if c == 0:
                    item.setData(Qt.UserRole, detail)
                self.bottleneck_table.setItem(i, c, item)

        self.bottleneck_table.setSortingEnabled(True)

        # 生産計画別進捗。
        self.order_table.setSortingEnabled(False)
        self.order_table.setRowCount(len(filtered_orders))
        for i, info in enumerate(sorted(filtered_orders, key=lambda x: (x["due_date"] or "9999", x["production_no"], x["assembly_process"]))):
            vals = [info["series"], info["production_no"], info.get("delivery_user", ""), info["assembly_process"], info["due_date"] or "",
                    info["parts_completion_date"] or "", len(info["part_ids"]), info["proc_count"]]
            for c, value in enumerate(vals):
                self.order_table.setItem(i, c, QTableWidgetItem(str(value)))
            self.order_table.setCellWidget(i, 8, self.make_progress_bar(info["progress_value"]))
            state_item = QTableWidgetItem(info["assessment"])
            state_item.setForeground(QColor(info["assessment_color"]))
            self.order_table.setItem(i, 9, state_item)
            self.order_table.setItem(i, 10, QTableWidgetItem(info["main_cause"]))
        self.order_table.setSortingEnabled(True)

        # 今日の要対応。優先度→納期→製番の順。
        visible_actions = []
        for priority, row, judgement in action_rows:
            info = orders[row["order_id"]]
            if series_selected != "すべて" and info["series"] != series_selected:
                continue
            searchable = " ".join(str(row[k] or "") for k in (
                "production_no", "delivery_user", "assembly_process", "part_no", "part_name", "processing_details", "vendor", "procurement_status"
            ))
            if term and term not in normalize_header(searchable):
                continue
            visible_actions.append((priority, row, judgement))
        visible_actions.sort(key=lambda x: (x[0], x[1]["procurement_due_date"] or "9999", x[1]["production_no"]))
        visible_actions = visible_actions[:50]
        self.action_table.setSortingEnabled(False)
        self.action_table.setRowCount(len(visible_actions))
        for i, (priority, row, judgement) in enumerate(visible_actions):
            label = "至急" if priority == 1 else "高" if priority == 2 else "確認"
            vals = [label, row["production_no"], row["assembly_process"], row["part_no"] or "",
                    row["processing_details"] or "", row["vendor"] or "", row["procurement_due_date"] or "",
                    row["procurement_status"] or "", judgement]
            for c, value in enumerate(vals):
                item = QTableWidgetItem(str(value))
                if c in (0, 8):
                    item.setForeground(QColor("red" if priority == 1 else "darkorange"))
                if c == 0:
                    item.setData(Qt.UserRole, int(row["procurement_id"]))
                self.action_table.setItem(i, c, item)
        self.action_table.setSortingEnabled(True)

        average = round(sum(x["progress_value"] for x in all_infos) / len(all_infos)) if all_infos else 0
        self.summary.setText(
            f"全体手配進捗 {average}% ／ 要対応工程 {len(action_rows)}件 ／ "
            f"表示中の生産計画 {len(filtered_orders)}件"
        )


class BottleneckDetailsDialog(QDialog):
    """加工内容別の集計を、その根拠となる全要対応手配工程へ展開する。"""

    def __init__(self, db: Database, processing_details: str, parent=None):
        super().__init__(parent)
        self.db = db
        self.processing_details = processing_details
        self.selected_procurement_id = None
        self.setWindowTitle(f"ボトルネック詳細：{processing_details}")
        self.resize(1160, 610)
        layout = QVBoxLayout(self)
        explanation = QLabel(
            f"加工内容：{processing_details}　｜　要対応の手配工程一覧（全製番・全組立工程）\n"
            "工程をダブルクリックすると、該当する製番・組立工程の要対応部品へ移動します。"
        )
        explanation.setWordWrap(True)
        layout.addWidget(explanation)
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "発注先", "加工内容", "判定", "製番", "組立工程", "品番", "部品名", "工程順", "手配納期", "状態"
        ])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.cellDoubleClicked.connect(self.choose_procurement)
        layout.addWidget(self.table, 1)
        self.count_label = QLabel()
        layout.addWidget(self.count_label)
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.reject)
        layout.addWidget(close_button)
        self.refresh()

    def refresh(self) -> None:
        # ダッシュボードの集計と同じ判定基準を再使用する。
        rows = []
        for row in self.db.progress_tree_rows():
            if row["procurement_id"] is None:
                continue
            detail = row["processing_details"] or "（加工内容未設定）"
            if detail != self.processing_details:
                continue
            judgement, _color, is_problem = DashboardWidget.procurement_judgement(row)
            if is_problem:
                rows.append((row, judgement))
        rows.sort(key=lambda pair: (
            0 if pair[1] == "納期超過" else 1 if pair[1] == "目標超過" else 2,
            pair[0]["procurement_due_date"] or "9999",
            pair[0]["production_no"], pair[0]["assembly_process"]
        ))
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(rows))
        for index, (row, judgement) in enumerate(rows):
            vals = [
                row["vendor"] or "未設定", row["processing_details"] or "（加工内容未設定）",
                judgement, row["production_no"], row["assembly_process"],
                row["part_no"] or "", row["part_name"] or "",
                row["process_order"], row["procurement_due_date"] or "",
                row["procurement_status"] or ""
            ]
            for column, value in enumerate(vals):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, int(row["procurement_id"]))
                if column == 2:
                    item.setForeground(QColor("red" if judgement in ("納期超過", "目標超過") else "darkorange"))
                self.table.setItem(index, column, item)
        self.table.setSortingEnabled(True)
        self.count_label.setText(f"要対応工程 {len(rows)}件（工程単位で集計）")

    def choose_procurement(self, row_index: int, _column: int) -> None:
        item = self.table.item(row_index, 0)
        if item is None or item.data(Qt.UserRole) is None:
            return
        self.selected_procurement_id = int(item.data(Qt.UserRole))
        self.accept()


class DataManagementDialog(QDialog):
    """タブ単位の一括アーカイブ・クリア・復帰を行う共通画面。"""

    def __init__(self, db: Database, default_scope: str = "orders", parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("一括データ管理")
        self.resize(760, 520)
        root = QVBoxLayout(self)

        info = QLabel(
            "アーカイブは対象データを復帰可能な状態で保存してから一覧を空にします。\n"
            "一括クリアは復帰できません。復帰時は同じ対象範囲の現在データをアーカイブ時点の内容で置換します。"
        )
        info.setWordWrap(True)
        info.setStyleSheet("padding:8px; background:#fff4ce; border:1px solid #d6b656;")
        root.addWidget(info)

        top = QHBoxLayout()
        top.addWidget(QLabel("対象"))
        self.scope = QComboBox()
        for key, label in db.ARCHIVE_SCOPES.items():
            self.scope.addItem(label, key)
        idx = self.scope.findData(default_scope)
        if idx >= 0:
            self.scope.setCurrentIndex(idx)
        self.label = QLineEdit()
        self.label.setPlaceholderText("アーカイブ名（任意。例：2026年8月締め）")
        top.addWidget(self.scope, 2)
        top.addWidget(self.label, 3)
        root.addLayout(top)

        actions = QHBoxLayout()
        archive_btn = QPushButton("アーカイブして空にする")
        clear_btn = QPushButton("一括クリア")
        restore_btn = QPushButton("選択アーカイブを復帰")
        refresh_btn = QPushButton("履歴更新")
        archive_btn.setStyleSheet("font-weight:bold;")
        clear_btn.setStyleSheet("color:#b00020;")
        actions.addWidget(archive_btn)
        actions.addWidget(clear_btn)
        actions.addStretch()
        actions.addWidget(restore_btn)
        actions.addWidget(refresh_btn)
        root.addLayout(actions)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["ID", "対象", "アーカイブ名", "保存日時", "件数", "最終復帰日時"])
        self.table.setColumnHidden(0, True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        root.addWidget(self.table, 1)

        close = QDialogButtonBox(QDialogButtonBox.Close)
        close.rejected.connect(self.reject)
        root.addWidget(close)

        self.scope.currentIndexChanged.connect(self.refresh)
        archive_btn.clicked.connect(self.archive_current)
        clear_btn.clicked.connect(self.clear_current)
        restore_btn.clicked.connect(self.restore_selected)
        refresh_btn.clicked.connect(self.refresh)
        self.refresh()

    def current_scope(self) -> str:
        return str(self.scope.currentData())

    def refresh(self):
        rows = self.db.list_archives(self.current_scope())
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            vals = [
                row["id"],
                self.db.ARCHIVE_SCOPES.get(row["scope"], row["scope"]),
                row["label"],
                row["archived_at"].replace("T", " "),
                row["record_count"],
                (row["restored_at"] or "").replace("T", " "),
            ]
            for c, value in enumerate(vals):
                item = QTableWidgetItem(str(value))
                if c == 0:
                    item.setData(Qt.UserRole, int(row["id"]))
                self.table.setItem(i, c, item)
        if rows:
            self.table.selectRow(0)

    def archive_current(self):
        scope = self.current_scope()
        name = self.db.ARCHIVE_SCOPES[scope]
        answer = QMessageBox.question(
            self,
            "一括アーカイブ確認",
            f"「{name}」の全データをアーカイブへ保存し、現在の一覧を空にしますか？\n\n"
            "復帰するときは、対象範囲の現在データがアーカイブ内容で置き換わります。",
        )
        if answer != QMessageBox.Yes:
            return
        archive_id, count = self.db.archive_scope(scope, self.label.text())
        if not archive_id:
            QMessageBox.information(self, "一括アーカイブ", "対象データは0件でした。")
            return
        QMessageBox.information(self, "一括アーカイブ", f"{count}件をアーカイブしました。\nアーカイブID: {archive_id}")
        self.label.clear()
        self.refresh()
        self.accept()

    def clear_current(self):
        scope = self.current_scope()
        name = self.db.ARCHIVE_SCOPES[scope]
        text, ok = QInputDialog.getText(
            self,
            "一括クリア確認",
            f"「{name}」を復帰不能な状態で一括削除します。\n実行するには「クリア」と入力してください。",
        )
        if not ok or text.strip() != "クリア":
            return
        count = self.db.clear_scope(scope)
        QMessageBox.information(self, "一括クリア", f"{count}件を削除しました。")
        self.accept()

    def selected_archive_id(self) -> int | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return int(item.text()) if item else None

    def restore_selected(self):
        archive_id = self.selected_archive_id()
        if archive_id is None:
            QMessageBox.information(self, "アーカイブ復帰", "復帰するアーカイブを選択してください。")
            return
        answer = QMessageBox.warning(
            self,
            "アーカイブ復帰確認",
            "復帰すると、同じ対象範囲の現在データは削除され、選択したアーカイブ時点の内容に置き換わります。\n\n実行しますか？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        try:
            scope, count = self.db.restore_archive(archive_id)
        except Exception as exc:
            QMessageBox.critical(self, "復帰エラー", str(exc))
            return
        QMessageBox.information(self, "アーカイブ復帰", f"{self.db.ARCHIVE_SCOPES.get(scope, scope)}を復帰しました。\n保存件数: {count}件")
        self.refresh()
        self.accept()


def open_data_management(parent: QWidget, db: Database, on_change, scope: str):
    dlg = DataManagementDialog(db, scope, parent)
    if dlg.exec():
        on_change()


class OrdersWidget(QWidget):
    def __init__(self, db: Database, on_change):
        super().__init__()
        self.db = db
        self.on_change = on_change
        root = QVBoxLayout(self)

        controls = QHBoxLayout()
        add_btn = QPushButton("新規登録")
        edit_btn = QPushButton("編集")
        delete_btn = QPushButton("削除")
        import_btn = QPushButton("CSV / Excel取り込み")
        bulk_btn = QPushButton("一括アーカイブ／クリア")
        refresh_btn = QPushButton("更新")
        controls.addWidget(add_btn)
        controls.addWidget(edit_btn)
        controls.addWidget(delete_btn)
        controls.addWidget(import_btn)
        controls.addWidget(bulk_btn)
        controls.addStretch()
        controls.addWidget(refresh_btn)
        root.addLayout(controls)

        filter_box = QGroupBox("検索・絞り込み")
        filter_layout = QHBoxLayout(filter_box)
        self.search = QLineEdit()
        self.search.setPlaceholderText("機種・製番・納入先ユーザー・組立工程などを検索")
        self.match_mode = QComboBox()
        self.match_mode.addItems(["部分一致", "完全一致"])
        self.state_filter = QComboBox()
        self.state_filter.addItem("状態: すべて")
        self.state_filter.addItems(ORDER_STATES)
        clear_filter_btn = QPushButton("解除")
        filter_layout.addWidget(QLabel("検索"))
        filter_layout.addWidget(self.search, 1)
        filter_layout.addWidget(self.match_mode)
        filter_layout.addWidget(self.state_filter)
        filter_layout.addWidget(clear_filter_btn)
        root.addWidget(filter_box)

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(
            ["ID", "機種", "製番", "納入先ユーザー", "納期", "組立工程", "部品終結", "出庫開始", "出庫完了", "状態"]
        )
        self.table.setColumnHidden(0, True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        root.addWidget(self.table)

        add_btn.clicked.connect(self.add_order)
        edit_btn.clicked.connect(self.edit_order)
        delete_btn.clicked.connect(self.delete_order)
        import_btn.clicked.connect(self.import_file)
        bulk_btn.clicked.connect(lambda: open_data_management(self, self.db, self.on_change, "orders"))
        refresh_btn.clicked.connect(self.refresh)
        self.search.textChanged.connect(self.apply_filters)
        self.match_mode.currentTextChanged.connect(self.apply_filters)
        self.state_filter.currentTextChanged.connect(self.apply_filters)
        clear_filter_btn.clicked.connect(self.clear_filters)
        self.table.doubleClicked.connect(self.edit_order)

    def clear_filters(self):
        self.search.clear()
        self.match_mode.setCurrentText("部分一致")
        self.state_filter.setCurrentIndex(0)

    def apply_filters(self):
        term = normalize_header(self.search.text())
        exact = self.match_mode.currentText() == "完全一致"
        state = self.state_filter.currentText().replace("状態: ", "")
        for row in range(self.table.rowCount()):
            values = [self.table.item(row, c).text() if self.table.item(row, c) else "" for c in range(1, self.table.columnCount())]
            normalized = [normalize_header(v) for v in values]
            text_ok = not term or (any(v == term for v in normalized) if exact else any(term in v for v in normalized))
            row_state = self.table.item(row, 9).text() if self.table.item(row, 9) else ""
            state_ok = state == "すべて" or row_state == state
            self.table.setRowHidden(row, not (text_ok and state_ok))

    def selected_id(self) -> int | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        return int(self.table.item(row, 0).text())

    def refresh(self):
        rows = self.db.list_orders()
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            vals = [
                r["id"], r["series"], r["production_no"], r["delivery_user"] or "", r["due_date"] or "",
                r["assembly_process"] or "", r["parts_completion_date"] or "",
                r["parts_issue_start_date"] or "", r["parts_issue_end_date"] or "",
                r["state"]
            ]
            for c, v in enumerate(vals):
                self.table.setItem(i, c, QTableWidgetItem(str(v)))
        self.apply_filters()

    def add_order(self):
        dlg = OrderDialog(self)
        if dlg.exec():
            try:
                self.db.save_order(dlg.data())
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "登録エラー", "同じ製番が既に登録されています。")
                return
            self.refresh()
            self.on_change()

    def edit_order(self):
        order_id = self.selected_id()
        if order_id is None:
            QMessageBox.information(self, "選択", "編集する製番を選択してください。")
            return
        row = self.db.get_order(order_id)
        dlg = OrderDialog(self, row)
        if dlg.exec():
            try:
                self.db.save_order(dlg.data(), order_id)
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "更新エラー", "同じ製番が既に登録されています。")
                return
            self.refresh()
            self.on_change()

    def import_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "生産計画ファイルを選択", str(APP_DIR),
            "表形式ファイル (*.csv *.xls *.xlsx)"
        )
        if not path:
            return
        try:
            data = read_tabular_file(path, ORDER_HEADER_ALIASES)
        except Exception as exc:
            QMessageBox.critical(self, "読込エラー", str(exc))
            return
        labels = {
            "series": "製品シリーズ", "production_no": "製番",
            "delivery_user": "納入先ユーザー", "due_date": "顧客納期", "assembly_process": "組立工程",
            "parts_completion_date": "部品完成日",
            "parts_issue_start_date": "部品出庫開始日",
            "parts_issue_end_date": "部品出庫完了日",
            "assembly_start_date": "組立開始日", "witness_date": "立会日",
            "shipping_date": "出荷予定日", "priority": "優先度",
            "state": "状態", "notes": "備考",
        }
        dlg = ImportDialog(
            data, labels, ORDER_HEADER_ALIASES, {"series", "production_no"},
            "生産計画 列対応", self
        )
        if not dlg.exec():
            return
        added, updated, errors = self.db.upsert_order_rows(dlg.mapped_rows(), path)
        message = f"追加 {added}件、更新 {updated}件を取り込みました。"
        if errors:
            message += f"\n\n警告: {len(errors)}件\n" + "\n".join(errors[:10])
        QMessageBox.information(self, "取り込み結果", message)
        self.refresh()
        self.on_change()

    def delete_order(self):
        order_id = self.selected_id()
        if order_id is None:
            return
        if QMessageBox.question(self, "削除確認", "選択した製番と、その部品表を削除しますか？") == QMessageBox.Yes:
            self.db.delete_order(order_id)
            self.refresh()
            self.on_change()



class BomWidget(QWidget):
    HEADERS = ["製番", "組立工程", "品番", "品名", "数量", "図番", "材質", "手配工程数", "完了工程数", "備考"]

    def __init__(self, db: Database, on_change):
        super().__init__()
        self.db = db
        self.on_change = on_change
        root = QVBoxLayout(self)
        controls = QHBoxLayout()
        self.production_filter = QComboBox()
        import_btn = QPushButton("CSV / Excel取り込み")
        bulk_btn = QPushButton("一括アーカイブ／クリア")
        clear_btn = QPushButton("絞り込み解除")
        refresh_btn = QPushButton("更新")
        controls.addWidget(QLabel("製番")); controls.addWidget(self.production_filter)
        controls.addWidget(import_btn); controls.addWidget(bulk_btn); controls.addWidget(clear_btn)
        controls.addStretch(); controls.addWidget(refresh_btn)
        root.addLayout(controls)

        filter_box = QGroupBox("列ごとの検索・絞り込み（部分一致）")
        grid = QGridLayout(filter_box)
        self.column_filters = []
        for i, h in enumerate(self.HEADERS):
            e = QLineEdit(); e.setPlaceholderText(h); e.setClearButtonEnabled(True)
            e.textChanged.connect(self.apply_filters); self.column_filters.append(e)
            grid.addWidget(QLabel(h), (i//5)*2, i%5); grid.addWidget(e, (i//5)*2+1, i%5)
        root.addWidget(filter_box)

        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        root.addWidget(self.table)
        import_btn.clicked.connect(self.import_file); bulk_btn.clicked.connect(lambda: open_data_management(self, self.db, self.on_change, "parts")); clear_btn.clicked.connect(self.clear_filters)
        refresh_btn.clicked.connect(self.refresh); self.production_filter.currentTextChanged.connect(self.refresh)

    def clear_filters(self):
        for e in self.column_filters: e.clear()
        self.production_filter.setCurrentText("すべて")

    def apply_filters(self):
        terms=[normalize_header(e.text()) for e in self.column_filters]
        for r in range(self.table.rowCount()):
            ok=all(not t or t in normalize_header(self.table.item(r,c).text() if self.table.item(r,c) else "") for c,t in enumerate(terms))
            self.table.setRowHidden(r, not ok)

    def refresh(self):
        values=sorted({r["production_no"] for r in self.db.list_orders()}); current=self.production_filter.currentText()
        self.production_filter.blockSignals(True); self.production_filter.clear(); self.production_filter.addItem("すべて"); self.production_filter.addItems(values)
        self.production_filter.setCurrentText(current if current in values else "すべて"); self.production_filter.blockSignals(False)
        selected=self.production_filter.currentText(); rows=self.db.list_parts(None if selected=="すべて" else selected)
        self.table.setSortingEnabled(False); self.table.setRowCount(len(rows))
        for i,r in enumerate(rows):
            vals=[r["production_no"],r["assembly_process"],r["part_no"],r["part_name"],int(r["quantity"]),r["drawing_no"],r["material"],r["procurement_count"],r["completed_count"] or 0,r["notes"]]
            for c,v in enumerate(vals):
                item=QTableWidgetItem(str(v));
                if c in (4,7,8): item.setTextAlignment(Qt.AlignRight|Qt.AlignVCenter)
                self.table.setItem(i,c,item)
        self.table.setSortingEnabled(True); self.apply_filters()

    def import_file(self):
        path,_=QFileDialog.getOpenFileName(self,"BOMファイルを選択",str(APP_DIR),"表形式ファイル (*.csv *.xls *.xlsx)")
        if not path: return
        try: data=read_tabular_file(path,BOM_HEADER_ALIASES)
        except Exception as exc: QMessageBox.critical(self,"読込エラー",str(exc)); return
        dlg=TargetMappingDialog(data,{"production_no":"製番","assembly_process":"組立工程","part_no":"品番","part_name":"品名","quantity":"数量","drawing_no":"図番","material":"材質","notes":"備考","vendor":"発注先","due_date":"納期","issue_date":"出庫予定日","status":"ステータス"},BOM_HEADER_ALIASES,{"production_no","assembly_process","part_no"},"BOM 列対応",self)
        if not dlg.exec(): return
        count,errors=self.db.upsert_bom_rows(dlg.mapped_rows(),path)
        msg=f"{count}件を取り込みました。"+(f"\n\n未取込・警告: {len(errors)}件\n"+"\n".join(errors[:10]) if errors else "")
        QMessageBox.information(self,"取り込み結果",msg); self.refresh(); self.on_change()


class ProcurementDialog(QDialog):
    def __init__(self, db: Database, row=None, parent=None):
        super().__init__(parent); self.db=db; self.setWindowTitle("手配工程")
        form=QFormLayout(self)
        self.part=QComboBox(); self.part_rows=db.part_choices()
        for p in self.part_rows:
            self.part.addItem(f'{p["production_no"]} / {p["assembly_process"]} / {p["part_no"]} {p["part_name"]}', p["id"])
        self.order=QSpinBox(); self.order.setRange(0,9990); self.order.setSingleStep(10); self.order.setValue(10)
        self.kind=QComboBox(); self.kind.addItems(db.master_values("procurement_type") or ["材料","加工","購入品","社内加工","検査","納品"]); self.kind.setEditable(True)
        self.details=QComboBox(); self.details.addItems(db.master_values("processing_details")); self.details.setEditable(True)
        self.vendor=QComboBox(); self.vendor.addItems(db.vendor_display_values()); self.vendor.setEditable(True)
        self.due=DateField(); self.issue=DateField(); self.received=DateField()
        self.status=QComboBox(); self.status.addItems(db.master_values("status") or STATUSES); self.status.setEditable(True); self.notes=QTextEdit(); self.notes.setMaximumHeight(80)
        form.addRow("対象部品",self.part); form.addRow("工程順",self.order); form.addRow("手配区分",self.kind)
        form.addRow("加工内容",self.details); form.addRow("発注先",self.vendor); form.addRow("納期",self.due)
        form.addRow("出庫予定日",self.issue); form.addRow("実績入荷日",self.received); form.addRow("ステータス",self.status); form.addRow("備考",self.notes)
        buttons=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel); buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject); form.addRow(buttons)
        if row:
            idx=self.part.findData(row["part_id"]); self.part.setCurrentIndex(max(0,idx)); self.order.setValue(row["process_order"])
            self.kind.setCurrentText(row["procurement_type"]); self.details.setCurrentText(row["processing_details"]); self.vendor.setCurrentText(row["vendor"])
            self._set_date(self.due,row["due_date"]); self._set_date(self.issue,row["issue_date"]); self._set_date(self.received,row["actual_received_date"])
            self.status.setCurrentText(row["status"]); self.notes.setPlainText(row["notes"])
    def _set_date(self, field, value):
        if value:
            d=QDate.fromString(value,"yyyy-MM-dd")
            if d.isValid(): field.setDate(d)
    def data(self):
        return {"part_id":self.part.currentData(),"process_order":self.order.value(),"procurement_type":self.kind.currentText().strip(),"processing_details":self.details.currentText().strip(),"vendor":self.vendor.currentText().strip(),"due_date":self.due.date().toString("yyyy-MM-dd"),"issue_date":self.issue.date().toString("yyyy-MM-dd"),"actual_received_date":self.received.date().toString("yyyy-MM-dd"),"status":self.status.currentText(),"notes":self.notes.toPlainText().strip()}


class ProcurementsWidget(QWidget):
    """部品を選び、その部品の持ち回り手配工程を工程順で確認する画面。"""

    PROCESS_HEADERS = ["完了", "工程順", "区分", "加工内容", "発注先", "納期", "出庫予定", "入荷実績", "状態", "判定"]

    def __init__(self, db, on_change):
        super().__init__()
        self.db = db
        self.on_change = on_change
        self.current_part_id = None
        root = QVBoxLayout(self)

        toolbar = QHBoxLayout()
        self.prod = QComboBox(); self.prod.addItem("すべて")
        self.search = QLineEdit(); self.search.setPlaceholderText("製番・組立工程・品番・品名を検索")
        self.status_filter = QComboBox(); self.status_filter.addItem("すべて"); self.status_filter.addItems(STATUSES)
        self.problem_only = QCheckBox("要対応工程のみ")
        import_btn = QPushButton("CSV / Excel取り込み")
        bulk_btn = QPushButton("一括アーカイブ／クリア")
        apply_master = QPushButton("品番工程マスタを適用")
        add = QPushButton("手配工程を追加")
        edit = QPushButton("選択工程を編集")
        delete = QPushButton("選択工程を削除")
        refresh = QPushButton("更新")
        toolbar.addWidget(QLabel("製番")); toolbar.addWidget(self.prod)
        toolbar.addWidget(QLabel("部品検索")); toolbar.addWidget(self.search, 1)
        toolbar.addWidget(QLabel("状態")); toolbar.addWidget(self.status_filter)
        toolbar.addWidget(self.problem_only)
        toolbar.addWidget(import_btn); toolbar.addWidget(bulk_btn); toolbar.addWidget(apply_master)
        toolbar.addWidget(add); toolbar.addWidget(edit); toolbar.addWidget(delete); toolbar.addWidget(refresh)
        root.addLayout(toolbar)

        note = QLabel("左で部品を選択すると、右にその部品の材料手配・加工手配・納品工程が工程順で表示されます。")
        note.setWordWrap(True)
        root.addWidget(note)

        splitter = QSplitter(Qt.Horizontal)
        left = QWidget(); left_layout = QVBoxLayout(left); left_layout.setContentsMargins(0,0,0,0)
        left_title = QLabel("部品一覧")
        left_title.setStyleSheet("font-size:16px; font-weight:bold; padding:4px;")
        left_layout.addWidget(left_title)
        self.parts = QTableWidget(0, 6)
        self.parts.setHorizontalHeaderLabels(["製番", "組立工程", "品番", "品名", "手配個数", "進捗"])
        self.parts.setSelectionBehavior(QTableWidget.SelectRows)
        self.parts.setSelectionMode(QTableWidget.SingleSelection)
        self.parts.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.parts.horizontalHeader().setStretchLastSection(True)
        self.parts.setAlternatingRowColors(True)
        left_layout.addWidget(self.parts)

        right = QWidget(); right_layout = QVBoxLayout(right); right_layout.setContentsMargins(0,0,0,0)
        self.part_title = QLabel("部品を選択してください")
        self.part_title.setStyleSheet("font-size:18px; font-weight:bold; padding:6px;")
        self.part_meta = QLabel("")
        self.part_meta.setStyleSheet("padding: 0 6px 6px 6px;")
        self.part_progress = QProgressBar(); self.part_progress.setRange(0,100); self.part_progress.setFormat("部品進捗 %p%")
        self.part_progress.setMaximumHeight(24)
        right_layout.addWidget(self.part_title)
        right_layout.addWidget(self.part_meta)
        right_layout.addWidget(self.part_progress)
        self.processes = QTableWidget(0, len(self.PROCESS_HEADERS))
        self.processes.setHorizontalHeaderLabels(self.PROCESS_HEADERS)
        self.processes.setSelectionBehavior(QTableWidget.SelectRows)
        self.processes.setSelectionMode(QTableWidget.SingleSelection)
        self.processes.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.processes.horizontalHeader().setStretchLastSection(True)
        self.processes.setAlternatingRowColors(True)
        right_layout.addWidget(self.processes)

        self.detail = QGroupBox("選択工程の詳細")
        dg = QGridLayout(self.detail)
        self.detail_labels = {}
        for idx, key in enumerate(["工程", "区分", "加工内容", "発注先", "納期", "状態", "判定", "備考"]):
            value = QLabel("-"); value.setWordWrap(True); value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.detail_labels[key] = value
            dg.addWidget(QLabel(key), idx//4, (idx%4)*2)
            dg.addWidget(value, idx//4, (idx%4)*2+1)
        right_layout.addWidget(self.detail)

        splitter.addWidget(left); splitter.addWidget(right); splitter.setStretchFactor(0, 2); splitter.setStretchFactor(1, 5)
        root.addWidget(splitter, 1)

        self.parts.itemSelectionChanged.connect(self.select_part)
        self.processes.itemSelectionChanged.connect(self.show_process_detail)
        self.processes.doubleClicked.connect(self.edit_item)
        self.prod.currentTextChanged.connect(self.refresh)
        self.search.textChanged.connect(self.apply_part_filter)
        self.status_filter.currentTextChanged.connect(self.load_processes)
        self.problem_only.toggled.connect(self.load_processes)
        import_btn.clicked.connect(self.import_file); bulk_btn.clicked.connect(lambda: open_data_management(self, self.db, self.on_change, "procurements")); apply_master.clicked.connect(self.apply_part_master)
        add.clicked.connect(self.add_item); edit.clicked.connect(self.edit_item); delete.clicked.connect(self.delete_item); refresh.clicked.connect(self.refresh)

    @staticmethod
    def progress_for_status(status):
        return ProgressTreeWidget.STATUS_PROGRESS.get(status or "未手配", 0)

    @staticmethod
    def judgement(row):
        target = row["target_completion_date"] or row["target_issue_start_date"]
        if not row["due_date"]: return "納期未設定"
        if row["status"] != "出庫済" and row["due_date"] < date.today().isoformat(): return "納期超過"
        if target and row["due_date"] > target: return "目標超過"
        return "正常"

    def refresh(self):
        parts = self.db.list_parts()
        values = sorted({r["production_no"] for r in parts})
        current = self.prod.currentText()
        self.prod.blockSignals(True); self.prod.clear(); self.prod.addItem("すべて"); self.prod.addItems(values)
        self.prod.setCurrentText(current if current in values else "すべて"); self.prod.blockSignals(False)
        selected_prod = self.prod.currentText()
        if selected_prod != "すべて": parts = [r for r in parts if r["production_no"] == selected_prod]
        all_proc = self.db.list_procurements()
        proc_by_part = {}
        for r in all_proc: proc_by_part.setdefault(r["part_id"], []).append(r)
        previous = self.current_part_id
        self.parts.setRowCount(len(parts))
        select_row = -1
        for i, r in enumerate(parts):
            procs = proc_by_part.get(r["id"], [])
            values_p = [self.progress_for_status(x["status"]) for x in procs]
            progress = round(sum(values_p)/len(values_p)) if values_p else 0
            vals = [r["production_no"], r["assembly_process"], r["part_no"], r["part_name"], int(r["quantity"] or 0), f"{progress}%"]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(str(v))
                if c == 0: item.setData(Qt.UserRole, r["id"])
                if c in (4,5): item.setTextAlignment(Qt.AlignRight|Qt.AlignVCenter)
                self.parts.setItem(i,c,item)
            if r["id"] == previous: select_row = i
        self.apply_part_filter()
        if select_row >= 0: self.parts.selectRow(select_row)
        elif self.parts.rowCount(): self.parts.selectRow(0)
        else:
            self.current_part_id=None; self.load_processes()

    def apply_part_filter(self):
        term = normalize_header(self.search.text())
        for r in range(self.parts.rowCount()):
            text = " ".join(self.parts.item(r,c).text() for c in range(4) if self.parts.item(r,c))
            self.parts.setRowHidden(r, bool(term) and term not in normalize_header(text))

    def select_part(self):
        items = self.parts.selectedItems()
        self.current_part_id = items[0].data(Qt.UserRole) if items else None
        self.load_processes()

    def load_processes(self):
        rows = self.db.list_procurements(self.current_part_id) if self.current_part_id else []
        if rows:
            r0=rows[0]
            self.part_title.setText(f'{r0["part_no"]}  {r0["part_name"]}')
            self.part_meta.setText(f'製番 {r0["production_no"]} / 組立工程 {r0["assembly_process"]} / 手配個数 {int(r0["quantity"] or 0)}')
        elif self.current_part_id:
            part = next((r for r in self.db.list_parts() if r["id"]==self.current_part_id), None)
            if part:
                self.part_title.setText(f'{part["part_no"]}  {part["part_name"]}')
                self.part_meta.setText(f'製番 {part["production_no"]} / 組立工程 {part["assembly_process"]} / 手配個数 {int(part["quantity"] or 0)} / 手配工程なし')
        else:
            self.part_title.setText("部品を選択してください"); self.part_meta.setText("")
        status_filter=self.status_filter.currentText(); problem_only=self.problem_only.isChecked()
        display=[]
        for r in rows:
            judge=self.judgement(r)
            if status_filter!="すべて" and r["status"]!=status_filter: continue
            if problem_only and judge=="正常" and r["status"]!="未手配": continue
            display.append((r,judge))
        self.processes.setRowCount(len(display))
        progress_values=[]
        for i,(r,judge) in enumerate(display):
            progress=self.progress_for_status(r["status"]); progress_values.append(progress)
            icon="✔" if progress==100 else "▶" if r["status"] not in ("未手配","") else "□"
            vals=[icon,r["process_order"],r["procurement_type"],r["processing_details"],r["vendor"],r["due_date"] or "",r["issue_date"] or "",r["actual_received_date"] or "",r["status"],judge]
            for c,v in enumerate(vals):
                item=QTableWidgetItem(str(v))
                if c==0: item.setData(Qt.UserRole,r["id"])
                if c in (0,1): item.setTextAlignment(Qt.AlignCenter)
                if c==8: item.setForeground(QColor({"未手配":"red","発注済":"darkorange","納入済":"blue","入庫済":"purple","出庫済":"green"}.get(str(v),"black")))
                if c==9: item.setForeground(QColor("green" if judge=="正常" else "red" if judge in ("納期超過","目標超過") else "darkorange"))
                self.processes.setItem(i,c,item)
        all_values=[self.progress_for_status(r["status"]) for r in rows]
        self.part_progress.setValue(round(sum(all_values)/len(all_values)) if all_values else 0)
        if self.processes.rowCount(): self.processes.selectRow(0)
        else: self.show_process_detail()

    def selected_id(self):
        items=self.processes.selectedItems(); return items[0].data(Qt.UserRole) if items else None

    def show_process_detail(self):
        pid=self.selected_id(); row=self.db.get_procurement(pid) if pid else None
        if not row:
            for lbl in self.detail_labels.values(): lbl.setText("-")
            return
        full=next((r for r in self.db.list_procurements() if r["id"]==pid),None)
        judge=self.judgement(full) if full else "-"
        values={"工程":str(row["process_order"]),"区分":row["procurement_type"] or "-","加工内容":row["processing_details"] or "-","発注先":row["vendor"] or "-","納期":row["due_date"] or "-","状態":row["status"] or "-","判定":judge,"備考":row["notes"] or "-"}
        for k,v in values.items(): self.detail_labels[k].setText(v)

    def import_file(self):
        path,_=QFileDialog.getOpenFileName(self,"手配工程ファイルを選択",str(APP_DIR),"表形式ファイル (*.csv *.xls *.xlsx)")
        if not path: return
        try: data=read_tabular_file(path,PROCUREMENT_HEADER_ALIASES)
        except Exception as exc: QMessageBox.critical(self,"読込エラー",str(exc)); return
        labels={"production_no":"製番","assembly_process":"組立工程","part_no":"品番","process_order":"工程順","procurement_type":"手配区分","processing_details":"加工内容","vendor":"発注先","due_date":"納期","issue_date":"出庫予定日","actual_received_date":"実績入荷日","status":"ステータス","notes":"備考"}
        dlg=ImportDialog(data,labels,PROCUREMENT_HEADER_ALIASES,{"production_no","assembly_process","part_no","process_order"},"手配工程 列対応",self)
        if not dlg.exec(): return
        added,updated,errors=self.db.upsert_procurement_rows(dlg.mapped_rows(),path)
        msg=f"追加 {added}件、更新 {updated}件を取り込みました。"
        if errors: msg+=f"\n\n未取込・警告: {len(errors)}件\n"+"\n".join(errors[:12])
        QMessageBox.information(self,"取り込み結果",msg); self.on_change()

    def apply_part_master(self):
        if self.current_part_id is None:
            QMessageBox.information(self,"工程マスタ適用","左側から対象部品を選択してください。")
            return
        part=next((r for r in self.db.list_parts() if r["id"]==self.current_part_id),None)
        if not part: return
        masters=self.db.matching_route_masters(part["part_no"])
        if not masters:
            QMessageBox.information(self,"工程マスタ適用",f'品番 {part["part_no"]} の工程マスタが登録されていません。\n「マスタ」タブで登録してください。')
            return
        labels=[f'{m["route_name"]}（{len(self.db.list_route_steps(m["id"]))}工程）' for m in masters]
        selected,ok=QInputDialog.getItem(self,"工程マスタ適用",f'{part["part_no"]} に適用する工程マスタ',labels,0,False)
        if not ok:return
        master=masters[labels.index(selected)]
        answer=QMessageBox.question(self,"適用方法","既存の手配工程を削除して、マスタ工程へ置き換えますか？\n\n［はい］置換する\n［いいえ］未登録の工程順だけ追加する",QMessageBox.Yes|QMessageBox.No|QMessageBox.Cancel)
        if answer==QMessageBox.Cancel:return
        added,skipped=self.db.apply_route_master(self.current_part_id,master["id"],answer==QMessageBox.Yes)
        QMessageBox.information(self,"適用結果",f"{added}工程を追加しました。"+(f"\n{skipped}工程は同じ工程順があるためスキップしました。" if skipped else ""))
        self.on_change()

    def add_item(self):
        if not self.db.part_choices(): QMessageBox.information(self,"手配追加","先に部品表へ部品を登録してください。"); return
        dlg=ProcurementDialog(self.db,parent=self)
        if self.current_part_id is not None:
            idx=dlg.part.findData(self.current_part_id)
            if idx>=0: dlg.part.setCurrentIndex(idx)
        if dlg.exec():
            try: self.db.save_procurement(dlg.data())
            except sqlite3.IntegrityError: QMessageBox.warning(self,"登録エラー","同じ部品に同じ工程順が登録されています。"); return
            self.on_change()
            self.refresh(); self.load_processes()

    def edit_item(self,*_):
        pid=self.selected_id()
        if pid is None: return
        dlg=ProcurementDialog(self.db,self.db.get_procurement(pid),self)
        if dlg.exec():
            try: self.db.save_procurement(dlg.data(),pid)
            except sqlite3.IntegrityError: QMessageBox.warning(self,"更新エラー","同じ部品に同じ工程順が登録されています。"); return
            self.on_change()
            self.refresh(); self.load_processes()

    def delete_item(self):
        pid=self.selected_id()
        if pid is not None and QMessageBox.question(self,"削除確認","選択した手配工程を削除しますか？")==QMessageBox.Yes:
            self.db.delete_procurement(pid); self.on_change()


class ProgressTreeWidget(QWidget):
    """進捗と停止原因を視覚的に追跡するドリルダウン画面。"""

    STATUS_PROGRESS = {"未手配": 0, "発注済": 25, "納入済": 65, "入庫済": 85, "出庫済": 100}

    def __init__(self, db: Database, on_change):
        super().__init__(); self.db=db; self.on_change=on_change
        root=QVBoxLayout(self)
        bar=QHBoxLayout()
        self._drilldown_order_id = None  # 製番だけでなく組立工程まで区別する生産計画ID
        self.production_filter=QComboBox(); self.production_filter.addItem("すべて")
        self.mode_filter=QComboBox(); self.mode_filter.addItems(["すべて","要対応のみ","未完了のみ"])
        self.search=QLineEdit(); self.search.setPlaceholderText("製番・部品・加工内容・発注先を検索")
        expand_btn=QPushButton("すべて展開"); collapse_btn=QPushButton("折りたたむ"); refresh_btn=QPushButton("更新")
        bar.addWidget(QLabel("製番")); bar.addWidget(self.production_filter); bar.addWidget(QLabel("表示")); bar.addWidget(self.mode_filter)
        bar.addWidget(QLabel("検索")); bar.addWidget(self.search,1); bar.addWidget(expand_btn); bar.addWidget(collapse_btn); bar.addWidget(refresh_btn)
        root.addLayout(bar)
        scope_bar=QHBoxLayout()
        self.drilldown_label=QLabel("")
        self.drilldown_clear_btn=QPushButton("要対応の絞り込みを解除")
        self.drilldown_label.setStyleSheet("font-weight:bold; color:#b45309; padding:4px;")
        self.drilldown_label.hide(); self.drilldown_clear_btn.hide()
        scope_bar.addWidget(self.drilldown_label,1); scope_bar.addWidget(self.drilldown_clear_btn)
        root.addLayout(scope_bar)
        guide=QLabel("①生産計画 → ②部品 → ③停止工程 → ④発注先・状態 の順にクリックすると、右側に原因が表示されます。")
        guide.setWordWrap(True); guide.setStyleSheet("padding:6px; font-weight:bold;"); root.addWidget(guide)

        splitter=QSplitter(Qt.Horizontal)
        self.tree=QTreeWidget(); self.tree.setColumnCount(6)
        self.tree.setHeaderLabels(["生産計画／部品／工程", "状態", "進捗", "目標・納期", "発注先", "判定"])
        self.tree.header().setSectionResizeMode(0,QHeaderView.Stretch)
        for col in range(1,6): self.tree.header().setSectionResizeMode(col,QHeaderView.ResizeToContents)
        self.tree.setAlternatingRowColors(True); self.tree.setSelectionMode(QTreeWidget.SingleSelection)
        splitter.addWidget(self.tree)

        detail=QWidget(); dl=QVBoxLayout(detail); dl.setContentsMargins(8,0,0,0)
        self.detail_title=QLabel("項目を選択してください"); self.detail_title.setStyleSheet("font-size:20px; font-weight:bold; padding:8px;")
        self.detail_progress=QProgressBar(); self.detail_progress.setRange(0,100); self.detail_progress.setFormat("進捗 %p%")
        self.detail_progress.setMinimumHeight(28)
        self.detail_status=QLabel("-"); self.detail_status.setStyleSheet("font-size:16px; font-weight:bold; padding:8px;")
        cause_box=QGroupBox("停止原因／次に確認すること"); cg=QVBoxLayout(cause_box)
        self.detail_cause=QLabel("左側のツリーから生産計画または部品を選択してください。")
        self.detail_cause.setWordWrap(True); self.detail_cause.setTextInteractionFlags(Qt.TextSelectableByMouse); cg.addWidget(self.detail_cause)
        info_box=QGroupBox("選択項目の詳細"); ig=QGridLayout(info_box); self.detail_fields={}
        for i,key in enumerate(["種別","製番","組立工程","品番","工程順","加工内容","発注先","納期","状態","判定"]):
            val=QLabel("-"); val.setWordWrap(True); val.setTextInteractionFlags(Qt.TextSelectableByMouse); self.detail_fields[key]=val
            ig.addWidget(QLabel(key),i,0); ig.addWidget(val,i,1)
        self.edit_btn=QPushButton("この手配工程を編集"); self.edit_btn.setEnabled(False); self.edit_btn.clicked.connect(self.edit_selected)
        dl.addWidget(self.detail_title); dl.addWidget(self.detail_progress); dl.addWidget(self.detail_status); dl.addWidget(cause_box); dl.addWidget(info_box); dl.addWidget(self.edit_btn); dl.addStretch()
        splitter.addWidget(detail); splitter.setStretchFactor(0,5); splitter.setStretchFactor(1,2)
        root.addWidget(splitter,1)
        self.summary=QLabel(); self.summary.setStyleSheet("font-weight:bold; padding:6px;"); root.addWidget(self.summary)

        self.production_filter.currentTextChanged.connect(self.on_manual_filter_change); self.mode_filter.currentTextChanged.connect(self.on_manual_filter_change); self.search.textChanged.connect(self.apply_search)
        self.drilldown_clear_btn.clicked.connect(self.clear_drilldown)
        expand_btn.clicked.connect(self.tree.expandAll); collapse_btn.clicked.connect(self.tree.collapseAll); refresh_btn.clicked.connect(self.refresh)
        self.tree.itemDoubleClicked.connect(self.edit_procurement); self.tree.itemSelectionChanged.connect(self.update_detail)

    @classmethod
    def procurement_progress(cls,status): return cls.STATUS_PROGRESS.get(status or "未手配",0)
    @staticmethod
    def target_date(row): return row["parts_completion_date"] or row["parts_issue_start_date"]
    @classmethod
    def judgement(cls,row):
        status=row["procurement_status"]
        if not row["procurement_id"]: return "手配なし","darkorange"
        if status=="出庫済": return "完了","green"
        due=row["procurement_due_date"]; target=cls.target_date(row); today=date.today().isoformat()
        if due and due<today: return "納期超過","red"
        if due and target and due>target: return "目標超過","red"
        if status=="未手配": return "未手配","darkorange"
        if not due: return "納期未設定","darkorange"
        return "進行中","blue"
    @staticmethod
    def make_progress_bar(value):
        bar=QProgressBar(); bar.setRange(0,100); bar.setValue(value); bar.setFormat(f"{value}%"); bar.setMaximumHeight(22)
        return bar

    def on_manual_filter_change(self, _value=None):
        """手動フィルターを優先し、ダッシュボードからの一時絞り込みを解除する。"""
        self._drilldown_order_id = None
        self.drilldown_label.hide(); self.drilldown_clear_btn.hide()
        self.refresh()

    def clear_drilldown(self):
        """対象工程を絞った一時表示から、元の全体一覧に戻る。"""
        self._drilldown_order_id = None
        self.production_filter.blockSignals(True)
        self.mode_filter.blockSignals(True)
        try:
            self.production_filter.setCurrentText("すべて")
            self.mode_filter.setCurrentText("すべて")
        finally:
            self.production_filter.blockSignals(False)
            self.mode_filter.blockSignals(False)
        self.search.clear()
        self.drilldown_label.hide(); self.drilldown_clear_btn.hide()
        self.refresh()

    def refresh(self):
        rows=self.db.progress_tree_rows(); values=sorted({r["production_no"] for r in rows}); current=self.production_filter.currentText()
        self.production_filter.blockSignals(True); self.production_filter.clear(); self.production_filter.addItem("すべて"); self.production_filter.addItems(values); self.production_filter.setCurrentText(current if current in values else "すべて"); self.production_filter.blockSignals(False)
        selected=self.production_filter.currentText()
        if selected!="すべて": rows=[r for r in rows if r["production_no"]==selected]
        if self._drilldown_order_id is not None:
            rows=[r for r in rows if r["order_id"] == self._drilldown_order_id]
        mode=self.mode_filter.currentText(); self.tree.clear(); orders={}; parts={}; total_proc=completed_proc=problem_proc=0
        for row in rows:
            # 一時絞り込みでは問題のある工程だけ残し、正常部品・正常工程を出さない。
            if self._drilldown_order_id is not None:
                if row["procurement_id"] is None or self.judgement(row)[0] not in (
                    "納期超過", "目標超過", "未手配", "手配なし", "納期未設定"
                ):
                    continue
            oid=row["order_id"]
            if oid not in orders:
                item=QTreeWidgetItem([f'▾ {row["production_no"]}  {row["assembly_process"]}',f'{row["series"]} / {row["order_state"]}',"",row["parts_completion_date"] or row["order_due_date"] or "","",""])
                item.setData(0,Qt.UserRole,"order"); item.setData(0,Qt.UserRole+1,oid); self.tree.addTopLevelItem(item); orders[oid]={"item":item,"progress":[],"problems":[]}
            pid=row["part_id"]
            if pid is None: continue
            if pid not in parts:
                item=QTreeWidgetItem([f'   ▾ {row["part_no"]}  {row["part_name"]}  ×{int(row["quantity"] or 0)}',"部品","",self.target_date(row) or "","",""])
                item.setData(0,Qt.UserRole,"part"); item.setData(0,Qt.UserRole+1,pid); item.setData(0,Qt.UserRole+2,oid); orders[oid]["item"].addChild(item); parts[pid]={"item":item,"progress":[],"problems":[],"order_id":oid}
            if row["procurement_id"] is None: continue
            progress=self.procurement_progress(row["procurement_status"]); judge,color=self.judgement(row); problem=judge in ("納期超過","目標超過","未手配","手配なし","納期未設定")
            total_proc+=1; completed_proc+=progress==100; problem_proc+=problem
            if mode=="要対応のみ" and not problem: continue
            if mode=="未完了のみ" and progress==100: continue
            mark="✔" if progress==100 else "⚠" if problem else "▶"
            item=QTreeWidgetItem([f'      {mark} {row["process_order"]}: {row["processing_details"] or "（内容未設定）"}',f'{row["procurement_type"]} / {row["procurement_status"]}',"",row["procurement_due_date"] or "",row["vendor"] or "",judge])
            item.setData(0,Qt.UserRole,"procurement"); item.setData(0,Qt.UserRole+1,row["procurement_id"]); item.setData(0,Qt.UserRole+2,pid)
            item.setData(0,Qt.UserRole+3,{"production_no":row["production_no"],"assembly_process":row["assembly_process"],"part_no":row["part_no"],"part_name":row["part_name"],"process_order":row["process_order"],"processing_details":row["processing_details"],"vendor":row["vendor"],"due":row["procurement_due_date"],"status":row["procurement_status"],"judge":judge,"progress":progress})
            item.setForeground(5,QColor(color)); parts[pid]["item"].addChild(item); self.tree.setItemWidget(item,2,self.make_progress_bar(progress)); parts[pid]["progress"].append(progress)
            if problem: parts[pid]["problems"].append(f'{row["processing_details"] or "工程未設定"}：{judge}（{row["vendor"] or "発注先未設定"}）')
        for pid,info in parts.items():
            item=info["item"]
            if mode!="すべて" and item.childCount()==0: item.parent().removeChild(item); continue
            progress=round(sum(info["progress"])/len(info["progress"])) if info["progress"] else 0; item.setData(0,Qt.UserRole+3,{"progress":progress,"problems":info["problems"]})
            self.tree.setItemWidget(item,2,self.make_progress_bar(progress)); item.setText(5,f'要対応 {len(info["problems"])}件' if info["problems"] else ('完了' if progress==100 else '進行中')); item.setForeground(5,QColor('red' if info["problems"] else 'green' if progress==100 else 'blue'))
            orders[info["order_id"]]["progress"].append(progress); orders[info["order_id"]]["problems"].extend(info["problems"])
        for oid,info in orders.items():
            item=info["item"]
            if mode!="すべて" and item.childCount()==0:
                idx=self.tree.indexOfTopLevelItem(item)
                if idx>=0:self.tree.takeTopLevelItem(idx)
                continue
            progress=round(sum(info["progress"])/len(info["progress"])) if info["progress"] else 0; item.setData(0,Qt.UserRole+3,{"progress":progress,"problems":info["problems"]})
            self.tree.setItemWidget(item,2,self.make_progress_bar(progress)); item.setText(5,f'要対応 {len(info["problems"])}件' if info["problems"] else ('完了' if progress==100 else '進行中')); item.setForeground(5,QColor('red' if info["problems"] else 'green' if progress==100 else 'blue'))
        self.summary.setText(f'表示: 生産計画 {self.tree.topLevelItemCount()}件 / 手配工程 {total_proc}件 / 完了 {completed_proc}件 / 要対応 {problem_proc}件')
        self.tree.expandToDepth(1); self.apply_search(); self.update_detail()

    def focus_procurement(self, procurement_id: int) -> bool:
        """対象の製番×組立工程の要対応工程だけ表示し、元の手配IDを選択する。"""
        rows = self.db.progress_tree_rows()
        target = next((r for r in rows if r["procurement_id"] == procurement_id), None)
        if target is None:
            return False
        # order_idを使うことで同じ製番の別組立工程が混ざらない。
        self._drilldown_order_id = target["order_id"]
        self.production_filter.blockSignals(True)
        self.mode_filter.blockSignals(True)
        self.search.blockSignals(True)
        try:
            self.production_filter.setCurrentText(target["production_no"])
            self.mode_filter.setCurrentText("要対応のみ")
            self.search.clear()
        finally:
            self.production_filter.blockSignals(False)
            self.mode_filter.blockSignals(False)
            self.search.blockSignals(False)
        self.drilldown_label.setText(
            f'ダッシュボードからの絞り込み：製番 {target["production_no"]} / '
            f'組立工程 {target["assembly_process"]} / 要対応工程のみ'
        )
        self.drilldown_label.show(); self.drilldown_clear_btn.show()
        self.refresh()
        for i in range(self.tree.topLevelItemCount()):
            order_item = self.tree.topLevelItem(i)
            for j in range(order_item.childCount()):
                part_item = order_item.child(j)
                for k in range(part_item.childCount()):
                    item = part_item.child(k)
                    if (item.data(0, Qt.UserRole) == "procurement"
                            and item.data(0, Qt.UserRole + 1) == procurement_id):
                        self.tree.expandItem(order_item)
                        self.tree.expandItem(part_item)
                        self.tree.setCurrentItem(item)
                        self.tree.scrollToItem(item)
                        self.update_detail()
                        return True
        # 表示中に判定が変わった場合も対象工程への移動は失敗扱いにせず、
        # 同じ製番・組立工程内で通常表示して対象を探す。
        self.mode_filter.blockSignals(True)
        try:
            self.mode_filter.setCurrentText("すべて")
        finally:
            self.mode_filter.blockSignals(False)
        self._drilldown_order_id = None
        self.drilldown_label.hide(); self.drilldown_clear_btn.hide()
        self.refresh()
        for i in range(self.tree.topLevelItemCount()):
            order_item = self.tree.topLevelItem(i)
            for j in range(order_item.childCount()):
                part_item = order_item.child(j)
                for k in range(part_item.childCount()):
                    item = part_item.child(k)
                    if item.data(0, Qt.UserRole + 1) == procurement_id and item.data(0, Qt.UserRole) == "procurement":
                        self.tree.expandItem(order_item); self.tree.expandItem(part_item)
                        self.tree.setCurrentItem(item); self.tree.scrollToItem(item)
                        self.update_detail()
                        return True
        return False

    def apply_search(self):
        term=normalize_header(self.search.text())
        for i in range(self.tree.topLevelItemCount()):
            order=self.tree.topLevelItem(i); order_visible=not term
            for j in range(order.childCount()):
                part=order.child(j); part_visible=not term
                for k in range(part.childCount()):
                    proc=part.child(k); visible=not term or term in normalize_header(' '.join(proc.text(c) for c in range(6))); proc.setHidden(not visible); part_visible=part_visible or visible
                if term and term in normalize_header(' '.join(part.text(c) for c in range(6))): part_visible=True; [part.child(k).setHidden(False) for k in range(part.childCount())]
                part.setHidden(not part_visible); order_visible=order_visible or part_visible
            if term and term in normalize_header(' '.join(order.text(c) for c in range(6))): order_visible=True; [order.child(j).setHidden(False) for j in range(order.childCount())]
            order.setHidden(not order_visible)
        if term:self.tree.expandAll()

    def update_detail(self):
        item=self.tree.currentItem()
        for v in self.detail_fields.values(): v.setText('-')
        self.edit_btn.setEnabled(False)
        if not item:
            self.detail_title.setText('項目を選択してください'); self.detail_progress.setValue(0); self.detail_status.setText('-'); self.detail_cause.setText('左側のツリーから生産計画または部品を選択してください。'); return
        kind=item.data(0,Qt.UserRole); data=item.data(0,Qt.UserRole+3) or {}; progress=int(data.get('progress',0))
        self.detail_progress.setValue(progress); self.detail_fields['種別'].setText({'order':'生産計画','part':'部品','procurement':'手配工程'}.get(kind,'-'))
        if kind=='procurement':
            self.detail_title.setText(f'{data.get("part_name","")} / {data.get("processing_details","")}'); self.detail_status.setText(f'{data.get("judge","-")} － {data.get("status","-")}')
            self.detail_cause.setText('この工程が現在の停止点です。発注先、納期、ステータスを確認し、必要なら下のボタンから更新してください。' if data.get('judge') not in ('完了','進行中') else '工程は進行中です。次工程へ渡せる状態か確認してください。')
            mapping={'製番':'production_no','組立工程':'assembly_process','品番':'part_no','工程順':'process_order','加工内容':'processing_details','発注先':'vendor','納期':'due','状態':'status','判定':'judge'}
            for label,key in mapping.items(): self.detail_fields[label].setText(str(data.get(key) or '-'))
            self.edit_btn.setEnabled(True)
        else:
            self.detail_title.setText(item.text(0).replace('▾','').strip()); problems=data.get('problems',[]); self.detail_status.setText(item.text(5) or '進行中')
            self.detail_cause.setText('\n'.join(f'• {p}' for p in problems[:8]) if problems else ('すべての工程が完了しています。' if progress==100 else '重大な遅延判定はありません。未完了工程を順に確認してください。'))
            if kind=='order':
                parts=item.text(0).replace('▾','').strip().split(); self.detail_fields['製番'].setText(parts[0] if parts else '-'); self.detail_fields['組立工程'].setText(parts[1] if len(parts)>1 else '-')
            elif kind=='part':
                text=item.text(0).replace('▾','').strip(); self.detail_fields['品番'].setText(text.split()[0] if text else '-')

    def edit_selected(self):
        item=self.tree.currentItem()
        if item: self.edit_procurement(item,0)

    def edit_procurement(self,item,_column):
        if item.data(0,Qt.UserRole)!="procurement": return
        pid=item.data(0,Qt.UserRole+1); row=self.db.get_procurement(pid)
        if not row:return
        dlg=ProcurementDialog(self.db,row,self)
        if dlg.exec():
            try:self.db.save_procurement(dlg.data(),pid)
            except sqlite3.IntegrityError: QMessageBox.warning(self,"更新エラー","同じ部品に同じ工程順が登録されています。"); return
            self.on_change()


class RouteMasterStepDialog(QDialog):
    def __init__(self, db: Database, master_id: int, row=None, parent=None):
        super().__init__(parent); self.master_id=master_id; self.setWindowTitle("品番別 工程マスタ")
        form=QFormLayout(self)
        self.order=QSpinBox(); self.order.setRange(0,9990); self.order.setSingleStep(10); self.order.setValue(10)
        self.kind=QComboBox(); self.kind.addItems(db.master_values("procurement_type") or ["材料","加工","購入品","社内加工","検査","納品"]); self.kind.setEditable(True)
        self.details=QComboBox(); self.details.addItems(db.master_values("processing_details")); self.details.setEditable(True)
        self.vendor=QComboBox(); self.vendor.addItems(db.vendor_display_values()); self.vendor.setEditable(True); self.lead=QSpinBox(); self.lead.setRange(0,3650); self.notes=QTextEdit(); self.notes.setMaximumHeight(70)
        form.addRow("工程順",self.order); form.addRow("手配区分",self.kind); form.addRow("加工内容",self.details); form.addRow("標準発注先",self.vendor); form.addRow("標準LT（日）",self.lead); form.addRow("備考",self.notes)
        buttons=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel); buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject); form.addRow(buttons)
        if row:
            self.order.setValue(row["process_order"]); self.kind.setCurrentText(row["procurement_type"]); self.details.setCurrentText(row["processing_details"]); self.vendor.setCurrentText(row["vendor"]); self.lead.setValue(row["lead_time_days"]); self.notes.setPlainText(row["notes"])
    def data(self):
        return {"route_master_id":self.master_id,"process_order":self.order.value(),"procurement_type":self.kind.currentText().strip(),"processing_details":self.details.currentText().strip(),"vendor":self.vendor.currentText().strip(),"lead_time_days":self.lead.value(),"notes":self.notes.toPlainText().strip()}


class VendorDialog(QDialog):
    def __init__(self, row=None, parent=None):
        super().__init__(parent); self.setWindowTitle("発注先マスタ")
        form=QFormLayout(self); self.code=QLineEdit(); self.name=QLineEdit(); self.notes=QTextEdit(); self.notes.setMaximumHeight(80)
        form.addRow("発注先コード",self.code); form.addRow("発注先名称",self.name); form.addRow("備考",self.notes)
        buttons=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel); buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject); form.addRow(buttons)
        if row: self.code.setText(row["vendor_code"]); self.name.setText(row["vendor_name"]); self.notes.setPlainText(row["notes"])
    def data(self): return self.code.text().strip(),self.name.text().strip(),self.notes.toPlainText().strip()


class ChoiceMasterDialog(QDialog):
    def __init__(self, title, is_status=False, row=None, parent=None):
        super().__init__(parent); self.setWindowTitle(title); self.is_status=is_status
        form=QFormLayout(self); self.value=QLineEdit(); self.order=QSpinBox(); self.order.setRange(0,9999); self.order.setValue(100); self.progress=QSpinBox(); self.progress.setRange(0,100); self.progress.setSuffix(" %")
        form.addRow("名称",self.value); form.addRow("表示順",self.order)
        if is_status: form.addRow("進捗率",self.progress)
        buttons=QDialogButtonBox(QDialogButtonBox.Ok|QDialogButtonBox.Cancel); buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject); form.addRow(buttons)
        if row:
            self.value.setText(row["value"]); self.order.setValue(row["sort_order"])
            if is_status and row["progress_percent"] is not None:self.progress.setValue(row["progress_percent"])
    def data(self): return self.value.text().strip(),self.order.value(),self.progress.value() if self.is_status else None


class CommonMastersWidget(QWidget):
    INFO={"processing_details":("加工内容",False),"procurement_type":("手配区分",False),"status":("ステータス",True)}
    def __init__(self,db,on_change):
        super().__init__(); self.db=db; self.on_change=on_change; self.tables={}; root=QVBoxLayout(self); self.tabs=QTabWidget(); root.addWidget(self.tabs); self._vendor_tab()
        for category,(label,is_status) in self.INFO.items():self._choice_tab(category,label,is_status)
    def _vendor_tab(self):
        page=QWidget(); lay=QVBoxLayout(page); bar=QHBoxLayout()
        for label,slot in (("追加",self.add_vendor),("編集",self.edit_vendor),("削除",self.delete_vendor),("CSV / Excel取り込み",self.import_vendors),("Excel貼り付け",self.paste_vendors)):
            b=QPushButton(label); b.clicked.connect(slot); bar.addWidget(b)
        bar.addStretch(); lay.addLayout(bar); hint=QLabel("コードと名称は重複不可です。Excelの複数行を『コード／名称／備考』の順でコピーし、そのまま貼り付けできます。"); hint.setWordWrap(True); lay.addWidget(hint)
        self.vendor_table=QTableWidget(0,3); self.vendor_table.setHorizontalHeaderLabels(["発注先コード","発注先名称","備考"]); self.vendor_table.setSelectionBehavior(QTableWidget.SelectRows); self.vendor_table.setSelectionMode(QTableWidget.SingleSelection); self.vendor_table.horizontalHeader().setStretchLastSection(True); self.vendor_table.doubleClicked.connect(self.edit_vendor); lay.addWidget(self.vendor_table); self.tabs.addTab(page,"発注先")
    def _choice_tab(self,category,label,is_status):
        page=QWidget(); lay=QVBoxLayout(page); bar=QHBoxLayout()
        for text,slot in (("追加",lambda _=False,c=category:self.add_choice(c)),("編集",lambda _=False,c=category:self.edit_choice(c)),("削除",lambda _=False,c=category:self.delete_choice(c)),("Excel貼り付け",lambda _=False,c=category:self.paste_choices(c))):
            b=QPushButton(text); b.clicked.connect(slot); bar.addWidget(b)
        bar.addStretch(); lay.addLayout(bar); cols=["名称","表示順"]+(["進捗率"] if is_status else []); table=QTableWidget(0,len(cols)); table.setHorizontalHeaderLabels(cols); table.setSelectionBehavior(QTableWidget.SelectRows); table.setSelectionMode(QTableWidget.SingleSelection); table.horizontalHeader().setStretchLastSection(True); table.doubleClicked.connect(lambda _idx,c=category:self.edit_choice(c)); lay.addWidget(table); self.tables[category]=table; self.tabs.addTab(page,label)
    def selected_id(self,table):
        items=table.selectedItems(); return items[0].data(Qt.UserRole) if items else None
    def refresh(self):
        rows=self.db.list_vendors(); self.vendor_table.setRowCount(len(rows))
        for i,r in enumerate(rows):
            for c,v in enumerate([r["vendor_code"],r["vendor_name"],r["notes"]]):
                item=QTableWidgetItem(str(v));
                if c==0:item.setData(Qt.UserRole,r["id"])
                self.vendor_table.setItem(i,c,item)
        for category,(label,is_status) in self.INFO.items():
            rows=self.db.list_choice_master(category); table=self.tables[category]; table.setRowCount(len(rows))
            for i,r in enumerate(rows):
                vals=[r["value"],r["sort_order"]]+([r["progress_percent"] if r["progress_percent"] is not None else 0] if is_status else [])
                for c,v in enumerate(vals):
                    item=QTableWidgetItem(str(v));
                    if c==0:item.setData(Qt.UserRole,r["id"])
                    table.setItem(i,c,item)
    def add_vendor(self):
        dlg=VendorDialog(parent=self)
        if dlg.exec():
            try:self.db.save_vendor(*dlg.data())
            except (sqlite3.IntegrityError,ValueError) as exc:QMessageBox.warning(self,"登録エラー",f"コードまたは名称が重複しています。\n{exc}");return
            self.on_change()
    def edit_vendor(self,*_):
        vid=self.selected_id(self.vendor_table); row=self.db.conn.execute("SELECT * FROM vendor_masters WHERE id=?",(vid,)).fetchone() if vid else None
        if not row:return
        dlg=VendorDialog(row,self)
        if dlg.exec():
            try:self.db.save_vendor(*dlg.data(),vendor_id=vid)
            except (sqlite3.IntegrityError,ValueError) as exc:QMessageBox.warning(self,"更新エラー",f"コードまたは名称が重複しています。\n{exc}");return
            self.on_change()
    def delete_vendor(self):
        vid=self.selected_id(self.vendor_table)
        if vid and QMessageBox.question(self,"削除確認","発注先マスタを削除しますか？既存の手配データは変更されません。")==QMessageBox.Yes:self.db.delete_vendor(vid);self.on_change()
    def import_vendors(self):
        path,_=QFileDialog.getOpenFileName(self,"発注先マスタファイルを選択",str(APP_DIR),"表形式ファイル (*.csv *.xls *.xlsx)")
        if not path:return
        try:data=read_tabular_file(path,VENDOR_MASTER_HEADER_ALIASES)
        except Exception as exc:QMessageBox.critical(self,"読込エラー",str(exc));return
        dlg=ImportDialog(data,{"vendor_code":"発注先コード","vendor_name":"発注先名称","notes":"備考"},VENDOR_MASTER_HEADER_ALIASES,{"vendor_code","vendor_name"},"発注先マスタ 列対応",self)
        if not dlg.exec():return
        added,updated,errors=self.db.upsert_vendor_rows(dlg.mapped_rows(),path); msg=f"追加 {added}件、更新 {updated}件。"+(f"\n\n警告 {len(errors)}件\n"+"\n".join(errors[:12]) if errors else ""); QMessageBox.information(self,"取り込み結果",msg); self.on_change()
    def paste_vendors(self):
        text=QApplication.clipboard().text().strip()
        if not text:QMessageBox.information(self,"貼り付け","クリップボードにデータがありません。");return
        rows=[]
        for line in text.splitlines():
            cols=line.split("\t")
            if len(cols)<2:cols=[x.strip() for x in line.split(",")]
            if len(cols)>=2:rows.append({"vendor_code":cols[0].strip(),"vendor_name":cols[1].strip(),"notes":cols[2].strip() if len(cols)>2 else ""})
        if rows and normalize_header(rows[0]["vendor_code"]) in {normalize_header(x) for x in VENDOR_MASTER_HEADER_ALIASES["vendor_code"]}:rows=rows[1:]
        added,updated,errors=self.db.upsert_vendor_rows(rows,"clipboard"); QMessageBox.information(self,"貼り付け結果",f"追加 {added}件、更新 {updated}件、警告 {len(errors)}件。"+("\n"+"\n".join(errors[:10]) if errors else "")); self.on_change()
    def add_choice(self,category):
        label,is_status=self.INFO[category]; dlg=ChoiceMasterDialog(label+"マスタ",is_status,parent=self)
        if dlg.exec():
            try:self.db.save_choice_master(category,*dlg.data())
            except (sqlite3.IntegrityError,ValueError) as exc:QMessageBox.warning(self,"登録エラー",f"同じ名称が既にあります。\n{exc}");return
            self.on_change()
    def edit_choice(self,category):
        table=self.tables[category]; iid=self.selected_id(table); row=self.db.conn.execute("SELECT * FROM choice_masters WHERE id=?",(iid,)).fetchone() if iid else None
        if not row:return
        label,is_status=self.INFO[category]; dlg=ChoiceMasterDialog(label+"マスタ",is_status,row,self)
        if dlg.exec():
            try:self.db.save_choice_master(category,*dlg.data(),item_id=iid)
            except (sqlite3.IntegrityError,ValueError) as exc:QMessageBox.warning(self,"更新エラー",f"同じ名称が既にあります。\n{exc}");return
            self.on_change()
    def delete_choice(self,category):
        iid=self.selected_id(self.tables[category])
        if iid and QMessageBox.question(self,"削除確認","選択したマスタ値を削除しますか？既存データは変更されません。")==QMessageBox.Yes:self.db.delete_choice_master(iid);self.on_change()
    def paste_choices(self,category):
        text=QApplication.clipboard().text().strip()
        if not text:QMessageBox.information(self,"貼り付け","クリップボードにデータがありません。");return
        added=skipped=0
        for index,line in enumerate(text.splitlines(),start=1):
            cols=line.split("\t"); value=cols[0].strip()
            if not value or normalize_header(value) in {"名称","加工内容","手配区分","ステータス"}:continue
            try:
                order=int(cols[1]) if len(cols)>1 and cols[1].strip() else index*10; progress=int(cols[2]) if category=="status" and len(cols)>2 and cols[2].strip() else (0 if category=="status" else None); self.db.save_choice_master(category,value,order,progress); added+=1
            except (sqlite3.IntegrityError,ValueError):skipped+=1
        QMessageBox.information(self,"貼り付け結果",f"追加 {added}件、重複・不正 {skipped}件。"); self.on_change()


class MastersWidget(QWidget):
    def __init__(self,db,on_change):
        super().__init__(); self.db=db; self.on_change=on_change; self.current_master_id=None
        root=QVBoxLayout(self)
        title_bar=QHBoxLayout(); title=QLabel("マスタ"); title.setStyleSheet("font-size: 18px; font-weight: bold;"); title_bar.addWidget(title); title_bar.addStretch(); bulk_btn=QPushButton("一括アーカイブ／クリア／復帰"); title_bar.addWidget(bulk_btn); root.addLayout(title_bar)
        bulk_btn.clicked.connect(lambda: open_data_management(self, self.db, self.on_change, "route_masters"))
        self.tabs=QTabWidget(); root.addWidget(self.tabs)
        route=QWidget(); layout=QVBoxLayout(route)
        bar=QHBoxLayout(); addm=QPushButton("品番工程マスタ追加"); editm=QPushButton("名称編集"); delm=QPushButton("削除"); imp=QPushButton("CSV / Excel取り込み"); adds=QPushButton("工程追加"); edits=QPushButton("工程編集"); dels=QPushButton("工程削除")
        for w in (addm,editm,delm,imp): bar.addWidget(w)
        bar.addStretch()
        for w in (adds,edits,dels): bar.addWidget(w)
        layout.addLayout(bar)
        split=QSplitter(Qt.Horizontal)
        self.master_table=QTableWidget(0,4); self.master_table.setHorizontalHeaderLabels(["品番","工程マスタ名","工程数","備考"]); self.master_table.setSelectionBehavior(QTableWidget.SelectRows); self.master_table.setSelectionMode(QTableWidget.SingleSelection); self.master_table.horizontalHeader().setStretchLastSection(True)
        self.step_table=QTableWidget(0,7); self.step_table.setHorizontalHeaderLabels(["工程順","区分","加工内容","標準発注先","標準LT（日）","備考",""]); self.step_table.setSelectionBehavior(QTableWidget.SelectRows); self.step_table.setSelectionMode(QTableWidget.SingleSelection); self.step_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents); self.step_table.horizontalHeader().setStretchLastSection(True)
        split.addWidget(self.master_table); split.addWidget(self.step_table); split.setStretchFactor(0,2); split.setStretchFactor(1,4); layout.addWidget(split)
        self.tabs.addTab(route,"品番別 持ち回り工程")
        self.common_masters=CommonMastersWidget(self.db,self.on_change); self.tabs.addTab(self.common_masters,"共通マスタ編集")
        self.master_table.itemSelectionChanged.connect(self.select_master); addm.clicked.connect(self.add_master); editm.clicked.connect(self.edit_master); delm.clicked.connect(self.delete_master); imp.clicked.connect(self.import_master); adds.clicked.connect(self.add_step); edits.clicked.connect(self.edit_step); dels.clicked.connect(self.delete_step); self.step_table.doubleClicked.connect(self.edit_step)

    def selected_master_id(self):
        items=self.master_table.selectedItems(); return items[0].data(Qt.UserRole) if items else None
    def selected_step_id(self):
        items=self.step_table.selectedItems(); return items[0].data(Qt.UserRole) if items else None
    def refresh(self):
        previous=self.current_master_id; rows=self.db.list_route_masters(); self.master_table.setRowCount(len(rows)); selected=-1
        for i,r in enumerate(rows):
            vals=[r["part_no"],r["route_name"],r["step_count"],r["notes"]]
            for c,v in enumerate(vals):
                it=QTableWidgetItem(str(v));
                if c==0: it.setData(Qt.UserRole,r["id"])
                self.master_table.setItem(i,c,it)
            if r["id"]==previous:selected=i
        if selected>=0:self.master_table.selectRow(selected)
        elif rows:self.master_table.selectRow(0)
        else:self.current_master_id=None; self.load_steps()
        self.common_masters.refresh()
    def select_master(self): self.current_master_id=self.selected_master_id(); self.load_steps()
    def load_steps(self):
        rows=self.db.list_route_steps(self.current_master_id) if self.current_master_id else []; self.step_table.setRowCount(len(rows))
        for i,r in enumerate(rows):
            vals=[r["process_order"],r["procurement_type"],r["processing_details"],r["vendor"],r["lead_time_days"],r["notes"],""]
            for c,v in enumerate(vals):
                it=QTableWidgetItem(str(v));
                if c==0:it.setData(Qt.UserRole,r["id"])
                self.step_table.setItem(i,c,it)
    def master_input(self,row=None):
        part,ok=QInputDialog.getText(self,"品番工程マスタ","品番",text=row["part_no"] if row else "")
        if not ok or not part.strip(): return None
        name,ok=QInputDialog.getText(self,"品番工程マスタ","工程マスタ名",text=row["route_name"] if row else "標準工程")
        if not ok:return None
        return part.strip(),name.strip() or "標準工程",row["notes"] if row else ""
    def add_master(self):
        data=self.master_input()
        if not data:return
        try:self.current_master_id=self.db.save_route_master(*data)
        except sqlite3.IntegrityError:QMessageBox.warning(self,"登録エラー","同じ品番・工程マスタ名が既にあります。");return
        self.on_change()
    def edit_master(self):
        mid=self.selected_master_id(); row=self.db.get_route_master(mid) if mid else None
        if not row:return
        data=self.master_input(row)
        if not data:return
        try:self.db.save_route_master(*data,master_id=mid)
        except sqlite3.IntegrityError:QMessageBox.warning(self,"更新エラー","同じ品番・工程マスタ名が既にあります。");return
        self.on_change()
    def delete_master(self):
        mid=self.selected_master_id()
        if mid and QMessageBox.question(self,"削除確認","選択した工程マスタと全工程を削除しますか？")==QMessageBox.Yes:self.db.delete_route_master(mid);self.current_master_id=None;self.on_change()
    def add_step(self):
        if not self.current_master_id:QMessageBox.information(self,"工程追加","先に品番工程マスタを選択してください。");return
        dlg=RouteMasterStepDialog(self.db,self.current_master_id,parent=self)
        if dlg.exec():
            try:self.db.save_route_step(dlg.data())
            except sqlite3.IntegrityError:QMessageBox.warning(self,"登録エラー","同じ工程順が登録されています。");return
            self.on_change()
    def edit_step(self,*_):
        sid=self.selected_step_id()
        if not sid:return
        row=self.db.conn.execute("SELECT * FROM route_master_steps WHERE id=?",(sid,)).fetchone(); dlg=RouteMasterStepDialog(self.db,self.current_master_id,row,self)
        if dlg.exec():
            try:self.db.save_route_step(dlg.data(),sid)
            except sqlite3.IntegrityError:QMessageBox.warning(self,"更新エラー","同じ工程順が登録されています。");return
            self.on_change()
    def delete_step(self):
        sid=self.selected_step_id()
        if sid and QMessageBox.question(self,"削除確認","選択したマスタ工程を削除しますか？")==QMessageBox.Yes:self.db.delete_route_step(sid);self.on_change()
    def import_master(self):
        path,_=QFileDialog.getOpenFileName(self,"工程マスタファイルを選択",str(APP_DIR),"表形式ファイル (*.csv *.xls *.xlsx)")
        if not path:return
        try:data=read_tabular_file(path,ROUTE_MASTER_HEADER_ALIASES)
        except Exception as exc:QMessageBox.critical(self,"読込エラー",str(exc));return
        labels={"part_no":"品番","route_name":"工程マスタ名","process_order":"工程順","procurement_type":"手配区分","processing_details":"加工内容","vendor":"標準発注先","lead_time_days":"標準LT（日）","notes":"備考"}
        dlg=ImportDialog(data,labels,ROUTE_MASTER_HEADER_ALIASES,{"part_no","process_order"},"品番別工程マスタ 列対応",self)
        if not dlg.exec():return
        count,errors=self.db.upsert_route_master_rows(dlg.mapped_rows(),path);msg=f"{count}工程を取り込みました。"+(f"\n\n警告: {len(errors)}件\n"+"\n".join(errors[:12]) if errors else "");QMessageBox.information(self,"取り込み結果",msg);self.on_change()



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.db=Database(DB_PATH)
        self.setWindowTitle("生産管理アプリ Ver.1.0α Phase 11"); self.resize(1480,900)
        self.tabs=QTabWidget(); self.dashboard=DashboardWidget(self.db); self.progress_tree=ProgressTreeWidget(self.db,self.refresh_all); self.orders=OrdersWidget(self.db,self.refresh_all); self.bom=BomWidget(self.db,self.refresh_all); self.procurements=ProcurementsWidget(self.db,self.refresh_all); self.masters=MastersWidget(self.db,self.refresh_all)
        self.tabs.addTab(self.dashboard,"ダッシュボード"); self.tabs.addTab(self.progress_tree,"進捗ツリー"); self.tabs.addTab(self.orders,"生産計画"); self.tabs.addTab(self.bom,"部品表"); self.tabs.addTab(self.procurements,"手配"); self.tabs.addTab(self.masters,"マスタ")
        self.setCentralWidget(self.tabs)
        self.dashboard.procurement_requested.connect(self.open_procurement_in_tree)
        self.dashboard.bottleneck_requested.connect(self.open_bottleneck_details)
        refresh_action=QAction("すべて更新",self); refresh_action.triggered.connect(self.refresh_all); self.menuBar().addAction(refresh_action)
        data_action=QAction("一括データ管理",self); data_action.triggered.connect(lambda: open_data_management(self, self.db, self.refresh_all, "orders")); self.menuBar().addAction(data_action)
        excel_action=QAction("Excel共有出力",self); excel_action.setShortcut("Ctrl+Shift+E"); excel_action.triggered.connect(lambda: export_excel_snapshot(self, self.db)); self.menuBar().addAction(excel_action)
        self.setStatusBar(QStatusBar()); self.refresh_all()
    def open_procurement_in_tree(self, procurement_id: int) -> None:
        """ダッシュボードの要対応行から該当手配工程を開く。"""
        if self.progress_tree.focus_procurement(procurement_id):
            self.tabs.setCurrentWidget(self.progress_tree)
            self.statusBar().showMessage(f"手配工程 ID:{procurement_id} を表示しました", 5000)
        else:
            QMessageBox.information(self, "対象が見つかりません", "指定された手配工程は現在のデータに存在しません。再読み込みしてください。")

    def open_bottleneck_details(self, processing_details: str) -> None:
        """集計から各工程の問題一覧を表示し、選択された工程へ遷移する。"""
        dialog = BottleneckDetailsDialog(self.db, processing_details, self)
        if dialog.exec() == QDialog.Accepted and dialog.selected_procurement_id is not None:
            self.open_procurement_in_tree(dialog.selected_procurement_id)

    def refresh_all(self):
        self.orders.refresh(); self.bom.refresh(); self.procurements.refresh(); self.masters.refresh(); self.dashboard.refresh(); self.progress_tree.refresh()
        self.statusBar().showMessage(f"更新: {datetime.now().strftime('%Y/%m/%d %H:%M:%S')}  DB: {DB_PATH}",5000)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
