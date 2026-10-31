# -*- coding: utf-8 -*-
"""JSON 批量转 Excel：模板字段树勾选 + 多文件导出"""

import json
import os
from collections import OrderedDict

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTreeWidget,
    QTreeWidgetItem, QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QGroupBox, QSplitter, QWidget, QAbstractItemView, QProgressDialog,
    QShortcut,
)
from openpyxl import Workbook


def _type_label(value):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, int):
        return "int"
    if isinstance(value, float):
        return "float"
    if isinstance(value, str):
        return "string"
    if isinstance(value, dict):
        return "object"
    if isinstance(value, list):
        return f"array[{len(value)}]"
    return type(value).__name__


def _preview(value, limit=48):
    if isinstance(value, (dict, list)):
        return ""
    text = "null" if value is None else str(value)
    text = text.replace("\n", " ")
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _path_to_header(path):
    parts = []
    for p in path:
        if p == "*":
            if parts:
                parts[-1] = parts[-1] + "[]"
            else:
                parts.append("[]")
        else:
            parts.append(str(p))
    return ".".join(parts)


def _default_custom_name(path):
    """默认列名：取路径最后一段有意义的字段名"""
    for p in reversed(path or []):
        if p != "*":
            return str(p)
    return _path_to_header(path) if path else ""


def _unique_headers(headers):
    """避免自定义列名重复导致 Excel 列混淆"""
    seen = {}
    result = []
    for h in headers:
        name = h if h else "列"
        if name not in seen:
            seen[name] = 1
            result.append(name)
        else:
            seen[name] += 1
            result.append(f"{name}_{seen[name]}")
    return result


def _sample_for_array(arr):
    """取数组中第一个非空对象作结构样例；否则取第一个非空元素。"""
    if not isinstance(arr, list) or not arr:
        return None
    for el in arr:
        if isinstance(el, dict):
            return el
    for el in arr:
        if el is not None:
            return el
    return None


def extract_rows(data, paths):
    """
    按带 '*' 的路径从 JSON 抽出表格行。
    paths: List[List] 例如 [['data','data','rows','*','organName'], ['timestamp']]
    返回 OrderedDict 列表（键为表头路径字符串）。
    """
    if not paths:
        return []

    headers = [_path_to_header(p) for p in paths]
    indexed = list(enumerate(paths))

    def navigate(node, key):
        if not isinstance(node, dict):
            return None
        return node.get(key)

    def expand(node, specs):
        """
        specs: list of (path_index, remaining_tokens)
        返回 list[dict[path_index -> value]]
        """
        if not specs:
            return [{}]

        # 若有路径已走完，记录当前节点为值，其余继续
        finished = []
        ongoing = []
        for idx, rem in specs:
            if not rem:
                finished.append((idx, node))
            else:
                ongoing.append((idx, rem))

        base = {idx: val for idx, val in finished}
        if not ongoing:
            return [base]

        # 是否需要在此展开数组
        if any(rem[0] == "*" for _, rem in ongoing):
            if not all(rem[0] == "*" for _, rem in ongoing):
                # 路径不对齐：有的要展开数组，有的还在取键——先按非 * 导航会更乱，
                # 约定：出现 * 时，当前 node 必须是 list
                pass
            items = node if isinstance(node, list) else ([] if node is None else [node])
            rows_out = []
            next_specs = [(idx, rem[1:]) for idx, rem in ongoing]
            for item in items:
                for partial in expand(item, next_specs):
                    merged = dict(base)
                    merged.update(partial)
                    rows_out.append(merged)
            return rows_out if rows_out else [base]

        # 按下一层 key 分组；同对象多字段分别递归后再合并为同一行
        keys_needed = OrderedDict()
        for idx, rem in ongoing:
            key = rem[0]
            keys_needed.setdefault(key, []).append((idx, rem[1:]))

        partial_groups = []
        for key, sub_specs in keys_needed.items():
            child = navigate(node, key)
            partial_groups.append(expand(child, sub_specs))

        rows_out = [dict(base)]
        for group in partial_groups:
            new_rows = []
            for left in rows_out:
                for right in group:
                    merged = dict(left)
                    merged.update(right)
                    new_rows.append(merged)
            rows_out = new_rows
        return rows_out

    raw_rows = expand(data, [(i, list(p)) for i, p in indexed])
    result = []
    for raw in raw_rows:
        row = OrderedDict()
        for i, header in enumerate(headers):
            val = raw.get(i)
            if isinstance(val, (dict, list)):
                val = json.dumps(val, ensure_ascii=False)
            row[header] = val
        result.append(row)
    return result


class JsonToExcelDialog(QDialog):
    # UserRole: 路径列表；UserRole+1: 是否可编辑自定义列名（叶子字段）
    ROLE_PATH = Qt.UserRole
    ROLE_RENAMEABLE = Qt.UserRole + 1

    def __init__(self, parent=None):
        super(JsonToExcelDialog, self).__init__(parent)
        self.setWindowTitle("JSON 转 Excel")
        self.resize(1280, 720)
        self.setMinimumSize(1100, 600)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)

        self.template_data = None
        self.json_files = []  # 有序去重路径列表
        self._custom_name_enabled = False

        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)

        tip = QLabel(
            "① 加载 JSON 模板并勾选导出字段；② 勾选「自定义字段名称」后，选中字段可编辑「自定义列名」；"
            "③ 添加文件/文件夹；④ 导出 Excel（表头使用自定义列名）。"
        )
        tip.setWordWrap(True)
        tip.setStyleSheet("color: #4b5563; font-size: 9pt; padding: 4px 0;")
        root.addWidget(tip)

        splitter = QSplitter(Qt.Horizontal)

        # ----- 左侧：模板字段树 -----
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)

        tpl_group = QGroupBox("字段模板（勾选导出列）")
        tpl_layout = QVBoxLayout(tpl_group)

        tpl_btns = QHBoxLayout()
        self.btn_load_tpl = QPushButton("加载 JSON 模板…")
        self.btn_load_tpl.clicked.connect(self.load_template)
        tpl_btns.addWidget(self.btn_load_tpl)
        self.btn_check_all = QPushButton("全选")
        self.btn_check_all.clicked.connect(lambda: self._set_all_checks(Qt.Checked))
        tpl_btns.addWidget(self.btn_check_all)
        self.btn_uncheck_all = QPushButton("全不选")
        self.btn_uncheck_all.clicked.connect(lambda: self._set_all_checks(Qt.Unchecked))
        tpl_btns.addWidget(self.btn_uncheck_all)
        tpl_btns.addSpacing(12)
        self.chk_custom_name = QPushButton("自定义字段名称：关")
        self.chk_custom_name.setCheckable(True)
        self.chk_custom_name.setChecked(False)
        self.chk_custom_name.setToolTip(
            "开启后：选中字段，双击或按 F2 编辑「自定义列名」；导出 Excel 时以此作为表头"
        )
        self.chk_custom_name.setStyleSheet(
            "QPushButton { background: #e2e8f0; color: #334155; border: none; "
            "border-radius: 6px; padding: 6px 12px; font-weight: 600; }"
            "QPushButton:checked { background: #2563eb; color: white; }"
        )
        self.chk_custom_name.toggled.connect(self._on_custom_name_toggled)
        tpl_btns.addWidget(self.chk_custom_name)
        self.btn_fill_names = QPushButton("用原字段名填充")
        self.btn_fill_names.setEnabled(False)
        self.btn_fill_names.setToolTip("将所有可导出字段的自定义列名重置为 JSON 原字段名")
        self.btn_fill_names.clicked.connect(self._fill_default_custom_names)
        tpl_btns.addWidget(self.btn_fill_names)
        tpl_btns.addStretch()
        tpl_layout.addLayout(tpl_btns)

        self.tpl_label = QLabel("未加载模板")
        self.tpl_label.setStyleSheet("color: #6b7280; font-size: 9pt;")
        tpl_layout.addWidget(self.tpl_label)

        self.custom_hint = QLabel("")
        self.custom_hint.setStyleSheet("color: #2563eb; font-size: 9pt;")
        self.custom_hint.hide()
        tpl_layout.addWidget(self.custom_hint)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["字段 / 结构", "自定义列名", "类型", "示例"])
        self.tree.setColumnWidth(0, 320)
        self.tree.setColumnWidth(1, 180)
        self.tree.setColumnWidth(2, 100)
        self.tree.setColumnWidth(3, 220)
        self.tree.setAlternatingRowColors(True)
        self.tree.setStyleSheet("QTreeWidget { font-size: 9pt; }")
        # 仅通过代码 editItem(..., 1) 编辑自定义列名，避免误改结构名
        self.tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tree.itemDoubleClicked.connect(self._on_tree_double_clicked)
        self.tree.itemChanged.connect(self._on_tree_item_changed)
        QShortcut(QKeySequence(Qt.Key_F2), self.tree, activated=self._edit_current_custom_name)
        tpl_layout.addWidget(self.tree)
        left_layout.addWidget(tpl_group)
        splitter.addWidget(left)

        # ----- 右侧：文件列表 -----
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)

        file_group = QGroupBox("待转换 JSON 文件")
        file_layout = QVBoxLayout(file_group)

        file_btns = QHBoxLayout()
        self.btn_add_files = QPushButton("添加文件…")
        self.btn_add_files.clicked.connect(self.add_files)
        file_btns.addWidget(self.btn_add_files)
        self.btn_add_folder = QPushButton("添加文件夹…")
        self.btn_add_folder.clicked.connect(self.add_folder)
        file_btns.addWidget(self.btn_add_folder)
        self.btn_remove = QPushButton("移除选中")
        self.btn_remove.clicked.connect(self.remove_selected_files)
        file_btns.addWidget(self.btn_remove)
        self.btn_clear_files = QPushButton("清空")
        self.btn_clear_files.clicked.connect(self.clear_files)
        file_btns.addWidget(self.btn_clear_files)
        file_btns.addStretch()
        file_layout.addLayout(file_btns)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.file_list.setStyleSheet("QListWidget { font-size: 9pt; }")
        file_layout.addWidget(self.file_list)

        self.file_count_label = QLabel("共 0 个文件")
        self.file_count_label.setStyleSheet("color: #6b7280; font-size: 9pt;")
        file_layout.addWidget(self.file_count_label)

        right_layout.addWidget(file_group)
        splitter.addWidget(right)

        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([860, 380])
        root.addWidget(splitter, 1)

        # ----- 底部操作 -----
        bottom = QHBoxLayout()
        bottom.addStretch()
        self.btn_export = QPushButton("一键导出 Excel…")
        self.btn_export.setStyleSheet(
            "QPushButton { background: #2563eb; color: white; font-weight: 600; "
            "padding: 8px 20px; border: none; border-radius: 6px; font-size: 10pt; }"
            "QPushButton:hover { background: #1d4ed8; }"
            "QPushButton:disabled { background: #93c5fd; }"
        )
        self.btn_export.clicked.connect(self.export_excel)
        bottom.addWidget(self.btn_export)
        self.btn_close = QPushButton("关闭")
        self.btn_close.clicked.connect(self.close)
        bottom.addWidget(self.btn_close)
        root.addLayout(bottom)

    # ========== 模板树 ==========
    def load_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 JSON 模板文件", "", "JSON 文件 (*.json);;所有文件 (*.*)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法读取模板：\n{e}")
            return
        self.template_data = data
        self.tpl_label.setText(f"模板：{path}")
        self._rebuild_tree(data)
        # 若尚未添加文件，可顺带加入该模板文件
        self._add_paths([path])

    def _on_custom_name_toggled(self, checked):
        self._custom_name_enabled = checked
        self.chk_custom_name.setText("自定义字段名称：开" if checked else "自定义字段名称：关")
        self.btn_fill_names.setEnabled(checked)
        if checked:
            self.custom_hint.setText(
                "已开启：选中叶子字段后，双击「自定义列名」列或按 F2 编辑；导出 Excel 时以此作表头。"
            )
            self.custom_hint.show()
        else:
            self.custom_hint.hide()
            cur = self.tree.currentItem()
            if cur is not None:
                self.tree.closePersistentEditor(cur, 1)

    def _edit_current_custom_name(self):
        item = self.tree.currentItem()
        if not self._custom_name_enabled or not item:
            return
        if not item.data(0, self.ROLE_RENAMEABLE):
            return
        self.tree.editItem(item, 1)

    def _on_tree_double_clicked(self, item, column):
        if not self._custom_name_enabled:
            return
        if not item or not item.data(0, self.ROLE_RENAMEABLE):
            return
        self.tree.editItem(item, 1)

    def _on_tree_item_changed(self, item, column):
        # 防止误改第 0 列结构名
        if column != 0:
            return
        original = item.data(0, Qt.UserRole + 2)
        if original is not None and item.text(0) != original:
            self.tree.blockSignals(True)
            item.setText(0, original)
            self.tree.blockSignals(False)

    def _fill_default_custom_names(self):
        def walk(item):
            if item.data(0, self.ROLE_RENAMEABLE):
                path = item.data(0, self.ROLE_PATH) or []
                item.setText(1, _default_custom_name(path))
            for i in range(item.childCount()):
                walk(item.child(i))

        self.tree.blockSignals(True)
        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))
        self.tree.blockSignals(False)

    def _make_tree_item(self, parent, labels, path, renameable=False, tristate=False):
        """labels: [结构名, 自定义列名, 类型, 示例]"""
        item = QTreeWidgetItem(parent, labels)
        item.setData(0, self.ROLE_PATH, path)
        item.setData(0, self.ROLE_RENAMEABLE, renameable)
        item.setData(0, Qt.UserRole + 2, labels[0])  # 原始结构名，防误改
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable
        if tristate:
            flags |= Qt.ItemIsTristate
        if renameable:
            flags |= Qt.ItemIsEditable
        item.setFlags(flags)
        item.setCheckState(0, Qt.Unchecked)
        return item

    def _rebuild_tree(self, data):
        self.tree.clear()
        root = self._make_tree_item(
            self.tree,
            ["(根对象)", "", _type_label(data), ""],
            [],
            renameable=False,
            tristate=True,
        )
        self._fill_tree(root, data, [])
        self.tree.expandToDepth(3)

    def _fill_tree(self, parent_item, value, path):
        if isinstance(value, dict):
            for key, child in value.items():
                child_path = path + [key]
                is_leaf = not isinstance(child, (dict, list))
                custom = _default_custom_name(child_path) if is_leaf else ""
                item = self._make_tree_item(
                    parent_item,
                    [str(key), custom, _type_label(child), _preview(child)],
                    child_path,
                    renameable=is_leaf,
                    tristate=isinstance(child, (dict, list)),
                )
                if isinstance(child, (dict, list)):
                    self._fill_tree(item, child, child_path)
        elif isinstance(value, list):
            arr_path = path + ["*"]
            sample = _sample_for_array(value)
            elem_label = f"[数组元素 × {len(value)}]"
            elem = self._make_tree_item(
                parent_item,
                [elem_label, "", _type_label(value), ""],
                arr_path,
                renameable=False,
                tristate=True,
            )
            if isinstance(sample, dict):
                self._fill_tree(elem, sample, arr_path)
            elif isinstance(sample, list):
                self._fill_tree(elem, sample, arr_path)
            elif sample is not None:
                self._make_tree_item(
                    elem,
                    ["(值)", _default_custom_name(arr_path), _type_label(sample), _preview(sample)],
                    arr_path,
                    renameable=True,
                    tristate=False,
                )

    def _set_all_checks(self, state):
        def walk(item):
            item.setCheckState(0, state)
            for i in range(item.childCount()):
                walk(item.child(i))

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))

    def get_checked_fields(self):
        """
        收集勾选的叶子字段。
        返回 list of (path, excel_header)
        excel_header：开启自定义时用自定义列名，否则用完整路径。
        """
        fields = []

        def walk(item):
            path = item.data(0, self.ROLE_PATH)
            if item.childCount() == 0:
                if item.checkState(0) == Qt.Checked and path:
                    if self._custom_name_enabled:
                        custom = (item.text(1) or "").strip()
                        header = custom if custom else _default_custom_name(path)
                    else:
                        header = _path_to_header(path)
                    fields.append((list(path), header))
                return
            for i in range(item.childCount()):
                walk(item.child(i))

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))

        seen = set()
        unique = []
        for path, header in fields:
            key = tuple(path)
            if key not in seen:
                seen.add(key)
                unique.append((path, header))
        return unique

    # ========== 文件列表 ==========
    def _add_paths(self, paths):
        added = 0
        existing = set(self.json_files)
        for p in paths:
            p = os.path.normpath(p)
            if not p.lower().endswith(".json"):
                continue
            if not os.path.isfile(p):
                continue
            if p in existing:
                continue
            existing.add(p)
            self.json_files.append(p)
            self.file_list.addItem(QListWidgetItem(p))
            added += 1
        self._refresh_file_count()
        return added

    def _refresh_file_count(self):
        self.file_count_label.setText(f"共 {len(self.json_files)} 个文件")

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择 JSON 文件", "", "JSON 文件 (*.json);;所有文件 (*.*)"
        )
        if files:
            n = self._add_paths(files)
            if n == 0:
                QMessageBox.information(self, "提示", "没有新增文件（可能已在列表中）。")

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择包含 JSON 的文件夹")
        if not folder:
            return
        found = []
        for root_dir, _dirs, files in os.walk(folder):
            for name in files:
                if name.lower().endswith(".json"):
                    found.append(os.path.join(root_dir, name))
        found.sort()
        n = self._add_paths(found)
        QMessageBox.information(
            self, "添加完成", f"在文件夹中找到 {len(found)} 个 JSON，新增 {n} 个。"
        )

    def remove_selected_files(self):
        rows = sorted({idx.row() for idx in self.file_list.selectedIndexes()}, reverse=True)
        for row in rows:
            self.file_list.takeItem(row)
            if 0 <= row < len(self.json_files):
                del self.json_files[row]
        self._refresh_file_count()

    def clear_files(self):
        self.file_list.clear()
        self.json_files = []
        self._refresh_file_count()

    # ========== 导出 ==========
    def export_excel(self):
        fields = self.get_checked_fields()
        if not fields:
            QMessageBox.warning(self, "提示", "请先加载模板并勾选至少一个字段。")
            return
        if not self.json_files:
            QMessageBox.warning(self, "提示", "请先添加要转换的 JSON 文件。")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存 Excel",
            "export.xlsx",
            "Excel 工作簿 (*.xlsx)",
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        paths = [p for p, _ in fields]
        # 内部仍用路径键取数；写出时用自定义表头
        path_headers = [_path_to_header(p) for p in paths]
        display_headers = _unique_headers([h for _, h in fields])
        all_rows = []
        errors = []

        progress = QProgressDialog("正在导出…", "取消", 0, len(self.json_files), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(300)

        for i, fp in enumerate(self.json_files):
            progress.setValue(i)
            progress.setLabelText(f"处理中 ({i + 1}/{len(self.json_files)})\n{os.path.basename(fp)}")
            if progress.wasCanceled():
                QMessageBox.information(self, "已取消", "导出已取消。")
                return
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    data = json.load(f)
                rows = extract_rows(data, paths)
                if not rows:
                    rows = [OrderedDict((h, None) for h in path_headers)]
                for row in rows:
                    ordered = OrderedDict()
                    for path_h, disp_h in zip(path_headers, display_headers):
                        ordered[disp_h] = row.get(path_h)
                    all_rows.append(ordered)
            except Exception as e:
                errors.append(f"{os.path.basename(fp)}: {e}")

        progress.setValue(len(self.json_files))

        if not all_rows:
            QMessageBox.warning(self, "提示", "没有可导出的数据。\n" + "\n".join(errors[:5]))
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "data"
            ws.append(display_headers)
            for row in all_rows:
                ws.append([row.get(h) for h in display_headers])
            wb.save(save_path)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))
            return

        msg = (
            f"导出成功！\n\n文件：{save_path}\n"
            f"行数：{len(all_rows)}（含表头共 {len(all_rows) + 1} 行）\n"
            f"列数：{len(display_headers)}"
        )
        if self._custom_name_enabled:
            msg += "\n表头：已使用自定义字段名称"
        if errors:
            msg += f"\n\n有 {len(errors)} 个文件失败：\n" + "\n".join(errors[:8])
            if len(errors) > 8:
                msg += f"\n… 另有 {len(errors) - 8} 个"
        QMessageBox.information(self, "完成", msg)
