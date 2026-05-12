import sys
import threading

from PySide6.QtWidgets import (
    QApplication, QWidget, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QCheckBox, QGroupBox, QMessageBox, QProgressBar
)
from PySide6.QtCore import Qt

from UI_checktest.page import UI_STRUCTURE
from UI_checktest.multilanguage import ALL_LANGUAGES
from UI_checktest.run import main as run_test
from PySide6.QtWidgets import QTreeWidgetItem
from PySide6.QtCore import QThread, Signal

print("Qt UI 已启动")

class Worker(QThread):
    finished = Signal()
    error = Signal(str)
    status = Signal(str)

    def __init__(self, run_func, langs, pages):
        super().__init__()
        self.run_func = run_func
        self.langs = langs
        self.pages = pages

    def run(self):
        try:
            self.status.emit("运行中...")

            # ❌ 不允许UI操作
            self.run_func(self.langs, self.pages, status_callback=self.status.emit)

        except Exception as e:
            self.error.emit(str(e))
        finally:
            self.finished.emit()
# =========================
# 主窗口
# =========================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("UI检查工具 - Qt工业版")
        self.resize(1000, 750)

        self.lang_boxes = {}
        self.page_tree_items = []

        self._build_ui()

    # =========================
    # UI结构
    # =========================
    def _build_ui(self):
        root_layout = QVBoxLayout(self)

        title = QLabel("多语言 UI 自动检查（PySide6 工业版）")
        title.setStyleSheet("font-size:18px;font-weight:bold;")
        root_layout.addWidget(title)

        center_layout = QHBoxLayout()
        root_layout.addLayout(center_layout)

        # ========== 左：语言 ==========
        lang_box = QGroupBox("语言选择")
        lang_layout = QVBoxLayout()
        lang_box.setLayout(lang_layout)

        for lang_id, info in ALL_LANGUAGES.items():
            cb = QCheckBox(f"{lang_id}. {info[0]}")
            cb.setChecked(True)
            self.lang_boxes[int(lang_id)] = cb
            lang_layout.addWidget(cb)

        center_layout.addWidget(lang_box)

        # ========== 右：Tree ==========
        page_box = QGroupBox("页面选择（Tree）")
        page_layout = QVBoxLayout()
        page_box.setLayout(page_layout)

        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemChanged.connect(self._on_item_changed)

        page_layout.addWidget(self.tree)
        center_layout.addWidget(page_box)

        self._build_tree()

        # =========================
        # 按钮区
        # =========================
        btn_layout = QHBoxLayout()

        self.btn_start = QPushButton("开始运行")
        self.btn_start.clicked.connect(self._start)

        btn_all_lang = QPushButton("全选语言")
        btn_all_lang.clicked.connect(self._select_all_lang)

        btn_clear_lang = QPushButton("取消语言")
        btn_clear_lang.clicked.connect(self._clear_lang)

        btn_all_page = QPushButton("全选页面")
        btn_all_page.clicked.connect(self._select_all_page)

        btn_clear_page = QPushButton("取消页面")
        btn_clear_page.clicked.connect(self._clear_page)

        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(btn_all_lang)
        btn_layout.addWidget(btn_clear_lang)
        btn_layout.addWidget(btn_all_page)
        btn_layout.addWidget(btn_clear_page)

        root_layout.addLayout(btn_layout)
        # 状态
        self.status = QLabel("就绪")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        root_layout.addWidget(self.status)
        root_layout.addWidget(self.progress)

    # =========================
    # Tree构建
    # =========================
    def _build_tree(self):
        self.tree.clear()
        self.page_items = []

        for main_page, config in UI_STRUCTURE.items():
            main_item = self._create_item(main_page, True)
            self.tree.addTopLevelItem(main_item)

            for sub_name, actions in config.get("sub_pages", {}).items():
                sub_item = self._create_item(sub_name, True)
                main_item.addChild(sub_item)

                if isinstance(actions, dict):
                    for sub2 in actions.get("sub_pages", {}):
                        sub2_item = self._create_item(sub2, True)
                        sub_item.addChild(sub2_item)

        self.tree.expandAll()

    def _create_item(self, text, checked=False):
        item = QTreeWidgetItem([text])
        item.setFlags(
            Qt.ItemIsUserCheckable | 
            Qt.ItemIsEnabled  |
            Qt.ItemIsSelectable
        )
        item.setCheckState(0, Qt.Checked if checked else Qt.Unchecked)
        self.page_items.append(item)
        return item

    # =========================
    # Tree变化（Qt自动联动）
    # =========================
    def _on_item_changed(self, item, column):
        if column != 0:
            return

        state = item.checkState(0)

        # 防止递归触发
        self.tree.blockSignals(True)

        # =========================
        # 向下传播
        # =========================
        def set_children(parent, state):
            for i in range(parent.childCount()):
                child = parent.child(i)
                child.setCheckState(0, state)
                set_children(child, state)

        set_children(item, state)

        # =========================
        # 向上更新
        # =========================
        def update_parent(parent):
            if parent is None:
                return

            checked = 0
            partial = 0
            total = parent.childCount()

            for i in range(total):
                c = parent.child(i).checkState(0)
                if c == Qt.Checked:
                    checked += 1
                elif c == Qt.PartiallyChecked:
                    partial += 1

            if checked == total:
                parent.setCheckState(0, Qt.Checked)
            elif checked > 0 or partial > 0:
                parent.setCheckState(0, Qt.PartiallyChecked)
            else:
                parent.setCheckState(0, Qt.Unchecked)

            update_parent(parent.parent())

        update_parent(item.parent())

        self.tree.blockSignals(False)
    # =========================
    # 获取选中页面
    # =========================
    def get_selected_pages(self):
        selected = []

        def walk(item, prefix=""):
            for i in range(item.childCount()):
                child = item.child(i)
                if child.checkState(0) == Qt.Checked:
                    selected.append(child.text(0))
                walk(child, prefix)

        root = self.tree.invisibleRootItem()

        for i in range(root.childCount()):
            top = root.child(i)
            if top.checkState(0) == Qt.Checked:
                selected.append(top.text(0))
            walk(top)

        return selected

    # =========================
    # 语言
    # =========================
    def _select_all_lang(self):
        for cb in self.lang_boxes.values():
            cb.setChecked(True)

    # =========================
    # 页面
    # =========================
    def _select_all_page(self):
        def set_all(item):
            item.setCheckState(0, Qt.Checked)
            for i in range(item.childCount()):
                set_all(item.child(i))

        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            set_all(root.child(i))

    def _clear_lang(self):
        for cb in self.lang_boxes.values():
            cb.setChecked(False)

    def _clear_page(self):
        def set_all(item):
            item.setCheckState(0, Qt.Unchecked)
            for i in range(item.childCount()):
                set_all(item.child(i))

        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            set_all(root.child(i))

    # =========================
    # 执行
    # =========================
    def _start(self):

        langs = [k for k, v in self.lang_boxes.items() if v.isChecked()]
        pages = self.get_selected_pages()

        if not langs:
            QMessageBox.warning(self, "提示", "请选择语言")
            return

        if not pages:
            QMessageBox.warning(self, "提示", "请选择页面")
            return

        self.status.setText("运行中...")
        self.progress.setValue(0)

        # ===== QThread 启动 =====
        self.worker = Worker(
            run_func=run_test,
            langs=langs,
            pages=pages
        )

        # 信号绑定
        self.worker.status.connect(self._on_status)
        self.worker.finished.connect(self._on_done)
        self.worker.error.connect(self._on_error)
        self.btn_start.setEnabled(False)

        self.worker.start()


    def _on_status(self, msg):
        if msg.startswith('PROGRESS:'):
            parts = msg.split(':', 2)
            if len(parts) == 3 and parts[1].isdigit():
                self.progress.setValue(int(parts[1]))
                self.status.setText(parts[2])
                return
        self.status.setText(msg)


    def _on_done(self):
        self.status.setText("完成")
        self.progress.setValue(100)
        self.btn_start.setEnabled(True)
        


    def _on_error(self, msg):
        self.status.setText("失败")
        QMessageBox.critical(self, "错误", msg)
        self.btn_start.setEnabled(True)

    def _run(self, langs, pages):
        try:
            run_test(selected_langs=langs, selected_pages=pages)
            self.status.setText("完成")
        except Exception as e:
            self.status.setText(f"失败: {e}")
            QMessageBox.critical(self, "错误", str(e))


# =========================
# main
# =========================
def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()