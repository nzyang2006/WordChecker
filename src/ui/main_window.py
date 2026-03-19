"""
主窗口界面 - 技术文档检查器（优化版）
"""
import sys
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTreeWidget,
    QTreeWidgetItem, QGroupBox, QProgressBar,
    QTabWidget, QTableWidget, QTableWidgetItem,
    QMessageBox, QSplitter, QHeaderView, QComboBox,
    QTextEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from core.rule_engine import RuleEngine
from models.check_result import DocumentCheckResult, CheckResult
from utils.result_exporter import ResultExporter


class CheckWorker(QThread):
    """检查工作线程"""
    progress = pyqtSignal(int, str)  # 进度百分比, 当前检查项
    log_message = pyqtSignal(str)  # 日志消息
    finished = pyqtSignal(object)  # DocumentCheckResult
    error = pyqtSignal(str)

    def __init__(self, engine: RuleEngine, document_path: str, rule_ids: list = None):
        super().__init__()
        self.engine = engine
        self.document_path = document_path
        self.rule_ids = rule_ids

    def run(self):
        """执行检查"""
        try:
            # 定义进度回调函数
            def progress_callback(current, total, rule_name):
                # 计算百分比
                percent = int((current / total) * 100)
                
                # 发送进度信号
                self.progress.emit(percent, rule_name)
                
                # 发送日志信号（打印到控制台）
                log_msg = f"[{current}/{total}] {rule_name} - {percent}%"
                self.log_message.emit(log_msg)
                
                # 打印到控制台，使用固定宽度避免残留
                # \033[K 清除从光标到行尾的内容
                print(f"\r\033[K{log_msg}", end="", flush=True)
            
            # 执行检查
            result = self.engine.check_document(
                self.document_path, 
                self.rule_ids,
                progress_callback=progress_callback
            )
            
            # 完成后换行
            print()  # 换行
            success_msg = f"[SUCCESS] Check completed! Total: {len(result.results)} items"
            self.log_message.emit(success_msg)
            print(success_msg)
            
            self.finished.emit(result)
        except Exception as e:
            error_msg = f"Check failed: {str(e)}"
            self.log_message.emit(f"[ERROR] {error_msg}")
            print(f"\n[ERROR] {error_msg}")
            self.error.emit(error_msg)


class MainWindow(QMainWindow):
    """主窗口"""

    # 定义类别展示顺序
    CATEGORY_ORDER = [
        "文档封面和签署页",
        "页码",
        "目次",
        "引用文件",
        "正文",
        "图表",
        "公式",
        "量、单位及其符号",
        "附录"
    ]

    def __init__(self):
        super().__init__()
        self.rule_engine = RuleEngine()
        self.current_document = None
        self.check_result = None
        self.check_worker = None

        self.init_ui()
        self.load_rules()

    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("技术文档检查器 v2.0")
        self.setGeometry(100, 100, 1400, 900)

        # 创建中心窗口
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 顶部工具栏 - 简化为一行
        toolbar = self.create_toolbar()
        main_layout.addWidget(toolbar)

        # 主内容区域
        content_layout = QHBoxLayout()

        # 左侧面板 - 规则选择
        left_panel = self.create_left_panel()

        # 右侧面板 - 结果显示
        right_panel = self.create_right_panel()

        # 使用分割器
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([400, 1000])

        content_layout.addWidget(splitter)
        main_layout.addLayout(content_layout)

        # 创建状态栏
        self.statusBar().showMessage("就绪 - 请选择Word文档并选择检查规则")

    def create_toolbar(self) -> QWidget:
        """创建顶部工具栏 - 简化版"""
        toolbar = QWidget()
        toolbar.setMaximumHeight(50)  # 限制高度
        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # 文件路径标签
        self.file_path_label = QLabel("未选择文件")
        self.file_path_label.setStyleSheet("""
            QLabel {
                color: #666;
                padding: 5px;
                background-color: #f5f5f5;
                border-radius: 3px;
                min-width: 250px;
            }
        """)
        layout.addWidget(self.file_path_label)

        # 选择文件按钮
        self.select_file_btn = QPushButton("选择文件")
        self.select_file_btn.clicked.connect(self.select_file)
        self.select_file_btn.setFixedHeight(32)
        self.select_file_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        layout.addWidget(self.select_file_btn)

        # 分隔符
        layout.addSpacing(20)

        # 开始检查按钮
        self.check_btn = QPushButton("开始检查")
        self.check_btn.setEnabled(False)
        self.check_btn.clicked.connect(self.start_check)
        self.check_btn.setFixedHeight(32)
        self.check_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px 20px;
                border-radius: 3px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        layout.addWidget(self.check_btn)

        # 导出格式选择
        self.export_combo = QComboBox()
        self.export_combo.addItems(["导出为Excel", "导出为HTML", "导出为JSON"])
        self.export_combo.setEnabled(False)
        self.export_combo.setFixedHeight(32)
        self.export_combo.setFixedWidth(120)
        layout.addWidget(self.export_combo)

        # 导出按钮
        self.export_btn = QPushButton("导出报告")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_result)
        self.export_btn.setFixedHeight(32)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        layout.addWidget(self.export_btn)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(28)
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 3px;
                text-align: center;
                background-color: #f5f5f5;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 2px;
            }
        """)
        layout.addWidget(self.progress_bar)

        layout.addStretch()

        return toolbar

    def create_left_panel(self) -> QWidget:
        """创建左侧面板 - 规则选择"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(5)

        # 规则树
        rules_group = QGroupBox("检查规则")
        rules_layout = QVBoxLayout(rules_group)

        # 按钮布局
        btn_layout = QHBoxLayout()

        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.setFixedHeight(28)
        self.select_all_btn.clicked.connect(self.select_all_rules)
        btn_layout.addWidget(self.select_all_btn)

        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.setFixedHeight(28)
        self.deselect_all_btn.clicked.connect(self.deselect_all_rules)
        btn_layout.addWidget(self.deselect_all_btn)

        rules_layout.addLayout(btn_layout)

        # 规则树形列表
        self.rules_tree = QTreeWidget()
        self.rules_tree.setHeaderLabel("检查规则（共0项）")
        self.rules_tree.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #ddd;
                font-size: 12px;
            }
            QTreeWidget::item {
                padding: 3px;
            }
            QTreeWidget::item:selected {
                background-color: #e3f2fd;
                color: black;
            }
        """)
        # 允许显示多行文本（用于显示描述）
        self.rules_tree.setWordWrap(True)
        rules_layout.addWidget(self.rules_tree)

        # 统计信息
        self.rules_stats_label = QLabel()
        self.rules_stats_label.setStyleSheet("color: #666; font-size: 11px; padding: 3px;")
        rules_layout.addWidget(self.rules_stats_label)

        layout.addWidget(rules_group)

        return panel

    def create_right_panel(self) -> QWidget:
        """创建右侧面板 - 结果显示"""
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 结果表格（去掉标签页，直接显示表格）
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["检查项", "状态", "详情", "位置"])

        # 设置列宽比例
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # 检查项 - 固定宽度
        header.setSectionResizeMode(1, QHeaderView.Fixed)  # 状态 - 固定宽度
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # 详情 - 自动拉伸
        header.setSectionResizeMode(3, QHeaderView.Interactive)  # 位置 - 可调整

        self.results_table.setColumnWidth(0, 180)  # 检查项宽度缩短
        self.results_table.setColumnWidth(1, 80)   # 状态宽度
        self.results_table.setColumnWidth(3, 150)  # 位置宽度

        self.results_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                gridline-color: #f0f0f0;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 5px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
        """)

        # 启用自动换行
        self.results_table.setWordWrap(True)
        self.results_table.setTextElideMode(Qt.ElideNone)

        layout.addWidget(self.results_table)

        return panel

    def load_rules(self):
        """加载规则到树形列表"""
        self.rules_tree.clear()
        rules = self.rule_engine.get_rules()

        # 按类别分组
        categories = {}
        for rule in rules:
            cat = rule.category
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(rule)

        # 按照指定顺序添加到树
        total_rules = 0
        for cat in self.CATEGORY_ORDER:
            if cat not in categories:
                continue

            cat_rules = categories[cat]

            # 创建类别节点
            cat_item = QTreeWidgetItem([f"{cat} ({len(cat_rules)}项)"])
            cat_item.setFlags(cat_item.flags() | Qt.ItemIsAutoTristate | Qt.ItemIsUserCheckable)
            cat_item.setCheckState(0, Qt.Unchecked)
            font = cat_item.font(0)
            font.setBold(True)
            cat_item.setFont(0, font)
            self.rules_tree.addTopLevelItem(cat_item)

            # 添加规则节点 - 显示名称和描述
            for rule in cat_rules:
                rule_item = QTreeWidgetItem([rule.name])
                rule_item.setFlags(rule_item.flags() | Qt.ItemIsUserCheckable)
                rule_item.setCheckState(0, Qt.Unchecked)
                rule_item.setData(0, Qt.UserRole, rule.id)

                # 在小字中显示描述
                desc_item = QTreeWidgetItem([f"  {rule.description}"])
                desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsSelectable)
                font = desc_item.font(0)
                font.setPointSize(9)
                font.setItalic(True)
                desc_item.setFont(0, font)
                desc_item.setForeground(0, QColor("#666"))

                # 设置为不可选中
                desc_item.setDisabled(True)

                rule_item.addChild(desc_item)
                cat_item.addChild(rule_item)
                total_rules += 1

            # 展开类别节点
            cat_item.setExpanded(True)

        self.rules_tree.setHeaderLabel(f"检查规则（共{total_rules}项）")

        # 默认全选所有规则
        self.select_all_rules()

        # 连接信号
        self.rules_tree.itemChanged.connect(self.on_rule_item_changed)

    def on_rule_item_changed(self, item, column):
        """规则项状态改变"""
        # 统计选中的规则数
        selected_count = 0
        total_count = 0

        root = self.rules_tree.invisibleRootItem()
        for i in range(root.childCount()):
            cat_item = root.child(i)
            for j in range(cat_item.childCount()):
                rule_item = cat_item.child(j)
                # 跳过描述项
                if rule_item.data(0, Qt.UserRole):
                    total_count += 1
                    if rule_item.checkState(0) == Qt.Checked:
                        selected_count += 1

        self.rules_stats_label.setText(f"已选择: {selected_count} / {total_count} 项")

    def select_all_rules(self):
        """全选所有规则"""
        root = self.rules_tree.invisibleRootItem()
        for i in range(root.childCount()):
            cat_item = root.child(i)
            cat_item.setCheckState(0, Qt.Checked)

    def deselect_all_rules(self):
        """取消选择所有规则"""
        root = self.rules_tree.invisibleRootItem()
        for i in range(root.childCount()):
            cat_item = root.child(i)
            cat_item.setCheckState(0, Qt.Unchecked)

    def select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )

        if file_path:
            self.current_document = file_path
            self.file_path_label.setText(Path(file_path).name)
            self.file_path_label.setStyleSheet("""
                QLabel {
                    color: #2196F3;
                    padding: 5px;
                    background-color: #e3f2fd;
                    border-radius: 3px;
                    font-weight: bold;
                }
            """)
            self.check_btn.setEnabled(True)
            self.statusBar().showMessage(f"已选择: {file_path}")

    def get_selected_rule_ids(self) -> list:
        """获取选中的规则ID列表"""
        rule_ids = []
        root = self.rules_tree.invisibleRootItem()

        for i in range(root.childCount()):
            cat_item = root.child(i)
            for j in range(cat_item.childCount()):
                rule_item = cat_item.child(j)
                # 跳过描述项（没有rule id的项）
                if rule_item.data(0, Qt.UserRole) and rule_item.checkState(0) == Qt.Checked:
                    rule_ids.append(rule_item.data(0, Qt.UserRole))

        return rule_ids

    def start_check(self):
        """开始检查"""
        if not self.current_document:
            QMessageBox.warning(self, "警告", "请先选择Word文档")
            return

        rule_ids = self.get_selected_rule_ids()
        if not rule_ids:
            QMessageBox.warning(self, "警告", "请至少选择一个检查规则")
            return

        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.check_btn.setEnabled(False)

        # 创建工作线程
        self.check_worker = CheckWorker(self.rule_engine, self.current_document, rule_ids)
        self.check_worker.progress.connect(self.update_progress)
        self.check_worker.finished.connect(self.on_check_finished)
        self.check_worker.error.connect(self.on_check_error)
        self.check_worker.start()

        self.statusBar().showMessage("正在检查文档...")

    def update_progress(self, percent, current_item):
        """更新进度"""
        self.progress_bar.setValue(percent)
        self.statusBar().showMessage(f"正在检查: {current_item}")

    def on_check_finished(self, result: DocumentCheckResult):
        """检查完成"""
        self.check_result = result
        self.progress_bar.setVisible(False)
        self.check_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.export_combo.setEnabled(True)

        # 显示结果
        self.display_results(result)

        self.statusBar().showMessage(f"检查完成 - 共 {len(result.results)} 项检查")

    def on_check_error(self, error_msg):
        """检查出错"""
        self.progress_bar.setVisible(False)
        self.check_btn.setEnabled(True)
        QMessageBox.critical(self, "错误", f"检查失败:\n{error_msg}")
        self.statusBar().showMessage("检查失败")

    def display_results(self, result: DocumentCheckResult):
        """显示检查结果"""
        # 清空表格
        self.results_table.setRowCount(0)

        # 添加结果
        for i, check_result in enumerate(result.results):
            row = self.results_table.rowCount()
            self.results_table.insertRow(row)

            # 检查项名称
            name_item = QTableWidgetItem(check_result.rule_name)
            self.results_table.setItem(row, 0, name_item)

            # 状态
            if check_result.passed:
                status_item = QTableWidgetItem("✓ 通过")
                status_item.setForeground(QColor("#4CAF50"))
            else:
                status_item = QTableWidgetItem("✗ 未通过")
                status_item.setForeground(QColor("#f44336"))
            status_item.setTextAlignment(Qt.AlignCenter)
            self.results_table.setItem(row, 1, status_item)

            # 详情 - 显示所有问题的详细信息
            if check_result.issues:
                details = "\n".join([f"• {issue.description}" for issue in check_result.issues])
            else:
                details = "符合要求"
            detail_item = QTableWidgetItem(details)
            detail_item.setTextAlignment(Qt.AlignLeft | Qt.AlignTop)
            self.results_table.setItem(row, 2, detail_item)

            # 位置
            if check_result.issues:
                locations = "\n".join([issue.position for issue in check_result.issues if issue.position])
            else:
                locations = "-"
            location_item = QTableWidgetItem(locations if locations else "-")
            location_item.setTextAlignment(Qt.AlignLeft | Qt.AlignTop)
            self.results_table.setItem(row, 3, location_item)

            # 设置行高（根据内容自动调整）
            self.results_table.resizeRowToContents(row)

    def export_result(self):
        """导出检查结果"""
        if not self.check_result:
            QMessageBox.warning(self, "警告", "没有可导出的检查结果")
            return

        # 获取导出格式
        export_format = self.export_combo.currentText()

        # 选择保存路径
        if "Excel" in export_format:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存Excel报告", "", "Excel文件 (*.xlsx)"
            )
            if file_path:
                try:
                    ResultExporter.export_to_excel(self.check_result, file_path)
                    QMessageBox.information(self, "成功", f"报告已导出到:\n{file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"导出失败:\n{str(e)}")

        elif "HTML" in export_format:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存HTML报告", "", "HTML文件 (*.html)"
            )
            if file_path:
                try:
                    ResultExporter.export_to_html(self.check_result, file_path)
                    QMessageBox.information(self, "成功", f"报告已导出到:\n{file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"导出失败:\n{str(e)}")

        elif "JSON" in export_format:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "保存JSON报告", "", "JSON文件 (*.json)"
            )
            if file_path:
                try:
                    ResultExporter.export_to_json(self.check_result, file_path)
                    QMessageBox.information(self, "成功", f"报告已导出到:\n{file_path}")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"导出失败:\n{str(e)}")


def main():
    """主函数"""
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import Qt

    # 设置高DPI支持
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # 创建应用
    app = QApplication(sys.argv)

    # 设置应用样式
    app.setStyle('Fusion')

    # 创建主窗口
    window = MainWindow()
    window.show()

    # 运行应用
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
