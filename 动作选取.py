import sys
import os
import random
import traceback
import json
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMainWindow, QPushButton, QListWidget, 
    QTextEdit, QVBoxLayout, QHBoxLayout, QLabel, QFileDialog, 
    QMessageBox, QSplitter, QScrollArea, QFrame, QCheckBox, 
    QInputDialog, QDialog, QGridLayout, QAbstractItemView,
    QStyleFactory, QLayout, QSizePolicy, QTableWidget, 
    QTableWidgetItem, QHeaderView, QSpinBox, QGroupBox, QToolButton
)
from PyQt5.QtGui import QPalette, QColor, QFont, QCursor
from PyQt5.QtCore import Qt, pyqtSignal, QSize, QEvent

# --- 辅助函数：生成圆圈数字 ---
def get_circled_num(n):
    if 1 <= n <= 20: return chr(9311 + n)
    elif 21 <= n <= 35: return chr(12860 + n)
    elif 36 <= n <= 50: return chr(12976 + n)
    else: return f"({n})"

# --- 默认配置 ---
DEFAULT_COL_MAPPING = {
  "第一类动作（A/C列）": [0, 2],
  "第二类动作（E/G列）": [4, 6],
  "第三类动作（I列）": {"cols": [8], "min_c": 4},
  "第四类动作（K列）": [10],
  "第五类动作（M/O列）": [12, 14],
  "辅助动作（S/U/W列）": [18, 20, 22],
  "表情（Y列）": [24]
}
DEFAULT_MIN_C = 3          
REPEAT_MIN = 3             
REPEAT_MAX = 5             

# --- [需求2 & 3] 现代 UI 主题设置 ---

def set_light_theme(app: QApplication):
    app.setStyle("Fusion")
    
    # 浅色模式调色板
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(245, 247, 250))       # 背景：极淡的灰蓝白
    palette.setColor(QPalette.WindowText, QColor(48, 49, 51))      # 文字：深灰
    palette.setColor(QPalette.Base, QColor(255, 255, 255))         # 输入框背景：纯白
    palette.setColor(QPalette.AlternateBase, QColor(250, 250, 250))
    palette.setColor(QPalette.Text, QColor(48, 49, 51))
    palette.setColor(QPalette.Button, QColor(255, 255, 255))       
    palette.setColor(QPalette.ButtonText, QColor(96, 98, 102))
    palette.setColor(QPalette.Highlight, QColor(64, 158, 255))     # 强调色：现代蓝
    palette.setColor(QPalette.HighlightedText, Qt.white)
    app.setPalette(palette)

    # 全局 QSS
    qss = """
    QWidget {
        font-family: "Calibri", "SimHei"; 
        font-size: 10pt;
    }
    
    /* 现代卡片风格 Frame */
    QFrame#CardFrame {
        background-color: #FFFFFF;
        border: 1px solid #EBEEF5;
        border-radius: 8px;
    }
    
    /* 主按钮 */
    QPushButton {
        background-color: #FFFFFF;
        border: 1px solid #DCDFE6;
        border-radius: 6px;
        padding: 6px 12px;
        color: #606266;
    }
    QPushButton:hover {
        color: #409EFF;
        border-color: #C6E2FF;
        background-color: #ECF5FF;
    }
    QPushButton:pressed {
        color: #3A8EE6;
        border-color: #3A8EE6;
        background-color: #E6F1FC;
    }
    QPushButton:checked {
        color: #FFFFFF;
        background-color: #409EFF;
        border-color: #409EFF;
    }

    /* 胶囊/标签风格按钮 (用于选择窗口) */
    QPushButton#TagButton {
        background-color: #F4F4F5;
        border: 1px solid #E9E9EB;
        border-radius: 15px; /* 圆润 */
        color: #909399;
        padding: 5px 15px;
        font-weight: bold;
        text-align: center;
    }
    QPushButton#TagButton:hover {
        background-color: #E6F1FC;
        color: #409EFF;
        border-color: #C6E2FF;
    }
    QPushButton#TagButton:checked {
        background-color: #409EFF;
        color: #FFFFFF;
        border-color: #409EFF;
    }
    
    /* 输入框和列表 */
    QTextEdit, QListWidget, QTableWidget {
        border: 1px solid #DCDFE6;
        border-radius: 6px;
        background-color: #FFFFFF;
        padding: 5px;
    }
    
    QTableWidget::item:selected {
        background-color: #ECF5FF;
        color: #409EFF;
    }

    QScrollBar:vertical {
        border: none;
        background: transparent;
        width: 8px;
        margin: 0px;
    }
    QScrollBar::handle:vertical {
        background: #C0C4CC;
        min-height: 20px;
        border-radius: 4px;
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
    """
    app.setStyleSheet(qss)

def set_dark_theme(app: QApplication):
    app.setStyle("Fusion")
    
    # [需求3] 优化夜间模式配色 (Modern Dark)
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(32, 33, 36))          # 背景：深炭色 (Google Dark)
    palette.setColor(QPalette.WindowText, QColor(224, 224, 224))   # 文字：灰白，不刺眼
    palette.setColor(QPalette.Base, QColor(45, 45, 48))            # 输入框：稍亮
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.Text, QColor(224, 224, 224))
    palette.setColor(QPalette.Button, QColor(45, 45, 48))
    palette.setColor(QPalette.ButtonText, QColor(224, 224, 224))
    palette.setColor(QPalette.Highlight, QColor(64, 158, 255))     # 保持蓝色高亮
    palette.setColor(QPalette.HighlightedText, Qt.white)
    palette.setColor(QPalette.Link, QColor(100, 181, 246))
    app.setPalette(palette)
    
    qss = """
    QWidget {
        font-family: "Calibri", "SimHei";
        font-size: 10pt;
    }
    
    QFrame#CardFrame {
        background-color: #2D2D30; /* 卡片背景 */
        border: 1px solid #3E3E42;
        border-radius: 8px;
    }
    
    QPushButton {
        background-color: #363636;
        border: 1px solid #4F4F4F;
        border-radius: 6px;
        padding: 6px 12px;
        color: #E0E0E0;
    }
    QPushButton:hover {
        background-color: #404040;
        border-color: #606060;
    }
    QPushButton:checked {
        background-color: #1e4f7a; /* 深蓝选中态 */
        border-color: #409EFF;
        color: #FFFFFF;
    }

    /* 夜间模式 Tag 按钮 */
    QPushButton#TagButton {
        background-color: #2D2D30;
        border: 1px solid #3E3E42;
        border-radius: 15px;
        color: #A0A0A0;
    }
    QPushButton#TagButton:hover {
        background-color: #3E3E42;
        color: #FFFFFF;
    }
    QPushButton#TagButton:checked {
        background-color: #164c7e;
        border-color: #409EFF;
        color: #FFFFFF;
    }
    
    QTextEdit, QListWidget, QTableWidget {
        border: 1px solid #3E3E42;
        background-color: #252526; /* 编辑器深色背景 */
        color: #E0E0E0;
    }
    
    QTableWidget::item:selected {
        background-color: #1e3a5f;
    }
    
    QScrollBar::handle:vertical {
        background: #555555;
    }
    """
    app.setStyleSheet(qss)

# --- 自定义组件 ---
class TruncatedCheckBox(QWidget):
    """
    使用 QCheckBox + QLabel 组合，以支持 HTML 渲染
    """
    def __init__(self, full_code, translation, parent=None):
        super().__init__(parent)
        self.full_code = full_code
        self.translation = translation
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8) # 稍微增加间距
        
        self.checkbox = QCheckBox()
        self.checkbox.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        
        self.label = QLabel()
        self.label.setTextFormat(Qt.RichText) 
        self.label.setCursor(QCursor(Qt.PointingHandCursor))
        self.label.setStyleSheet("border: none; background: transparent;") # 确保标签透明
        
        display_code = full_code
        if len(display_code) > 25: 
            display_code = display_code[:23] + "..."
            
        # 使用 css class 或默认颜色，不在 html 中硬编码颜色，以便适配夜间模式
        label_text = f"<b>{display_code}</b>"
        if translation:
            label_text += f" <span style='font-weight:normal; opacity: 0.7;'>({translation})</span>"
            
        self.label.setText(label_text)
        self.label.setToolTip(f"{full_code}\n{translation}") 

        layout.addWidget(self.checkbox)
        layout.addWidget(self.label)
        layout.addStretch() 
        
        self.setLayout(layout)
        self.checkbox.stateChanged.connect(self.on_state_changed)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.checkbox.setChecked(not self.checkbox.isChecked())
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        dialog = QDialog(self)
        dialog.setWindowTitle("完整内容查看")
        dialog.resize(400, 300)
        # [需求2] 美化双击查看窗口
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        info = QLabel(f"<b>动作代码:</b>")
        info.setStyleSheet("font-size: 14px; margin-bottom: 5px;")
        layout.addWidget(info)
        
        content = QTextEdit()
        content.setPlainText(self.full_code)
        content.setReadOnly(True)
        content.setStyleSheet("font-size: 13px; color: #409EFF;")
        layout.addWidget(content)
        
        trans_lbl = QLabel(f"<b>翻译:</b> {self.translation}" if self.translation else "<b>翻译:</b> 无")
        trans_lbl.setStyleSheet("margin-top: 10px; font-size: 13px;")
        trans_lbl.setWordWrap(True)
        layout.addWidget(trans_lbl)
        
        btn = QPushButton("关闭")
        btn.clicked.connect(dialog.accept)
        layout.addWidget(btn)
        
        dialog.exec_()

    def on_state_changed(self, state): pass 
    def isChecked(self): return self.checkbox.isChecked()
    def setChecked(self, val): self.checkbox.setChecked(val)
    def property(self, name):
        if name == "code": return self.full_code
        return super().property(name)


class MainWindow(QMainWindow):
    def __init__(self, app: QApplication):
        super().__init__()
        self.app = app
        self.dark_mode = False
        
        self.excel_data = None
        self.current_excel_path = None
        self.translation_map = {}
        self.combined_plan = [] 
        
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.config_data = self.load_config()
        self.col_mapping = self.config_data.get("mapping", DEFAULT_COL_MAPPING)
        
        self.action_categories = {}
        for k in self.col_mapping.keys():
            if "表情" in k: self.action_categories[k] = []
            elif "辅助" in k: self.action_categories[k] = {}
            else: self.action_categories[k] = {}
            
        self.category_widgets = {} 
        self.use_original_text = False

        self.init_ui()
        
        saved_path = self.config_data.get("excel_path", "")
        if saved_path and os.path.exists(saved_path):
            self.load_excel_file(saved_path)
        else:
            self.file_label.setText("请选择 Excel 文件")

    def load_config(self):
        config_path = os.path.join(self.base_dir, "config.json")
        default_config = {
            "excel_path": "",
            "mapping": DEFAULT_COL_MAPPING,
            "last_export_dir": "" 
        }
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    if "mapping" not in config: config["mapping"] = DEFAULT_COL_MAPPING
                    return config
            except Exception: return default_config
        else:
            return default_config

    def save_config(self):
        config_path = os.path.join(self.base_dir, "config.json")
        last_export = self.config_data.get("last_export_dir", "")
        
        data = {
            "excel_path": self.current_excel_path if self.current_excel_path else "",
            "mapping": self.col_mapping,
            "last_export_dir": last_export
        }
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except PermissionError:
            QMessageBox.warning(self, "配置保存失败", f"无法写入配置文件:\n{config_path}\n请检查权限或文件是否被占用。")
        except Exception as e:
            print(f"保存配置失败: {e}")

    def parse_config_setting(self, cat_name):
        setting = self.col_mapping.get(cat_name)
        final_cols = []
        final_min_c = DEFAULT_MIN_C
        if setting is None: return final_cols, final_min_c
        if isinstance(setting, dict):
            final_cols = setting.get("cols", [])
            final_min_c = setting.get("min_c", DEFAULT_MIN_C)
        elif isinstance(setting, list):
            final_cols = setting
            final_min_c = DEFAULT_MIN_C
        return final_cols, final_min_c

    def init_ui(self):
        self.setWindowTitle("动作训练计划生成器 (Modern UI)")
        self.resize(1600, 950)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        top_bar = QHBoxLayout()
        self.theme_btn = QPushButton("切换深色 / 浅色模式")
        self.theme_btn.clicked.connect(self.toggle_theme)
        self.lang_toggle = QCheckBox("左侧显示原始Prompt")
        self.lang_toggle.stateChanged.connect(self.toggle_language_display)
        top_bar.addStretch()
        top_bar.addWidget(self.lang_toggle)
        top_bar.addWidget(self.theme_btn)
        main_layout.addLayout(top_bar)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(2) # 细分割线
        main_layout.addWidget(splitter)

        # --- 左侧面板 ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0) # 右侧留点空隙
        
        file_layout = QHBoxLayout()
        self.file_label = QLabel("等待加载...")
        self.file_label.setStyleSheet("color: #909399; font-style: italic;")
        btn_open = QPushButton("更改 Excel")
        btn_open.clicked.connect(self.change_excel_path)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(btn_open)
        left_layout.addLayout(file_layout)

        ctrl_layout = QHBoxLayout()
        self.cat_list = QListWidget() 
        self.cat_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.cat_list.setFixedHeight(100)
        self.cat_list.model().rowsMoved.connect(self.sync_category_ui_order)
        initial_cats = [k for k in self.col_mapping.keys() if "辅助" not in k and "表情" not in k]
        self.cat_list.addItems(initial_cats)
        
        btn_box = QVBoxLayout()
        btn_reset = QPushButton("重置所有")
        btn_reset.clicked.connect(self.reset_all_actions)
        btn_gen = QPushButton("生成 Prompt")
        btn_gen.clicked.connect(self.generate_plan)
        # [需求2] 按钮高亮样式
        btn_gen.setStyleSheet("background-color: #409EFF; color: white; border-color: #409EFF; font-weight: bold;")
        
        btn_export = QPushButton("导出 Excel")
        btn_export.clicked.connect(self.export_excel)
        btn_box.addWidget(btn_reset)
        btn_box.addWidget(btn_gen)
        btn_box.addWidget(btn_export)
        
        ctrl_layout.addWidget(self.cat_list, 1)
        ctrl_layout.addLayout(btn_box, 1)
        left_layout.addLayout(ctrl_layout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        self.cats_container_widget = QWidget()
        self.cats_layout = QVBoxLayout(self.cats_container_widget)
        self.cats_layout.setAlignment(Qt.AlignTop)
        # [需求1] 增加左侧间距，确保内容不贴边
        self.cats_layout.setContentsMargins(5, 5, 5, 5)
        self.cats_layout.setSpacing(15) 
        
        scroll.setWidget(self.cats_container_widget)
        left_layout.addWidget(scroll)
        
        self.status_label = QLabel("当前总计: 0 张")
        self.status_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #409EFF; padding: 5px;")
        left_layout.addWidget(self.status_label)
        self.refresh_category_widgets()

        # --- 右侧面板 ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(10, 0, 0, 0) # 左侧留点空隙

        tool_bar = QHBoxLayout()
        
        btn_all = QPushButton("全部全选")
        btn_all.clicked.connect(lambda: self.batch_check(True))

        btn_sel_drag = QPushButton("选择高亮")
        btn_sel_drag.setToolTip("将鼠标拖蓝选中的行打上勾")
        btn_sel_drag.clicked.connect(self.select_dragged_rows)
        
        btn_inv = QPushButton("反选")
        btn_inv.clicked.connect(self.global_invert_selection) 
        
        btn_up = QPushButton("▲")
        btn_up.clicked.connect(lambda: self.move_prompt(-1))
        btn_down = QPushButton("▼")
        btn_down.clicked.connect(lambda: self.move_prompt(1))
        
        btn_copy = QPushButton("复制")
        btn_copy.clicked.connect(self.copy_selected_prompt)
        
        btn_rm_tag = QPushButton("删除 Tag")
        btn_rm_tag.clicked.connect(self.remove_specific_tag)
        
        btn_del = QPushButton("删除行")
        btn_del.clicked.connect(self.delete_selected_prompt)
        
        btn_tag = QPushButton("添加 Tag")
        btn_tag.clicked.connect(self.add_extra_prompt)
        btn_aux = QPushButton("添加辅助/表情")
        btn_aux.clicked.connect(self.open_add_action_window)
        # 按钮样式微调
        btn_aux.setStyleSheet("font-weight: bold;")

        tool_bar.addWidget(btn_all)
        tool_bar.addWidget(btn_sel_drag) 
        tool_bar.addWidget(btn_inv)
        tool_bar.addSpacing(10)
        tool_bar.addWidget(btn_up)
        tool_bar.addWidget(btn_down)
        tool_bar.addSpacing(10)
        tool_bar.addWidget(btn_copy)
        tool_bar.addWidget(btn_rm_tag) 
        tool_bar.addWidget(btn_del)
        tool_bar.addSpacing(10)
        tool_bar.addWidget(btn_tag)
        tool_bar.addWidget(btn_aux)
        tool_bar.addStretch()
        
        self.count_label = QLabel("Prompt数: 0")
        tool_bar.addWidget(self.count_label)
        right_layout.addLayout(tool_bar)

        self.prompt_table = QTableWidget()
        self.prompt_table.setColumnCount(5)
        self.prompt_table.setHorizontalHeaderLabels(["序号", "选择", "标签", "操作", "内容"])
        self.prompt_table.verticalHeader().setMinimumSectionSize(45) # 增加行高，更现代
        self.prompt_table.setWordWrap(True)
        self.prompt_table.setTextElideMode(Qt.ElideNone)
        self.prompt_table.scrollTo = lambda index, hint=None: None
        self.prompt_table.setAutoScroll(False)
        self.prompt_table.setFocusPolicy(Qt.NoFocus)

        header = self.prompt_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents) 
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents) 
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents) 
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents) 
        header.setSectionResizeMode(4, QHeaderView.Stretch) 
        
        self.prompt_table.verticalHeader().setVisible(False)
        self.prompt_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.prompt_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.prompt_table.setShowGrid(False) 
        self.prompt_table.setAlternatingRowColors(True) 
        
        self.prompt_table.cellDoubleClicked.connect(self.open_prompt_editor)
        
        right_layout.addWidget(self.prompt_table)

        left_panel.setMinimumWidth(600) 
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([850, 650])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

    # --- 逻辑功能 ---

    def select_dragged_rows(self):
        selected_indexes = self.prompt_table.selectedIndexes()
        if not selected_indexes: return
        rows = set(index.row() for index in selected_indexes)
        for row in rows:
            if 0 <= row < len(self.combined_plan):
                self.combined_plan[row]['checked'] = True
        self.refresh_prompt_list()

    def remove_specific_tag(self):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: QMessageBox.information(self, "提示", "请先勾选需要处理的行"); return
        text, ok = QInputDialog.getText(self, "删除 Tag", "请输入要删除的 tag (区分大小写)：")
        if ok and text:
            text = text.strip()
            if not text: return
            for i in indices:
                orig = self.combined_plan[i]["original_action"]
                parts = [p.strip() for p in orig.split(',')]
                new_parts = [p for p in parts if p != text]
                self.combined_plan[i]["original_action"] = ", ".join(new_parts)
            self.refresh_prompt_list()

    def open_prompt_editor(self, row, column):
        if row < 0 or row >= len(self.combined_plan): return
        item_data = self.combined_plan[row]
        current_text = item_data["original_action"]
        dialog = QDialog(self)
        dialog.setWindowTitle(f"编辑 Prompt (第 {row+1} 条)")
        dialog.resize(600, 400)
        layout = QVBoxLayout(dialog)
        label = QLabel("Prompt 内容 (双击此处可修改):")
        layout.addWidget(label)
        text_edit = QTextEdit()
        text_edit.setPlainText(current_text)
        text_edit.setFont(QFont("Calibri", 12))
        layout.addWidget(text_edit)
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("保存修改")
        btn_cancel = QPushButton("取消")
        def save():
            new_text = text_edit.toPlainText().strip()
            if new_text:
                self.combined_plan[row]["original_action"] = new_text
                self.refresh_prompt_list()
                dialog.accept()
        btn_save.clicked.connect(save)
        btn_cancel.clicked.connect(dialog.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        dialog.exec_()

    def global_invert_selection(self):
        for i in range(len(self.combined_plan)):
            self.combined_plan[i]['checked'] = not self.combined_plan[i]['checked']
        self.refresh_prompt_list()

    def toggle_theme(self):
        if self.dark_mode: set_light_theme(self.app)
        else: set_dark_theme(self.app)
        self.dark_mode = not self.dark_mode
        self.update_ui_display()
        self.refresh_prompt_list()

    def toggle_language_display(self, state):
        self.use_original_text = (state == Qt.Checked)
        self.update_ui_display()
        self.refresh_prompt_list()

    def change_excel_path(self):
        start_dir = self.base_dir
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", start_dir, "Excel Files (*.xlsx)")
        if path: self.load_excel_file(path)

    def load_excel_file(self, path):
        try:
            self.current_excel_path = path
            self.save_config()
            self.excel_data = pd.read_excel(path, header=None)
            self.translation_map = {}
            for cat in self.col_mapping.keys():
                _, min_c = self.parse_config_setting(cat)
                self.process_category_data(cat, target_c_count=min_c)
            self.update_ui_display()
            self.file_label.setText(os.path.basename(path))
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法读取文件: {str(e)}")

    def process_category_data(self, cat_name, target_c_count=DEFAULT_MIN_C):
        col_indices, _ = self.parse_config_setting(cat_name)
        is_aux = "辅助" in cat_name
        is_emo = "表情" in cat_name

        if is_emo: self.action_categories[cat_name] = []
        elif is_aux: self.action_categories[cat_name] = {}
        else: self.action_categories[cat_name] = {}
        if self.excel_data is None: return

        all_actions_pool = [] 
        for col_idx in col_indices:
            if col_idx >= self.excel_data.shape[1]: continue
            col_data = self.excel_data.iloc[:, col_idx]
            trans_data = None
            if col_idx + 1 < self.excel_data.shape[1]: trans_data = self.excel_data.iloc[:, col_idx + 1]
            if len(col_data) > 0: sub_cat = str(col_data.iloc[0]).strip()
            else: continue
            for i in range(1, len(col_data)):
                act_val = col_data.iloc[i]
                if pd.notna(act_val):
                    act = str(act_val).strip()
                    if act:
                        all_actions_pool.append((act, sub_cat))
                        if trans_data is not None and i < len(trans_data):
                            t_val = trans_data.iloc[i]
                            if pd.notna(t_val):
                                t_str = str(t_val).strip()
                                if t_str and t_str.lower() != 'nan':
                                    self.translation_map[act] = t_str
        if not all_actions_pool: return

        if is_emo: 
            self.action_categories[cat_name] = [x[0] for x in all_actions_pool]
        elif is_aux: 
            for act, sub in all_actions_pool:
                if sub not in self.action_categories[cat_name]:
                    self.action_categories[cat_name][sub] = {}
                self.action_categories[cat_name][sub][act] = 1
        else:
            final_selection = []
            shuffled_pool = all_actions_pool.copy()
            random.shuffle(shuffled_pool)
            unique_actions = {} 
            for a, s in shuffled_pool:
                if a not in unique_actions: unique_actions[a] = s
            unique_keys = list(unique_actions.keys())
            count_to_take = min(target_c_count, len(unique_keys))
            selected_keys = unique_keys[:count_to_take]

            for act in selected_keys:
                sub = unique_actions[act]
                repeat = random.randint(REPEAT_MIN, REPEAT_MAX)
                final_selection.append({'action': act, 'count': repeat, 'sub': sub})
            for _, sub in all_actions_pool:
                if sub not in self.action_categories[cat_name]:
                    self.action_categories[cat_name][sub] = []
            for item in final_selection:
                sub = item['sub']
                self.action_categories[cat_name][sub].append({item['action']: item['count']})

    def refresh_category_widgets(self):
        for i in reversed(range(self.cats_layout.count())): 
            self.cats_layout.itemAt(i).widget().setParent(None)
        self.category_widgets = {}

        order = [self.cat_list.item(i).text() for i in range(self.cat_list.count())]
        for cat_name in order:
            _, min_c_val = self.parse_config_setting(cat_name)
            
            # 使用 ObjectName 应用 CardFrame 样式
            frame = QFrame()
            frame.setObjectName("CardFrame")
            
            f_layout = QVBoxLayout(frame)
            h_layout = QHBoxLayout()
            lbl = QLabel(cat_name)
            lbl.setFont(QFont("SimHei", 10, QFont.Bold)) 
            lbl.setStyleSheet("border: none; color: #303133;" if not self.dark_mode else "border: none; color: #E0E0E0;")
            lbl.setMinimumWidth(150)
            
            lbl_c = QLabel("保底类别:")
            lbl_c.setStyleSheet("border: none; color: #606266;" if not self.dark_mode else "border: none; color: #B0B0B0;")
            spin_c = QSpinBox()
            spin_c.setRange(1, 50) 
            spin_c.setValue(min_c_val) 
            spin_c.setFixedWidth(50) 
            
            lbl_cur_c = QLabel("(当前: 0类)")
            lbl_cur_c.setStyleSheet("color: #909399; font-size: 11px; border: none;")
            
            lbl_info = QLabel("共 0 张")
            lbl_info.setStyleSheet("color: #606266; border: none;" if not self.dark_mode else "color: #A0A0A0; border: none;")
            lbl_info.setFixedWidth(60)
            lbl_info.setAlignment(Qt.AlignCenter)
            
            btn_sel = QPushButton("≡ 选择")
            btn_sel.setFixedWidth(60) 
            btn_sel.clicked.connect(lambda _, c=cat_name: self.open_manual_selection_window(c))
            
            btn_clear = QPushButton("× 清空")
            btn_clear.setFixedWidth(60)
            # 保持清空按钮的特殊红色样式
            if not self.dark_mode:
                btn_clear.setStyleSheet("QPushButton { color: #F56C6C; border-color: #fbc4c4; background-color: #fef0f0; } QPushButton:hover { background-color: #F56C6C; color: white; border-color: #F56C6C; }")
            else:
                btn_clear.setStyleSheet("QPushButton { color: #ff8080; border-color: #703030; background-color: #3a1c1c; } QPushButton:hover { background-color: #703030; color: white; }")
                
            btn_clear.clicked.connect(lambda _, c=cat_name: self.clear_single_category(c))

            btn_re = QPushButton("⟳ 重抽")
            btn_re.setFixedWidth(60) 
            btn_re.clicked.connect(lambda _, c=cat_name: self.reset_single_category(c))

            h_layout.addWidget(lbl)
            h_layout.addStretch()
            h_layout.addWidget(lbl_c)
            h_layout.addWidget(spin_c)
            h_layout.addWidget(lbl_cur_c)
            h_layout.addSpacing(10)
            h_layout.addWidget(lbl_info)
            h_layout.addSpacing(15)
            h_layout.addWidget(btn_sel)
            h_layout.addWidget(btn_clear)
            h_layout.addWidget(btn_re)
            f_layout.addLayout(h_layout)

            text_edit = QTextEdit()
            text_edit.setReadOnly(True)  
            text_edit.setCursor(QCursor(Qt.PointingHandCursor)) 
            # [需求1] 增加高度至100px，让内容展示更完整
            text_edit.setFixedHeight(100) 
            
            # [需求1] 移除横向滚动条，开启软换行，增加内边距
            text_edit.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            text_edit.setLineWrapMode(QTextEdit.WidgetWidth)
            # 给内部文字留点呼吸空间
            text_edit.setStyleSheet("padding: 5px;") 
            
            text_edit.mousePressEvent = lambda event, c=cat_name: self.open_manual_selection_window(c)
            f_layout.addWidget(text_edit)
            
            self.category_widgets[cat_name] = {'text': text_edit, 'spin_c': spin_c, 'lbl_info': lbl_info, 'lbl_cur_c': lbl_cur_c}
            self.cats_layout.addWidget(frame)
        self.update_ui_display()

    def sync_category_ui_order(self): self.refresh_category_widgets()

    def update_ui_display(self):
        total_prompts_count = 0 
        for cat, widgets in self.category_widgets.items():
            widget = widgets['text'] 
            lbl_info = widgets['lbl_info']
            lbl_cur_c = widgets['lbl_cur_c']
            widget.clear()
            data = self.action_categories.get(cat, {})
            text_content = ""
            cat_count = 0 
            distinct_categories_count = 0 

            if isinstance(data, dict):
                for sub, acts_list in data.items():
                    if isinstance(acts_list, list):
                        groups_str = []
                        for action_group in acts_list:
                            if action_group:
                                action = list(action_group.keys())[0]
                                repeat_count = action_group[action]
                                cat_count += repeat_count 
                                distinct_categories_count += 1 
                                preview_text = self.translation_map.get(action, action) if not self.use_original_text else action
                                groups_str.append(f"{preview_text}*{repeat_count}") 
                        if groups_str: text_content += f"{sub}: {' | '.join(groups_str)}\n"
                    elif isinstance(acts_list, dict):
                        items = []
                        for k, v in acts_list.items():
                             display_k = self.translation_map.get(k, k) if not self.use_original_text else k
                             items.append(f"{display_k}*{v}")
                             cat_count += v
                             distinct_categories_count += 1
                        if items: text_content += f"{sub}: {', '.join(items)}\n"
            elif isinstance(data, list):
                distinct_categories_count = len(data)
                cat_count = len(data) 
                if data:
                    display_list = [self.translation_map.get(x, x) if not self.use_original_text else x for x in data]
                    text_content = ", ".join(display_list)
            widget.setPlainText(text_content)
            lbl_info.setText(f"共 {cat_count} 张")
            lbl_cur_c.setText(f"(当前: {distinct_categories_count}类)")
            total_prompts_count += cat_count
        self.status_label.setText(f"当前总计: {total_prompts_count} 张")

    def clear_single_category(self, cat):
        if "表情" in cat: self.action_categories[cat] = []
        elif "辅助" in cat: self.action_categories[cat] = {}
        else: self.action_categories[cat] = {} 
        self.update_ui_display()

    def reset_single_category(self, cat):
        if self.excel_data is not None:
            target_c = DEFAULT_MIN_C
            if cat in self.category_widgets: target_c = self.category_widgets[cat]['spin_c'].value()
            self.process_category_data(cat, target_c_count=target_c)
            self.update_ui_display()

    def reset_all_actions(self):
        if self.excel_data is not None:
            self.combined_plan = []
            self.refresh_prompt_list()
            for cat_name in self.action_categories.keys():
                target_c = DEFAULT_MIN_C
                if cat_name in self.category_widgets: target_c = self.category_widgets[cat_name]['spin_c'].value()
                self.process_category_data(cat_name, target_c_count=target_c)
            self.update_ui_display()
        else: QMessageBox.warning(self, "警告", "请先点击'更改 Excel'加载数据")

    def generate_plan(self):
        if self.excel_data is None:
            QMessageBox.warning(self, "警告", "请先加载Excel文件")
            return
        order = [self.cat_list.item(i).text() for i in range(self.cat_list.count())]
        raw_plan = []
        for cat_name in order:
            if cat_name not in self.category_widgets: continue
            widget = self.category_widgets[cat_name]['text']
            if not widget: continue
            text_content = widget.toPlainText()
            lines = text_content.split('\n')
            for line in lines:
                line = line.strip()
                if not line: continue
                if ':' in line: parts = line.split(':', 1); content_part = parts[1].strip()
                else: content_part = line 
                if '|' in content_part: actions = content_part.split('|')
                else: actions = content_part.split(',')
                for item in actions:
                    item = item.strip()
                    if not item: continue
                    action_name = item
                    count = 1
                    if '*' in item:
                        a_parts = item.rsplit('*', 1) 
                        if len(a_parts) == 2 and a_parts[1].isdigit():
                            action_name = a_parts[0].strip(); count = int(a_parts[1])
                    original_code = action_name
                    if not self.use_original_text: 
                         for code, trans in self.translation_map.items():
                             if trans == action_name: original_code = code; break
                    raw_plan.extend([original_code] * count)
        self.combined_plan = []
        for action in raw_plan:
            self.combined_plan.append({"original_action": action, "translation": self.translation_map.get(action, "无标签"), "checked": False})
        self.refresh_prompt_list()

    def refresh_prompt_list(self):
        self.prompt_table.setRowCount(0) 
        self.prompt_table.setRowCount(len(self.combined_plan))
        for idx, item in enumerate(self.combined_plan):
            item_idx = QTableWidgetItem(f"No.{idx + 1}")
            item_idx.setFlags(item_idx.flags() ^ Qt.ItemIsEditable) 
            item_idx.setTextAlignment(Qt.AlignCenter)
            self.prompt_table.setItem(idx, 0, item_idx)
            
            chk_widget = QWidget(); chk_layout = QHBoxLayout(chk_widget); chk_layout.setContentsMargins(0,0,0,0); chk_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox(); chk.setChecked(item['checked'])
            chk.clicked.connect(lambda checked, i=idx: self.on_checkbox_toggled(checked, i))
            chk_layout.addWidget(chk)
            self.prompt_table.setCellWidget(idx, 1, chk_widget)
            
            trans_text = item['translation'] if item['translation'] else "未定义"
            item_tag = QTableWidgetItem(f"[{trans_text}]")
            item_tag.setFlags(item_tag.flags() ^ Qt.ItemIsEditable); item_tag.setTextAlignment(Qt.AlignCenter)
            self.prompt_table.setItem(idx, 2, item_tag)
            
            link_lbl = QLabel("[全选]")
            link_lbl.setStyleSheet("color: #409EFF; text-decoration: underline; margin-left: 5px;")
            link_lbl.setCursor(QCursor(Qt.PointingHandCursor))
            link_lbl.mousePressEvent = lambda e, t=trans_text: self.select_group_by_translation(t)
            self.prompt_table.setCellWidget(idx, 3, link_lbl)
            
            item_content = QTableWidgetItem(item['original_action'])
            item_content.setFlags(item_content.flags() ^ Qt.ItemIsEditable)
            self.prompt_table.setItem(idx, 4, item_content)
        self.prompt_table.resizeColumnsToContents()
        header = self.prompt_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        self.prompt_table.resizeRowsToContents()
        self.count_label.setText(f"Prompt数: {len(self.combined_plan)}")

    def on_checkbox_toggled(self, checked, index):
        if 0 <= index < len(self.combined_plan): self.combined_plan[index]['checked'] = checked

    def select_group_by_translation(self, trans_text):
        target_indices = [i for i, x in enumerate(self.combined_plan) if x['translation'] == trans_text]
        if not target_indices: return
        all_checked = all(self.combined_plan[i]['checked'] for i in target_indices)
        new_state = not all_checked
        for i in target_indices:
            self.combined_plan[i]['checked'] = new_state
            cb_widget = self.prompt_table.cellWidget(i, 1)
            if cb_widget:
                cb = cb_widget.findChild(QCheckBox)
                if cb: cb.blockSignals(True); cb.setChecked(new_state); cb.blockSignals(False)

    def move_prompt(self, direction):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: return
        moved = False
        if direction == -1: 
            for i in sorted(indices):
                if i > 0 and not self.combined_plan[i-1]["checked"]:
                    self.combined_plan[i], self.combined_plan[i-1] = self.combined_plan[i-1], self.combined_plan[i]; moved = True
        else: 
            for i in sorted(indices, reverse=True):
                if i < len(self.combined_plan) - 1 and not self.combined_plan[i+1]["checked"]:
                    self.combined_plan[i], self.combined_plan[i+1] = self.combined_plan[i+1], self.combined_plan[i]; moved = True
        if moved: self.refresh_prompt_list()

    def batch_check(self, state):
        for i in range(len(self.combined_plan)): self.combined_plan[i]['checked'] = state
        self.refresh_prompt_list()

    def delete_selected_prompt(self):
        self.combined_plan = [x for x in self.combined_plan if not x['checked']]
        self.refresh_prompt_list()

    def copy_selected_prompt(self):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: return
        offset = 0
        for i in sorted(indices):
            current_idx = i + offset
            new_item = self.combined_plan[current_idx].copy()
            new_item['checked'] = True
            self.combined_plan.insert(current_idx + 1, new_item)
            offset += 1
        self.refresh_prompt_list()

    def add_extra_prompt(self):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: QMessageBox.information(self, "提示", "请先勾选需要添加tag的行"); return
        text, ok = QInputDialog.getText(self, "添加tag", "请输入要添加的 tag (例如: masterpiece)：")
        if ok and text:
            for i in indices:
                orig = self.combined_plan[i]["original_action"]
                self.combined_plan[i]["original_action"] = f"{text}, {orig}" if not orig.startswith(",") else f"{text}{orig}"
            self.refresh_prompt_list()

    def open_manual_selection_window(self, cat_name):
        if self.excel_data is None: return
        
        options = self.get_all_options_for_category(cat_name)
        if not options: QMessageBox.information(self, "提示", "该分类下无可用选项"); return

        current_data = self.action_categories.get(cat_name, {})
        current_actions_map = {}
        selection_order = [] 

        if isinstance(current_data, dict):
            for sub, items in current_data.items():
                if isinstance(items, list):
                    for item in items:
                        if item:
                            act = list(item.keys())[0]
                            cnt = item[act]
                            current_actions_map[act] = cnt
                            if act not in selection_order: selection_order.append(act)
        
        dialog = QDialog(self)
        dialog.setWindowTitle(f"选择动作 - {cat_name} (按点击顺序生成)")
        dialog.resize(1100, 700) 
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinMaxButtonsHint)
        
        main_layout = QVBoxLayout(dialog)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        lbl_hint = QLabel(f"分类：<b>{cat_name}</b> (动作将按照你点击的顺序生成，按钮前会显示顺序编号)")
        lbl_hint.setStyleSheet("font-size: 14px; margin-bottom: 5px;")
        main_layout.addWidget(lbl_hint)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(20) 
        
        grouped = {}
        for opt in options:
            s = opt["sub"]; 
            if s not in grouped: grouped[s] = []
            grouped[s].append(opt)

        all_item_widgets = {} 

        # [需求2] 使用 modern Tag 样式按钮
        # 已经在全局 QSS 中定义了 #TagButton 样式，这里直接设置 ObjectName

        def refresh_button_labels():
            for act_key, widget_pack in all_item_widgets.items():
                btn_obj = widget_pack["btn"]
                orig_text = widget_pack.get("orig_text", "")
                if act_key in selection_order:
                    idx = selection_order.index(act_key) + 1
                    tag = get_circled_num(idx)
                    btn_obj.setText(f"{tag} {orig_text}")
                else:
                    btn_obj.setText(orig_text)

        for sub_cat, items in grouped.items():
            group_box = QGroupBox(sub_cat)
            group_box.setFont(QFont("Calibri", 10, QFont.Bold))
            grid = QGridLayout(group_box)
            grid.setSpacing(10)
            row, col = 0, 0
            COL_MAX = 4  
            
            for item in items:
                act_text = item["act"]
                display_text = item["trans"] if item["trans"] and str(item["trans"]).lower() != 'nan' else act_text
                
                cell_widget = QWidget(); cell_layout = QHBoxLayout(cell_widget); cell_layout.setContentsMargins(0,0,0,0); cell_layout.setSpacing(5)
                
                btn = QPushButton(display_text)
                btn.setObjectName("TagButton") # 应用 Tag 样式
                btn.setCheckable(True)
                btn.setCursor(QCursor(Qt.PointingHandCursor))
                btn.setToolTip(f"原始Prompt: {act_text}")
                btn.setFixedHeight(35)
                
                spin = QSpinBox(); spin.setRange(1, 99); spin.setFixedWidth(45); spin.setFixedHeight(35); spin.setAlignment(Qt.AlignCenter)
                normal_font = QFont("Calibri", 9); normal_font.setBold(False); spin.setFont(normal_font) 

                if act_text in current_actions_map:
                    btn.setChecked(True)
                    spin.setValue(current_actions_map[act_text])
                    spin.setEnabled(True)
                else:
                    btn.setChecked(False)
                    spin.setValue(3) 
                    spin.setEnabled(False)

                all_item_widgets[act_text] = {
                    "btn": btn,
                    "spin": spin,
                    "sub": sub_cat,
                    "orig_text": display_text 
                }

                def toggle_spin(checked, s=spin, a=act_text):
                    s.setEnabled(checked)
                    if checked:
                        if a not in selection_order: selection_order.append(a)
                    else:
                        if a in selection_order: selection_order.remove(a)
                    refresh_button_labels() 
                        
                btn.toggled.connect(toggle_spin)
                cell_layout.addWidget(btn); cell_layout.addWidget(spin)
                grid.addWidget(cell_widget, row, col)
                col += 1
                if col >= COL_MAX: col = 0; row += 1
            content_layout.addWidget(group_box)

        refresh_button_labels()

        content_layout.addStretch()
        scroll.setWidget(content_widget)
        main_layout.addWidget(scroll)

        btn_confirm = QPushButton("确认修改")
        btn_confirm.setFixedHeight(40)
        btn_confirm.setStyleSheet("font-weight: bold; font-size: 14px; background-color: #409EFF; color: white;")
        main_layout.addWidget(btn_confirm)

        def apply():
            new_data = {}
            if not selection_order:
                if QMessageBox.question(dialog, "确认", "未选择任何动作，这将清空该类目，是否继续？") == QMessageBox.Yes:
                    self.action_categories[cat_name] = {}
                else: return
            else:
                for act in selection_order:
                    if act in all_item_widgets:
                        s = all_item_widgets[act]["sub"]
                        if s not in new_data: new_data[s] = []
                for act in selection_order:
                    if act in all_item_widgets:
                        widget_data = all_item_widgets[act]
                        cnt = widget_data["spin"].value()
                        sub = widget_data["sub"]
                        new_data[sub].append({act: cnt})
                self.action_categories[cat_name] = new_data
            self.update_ui_display()
            dialog.accept()

        btn_confirm.clicked.connect(apply)
        dialog.exec_()

    def get_all_options_for_category(self, cat_name):
        col_indices, _ = self.parse_config_setting(cat_name)
        options = [] 
        for col_idx in col_indices:
            if col_idx >= self.excel_data.shape[1]: continue
            col_data = self.excel_data.iloc[:, col_idx]
            trans_data = None
            if col_idx + 1 < self.excel_data.shape[1]: trans_data = self.excel_data.iloc[:, col_idx + 1]
            if len(col_data) > 0: sub_cat = str(col_data.iloc[0]).strip()
            else: continue
            for i in range(1, len(col_data)):
                act_val = col_data.iloc[i]
                if pd.notna(act_val):
                    act = str(act_val).strip()
                    if act:
                        t_str = ""
                        if trans_data is not None and i < len(trans_data):
                            t_val = trans_data.iloc[i]
                            if pd.notna(t_val): t_str = str(t_val).strip()
                        options.append({"sub": sub_cat, "act": act, "trans": t_str})
        return options

    # [需求2] 重构辅助动作选择窗口：美观的网格 Tag 布局
    def open_add_action_window(self):
        if not self.combined_plan: return
        dialog = QDialog(self)
        dialog.setWindowTitle("添加辅助动作与表情")
        
        screen_geo = QApplication.desktop().screenGeometry()
        dialog.resize(int(screen_geo.width() * 0.8), int(screen_geo.height() * 0.8))
        
        layout = QVBoxLayout(dialog)
        
        main_scroll = QScrollArea()
        main_scroll.setWidgetResizable(True)
        # [需求1] 确保无横向滚动条
        main_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff) 
        
        container = QWidget()
        grid = QGridLayout(container)
        grid.setContentsMargins(20, 20, 20, 20)
        grid.setSpacing(15)

        cols_data = []
        aux_data = self.action_categories.get("辅助动作（S/U/W列）", {})
        for sub, acts in aux_data.items(): cols_data.append((sub, list(acts.keys())))
        
        emo_acts = self.action_categories.get("表情（Y列）", [])
        
        all_checkboxes = []
        current_idx = 0
        COL_COUNT = 4 

        # --- 辅助动作 ---
        for title, items in cols_data:
            lbl = QLabel(f"【{title}】")
            # 标题样式优化
            lbl.setStyleSheet(
                "font-weight: bold; font-size: 15px; margin-top: 20px; margin-bottom: 10px; " + 
                ("color: #333;" if not self.dark_mode else "color: #E0E0E0;")
            )
            
            current_row = current_idx // COL_COUNT
            if current_idx % COL_COUNT != 0: 
                current_row += 1
                current_idx = current_row * COL_COUNT 
            
            grid.addWidget(lbl, current_row, 0, 1, COL_COUNT)
            current_idx += COL_COUNT 

            for item in items:
                trans = self.translation_map.get(item, "")
                cb = TruncatedCheckBox(item, trans)
                
                r = current_idx // COL_COUNT
                c = current_idx % COL_COUNT
                grid.addWidget(cb, r, c)
                all_checkboxes.append(cb)
                current_idx += 1

        # --- 表情 ---
        if emo_acts:
            current_row = current_idx // COL_COUNT
            if current_idx % COL_COUNT != 0: 
                current_row += 1
                current_idx = current_row * COL_COUNT 

            lbl_emo = QLabel("【表情】")
            lbl_emo.setStyleSheet(
                "font-weight: bold; font-size: 15px; margin-top: 20px; margin-bottom: 10px; " + 
                ("color: #333;" if not self.dark_mode else "color: #E0E0E0;")
            )
            grid.addWidget(lbl_emo, current_row, 0, 1, COL_COUNT)
            current_idx += COL_COUNT

            for item in emo_acts:
                trans = self.translation_map.get(item, "")
                cb = TruncatedCheckBox(item, trans)
                
                r = current_idx // COL_COUNT
                c = current_idx % COL_COUNT
                grid.addWidget(cb, r, c)
                all_checkboxes.append(cb)
                current_idx += 1

        grid.setRowStretch(grid.rowCount(), 1)
        
        main_scroll.setWidget(container)
        layout.addWidget(main_scroll)

        btn = QPushButton("确认添加")
        btn.setFixedHeight(45)
        btn.setStyleSheet("font-weight: bold; font-size: 14px; background-color: #409EFF; color: white; border-radius: 6px;")
        def apply():
            tags = [cb.property("code") for cb in all_checkboxes if cb.isChecked()]
            if tags:
                indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
                for i in indices:
                    orig = self.combined_plan[i]["original_action"]
                    self.combined_plan[i]["original_action"] = f"{','.join(tags)}, {orig}"
                self.refresh_prompt_list()
            dialog.accept()
        btn.clicked.connect(apply)
        layout.addWidget(btn)
        
        dialog.exec_()

    # [需求3] 智能导出路径
    def export_excel(self):
        if not self.combined_plan: return
        
        last_dir = self.config_data.get("last_export_dir", "")
        if not last_dir or not os.path.exists(last_dir):
            last_dir = self.base_dir
            
        default_path = os.path.join(last_dir, "actions.xlsx")
        
        path, _ = QFileDialog.getSaveFileName(self, "导出 Excel", default_path, "Excel Files (*.xlsx)")
        
        if path:
            new_dir = os.path.dirname(path)
            if new_dir != last_dir:
                self.config_data["last_export_dir"] = new_dir
                self.save_config()
                
            try:
                final = [x["original_action"] for x in self.combined_plan]
                seeds = [random.randint(10000000000000, 99999999999999) for _ in final]
                df = pd.DataFrame({
                    "序号": range(1, len(final)+1),
                    "动作Prompt": final,
                    "种子": seeds,
                    "完成情况": [""] * len(final)
                })
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    ws = writer.sheets['Sheet1']
                    ws.column_dimensions['C'].width = 20
                    for row in range(2, len(final) + 2): ws.cell(row=row, column=3).number_format = '0'
                QMessageBox.information(self, "成功", "导出完成")
            except Exception as e: QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    set_light_theme(app)
    window = MainWindow(app)
    window.show()
    sys.exit(app.exec_())