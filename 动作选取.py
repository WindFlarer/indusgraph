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
    QTableWidgetItem, QHeaderView, QSpinBox, QGroupBox
)
from PyQt5.QtGui import QPalette, QColor, QFont, QCursor
from PyQt5.QtCore import Qt, pyqtSignal, QSize

# --- 默认常量配置 ---
DEFAULT_COL_MAPPING = {
  "第一类动作（A/C列）": [0, 2],
  "第二类动作（E/G列）": [4, 6],
  "第三类动作（I列）": {
    "cols": [8], 
    "min_c": 4, 
  },
  "第四类动作（K列）": [10],
  "第五类动作（M/O列）": [12, 14],
  "辅助动作（S/U/W列）": [18, 20, 22],
  "表情（Y列）": [24]
}

# --- 默认参数设置 ---
DEFAULT_MIN_C = 3          # 默认保底类别数 (min_c)
REPEAT_MIN = 3             # 单个动作最少重复次数
REPEAT_MAX = 5             # 单个动作最多重复次数

# --- 主题设置 ---
def set_light_theme(app: QApplication):
    app.setStyle("Fusion")
    palette = app.style().standardPalette()
    palette.setColor(QPalette.Highlight, QColor(225, 235, 250))
    palette.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(palette)

def set_dark_theme(app: QApplication):
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.WindowText, Qt.white)
    palette.setColor(QPalette.Base, QColor(35, 35, 35))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipBase, Qt.white)
    palette.setColor(QPalette.ToolTipText, Qt.white)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Highlight, QColor(60, 75, 90))
    palette.setColor(QPalette.HighlightedText, Qt.white)
    app.setPalette(palette)

class MainWindow(QMainWindow):
    def __init__(self, app: QApplication):
        super().__init__()
        self.app = app
        self.dark_mode = False
        
        self.excel_data = None
        self.current_excel_path = None
        self.translation_map = {}
        self.combined_plan = [] 
        
        # 加载配置文件
        self.col_mapping = self.load_config()
        
        self.action_categories = {}
        for k in self.col_mapping.keys():
            if "表情" in k: self.action_categories[k] = []
            elif "辅助" in k: self.action_categories[k] = {}
            else: self.action_categories[k] = {}
            
        self.category_widgets = {} 
        self.use_original_text = False

        self.init_ui()
        self.load_excel() 

    def load_config(self):
        """加载 config.json 配置文件，如果不存在则写入默认配置"""
        config_path = "config.json"
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    print(f"已加载配置文件: {config_path}")
                    return config
            except Exception as e:
                print(f"配置文件加载失败，使用默认设置: {e}")
                return DEFAULT_COL_MAPPING.copy()
        else:
            try:
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(DEFAULT_COL_MAPPING, f, indent=4, ensure_ascii=False)
            except Exception: pass
            return DEFAULT_COL_MAPPING.copy()

    def parse_config_setting(self, cat_name):
        """
        解析配置项，移除 min_p，保留 min_c
        返回: (col_indices_list, min_c)
        """
        setting = self.col_mapping.get(cat_name)
        
        final_cols = []
        final_min_c = DEFAULT_MIN_C

        if setting is None:
            return final_cols, final_min_c

        if isinstance(setting, dict):
            final_cols = setting.get("cols", [])
            # 移除 min_p 读取
            final_min_c = setting.get("min_c", DEFAULT_MIN_C)
        elif isinstance(setting, list):
            final_cols = setting
            final_min_c = DEFAULT_MIN_C
        
        return final_cols, final_min_c

    def init_ui(self):
        self.setWindowTitle("动作训练计划生成器 (按类别保底版)")
        # [修改] 增加窗口初始宽度，确保能容纳两边内容
        self.resize(1500, 900)
        
        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 顶部栏
        top_bar = QHBoxLayout()
        self.theme_btn = QPushButton("切换深色 / 浅色模式")
        self.theme_btn.clicked.connect(self.toggle_theme)
        
        self.lang_toggle = QCheckBox("左侧显示原始Prompt")
        self.lang_toggle.stateChanged.connect(self.toggle_language_display)
        
        top_bar.addStretch()
        top_bar.addWidget(self.lang_toggle)
        top_bar.addWidget(self.theme_btn)
        main_layout.addLayout(top_bar)

        # 分割窗格
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # --- 左侧面板 ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        # 文件操作区
        file_layout = QHBoxLayout()
        self.file_label = QLabel("等待加载...")
        btn_open = QPushButton("打开 Excel")
        btn_open.clicked.connect(self.load_excel)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(btn_open)
        left_layout.addLayout(file_layout)

        # 控制区
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
        btn_export = QPushButton("导出 Excel")
        btn_export.clicked.connect(self.export_excel)
        
        btn_box.addWidget(btn_reset)
        btn_box.addWidget(btn_gen)
        btn_box.addWidget(btn_export)
        
        ctrl_layout.addWidget(self.cat_list, 1)
        ctrl_layout.addLayout(btn_box, 1)
        left_layout.addLayout(ctrl_layout)

        # 分类详情滚动区
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.cats_container_widget = QWidget()
        self.cats_layout = QVBoxLayout(self.cats_container_widget)
        self.cats_layout.setAlignment(Qt.AlignTop)
        
        scroll.setWidget(self.cats_container_widget)
        left_layout.addWidget(scroll)
        
        # 左下角状态标签
        self.status_label = QLabel("当前总计: 0 张")
        self.status_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #4282da; padding: 5px;")
        left_layout.addWidget(self.status_label)

        self.refresh_category_widgets()

        # --- 右侧面板 ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        # 工具栏
        tool_bar = QHBoxLayout()
        btn_all = QPushButton("全选")
        btn_all.clicked.connect(self.smart_select_all) 
        btn_inv = QPushButton("反选")
        btn_inv.clicked.connect(self.smart_invert_selection) 
        
        btn_up = QPushButton("▲")
        btn_up.clicked.connect(lambda: self.move_prompt(-1))
        btn_down = QPushButton("▼")
        btn_down.clicked.connect(lambda: self.move_prompt(1))
        
        btn_copy = QPushButton("复制")
        btn_copy.clicked.connect(self.copy_selected_prompt)
        btn_del = QPushButton("删除")
        btn_del.clicked.connect(self.delete_selected_prompt)
        
        btn_tag = QPushButton("添加 Tag")
        btn_tag.clicked.connect(self.add_extra_prompt)
        btn_aux = QPushButton("添加辅助/表情")
        btn_aux.clicked.connect(self.open_add_action_window)

        tool_bar.addWidget(btn_all)
        tool_bar.addWidget(btn_inv)
        tool_bar.addSpacing(10)
        tool_bar.addWidget(btn_up)
        tool_bar.addWidget(btn_down)
        tool_bar.addSpacing(10)
        tool_bar.addWidget(btn_copy)
        tool_bar.addWidget(btn_del)
        tool_bar.addSpacing(10)
        tool_bar.addWidget(btn_tag)
        tool_bar.addWidget(btn_aux)
        tool_bar.addStretch()
        
        self.count_label = QLabel("Prompt数: 0")
        tool_bar.addWidget(self.count_label)
        
        right_layout.addLayout(tool_bar)

        # --- 右侧面板表格配置 ---
        self.prompt_table = QTableWidget()
        self.prompt_table.setColumnCount(5)
        self.prompt_table.setHorizontalHeaderLabels(["序号", "选择", "标签", "操作", "内容"])
        
        self.prompt_table.verticalHeader().setMinimumSectionSize(60)
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
        
        right_layout.addWidget(self.prompt_table)

        # [修改] 增加左侧最小宽度，防止被挤压
        left_panel.setMinimumWidth(600) 
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        # [修改] 调整初始比例，给左侧更多空间 (850 vs 650)
        splitter.setSizes([850, 650])
        splitter.setStretchFactor(0, 0) # 左侧尽量保持尺寸
        splitter.setStretchFactor(1, 1) # 窗口拉大时优先拉大右侧

    # --- 逻辑功能 ---

    def smart_select_all(self):
        selected_rows = set(index.row() for index in self.prompt_table.selectedIndexes())
        if selected_rows:
            for row in selected_rows:
                self.combined_plan[row]['checked'] = True
                cb_widget = self.prompt_table.cellWidget(row, 1)
                if cb_widget:
                    cb = cb_widget.findChild(QCheckBox)
                    if cb: 
                        cb.blockSignals(True)
                        cb.setChecked(True)
                        cb.blockSignals(False)
        else:
            self.batch_check(True)

    def smart_invert_selection(self):
        selected_rows = set(index.row() for index in self.prompt_table.selectedIndexes())
        if selected_rows:
            for row in selected_rows:
                self.combined_plan[row]['checked'] = False
                cb_widget = self.prompt_table.cellWidget(row, 1)
                if cb_widget:
                    cb = cb_widget.findChild(QCheckBox)
                    if cb:
                        cb.blockSignals(True)
                        cb.setChecked(False)
                        cb.blockSignals(False)
        else:
            self.invert_check()

    def toggle_theme(self):
        if self.dark_mode:
            set_light_theme(self.app)
        else:
            set_dark_theme(self.app)
        self.dark_mode = not self.dark_mode
        self.update_ui_display()
        self.refresh_prompt_list()

    def toggle_language_display(self, state):
        self.use_original_text = (state == Qt.Checked)
        self.update_ui_display()
        self.refresh_prompt_list()

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx)")
        if path:
            try:
                self.current_excel_path = path
                self.excel_data = pd.read_excel(path, header=None)
                self.translation_map = {}
                
                # --- 加载时同时应用 min_c ---
                for cat in self.col_mapping.keys():
                    _, min_c = self.parse_config_setting(cat)
                    self.process_category_data(cat, target_c_count=min_c)
                    
                self.update_ui_display()
                self.file_label.setText(os.path.basename(path))
            except Exception as e:
                if self.sender():
                    QMessageBox.critical(self, "错误", f"无法读取文件: {str(e)}")
                else:
                    print(f"初始化加载取消或失败: {e}")

    def process_category_data(self, cat_name, target_c_count=DEFAULT_MIN_C):
        # 移除了 min_p
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
            if col_idx + 1 < self.excel_data.shape[1]:
                trans_data = self.excel_data.iloc[:, col_idx + 1]
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
            # --- 智能抽取逻辑 ---
            final_selection = []
            shuffled_pool = all_actions_pool.copy()
            random.shuffle(shuffled_pool)
            
            # 只抽取 min_c 个不同的动作
            needed_unique = min(target_c_count, len(shuffled_pool))
            
            # 使用字典去重动作名，确保抽取的是不同的动作
            unique_actions = {} # key: action, val: sub
            for a, s in shuffled_pool:
                if a not in unique_actions:
                    unique_actions[a] = s
            
            unique_keys = list(unique_actions.keys())
            # 如果去重后不够 min_c，就取全部；否则取前 min_c 个
            count_to_take = min(target_c_count, len(unique_keys))
            selected_keys = unique_keys[:count_to_take]

            for act in selected_keys:
                sub = unique_actions[act]
                repeat = random.randint(REPEAT_MIN, REPEAT_MAX)
                final_selection.append({'action': act, 'count': repeat, 'sub': sub})
                
            # 整理数据结构
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

            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            f_layout = QVBoxLayout(frame)
            
            h_layout = QHBoxLayout()
            lbl = QLabel(cat_name)
            lbl.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
            lbl.setMinimumWidth(150)
            
            lbl_c = QLabel("保底类别:")
            spin_c = QSpinBox()
            spin_c.setRange(1, 50) 
            spin_c.setValue(min_c_val) 
            spin_c.setFixedWidth(50) 
            spin_c.setToolTip(f"最少抽取 {min_c_val} 种不同的动作")
            
            # --- 新增：显示当前实际类别的 Label ---
            lbl_cur_c = QLabel("(当前: 0类)")
            lbl_cur_c.setStyleSheet("color: #555; font-size: 11px;")
            
            # ---
            
            lbl_info = QLabel("共 0 张")
            lbl_info.setStyleSheet("color: #666;")
            lbl_info.setFixedWidth(60)
            lbl_info.setAlignment(Qt.AlignCenter)
            
            btn_sel = QPushButton("≡ 选择")
            btn_sel.setFixedWidth(50) 
            btn_sel.clicked.connect(lambda _, c=cat_name: self.open_manual_selection_window(c))
            
            btn_re = QPushButton("⟳ 重抽")
            btn_re.setFixedWidth(50) 
            btn_re.clicked.connect(lambda _, c=cat_name: self.reset_single_category(c))

            h_layout.addWidget(lbl)
            h_layout.addStretch()
            
            h_layout.addWidget(lbl_c)
            h_layout.addWidget(spin_c)
            h_layout.addWidget(lbl_cur_c) # 添加当前类别显示
            h_layout.addSpacing(10)
            h_layout.addWidget(lbl_info)
            h_layout.addSpacing(15)
            h_layout.addWidget(btn_sel)
            h_layout.addWidget(btn_re)
            f_layout.addLayout(h_layout)

            text_edit = QTextEdit()
            text_edit.setReadOnly(True)  
            text_edit.setCursor(QCursor(Qt.PointingHandCursor)) 
            text_edit.setFixedHeight(80) 
            
            text_edit.mousePressEvent = lambda event, c=cat_name: self.open_manual_selection_window(c)
            
            f_layout.addWidget(text_edit)
            
            self.category_widgets[cat_name] = {
                'text': text_edit, 
                'spin_c': spin_c,      
                'lbl_info': lbl_info,
                'lbl_cur_c': lbl_cur_c # 存入引用
            }
            self.cats_layout.addWidget(frame)
        self.update_ui_display()

    def sync_category_ui_order(self):
        self.refresh_category_widgets()

    def update_ui_display(self):
        total_prompts_count = 0 

        for cat, widgets in self.category_widgets.items():
            widget = widgets['text'] 
            lbl_info = widgets['lbl_info']
            lbl_cur_c = widgets['lbl_cur_c'] # 获取引用
            
            widget.clear()
            data = self.action_categories.get(cat, {})
            
            text_content = ""
            cat_count = 0 
            distinct_categories_count = 0 # 统计当前有多少个不同的动作/子类

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
                        if groups_str:
                             text_content += f"{sub}: {' | '.join(groups_str)}\n"
                    elif isinstance(acts_list, dict):
                        # 辅助动作
                        items = []
                        for k, v in acts_list.items():
                             display_k = self.translation_map.get(k, k) if not self.use_original_text else k
                             items.append(f"{display_k}*{v}")
                             cat_count += v
                             distinct_categories_count += 1
                        if items:
                             text_content += f"{sub}: {', '.join(items)}\n"
            
            elif isinstance(data, list):
                # 表情
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

    def reset_single_category(self, cat):
        if self.excel_data is not None:
            target_c = DEFAULT_MIN_C
            if cat in self.category_widgets:
                target_c = self.category_widgets[cat]['spin_c'].value()
            self.process_category_data(cat, target_c_count=target_c)
            self.update_ui_display()

    def reset_all_actions(self):
        if self.excel_data is not None:
            self.combined_plan = []
            self.refresh_prompt_list()
            for cat_name in self.action_categories.keys():
                target_c = DEFAULT_MIN_C
                if cat_name in self.category_widgets:
                    target_c = self.category_widgets[cat_name]['spin_c'].value()
                self.process_category_data(cat_name, target_c_count=target_c)
            self.update_ui_display()
        else:
            QMessageBox.warning(self, "警告", "请先点击'打开 Excel'加载数据")

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
                
                if ':' in line:
                    parts = line.split(':', 1)
                    content_part = parts[1].strip()
                else:
                    content_part = line 
                
                if '|' in content_part:
                    actions = content_part.split('|')
                else:
                    actions = content_part.split(',')
                
                for item in actions:
                    item = item.strip()
                    if not item: continue
                    
                    action_name = item
                    count = 1
                    
                    if '*' in item:
                        a_parts = item.rsplit('*', 1) 
                        if len(a_parts) == 2 and a_parts[1].isdigit():
                            action_name = a_parts[0].strip()
                            count = int(a_parts[1])
                    
                    original_code = action_name
                    if not self.use_original_text: 
                         for code, trans in self.translation_map.items():
                             if trans == action_name:
                                 original_code = code
                                 break
                    
                    raw_plan.extend([original_code] * count)

        self.combined_plan = []
        for action in raw_plan:
            self.combined_plan.append({
                "original_action": action,
                "translation": self.translation_map.get(action, "无标签"),
                "checked": False
            })
        self.refresh_prompt_list()

    def refresh_prompt_list(self):
        self.prompt_table.setRowCount(0) 
        self.prompt_table.setRowCount(len(self.combined_plan))
        
        for idx, item in enumerate(self.combined_plan):
            item_idx = QTableWidgetItem(f"No.{idx + 1}")
            item_idx.setFlags(item_idx.flags() ^ Qt.ItemIsEditable) 
            item_idx.setTextAlignment(Qt.AlignCenter)
            self.prompt_table.setItem(idx, 0, item_idx)
            
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.setContentsMargins(0,0,0,0)
            chk_layout.setAlignment(Qt.AlignCenter)
            chk = QCheckBox()
            chk.setChecked(item['checked'])
            chk.clicked.connect(lambda checked, i=idx: self.on_checkbox_toggled(checked, i))
            chk_layout.addWidget(chk)
            self.prompt_table.setCellWidget(idx, 1, chk_widget)
            
            trans_text = item['translation'] if item['translation'] else "未定义"
            item_tag = QTableWidgetItem(f"[{trans_text}]")
            item_tag.setFlags(item_tag.flags() ^ Qt.ItemIsEditable)
            item_tag.setTextAlignment(Qt.AlignCenter)
            self.prompt_table.setItem(idx, 2, item_tag)
            
            link_lbl = QLabel("[全选]")
            link_lbl.setStyleSheet("color: #4282da; text-decoration: underline; margin-left: 5px;")
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
        if 0 <= index < len(self.combined_plan):
            self.combined_plan[index]['checked'] = checked

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
                if cb:
                    cb.blockSignals(True)
                    cb.setChecked(new_state)
                    cb.blockSignals(False)

    def move_prompt(self, direction):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: return
        moved = False
        if direction == -1: 
            for i in sorted(indices):
                if i > 0 and not self.combined_plan[i-1]["checked"]:
                    self.combined_plan[i], self.combined_plan[i-1] = self.combined_plan[i-1], self.combined_plan[i]
                    moved = True
        else: 
            for i in sorted(indices, reverse=True):
                if i < len(self.combined_plan) - 1 and not self.combined_plan[i+1]["checked"]:
                    self.combined_plan[i], self.combined_plan[i+1] = self.combined_plan[i+1], self.combined_plan[i]
                    moved = True
        if moved: self.refresh_prompt_list()

    def batch_check(self, state):
        for i in range(len(self.combined_plan)):
            self.combined_plan[i]['checked'] = state
        self.refresh_prompt_list()

    def invert_check(self):
        for i in range(len(self.combined_plan)):
            self.combined_plan[i]['checked'] = not self.combined_plan[i]['checked']
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
        if not indices:
            QMessageBox.information(self, "提示", "请先勾选需要添加tag的行")
            return
        text, ok = QInputDialog.getText(self, "添加tag", "请输入要添加的 tag (例如: masterpiece)：")
        if ok and text:
            for i in indices:
                orig = self.combined_plan[i]["original_action"]
                self.combined_plan[i]["original_action"] = f"{text}, {orig}" if not orig.startswith(",") else f"{text}{orig}"
            self.refresh_prompt_list()

    def open_manual_selection_window(self, cat_name):
        if self.excel_data is None: return
        
        # 获取该大类下的所有选项
        options = self.get_all_options_for_category(cat_name)
        if not options:
            QMessageBox.information(self, "提示", "该分类下无可用选项")
            return

        # --- 准备数据：获取当前分类下已生成的动作和张数 ---
        current_data = self.action_categories.get(cat_name, {})
        # 将结构化的 current_data 扁平化为 {action_name: count}，方便查找
        current_actions_map = {}
        if isinstance(current_data, dict):
            for sub, items in current_data.items():
                if isinstance(items, list):
                    for item in items:
                        if item:
                            act = list(item.keys())[0]
                            cnt = item[act]
                            current_actions_map[act] = cnt
        # ---------------------------------------------------

        # 1. 弹窗基础设置
        dialog = QDialog(self)
        dialog.setWindowTitle(f"选择动作 - {cat_name}")
        dialog.resize(1100, 700) 
        dialog.setWindowFlags(dialog.windowFlags() | Qt.WindowMinMaxButtonsHint)
        
        main_layout = QVBoxLayout(dialog)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 2. 顶部提示
        lbl_hint = QLabel(f"分类：<b>{cat_name}</b> (点击选择动作，右侧设置张数)")
        lbl_hint.setStyleSheet("font-size: 14px; margin-bottom: 5px;")
        main_layout.addWidget(lbl_hint)

        # 3. 内容滚动区
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(20) 
        
        # 4. 生成选项按钮
        grouped = {}
        for opt in options:
            s = opt["sub"]
            if s not in grouped: grouped[s] = []
            grouped[s].append(opt)

        all_item_widgets = [] # 存储每个选项的控件引用 (Button, SpinBox, ItemData)

        # 定义按钮样式
        btn_css_light = """
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 6px;
                color: #888;
                text-align: left;
            }
            QPushButton:checked {
                background-color: #0e639c;
                color: white;
                border: 1px solid #0a4b75;
                font-weight: bold;
            }
        """
        btn_css_dark = """
            QPushButton {
                background-color: #333;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 6px;
                color: #888;
                text-align: left;
            }
            QPushButton:checked {
                background-color: #0e639c;
                color: white;
                border: 1px solid #4282da;
                font-weight: bold;
            }
        """
        current_css = btn_css_dark if self.dark_mode else btn_css_light

        for sub_cat, items in grouped.items():
            group_box = QGroupBox(sub_cat)
            group_box.setFont(QFont("Microsoft YaHei", 10, QFont.Bold))
            
            grid = QGridLayout(group_box)
            grid.setSpacing(10)
            
            row, col = 0, 0
            COL_MAX = 4  # 每行4个
            
            for item in items:
                act_text = item["act"]
                display_text = item["trans"] if item["trans"] and str(item["trans"]).lower() != 'nan' else act_text
                
                # --- 每个单元格的容器 ---
                cell_widget = QWidget()
                cell_layout = QHBoxLayout(cell_widget)
                cell_layout.setContentsMargins(0,0,0,0)
                cell_layout.setSpacing(5)
                
                # 按钮
                btn = QPushButton(display_text)
                btn.setCheckable(True)
                btn.setCursor(QCursor(Qt.PointingHandCursor))
                btn.setStyleSheet(current_css)
                btn.setToolTip(f"原始Prompt: {act_text}")
                btn.setFixedHeight(35)
                
                # 数字框
                spin = QSpinBox()
                spin.setRange(1, 99)
                spin.setFixedWidth(45)
                spin.setFixedHeight(35)
                spin.setAlignment(Qt.AlignCenter)
                
                # [FIX] 显式设置字体为正常粗细，防止继承 GroupBox 的粗体
                normal_font = QFont("Microsoft YaHei", 9)
                normal_font.setBold(False)
                spin.setFont(normal_font) 

                # --- 状态初始化 ---
                if act_text in current_actions_map:
                    # 如果已存在：选中，填入现有数量
                    btn.setChecked(True)
                    spin.setValue(current_actions_map[act_text])
                    spin.setEnabled(True)
                    spin.setStyleSheet("") # 正常颜色
                else:
                    # 如果不存在：未选中，灰色，默认3
                    btn.setChecked(False)
                    spin.setValue(3) 
                    spin.setEnabled(False)
                    spin.setStyleSheet("color: #aaa;") # 灰色文字

                # --- 交互逻辑：点击按钮切换数字框状态 ---
                def toggle_spin(checked, s=spin):
                    s.setEnabled(checked)
                    if checked:
                        s.setStyleSheet("")
                    else:
                        s.setStyleSheet("color: #aaa;")
                        
                btn.toggled.connect(toggle_spin)
                
                cell_layout.addWidget(btn)
                cell_layout.addWidget(spin)
                
                grid.addWidget(cell_widget, row, col)
                
                # 记录引用以便 Apply 时读取
                all_item_widgets.append({
                    "btn": btn,
                    "spin": spin,
                    "data": item
                })
                
                col += 1
                if col >= COL_MAX:
                    col = 0
                    row += 1
            
            content_layout.addWidget(group_box)

        content_layout.addStretch()
        scroll.setWidget(content_widget)
        main_layout.addWidget(scroll)

        # 5. 底部按钮
        btn_confirm = QPushButton("确认修改")
        btn_confirm.setFixedHeight(40)
        btn_confirm.setStyleSheet("font-weight: bold; font-size: 14px;")
        main_layout.addWidget(btn_confirm)

        def apply():
            new_data = {}
            has_selection = False
            
            for widget_pack in all_item_widgets:
                btn = widget_pack["btn"]
                spin = widget_pack["spin"]
                item = widget_pack["data"]
                
                if btn.isChecked():
                    has_selection = True
                    sub = item["sub"]
                    act = item["act"]
                    count = spin.value()
                    
                    if sub not in new_data: new_data[sub] = []
                    new_data[sub].append({act: count})
            
            if not has_selection:
                if QMessageBox.question(dialog, "确认", "未选择任何动作，这将清空该类目，是否继续？") == QMessageBox.Yes:
                    self.action_categories[cat_name] = {}
                else:
                    return
            else:
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

    def open_add_action_window(self):
        if not self.combined_plan: return
        dialog = QDialog(self)
        dialog.setWindowTitle("添加辅助动作与表情")
        dialog.resize(1000, 800)
        layout = QVBoxLayout(dialog)
        scroll = QScrollArea()
        container = QWidget()
        h_layout = QHBoxLayout(container)
        cols_data = []
        aux_data = self.action_categories.get("辅助动作（S/U/W列）", {})
        for sub, acts in aux_data.items(): cols_data.append((sub, list(acts.keys())))
        emo_acts = self.action_categories.get("表情（Y列）", [])
        if emo_acts: cols_data.append(("表情", emo_acts))
        check_boxes = []
        for title, items in cols_data:
            grp = QFrame(); v = QVBoxLayout(grp)
            v.addWidget(QLabel(title))
            for item in items:
                trans = self.translation_map.get(item, "")
                display = f"{item}\n({trans})" if trans else item
                cb = QCheckBox(display); cb.setProperty("code", item)
                v.addWidget(cb); check_boxes.append(cb)
            v.addStretch(); h_layout.addWidget(grp)
        scroll.setWidget(container); layout.addWidget(scroll)
        btn = QPushButton("确认添加")
        def apply():
            tags = [cb.property("code") for cb in check_boxes if cb.isChecked()]
            if tags:
                indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
                for i in indices:
                    orig = self.combined_plan[i]["original_action"]
                    self.combined_plan[i]["original_action"] = f"{','.join(tags)}, {orig}"
                self.refresh_prompt_list()
            dialog.accept()
        btn.clicked.connect(apply); layout.addWidget(btn); dialog.exec_()

    def export_excel(self):
        if not self.combined_plan: return
        path, _ = QFileDialog.getSaveFileName(self, "导出 Excel", "actions.xlsx", "Excel Files (*.xlsx)")
        if path:
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