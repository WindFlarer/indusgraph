import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import random
import sys
import traceback
import platform
import re

# --- 常量配置 ---
INITIAL_COL_MAPPING = {
    "第一类动作（A/C列）": [0, 2],
    "第二类动作（E/G列）": [4, 6],
    "第三类动作（I列）": [8],
    "第四类动作（K列）": [10],
    "第五类动作（M/O列）": [12, 14],
    "第六类动作（Q列）": [16],
    "辅助动作（S/U/W列）": [18, 20, 22],
    "表情（Y列）": [24]
}

# --- 滚轮控制辅助类 ---
class ScrollHelper:
    @staticmethod
    def bind_mouse_scroll(widget, parent_canvas):
        def _on_mousewheel(event):
            if parent_canvas.winfo_height() < parent_canvas.bbox("all")[3]:
                if platform.system() == 'Windows':
                    parent_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif platform.system() == 'Darwin':
                    parent_canvas.yview_scroll(int(-1 * event.delta), "units")
                else:
                    if event.num == 4: parent_canvas.yview_scroll(-1, "units")
                    elif event.num == 5: parent_canvas.yview_scroll(1, "units")

        def _bind_recursive(w):
            if platform.system() == 'Linux':
                w.bind("<Button-4>", _on_mousewheel, add="+")
                w.bind("<Button-5>", _on_mousewheel, add="+")
            else:
                w.bind("<MouseWheel>", _on_mousewheel, add="+")
            for child in w.winfo_children():
                _bind_recursive(child)

        def _on_enter(event):
            _bind_recursive(widget)

        widget.bind("<Enter>", _on_enter)

# --- 多行输入对话框 ---
class BigTextInputDialog(tk.simpledialog.Dialog):
    def __init__(self, parent, title, prompt):
        self.prompt_text = prompt
        self.result_text = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text=self.prompt_text).pack(pady=5)
        self.text_box = tk.Text(master, height=10, width=50, font=("Microsoft YaHei", 10))
        self.text_box.pack(padx=10, pady=5)
        return self.text_box

    def apply(self):
        self.result_text = self.text_box.get("1.0", tk.END).strip()

class StyleConfig:
    def __init__(self, root):
        self.root = root
        self.is_dark_mode = tk.BooleanVar(value=False)
        self.setup_styles()

    def setup_styles(self):
        self.style = ttk.Style(self.root)
        try: self.style.theme_use('clam')
        except: pass

        self.light_style = {
            "bg": "#ffffff", "fg": "#000000", 
            "frame_bg": "#f3f3f3",
            "text_bg": "#ffffff", "text_fg": "#000000", "text_select": "#0078d7",
            "button_bg": "#e1e1e1", "button_fg": "#000000", "button_active": "#c0c0c0",
            "labelframe_bg": "#f9f9f9", "border": "#cccccc",
            "select_active": "#cce8ff" 
        }

        self.dark_style = {
            "bg": "#1e1e1e", "fg": "#d4d4d4",
            "frame_bg": "#252526",
            "text_bg": "#1e1e1e", "text_fg": "#cccccc", "text_select": "#264f78",
            "button_bg": "#3c3c3c", "button_fg": "#ffffff", "button_active": "#505050",
            "labelframe_bg": "#2d2d30", "border": "#3e3e42",
            "select_active": "#0e639c" 
        }
        self.apply_style(self.light_style)

    def apply_style(self, style_dict):
        self.current_style = style_dict
        s = self.style
        s.configure(".", background=style_dict["bg"], foreground=style_dict["fg"])
        s.configure("TFrame", background=style_dict["frame_bg"])
        s.configure("TLabel", background=style_dict["frame_bg"], foreground=style_dict["fg"])
        s.configure("TLabelframe", background=style_dict["labelframe_bg"], bordercolor=style_dict["border"])
        s.configure("TLabelframe.Label", background=style_dict["labelframe_bg"], foreground=style_dict["fg"], font=("Microsoft YaHei", 9, "bold"))
        
        s.configure("Action.TButton", font=("Microsoft YaHei", 9), padding=4, 
                   background=style_dict["button_bg"], foreground=style_dict["button_fg"],
                   borderwidth=1, relief="flat")
        s.map("Action.TButton", 
             background=[('active', style_dict["button_active"]), ('pressed', style_dict["text_select"])],
             foreground=[('active', style_dict["fg"])])

        s.configure("TCheckbutton", background=style_dict["frame_bg"], foreground=style_dict["fg"])
        s.map("TCheckbutton", background=[('active', style_dict["frame_bg"]), ('selected', style_dict["frame_bg"])])
        
        s.configure("Prompt.TCheckbutton", background=style_dict["frame_bg"], foreground=style_dict["fg"])
        s.configure("Tag.TLabel", foreground="#0078d7", font=("Microsoft YaHei", 9, "bold"), background=style_dict["frame_bg"])
        s.configure("Index.TLabel", foreground="#888888", font=("Arial", 9), background=style_dict["frame_bg"])
        
        s.configure("GroupSelect.TLabel", foreground="#0078d7", font=("Microsoft YaHei", 8, "underline"), background=style_dict["frame_bg"])

        s.configure("Vertical.TScrollbar", troughcolor=style_dict["frame_bg"], background=style_dict["button_bg"], borderwidth=0, arrowcolor=style_dict["fg"])
        s.configure("Horizontal.TScrollbar", troughcolor=style_dict["frame_bg"], background=style_dict["button_bg"], borderwidth=0, arrowcolor=style_dict["fg"])

    def toggle_dark_mode(self):
        current_style = self.dark_style if self.is_dark_mode.get() else self.light_style
        self.apply_style(current_style)
        self.root.configure(bg=current_style["bg"])
        return current_style

# 可拖拽列表框（带回调）
class DraggableListbox(tk.Listbox):
    def __init__(self, master, on_reorder=None, **kwargs):
        super().__init__(master, **kwargs)
        self.on_reorder = on_reorder
        self.bind("<Button-1>", self.on_press)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.dragging_index = None

    def on_press(self, e):
        self.dragging_index = self.nearest(e.y)

    def on_drag(self, e):
        if self.dragging_index is None: return
        current = self.nearest(e.y)
        if current != self.dragging_index:
            text = self.get(self.dragging_index)
            self.delete(self.dragging_index)
            self.insert(current, text)
            self.dragging_index = current
            
    def on_release(self, e):
        if self.dragging_index is not None:
            self.dragging_index = None
            if self.on_reorder:
                self.on_reorder()

class ActionPlanGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("动作训练计划生成器 v4.8 (Excel格式优化版)")
        self.root.geometry("1400x900")
        
        self.style_config = StyleConfig(root)
        self.script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

        self.excel_data = None
        self.current_excel_path = None
        self.translation_map = {}
        self.combined_plan = [] 
        
        self.ui_rows = [] 
        
        self.col_mapping = INITIAL_COL_MAPPING.copy()
        
        self.action_categories = {}
        for k in self.col_mapping.keys():
            if "表情" in k:
                self.action_categories[k] = []
            elif "辅助" in k:
                self.action_categories[k] = {}
            else:
                self.action_categories[k] = {}
        self.category_order = [k for k in self.col_mapping.keys() if "辅助" not in k and "表情" not in k]

        self.check_vars = [] 
        self.text_widgets = {} 
        self.category_frames = {} 
        self.row_count_var = tk.StringVar(value="prompt数: 0")
        
        self.use_original_text = tk.BooleanVar(value=False) 

        self.setup_ui()
        self.root.after(100, self.select_excel_file_on_start)

    def select_excel_file_on_start(self):
        path = filedialog.askopenfilename(
            initialdir=self.script_dir,
            title="请选择Excel文件",
            filetypes=[("Excel", "*.xlsx")]
        )
        if path:
            self.current_excel_path = path
            self.load_excel(path)

    def update_theme(self, style_dict):
        for t_key in self.text_widgets:
            self.text_widgets[t_key].config(
                bg=style_dict["text_bg"], 
                fg=style_dict["text_fg"], 
                insertbackground=style_dict["fg"],
                selectbackground=style_dict["text_select"]
            )
        self.cat_listbox.config(
            bg=style_dict["text_bg"], 
            fg=style_dict["text_fg"],
            selectbackground=style_dict["text_select"]
        )
        self.canvas_l.config(bg=style_dict["frame_bg"])
        self.canvas_r.config(bg=style_dict["frame_bg"])
        self.root.configure(bg=style_dict["bg"])
        
    def toggle_language_display(self):
        self.update_ui_display()
        self.status_var.set(f"左侧预览已切换至{'原始Prompt' if self.use_original_text.get() else '翻译标签'}")

    def setup_ui(self):
        top_bar = ttk.Frame(self.root)
        top_bar.pack(fill="x", padx=10, pady=5)
        ttk.Checkbutton(top_bar, text="夜间模式", variable=self.style_config.is_dark_mode, 
                       command=lambda: self.update_theme(self.style_config.toggle_dark_mode())).pack(side="right", padx=5)
        ttk.Checkbutton(top_bar, text="左侧显示原始Prompt", variable=self.use_original_text, 
                       command=self.toggle_language_display).pack(side="right")

        main_paned = ttk.PanedWindow(self.root, orient="horizontal")
        main_paned.pack(fill="both", expand=True, padx=5, pady=5)

        self.left_frame = ttk.Frame(main_paned)
        main_paned.add(self.left_frame, weight=3)
        self.right_frame = ttk.Frame(main_paned)
        main_paned.add(self.right_frame, weight=4)

        self.create_left_panel()
        self.create_right_panel()

    def create_left_panel(self):
        self.canvas_l = tk.Canvas(self.left_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.left_frame, orient="vertical", command=self.canvas_l.yview)
        self.left_scroll_frame = ttk.Frame(self.canvas_l)
        
        self.canvas_l.configure(yscrollcommand=scrollbar.set, bg=self.style_config.current_style["frame_bg"])
        self.canvas_l_window = self.canvas_l.create_window((0, 0), window=self.left_scroll_frame, anchor="nw")
        
        self.left_scroll_frame.bind("<Configure>", lambda e: self.canvas_l.configure(scrollregion=self.canvas_l.bbox("all")))
        self.canvas_l.bind("<Configure>", lambda e: self.canvas_l.itemconfig(self.canvas_l_window, width=e.width))

        self.canvas_l.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        ScrollHelper.bind_mouse_scroll(self.left_scroll_frame, self.canvas_l)

        file_box = ttk.Frame(self.left_scroll_frame)
        file_box.pack(fill="x", pady=5, padx=5)
        self.file_label = ttk.Label(file_box, text="等待加载...", foreground="gray")
        self.file_label.pack(side="left", fill="x", expand=True)
        ttk.Button(file_box, text="打开Excel", command=self.select_excel_file_on_start, style="Action.TButton").pack(side="right")

        ctrl_frame = ttk.Frame(self.left_scroll_frame)
        ctrl_frame.pack(fill="x", pady=5, padx=5)
        
        order_box = ttk.LabelFrame(ctrl_frame, text="生成顺序（拖拽排序）")
        order_box.pack(side="left", fill="y", padx=5)
        
        self.cat_listbox = DraggableListbox(order_box, 
                                           on_reorder=self.sync_category_ui_order,
                                           height=5, width=25, borderwidth=0,
                                           bg=self.style_config.current_style["text_bg"], 
                                           fg=self.style_config.current_style["text_fg"])
        self.cat_listbox.pack(padx=5, pady=5)
        for c in self.category_order: self.cat_listbox.insert(tk.END, c)
        
        btn_box = ttk.Frame(ctrl_frame)
        btn_box.pack(side="left", fill="both", expand=True, padx=5)
        ttk.Button(btn_box, text="重置所有", command=self.reset_all_actions, style="Action.TButton").pack(fill="x", pady=2)
        ttk.Button(btn_box, text="生成 Prompt", command=self.generate_plan, style="Action.TButton").pack(fill="x", pady=2)
        ttk.Button(btn_box, text="导出 Excel", command=self.export_excel, style="Action.TButton").pack(fill="x", pady=2)
        ttk.Button(btn_box, text="+ 新增大类", command=self.add_new_category_dialog, style="Action.TButton").pack(fill="x", pady=2)

        self.categories_container = ttk.Frame(self.left_scroll_frame)
        self.categories_container.pack(fill="x", pady=5)
        
        for cat in self.col_mapping.keys():
            if "辅助" in cat or "表情" in cat: continue
            self.create_single_category_ui(cat)

        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(self.left_scroll_frame, textvariable=self.status_var, foreground="#0078d7").pack(fill="x", pady=10, padx=5)

    def sync_category_ui_order(self):
        new_order = list(self.cat_listbox.get(0, tk.END))
        self.category_order = new_order
        for cat_name in new_order:
            frame = self.category_frames.get(cat_name)
            if frame:
                frame.pack(in_=self.categories_container, fill="x", pady=5, padx=5)

    def create_single_category_ui(self, cat_name):
        f = ttk.Frame(self.categories_container)
        f.pack(fill="x", pady=5, padx=5)
        self.category_frames[cat_name] = f
        
        h = ttk.Frame(f)
        h.pack(fill="x", pady=2)
        ttk.Label(h, text=cat_name, font=("Microsoft YaHei", 10, "bold")).pack(side="left")
        
        b_frame = ttk.Frame(h)
        b_frame.pack(side="right")
        ttk.Button(b_frame, text="≡ 选择", width=6, command=lambda c=cat_name: self.open_manual_selection_window(c), style="Action.TButton").pack(side="left", padx=2)
        ttk.Button(b_frame, text="+ 添加", width=6, command=lambda c=cat_name: self.manual_add_action(c), style="Action.TButton").pack(side="left", padx=2)
        ttk.Button(b_frame, text="⟳ 重抽", width=6, command=lambda c=cat_name: self.reset_single_category(c), style="Action.TButton").pack(side="left", padx=2)
        
        t = tk.Text(f, height=4, font=("Consolas", 9), relief="flat", padx=5, pady=5,
                   bg=self.style_config.current_style["text_bg"], 
                   fg=self.style_config.current_style["text_fg"])
        t.pack(fill="x", pady=2)
        self.text_widgets[cat_name] = t

    def create_right_panel(self):
        tool_bar = ttk.Frame(self.right_frame)
        tool_bar.pack(fill="x", pady=5, padx=5)
        
        sel_box = ttk.Frame(tool_bar)
        sel_box.pack(side="left", padx=2)
        ttk.Button(sel_box, text="全选", width=4, command=self.select_all_prompt, style="Action.TButton").pack(side="left", padx=1)
        ttk.Button(sel_box, text="反选", width=4, command=self.invert_prompt, style="Action.TButton").pack(side="left", padx=1)
        
        mov_box = ttk.Frame(tool_bar)
        mov_box.pack(side="left", padx=5)
        ttk.Button(mov_box, text="▲", width=3, command=lambda: self.move_prompt(-1), style="Action.TButton").pack(side="left", padx=1)
        ttk.Button(mov_box, text="▼", width=3, command=lambda: self.move_prompt(1), style="Action.TButton").pack(side="left", padx=1)

        action_box = ttk.Frame(tool_bar)
        action_box.pack(side="left", padx=5)
        ttk.Button(action_box, text="复制", width=4, command=self.copy_selected_prompt, style="Action.TButton").pack(side="left", padx=1)
        ttk.Button(action_box, text="删除", width=4, command=self.delete_selected_prompt, style="Action.TButton").pack(side="left", padx=1)

        mod_box = ttk.Frame(tool_bar)
        mod_box.pack(side="left", padx=5)
        ttk.Button(mod_box, text="添加tag", width=8, command=self.add_extra_prompt, style="Action.TButton").pack(side="left", padx=1)
        ttk.Button(mod_box, text="添加辅助/表情", width=12, command=self.open_add_action_window, style="Action.TButton").pack(side="left", padx=1)

        ttk.Label(tool_bar, textvariable=self.row_count_var, foreground="#0078d7").pack(side="right", padx=5)

        list_frame = ttk.LabelFrame(self.right_frame, text="Prompt 预览（生成的训练集）")
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.canvas_r = tk.Canvas(list_frame, highlightthickness=0, bg=self.style_config.current_style["frame_bg"])
        scr_r = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas_r.yview)
        
        self.prompt_inner = ttk.Frame(self.canvas_r)
        
        self.canvas_r_window = self.canvas_r.create_window((0,0), window=self.prompt_inner, anchor="nw")
        self.canvas_r.configure(yscrollcommand=scr_r.set)
        
        self.prompt_inner.bind("<Configure>", lambda e: self.canvas_r.configure(scrollregion=self.canvas_r.bbox("all")))
        self.canvas_r.bind("<Configure>", lambda e: self.canvas_r.itemconfig(self.canvas_r_window, width=e.width))

        self.canvas_r.pack(side="left", fill="both", expand=True)
        scr_r.pack(side="right", fill="y")
        
        ScrollHelper.bind_mouse_scroll(self.prompt_inner, self.canvas_r)

    # --- 逻辑功能 ---

    def load_excel(self, path):
        try:
            self.excel_data = pd.read_excel(path, header=None)
            self.translation_map = {}
            for cat in self.col_mapping.keys():
                self.process_category_data(cat)
            self.update_ui_display()
            self.file_label.config(text=os.path.basename(path))
            self.status_var.set("Excel 加载成功")
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件: {str(e)}")

    def process_category_data(self, cat_name):
        col_indices = self.col_mapping.get(cat_name, [])
        is_aux = "辅助" in cat_name
        is_emo = "表情" in cat_name

        if is_emo:
            self.action_categories[cat_name] = []
        elif is_aux:
            self.action_categories[cat_name] = {}
        else:
            self.action_categories[cat_name] = {}

        for col_idx in col_indices:
            if col_idx >= self.excel_data.shape[1]: continue
            
            col_data = self.excel_data.iloc[:, col_idx]
            trans_data = None
            if col_idx + 1 < self.excel_data.shape[1]:
                trans_data = self.excel_data.iloc[:, col_idx + 1]
            
            if len(col_data) > 0:
                sub_cat = str(col_data.iloc[0]).strip()
            else:
                continue
            
            actions = []
            for i in range(1, len(col_data)):
                act_val = col_data.iloc[i]
                if pd.notna(act_val):
                    act = str(act_val).strip()
                    if act:
                        actions.append(act)
                        if trans_data is not None and i < len(trans_data):
                            t_val = trans_data.iloc[i]
                            if pd.notna(t_val):
                                t_str = str(t_val).strip()
                                if t_str and t_str.lower() != 'nan':
                                    self.translation_map[act] = t_str
            
            if not actions: continue

            if is_emo: 
                self.action_categories[cat_name].extend(actions)
            elif is_aux: 
                if sub_cat not in self.action_categories[cat_name]:
                    self.action_categories[cat_name][sub_cat] = {}
                for a in actions:
                    self.action_categories[cat_name][sub_cat][a] = 1
            else:
                if sub_cat not in self.action_categories[cat_name]:
                    self.action_categories[cat_name][sub_cat] = []
                
                min_groups = 3
                if len(actions) < min_groups:
                    selected_indices = list(range(len(actions)))
                else:
                    num_selected = random.randint(min_groups, min(len(actions), 6))
                    selected_indices = random.sample(range(len(actions)), num_selected)
                
                for idx in selected_indices:
                    action = actions[idx]
                    repeat_count = random.randint(3, 5)
                    self.action_categories[cat_name][sub_cat].append({action: repeat_count})

    def col_str_to_index(self, col_str):
        col_str = col_str.upper().strip()
        num = 0
        for c in col_str:
            if 'A' <= c <= 'Z':
                num = num * 26 + (ord(c) - ord('A') + 1)
            else:
                return -1
        return num - 1

    def get_all_options_for_category(self, cat_name):
        col_indices = self.col_mapping.get(cat_name, [])
        options = [] 
        
        for col_idx in col_indices:
            if col_idx >= self.excel_data.shape[1]: continue
            
            col_data = self.excel_data.iloc[:, col_idx]
            trans_data = None
            if col_idx + 1 < self.excel_data.shape[1]:
                trans_data = self.excel_data.iloc[:, col_idx + 1]
            
            if len(col_data) > 0:
                sub_cat = str(col_data.iloc[0]).strip()
            else:
                continue
            
            for i in range(1, len(col_data)):
                act_val = col_data.iloc[i]
                if pd.notna(act_val):
                    act = str(act_val).strip()
                    if act:
                        t_str = ""
                        if trans_data is not None and i < len(trans_data):
                            t_val = trans_data.iloc[i]
                            if pd.notna(t_val):
                                val = str(t_val).strip()
                                if val and val.lower() != 'nan':
                                    t_str = val
                        options.append({
                            "sub": sub_cat,
                            "act": act,
                            "trans": t_str
                        })
        return options

    def open_manual_selection_window(self, cat_name):
        if self.excel_data is None:
            messagebox.showwarning("提示", "请先加载 Excel 文件")
            return

        options = self.get_all_options_for_category(cat_name)
        if not options:
            messagebox.showinfo("提示", "该类目下没有可选项")
            return

        win = tk.Toplevel(self.root)
        win.title(f"选择 - {cat_name}")
        win.geometry("800x600")
        win.configure(bg=self.style_config.current_style["bg"])

        top_f = ttk.Frame(win)
        top_f.pack(fill="x", pady=5, padx=10)
        
        ttk.Label(top_f, text="点击文本框选中（标记序号），再次点击取消。", foreground="#0078d7").pack(side="left")
        
        ttk.Label(top_f, text="|  重复次数:").pack(side="left", padx=(15, 5))
        count_var = tk.StringVar(value="5")
        tk.Entry(top_f, textvariable=count_var, width=5, justify="center").pack(side="left", padx=5)
        
        ttk.Button(top_f, text="确认修改", command=lambda: self.apply_manual_selection(win, cat_name, count_var), style="Action.TButton").pack(side="left", padx=15)

        main_container = tk.Canvas(win, bg=self.style_config.current_style["bg"], highlightthickness=0)
        v_scroll = ttk.Scrollbar(win, orient="vertical", command=main_container.yview)
        content_frame = ttk.Frame(main_container)
        
        main_container.create_window((0, 0), window=content_frame, anchor="nw")
        main_container.configure(yscrollcommand=v_scroll.set)
        
        v_scroll.pack(side="right", fill="y")
        main_container.pack(side="left", fill="both", expand=True)
        
        content_frame.bind("<Configure>", lambda e: main_container.configure(scrollregion=main_container.bbox("all")))
        ScrollHelper.bind_mouse_scroll(content_frame, main_container)

        self.temp_selected = [] 
        self.btn_map = {} 

        grouped = {}
        for opt in options:
            s = opt["sub"]
            if s not in grouped: grouped[s] = []
            grouped[s].append(opt)

        row_idx = 0
        COLUMNS = 4 
        
        for sub, items in grouped.items():
            ttk.Label(content_frame, text=sub, font=("Microsoft YaHei", 10, "bold")).grid(row=row_idx, column=0, columnspan=COLUMNS, sticky="w", padx=10, pady=(10, 5))
            row_idx += 1
            
            for i, item in enumerate(items):
                r = row_idx + i // COLUMNS
                c = i % COLUMNS
                
                display_text = item["trans"] if item["trans"] else item["act"]
                
                btn = tk.Button(content_frame, text=display_text, font=("Microsoft YaHei", 9),
                                bg=self.style_config.current_style["button_bg"],
                                fg=self.style_config.current_style["button_fg"],
                                relief="flat", padx=5, pady=2, width=20, wraplength=140)
                
                btn.config(command=lambda b=btn, it=item: self.toggle_selection(b, it))
                btn.grid(row=r, column=c, padx=5, pady=2, sticky="nsew")
                
                self.btn_map[item["act"]] = btn 
            
            row_idx += (len(items) + COLUMNS - 1) // COLUMNS

    def toggle_selection(self, btn, item):
        if item in self.temp_selected:
            self.temp_selected.remove(item)
        else:
            self.temp_selected.append(item)
        self.refresh_selection_visuals()

    def refresh_selection_visuals(self):
        style = self.style_config.current_style
        
        for act, btn in self.btn_map.items():
             btn.config(bg=style["button_bg"], fg=style["button_fg"], text=btn.cget("text").split("] ")[-1])
        
        for idx, item in enumerate(self.temp_selected):
            btn = self.btn_map.get(item["act"])
            if btn:
                raw_text = item["trans"] if item["trans"] else item["act"]
                new_text = f"[{idx+1}] {raw_text}"
                btn.config(bg=style.get("select_active", "#cce8ff"), fg=style["text_select"], text=new_text)

    def apply_manual_selection(self, win, cat_name, count_var):
        try:
            count = int(count_var.get())
            if count <= 0: raise ValueError
        except:
            messagebox.showerror("错误", "重复次数必须为正整数")
            return

        if not self.temp_selected:
            if not messagebox.askyesno("确认", "未选择任何动作，是否清空该类目？"):
                return
            self.action_categories[cat_name] = {}
        else:
            new_data = {}
            for item in self.temp_selected:
                sub = item["sub"]
                act = item["act"]
                if sub not in new_data:
                    new_data[sub] = []
                new_data[sub].append({act: count})
            
            self.action_categories[cat_name] = new_data
        
        self.update_single_category_ui(cat_name)
        win.destroy()

    def add_new_category_dialog(self):
        if self.excel_data is None:
            messagebox.showwarning("提示", "请先加载 Excel 文件")
            return
            
        dialog = simpledialog.askstring("新增大类", 
            "请输入列号 (例如: AA, AB 或 AF)\n注意：输入的是Prompt列，程序会自动读取下一列作为翻译。")
        
        if not dialog: return
        
        try:
            col_strs = [s.strip() for s in dialog.split(',') if s.strip()]
            indices = []
            for s in col_strs:
                idx = self.col_str_to_index(s)
                if idx == -1: raise ValueError(f"列号格式错误: {s}")
                indices.append(idx)
            
            cat_idx = len([k for k in self.col_mapping if "辅助" not in k and "表情" not in k]) + 1
            new_cat_name = f"第{cat_idx}类动作（自定义）"
            
            self.col_mapping[new_cat_name] = indices
            self.category_order.append(new_cat_name)
            self.cat_listbox.insert(tk.END, new_cat_name)
            
            self.process_category_data(new_cat_name)
            self.create_single_category_ui(new_cat_name)
            self.update_single_category_ui(new_cat_name)
            
            self.canvas_l.update_idletasks()
            self.canvas_l.yview_moveto(1)
            
            self.status_var.set(f"已添加 {new_cat_name}")
            
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def manual_add_action(self, cat):
        dialog = BigTextInputDialog(self.root, f"添加到 [{cat}]", "请输入动作名称：")
        text = dialog.result_text
        if text:
            repeats = random.randint(3, 5)
            sub_cats = list(self.action_categories[cat].keys())
            target_sub = "手动添加"
            if sub_cats and "手动" not in sub_cats:
                target_sub = sub_cats[0]
            
            if target_sub not in self.action_categories[cat]:
                self.action_categories[cat][target_sub] = {}
            
            self.action_categories[cat][target_sub][text] = repeats
            self.update_single_category_ui(cat)

    def update_single_category_ui(self, cat):
        widget = self.text_widgets.get(cat)
        if not widget: return
        widget.delete(1.0, tk.END)
        data = self.action_categories.get(cat, {})
        
        def truncate_text(text, max_len=30):
            if max_len and len(text) > max_len:
                return text[:max_len] + "..."
            return text
            
        for sub, acts_list in data.items():
            if isinstance(acts_list, list):
                groups_str = []
                for action_group in acts_list:
                    if action_group:
                        action = list(action_group.keys())[0]
                        repeat_count = action_group[action]
                        
                        if not self.use_original_text.get():
                            preview_text = self.translation_map.get(action, action)
                        else:
                            preview_text = action
                            
                        truncated_preview = truncate_text(preview_text)
                        groups_str.append(f"{truncated_preview}*{repeat_count}")
                
                acts_str = " | ".join(groups_str)
                line = f"{sub}: {acts_str}\n"
                widget.insert(tk.END, line)
            else:
                items = []
                for k,v in acts_list.items():
                     if not self.use_original_text.get():
                         display_k = self.translation_map.get(k, k)
                     else:
                         display_k = k
                     items.append(f"{truncate_text(display_k)}*{v}")
                acts_str = ", ".join(items)
                line = f"{sub}: {acts_str}\n"
                widget.insert(tk.END, line)

    def update_ui_display(self):
        for cat in self.text_widgets:
            self.update_single_category_ui(cat)

    def reset_single_category(self, cat):
        if self.excel_data is None: return
        self.process_category_data(cat)
        self.update_single_category_ui(cat)

    def generate_plan(self):
        if self.excel_data is None: 
            messagebox.showwarning("警告", "请先加载Excel文件")
            return
            
        self.category_order = [self.cat_listbox.get(i) for i in range(self.cat_listbox.size())]
        raw_plan = []
        for cat in self.category_order:
            data = self.action_categories.get(cat, {})
            for sub_acts_list in data.values():
                if sub_acts_list:
                    if isinstance(sub_acts_list, list):
                        for action_group in sub_acts_list:
                            if action_group:
                                action = list(action_group.keys())[0]
                                times = action_group[action]
                                raw_plan.extend([action] * times)
                    else:
                        for k, v in sub_acts_list.items():
                            raw_plan.extend([k] * v)
        
        self.combined_plan = []
        for action in raw_plan:
            self.combined_plan.append({
                "original_action": action, 
                "translation": self.translation_map.get(action, "无标签"), 
                "checked": False
            })
        self.refresh_prompt_list()

    def create_one_ui_row(self, item):
        row = ttk.Frame(self.prompt_inner)
        
        var = tk.BooleanVar(value=item["checked"])
        var.trace_add("write", lambda *args, v=var, r=row: self.on_check_change_fast(v, r))
        self.check_vars.append(var)
        
        trans_text = item['translation']
        if not trans_text or trans_text == "无标签":
            trans_text = "未定义"
        
        lbl_tag = ttk.Label(row, text=f"[{trans_text}]", style="Tag.TLabel", width=15, anchor="e")
        lbl_tag.pack(side="left", padx=(5, 2))
        
        btn_grp = ttk.Label(row, text="[全选]", style="GroupSelect.TLabel", cursor="hand2")
        btn_grp.pack(side="left", padx=(0, 5))
        btn_grp.bind("<Button-1>", lambda e, t=trans_text: self.select_group_by_translation(t))
        
        lbl_idx = ttk.Label(row, text="No.-", style="Index.TLabel", width=5)
        lbl_idx.pack(side="left", padx=(0, 5))
        
        action_text = item['original_action'].replace('\n', '')
        cb = ttk.Checkbutton(row, text=action_text, variable=var, style="Prompt.TCheckbutton")
        cb.pack(side="left", fill="x", expand=True)

        return {
            "frame": row,
            "label_idx": lbl_idx,
            "var": var
        }
    
    def select_group_by_translation(self, trans_text):
        if not trans_text: return
        
        target_indices = []
        all_checked = True
        
        for idx, item in enumerate(self.combined_plan):
            if item['translation'] == trans_text:
                target_indices.append(idx)
                if not item['checked']:
                    all_checked = False
        
        if not target_indices: return
        
        new_state = not all_checked
        for idx in target_indices:
            self.check_vars[idx].set(new_state)

    def refresh_prompt_list(self):
        for w in self.prompt_inner.winfo_children(): w.destroy()
        self.check_vars = [] 
        self.ui_rows = [] 
        
        self.row_count_var.set(f"prompt数: {len(self.combined_plan)}")
        
        for idx, item in enumerate(self.combined_plan):
            ui_data = self.create_one_ui_row(item)
            ui_data["frame"].pack(fill="x", padx=2, pady=1)
            ui_data["label_idx"].config(text=f"No.{idx+1}")
            self.ui_rows.append(ui_data)

    def on_check_change_fast(self, var, row_frame):
        current_idx = -1
        for i, uiro in enumerate(self.ui_rows):
            if uiro["frame"] == row_frame:
                current_idx = i
                break
        
        if current_idx != -1:
            self.combined_plan[current_idx]["checked"] = var.get()

    def move_prompt(self, direction):
        if not self.combined_plan: return
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: return
        
        moved = False
        
        if direction == -1: 
            for i in sorted(indices):
                if i > 0 and not self.combined_plan[i-1]["checked"]:
                    self.combined_plan[i], self.combined_plan[i-1] = self.combined_plan[i-1], self.combined_plan[i]
                    self.check_vars[i], self.check_vars[i-1] = self.check_vars[i-1], self.check_vars[i]
                    self.ui_rows[i], self.ui_rows[i-1] = self.ui_rows[i-1], self.ui_rows[i]
                    moved = True
        else: 
            for i in sorted(indices, reverse=True):
                if i < len(self.combined_plan) - 1 and not self.combined_plan[i+1]["checked"]:
                    self.combined_plan[i], self.combined_plan[i+1] = self.combined_plan[i+1], self.combined_plan[i]
                    self.check_vars[i], self.check_vars[i+1] = self.check_vars[i+1], self.check_vars[i]
                    self.ui_rows[i], self.ui_rows[i+1] = self.ui_rows[i+1], self.ui_rows[i]
                    moved = True

        if moved:
            self.reorder_ui_rows()

    def reorder_ui_rows(self):
        for row_data in self.ui_rows:
            row_data["frame"].pack_forget()
            
        for idx, row_data in enumerate(self.ui_rows):
            frame = row_data["frame"]
            frame.pack(side="top", fill="x", padx=2, pady=1)
            row_data["label_idx"].config(text=f"No.{idx+1}")

    def delete_selected_prompt(self):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: return

        for i in sorted(indices, reverse=True):
            self.ui_rows[i]["frame"].destroy() 
            del self.combined_plan[i]
            del self.check_vars[i]
            del self.ui_rows[i]

        self.row_count_var.set(f"prompt数: {len(self.combined_plan)}")
        for idx, row_data in enumerate(self.ui_rows):
            row_data["label_idx"].config(text=f"No.{idx+1}")

    def copy_selected_prompt(self):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices: return

        insert_tasks = [] 
        
        offset = 0
        for i in sorted(indices):
            current_idx = i + offset
            original_item = self.combined_plan[current_idx]
            
            copy_item = original_item.copy()
            copy_item["checked"] = True 
            
            insert_pos = current_idx + 1
            insert_tasks.append((insert_pos, copy_item))
            offset += 1
            
        for pos, item in insert_tasks:
            self.combined_plan.insert(pos, item)
            ui_data = self.create_one_ui_row(item)
            self.ui_rows.insert(pos, ui_data)
            
        self.reorder_ui_rows()
        self.row_count_var.set(f"prompt数: {len(self.combined_plan)}")
        self.status_var.set(f"已复制 {len(insert_tasks)} 条 Prompt")

    def add_extra_prompt(self):
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices:
            messagebox.showinfo("提示", "请先勾选需要添加tag的行")
            return
            
        dialog = BigTextInputDialog(self.root, "添加tag", "请输入要添加的 tag (例如: masterpiece)：")
        tag = dialog.result_text
        
        if tag:
            for i in indices:
                original = self.combined_plan[i]["original_action"]
                if original.startswith(","):
                    self.combined_plan[i]["original_action"] = f"{tag}{original}"
                else:
                    self.combined_plan[i]["original_action"] = f"{tag}, {original}"
            self.refresh_prompt_list()

    def select_all_prompt(self):
        for var in self.check_vars:
            var.set(True)
    
    def invert_prompt(self):
        for var in self.check_vars:
            var.set(not var.get())

    def reset_all_actions(self):
        if self.current_excel_path:
            self.col_mapping = INITIAL_COL_MAPPING.copy()
            for w in self.categories_container.winfo_children():
                w.destroy()
            self.text_widgets = {}
            self.category_order = [k for k in self.col_mapping.keys() if "辅助" not in k and "表情" not in k]
            
            self.load_excel(self.current_excel_path)
            
            for cat in self.col_mapping.keys():
                if "辅助" in cat or "表情" in cat: continue
                self.create_single_category_ui(cat)
            self.update_ui_display()

            self.combined_plan = []
            for w in self.prompt_inner.winfo_children(): w.destroy()
            self.ui_rows = []
            self.status_var.set("已完全重置")

    def open_add_action_window(self):
        if not self.combined_plan: 
            messagebox.showinfo("提示", "请先生成基础计划")
            return
            
        win = tk.Toplevel(self.root)
        win.title("添加辅助动作与表情")
        screen_h = self.root.winfo_screenheight()
        win_h = int(screen_h * 0.85)
        win.geometry(f"1300x{win_h}")
        win.configure(bg=self.style_config.current_style["bg"])
        
        top_f = ttk.Frame(win)
        top_f.pack(fill="x", pady=5)
        ttk.Label(top_f, text="勾选下方内容，点击底部按钮确认添加到主界面选中的 Prompt 后").pack()

        main_container = tk.Canvas(win, bg=self.style_config.current_style["bg"], highlightthickness=0)
        h_scroll = ttk.Scrollbar(win, orient="horizontal", command=main_container.xview)
        
        content_frame = ttk.Frame(main_container)
        main_container.create_window((0, 0), window=content_frame, anchor="nw")
        main_container.configure(xscrollcommand=h_scroll.set)
        
        h_scroll.pack(side="bottom", fill="x")
        main_container.pack(side="top", fill="both", expand=True)

        def update_scroll_region(event):
            main_container.configure(scrollregion=main_container.bbox("all"))
        content_frame.bind("<Configure>", update_scroll_region)
        
        def _h_scroll(event):
             if platform.system() == 'Windows':
                main_container.xview_scroll(int(-1 * (event.delta / 120)), "units")
             else:
                if event.num == 4: main_container.xview_scroll(-1, "units")
                elif event.num == 5: main_container.xview_scroll(1, "units")
        
        if platform.system() == 'Linux':
            win.bind("<Button-4>", _h_scroll)
            win.bind("<Button-5>", _h_scroll)
        else:
            win.bind("<MouseWheel>", _h_scroll)

        self.add_vars = {"aux": [], "emo": []}
        
        def truncate_text(text, max_len=30):
            if max_len and len(text) > max_len:
                return text[:max_len] + "..."
            return text
            
        columns_data = []
        aux_data = self.action_categories.get("辅助动作（S/U/W列）", {})
        for sub_cat, acts in aux_data.items():
            item_list = []
            for a in acts.keys():
                t = self.translation_map.get(a, "")
                truncated_a = truncate_text(a)
                label = f"{truncated_a}\n({t})" if t else truncated_a
                item_list.append((a, label))
            columns_data.append((sub_cat, item_list))
            
        emo_acts = self.action_categories.get("表情（Y列）", [])
        emo_list = []
        for a in emo_acts:
            t = self.translation_map.get(a, "")
            truncated_a = truncate_text(a)
            label = f"{truncated_a}\n({t})" if t else truncated_a
            emo_list.append((a, label))
        columns_data.append(("表情", emo_list))

        ITEM_HEIGHT = 45 
        available_height = win_h - 100 
        MAX_ITEMS_PER_COL = max(10, available_height // ITEM_HEIGHT)

        current_col_idx = 0
        
        for title, items in columns_data:
            total_items = len(items)
            num_sub_cols = (total_items + MAX_ITEMS_PER_COL - 1) // MAX_ITEMS_PER_COL
            
            cat_frame = ttk.LabelFrame(content_frame, text=title)
            cat_frame.grid(row=0, column=current_col_idx, padx=5, pady=5, sticky="n")
            
            for i in range(num_sub_cols):
                sub_col_frame = ttk.Frame(cat_frame)
                sub_col_frame.pack(side="left", fill="y", padx=5, anchor="n")
                
                start = i * MAX_ITEMS_PER_COL
                end = min(start + MAX_ITEMS_PER_COL, total_items)
                
                for idx in range(start, end):
                    code, label = items[idx]
                    v = tk.BooleanVar()
                    cb = ttk.Checkbutton(sub_col_frame, text=label, variable=v)
                    cb.pack(anchor="w", pady=2)
                    if title == "表情": self.add_vars["emo"].append((code, v))
                    else: self.add_vars["aux"].append((code, v))
            
            current_col_idx += 1

        btn = ttk.Button(win, text="确认添加", width=25, command=lambda: self.apply_add(win), style="Action.TButton")
        btn.place(relx=0.5, rely=0.95, anchor="center")

    def apply_add(self, win):
        tags = []
        for a, v in self.add_vars["aux"]:
            if v.get(): tags.append(a)
        for a, v in self.add_vars["emo"]:
            if v.get(): tags.append(a)
        
        if not tags: 
            win.destroy()
            return
        
        indices = [i for i, x in enumerate(self.combined_plan) if x["checked"]]
        if not indices:
            messagebox.showinfo("提示", "请在主界面勾选需要添加tag的 Prompt 行")
            return

        for i in indices:
            original = self.combined_plan[i]["original_action"]
            self.combined_plan[i]["original_action"] = f"{','.join(tags)}, {original}"
        
        self.refresh_prompt_list()
        win.destroy()

    def export_excel(self):
        final = [x["original_action"] for x in self.combined_plan]
        if not final: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            # 1. 恢复为整数
            seeds = [random.randint(10000000000000, 99999999999999) for _ in final]
            completion = ["" for _ in final]
            df = pd.DataFrame({
                "序号": range(1, len(final)+1),
                "动作Prompt": final,
                "种子": seeds,
                "完成情况": completion
            })
            
            # 2. 使用 ExcelWriter + openpyxl 引擎
            try:
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    # 3. 设置格式：C列为数值但不显示科学计数法
                    ws = writer.sheets['Sheet1']
                    # 设置C列宽度以便完整显示
                    ws.column_dimensions['C'].width = 20
                    
                    # 遍历C列所有单元格(跳过表头)并设置 number_format
                    # openpyxl 中列索引 C 对应第3列
                    for row in range(2, len(final) + 2):
                        cell = ws.cell(row=row, column=3)
                        cell.number_format = '0' 
                        
                messagebox.showinfo("成功", "导出完成")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        if os.name == 'nt':
            root.option_add("*Font", ("Microsoft YaHei", 9))
        else:
            root.option_add("*Font", ("WenQuanYi Micro Hei", 9))
        
        app = ActionPlanGenerator(root)
        root.mainloop()
    except Exception:
        print(traceback.format_exc())