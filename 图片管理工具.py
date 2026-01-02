import sys
import os

# 1. 禁用高DPI缩放
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
os.environ["QT_SCALE_FACTOR"] = "1"

import re
import json
import math
import subprocess
from PIL import Image, ImageFile

ImageFile.LOAD_TRUNCATED_IMAGES = True

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, QComboBox, 
                             QMessageBox, QProgressBar, QGraphicsView, 
                             QGraphicsScene, QGraphicsItem, QGraphicsObject,
                             QGraphicsDropShadowEffect, QDialog, QFrame)
from PyQt6.QtCore import (Qt, QThread, pyqtSignal, QSettings, QRectF, 
                          QPointF, QTimer)
from PyQt6.QtGui import (QPixmap, QPainter, QPen, QColor, 
                         QTransform, QImage, QPainterPath)

from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader as PDFImageReader

# 尝试导入 win32com 用于控制 Photoshop
try:
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# =========================================================
# 0. UI 样式配置
# =========================================================

STYLE_DARK = """
QMainWindow { background-color: #1e1e1e; }
QWidget { color: #b0b0b0; font-family: "Segoe UI", "Microsoft YaHei"; font-size: 14px; }
QPushButton {
    background-color: #2d2d2d; color: #e0e0e0; border: 1px solid #3e3e3e;
    border-radius: 8px; padding: 8px 16px; font-weight: 500;
}
QPushButton:hover { background-color: #3e3e3e; border: 1px solid #505050; }
QPushButton:pressed { background-color: #007acc; border: 1px solid #007acc; color: white; }
QComboBox {
    background-color: #2d2d2d; color: white; border: 1px solid #3e3e3e;
    border-radius: 6px; padding: 5px;
}
QProgressBar { border: none; background-color: #2d2d2d; border-radius: 2px; height: 4px; }
QProgressBar::chunk { background-color: #007acc; border-radius: 2px; }
QGraphicsView { border: none; background-color: #1e1e1e; }
"""

STYLE_LIGHT = """
QMainWindow { background-color: #f5f5f5; }
QWidget { color: #333333; font-family: "Segoe UI", "Microsoft YaHei"; font-size: 14px; }
QPushButton {
    background-color: #ffffff; color: #333333; border: 1px solid #cccccc;
    border-radius: 8px; padding: 8px 16px; font-weight: 500;
}
QPushButton:hover { background-color: #e6e6e6; border: 1px solid #b3b3b3; }
QPushButton:pressed { background-color: #007acc; border: 1px solid #007acc; color: white; }
QComboBox {
    background-color: #ffffff; color: #333333; border: 1px solid #cccccc;
    border-radius: 6px; padding: 5px;
}
QProgressBar { border: none; background-color: #e0e0e0; border-radius: 2px; height: 4px; }
QProgressBar::chunk { background-color: #007acc; border-radius: 2px; }
QGraphicsView { border: none; background-color: #f0f0f0; }
"""

# =========================================================
# 1. 加载线程
# =========================================================
class ImageLoaderThread(QThread):
    progress_signal = pyqtSignal(int, int)
    item_loaded_signal = pyqtSignal(object, str) 

    def __init__(self, image_paths, target_width):
        super().__init__()
        self.image_paths = image_paths
        self.target_width = target_width
        self.is_running = True

    def run(self):
        total = len(self.image_paths)
        for i, path in enumerate(self.image_paths):
            if not self.is_running: break
            try:
                pil_img = Image.open(path)
                if pil_img.mode != "RGBA":
                    pil_img = pil_img.convert("RGBA")
                w, h = pil_img.size
                if w > self.target_width * 1.5: 
                    ratio = self.target_width / float(w)
                    new_h = int(h * ratio)
                    pil_img = pil_img.resize((self.target_width, new_h), Image.Resampling.LANCZOS)
                self.item_loaded_signal.emit(pil_img, path)
            except Exception:
                pass 
            self.progress_signal.emit(i + 1, total)

    def stop(self):
        self.is_running = False
        self.wait()

# =========================================================
# 2. 图元类
# =========================================================
class ThumbnailItem(QGraphicsObject):
    def __init__(self, pixmap, path, index):
        super().__init__()
        self.internal_pixmap = pixmap 
        self.path = path
        self.index_id = index 
        self.setAcceptHoverEvents(True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable)
        self.is_hovered = False
        self.path_clip = None
        self.rect_cache = QRectF()

    def boundingRect(self):
        if self.internal_pixmap and not self.internal_pixmap.isNull():
            return QRectF(0, 0, self.internal_pixmap.width(), self.internal_pixmap.height())
        return QRectF(0, 0, 100, 100)

    def paint(self, painter, option, widget=None):
        rect = self.boundingRect()
        if rect != self.rect_cache:
            self.rect_cache = rect
            self.path_clip = QPainterPath()
            self.path_clip.addRoundedRect(rect, 12, 12)

        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.save()
        painter.setClipPath(self.path_clip)
        if self.internal_pixmap and not self.internal_pixmap.isNull():
            painter.drawPixmap(0, 0, self.internal_pixmap)
        painter.restore()

        is_selected = self.isSelected()
        if is_selected or self.is_hovered:
            painter.save()
            if is_selected:
                painter.fillPath(self.path_clip, QColor(0, 122, 204, 60))
                pen = QPen(QColor("#007acc"), 4)
                painter.setPen(pen)
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawPath(self.path_clip)
                
                check_size = 24
                check_x = rect.width() - check_size - 8
                check_y = rect.height() - check_size - 8
                painter.setPen(Qt.PenStyle.NoPen)
                painter.setBrush(QColor("#007acc"))
                painter.drawEllipse(int(check_x), int(check_y), check_size, check_size)
                
                pen_check = QPen(QColor("white"), 2)
                painter.setPen(pen_check)
                painter.drawLine(int(check_x + 6), int(check_y + 12), int(check_x + 10), int(check_y + 16))
                painter.drawLine(int(check_x + 10), int(check_y + 16), int(check_x + 18), int(check_y + 8))
            elif self.is_hovered:
                painter.fillPath(self.path_clip, QColor(255, 255, 255, 40))
            painter.restore()

    def hoverEnterEvent(self, event):
        self.is_hovered = True
        self.update()
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        self.is_hovered = False
        self.update()
        super().hoverLeaveEvent(event)

# =========================================================
# 3. 插入指示器
# =========================================================
class DropIndicator(QGraphicsItem):
    def __init__(self, height=100):
        super().__init__()
        self.height = height
        
    def boundingRect(self):
        return QRectF(-6, -6, 12, self.height + 12)
        
    def paint(self, painter, option, widget=None):
        color = QColor("#00FFFF") 
        pen = QPen(color)
        pen.setWidth(6)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(pen)
        painter.drawLine(0, 0, 0, self.height)
        
        painter.setBrush(color)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(-6, -6, 12, 12)
        painter.drawEllipse(-6, self.height-6, 12, 12)

# =========================================================
# 4. 流式布局
# =========================================================
class FlowLayoutView(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.scene = QGraphicsScene(self)
        self.setScene(self.scene)
        self.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setViewportUpdateMode(QGraphicsView.ViewportUpdateMode.FullViewportUpdate)
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        
        self.items_list = []
        self.item_width = 200
        self.spacing = 20
        self.margin = 30
        
        self.dragging_items = []
        self.drag_start_pos = QPointF()
        self.is_dragging_mode = False 
        self.insert_index = -1
        
        self.indicator = DropIndicator()
        self.indicator.setZValue(999)
        self.scene.addItem(self.indicator)
        self.indicator.hide()
        
        self.scroll_timer = QTimer()
        self.scroll_timer.timeout.connect(self.auto_scroll)
        self.scroll_step = 0    

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.relayout()

    def add_item_from_pil(self, pil_image, path):
        try:
            if pil_image.mode != "RGBA":
                pil_image = pil_image.convert("RGBA")
            data = pil_image.tobytes("raw", "RGBA")
            qimage = QImage(data, pil_image.width, pil_image.height, QImage.Format.Format_RGBA8888)
            pixmap = QPixmap.fromImage(qimage.copy())
            
            index = len(self.items_list)
            item = ThumbnailItem(pixmap, path, index)
            if not pixmap.isNull():
                scale = self.item_width / pixmap.width()
                item.setTransform(QTransform().scale(scale, scale))
            self.scene.addItem(item)
            self.items_list.append(item)
            self.relayout()
        except Exception:
            pass

    def clear_items(self):
        self.items_list.clear()
        self.scene.clear()
        self.indicator = DropIndicator()
        self.indicator.setZValue(999)
        self.scene.addItem(self.indicator)
        self.indicator.hide()

    def set_scale(self, width):
        self.item_width = width
        for item in self.items_list:
            if not item.internal_pixmap.isNull():
                scale = width / item.internal_pixmap.width()
                item.setTransform(QTransform().scale(scale, scale))
        self.relayout()

    def relayout(self):
        if not self.items_list: return
        view_width = self.viewport().width()
        x = self.margin
        y = self.margin
        
        for item in self.items_list:
            if item in self.dragging_items: continue
            current_h = item.boundingRect().height() * item.transform().m11()
            item.setPos(x, y)
            x += self.item_width + self.spacing
            if x + self.item_width > view_width - self.margin:
                x = self.margin
                y += current_h + self.spacing
        total_height = y + 400
        self.scene.setSceneRect(0, 0, view_width, max(total_height, self.viewport().height()))

    def update_indicator(self, index):
        if index == -1: 
            self.indicator.hide()
            return
        view_width = self.viewport().width()
        x = self.margin
        y = self.margin
        for i in range(index):
            if i < len(self.items_list):
                target_item = self.items_list[i]
                if target_item in self.dragging_items: continue
                current_h = target_item.boundingRect().height() * target_item.transform().m11()
            else:
                current_h = self.item_width * 1.5
            x += self.item_width + self.spacing
            if x + self.item_width > view_width - self.margin:
                x = self.margin
                y += current_h + self.spacing
        self.indicator.height = self.item_width * 1.4
        self.indicator.prepareGeometryChange()
        indicator_x = x - (self.spacing / 2)
        self.indicator.setPos(indicator_x, y + 10)
        self.indicator.show()

    def mousePressEvent(self, event):
        item = self.itemAt(event.pos())
        if event.button() == Qt.MouseButton.RightButton:
            if isinstance(item, ThumbnailItem):
                item.setSelected(False)
            event.accept()
            return
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_start_pos = event.pos()
            self.is_dragging_mode = False
            if isinstance(item, ThumbnailItem):
                modifiers = QApplication.keyboardModifiers()
                if not (modifiers & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier)):
                    if not item.isSelected():
                        for sel in self.scene.selectedItems(): sel.setSelected(False)
                        item.setSelected(True)
                else:
                    item.setSelected(not item.isSelected())
            else:
                super().mousePressEvent(event)
                return
        
    def mouseMoveEvent(self, event):
        if not self.is_dragging_mode and event.buttons() & Qt.MouseButton.LeftButton:
            dist = (event.pos() - self.drag_start_pos).manhattanLength()
            if dist > QApplication.startDragDistance():
                item = self.itemAt(self.drag_start_pos)
                if isinstance(item, ThumbnailItem): self.start_drag_reorder(item)
        if self.is_dragging_mode:
            delta = self.mapToScene(event.pos()) - self.mapToScene(self.drag_start_pos)
            self.drag_start_pos = event.pos()
            for item in self.dragging_items: item.moveBy(delta.x(), delta.y())
            self.insert_index = self.calculate_index_at(event.pos())
            self.update_indicator(self.insert_index)
            view_y = event.pos().y()
            viewport_h = self.viewport().height()
            if view_y < 50: self.scroll_step = -20
            elif view_y > viewport_h - 50: self.scroll_step = 20
            else: self.scroll_step = 0
        else:
            super().mouseMoveEvent(event)

    def start_drag_reorder(self, trigger_item):
        self.is_dragging_mode = True
        self.dragging_items = [i for i in self.scene.selectedItems() if isinstance(i, ThumbnailItem)]
        if trigger_item not in self.dragging_items:
            trigger_item.setSelected(True)
            self.dragging_items.append(trigger_item)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(0, 0, 0, 80))
        for drag_item in self.dragging_items:
            drag_item.setZValue(100)
            drag_item.setOpacity(0.8)
            drag_item.setGraphicsEffect(shadow)
        self.scroll_timer.start(20)

    def mouseReleaseEvent(self, event):
        if self.is_dragging_mode:
            self.scroll_timer.stop()
            self.indicator.hide()
            for item in self.dragging_items:
                item.setZValue(0)
                item.setOpacity(1.0)
                item.setGraphicsEffect(None)
            self.dragging_items.sort(key=lambda x: self.items_list.index(x) if x in self.items_list else -1)
            for item in self.dragging_items:
                if item in self.items_list: self.items_list.remove(item)
            final_index = self.insert_index
            if final_index < 0: final_index = len(self.items_list)
            if final_index > len(self.items_list): final_index = len(self.items_list)
            self.items_list[final_index:final_index] = self.dragging_items
            self.dragging_items = []
            self.insert_index = -1
            self.is_dragging_mode = False
            self.relayout()
        else:
            super().mouseReleaseEvent(event)

    def auto_scroll(self):
        if self.scroll_step != 0:
            bar = self.verticalScrollBar()
            bar.setValue(bar.value() + self.scroll_step)

    def calculate_index_at(self, view_pos):
        scene_pos = self.mapToScene(view_pos)
        min_dist = float('inf')
        closest_idx = -1
        view_width = self.viewport().width()
        x = self.margin
        y = self.margin
        temp_list = [i for i in self.items_list if i not in self.dragging_items]
        for i in range(len(temp_list) + 1):
            est_h = self.item_width * 1.5
            center = QPointF(x + self.item_width/2, y + est_h/2)
            dist = abs(scene_pos.x() - center.x()) + abs(scene_pos.y() - center.y())
            if dist < min_dist:
                min_dist = dist
                closest_idx = i
            x += self.item_width + self.spacing
            if x + self.item_width > view_width - self.margin:
                x = self.margin
                y += est_h + self.spacing
        return closest_idx

# =========================================================
# 5. 封面裁剪弹窗
# =========================================================
class CropOverlayItem(QGraphicsObject):
    def __init__(self, rect_size, parent=None):
        super().__init__(parent)
        self.img_w, self.img_h = rect_size
        self.crop_ratio = 1.5 
        self.crop_rect = QRectF()
        self.update_crop_rect()
        
    def set_ratio(self, ratio):
        self.crop_ratio = ratio
        self.update_crop_rect()
        self.update()

    def update_crop_rect(self):
        tgt_h = self.img_w / self.crop_ratio
        tgt_w = self.img_w
        if tgt_h > self.img_h:
            tgt_h = self.img_h
            tgt_w = tgt_h * self.crop_ratio
        x = (self.img_w - tgt_w) / 2
        y = (self.img_h - tgt_h) / 2
        self.crop_rect = QRectF(x, y, tgt_w, tgt_h)

    def boundingRect(self):
        return QRectF(0, 0, self.img_w, self.img_h)

    def paint(self, painter, option, widget=None):
        path_full = QPainterPath()
        path_full.addRect(0, 0, self.img_w, self.img_h)
        path_crop = QPainterPath()
        path_crop.addRect(self.crop_rect)
        path_mask = path_full.subtracted(path_crop)
        
        painter.setBrush(QColor(0, 0, 0, 200))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawPath(path_mask)
        
        painter.setBrush(Qt.BrushStyle.NoBrush)
        pen = QPen(QColor("#00FFFF"), 2) 
        pen.setStyle(Qt.PenStyle.DashLine)
        painter.setPen(pen)
        painter.drawRect(self.crop_rect)

    def move_crop_rect(self, dy):
        new_y = self.crop_rect.y() + dy
        if new_y < 0: new_y = 0
        if new_y + self.crop_rect.height() > self.img_h:
            new_y = self.img_h - self.crop_rect.height()
        self.crop_rect.moveTop(new_y)
        self.update()

class CoverCropDialog(QDialog):
    def __init__(self, image_path, parent_dir, parent=None):
        super().__init__(parent)
        self.setWindowTitle("制作封面图")
        self.resize(1000, 800)
        self.image_path = image_path
        self.parent_dir = parent_dir 
        self.setup_ui()
        self.load_image()

    def setup_ui(self):
        layout = QHBoxLayout(self)
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setBackgroundBrush(QColor("#2d2d2d"))
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        layout.addWidget(self.view, 1) 
        
        ctrl_panel = QFrame()
        ctrl_panel.setFixedWidth(200)
        ctrl_panel.setStyleSheet("background-color: #333; border-radius: 10px;")
        vbox = QVBoxLayout(ctrl_panel)
        vbox.setSpacing(20)
        
        lbl_title = QLabel("横向比例")
        lbl_title.setStyleSheet("color: white; font-size: 16px; font-weight: bold;")
        vbox.addWidget(lbl_title)
        
        self.btn_r1 = QPushButton("1.5 : 1 (标准)")
        self.btn_r1.setCheckable(True)
        self.btn_r1.setChecked(True)
        self.btn_r1.clicked.connect(lambda: self.change_ratio(1.5))
        
        self.btn_r2 = QPushButton("2.8 : 1 (超宽)")
        self.btn_r2.setCheckable(True)
        self.btn_r2.clicked.connect(lambda: self.change_ratio(2.8))
        
        self.btn_group = [self.btn_r1, self.btn_r2]
        for btn in self.btn_group:
            btn.setStyleSheet("QPushButton{color: white; background: #444; border: none; padding: 10px;} QPushButton:checked{background: #007acc; border: 1px solid white;}")
            vbox.addWidget(btn)
        
        vbox.addStretch()
        lbl_info = QLabel("提示：\n直接上下拖动\n高亮区域即可调整")
        lbl_info.setStyleSheet("color: #aaa;")
        vbox.addWidget(lbl_info)
        
        btn_save = QPushButton("💾 裁剪并保存")
        btn_save.setStyleSheet("background-color: #28a745; color: white; padding: 12px; font-weight: bold;")
        btn_save.clicked.connect(self.save_cover)
        vbox.addWidget(btn_save)
        layout.addWidget(ctrl_panel)

    def load_image(self):
        try:
            self.pil_img = Image.open(self.image_path)
            if self.pil_img.mode != "RGBA": self.pil_img = self.pil_img.convert("RGBA")
            qim = QImage(self.pil_img.tobytes("raw", "RGBA"), self.pil_img.width, self.pil_img.height, QImage.Format.Format_RGBA8888)
            pixmap = QPixmap.fromImage(qim)
            self.bg_item = self.scene.addPixmap(pixmap)
            self.overlay = CropOverlayItem((pixmap.width(), pixmap.height()))
            self.overlay.setZValue(10)
            self.scene.addItem(self.overlay)
            self.view.fitInView(self.bg_item, Qt.AspectRatioMode.KeepAspectRatio)
            self.view.viewport().installEventFilter(self)
            self.last_pos = None
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法加载图片: {e}")
            self.close()

    def change_ratio(self, ratio):
        if ratio == 1.5:
            self.btn_r1.setChecked(True)
            self.btn_r2.setChecked(False)
        else:
            self.btn_r1.setChecked(False)
            self.btn_r2.setChecked(True)
        self.overlay.set_ratio(ratio)

    def eventFilter(self, source, event):
        if source == self.view.viewport():
            if event.type() == event.Type.MouseButtonPress:
                self.last_pos = event.pos()
                return True
            elif event.type() == event.Type.MouseMove and self.last_pos:
                delta_view = event.pos() - self.last_pos
                self.last_pos = event.pos()
                map_delta = self.view.mapToScene(delta_view.x(), delta_view.y()) - self.view.mapToScene(0, 0)
                self.overlay.move_crop_rect(map_delta.y())
                return True
            elif event.type() == event.Type.MouseButtonRelease:
                self.last_pos = None
                return True
        return super().eventFilter(source, event)

    def save_cover(self):
        try:
            rect = self.overlay.crop_rect
            box = (int(rect.x()), int(rect.y()), int(rect.right()), int(rect.bottom()))
            cropped = self.pil_img.crop(box)
            if cropped.mode == "RGBA": cropped = cropped.convert("RGB") 
            target_dir = os.path.dirname(self.parent_dir)
            if not target_dir or not os.path.exists(target_dir): target_dir = self.parent_dir 
            idx = 1
            while True:
                filename = f"封面图{idx}.jpg"
                save_path = os.path.join(target_dir, filename)
                if not os.path.exists(save_path): break
                idx += 1
            cropped.save(save_path, quality=95)
            QMessageBox.information(self, "成功", f"封面已保存至:\n{save_path}")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))

# =========================================================
# 6. 主窗口
# =========================================================
class FocusImageMain(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FocusImage Pro - Ultimate")
        self.resize(1400, 950)
        self.settings = QSettings("FocusImage", "Config")
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        self.scale_levels = {"XS (极小)": 100, "S (小)": 150, "M (中)": 200, "L (大)": 280, "XL (极大)": 350}
        self.current_scale_key = "M (中)"
        self.last_open_dir = ""  
        self.ps_path = "" 
        self.is_dark_mode = True 
        self.setup_ui()
        self.load_last_config()
        self.apply_theme()

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # --- 顶部工具栏 ---
        top_bar = QHBoxLayout()
        self.lbl_path = QLabel("未选择文件夹")
        self.lbl_path.setStyleSheet("font-size: 16px; font-weight: bold;")
        top_bar.addWidget(self.lbl_path)
        top_bar.addStretch()
        
        btn_sel_all = QPushButton("☑️ 全选")
        btn_sel_all.setFixedWidth(80)
        btn_sel_all.clicked.connect(self.select_all)
        top_bar.addWidget(btn_sel_all)

        btn_sel_inv = QPushButton("🔄 反选")
        btn_sel_inv.setFixedWidth(80)
        btn_sel_inv.clicked.connect(self.invert_selection)
        top_bar.addWidget(btn_sel_inv)

        self.btn_theme = QPushButton("☀️/🌙")
        self.btn_theme.setFixedWidth(60)
        self.btn_theme.clicked.connect(self.toggle_theme)
        top_bar.addWidget(self.btn_theme)
        
        top_bar.addWidget(QLabel("视图大小:"))
        self.combo_scale = QComboBox()
        self.combo_scale.addItems(list(self.scale_levels.keys()))
        self.combo_scale.setCurrentText(self.current_scale_key)
        self.combo_scale.currentTextChanged.connect(self.change_scale)
        self.combo_scale.setFixedWidth(120)
        top_bar.addWidget(self.combo_scale)
        layout.addLayout(top_bar)
        
        self.view = FlowLayoutView()
        self.view.item_width = self.scale_levels[self.current_scale_key]
        layout.addWidget(self.view)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        
        # --- 底部工具栏 ---
        btm_bar = QHBoxLayout()
        btm_bar.setSpacing(15)
        
        btn_open = QPushButton("📂 打开文件夹")
        btn_open.setFixedHeight(45)
        btn_open.clicked.connect(self.select_folder)
        btm_bar.addWidget(btn_open)
        
        btn_cover = QPushButton("🖼️ 制作封面")
        btn_cover.setFixedHeight(45)
        btn_cover.setStyleSheet("""
            QPushButton { background-color: #e67e22; border: 1px solid #d35400; color: white; }
            QPushButton:hover { background-color: #d35400; }
        """)
        btn_cover.clicked.connect(self.open_cover_maker)
        btm_bar.addWidget(btn_cover)
        
        btn_ps = QPushButton("🎨 PS 打开")
        btn_ps.setFixedHeight(45)
        btn_ps.setStyleSheet("""
            QPushButton { background-color: #001e36; border: 1px solid #31a8ff; color: #31a8ff; font-weight: bold; }
            QPushButton:hover { background-color: #002b4d; color: white; }
        """)
        btn_ps.clicked.connect(self.open_in_ps)
        btm_bar.addWidget(btn_ps)
        
        # === 新增：PS 保存并关闭按钮 ===
        btn_ps_close = QPushButton("💾 PS: 保存并关闭所有")
        btn_ps_close.setFixedHeight(45)
        # 红色警告样式
        btn_ps_close.setStyleSheet("""
            QPushButton { background-color: #3b0000; border: 1px solid #ff4d4d; color: #ff4d4d; font-weight: bold; }
            QPushButton:hover { background-color: #ff4d4d; color: white; }
        """)
        btn_ps_close.clicked.connect(self.save_and_close_ps_docs)
        btm_bar.addWidget(btn_ps_close)
        # ============================

        btm_bar.addStretch()
        
        lbl_start = QLabel("起始编号:")
        btm_bar.addWidget(lbl_start)
        self.spin_start = QComboBox()
        self.spin_start.setEditable(True)
        self.spin_start.setCurrentText("10")
        self.spin_start.setFixedWidth(80)
        self.spin_start.setFixedHeight(40)
        btm_bar.addWidget(self.spin_start)
        
        btn_jpg = QPushButton("🚀 导出 JPG")
        btn_jpg.setFixedHeight(45)
        btn_jpg.clicked.connect(self.export_jpg)
        btm_bar.addWidget(btn_jpg)
        
        btn_pdf = QPushButton("📄 生成 PDF")
        btn_pdf.setFixedHeight(45)
        btn_pdf.clicked.connect(self.export_pdf)
        btm_bar.addWidget(btn_pdf)
        layout.addLayout(btm_bar)

    def select_all(self):
        for item in self.view.items_list: item.setSelected(True)

    def invert_selection(self):
        for item in self.view.items_list: item.setSelected(not item.isSelected())

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()

    def apply_theme(self):
        if self.is_dark_mode: QApplication.instance().setStyleSheet(STYLE_DARK)
        else: QApplication.instance().setStyleSheet(STYLE_LIGHT)

    def change_scale(self, text):
        self.current_scale_key = text
        self.view.set_scale(self.scale_levels[text])

    def get_initial_dir(self):
        if self.last_open_dir and os.path.exists(self.last_open_dir): return os.path.dirname(self.last_open_dir)
        return ""

    def select_folder(self):
        start_dir = self.get_initial_dir()
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", start_dir)
        if folder: self.load_images(folder)

    def load_images(self, folder):
        self.last_open_dir = folder
        self.save_config(folder)
        self.lbl_path.setText(f"{os.path.basename(folder)}")
        self.lbl_path.setToolTip(folder)
        self.view.clear_items()
        exts = ('.jpg', '.png', '.jpeg', '.bmp', '.gif', '.webp')
        try:
            files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(exts)]
            files.sort(key=lambda s: [int(x) if x.isdigit() else x.lower() for x in re.split(r'(\d+)', os.path.basename(s))])
        except Exception: files = []
        if not files:
            QMessageBox.information(self, "提示", "没有找到图片")
            return
        self.progress_bar.show()
        width = self.scale_levels[self.current_scale_key]
        self.loader = ImageLoaderThread(files, width)
        self.loader.item_loaded_signal.connect(self.view.add_item_from_pil)
        self.loader.progress_signal.connect(lambda c, t: self.progress_bar.setValue(int(c/t*100)))
        self.loader.finished.connect(self.progress_bar.hide)
        self.loader.start()

    def open_cover_maker(self):
        selected_items = self.view.scene.selectedItems()
        img_items = [i for i in selected_items if isinstance(i, ThumbnailItem)]
        if len(img_items) != 1:
            QMessageBox.warning(self, "提示", "请先选中一张图片作为封面素材！\n(只能选中一张)")
            return
        target_item = img_items[0]
        if not os.path.exists(target_item.path):
            QMessageBox.warning(self, "错误", "图片文件不存在")
            return
        current_dir = self.last_open_dir
        dialog = CoverCropDialog(target_item.path, current_dir, self)
        dialog.exec()

    def open_in_ps(self):
        selected_items = self.view.scene.selectedItems()
        img_paths = [i.path for i in selected_items if isinstance(i, ThumbnailItem)]
        if not img_paths:
            QMessageBox.warning(self, "提示", "请先选择需要编辑的图片")
            return
        if not self.ps_path or not os.path.exists(self.ps_path):
            QMessageBox.information(self, "首次设置", "请找到您的 Photoshop.exe 主程序文件")
            exe_path, _ = QFileDialog.getOpenFileName(self, "选择 Photoshop.exe", "", "Executables (*.exe)")
            if exe_path and os.path.exists(exe_path):
                self.ps_path = exe_path
                self.save_config(self.last_open_dir) 
            else: return
        try:
            if len(img_paths) > 20:
                reply = QMessageBox.question(self, "警告", f"确定要一次性打开 {len(img_paths)} 张图片吗？\n这可能会导致电脑卡顿。", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.No: return
            cmd = [self.ps_path] + img_paths
            subprocess.Popen(cmd)
        except Exception as e:
            QMessageBox.critical(self, "启动失败", f"无法启动 Photoshop:\n{str(e)}")

# === 核心逻辑：使用COM控制PS保存并关闭 (修复JPEG弹窗问题) ===
    def save_and_close_ps_docs(self):
        if not HAS_WIN32:
            QMessageBox.critical(self, "错误", "缺少必要的库 'pywin32'。\n请在终端运行: pip install pywin32")
            return
            
        try:
            # 获取 PS 实例
            ps_app = win32com.client.GetActiveObject("Photoshop.Application")
        except Exception:
            try:
                ps_app = win32com.client.Dispatch("Photoshop.Application")
            except Exception:
                QMessageBox.warning(self, "提示", "未能连接到 Photoshop，请确认它是否正在运行。")
                return

        count = 0
        try:
            count = ps_app.Documents.Count
        except: pass
        
        if count == 0:
            QMessageBox.information(self, "提示", "Photoshop 中没有打开的文档。")
            return

        reply = QMessageBox.question(self, "确认", f"Photoshop 中有 {count} 个文档。\n确定要全部【保存并关闭】吗？\n(将自动应用最高质量保存 JPEG)", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.No:
            return

        # --- 关键设置：临时禁止 PS 弹窗 ---
        # 3 = psDisplayNoDialogs (不显示弹窗)
        original_dialog_mode = ps_app.DisplayDialogs
        ps_app.DisplayDialogs = 3 

        try:
            # 准备 JPEG 保存选项 (预先填好回答，防止弹窗)
            jpg_opts = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
            jpg_opts.EmbedColorProfile = True
            jpg_opts.FormatOptions = 1 # 1 = Standard
            jpg_opts.Matte = 1 # 1 = No Matte
            jpg_opts.Quality = 12 # 0-12, 12是最佳质量

            # 倒序遍历关闭，防止索引偏移
            while ps_app.Documents.Count > 0:
                doc = ps_app.ActiveDocument
                file_name = doc.Name.lower()
                
                try:
                    # 如果是 JPG/JPEG，使用 SaveAs 覆盖原文件以应用选项
                    if file_name.endswith('.jpg') or file_name.endswith('.jpeg'):
                        # 参数: 路径, 选项, 是否作为副本(False), 扩展名大小写(2=lower)
                        doc.SaveAs(doc.FullName, jpg_opts, False, 2)
                        # 因为上面已经SaveAs存过了，这里选 "2" (DoNotSaveChanges) 直接关闭
                        doc.Close(2) 
                    else:
                        # 对于 PSD、PNG 等其他格式，通常 Close(1) 就够了
                        # 配合 DisplayDialogs = 3，大部分也不会弹窗
                        doc.Close(1) # 1 = SaveChanges
                except Exception as e_doc:
                    print(f"关闭文档出错: {e_doc}")
                    # 如果保存失败，强制不保存关闭，防止卡死循环
                    doc.Close(2)

            QMessageBox.information(self, "完成", "所有文档已保存并关闭。")
            
        except Exception as e:
            QMessageBox.critical(self, "执行出错", f"操作过程中发生错误:\n{str(e)}")
        finally:
            # 恢复 PS 的弹窗设置，以免影响用户后续手动操作
            ps_app.DisplayDialogs = original_dialog_mode
    # ========================================
    # ========================================

    def export_jpg(self):
        paths = [item.path for item in self.view.items_list]
        if not paths: return
        start_dir = self.get_initial_dir()
        save_dir = QFileDialog.getExistingDirectory(self, "选择导出位置", start_dir)
        if not save_dir: return
        try: start = int(self.spin_start.currentText())
        except: start = 10
        count = 0
        try:
            for i, p in enumerate(paths):
                try:
                    with Image.open(p) as img:
                        if img.mode != 'RGB': img = img.convert('RGB')
                        name = f"{start + i*10:03d}.jpg"
                        img.save(os.path.join(save_dir, name), quality=95)
                    count += 1
                except: pass
            QMessageBox.information(self, "完成", f"已导出 {count} 张")
        except Exception as e: QMessageBox.critical(self, "错误", str(e))

    def export_pdf(self):
        paths = [item.path for item in self.view.items_list]
        if not paths: return
        start_dir = self.get_initial_dir()
        save_path, _ = QFileDialog.getSaveFileName(self, "保存PDF", os.path.join(start_dir, "output.pdf"), "*.pdf")
        if not save_path: return
        try:
            valid_imgs = []
            max_w = 0
            for p in paths:
                try:
                    with Image.open(p) as img:
                        max_w = max(max_w, img.width)
                        valid_imgs.append(p)
                except: pass
            if not valid_imgs: return
            margin = 20
            page_w = max_w + margin*2
            c = canvas.Canvas(save_path, pagesize=(page_w, 100))
            for p in valid_imgs:
                try:
                    with Image.open(p) as img:
                        w, h = img.size
                        draw_w = page_w - margin*2
                        scale = draw_w / w
                        draw_h = h * scale
                        c.setPageSize((page_w, draw_h + margin*2))
                        c.drawImage(PDFImageReader(img), margin, margin, width=draw_w, height=draw_h)
                        c.showPage()
                except: pass
            c.save()
            QMessageBox.information(self, "完成", "PDF已生成")
        except Exception as e: QMessageBox.critical(self, "错误", str(e))

    def load_last_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    last = data.get("last_dir", "")
                    self.ps_path = data.get("ps_path", "") 
                    if last and os.path.exists(last): self.load_images(last)
                    else: self.last_open_dir = last 
            except: pass

    def save_config(self, path):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({"last_dir": path, "ps_path": self.ps_path}, f)
        except: pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FocusImageMain()
    window.show()
    sys.exit(app.exec())