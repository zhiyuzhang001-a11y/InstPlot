import sys
import os
import re
import copy
import numpy as np
import pandas as pd
from license_manager_secure import check_license, activate_app, get_machine_code
from io import StringIO
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QWidget,
    QVBoxLayout, QHBoxLayout, QPushButton, QComboBox,
    QDialog, QTableWidget, QTableWidgetItem, QLabel, QToolBar,
    QMessageBox, QLineEdit
)
from PySide6.QtGui import QAction, QPixmap
from PySide6.QtCore import QSize, Qt
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# 最小化的启动时 rcParams 设置（仅必需项，加快启动）
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 延迟加载的样式初始化标志
_mpl_style_initialized = False

def _initialize_mpl_style():
    """延迟初始化 matplotlib 样式，仅在首次绘图时调用"""
    global _mpl_style_initialized
    if _mpl_style_initialized:
        return
    _mpl_style_initialized = True
    
    # 使用更现代的 matplotlib 风格
    try:
        plt.style.use('seaborn-v0_8-paper')
    except Exception:
        pass
    
    # 统一一些 rc 参数以获得更清晰的展示
    plt.rcParams.update({
        'figure.dpi': 120,
        'axes.titlesize': 16,
        'axes.labelsize': 15,
        'xtick.labelsize': 13,
        'ytick.labelsize': 13,
        'legend.fontsize': 12,
        'lines.linewidth': 2,
        'axes.linewidth': 1.2,
        'axes.edgecolor': '#000000',
        'xtick.color': '#000000',
        'ytick.color': '#000000',
        'text.color': '#000000',
        'lines.markersize': 5,
        'grid.color': "#DDDDDD86",
        'grid.linestyle': '--',
        'grid.alpha': 1,
    })

# 内嵌的浅色 QSS，作为缺省/回退样式（如果外部 style_light.qss 不存在或不可读）
STYLE_LIGHT_QSS = r"""
/* 非全白的“浅色”主题：整体为深灰但不是纯黑，绘图区仍由 Matplotlib 单独控制为白色 */
QWidget { background-color: #2b3036; color: #e6eef6; }
QToolBar { background: #32363c; spacing: 6px; border-bottom: 1px solid #3a3f45; }
QPushButton { background-color: #374151; border: 1px solid #424750; border-radius: 8px; padding: 8px 12px; color: #ffffff; }
QPushButton:hover { background-color: #3f4a56; }
QPushButton:pressed { background-color: #2f3a45; }
QComboBox { padding: 6px 10px; border: 1px solid #424750; border-radius: 8px; background-color: #374151; color: #ffffff; }
QComboBox::drop-down { width: 18px; border: none; background: transparent; }
/* 小箭头提示（浅色主题为深色箭头） */
/* 使用内嵌 SVG 作为小箭头，避免 CSS 三角在某些平台显示为矩形的问题 */
QComboBox::down-arrow {
    image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='8' height='6'><polygon points='0,0 8,0 4,6' fill='%23111111'/></svg>");
    width: 8px;
    height: 6px;
    margin-right: 6px;
}
QStatusBar { background: #2b3036; color: #d0d7de; }

/* 保持表格和头部配色 */
QTableWidget { background-color: #22262a; color: #e6eef6; gridline-color: #33383d; }
QHeaderView::section { background-color: #2b3036; color: #e6eef6; }

/* 下拉弹出列表使用浅色背景以便可读（即使整体界面为深灰） */
QComboBox QAbstractItemView, QListView, QMenu {
    background-color: #ffffff;
    color: #111111;
    selection-background-color: #e8eef7;
    outline: none;
}
QComboBox QAbstractItemView::item:selected, QListView::item:selected {
    background-color: #e8eef7;
    color: #111111;
}
"""

def latex_to_unicode(name):
    replacements = {
        r'\theta': '\u03B8',   # θ
        r'\mu': '\u03BC',      # μ
        r'\Omega': '\u03A9',   # Ω
        r'\alpha': '\u03B1',   # α
        r'\beta': '\u03B2',    # β
        r'\gamma': '\u03B3',   # γ
        r'\Delta': '\u0394',   # Δ
        r'\sigma': '\u03C3',   # σ
    }
    for k, v in replacements.items():
        name = name.replace(k, v)
    return name

class SquareFigureCanvas(FigureCanvas):
    """自定义 Canvas 类，保持绘图区域为正方形"""
    def __init__(self, figure):
        super().__init__(figure)
        from PySide6.QtWidgets import QSizePolicy
        # 设置尺寸策略，支持高度随宽度变化
        size_policy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        size_policy.setHeightForWidth(True)
        self.setSizePolicy(size_policy)
    
    def hasHeightForWidth(self):
        """告诉布局系统这个控件的高度依赖于宽度"""
        return True
    
    def heightForWidth(self, width):
        """返回给定宽度所对应的高度（正方形，所以返回相同值）"""
        return width

    def sizeHint(self):
        """向布局建议一个合适的默认大小（正方形）。"""
        try:
            from PySide6.QtCore import QSize
            return QSize(800, 800)
        except Exception:
            return super().sizeHint()

    def minimumSizeHint(self):
        """给出一个较小但可用的正方形最小尺寸。"""
        try:
            from PySide6.QtCore import QSize
            return QSize(500, 500)
        except Exception:
            return super().minimumSizeHint()

class PlotApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("InstPlot")
        
        # 根据屏幕可用区域自适应窗口尺寸（保持正方形，考虑高 DPI）
        screen = QApplication.primaryScreen()
        # 使用 availableGeometry() 获取去除任务栏后的可用区域
        available = screen.availableGeometry()
        
        # 获取 DPI 缩放比例
        dpi = screen.logicalDotsPerInch()
        dpi_scale = dpi / 96.0  # 96 是标准 DPI
        
        # 根据 DPI 调整窗口占据比例
        # 标准 DPI (100%): 占高度的 2/3
        # 高 DPI (150%): 占高度的 55%（2/3 / 1.2）
        # 更高 DPI: 进一步降低比例
        if dpi_scale <= 1.0:
            height_ratio = 0.67  # 2/3
        elif dpi_scale <= 1.25:
            height_ratio = 0.60  # 125% DPI
        elif dpi_scale <= 1.5:
            height_ratio = 0.55  # 150% DPI
        else:
            height_ratio = 0.50  # 200% 及以上
        
        window_size = int(available.height() * height_ratio)
        # 确保不小于最小尺寸，也不超过可用宽度
        # 根据 DPI 设置不同的最小窗口尺寸
        min_window_size = 700 if dpi_scale <= 1.25 else 500
        window_size = max(window_size, min_window_size)
        window_size = min(window_size, available.width() - 50)  # 留 50px 边距
        self.resize(window_size, window_size)
        
        # 设置最小窗口尺寸
        self.setMinimumSize(600, 500)
        
        # 将窗口居中显示在可用区域内
        self.move(
            available.x() + (available.width() - window_size) // 2,
            available.y() + (available.height() - window_size) // 2
        )
        
        self.setAcceptDrops(True)
        self.history = []  # 保存每次操作前的 loaded_files 状态，用于撤回
        self.max_history = 10  # 最多保存 10 步历史
        self.dragging = False
        self.last_mouse_pos = None
        # 矩形选择相关
        self._rect_selector = None
        self._rect_start = None  # (xdata, ydata)
        self._mouse_press_pix = None  # (xpix, ypix)
        self._is_selecting = False

        # 状态栏
        self.statusBar().showMessage("拖入数据文件或点击打开文件按钮")

        # 初始化字体大小（会在 resizeEvent 中动态更新）
        self._update_font_sizes()

        # =============== 工具栏 ===============
        self.toolbar = QToolBar("主工具栏", self)
        self.toolbar.setIconSize(QSize(25, 25))
        self.addToolBar(self.toolbar)

        # 应用全局样式（QSS）——轻量美化：圆角按钮、统一配色、悬停效果
        app = QApplication.instance()
        if app is not None:
            # 优先尝试读取外部样式表文件 style_light.qss（放在与本文件相同目录）
            qss_path = os.path.join(os.path.dirname(__file__), 'style_light.qss')
            try:
                if os.path.exists(qss_path):
                    with open(qss_path, 'r', encoding='utf-8') as f:
                        app.setStyleSheet(f.read())
                else:
                    # 如果外部文件不存在，使用内嵌的 STYLE_LIGHT_QSS 常量作为回退
                    app.setStyleSheet(STYLE_LIGHT_QSS)
            except Exception:
                # 最后退回到非常简单的内联样式，保证应用能显示
                try:
                    qss = """
                    QWidget { background-color: #fafafa; }
                    QToolBar { background: #ffffff; spacing: 6px; border-bottom: 1px solid #e6e6e6; }
                    QPushButton { background-color: #f0f3f7; border: 1px solid #d9e1ec; border-radius: 6px; padding: 6px 10px; }
                    QPushButton:hover { background-color: #e8eef7; }
                    QPushButton:pressed { background-color: #dfeaf9; }
                    QComboBox { padding: 4px 8px; border: 1px solid #d9dfe6; border-radius: 6px; }
                    QStatusBar { background: #ffffff; }
                    """
                    app.setStyleSheet(qss)
                except Exception:
                    pass

        def make_action(icon_name, text, slot):
            """快速创建带 FontAwesome 图标的 QAction"""
            act = QAction(text, self)
            try:
                # qtawesome 需要在 QApplication 创建后才能使用
                from qtawesome import icon as qta_icon
                act.setIcon(qta_icon(icon_name, color='#5f6368'))
            except Exception as e:
                # 如果图标加载失败，只使用文本
                print(f"图标加载失败 ({icon_name}): {e}")
            act.setStatusTip(text)
            act.triggered.connect(slot)
            return act

        
        # 图标参考：https://github.com/spyder-ide/qtawesome/blob/master/qtawesome/iconic-fonts.md
        # fa5s = FontAwesome 5 Solid, fa5r = FontAwesome 5 Regular
        self.toolbar.addAction(make_action("fa5s.folder-open", "打开文件", self.open_file))
        self.toolbar.addAction(make_action("fa5s.save", "导出数据", self.export_data))
        self.toolbar.addAction(make_action("fa5s.image", "保存图片", self.save_figure))
        self.toolbar.addAction(make_action("fa5s.undo", "撤回", self.undo))
        self.toolbar.addSeparator()
        # 主题（仅浅色），不提供深色切换

        # 画布和工具栏（使用轻量创建方式）
        from matplotlib.figure import Figure
        self.figure = Figure(figsize=(8, 8), dpi=100, facecolor='white')
        self.ax = self.figure.add_subplot(111, facecolor='white')
        self.canvas = SquareFigureCanvas(self.figure)
        self.canvas.mpl_connect('motion_notify_event', self.on_mouse_move)
        self.canvas.mpl_connect('button_press_event', self.on_mouse_press)
        # 鼠标交互绑定
        self.canvas.mpl_connect("button_release_event", self.on_mouse_release)
        self.canvas.mpl_connect("motion_notify_event", self.on_mouse_drag)
        self.canvas.mpl_connect("scroll_event", self.on_scroll)

        # =============== 下拉菜单和绘制按钮 ===============
        self.btn_center = QPushButton("对称处理")
        self.btn_normalize = QPushButton("归一化")
        self.btn_remove_bg = QPushButton("去背底")
        self.btn_clear = QPushButton("清空图形")
        # Matplotlib 核心导航按钮
        self.btn_save = QPushButton("保存图片")
        
        # 下拉选择（样式将在 _update_button_styles 中动态设置）
        self.combo_x = QComboBox()
        self.combo_y = QComboBox()
        
        # 防止首次加载时被右侧工具栏或其他控件遮挡，设置一个合理的最小/固定宽度
        try:
            # 使下拉宽度与“清空图形”按钮宽度一致，视觉更紧凑
            clear_w = self.btn_clear.sizeHint().width()
            if clear_w and clear_w > 0:
                w = int(clear_w)
            else:
                w = 110
            self.combo_x.setMinimumWidth(w)
            self.combo_y.setMinimumWidth(w)
            # 设置下拉列表视图的最小宽度
            self.combo_x.view().setMinimumWidth(200)
            self.combo_y.view().setMinimumWidth(200)
        except Exception:
            try:
                self.combo_x.setMinimumWidth(120)
                self.combo_y.setMinimumWidth(120)
            except Exception:
                pass
        # 不使用 X/Y 标签，保持简洁的下拉控件
        
        self.btn_plot = QPushButton("绘制曲线")

        # 顶部布局
        top_layout = QHBoxLayout()
        for w in [self.btn_center, self.btn_normalize,
                  self.btn_remove_bg]:
            top_layout.addWidget(w)

        top_layout.addStretch()
        for w in [self.btn_clear,self.combo_x, self.combo_y, self.btn_plot]:
            top_layout.addWidget(w)

        # =============== 主体布局 ===============
        layout = QVBoxLayout()
        layout.addLayout(top_layout)
        layout.addWidget(self.canvas)


        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # 初始化按钮样式（在所有控件创建后）
        self._update_button_styles()

        # =============== 按钮功能绑定 ===============
        self.btn_center.clicked.connect(self.apply_center)
        self.btn_normalize.clicked.connect(self.apply_normalize)
        self.btn_remove_bg.clicked.connect(self.remove_background)
        self.btn_clear.clicked.connect(self.clear_plot)
        self.btn_plot.clicked.connect(self.plot_selected)

        # 数据存储
        self.loaded_files = []  # 存储 (file_path, df)
        self.col_unicode_map = {}
        self.last_x_col = ""
        self.last_y_col = ""

        # 不加载深色主题偏好（深色主题支持已移除）
        # 固定宽度，不需要自适应调整

    def _update_font_sizes(self):
        """根据窗口大小动态计算字体大小"""
        try:
            # 基于窗口宽度计算字体大小
            window_width = self.width()
            # 基础字号：窗口宽度的 1.5%，范围 11-16pt
            self.base_font_size = max(11, min(16, int(window_width * 0.015)))
            # 按钮字号：比基础字号大 1-2pt，范围 12-18pt
            self.button_font_size = max(12, min(18, self.base_font_size + 2))
        except Exception:
            self.base_font_size = 12
            self.button_font_size = 14

    def _update_button_styles(self):
        """更新所有按钮和控件的样式"""
        try:
            # 计算 padding
            padding_v = max(int(self.button_font_size * 0.4), 4)
            padding_h = max(int(self.button_font_size * 0.7), 8)
            
            # 更新按钮样式
            self.top_button_style = (
                f"font-size: {self.button_font_size}pt; font-weight: 600; "
                f"padding: {padding_v}px {padding_h}px; "
                "border-radius: 8px; background-color: #f0f3f7; border: 1px solid #d9e1ec; color: #111111;"
            )
            
            # 应用到所有按钮
            for btn in [self.btn_center, self.btn_normalize, self.btn_remove_bg, 
                       self.btn_clear, self.btn_save, self.btn_plot]:
                btn.setStyleSheet(self.top_button_style)
            
            # 更新下拉框样式
            combo_padding_v = max(int(self.base_font_size * 0.4), 5)
            combo_padding_h = max(int(self.base_font_size * 0.8), 10)
            self.combo_style_light = (
                f"QComboBox {{ font-size: {self.base_font_size}pt; "
                f"padding: {combo_padding_v}px {combo_padding_h}px; border-radius: 8px; "
                "background-color: #f0f3f7; border: 1px solid #d9e1ec; color: #111111; }"
                "QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: top right; width: 30px; "
                "border: none; border-left: 1px solid #d9e1ec; border-top-right-radius: 8px; border-bottom-right-radius: 8px; "
                "background-color: transparent; }"
            )
            self.combo_x.setStyleSheet(self.combo_style_light)
            self.combo_y.setStyleSheet(self.combo_style_light)
        except Exception as e:
            pass

    def resizeEvent(self, event):
        """窗口大小改变时重新计算字体大小"""
        super().resizeEvent(event)
        try:
            self._update_font_sizes()
            self._update_button_styles()
        except Exception:
            pass

    def _calculate_scaled_font_size(self, base_size):
        """根据屏幕 DPI 计算缩放后的字体大小（单位：pt）"""
        try:
            screen = QApplication.primaryScreen()
            # 获取逻辑 DPI（通常是 96 on Windows, 72 on macOS）
            dpi = screen.logicalDotsPerInch()
            # 标准 DPI 是 96
            scale_factor = dpi / 96.0
            # 返回缩放后的字体大小，确保至少为基础大小
            return max(int(base_size * scale_factor), base_size)
        except Exception:
            return base_size

    # 鼠标移动时，显示坐标
    def on_mouse_move(self, event):
        if event.inaxes:  # 鼠标在绘图区内
            x, y = event.xdata, event.ydata
            self.statusBar().showMessage(f"x={x:.4g}, y={y:.4g}")
        else:
            self.statusBar().clearMessage()

    def on_click_point(self, event):
        if event.inaxes is None or event.button != 1:  # 只响应左键
            return

        x, y = event.xdata, event.ydata
        print(f"点击坐标: x={x:.3f}, y={y:.3f}")

        # 移除上一次高亮点
        if hasattr(self, "_highlight") and self._highlight is not None:
            try:
                self._highlight.remove()
            except Exception:
                pass
            self._highlight = None
        # 在所有已加载的曲线数据中寻找距离点击点最近的点
        nearest = None
        nearest_dist = None
        nearest_info = None  # (file_index, idx_in_df, xcol, ycol)
        xcol = self.combo_x.currentText()
        ycol = self.combo_y.currentText()
        if not xcol or not ycol:
            # 如果未选择列，直接显示点击高亮点
            self._highlight = event.inaxes.plot(
                x, y, marker='o', markersize=8, color='red', markeredgecolor='black', zorder=10
            )[0]
            self.canvas.draw()
            return

        # 使用像素坐标比较（更加符合可视上的点击定位），并设置最大像素容限
        tol_pixels = 10  # 容忍范围(px)，可调整
        event_xpix = getattr(event, 'x', None)
        event_ypix = getattr(event, 'y', None)
        if event_xpix is None or event_ypix is None:
            # 回退到数据坐标方式
            event_xpix = None
        for fi, (file_path, df) in enumerate(self.loaded_files):
            if xcol not in df.columns or ycol not in df.columns:
                continue
            xs = pd.to_numeric(df[xcol], errors='coerce')
            ys = pd.to_numeric(df[ycol], errors='coerce')
            valid_mask = ~(xs.isna() | ys.isna())
            if not valid_mask.any():
                continue
            xs_v = xs[valid_mask].to_numpy()
            ys_v = ys[valid_mask].to_numpy()
            try:
                if event_xpix is not None:
                    pts_disp = self.ax.transData.transform(np.column_stack((xs_v, ys_v)))
                    dx = pts_disp[:, 0] - event_xpix
                    dy = pts_disp[:, 1] - event_ypix
                    dists = np.hypot(dx, dy)
                    min_idx = int(np.argmin(dists))
                    min_dist_pix = float(dists[min_idx])
                    if min_dist_pix <= tol_pixels and (nearest_dist is None or min_dist_pix < nearest_dist):
                        nearest_dist = min_dist_pix
                        nearest = (xs_v[min_idx], ys_v[min_idx])
                        orig_indices = df[valid_mask].index.to_numpy()
                        nearest_info = (fi, int(orig_indices[min_idx]))
                else:
                    # 如果没有像素坐标，回退到数据坐标距离
                    dx = xs_v - x
                    dy = ys_v - y
                    dists = np.hypot(dx, dy)
                    min_idx = int(np.argmin(dists))
                    min_dist = float(dists[min_idx])
                    if nearest_dist is None or min_dist < nearest_dist:
                        nearest_dist = min_dist
                        nearest = (xs_v[min_idx], ys_v[min_idx])
                        orig_indices = df[valid_mask].index.to_numpy()
                        nearest_info = (fi, int(orig_indices[min_idx]))
            except Exception:
                continue

        if nearest is None:
            # 无数据点可选，直接绘制点击点
            self._highlight = event.inaxes.plot(
                x, y, marker='o', markersize=8, color='red', markeredgecolor='black', zorder=10
            )[0]
            self.canvas.draw()
            return

        # 高亮最近点
        hx, hy = nearest
        try:
            self._highlight = event.inaxes.plot(
                hx, hy, marker='o', markersize=10, color='red', markeredgecolor='black', zorder=12
            )[0]
        except Exception:
            self._highlight = None
        self.canvas.draw()

        # 弹出确认框
        try:
            mb = QMessageBox(self)
            mb.setWindowTitle("删除点")
            mb.setText(f"检测到最近点 (x={hx:.4g}, y={hy:.4g})，是否删除？")
            mb.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            ret = mb.exec()
            if ret == QMessageBox.Yes:
                # 记录历史以便撤回
                try:
                    self.history.append(copy.deepcopy(self.loaded_files))
                    if len(self.history) > self.max_history:
                        self.history.pop(0)
                except Exception:
                    pass
                # 从对应 DataFrame 删除该行并重绘
                try:
                    fi, orig_idx = nearest_info
                    file_path, df = self.loaded_files[fi]
                    # 删除行
                    df = df.drop(index=orig_idx).reset_index(drop=True)
                    # 更新存储
                    self.loaded_files[fi] = (file_path, df)
                    # 删除点后保留当前缩放/平移状态
                    self.replot_all(preserve_view=True)
                    self.statusBar().showMessage(f"已删除点 (x={hx:.4g}, y={hy:.4g})")
                except Exception as e:
                    print("删除点失败:", e)
                    self.statusBar().showMessage("删除点失败")
            else:
                # 如果取消，移除高亮
                try:
                    if hasattr(self, '_highlight') and self._highlight is not None:
                        self._highlight.remove()
                        self._highlight = None
                        self.canvas.draw()
                except Exception:
                    pass
        except Exception:
            pass

    # 鼠标按下
    def on_mouse_press(self, event):
        # 右键开始平移
        if event.button == 3 and event.inaxes:
            self.dragging = True
            self.last_mouse_pos = (event.x, event.y)
            return

        # 左键：可能是单击也可能是矩形选择，记录起点（像素与数据坐标）
        if event.button == 1 and event.inaxes:
            try:
                self._mouse_press_pix = (event.x, event.y)
                self._rect_start = (event.xdata, event.ydata)
                self._is_selecting = False
            except Exception:
                self._mouse_press_pix = None
                self._rect_start = None
                self._is_selecting = False

    # 鼠标释放
    def on_mouse_release(self, event):
        # 结束右键平移
        if event.button == 3:
            self.dragging = False
            self.last_mouse_pos = None
            return

        # 左键松开：处理矩形选择结束或单击
        if event.button == 1:
            # 如果处于矩形选择中，完成批量删除流程
            if self._is_selecting and self._rect_selector is not None:
                try:
                    bbox = self._rect_selector.get_bbox()
                    xmin, ymin = bbox.x0, bbox.y0
                    xmax, ymax = bbox.x1, bbox.y1
                except Exception:
                    xmin = ymin = xmax = ymax = None

                # 移除矩形补丁
                try:
                    self._rect_selector.remove()
                except Exception:
                    pass
                self._rect_selector = None
                self._is_selecting = False

                if xmin is None:
                    self._mouse_press_pix = None
                    self._rect_start = None
                    return

                xcol = self.combo_x.currentText()
                ycol = self.combo_y.currentText()
                if not xcol or not ycol:
                    self.statusBar().showMessage("请先选择 X/Y 列以进行区域删除")
                    self._mouse_press_pix = None
                    self._rect_start = None
                    return

                to_delete = {}
                total_count = 0
                for fi, (file_path, df) in enumerate(self.loaded_files):
                    if xcol not in df.columns or ycol not in df.columns:
                        continue
                    xs = pd.to_numeric(df[xcol], errors='coerce')
                    ys = pd.to_numeric(df[ycol], errors='coerce')
                    valid_mask = ~(xs.isna() | ys.isna())
                    if not valid_mask.any():
                        continue
                    sub = df[valid_mask]
                    mask_in = (sub[xcol] >= xmin) & (sub[xcol] <= xmax) & (sub[ycol] >= ymin) & (sub[ycol] <= ymax)
                    if mask_in.any():
                        inds = sub[mask_in].index.to_list()
                        to_delete[fi] = inds
                        total_count += len(inds)

                if total_count == 0:
                    self.statusBar().showMessage("矩形内未找到数据点")
                    self._mouse_press_pix = None
                    self._rect_start = None
                    try:
                        self.canvas.draw()
                    except Exception:
                        pass
                    return

                try:
                    mb = QMessageBox(self)
                    mb.setWindowTitle("删除多个点")
                    mb.setText(f"检测到 {total_count} 个点在选区内，是否删除？")
                    mb.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    ret = mb.exec()
                    if ret == QMessageBox.Yes:
                        try:
                            self.history.append(copy.deepcopy(self.loaded_files))
                            if len(self.history) > self.max_history:
                                self.history.pop(0)
                        except Exception:
                            pass
                        try:
                            for fi, inds in to_delete.items():
                                path, df = self.loaded_files[fi]
                                df = df.drop(index=inds).reset_index(drop=True)
                                self.loaded_files[fi] = (path, df)
                            # 批量删除后保留当前视图范围
                            self.replot_all(preserve_view=True)
                            self.statusBar().showMessage(f"已删除选区内 {total_count} 个点")
                        except Exception as e:
                            print("批量删除失败:", e)
                            self.statusBar().showMessage("批量删除失败")
                    else:
                        self.statusBar().showMessage("已取消批量删除")
                except Exception:
                    pass

                self._mouse_press_pix = None
                self._rect_start = None
                return

            # 否则按下与释放位置接近，视为单击
            try:
                if self._mouse_press_pix is not None:
                    px0, py0 = self._mouse_press_pix
                    if np.hypot(event.x - px0, event.y - py0) < 6:
                        self.on_click_point(event)
            except Exception:
                pass

            self._mouse_press_pix = None
            self._rect_start = None

    # 鼠标拖动
    def on_mouse_drag(self, event):
        # 右键平移优先
        if self.dragging:
            if not event.inaxes or self.last_mouse_pos is None:
                return
            dx = event.x - self.last_mouse_pos[0]
            dy = event.y - self.last_mouse_pos[1]
            ax = event.inaxes
            xlim = ax.get_xlim()
            ylim = ax.get_ylim()
            x_range = xlim[1] - xlim[0]
            y_range = ylim[1] - ylim[0]
            width, height = self.canvas.width(), self.canvas.height()
            dx_data = -dx * x_range / width
            dy_data = -dy * y_range / height
            ax.set_xlim(xlim[0] + dx_data, xlim[1] + dx_data)
            ax.set_ylim(ylim[0] + dy_data, ylim[1] + dy_data)
            self.canvas.draw()
            self.last_mouse_pos = (event.x, event.y)
            return

        # 左键矩形选择处理
        if self._mouse_press_pix is None or not event.inaxes:
            return
        try:
            px0, py0 = self._mouse_press_pix
            cur_px, cur_py = event.x, event.y
            dist = np.hypot(cur_px - px0, cur_py - py0)
            start_xdata, start_ydata = self._rect_start if self._rect_start is not None else (None, None)
            if dist > 6 and start_xdata is not None:
                x0, y0 = start_xdata, start_ydata
                x1, y1 = event.xdata, event.ydata
                if x1 is None or y1 is None:
                    return
                xmin, xmax = sorted([x0, x1])
                ymin, ymax = sorted([y0, y1])
                if not self._is_selecting:
                    try:
                        self._rect_selector = Rectangle((xmin, ymin), xmax - xmin, ymax - ymin,
                                                        fill=False, edgecolor='red', linewidth=1.2,
                                                        linestyle='--', zorder=11)
                        self.ax.add_patch(self._rect_selector)
                        self._is_selecting = True
                    except Exception:
                        self._rect_selector = None
                        self._is_selecting = False
                else:
                    try:
                        self._rect_selector.set_xy((xmin, ymin))
                        self._rect_selector.set_width(xmax - xmin)
                        self._rect_selector.set_height(ymax - ymin)
                    except Exception:
                        pass
                try:
                    self.canvas.draw()
                except Exception:
                    pass
        except Exception:
            pass

    def on_scroll(self, event):
        # 滚轮缩放
        if not event.inaxes:
            return
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        xdata, ydata = event.xdata, event.ydata

        scale_factor = 1.1 if event.button == 'down' else 1/1.1
        new_width = (xlim[1] - xlim[0]) * scale_factor
        new_height = (ylim[1] - ylim[0]) * scale_factor

        relx = (xdata - xlim[0]) / (xlim[1] - xlim[0])
        rely = (ydata - ylim[0]) / (ylim[1] - ylim[0])

        self.ax.set_xlim([xdata - new_width * relx, xdata + new_width * (1 - relx)])
        self.ax.set_ylim([ydata - new_height * rely, ydata + new_height * (1 - rely)])
        self.canvas.draw_idle()
    
    # 保存图片
    def save_figure(self):
        fname, _ = QFileDialog.getSaveFileName(
            self,
            "保存图片",
            "",
            "PNG 文件 (*.png);;JPG 文件 (*.jpg *.jpeg);;TIFF 文件 (*.tif *.tiff);;BMP 文件 (*.bmp);;PDF 文件 (*.pdf);;SVG 文件 (*.svg);;所有文件 (*)"
        )
        if fname:
            try:
                self.canvas.figure.savefig(fname, dpi=600)
                self.statusBar().showMessage(f"图片已保存: {fname}")
            except Exception as e:
                self.statusBar().showMessage(f"保存失败: {e}")

    
    # 打开文件
    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择数据文件", "", "Text Files (*.txt *.csv);;All Files (*)"
        )
        if file_path:
            self.load_file(file_path)
            self.plot_selected()

    # 拖拽事件
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            self.load_file(file_path)
        # 拖入后自动绘图
        self.plot_selected()
    
    # 加载文件
    def load_file(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        chosen_sep = None
        try:
            if ext in [".xls", ".xlsx"]:
                # 读取 Excel 文件
                df = pd.read_excel(file_path, header=0)  # 默认第一行作为列名
                chosen_sep = None  # Excel 不涉及分隔符
                enc_used = "Excel"
                
            else:
                # 读取文本文件
                def try_read_text_with_encodings(path, encodings):
                    for enc_try in encodings:
                        if not enc_try:
                            continue
                        try:
                            with open(path, 'r', encoding=enc_try) as f:
                                return f.read(), enc_try
                        except Exception:
                            continue
                    return None, None
        
                # 判断是否 VSM 文件
                with open(file_path, 'r', encoding='ascii', errors='ignore') as f:
                    preview_lines = [ln.strip() for ln in f.readlines()[:10] if ln.strip()]
                is_vsm = any("vsm" in ln.lower() for ln in preview_lines)

                if is_vsm:
                    # VSM 文件固定读取方式
                    df = pd.read_csv(
                        file_path,
                        skiprows=31,       # 跳过头信息
                        header=None,
                        usecols=[3, 4]     # Bz 和 emu
                    )
                    df.columns = ['B (Oe)', 'M (emu)']
                    chosen_sep = ','  # 方便状态栏显示
                    enc_used = "VSM"

                else:
                # 编码检测
                    import chardet
                    with open(file_path, 'rb') as f:
                        raw = f.read(5000)
                        detected = chardet.detect(raw)
                    enc_candidates = [detected.get('encoding'), 'utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'gbk', 'big5', 'mac_roman']

                    text, enc_used = try_read_text_with_encodings(file_path, enc_candidates)
                    if text is None:
                        raise ValueError("无法用常见编码读取文件")
                    
                    else:
                        # 分隔符检测
                        lines = [ln for ln in text.splitlines() if ln.strip()][:10]
                        if not lines:
                            raise ValueError("文件为空或只包含空行")
                        header_line = lines[0]
                        data_line = lines[1] if len(lines) > 1 else lines[0]
                        sep_candidates = ['\t', ',', ';', r'\s+']
                        chosen_sep = None
                        for sep in sep_candidates:
                            try:
                                if sep == r'\s+':
                                    hcols = re.split(r'\s+', header_line.strip())
                                    dcols = re.split(r'\s+', data_line.strip())
                                else:
                                    hcols = header_line.split(sep)
                                    dcols = data_line.split(sep)
                                if len(hcols) > 1 and abs(len(hcols) - len(dcols)) <= 0:
                                    chosen_sep = sep
                                    break
                            except Exception:
                                continue

                    if chosen_sep:
                        df = pd.read_csv(StringIO(text), sep=chosen_sep, engine='python')
                    else:
                        df = pd.read_fwf(StringIO(text))
                        chosen_sep = 'fwf'

            # 列名清理
            def clean_col_name(s):
                s = str(s).strip()
                s = re.sub(r'\s+', ' ', s)
                return s
            df.columns = [clean_col_name(c) for c in df.columns]

            # 修复常见乱码
            def fix_garbled(s: str) -> str:
                return (
                    s.replace('¦È', 'θ')
                    .replace('¡ã', '°')
                    .replace('¦¸', 'Ω')
                    .replace('Â', '')
                    .strip()
                )
            df.columns = [fix_garbled(c) for c in df.columns]

            # 更新数据存储
            self.loaded_files.append((file_path, df))
            self.col_unicode_map.update({col: latex_to_unicode(str(col)) for col in df.columns})

            # 更新下拉菜单（使用最新文件列名）
            self.combo_x.clear()
            self.combo_y.clear()
            self.combo_x.addItems(df.columns)
            self.combo_y.addItems(df.columns)

            # 记录默认列
            if not self.last_x_col:
                self.last_x_col = df.columns[0]
            if not self.last_y_col and len(df.columns) > 1:
                self.last_y_col = df.columns[1]
            self.combo_x.setCurrentText(self.last_x_col)
            self.combo_y.setCurrentText(self.last_y_col)

            # 状态栏
            self.statusBar().showMessage(f"已加载文件：{file_path} (编码: {enc_used}, 分隔符: {repr(chosen_sep)})")
            print(f"已加载文件: {file_path}, 编码: {enc_used}, 分隔符: {repr(chosen_sep)}")
            print("列名:", df.columns.tolist())
            print(df.head())

        except Exception as e:
            self.statusBar().showMessage(f"文件读取失败: {e}")
            print("读取文件错误:", e)

    # 绘图
    def plot_selected(self):
        if not self.loaded_files:
            self.statusBar().showMessage("请先加载数据文件")
            return

        x_col = self.combo_x.currentText()
        y_col = self.combo_y.currentText()
        if not x_col or not y_col:
            self.statusBar().showMessage("请选择 X 列和 Y 列")
            return
        
        self._draw_all_files(x_col, y_col)

        # 记住列名
        self.last_x_col = x_col
        self.last_y_col = y_col

    # 清空
    def clear_plot(self):
        self.ax.clear()
        self.canvas.draw()
        self.loaded_files.clear()
        # 清空下拉选择并重置记录的列
        try:
            self.combo_x.clear()
            self.combo_y.clear()
        except Exception:
            pass
        self.last_x_col = ""
        self.last_y_col = ""
        self.statusBar().showMessage("已清空图形、文件数据和 X/Y 选择")

    #对所有已加载文件的 Y 列执行纵向对称处理
    def apply_center(self):
        if not getattr(self, "loaded_files", None):
            self.statusBar().showMessage("请先加载数据文件")
            return

        y_col = self.combo_y.currentText()
        if not y_col:
            self.statusBar().showMessage("请选择 Y 列")
            return
        
        # 保存当前状态以便撤回
        self.history.append(copy.deepcopy(self.loaded_files))
        if len(self.history) > self.max_history:
            self.history.pop(0)

        changed = False
        for file_path, df in self.loaded_files:
            if y_col in df.columns:
                try:
                    df[y_col] = center_data(df[y_col])
                    changed = True
                    print(f"[center] applied to {os.path.basename(file_path)} ({y_col})")
                except Exception as e:
                    print(f"[center] failed on {file_path}: {e}")
            else:
                print(f"[center] skip {os.path.basename(file_path)}: no column {y_col}")

        if changed:
            self.statusBar().showMessage(f"对称处理完成（列: {y_col})")
            self.replot_all()
        else:
            self.statusBar().showMessage("没有文件包含所选 Y 列，未做处理")

    #对所有已加载文件的 Y 列执行归一化处理
    def apply_normalize(self):
        if not getattr(self, "loaded_files", None):
            self.statusBar().showMessage("请先加载数据文件")
            return

        y_col = self.combo_y.currentText()
        if not y_col:
            self.statusBar().showMessage("请选择 Y 列")
            return
        
        # 保存当前状态以便撤回
        self.history.append(copy.deepcopy(self.loaded_files))
        if len(self.history) > self.max_history:
            self.history.pop(0)

        changed = False
        for file_path, df in self.loaded_files:
            if y_col in df.columns:
                try:
                    # 先对称
                    df[y_col] = center_data(df[y_col])
                    # 再归一化
                    normalized_Y, top_n_avg = normalize_data(df[y_col])
                    df[y_col] = normalized_Y
                    changed = True
                    print(f"[normalize] applied to {os.path.basename(file_path)} ({y_col}), top_n_avg={top_n_avg}")
                except Exception as e:
                    print(f"[normalize] failed on {file_path}: {e}")
            else:
                print(f"[normalize] skip {os.path.basename(file_path)}: no column {y_col}")

        if changed:
            self.statusBar().showMessage(f"归一化完成（列: {y_col})")
            self.replot_all()
        else:
            self.statusBar().showMessage("没有文件包含所选 Y 列，未做处理")

    #对已加载的所有文件进行背景信号去除（多项式拟合）
    def remove_background(self):
        if not self.loaded_files:
            self.statusBar().showMessage("请先加载数据文件")
            return

        x_col = self.combo_x.currentText()
        y_col = self.combo_y.currentText()

        if not x_col or not y_col:
            self.statusBar().showMessage("请先选择 X/Y 列")
            return
        
        # 保存当前状态以便撤回
        self.history.append(copy.deepcopy(self.loaded_files))
        if len(self.history) > self.max_history:
            self.history.pop(0)

        # 弹出表格让用户输入每条曲线的区间
        dlg = QDialog(self)
        dlg.setWindowTitle("设置去线性背景拟合区间")
        dlg.resize(400, 400)
        layout = QVBoxLayout(dlg)

        label = QLabel("单位请与数据列一致，留空表示跳过该曲线")
        layout.addWidget(label)

        table = QTableWidget(dlg)
        table.setRowCount(len(self.loaded_files))
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["文件名", "x_min", "x_max"])

        for i, (path, df) in enumerate(self.loaded_files):
            table.setItem(i, 0, QTableWidgetItem(os.path.basename(path)))
            table.setItem(i, 1, QTableWidgetItem(""))  # 默认空
            table.setItem(i, 2, QTableWidgetItem(""))

        layout.addWidget(table)

        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_cancel = QPushButton("取消")
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

        dlg.setLayout(layout)

        def on_ok():
            for i, (path, df) in enumerate(self.loaded_files):
                item_min = table.item(i, 1)
                item_max = table.item(i, 2)
                try:
                    B_min = float(item_min.text())
                    B_max = float(item_max.text())
                    if x_col in df.columns and y_col in df.columns:
                        mask = (df[x_col] >= B_min) & (df[x_col] <= B_max)
                        if mask.any():
                            p = np.polyfit(df.loc[mask, x_col], df.loc[mask, y_col], 1)
                            df[y_col] = df[y_col] - np.polyval(p, df[x_col])
                            print(f"[background] 去线性基底: {os.path.basename(path)} ({y_col})")
                except Exception:
                    # 空或者无效输入则跳过
                    continue
            dlg.accept()
            self.statusBar().showMessage("去背景处理完成")
            self.replot_all()

        btn_ok.clicked.connect(on_ok)
        btn_cancel.clicked.connect(dlg.reject)
        dlg.exec()

    #重新绘制所有当前曲线（使用当前 combo 中的列
    def replot_all(self, preserve_view=False):
        if not getattr(self, "loaded_files", None):
            self.statusBar().showMessage("尚未加载任何文件", 3000)
            return

        x_col = self.combo_x.currentText()
        y_col = self.combo_y.currentText()
        if not x_col or not y_col:
            self.statusBar().showMessage("请先选择 X/Y 列以重新绘制", 3000)
            return
        
        # 如果需要保留当前视图（缩放/平移），在绘制前保存坐标范围并在绘制后恢复
        cur_xlim = None
        cur_ylim = None
        if preserve_view:
            try:
                cur_xlim = self.ax.get_xlim()
                cur_ylim = self.ax.get_ylim()
            except Exception:
                cur_xlim = None
                cur_ylim = None

        try:
            self._draw_all_files(x_col, y_col)
            # 在成功重绘之后，恢复视图范围（如果之前保存了）
            if preserve_view and (cur_xlim is not None) and (cur_ylim is not None):
                try:
                    self.ax.set_xlim(cur_xlim)
                    self.ax.set_ylim(cur_ylim)
                    # 触发一次绘制以确保视图更新
                    self.canvas.draw()
                except Exception:
                    pass

            self.statusBar().showMessage(f"已重新绘制所有文件(X={x_col}, Y={y_col})", 3000)
        except Exception as e:
            self.statusBar().showMessage(f"绘图时出错：{e}", 5000)
    
    #核心绘图函数：根据当前 loaded_files 绘制曲线并统一样式
    def _draw_all_files(self, x_col, y_col):
        # 延迟初始化 matplotlib 样式（仅首次绘图时执行）
        _initialize_mpl_style()
        
        self.ax.clear()
        for file_path, df in self.loaded_files:
            if x_col not in df.columns or y_col not in df.columns:
                continue
            df[x_col] = pd.to_numeric(df[x_col], errors='coerce')
            df[y_col] = pd.to_numeric(df[y_col], errors='coerce')
            label_name = os.path.splitext(os.path.basename(file_path))[0]
            # 使用更柔和的点样式，避免每条曲线都显示 marker（影响大量点的渲染）
            # 绘制曲线并增加小 marker（在点多时不会过于拥挤）
            self.ax.plot(df[x_col], df[y_col], label=label_name, linewidth=2,
                         marker='o', markersize=4, markeredgewidth=0.6, alpha=0.9)

        # 局部样式设置（避免修改全局 rc）
        self.ax.set_xlabel(f"{x_col}", fontsize=16, labelpad=8)
        self.ax.set_ylabel(f"{y_col}", fontsize=16, labelpad=8)
        # 主/次刻度样式：在四周都显示（top/right），刻度向内，适度加粗
        self.ax.tick_params(axis='both', which='major', labelsize=13, length=6, width=1.2,
                            direction='in', top=True, right=True)
        self.ax.tick_params(axis='both', which='minor', labelsize=11, length=4, width=1.2,
                            direction='in', top=True, right=True)
        # 确保刻度位置设置为 both 以在上下左右显示刻度线
        try:
            self.ax.xaxis.set_ticks_position('both')
            self.ax.yaxis.set_ticks_position('both')
        except Exception:
            pass
        # 将图例放在图外右侧，避免遮挡数据（如果图中曲线较少则自然放置）
        # 将图例放回绘图区内部，使用自动最佳位置
        try:
            leg = self.ax.legend(fontsize=12, loc='best')
        except Exception:
            leg = self.ax.legend(fontsize=12)
        # 图例使用默认配色（主题仅为浅色），若需微调可在 style_light.qss 中修改
        self.ax.grid(True, linestyle='--', alpha=0.6)
        # 让布局适应右侧图例
        self.figure.tight_layout(rect=[0, 0, 0.92, 1])
        self.canvas.draw()

        # 状态栏信息
        x_unicode = self.col_unicode_map.get(x_col, x_col)
        y_unicode = self.col_unicode_map.get(y_col, y_col)
        
        self.statusBar().showMessage(f"绘制完成: {y_unicode} vs {x_unicode}")
    #撤回上一步操作
    def undo(self):
        if self.history:
            self.loaded_files = self.history.pop()
            self.replot_all()
            self.statusBar().showMessage("已撤回上一步操作")
        else:
            self.statusBar().showMessage("没有可撤回的操作")

    #导出当前加载的数据到文件（Excel/CSV/TXT）
    def export_data(self):
        if not self.loaded_files:
            self.statusBar().showMessage("没有数据可以导出")
            return

        # 打开保存文件对话框
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "导出数据", 
            "", 
            "Excel Files (*.xlsx);;CSV Files (*.csv);;Text Files (*.txt)", 
            options=options
        )
        if not file_path:
            return

        ext = os.path.splitext(file_path)[1].lower()

        try:
            if ext == ".xlsx":
                # 多 sheet 导出
                try:
                    import openpyxl
                except ImportError:
                    self.statusBar().showMessage("导出 Excel 需要安装 openpyxl")
                    return

                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    for i, (path, df) in enumerate(self.loaded_files):
                        # sheet 名称不能太长，且不能重复
                        sheet_name = f"{i}_{os.path.basename(path)[:20]}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

            else:
                # CSV 或 TXT，合并到一个文件
                sep = ',' if ext == '.csv' else '\t'
                with open(file_path, 'w', encoding='utf-8') as f:
                    for path, df in self.loaded_files:
                        f.write(f"# 文件: {os.path.basename(path)}\n")
                        df.to_csv(f, sep=sep, index=False)
                        f.write("\n\n")

            self.statusBar().showMessage(f"数据导出成功: {file_path}")
        except Exception as e:
            self.statusBar().showMessage(f"导出数据失败: {e}")
            print("导出数据错误:", e)

    
    def apply_light_theme(self):
        app = QApplication.instance()
        if app is None:
            return
        # 恢复轻主题（清空样式表或重置为默认）
        try:
            # 尝试从文件加载浅色样式表
            qss_path = os.path.join(os.path.dirname(__file__), 'style_light.qss')
            if os.path.exists(qss_path):
                with open(qss_path, 'r', encoding='utf-8') as f:
                    app.setStyleSheet(f.read())
            else:
                app.setStyleSheet('')
        except Exception:
            pass
        try:
            plt.style.use('seaborn-v0_8-paper')
        except Exception:
            pass
        # 绘图区设为白色
        self.figure.patch.set_facecolor('white')
        self.ax.set_facecolor('white')
        # 提高浅色主题下网格和坐标轴的对比度，避免网格太淡
        try:
            # 设置轴线颜色（深灰），刻度颜色，以及网格颜色
            for spine in self.ax.spines.values():
                spine.set_color('#000000')
                spine.set_linewidth(1.6)
            self.ax.xaxis.label.set_color('#222222')
            self.ax.yaxis.label.set_color('#222222')
            # 在四周显示刻度，刻度向内，颜色为深色
            self.ax.tick_params(axis='both', which='major', colors='#000000', direction='in', top=True, right=True, width=1.6)
            self.ax.tick_params(axis='both', which='minor', colors='#000000', direction='in', top=True, right=True, width=1.2)
            self.ax.grid(True, color='#cccccc', linestyle='--', alpha=0.9)
            # 更新已存在的图例边框颜色（如果存在）
            legend = self.ax.get_legend()
            if legend is not None:
                legend.get_frame().set_edgecolor('#000000')
                legend.get_frame().set_linewidth(0.8)
        except Exception:
            pass
        # 恢复顶部按钮浅色样式
        try:
            for btn in [self.btn_center, self.btn_normalize, self.btn_remove_bg, self.btn_clear, self.btn_save, self.btn_plot]:
                btn.setStyleSheet(self.top_button_style)
        except Exception:
            pass
        # 恢复下拉为浅色
        try:
            self.combo_x.setStyleSheet(self.combo_style_light)
            self.combo_y.setStyleSheet(self.combo_style_light)
        except Exception:
            pass
    # 深色主题支持已移除
        # 强制重绘画布以立即反映浅色背景
        try:
            self.canvas.draw()
        except Exception:
            pass
        self.replot_all()

#纵坐标对称以及归一化数据
def center_data(Y):
    Y = np.asarray(Y)
    center_value = (np.nanmax(Y) + np.nanmin(Y)) / 2
    centered_Y = Y - center_value
    return centered_Y

def normalize_data(Y, top_n=20):
    Y = pd.to_numeric(Y, errors='coerce')
    Y = np.asarray(Y)

    valid_Y = Y[~np.isnan(Y)]

    if len(valid_Y) == 0:
        return Y, np.nan
    if len(valid_Y) < top_n:
        top_n = len(valid_Y)

    top_n_avg = np.nanmean(np.partition(valid_Y, -top_n)[-top_n:])
    if top_n_avg == 0:
        return Y, top_n_avg

    normalized_Y = np.where(Y > top_n_avg, 
                            1, 
                            np.where(Y < -top_n_avg, 
                                     -1, 
                                     Y / top_n_avg
                                    )
                           )
    return normalized_Y, top_n_avg

if __name__ == "__main__":
    # 启用高 DPI 支持（必须在创建 QApplication 之前设置）
    # 启用高 DPI 像素图
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    # 兼容在嵌入/交互环境中已存在 QApplication 的情况
    app = QApplication.instance() or QApplication(sys.argv)
    
    # License 验证：检查许可证或试用状态
    is_valid, remaining_days = check_license()
    
    if not is_valid:
        # 试用已过期，弹出激活窗口
        machine_code = get_machine_code()
        
        # 创建自定义对话框，让机器码可以选中和复制
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QHBoxLayout
        
        dialog = QDialog(None)
        dialog.setWindowTitle("试用已结束")
        dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        # 提示文本
        label1 = QLabel("您的 30 天试用期已结束。")
        label1.setStyleSheet("font-size: 12pt; margin-bottom: 10px;")
        layout.addWidget(label1)
        
        # 机器码标签
        label2 = QLabel("您的机器码：")
        label2.setStyleSheet("font-size: 11pt; margin-top: 10px;")
        layout.addWidget(label2)
        
        # 机器码输入框（只读，但可选中复制）
        machine_code_edit = QLineEdit(machine_code)
        machine_code_edit.setReadOnly(True)
        machine_code_edit.setStyleSheet(
            "font-size: 14pt; font-weight: bold; padding: 8px; "
            "background-color: #f0f0f0; border: 2px solid #4CAF50; "
            "border-radius: 5px; color: #2E7D32;"
        )
        machine_code_edit.selectAll()  # 默认全选，方便复制
        layout.addWidget(machine_code_edit)
        
        # 说明文本
        label3 = QLabel("请将上述机器码复制并发送给作者，获取授权码后点击下方按钮激活。")
        label3.setWordWrap(True)
        label3.setStyleSheet("font-size: 10pt; margin-top: 10px; color: #666;")
        layout.addWidget(label3)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_purchase = QPushButton("💰 立即购买")
        btn_purchase.setStyleSheet(
            "font-size: 11pt; padding: 8px 20px; background-color: #FF9800; "
            "color: white; border-radius: 5px; font-weight: bold;"
        )
        
        btn_activate = QPushButton("输入授权码激活")
        btn_activate.setStyleSheet(
            "font-size: 11pt; padding: 8px 20px; background-color: #4CAF50; "
            "color: white; border-radius: 5px; font-weight: bold;"
        )
        
        btn_exit = QPushButton("退出")
        btn_exit.setStyleSheet(
            "font-size: 11pt; padding: 8px 20px; background-color: #f44336; "
            "color: white; border-radius: 5px;"
        )
        
        button_layout.addWidget(btn_purchase)
        button_layout.addWidget(btn_activate)
        button_layout.addWidget(btn_exit)
        layout.addLayout(button_layout)
        
        dialog.setLayout(layout)
        
        # 定义购买对话框函数
        def show_purchase_dialog():
            """显示购买引导对话框"""
            purchase_dialog = QDialog(dialog)
            purchase_dialog.setWindowTitle("购买 InstPlot")
            purchase_dialog.setMinimumWidth(480)
            
            p_layout = QVBoxLayout()
            p_layout.setSpacing(8)  # 设置紧凑的间距
            
            # 标题
            title_label = QLabel("🎉 购买 InstPlot 完整版")
            title_label.setStyleSheet("font-size: 15pt; font-weight: bold; color: #2196F3; margin-bottom: 5px;")
            title_label.setAlignment(Qt.AlignCenter)
            p_layout.addWidget(title_label)
            
            # 价格
            price_label = QLabel("永久使用 - 仅需 ¥6.99")
            price_label.setStyleSheet("font-size: 13pt; color: #FF5722; font-weight: bold; margin-bottom: 8px;")
            price_label.setAlignment(Qt.AlignCenter)
            p_layout.addWidget(price_label)
            
            # 分隔线
            line1 = QLabel()
            line1.setStyleSheet("border-bottom: 2px solid #E0E0E0; margin: 5px 0;")
            p_layout.addWidget(line1)
            
            # 步骤1标题 - 支付
            step1_label = QLabel("【步骤 1】微信扫码支付， 并备注邮箱账号")
            step1_label.setStyleSheet(
                "font-size: 11pt; font-weight: bold; margin-top: 8px; "
                "color: white; background-color: #07C160; padding: 6px; "
                "border-radius: 5px;"
            )
            step1_label.setAlignment(Qt.AlignCenter)
            p_layout.addWidget(step1_label)
            
            # 收款码图片
            qr_label = QLabel()
            qr_label.setAlignment(Qt.AlignCenter)
            qr_label.setStyleSheet("margin: 5px 0; padding: 5px;")
            
            # 尝试加载收款码图片（支持多种格式）
            qr_image_found = False
            for qr_image_path in ["payment_qrcode.jpg", "payment_qrcode.png", "payment_qrcode.jpeg"]:
                if os.path.exists(qr_image_path):
                    qr_pixmap = QPixmap(qr_image_path)
                    # 缩放图片（保持宽高比）
                    scaled_pixmap = qr_pixmap.scaled(220, 220, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    qr_label.setPixmap(scaled_pixmap)
                    qr_image_found = True
                    break
            
            if not qr_image_found:
                # 如果图片不存在，显示提示
                qr_label.setText(
                    "⚠️ 收款码图片未找到\n\n"
                    "请将微信收款码保存为:\n"
                    "payment_qrcode.jpg\n"
                    "并放在程序目录下"
                )
                qr_label.setStyleSheet(
                    "font-size: 11pt; color: #FF5722; padding: 30px; "
                    "background-color: #FFF3E0; border: 2px dashed #FF9800; border-radius: 10px;"
                )
            p_layout.addWidget(qr_label)
            
            # 分隔线
            line2 = QLabel()
            line2.setStyleSheet("border-bottom: 2px solid #E0E0E0; margin: 10px 0;")
            p_layout.addWidget(line2)
            
            # 步骤2标题 - 发送机器码
            step2_label = QLabel("【步骤 2】发送机器码给我")
            step2_label.setStyleSheet(
                "font-size: 11pt; font-weight: bold; "
                "color: white; background-color: #FF9800; padding: 6px; "
                "border-radius: 5px;"
            )
            step2_label.setAlignment(Qt.AlignCenter)
            p_layout.addWidget(step2_label)
            
            # 提示文字（包含可复制的邮箱）
            hint_text = QLabel("可以手动复制机器码发送到以下邮箱，或点击按钮自动发送")
            hint_text.setStyleSheet(
                "font-size: 11pt; color: #2196F3; font-weight: bold;"
            )
            hint_text.setAlignment(Qt.AlignCenter)
            hint_text.setContentsMargins(0, 8, 0, 5)
            p_layout.addWidget(hint_text)
            
            # 邮箱地址（可选中复制）
            email_display = QLineEdit("862937649@qq.com")
            email_display.setReadOnly(True)
            email_display.setAlignment(Qt.AlignCenter)
            email_display.setStyleSheet(
                "font-size: 12pt; font-weight: bold; padding: 8px; "
                "background-color: #E3F2FD; border: 2px solid #2196F3; "
                "border-radius: 5px; color: #1565C0;"
            )
            email_display.setMinimumWidth(200)
            
            # 创建一个水平布局来居中邮箱输入框
            email_layout = QHBoxLayout()
            email_layout.addStretch()
            email_layout.addWidget(email_display)
            email_layout.addStretch()
            
            p_layout.addLayout(email_layout)
            
            # 机器码输入框
            machine_edit = QLineEdit(machine_code)
            machine_edit.setReadOnly(True)
            machine_edit.setStyleSheet(
                "font-size: 13pt; font-weight: bold; padding: 10px; "
                "background-color: #FFF9C4; border: 2px solid #FFA000; "
                "border-radius: 5px; color: #E65100; margin-top: 5px;"
            )
            machine_edit.selectAll()
            p_layout.addWidget(machine_edit)
            
            # 按钮容器（复制和发邮件按钮并排）
            btn_container = QHBoxLayout()
            btn_container.setSpacing(10)
            
            # 复制按钮
            copy_btn = QPushButton("📋 复制机器码")
            copy_btn.setStyleSheet(
                "font-size: 10pt; padding: 8px; background-color: #2196F3; "
                "color: white; border-radius: 5px; margin-top: 5px; font-weight: bold;"
            )
            def copy_machine_code():
                clipboard = QApplication.clipboard()
                clipboard.setText(machine_code)
                copy_btn.setText("✅ 已复制")
                copy_btn.setStyleSheet(
                    "font-size: 10pt; padding: 8px; background-color: #4CAF50; "
                    "color: white; border-radius: 5px; margin-top: 5px; font-weight: bold;"
                )
            copy_btn.clicked.connect(copy_machine_code)
            btn_container.addWidget(copy_btn)
            
            # 发送邮件按钮
            from PySide6.QtGui import QDesktopServices
            from PySide6.QtCore import QUrl
            import urllib.parse
            
            email_btn = QPushButton("📧 自动发送邮件")
            email_btn.setStyleSheet(
                "font-size: 10pt; padding: 8px; background-color: #9C27B0; "
                "color: white; border-radius: 5px; font-weight: bold; margin-top: 5px;"
            )
            
            def send_email():
                # 构建邮件链接
                to_email = "862937649@qq.com"
                subject = "InstPlot 软件购买 - 机器码"
                body = f"您好！\n\n我已完成支付，请为我生成授权码。\n\n我的机器码是：{machine_code}\n\n谢谢！"
                
                # URL 编码
                mailto_url = f"mailto:{to_email}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
                
                # 打开邮件客户端
                QDesktopServices.openUrl(QUrl(mailto_url))
                
                # 按钮反馈
                email_btn.setText("✅ 邮件已打开")
                email_btn.setStyleSheet(
                    "font-size: 10pt; padding: 8px; background-color: #4CAF50; "
                    "color: white; border-radius: 5px; font-weight: bold; margin-top: 5px;"
                )
            
            email_btn.clicked.connect(send_email)
            btn_container.addWidget(email_btn)
            
            p_layout.addLayout(btn_container)
            
            # 提示文本
            tip_label = QLabel("支付后会在 5 分钟内回复授权码到您的邮箱")
            tip_label.setStyleSheet(
                "font-size: 10pt; color: #666; margin-top: 10px;"
            )
            tip_label.setAlignment(Qt.AlignCenter)
            p_layout.addWidget(tip_label)
            
            # 底部按钮
            p_button_layout = QHBoxLayout()
            p_button_layout.addStretch()
            
            ok_btn = QPushButton("我知道了")
            ok_btn.setStyleSheet(
                "font-size: 11pt; padding: 8px 30px; background-color: #4CAF50; "
                "color: white; border-radius: 5px;"
            )
            ok_btn.clicked.connect(purchase_dialog.accept)
            p_button_layout.addWidget(ok_btn)
            
            p_layout.addLayout(p_button_layout)
            purchase_dialog.setLayout(p_layout)
            purchase_dialog.exec()
        
        # 按钮事件
        btn_purchase.clicked.connect(show_purchase_dialog)
        btn_activate.clicked.connect(dialog.accept)
        btn_exit.clicked.connect(dialog.reject)
        
        ret = dialog.exec()
        if ret == QDialog.Accepted:
            # 用户选择激活
            if activate_app():
                # 激活成功，继续运行
                pass
            else:
                # 激活失败或取消，退出
                sys.exit(0)
        else:
            # 用户选择退出
            sys.exit(0)
    else:
        # 许可证有效或在试用期内
        if remaining_days > 0:
            # 试用模式，显示剩余天数
            QMessageBox.information(None, "试用模式", f"您正在试用模式下使用本软件。\n剩余试用天数：{remaining_days} 天")
        # remaining_days == 0 表示已激活，不显示任何提示
    
    # 创建主窗口（窗口尺寸和位置已在 __init__ 中自动设置）
    window = PlotApp()
    
    window.show()
    sys.exit(app.exec())