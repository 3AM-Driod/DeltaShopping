import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import Canvas, Label, Frame, Entry, messagebox
import threading
import time
from mss import mss
import sys
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController, Listener as KeyboardListener
import os
import json
from pathlib import Path
import ctypes
import platform
from datetime import datetime, timedelta


def resource_path(relative_path):
    """获取资源文件的绝对路径。在开发时和打包后均可使用"""
    try:
        # 如果存在_MEIPASS属性，说明程序是打包后运行
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent

    return str(base_path / relative_path)


# ====== Tesseract配置 ======
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# 配置文件路径
CONFIG_FILE = "ocr_config.json"


def load_config():
    """加载配置文件，返回阈值设置和区域配置"""
    config = {
        'threshold1': "愣着干嘛？",
        'threshold2': "等我给填？",
        'max_attempts': 1,
        'monitor_region': None,
        'click_region': None,
        'num_region': None,
        'shutdown_time': None  # 新增关机时间配置
    }

    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                loaded_config = json.load(f)
                # 合并配置
                for key in loaded_config:
                    if key in config:
                        config[key] = loaded_config[key]
        except:
            pass

    return config


def save_config(config):
    """保存配置文件"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)


# ====== 新增关机功能 ======
def shutdown_computer():
    """执行关机命令，支持Windows和Linux系统"""
    try:
        system_name = platform.system()
        if system_name == "Windows":
            # Windows关机命令（立即关机）
            os.system("shutdown /s /t 0")
        elif system_name == "Linux":
            # Linux关机命令
            os.system("sudo shutdown -h now")
        else:
            print(f"不支持的系统: {system_name}")
    except Exception as e:
        print(f"关机失败: {str(e)}")


class RegionSelector:
    """区域选择器（全屏透明覆盖层）"""

    def __init__(self, parent, title="请框选区域"):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.attributes('-fullscreen', True)  # 全屏
        self.top.attributes('-alpha', 0.2)  # 半透明
        self.top.attributes('-topmost', True)  # 置顶
        self.scaling_factor = self._get_windows_scaling()

        # 创建画布覆盖整个屏幕
        self.canvas = tk.Canvas(self.top,
                                bg='black',
                                highlightthickness=0,
                                cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 绑定事件
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.top.bind("<Escape>", self.cancel)
        self.top.bind("<Destroy>", self.on_window_close)

        # 初始化变量
        self.start_x = None
        self.start_y = None
        self.current_rect = None
        self.region = None
        self.closed_by_user = False
        self.persistent_rects = []  # 存储持久化矩形对象

        # 提示文本
        self.tip_label = self.canvas.create_text(
            self.top.winfo_screenwidth() // 2,
            50,
            text="按住鼠标左键框选区域，释放后确认（ESC取消）",
            font=("微软雅黑", 14, "bold"),
            fill="white"
        )

        # 设置为模态窗口并等待
        self.top.grab_set()
        self.top.wait_window()

    def _get_windows_scaling(self):
        """获取Windows系统DPI缩放比例"""
        try:
            # 获取缩放因子（125%缩放返回1.25）
            scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0
            return scale_factor
        except:
            return 1.0  # 默认无缩放

    def on_press(self, event):
        """鼠标按下事件"""
        self.start_x = event.x
        self.start_y = event.y

        # 创建新矩形
        self.current_rect = self.canvas.create_rectangle(
            self.start_x, self.start_y,
            self.start_x, self.start_y,
            outline="red",
            width=2,
            dash=(4, 4)
        )

    def on_drag(self, event):
        """鼠标拖动事件"""
        if self.current_rect:
            # 更新矩形尺寸
            self.canvas.coords(
                self.current_rect,
                self.start_x, self.start_y,
                event.x, event.y
            )

    def on_release(self, event):
        """鼠标释放事件"""
        end_x = event.x
        end_y = event.y

        # 确保坐标有效（左上->右下）
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)

        self.region = (x1, y1, x2 - x1, y2 - y1)

        # 删除临时虚线框
        if self.current_rect:
            self.canvas.delete(self.current_rect)

        # 创建绿色实线半透明矩形（80%透明度）
        persistent_rect = self.canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="#00FF00",  # 绿色边框
            fill="",  # 透明填充
            stipple="gray12",  # 点状纹理（半透明效果）
            width=2  # 边框粗细
        )
        self.persistent_rects.append(persistent_rect)

        # 添加确认视觉反馈
        self.top.update()
        time.sleep(0.3)

        self.top.destroy()

    def cancel(self, event=None):
        """取消选择"""
        self.region = None
        self.closed_by_user = True
        self.top.destroy()

    def on_window_close(self, event):
        """窗口关闭事件处理"""
        if not self.region:
            self.closed_by_user = True

    def get_region(self):
        """返回选择的区域（x, y, width, height）"""
        return self.region


class ParameterSelector:
    """分辨率选择器（带阈值设置）"""

    def __init__(self):
        self.rs = tk.Tk()
        self.rs.iconbitmap(resource_path('mouse.ico'))
        self.rs.title("参数配置")
        self.center_window()
        self.rs.resizable(False, False)
        self.rs.attributes('-topmost', True)  # 窗口置顶
        self.rs.protocol("WM_DELETE_WINDOW", self.on_close)  # 处理窗口关闭事件

        # 加载上次配置
        self.config = load_config()
        self.max_attempts = self.config['max_attempts']
        self.threshold1 = self.config['threshold1']
        self.threshold2 = self.config['threshold2']
        self.monitor_region = self.config['monitor_region']
        self.click_region = self.config['click_region']
        self.num_region = self.config['num_region']
        self.shutdown_time = self.config['shutdown_time']  # 关机时间

        # 设置样式
        self.rs.configure(bg="#f0f0f0")
        tk.Label(self.rs,
                 text="三脚粥鼠鼠交流群qq:162103846",
                 font=("微软雅黑", 12, "bold"),
                 bg="#f0f0f0").pack(pady=10)

        # === 最大尝试次数区域 ===
        attempts_frame = tk.Frame(self.rs, bg="#f0f0f0")
        attempts_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(attempts_frame,
                 text="最大购买尝试次数:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        self.max_attempts_entry = tk.Entry(attempts_frame, width=10)
        self.max_attempts_entry.grid(row=0, column=1, padx=5, pady=2)
        self.max_attempts_entry.insert(0, str(self.max_attempts))
        tk.Label(attempts_frame, text="次(达到此次数后自动暂停)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=0, column=2, sticky="w")
        # === 区域选择区域 ===
        region_frame = tk.Frame(self.rs, bg="#f0f0f0")
        region_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(region_frame,
                 text="监控区域设置:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        # 监控区域按钮
        self.monitor_btn = tk.Button(region_frame,
                                     text="价格监控区域",
                                     command=self.select_monitor_region,
                                     font=("微软雅黑", 9))
        self.monitor_btn.grid(row=0, column=1, padx=5)

        # 点击按钮区域
        self.click_btn = tk.Button(region_frame,
                                   text="购买按钮区域",
                                   command=self.select_click_region,
                                   font=("微软雅黑", 9))
        self.click_btn.grid(row=0, column=2, padx=5)
        # 点击数量条区域
        self.click_btn = tk.Button(region_frame,
                                   text="数量滑块区域",
                                   command=self.select_num_region,
                                   font=("微软雅黑", 9))
        self.click_btn.grid(row=0, column=3, padx=5)

        # 显示当前区域信息
        self.monitor_label = tk.Label(region_frame,
                                      text="未选择" if not self.monitor_region else
                                      f"价格监控区: ({self.monitor_region[0]},{self.monitor_region[1]})-({self.monitor_region[2]},{self.monitor_region[3]})",
                                      font=("微软雅黑", 8),
                                      fg="#666",
                                      bg="#f0f0f0")
        self.monitor_label.grid(row=1, column=1, columnspan=2, sticky="w")

        self.click_label = tk.Label(region_frame,
                                    text="未选择" if not self.click_region else
                                    f"购买点击点: ({self.click_region[0]}, {self.click_region[1]})",
                                    font=("微软雅黑", 8),
                                    fg="#666",
                                    bg="#f0f0f0")
        self.click_label.grid(row=2, column=1, columnspan=2, sticky="w")
        self.num_label = tk.Label(region_frame,
                                  text="未选择" if not self.num_region else
                                  f"数量点击点: ({self.num_region[0]}, {self.num_region[1]})",
                                  font=("微软雅黑", 8),
                                  fg="#666",
                                  bg="#f0f0f0")
        self.num_label.grid(row=3, column=1, columnspan=2, sticky="w")
        # === 阈值设置区域 ===
        threshold_frame = tk.Frame(self.rs, bg="#f0f0f0")
        threshold_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(threshold_frame,
                 text="设置触发阈值:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        # 下限阈值
        tk.Label(threshold_frame, text="价格下限:", bg="#f0f0f0").grid(row=1, column=0, sticky="e")
        self.threshold1_entry = tk.Entry(threshold_frame, width=10)
        self.threshold1_entry.grid(row=1, column=1, padx=5, pady=2)
        self.threshold1_entry.insert(0, str(self.threshold1))
        tk.Label(threshold_frame, text="哈夫币(低于此值不触发)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=1, column=2, sticky="w")

        # 上限阈值
        tk.Label(threshold_frame, text="价格上限:", bg="#f0f0f0").grid(row=2, column=0, sticky="e")
        self.threshold2_entry = tk.Entry(threshold_frame, width=10)
        self.threshold2_entry.grid(row=2, column=1, padx=5, pady=2)
        self.threshold2_entry.insert(0, str(self.threshold2))
        tk.Label(threshold_frame, text="哈夫币(高于此值不触发)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=2, column=2, sticky="w")

        # 默认值按钮
        def set_last_thresholds():
            self.threshold1_entry.delete(0, tk.END)
            self.threshold1_entry.insert(0, str(self.threshold1))
            self.threshold2_entry.delete(0, tk.END)
            self.threshold2_entry.insert(0, str(self.threshold2))

        tk.Button(threshold_frame, text="恢复为上次参数",
                  command=set_last_thresholds,
                  font=("微软雅黑", 8), bg="#e0e0e0").grid(row=3, column=1, pady=5)

        # === 新增定时关机设置区域 ===
        shutdown_frame = tk.Frame(self.rs, bg="#f0f0f0")
        shutdown_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(shutdown_frame,
                 text="定时关机设置:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        # 关机时间输入框
        tk.Label(shutdown_frame, text="关机时间:", bg="#f0f0f0").grid(row=1, column=0, sticky="e")
        self.shutdown_entry = tk.Entry(shutdown_frame, width=10)
        self.shutdown_entry.grid(row=1, column=1, padx=5, pady=2)
        self.shutdown_entry.insert(0, self.shutdown_time if self.shutdown_time else "")
        tk.Label(shutdown_frame, text="格式: HH:MM (24小时制)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=1, column=2, sticky="w")

        # 当前时间显示
        current_time = datetime.now().strftime("%H:%M")
        tk.Label(shutdown_frame, text=f"当前时间: {current_time}", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=2, column=1, sticky="w", pady=5)

        # 确认按钮
        tk.Button(self.rs, text="开始抢购",
                  command=self.start_monitoring,
                  font=("微软雅黑", 10), bg="#4CAF50", fg="white",
                  padx=20, pady=5).pack(pady=15)

        # 状态变量
        self.threshold1_val = None
        self.threshold2_val = None
        self.shutdown_time_val = None  # 新增关机时间变量
        self.closed_by_user = False

        self.rs.mainloop()

    def center_window(self):
        """自适应内容并居中显示"""
        # 强制更新布局以计算实际尺寸
        self.rs.update_idletasks()  # 关键！确保组件布局完成

        # 获取窗口实际宽高（基于内容）
        width = self.rs.winfo_width()
        height = self.rs.winfo_height()

        # 计算居中位置
        screen_width = self.rs.winfo_screenwidth()
        screen_height = self.rs.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # 应用新位置（保持原窗口尺寸）
        self.rs.geometry(f"+{x}+{y}")

    def select_monitor_region(self):
        """选择监控区域"""
        selector = RegionSelector(self.rs, "请框选监控区域（价格显示区）")
        region = selector.get_region()
        if region:
            self.monitor_region = region
            self.monitor_label.config(
                text=f"价格监控区: {region[0]}x{region[1]} ({region[2]}x{region[3]})"
            )
            # 更新配置
            self.config['monitor_region'] = region
            save_config(self.config)

    def select_click_region(self):
        """选择点击区域"""
        selector = RegionSelector(self.rs, "请框选点击区域（购买按钮）")
        region = selector.get_region()
        if region:
            # 计算中心点作为点击位置
            center_x = region[0] + region[2] // 2
            center_y = region[1] + region[3] // 2
            self.click_region = (center_x, center_y)
            self.click_label.config(
                text=f"购买点击点: ({center_x}, {center_y})"
            )
            # 更新配置
            self.config['click_region'] = self.click_region
            save_config(self.config)

    def select_num_region(self):
        """选择点击区域"""
        selector = RegionSelector(self.rs, "请框选点击区域（数量滑块）")
        region = selector.get_region()
        if region:
            # 计算中心点作为点击位置
            center_x = region[0] + region[2] // 2
            center_y = region[1] + region[3] // 2
            self.num_region = (center_x, center_y)
            self.num_label.config(
                text=f"数量点击点: ({center_x}, {center_y})"
            )
            # 更新配置
            self.config['num_region'] = self.num_region
            save_config(self.config)

    def start_monitoring(self):
        """验证并保存设置"""
        try:

            # 获取并验证最大尝试次数
            max_attempts = int(self.max_attempts_entry.get().replace(",", ""))
            # 获取并验证阈值
            threshold1 = int(self.threshold1_entry.get().replace(",", ""))
            threshold2 = int(self.threshold2_entry.get().replace(",", ""))
            if max_attempts <= 0:
                messagebox.showerror("错误", "最大尝试次数>=1且为整数")
                return
            self.max_attempts_val = max_attempts
            # 保存配置
            self.config['max_attempts'] = max_attempts
            save_config(self.config)

            if threshold1 >= threshold2:
                messagebox.showerror("错误", "上限阈值必须大于下限阈值")
                return

            if threshold1 <= 0 or threshold2 <= 0:
                messagebox.showerror("错误", "阈值必须为正整数")
                return

            self.threshold1_val = threshold1
            self.threshold2_val = threshold2

            # 检查区域是否已选择
            if not self.monitor_region:
                messagebox.showerror("错误", "请先选择监控区域")
                return

            if not self.click_region:
                messagebox.showerror("错误", "请先选择点击区域")
                return

            # 检查数量区域是否已选择
            if not self.num_region:
                messagebox.showerror("错误", "请先选择数量区域")
                return

            # 处理定时关机设置
            shutdown_time_str = self.shutdown_entry.get().strip()
            if shutdown_time_str:
                try:
                    # 验证时间格式
                    if not self.validate_time_format(shutdown_time_str):
                        messagebox.showerror("错误", "关机时间格式错误，请使用HH:MM格式（24小时制）")
                        return

                    # 保存关机时间
                    self.shutdown_time_val = shutdown_time_str
                    self.config['shutdown_time'] = shutdown_time_str
                except ValueError:
                    messagebox.showerror("错误", "关机时间格式错误，请使用HH:MM格式（24小时制）")
                    return
            else:
                self.shutdown_time_val = None
                self.config['shutdown_time'] = None

            # 保存配置
            self.config['threshold1'] = threshold1
            self.config['threshold2'] = threshold2
            save_config(self.config)

            # 关闭窗口
            self.rs.destroy()

        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数阈值")

    def validate_time_format(self, time_str):
        """验证时间格式是否为HH:MM"""
        try:
            # 尝试解析时间
            datetime.strptime(time_str, "%H:%M")
            return True
        except ValueError:
            return False

    def on_close(self):
        """处理窗口关闭事件"""
        self.closed_by_user = True
        self.rs.destroy()


class OverlayApp:
    def __init__(self, root, threshold1, threshold2, max_attempts, monitor_region, click_region, num_region,
                 shutdown_time=None):
        """初始化主应用，接收配置参数"""
        self.root = root
        self.root.title("鼠鼠伴生器灵Ver1.4")

        # 保存阈值配置
        self.THRESHOLD1 = threshold1
        self.THRESHOLD2 = threshold2
        # 保存最大尝试次数
        self.MAX_ATTEMPTS = max_attempts
        # 使用用户选择的区域
        self.MONITOR_REGION = {
            'left': monitor_region[0],
            'top': monitor_region[1],
            'width': monitor_region[2],
            'height': monitor_region[3]
        }
        self.CLICK_POSITION = click_region
        self.NUM_POSITION = num_region
        self.shutdown_time = shutdown_time  # 保存关机时间
        self.shutdown_timer = None  # 关机定时器
        self.shutdown_delay = None  # 关机倒计时（秒）

        # 显示当前配置信息
        config_frame = tk.Frame(root)
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(config_frame,
                 text=f"最大尝试次数: {self.MAX_ATTEMPTS}次",
                 font=("微软雅黑", 8),
                 fg="blue").pack(side=tk.LEFT, padx=10)

        tk.Label(config_frame,
                 text=f"最低价{self.THRESHOLD1:,}HV＄\n最高价{self.THRESHOLD2:,}HV＄",
                 font=("微软雅黑", 9),
                 fg="green").pack(side=tk.RIGHT, padx=10)

        # 显示监控区域信息
        tk.Label(config_frame,
                 text=f"监控区域: {monitor_region[0]}x{monitor_region[1]} "
                      f"({monitor_region[2]}x{monitor_region[3]})",
                 font=("微软雅黑", 8),
                 fg="gray").pack(side=tk.LEFT, padx=10)

        # 窗口设置
        self.root.attributes('-topmost', True)  # 窗口置顶
        self.root.attributes('-alpha', 1)  # 70%透明度

        # 初始化鼠标控制器
        self.mouse = MouseController()
        # 初始化键盘控制器
        self.keyboard = KeyboardController()

        # 创建画布用于显示截图
        self.canvas = Canvas(root,
                             width=self.MONITOR_REGION['width'],
                             height=self.MONITOR_REGION['height'])
        self.canvas.pack()

        # 识别结果显示区域
        self.text_label = tk.Label(root,
                                   text="识别结果将显示在这里",
                                   font=("微软雅黑", 10),
                                   bg='white')
        self.text_label.pack(fill=tk.X, padx=5, pady=5)

        # 状态显示区域
        self.status_frame = Frame(root)
        self.status_frame.pack(fill=tk.X, padx=5, pady=2)

        self.status_label1 = Label(self.status_frame,
                                   text="自动采购: 已暂停",
                                   font=("微软雅黑", 9),
                                   fg="gray")  # 初始为灰色暂停状态
        self.status_label1.pack(side=tk.LEFT)

        self.status_label2 = Label(self.status_frame,
                                   text="自动刷新: 未启动",
                                   font=("微软雅黑", 9),
                                   fg="gray")  # 初始为灰色暂停状态
        self.status_label2.pack(side=tk.RIGHT)

        # ====== 新增关机倒计时显示 ======
        self.shutdown_label = Label(self.status_frame,
                                    text="",
                                    font=("微软雅黑", 9),
                                    fg="purple")
        self.shutdown_label.pack(side=tk.RIGHT, padx=10)

        # 控制按钮
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(fill=tk.X, padx=5, pady=5)

        # 点击计数器
        self.click_count = 0
        self.click_count_label = Label(self.btn_frame,
                                       text=f"点击: 0次",
                                       font=("微软雅黑", 9))
        self.click_count_label.pack(side=tk.LEFT)
        # 开启/暂停按钮
        self.click_paused = True  # 初始状态为暂停
        self.toggle_button = tk.Button(self.btn_frame,
                                       text="允许购买",
                                       command=self.toggle_click,
                                       bg="#4CAF50",
                                       fg="white")
        self.toggle_button.pack(side=tk.LEFT, padx=5)
        tk.Button(self.btn_frame,
                  text="退出",
                  command=self.close_app,
                  bg="#FF6B6B").pack(side=tk.RIGHT, padx=5)

        # === 自动刷新功能 ===
        self.auto_refresh_running = False
        # 使用按钮作为状态指示器（不可点击）
        self.auto_refresh_label = tk.Label(
            self.btn_frame,
            text="按F5自动刷新",
            bg="#2196F3",  # 蓝色背景
            fg="white",
            padx=10,
            pady=5
        )
        self.auto_refresh_label.pack(side=tk.RIGHT, padx=5)

        # 添加线程锁
        self.lock = threading.Lock()

        # 启动全局键盘监听器（解决焦点问题）
        self.start_global_keyboard_listener()

        # 启动截图线程
        self.running = True
        self.thread = threading.Thread(target=self.update_overlay)
        self.thread.daemon = True
        self.thread.start()
        # 添加重新配置按钮
        self.reconfig_btn = tk.Button(self.btn_frame,
                                      text="重新配置",
                                      command=self.initiate_reconfiguration,
                                      bg="#9C27B0",  # 紫色背景
                                      fg="white")
        self.reconfig_btn.pack(side=tk.RIGHT, padx=5)

        # 保存原始配置以便重新初始化
        self.original_config = {
            'threshold1': threshold1,
            'threshold2': threshold2,
            'monitor_region': monitor_region,
            'click_region': click_region,
            'shutdown_time': shutdown_time
        }

        # ====== 启动定时关机功能 ======
        if self.shutdown_time:
            self.start_shutdown_timer()

    def start_shutdown_timer(self):
        """启动定时关机任务"""
        try:
            # 解析关机时间
            shutdown_hour, shutdown_minute = map(int, self.shutdown_time.split(':'))

            # 获取当前时间
            now = datetime.now()
            # 构建目标关机时间（今天）
            shutdown_datetime = now.replace(hour=shutdown_hour, minute=shutdown_minute, second=0, microsecond=0)

            # 如果目标时间已过，设置为明天
            if shutdown_datetime < now:
                shutdown_datetime += timedelta(days=1)

            # 计算时间差（秒）
            self.shutdown_delay = (shutdown_datetime - now).total_seconds()

            # 更新状态显示
            self.shutdown_label.config(
                text=f"定时关机: {self.shutdown_time} (倒计时: {self.format_time(self.shutdown_delay)})")

            # 启动定时关机线程
            self.shutdown_timer = threading.Timer(self.shutdown_delay, self.initiate_shutdown)
            self.shutdown_timer.daemon = True
            self.shutdown_timer.start()

            # 启动倒计时更新
            self.update_shutdown_countdown()

        except Exception as e:
            print(f"定时关机设置失败: {str(e)}")
            self.shutdown_label.config(text=f"定时关机设置失败", fg="red")

    def format_time(self, seconds):
        """将秒数格式化为HH:MM:SS"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def update_shutdown_countdown(self):
        """更新关机倒计时显示"""
        if self.shutdown_delay is None or self.shutdown_delay <= 0:
            return

        # 减少倒计时
        self.shutdown_delay -= 1

        # 更新显示
        if self.shutdown_delay > 0:
            self.shutdown_label.config(
                text=f"定时关机: {self.shutdown_time} (倒计时: {self.format_time(self.shutdown_delay)})")
            # 每秒更新一次
            self.root.after(1000, self.update_shutdown_countdown)
        else:
            self.shutdown_label.config(text="正在关机...", fg="red")

    def initiate_shutdown(self):
        """执行关机操作"""
        # 在主线程中更新UI
        self.root.after(0, lambda: self.shutdown_label.config(text="正在关机...", fg="red"))
        self.root.update()

        # 给用户1秒钟时间看到提示
        time.sleep(1)

        # 执行关机命令
        shutdown_computer()

    def initiate_reconfiguration(self):
        """启动重新配置流程"""
        # 暂停所有操作
        self.click_paused = True
        self.auto_refresh_running = False

        # 取消关机定时器
        if self.shutdown_timer and self.shutdown_timer.is_alive():
            self.shutdown_timer.cancel()

        # 更新状态
        self.status_label1.config(text="自动采购: 配置中...", fg="blue")
        self.status_label2.config(text="自动刷新: 已暂停", fg="gray")
        self.auto_refresh_label.config(text="F5启动自动刷新", bg="#2196F3")

        # 安全关闭当前窗口
        self.safe_shutdown()

        # 创建新配置窗口
        self.root.after(100, self.launch_new_configuration)

    def safe_shutdown(self):
        """安全停止所有线程和资源"""
        # 停止自动刷新
        self.auto_refresh_running = False

        # 停止主循环
        self.running = False

        # 停止键盘监听器
        if hasattr(self, 'key_listener'):
            self.key_listener.stop()

        # 取消关机定时器
        if self.shutdown_timer and self.shutdown_timer.is_alive():
            self.shutdown_timer.cancel()

        # 短暂等待确保线程退出
        time.sleep(0.1)

    def launch_new_configuration(self):
        """启动新的配置窗口"""
        # 关闭当前主窗口
        self.root.destroy()

        # 创建新的配置选择器
        selector = ParameterSelector()

        # 如果用户完成新配置，创建新的主窗口
        if not selector.closed_by_user and selector.threshold1_val and selector.threshold2_val:
            # 创建新的主窗口
            new_root = tk.Tk()
            new_root.iconbitmap(resource_path('mouse.ico'))

            # 计算并设置右上角位置
            screen_width = new_root.winfo_screenwidth()
            new_root.update_idletasks()
            width = new_root.winfo_width()
            x = screen_width - width
            new_root.geometry(f"+{x}+0")

            # 使用新配置启动应用
            OverlayApp(new_root,
                       selector.threshold1_val,
                       selector.threshold2_val,
                       selector.max_attempts_val,
                       selector.monitor_region,
                       selector.click_region,
                       selector.num_region,
                       selector.shutdown_time_val)  # 传递关机时间
            new_root.protocol("WM_DELETE_WINDOW", lambda: self.close_app(new_root))
            new_root.mainloop()

    def close_app(self, window=None):
        """安全退出应用（支持指定窗口）"""
        if not window:
            window = self.root

        # 停止所有操作
        self.auto_refresh_running = False
        self.running = False

        # 停止键盘监听器
        if hasattr(self, 'key_listener'):
            self.key_listener.stop()

        # 取消关机定时器
        if self.shutdown_timer and self.shutdown_timer.is_alive():
            self.shutdown_timer.cancel()

        # 销毁窗口
        window.destroy()
        sys.exit()

    def start_global_keyboard_listener(self):
        """启动全局键盘监听器，解决窗口焦点问题"""
        self.key_listener = KeyboardListener(on_press=self.on_key_press)
        self.key_listener.daemon = True
        self.key_listener.start()

    def on_key_press(self, key):
        """全局键盘事件处理"""
        try:
            # 检测F5按键
            if key == Key.f5:
                # 通过线程安全方式切换状态
                self.root.after(0, self.toggle_auto_refresh)
        except AttributeError:
            pass

    def toggle_auto_refresh(self, event=None):
        """通过F5键切换自动刷新状态"""
        self.auto_refresh_running = not self.auto_refresh_running

        if self.auto_refresh_running:
            # 启动状态
            self.auto_refresh_label.config(text="F5暂停自动刷新", bg="#FF9800")
            # 启动循环点击线程
            self.auto_refresh_thread = threading.Thread(
                target=self.auto_refresh_action,
                daemon=True
            )
            self.auto_refresh_thread.start()
            # 状态提示
            self.status_label2.config(text="自动刷新: 进行中(按F5停止)", fg="orange")
        else:
            # 停止状态
            self.auto_refresh_label.config(text="F5启动自动刷新", bg="#2196F3")
            # 状态提示
            self.status_label2.config(text="自动刷新:已暂停", fg="gray")

    def auto_refresh_action(self):
        """执行循环点击操作"""
        try:
            while self.auto_refresh_running:
                # 左键点击
                self.mouse.press(Button.left)
                time.sleep(0.05)
                self.mouse.release(Button.left)
                time.sleep(0.9)
                # 按下ESC键
                self.keyboard.press(Key.esc)
                time.sleep(0.05)
                self.keyboard.release(Key.esc)

                # 控制点击频率（0.5秒/次）
                for _ in range(10):
                    if not self.auto_refresh_running:
                        return
                    time.sleep(0.05)
        except Exception as e:
            print(f"自动刷新异常: {str(e)}")
            self.status_label2.config(text=f"自动刷新异常: {str(e)}", fg="red")

    def toggle_click(self):
        """切换点击启用状态"""
        with self.lock:
            self.click_paused = not self.click_paused

            if self.click_paused:
                self.toggle_button.config(text="允许购买", bg="#4CAF50")
                self.status_label1.config(text="自动采购: 已暂停", fg="gray")
            else:
                self.toggle_button.config(text="禁止购买", bg="#FF9800")
                self.status_label1.config(text="自动采购: 进行中", fg="green")

    def update_overlay(self):
        """持续更新识别区域内容"""
        with mss() as sct:
            while self.running:
                try:
                    # 1. 截取指定区域
                    screenshot = sct.grab(self.MONITOR_REGION)
                    img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")

                    # 2. 在画布上实时显示
                    self.display_on_canvas(img)

                    # 3. OCR识别
                    self.process_ocr(np.array(img))

                    time.sleep(0.005)  # 控制刷新频率
                except Exception as e:
                    print(f"更新异常: {str(e)}")
                    break

    def display_on_canvas(self, pil_img):
        """在Canvas上显示图像"""
        tk_img = ImageTk.PhotoImage(pil_img.resize(
            (self.MONITOR_REGION['width'], self.MONITOR_REGION['height'])))

        self.canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)
        self.canvas.image = tk_img  # 防止被垃圾回收

    def process_ocr(self, cv_img):
        """执行OCR并更新UI"""
        # 图像预处理
        gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        # _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

        # OCR识别
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789'
        text = pytesseract.image_to_string(gray, config=custom_config).strip()

        # 更新UI显示
        display_text = f"识别结果: {text or '无数字'}"
        self.text_label.config(text=display_text)

        # 检查是否满足点击条件
        if text and text.isdigit():
            current_num = int(text)
            if current_num < self.THRESHOLD2 and current_num > self.THRESHOLD1:
                # 检查是否处于暂停状态
                if not self.click_paused:  # 只有在非暂停状态才执行点击
                    self.perform_click(current_num)
                else:
                    # 显示满足条件但因暂停未点击的状态
                    self.status_label1.config(text=f"满足条件但未购买({current_num})", fg="orange")

    def perform_click(self, current_num):
        """在指定位置执行鼠标点击"""
        try:
            # 更新状态显示
            self.status_label1.config(text=f"自动购买: 尝试以（{current_num}哈夫币）购买", fg="red")

            # 使用pynput执行精确点击
            original_pos = self.mouse.position  # 保存原始位置

            # 移动鼠标到目标位置
            self.mouse.position = self.NUM_POSITION
            time.sleep(0.02)  # 确保移动到位

            # 执行左键点击
            self.mouse.click(Button.left, 1)  # 单次点击
            time.sleep(0.02)  # 确保移动到位

            # 移动鼠标到目标位置
            self.mouse.position = self.CLICK_POSITION
            time.sleep(0.02)  # 确保移动到位

            # 执行左键点击
            self.mouse.click(Button.left, 1)  # 单次点击
            time.sleep(0.02)  # 确保移动到位

            # 可选：返回原始位置（根据需求决定）
            self.mouse.position = original_pos
            time.sleep(0.5)  # 确保移动到位

            # 更新点击计数器
            self.click_count += 1
            self.click_count_label.config(text=f"点击: {self.click_count}次")

            if self.click_count >= self.MAX_ATTEMPTS:
                # 自动刷新暂停
                self.root.after(0, self.toggle_auto_refresh)

                # 自动暂停自动采购
                self.click_paused = True
                self.toggle_button.config(text="允许购买", bg="#4CAF50")
                self.status_label1.config(text=f"已达最大尝试次数({self.MAX_ATTEMPTS})", fg="red")

                # 添加视觉反馈
                self.flash_canvas("red")

                # 显示提示信息
                messagebox.showinfo("提示", f"已达到最大尝试次数({self.MAX_ATTEMPTS})，点击计数归零，自动采购已暂停")

                self.click_count = 0
                self.click_count_label.config(text=f"点击: {self.click_count}次")

            else:
                # 添加视觉反馈
                self.flash_canvas("green")

        except Exception as e:
            self.status_label1.config(text=f"状态: 点击失败 - {str(e)}", fg="red")
            print(f"点击失败: {str(e)}")

    def flash_canvas(self, color):
        """点击成功视觉反馈"""
        self.canvas.config(bg=color)
        self.root.after(100, lambda: self.canvas.config(bg='white'))

    def close_app(self):
        """安全退出应用"""
        self.auto_refresh_running = False  # 停止自动刷新
        self.running = False
        if hasattr(self, 'key_listener'):
            self.key_listener.stop()  # 停止全局键盘监听器
        # 取消关机定时器
        if self.shutdown_timer and self.shutdown_timer.is_alive():
            self.shutdown_timer.cancel()
        time.sleep(0.3)  # 等待线程结束
        self.root.destroy()
        sys.exit()


# 启动应用
if __name__ == "__main__":
    # 设置DPI感知（Windows）
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except:
        pass

    # 先显示配置选择器
    selector = ParameterSelector()

    # 如果用户直接关闭了配置窗口，则退出程序
    if selector.closed_by_user:
        sys.exit()

    # 检查是否完成配置
    if not selector.threshold1_val or not selector.threshold2_val:
        sys.exit()  # 用户未完成配置则退出

    # 检查是否选择了区域
    if not selector.monitor_region or not selector.click_region or not selector.num_region:
        messagebox.showwarning("警告", "请先选择监控区域、点击区域和数量区域！")
        sys.exit()

    # 创建主窗口并传递配置
    root = tk.Tk()
    # 获取窗口实际宽高
    width = root.winfo_width()
    height = root.winfo_height()
    # 获取屏幕宽高
    screen_width = root.winfo_screenwidth()

    # 计算右上角位置（x: 屏幕右边减去窗口宽度，y: 0）
    x = screen_width - width
    y = 0  # 可以根据需要调整，比如设置为50 就是距离顶部50像素

    # 设置窗口位置（不改变尺寸）
    root.iconbitmap(resource_path('mouse.ico'))
    root.geometry(f"+{x}+{y}")

    # 传递配置给主应用
    app = OverlayApp(root,
                     selector.threshold1_val,
                     selector.threshold2_val,
                     selector.max_attempts_val,
                     selector.monitor_region,
                     selector.click_region,
                     selector.num_region,
                     selector.shutdown_time_val)  # 传递关机时间
    root.protocol("WM_DELETE_WINDOW", app.close_app)
    root.mainloop()
