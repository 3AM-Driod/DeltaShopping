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
from pynput.keyboard import Key, Controller as KeyboardController, Listener as KeyboardListener  # 新增全局键盘监听器
import os
import json
from pathlib import Path

def resource_path(relative_path):
    """ 获取资源文件的绝对路径。在开发时和打包后均可使用 """
    try:
        # 如果存在_MEIPASS属性，说明程序是打包后运行
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent

    return str(base_path / relative_path)

# ====== 默认阈值配置 ======
DEFAULT_THRESHOLD1 = 200000
DEFAULT_THRESHOLD2 = 320000

# ====== 分辨率配置映射 ======
RESOLUTION_MAP = {
    "1080p": {"width": 1920, "height": 1080},
    "2K": {"width": 2560, "height": 1440},
    "4K": {"width": 3840, "height": 2160}
}
# ====== Tesseract配置 ======
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

# 配置文件路径
CONFIG_FILE = "ocr_config.json"


def load_config():
    """加载配置文件，返回阈值设置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                if 'threshold1' in config and 'threshold2' in config:
                    return config['threshold1'], config['threshold2']
        except:
            pass
    return DEFAULT_THRESHOLD1, DEFAULT_THRESHOLD2


def save_config(threshold1, threshold2):
    """保存配置文件"""
    config = {
        'threshold1': threshold1,
        'threshold2': threshold2
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)


class ResolutionSelector:
    """分辨率选择器（带阈值设置）"""

    def __init__(self):

        self.rs = tk.Tk()
        self.rs.iconbitmap(resource_path('mouse.ico'))
        self.rs.title("配置监控参数")
        self.center_window()
        self.rs.resizable(False, False)
        self.rs.attributes('-topmost', True)  # 窗口置顶

        # 加载上次配置
        self.threshold1, self.threshold2 = load_config()

        # 设置样式
        self.rs.configure(bg="#f0f0f0")
        tk.Label(self.rs,
                 text="请配置监控参数",
                 font=("微软雅黑", 12, "bold"),
                 bg="#f0f0f0").pack(pady=10)

        # === 分辨率选择区域 ===
        tk.Label(self.rs,
                 text="选择屏幕分辨率:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").pack(anchor="w", padx=20, pady=(10, 5))

        # 创建分辨率选择按钮框架
        res_frame = tk.Frame(self.rs, bg="#f0f0f0")
        res_frame.pack(fill=tk.X, padx=20, pady=5)

        # 创建分辨率选择按钮
        resolutions = ["1080p", "2K", "4K"]
        self.selected_resolution = tk.StringVar(value="2K")  # 默认选择2K
        for res in resolutions:
            tk.Radiobutton(res_frame, text=res, variable=self.selected_resolution,
                           value=res, bg="#f0f0f0", font=("微软雅黑", 9),
                           selectcolor="#e0e0e0").pack(side=tk.LEFT, padx=10)

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
        def set_default_thresholds():
            self.threshold1_entry.delete(0, tk.END)
            self.threshold1_entry.insert(0, str(DEFAULT_THRESHOLD1))
            self.threshold2_entry.delete(0, tk.END)
            self.threshold2_entry.insert(0, str(DEFAULT_THRESHOLD2))

        tk.Button(threshold_frame, text="恢复默认",
                  command=set_default_thresholds,
                  font=("微软雅黑", 8), bg="#e0e0e0").grid(row=3, column=1, pady=5)

        # 确认按钮
        tk.Button(self.rs, text="开始监控",
                  command=self.start_monitoring,
                  font=("微软雅黑", 10), bg="#4CAF50", fg="white",
                  padx=20, pady=5).pack(pady=15)

        # 状态变量
        self.resolution = None
        self.threshold1_val = None
        self.threshold2_val = None

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
    def start_monitoring(self):
        """验证并保存设置"""
        try:
            # 获取分辨率
            self.resolution = self.selected_resolution.get()

            # 获取并验证阈值
            threshold1 = int(self.threshold1_entry.get().replace(",", ""))
            threshold2 = int(self.threshold2_entry.get().replace(",", ""))

            if threshold1 >= threshold2:
                messagebox.showerror("错误", "上限阈值必须大于下限阈值")
                return

            if threshold1 <= 0 or threshold2 <= 0:
                messagebox.showerror("错误", "阈值必须为正整数")
                return

            self.threshold1_val = threshold1
            self.threshold2_val = threshold2

            # 保存配置
            save_config(threshold1, threshold2)

            # 关闭窗口
            self.rs.destroy()

        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数阈值")


class OverlayApp:
    def __init__(self, root, resolution, threshold1, threshold2):
        """初始化主应用，接收配置参数"""
        self.root = root
        self.root.title("鼠鼠伴生器灵Ver1.1")

        # 保存阈值配置
        self.THRESHOLD1 = threshold1
        self.THRESHOLD2 = threshold2

        # 根据选择的分辨率设置参数
        res_info = RESOLUTION_MAP[resolution]
        self.WIDTH = res_info["width"]
        self.HEIGHT = res_info["height"]

        # ====== 动态计算区域 ======
        self.MONITOR_REGION = {
            'top': int(self.HEIGHT * 1240 / 1440),
            'left': int(self.WIDTH * 180 / 2560),
            'width': int(self.WIDTH * 200 / 2560),
            'height': int(self.HEIGHT * 40 / 1440)
        }
        self.CLICK_POSITION = (
            int(self.WIDTH * 2200 / 2560),
            int(self.HEIGHT * 1220 / 1440)
        )
        # ========================

        # 显示当前配置信息
        config_frame = tk.Frame(root)
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(config_frame,
                 text=f"分辨率: {resolution} ({self.WIDTH}x{self.HEIGHT})",
                 font=("微软雅黑", 9),
                 fg="blue").pack(side=tk.LEFT, padx=10)

        tk.Label(config_frame,
                 text=f"最低价{self.THRESHOLD1:,}HV＄\n最高价{self.THRESHOLD2:,}HV＄",
                 font=("微软雅黑", 9),
                 fg="green").pack(side=tk.RIGHT, padx=10)

        # 窗口设置
        self.root.attributes('-topmost', True)  # 窗口置顶
        self.root.attributes('-alpha', 0.7)  # 70%透明度

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

    def start_global_keyboard_listener(self):
        """启动全局键盘监听器，解决窗口焦点问题[6,7](@ref)"""
        self.key_listener = KeyboardListener(on_press=self.on_key_press)
        self.key_listener.daemon = True
        self.key_listener.start()

    def on_key_press(self, key):
        """全局键盘事件处理[6,7](@ref)"""
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
            self.status_label1.config(text=f"自动购买: 尝试以 ({current_num} 哈夫币)购买", fg="red")

            # 使用pynput执行精确点击
            original_pos = self.mouse.position  # 保存原始位置

            # 移动鼠标到目标位置
            self.mouse.position = self.CLICK_POSITION
            time.sleep(0.05)  # 确保移动到位

            # 执行左键点击
            self.mouse.click(Button.left, 1)  # 单次点击
            time.sleep(0.05)  # 确保移动到位

            # 可选：返回原始位置（根据需求决定）
            self.mouse.position = original_pos
            time.sleep(0.5)  # 确保移动到位

            # 更新点击计数器
            self.click_count += 1
            self.click_count_label.config(text=f"点击: {self.click_count}次")

            # 添加视觉反馈
            self.flash_canvas("green")

            print(f"已点击: {self.CLICK_POSITION} (识别值: {current_num})")
        except Exception as e:
            self.status_label.config(text=f"状态: 点击失败 - {str(e)}", fg="red")
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
        time.sleep(0.3)  # 等待线程结束
        self.root.destroy()
        sys.exit()

# 启动应用
if __name__ == "__main__":
    # 先显示配置选择器
    selector = ResolutionSelector()

    if not selector.resolution or not selector.threshold1_val or not selector.threshold2_val:
        sys.exit()  # 用户未完成配置则退出

    # 创建主窗口并传递配置
    root = tk.Tk()
    # 获取窗口实际宽高
    width = root.winfo_width()
    height = root.winfo_height()
    # 获取屏幕宽高
    screen_width = root.winfo_screenwidth()

    # 计算右上角位置（x: 屏幕右边减去窗口宽度，y: 0）
    x = screen_width - width
    y = 0  # 可以根据需要调整，比如设置为50就是距离顶部50像素

    # 设置窗口位置（不改变尺寸）
    root.geometry(f"+{x}+{y}")
    root.iconbitmap(resource_path('mouse.ico'))

    # 传递配置给主应用
    app = OverlayApp(root,
                     selector.resolution,
                     selector.threshold1_val,
                     selector.threshold2_val)
    root.protocol("WM_DELETE_WINDOW", app.close_app)
    root.mainloop()
