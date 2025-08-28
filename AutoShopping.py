import cv2
import numpy as np
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
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
import hashlib
import socket
import struct
import base64
import uuid


def get_network_time():
    """从NTP服务器获取网络时间戳"""
    ntp_servers = [
        "ntp.aliyun.com"  # 阿里云时间服务器
    ]

    # NTP协议常量
    NTP_PACKET_FORMAT = "!12I"
    NTP_DELTA = 2208988800  # 1900-01-01 00:00:00 到 1970-01-01 00:00:00 的秒数

    for server in ntp_servers:
        try:
            # 创建UDP套接字
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.settimeout(2)  # 设置2秒超时

            # 构建NTP数据包
            packet = bytearray(48)
            packet[0] = 0x1B  # 设置协议版本和模式

            # 发送请求
            sock.sendto(packet, (server, 123))

            # 接收响应
            data, _ = sock.recvfrom(1024)

            if len(data) >= 48:
                # 解析NTP时间戳
                unpacked = struct.unpack(NTP_PACKET_FORMAT, data[:48])
                # 提取传输时间戳（第10个字段）
                ntp_time = unpacked[10] - NTP_DELTA
                return ntp_time
        except (socket.timeout, socket.gaierror, ConnectionRefusedError, OSError):
            continue
        finally:
            try:
                sock.close()
            except:
                pass

    print(f"请检查网络是否连接！")
    return 0
def get_accurate_time():
    """获取准确时间（优先网络时间），返回datetime对象"""
    timestamp = get_network_time()
    return datetime.fromtimestamp(timestamp)
def resource_path(relative_path):
    """获取资源文件的绝对路径。在开发时和打包后均可使用"""
    try:
        # 如果存在_MEIPASS属性，说明程序是打包后运行
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent

    return str(base_path / relative_path)


# 配置文件路径
def get_config_path():
    """获取安全的配置文件路径"""
    if platform.system() == "Windows":
        # 在Windows上使用AppData目录
        appdata_dir = Path(os.getenv('APPDATA'))
    else:
        # 在Linux/macOS上使用主目录
        appdata_dir = Path.home()

    config_dir = appdata_dir / ".mouse_ocr"
    config_dir.mkdir(exist_ok=True, parents=True)
    return config_dir / "config.enc"

# 更新CONFIG_FILE使用
CONFIG_FILE = get_config_path()


# ==================== 配置文件加密部分 ====================
def derive_key():
    """派生加密密钥（基于系统唯一标识符）"""
    # 获取系统标识符（MAC地址+主机名）
    system_id = f"{uuid.getnode()}-{platform.node()}".encode()

    # 使用PBKDF2生成密钥
    salt = b"mouse_ocr_salt"  # 固定盐值
    return hashlib.pbkdf2_hmac('sha256', system_id, salt, 100000, 32)
def encrypt_config(data):
    """加密配置数据"""
    key = derive_key()
    iv = os.urandom(16)  # 随机初始化向量

    # 添加版本标识和校验和
    version = b"MOUSE_OCRV1"
    data_json = json.dumps(data).encode()
    checksum = hashlib.sha256(data_json).digest()[:4]
    plaintext = version + checksum + data_json

    # AES加密
    cipher = AES.new(key, AES.MODE_CBC, iv)
    ciphertext = cipher.encrypt(pad(plaintext, AES.block_size))

    return base64.b64encode(iv + ciphertext).decode()

def decrypt_config(encrypted):
    """解密配置数据"""
    try:
        key = derive_key()
        data = base64.b64decode(encrypted)

        # 提取IV和密文
        iv = data[:16]
        ciphertext = data[16:]

        # AES解密
        cipher = AES.new(key, AES.MODE_CBC, iv)
        plaintext = cipher.decrypt(ciphertext)

        # 移除填充
        plaintext = unpad(plaintext, AES.block_size)

        # 验证版本和校验和
        if plaintext[:11] != b"MOUSE_OCRV1":
            raise ValueError("Invalid config version")

        checksum = plaintext[11:15]
        data_json = plaintext[15:]

        if hashlib.sha256(data_json).digest()[:4] != checksum:
            raise ValueError("Config checksum mismatch")

        return json.loads(data_json.decode())
    except Exception as e:
        print(f"Config decryption failed: {str(e)}")
        return None


def load_config():
    """加载并解密配置文件"""
    # 默认配置
    default_config = {
        'threshold1': "愣着干嘛？",
        'threshold2': "等我给填？",
        'max_attempts': 1,
        'max_success': 1,
        'monitor_region': None,
        'click_region': None,
        'num_region': None,
        'text_region': None,
        'shutdown_time': None,
        'auto_refresh_time': None,
        'refresh_interval_steps': 10,
        'activation_timestamp': None,
        'valid_until_timestamp': None,
        'last_activation_date': None
    }

    if not CONFIG_FILE.exists():
        return default_config

    try:
        with open(CONFIG_FILE, 'r') as f:
            encrypted = f.read()
            config = decrypt_config(encrypted)
            if config:
                return config
    except Exception as e:
        print(f"Error loading config: {str(e)}")

    return default_config


def save_config(config):
    """加密并保存配置"""
    encrypted = encrypt_config(config)
    config_path = get_config_path()

    # 确保目录存在
    config_path.parent.mkdir(exist_ok=True, parents=True)

    with open(config_path, 'w') as f:
        f.write(encrypted)

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
        self.max_success = self.config['max_success']
        self.threshold1 = self.config['threshold1']
        self.threshold2 = self.config['threshold2']
        self.monitor_region = self.config['monitor_region']
        self.click_region = self.config['click_region']
        self.num_region = self.config['num_region']
        self.text_region = self.config['text_region']
        self.shutdown_time = self.config['shutdown_time']  # 关机时间
        self.auto_refresh_time = self.config['auto_refresh_time']  # 关机时间
        self.refresh_interval_steps = self.config['refresh_interval_steps']
        # 设置样式
        self.rs.configure(bg="#f0f0f0")
        tk.Label(self.rs,
                 text="三脚粥鼠鼠交流群qq:162103846"
                      "\n本软件完全免费，如果你是花钱购买的，请立即退款并举报卖家",
                 font=("微软雅黑", 8, "bold"),
                 bg="#f0f0f0").pack(pady=10)
        # 创建刷新间隔控制面板
        self.create_refresh_control_panel()
        # === 最多尝试次数区域 ===
        attempts_frame = tk.Frame(self.rs, bg="#f0f0f0")
        attempts_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(attempts_frame,
                 text="最多点击次数:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        self.max_attempts_entry = tk.Entry(attempts_frame, width=10)
        self.max_attempts_entry.grid(row=0, column=1, padx=5, pady=2)
        self.max_attempts_entry.insert(0, str(self.max_attempts))
        tk.Label(attempts_frame, text="次(达到此次数后自动暂停)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=0, column=2, sticky="w")
        success_frame = tk.Frame(self.rs, bg="#f0f0f0")
        success_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(success_frame,
                 text="最多成功次数:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        self.max_success_entry = tk.Entry(success_frame, width=10)
        self.max_success_entry.grid(row=0, column=1, padx=5, pady=2)
        self.max_success_entry.insert(0, str(self.max_success))
        tk.Label(success_frame, text="次(达到此次数后自动暂停)", bg="#f0f0f0",
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
                                     text="数字监控",
                                     command=self.select_monitor_region,
                                     font=("微软雅黑", 9))
        self.monitor_btn.grid(row=0, column=1, padx=5)

        # 点击按钮区域
        self.click_btn = tk.Button(region_frame,
                                   text="触发按钮",
                                   command=self.select_click_region,
                                   font=("微软雅黑", 9))
        self.click_btn.grid(row=0, column=2, padx=5)
        # 点击数量条区域
        self.click_btn = tk.Button(region_frame,
                                   text="数量滑块",
                                   command=self.select_num_region,
                                   font=("微软雅黑", 9))
        self.click_btn.grid(row=0, column=3, padx=5)

        # 文本监控区域按钮
        self.text_btn = tk.Button(region_frame,
                                  text="文本监控",
                                  command=self.select_text_region,
                                  font=("微软雅黑", 9))
        self.text_btn.grid(row=0, column=4, padx=5)
        # 显示当前区域信息
        self.monitor_label = tk.Label(region_frame,
                                      text="未选择" if not self.monitor_region else
                                      f"价格监控区: ({self.monitor_region[0]},{self.monitor_region[1]})-WH({self.monitor_region[2]}x{self.monitor_region[3]})",
                                      font=("微软雅黑", 8),
                                      fg="#666",
                                      bg="#f0f0f0")
        self.monitor_label.grid(row=1, column=1, columnspan=2, sticky="w")

        self.click_label = tk.Label(region_frame,
                                    text="未选择" if not self.click_region else
                                    f"点击触发点: ({self.click_region[0]}, {self.click_region[1]})",
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
        self.text_label = tk.Label(region_frame,
                                      text="未选择" if not self.text_region else
                                      f"文本监控区: ({self.text_region[0]},{self.text_region[1]})-WH({self.text_region[2]}x{self.text_region[3]})",
                                      font=("微软雅黑", 8),
                                      fg="#666",
                                      bg="#f0f0f0")
        self.text_label.grid(row=4, column=1, columnspan=2, sticky="w")
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
        timer_frame = tk.Frame(self.rs, bg="#f0f0f0")
        timer_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(timer_frame,
                 text="定时设置:",
                 font=("微软雅黑", 10),
                 bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)

        # 关机时间输入框
        tk.Label(timer_frame, text="          执行关机时间:", bg="#f0f0f0",
                 font=("微软雅黑", 9),).grid(row=1, column=0, sticky="e")
        self.shutdown_entry = tk.Entry(timer_frame, width=10)
        self.shutdown_entry.grid(row=1, column=1, padx=5, pady=2)
        self.shutdown_entry.insert(0, self.shutdown_time if self.shutdown_time else "")
        tk.Label(timer_frame, text="格式: HH:MM (24小时制)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=1, column=2, sticky="w")
        # === 新增自动刷新时间设置区域 ===
        tk.Label(timer_frame, text="         自动刷新时间:", bg="#f0f0f0",
                 font=("微软雅黑", 9),).grid(row=2, column=0, sticky="e")
        self.refresh_entry = tk.Entry(timer_frame, width=10)
        self.refresh_entry.grid(row=2, column=1, padx=5, pady=2)
        self.refresh_entry.insert(0, self.config.get('auto_refresh_time', "") or "")
        tk.Label(timer_frame, text="格式: HH:MM (24小时制)", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=2, column=2, sticky="w")
        # 当前时间显示
        current_time = get_accurate_time().strftime("%H:%M")
        tk.Label(timer_frame, text=f"当前时间: {current_time}", bg="#f0f0f0",
                 font=("微软雅黑", 8), fg="#666").grid(row=0, column=1, sticky="w", pady=5)


        # 确认按钮
        tk.Button(self.rs, text="保存参数",
                  command=self.start_monitoring,
                  font=("微软雅黑", 10), bg="#4CAF50", fg="white",
                  padx=20, pady=5).pack(pady=15)
        # 状态变量
        self.threshold1_val = None
        self.threshold2_val = None
        self.shutdown_time_val = None  # 新增关机时间变量
        self.auto_refresh_time_val = None  # 新增关机时间变量
        self.closed_by_user = False

        self.rs.mainloop()
    def create_activation_status(self):
        """创建激活状态显示区域"""
        frame = tk.Frame(self.rs, bg="#f0f0f0")

        self.activation_label = tk.Label(
            frame,
            text="激活状态: 未激活",
            font=("微软雅黑", 9),
            fg="red",
            bg="#f0f0f0"
        )
        self.activation_label.pack(side=tk.LEFT)

        self.valid_until_label = tk.Label(
            frame,
            text="",
            font=("微软雅黑", 8),
            fg="#666",
            bg="#f0f0f0"
        )
        self.valid_until_label.pack(side=tk.LEFT, padx=10)

        return frame
    def create_refresh_control_panel(self):
        """创建刷新间隔控制面板"""
        control_frame = tk.Frame(self.rs, bg="#f0f0f0", padx=10, pady=5)
        control_frame.pack(fill=tk.X)
        self.seconds = self.refresh_interval_steps * 0.05

        # 标题标签
        tk.Label(control_frame,
                 text="  自动刷新间隔:",
                 font=("微软雅黑",10),
                 bg="#f0f0f0").pack(side=tk.LEFT, padx=(0, 5))

        # 减号按钮
        self.minus_btn = tk.Button(
            control_frame,
            text="-",
            command=self.decrease_refresh_interval,
            font=("微软雅黑", 9, "bold"),
            width=3,
            bg="#FF6B6B",
            fg="white"
        )
        self.minus_btn.pack(side=tk.LEFT)

        # 当前值显示
        self.refresh_label = tk.Label(
            control_frame,
            text=f"{self.seconds:.2f} ",
            font=("微软雅黑", 9),
            bg="white",
            relief="sunken",
            width=12,
            padx=5
        )
        self.refresh_label.pack(side=tk.LEFT, padx=5)

        # 加号按钮
        self.plus_btn = tk.Button(
            control_frame,
            text="+",
            command=self.increase_refresh_interval,
            font=("微软雅黑", 9, "bold"),
            width=3,
            bg="#4CAF50",
            fg="white"
        )
        self.plus_btn.pack(side=tk.LEFT)

        # 提示标签
        tk.Label(control_frame,
                 text="(单位:秒)",
                 font=("微软雅黑", 8),
                 fg="#666",
                 bg="#f0f0f0").pack(side=tk.LEFT, padx=(10, 0))

    def increase_refresh_interval(self):
        """增加刷新间隔步数"""
        self.refresh_interval_steps += 1
        self.update_refresh_display()

    def decrease_refresh_interval(self):
        """减少刷新间隔步数（最小值1）"""
        if self.refresh_interval_steps > 1:
            self.refresh_interval_steps -= 1
            self.update_refresh_display()
        else:
            # 达到最小值时提供视觉反馈
            self.refresh_label.config(bg="#FFD700")
            self.rs.after(300, lambda: self.refresh_label.config(bg="white"))

    def update_refresh_display(self):
        """更新显示值并计算对应秒数"""
        self.seconds = self.refresh_interval_steps * 0.05
        self.refresh_label.config(
            text=f"{self.seconds:.2f}",
            bg="white"
        )
        # 保存配置
        self.config['refresh_interval_steps'] = self.refresh_interval_steps
        save_config(self.config)
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
        selector = RegionSelector(self.rs, "请框选监控区域（数字显示区）")
        region = selector.get_region()
        if region:
            self.monitor_region = region
            self.monitor_label.config(
                text=f"价格监控区: ({region[0]},{region[1]})-WH({region[2]}x{region[3]})"
            )
            # 更新配置
            self.config['monitor_region'] = region
            save_config(self.config)

    def select_click_region(self):
        """选择点击区域"""
        selector = RegionSelector(self.rs, "请框选点击区域（点击触发按钮）")
        region = selector.get_region()
        if region:
            # 计算中心点作为点击位置
            center_x = region[0] + region[2] // 2
            center_y = region[1] + region[3] // 2
            self.click_region = (center_x, center_y)
            self.click_label.config(
                text=f"点击触发点: ({center_x}, {center_y})"
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
    def select_text_region(self):
        """选择监控区域"""
        selector = RegionSelector(self.rs, "请框选监控区域（文本显示区）")
        text_region = selector.get_region()
        if text_region:
            self.text_region = text_region
            self.text_label.config(
                text=f"文本监控区: ({text_region[0]},{text_region[1]})-WH({text_region[2]}x{text_region[3]})"
            )
            # 更新配置
            self.config['text_region'] = text_region
            save_config(self.config)
    def start_monitoring(self):
        try:

            # 获取并验证最多尝试次数
            max_attempts = int(self.max_attempts_entry.get().replace(",", ""))
            # 获取并验证最多尝试次数
            max_success = int(self.max_success_entry.get().replace(",", ""))
            # 获取并验证阈值
            threshold1 = int(self.threshold1_entry.get().replace(",", ""))
            threshold2 = int(self.threshold2_entry.get().replace(",", ""))
            if max_attempts <= 0:
                messagebox.showerror("错误", "最多尝试次数>=1且为整数")
                return
            self.max_attempts_val = max_attempts
            # 保存配置
            self.config['max_attempts'] = max_attempts
            save_config(self.config)
            if max_success <= 0:
                messagebox.showerror("错误", "最多成功次数>=1且为整数")
                return
            self.max_success_val = max_success
            # 保存配置
            self.config['max_success'] = max_success
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
            # 处理自动刷新时间设置
            auto_refresh_time_str = self.refresh_entry.get().strip()
            if auto_refresh_time_str:
                try:

                    if not self.validate_time_format(auto_refresh_time_str):
                        messagebox.showerror("错误", "自动刷新时间格式错误，请使用HH:MM格式（24小时制）")
                        return
                    self.auto_refresh_time_val = auto_refresh_time_str
                    self.config['auto_refresh_time'] = auto_refresh_time_str
                except ValueError:
                    messagebox.showerror("错误", "自动刷新时间格式错误，请使用HH:MM格式（24小时制）")
                    return
            else:
                self.auto_refresh_time_val = None
                self.config['auto_refresh_time'] = None
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
    def __init__(self, root, threshold1, threshold2, max_attempts, max_success, monitor_region, click_region, num_region,text_region,
                 shutdown_time=None, auto_refresh_time=None,refresh_interval_steps=10):
        """初始化主应用，接收配置参数"""
        self.root = root
        self.root.title("鼠鼠伴生器灵Ver2.3")
        # 添加新的事件和状态标志

        # 保存阈值配置
        self.THRESHOLD1 = threshold1
        self.THRESHOLD2 = threshold2
        # 保存最多尝试次数
        self.MAX_ATTEMPTS = max_attempts
        self.MAX_SUCCESS = max_success

        # 使用用户选择的区域
        self.MONITOR_REGION = {
            'left': monitor_region[0],
            'top': monitor_region[1],
            'width': monitor_region[2],
            'height': monitor_region[3]
        }
        self.TEXT_REGION = {
            'left': text_region[0],
            'top': text_region[1],
            'width': text_region[2],
            'height': text_region[3]
        }
        self.CLICK_POSITION = click_region
        self.NUM_POSITION = num_region

        self.shutdown_time = shutdown_time  # 保存关机时间
        self.shutdown_timer = None  # 关机定时器
        self.shutdown_delay = None  # 关机倒计时（秒）

        self.auto_refresh_time = auto_refresh_time  # 保存自动刷新时间
        self.auto_refresh_timer = None  # 自动刷新定时器
        self.auto_refresh_delay = None  # 自动刷新倒计时（秒）
        self.refresh_interval_steps = refresh_interval_steps

        self.shutdown_after_id = None
        self.auto_refresh_after_id = None

        # 显示当前配置信息
        config_frame = tk.Frame(root)
        config_frame.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(config_frame,
                 text=f"最多尝试次数: {self.MAX_ATTEMPTS}次\n最多成功次数: {self.MAX_SUCCESS}次",
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
        self.canvas.pack(pady=(5, 0))
        # 文本监控区域画布
        self.text_canvas = Canvas(root,
                                  width=self.TEXT_REGION['width'],
                                  height=self.TEXT_REGION['height'])
        self.text_canvas.pack(pady=(5, 0))
        # 识别结果显示区域
        self.result_label = tk.Label(root,
                                   text="数字识别未就绪",
                                   font=("微软雅黑", 10),
                                   bg='white')
        self.result_label.pack(fill=tk.X, padx=5, pady=5)
        # 识别结果显示区域
        self.text_result_label = tk.Label(root,
                                   text="文本识别未就绪",
                                   font=("微软雅黑", 10),
                                   bg='white')
        self.text_result_label.pack(fill=tk.X, padx=5, pady=5)

        # 状态显示区域
        self.status_frame = Frame(root)
        self.status_frame.pack(fill=tk.X, padx=5, pady=2)

        self.status_label1 = Label(self.status_frame,
                                   text="自动点击: 已暂停",
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
        # ====== 新增自动刷新倒计时显示 ======
        self.auto_refresh_label = Label(self.status_frame,
                                       text="",
                                       font=("微软雅黑", 9),
                                       fg="blue")
        self.auto_refresh_label.pack(side=tk.RIGHT, padx=10)
        # 控制按钮
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(fill=tk.X, padx=5, pady=5)

        # 点击计数器
        self.click_count = 0
        self.click_count_label = Label(self.btn_frame,
                                       text=f"点击: 0次",
                                       font=("微软雅黑", 9))
        self.click_count_label.pack(side=tk.LEFT)
        # 成功次数计数器
        self.success_count = 0
        self.success_count_label = Label(self.btn_frame,
                                        text=f"成功: 0次",
                                        font=("微软雅黑", 9))
        self.success_count_label.pack(side=tk.LEFT, padx=10)  # 放在点击次数旁边
        # 开启/暂停按钮
        self.click_paused = True  # 初始状态为暂停
        self.toggle_button = tk.Button(self.btn_frame,
                                       text="允许点击",
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

        # 添加线程锁
        self.action_lock = threading.Lock()

        # 启动全局键盘监听器（解决焦点问题）
        self.start_global_keyboard_listener()

        # 启动截图线程
        self.running = True
        self.thread = threading.Thread(target=self.update_overlay)
        self.thread.daemon = True
        self.thread.start()
        # +++ 启动文本监控线程 +++
        self.text_thread = threading.Thread(target=self.update_text_overlay)
        self.text_thread.daemon = True
        self.text_thread.start()
        # 添加重新配置按钮
        self.reconfig_btn = tk.Button(self.btn_frame,
                                      text="重新配置",
                                      command=self.initiate_reconfiguration,
                                      bg="#9C27B0",  # 紫色背景
                                      fg="white")
        self.reconfig_btn.pack(side=tk.RIGHT, padx=5)


        # ====== 启动自动刷新定时功能 ======
        if self.auto_refresh_time:
            self.start_auto_refresh_timer()
        # ====== 启动定时关机功能 ======
        if self.shutdown_time:
            self.start_shutdown_timer()
        # 获取窗口实际宽高（基于内容）
        width = self.root.winfo_width()
        height = self.root.winfo_height()

        # 计算居中位置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = 3*(screen_width - width) // 4
        y = (screen_height - height) // 10

        # 应用新位置（保持原窗口尺寸）
        self.root.geometry(f"+{x}+{y}")

    def start_auto_refresh_timer(self):
        """启动自动刷新定时任务"""
        try:
            # 解析自动刷新时间
            refresh_hour, refresh_minute = map(int, self.auto_refresh_time.split(':'))

            # 获取当前时间
            now = get_accurate_time()
            # 构建目标自动刷新时间（今天）
            refresh_datetime = now.replace(hour=refresh_hour, minute=refresh_minute,
                                           second=0, microsecond=0)

            # 如果目标时间已过，设置为明天
            if refresh_datetime < now:
                refresh_datetime += timedelta(days=1)

            # 计算时间差（秒）
            self.auto_refresh_delay = (refresh_datetime - now).total_seconds()

            # 更新状态显示
            self.auto_refresh_label.config(
                text=f"自动刷新: {self.auto_refresh_time} (倒计时: {self.format_time(self.auto_refresh_delay)})")

            # 启动定时器
            self.auto_refresh_timer = threading.Timer(self.auto_refresh_delay, self.initiate_auto_refresh)
            self.auto_refresh_timer.daemon = True
            self.auto_refresh_timer.start()

            # 启动倒计时更新
            self.update_auto_refresh_countdown()

        except Exception as e:
            print(f"自动刷新设置失败: {str(e)}")
            self.auto_refresh_label.config(text=f"自动刷新设置失败", fg="red")

    def initiate_auto_refresh(self):
        """执行自动刷新操作"""
        # 在主线程中更新UI
        if self.click_paused:  # 若点击功能被暂停
            self.toggle_click()  # 启用点击
        self.root.after(0, lambda: self.auto_refresh_label.config(text="正在启动自动刷新...", fg="white"))
        self.root.update()

        # 执行自动刷新
        if not self.auto_refresh_running:
            self.toggle_auto_refresh()

        # 更新显示
        self.auto_refresh_label.config(text="自动刷新已启动", fg="white")

    def update_auto_refresh_countdown(self):
        """更新自动刷新倒计时显示"""
        if self.auto_refresh_delay is None or self.auto_refresh_delay <= 0:
            return

        # 减少倒计时
        self.auto_refresh_delay -= 1

        # 更新显示
        if self.auto_refresh_delay > 0:
            self.auto_refresh_label.config(
                text=f"自动刷新: {self.auto_refresh_time} (倒计时: {self.format_time(self.auto_refresh_delay)})"
                     f"\n---按下F5以取消定时立刻开始自动刷新---")
            if self.auto_refresh_after_id  is not None:
                self.root.after_cancel(self.auto_refresh_after_id )
            # 每秒更新一次
            self.auto_refresh_after_id = self.root.after(1000, self.update_auto_refresh_countdown)
        else:
            self.auto_refresh_label.config(text="正在启动自动刷新...", fg="white")

    def start_shutdown_timer(self):
        """启动定时关机任务"""
        try:
            # 解析关机时间
            shutdown_hour, shutdown_minute = map(int, self.shutdown_time.split(':'))

            # 获取当前时间
            now = get_accurate_time()
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
            if self.shutdown_after_id is not None:
                self.root.after_cancel(self.shutdown_after_id)
            # 每秒更新一次
            self.shutdown_after_id = self.root.after(1000, self.update_shutdown_countdown)
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
        # 取消自动刷新定时器
        if self.auto_refresh_timer and self.auto_refresh_timer.is_alive():
            self.auto_refresh_timer.cancel()
        # 更新状态
        self.status_label1.config(text="自动点击: 配置中...", fg="blue")
        self.status_label2.config(text="自动刷新: 已暂停", fg="gray")
        self.auto_refresh_label.config(text="F5启动自动刷新", bg="#2196F3")

        # 安全关闭当前窗口
        self.safe_shutdown()

        # 创建新配置窗口
        self.root.after(100, self.launch_new_configuration)

    def safe_shutdown(self):
        """安全停止所有线程和资源"""
        # 停止倒计时的after调度
        if self.shutdown_after_id is not None:
            self.root.after_cancel(self.shutdown_after_id)
            self.shutdown_after_id = None
        if self.auto_refresh_after_id is not None:
            self.root.after_cancel(self.auto_refresh_after_id)
            self.auto_refresh_after_id = None

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
                       selector.max_success_val,
                       selector.monitor_region,
                       selector.click_region,
                       selector.num_region,
                       selector.text_region,
                       selector.shutdown_time_val,  # 传递关机时间
                       selector.auto_refresh_time_val,# 传递自动刷新时间
                       selector.refresh_interval_steps)
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
        # 取消自动刷新定时器
        if self.auto_refresh_timer and self.auto_refresh_timer.is_alive():
            self.auto_refresh_timer.cancel()
        if hasattr(self, 'thread') and self.thread.is_alive():
            self.thread.join(0.5)

        if hasattr(self, 'text_thread') and self.text_thread.is_alive():
            self.text_thread.join(0.5)
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
        # 取消自动刷新定时器
        self.auto_refresh_delay = None
        if self.auto_refresh_running:

            if self.click_paused:  # 若点击功能被暂停
                self.toggle_click()  # 启用点击
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
            self.toggle_click()  # 启用点击
            # 停止状态
            self.auto_refresh_label.config(text="F5启动自动刷新", bg="#2196F3")
            # 状态提示
            self.status_label2.config(text="自动刷新:已暂停", fg="gray")

    def auto_refresh_action(self):
        event = threading.Event()
        """执行循环点击操作"""
        steps = self.refresh_interval_steps

        try:
            while self.auto_refresh_running:

                # 如果点击完成后自动刷新被停止了，则退出
                if not self.auto_refresh_running:
                    return
                event.wait(0.05)
                self.mouse.press(Button.left)
                event.wait(0.05)
                self.mouse.release(Button.left)
                # 使用短时循环替代长sleep
                for _ in range(steps):  # 拆分成N次*0.05秒
                    if not self.auto_refresh_running:
                        return
                    event.wait(0.05)
                if self.price_value is not None and self.THRESHOLD1 < self.price_value < self.THRESHOLD2:
                    self.perform_click(self.price_value)
                # ESC按键
                if not self.auto_refresh_running:
                    return

                self.keyboard.press(Key.esc)
                event.wait(0.05)
                self.keyboard.release(Key.esc)
                # 使用短时循环替代长sleep
                for _ in range(steps):  # 拆分成N次*0.05秒
                    if not self.auto_refresh_running:
                        return
                    event.wait(0.05)

        except Exception as e:
            print(f"自动刷新异常: {str(e)}")
            self.root.after(0, lambda: self.status_label2.config(
                text=f"自动刷新异常: {str(e)}", fg="red"))
    def toggle_click(self):
        """切换点击启用状态"""
        with self.lock:
            self.click_paused = not self.click_paused
            if self.click_paused:
                self.toggle_button.config(text="允许点击", bg="#4CAF50")
                self.status_label1.config(text="自动点击: 已暂停", fg="gray")
            else:
                self.toggle_button.config(text="禁止点击", bg="#FF9800")
                self.status_label1.config(text="自动点击: 进行中", fg="green")

    def update_overlay(self):
        """持续更新识别区域内容"""
        with mss() as sct:
            event = threading.Event()
            # 创建字典存储不同分辨率的模板 [1](@ref)
            templates = {resolution: {} for resolution in ['1k', '2k', '4k']}

            # 加载三个分辨率的所有模板
            for resolution in templates.keys():
                for digit in range(10):
                    template_path = resource_path(f"digits{resolution}/{digit}.png")
                    try:
                        with open(template_path, 'rb') as f:
                            img_data = np.frombuffer(f.read(), dtype=np.uint8)
                            template = cv2.imdecode(img_data, cv2.IMREAD_GRAYSCALE)
                        templates[resolution][digit] = template
                    except Exception as e:
                        print(f"加载模板失败 [{resolution}/{digit}]: {e}")

            while self.running:
                start_time = time.time()

                try:
                    # 1. 截取指定区域
                    screenshot = sct.grab(self.MONITOR_REGION)
                    img_bgra = np.array(screenshot)
                    if img_bgra.size == 0:
                        print("警告：空截图，跳过本帧")
                        continue

                    # 转换为灰度图并进行预处理
                    gray = cv2.cvtColor(img_bgra, cv2.COLOR_BGRA2GRAY)
                    self.adaptive_threshold = self.preprocess_image(gray)
                    img = Image.fromarray(self.adaptive_threshold)

                    # 2. 在画布上实时显示
                    self.display_on_canvas(img)

                    # 3. 多分辨率OCR识别
                    self.process_ocr(img_bgra, templates)

                    process_time =  time.time() - start_time
                    wait_time = max(0.05 - process_time, 0.001)
                    event.wait(timeout=wait_time)
                except Exception as e:
                    print(f"更新覆盖层出错: {e}")
                    break
    def display_on_canvas(self, pil_img):
        """在Canvas上显示图像"""
        tk_img = ImageTk.PhotoImage(pil_img.resize(
            (self.MONITOR_REGION['width'], self.MONITOR_REGION['height'])))

        self.canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)
        self.canvas.image = tk_img  # 防止被垃圾回收

    def process_ocr(self, img_bgra, templates):
        """执行OCR并更新UI，支持多分辨率模板匹配与优化"""
        self.price_value = None  # 默认设置为None，表示当前没有识别到价格
        # 1. 图像预处理强化
        gray = cv2.cvtColor(img_bgra, cv2.COLOR_BGRA2GRAY)
        self.adaptive_threshold = self.preprocess_image(gray)
        # 在预处理后，进行填充
        padding_width = 20
        self.adaptive_threshold = cv2.copyMakeBorder(self.adaptive_threshold,
                                                padding_width, padding_width,
                                                padding_width, padding_width,
                                                cv2.BORDER_CONSTANT, value=0)
        # 2. 存储匹配结果 (x, y, w, h, digit, confidence, resolution)
        matches = []

        # 多分辨率模板匹配
        for resolution, digit_templates in templates.items():
            for digit, template in digit_templates.items():
                # 跳过空模板
                if template.size == 0:
                    continue

                # 模板匹配
                res = cv2.matchTemplate(self.adaptive_threshold, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(res)

                # 动态阈值：最高置信度的75%且不低于0.72
                threshold = max(0.7, max_val * 0.75)
                loc = np.where(res >= threshold)

                # 提取匹配位置
                w, h = template.shape[::-1]
                for pt in zip(*loc[::-1]):
                    confidence = res[pt[1], pt[0]]
                    # 存储完整信息（新增resolution字段）
                    matches.append((pt[0], pt[1], w, h, digit, confidence, resolution))

        # 3. 非极大值抑制（NMS）处理重复匹配
        if not matches:
            self.result_label.config(text="识别结果: 无数字")
            return

        boxes_wh = np.array([[x, y, w, h] for (x, y, w, h, _, _, _) in matches])
        confidences = [conf for (_, _, _, _, _, conf, _) in matches]

        # OpenCV NMS处理
        indices = cv2.dnn.NMSBoxes(
            boxes_wh.tolist(),
            confidences,
            score_threshold=0.5,
            nms_threshold=0.3  # 重叠度>30%则抑制
        )

        # 4. 保留NMS筛选后的匹配
        filtered_matches = []
        if indices is not None:
            for i in indices.flatten():
                x, y, w, h, digit, conf, res = matches[i]
                filtered_matches.append((x, digit, conf, res))
        else:
            self.result_label.config(text="识别结果: 无有效匹配")
            return

        # 5. 空间聚类与数字序列组合
        filtered_matches.sort(key=lambda x: x[0])  # 按x坐标排序
        digits = []
        prev_x = -100
        used_resolutions = set()

        for x, digit, conf, res in filtered_matches:
            if abs(x - prev_x) > 5:  # 基础去重
                digits.append(str(digit))
                prev_x = x
                used_resolutions.add(res)  # 记录使用的分辨率

        # 6. 结果格式化与校验
        price_str = ''.join(digits)
        if not price_str.isdigit():
            self.result_label.config(text=f"非法字符: {price_str}")
            return

        self.price_value = int(price_str)
        self.price_formatted = f"{self.price_value:,}" if len(price_str) > 3 else price_str
        # 7. 更新UI与条件检查（添加分辨率信息）
        res_info = "&".join(used_resolutions) if used_resolutions else "无"
        display_text = f"识别结果: {self.price_formatted} (置信度: {np.mean(confidences):.2f}, 模板: {res_info})"
        self.result_label.config(text=display_text)


    def preprocess_image(self,image):
        """
        对图像进行自适应二值化预处理，可选择添加形态学操作
        参数:
            image: 输入灰度图像
            apply_morphology: 是否应用形态学操作
        返回:
            预处理后的二值图像
        """
        # 自适应二值化参数设置
        block_size = 31  # 邻域大小，必须是奇数
        c_value = -2  # 从计算阈值中减去的常数

        # 应用自适应二值化（高斯加权）
        binary = cv2.adaptiveThreshold(
            image,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            block_size,
            c_value
        )

        return binary
    def load_success_templates(self):
        """加载成功标志模板"""
        self.success_templates = []
        template_dir = resource_path('success')

        for i in range(3):  # 0,1,2
            template_path = os.path.join(template_dir, f'{i}.png')
            try:
                template_img = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
                if template_img is not None:
                    self.success_templates.append(template_img)
            except Exception as e:
                print(f"加载模板 {i}.png 失败: {e}")


    def match_success_templates(self, text_img):
        """在文本监控区域匹配成功标志"""
        if not hasattr(self, 'success_templates') or not self.success_templates:
            return False

        # 预处理文本区域图像
        # 对每个模板进行匹配
        for template in self.success_templates:
            res = cv2.matchTemplate(text_img, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, _ = cv2.minMaxLoc(res)

            # 如果匹配度超过阈值则认为匹配成功
            if max_val > 0.68:
                return max_val,True

        return 0,False

    def update_text_overlay(self):
        """更新文本监控区域内容"""
        with mss() as sct:
            event = threading.Event()
            self.load_success_templates()  # 加载模板

            # 记录上一帧是否匹配成功（用于边缘检测）
            last_match = False

            while self.running:
                start_time = time.time()
                try:
                    # 截取文本监控区域
                    screenshot = sct.grab(self.TEXT_REGION)
                    img_bgra = np.array(screenshot)

                    if img_bgra.size == 0:
                        event.wait(0.05)
                        continue

                    # 转换为灰度图并进行预处理
                    gray = cv2.cvtColor(img_bgra, cv2.COLOR_BGRA2GRAY)
                    self.adaptive_threshold_text = self.preprocess_image(gray)
                    img = Image.fromarray(self.adaptive_threshold_text)
                    # 在预处理后，进行填充
                    padding_width = 70
                    padded_img  = cv2.copyMakeBorder(self.adaptive_threshold_text,
                                                                 padding_width, padding_width,
                                                                 padding_width, padding_width,
                                                                 cv2.BORDER_CONSTANT, value=0)
                    # 在画布上实时显示
                    self.display_on_text_canvas(img)

                    # 尝试匹配成功标志
                    confidences, current_match = self.match_success_templates(padded_img)

                    # 只在从非匹配状态变为匹配状态时计数（上升沿触发）
                    if current_match and not last_match:
                        # 成功计数器加1
                        self.success_count += 1
                        self.root.after(0, lambda: self.success_count_label.config(
                            text=f"成功: {self.success_count}次"))

                        # 更新识别结果显示
                        self.root.after(0, lambda: self.text_result_label.config(
                            text=f"文本识别: 成功匹配! (置信度: {np.mean(confidences):.2f})",
                            fg="green"))

                        # 添加视觉反馈
                        self.flash_text_canvas("green")

                    # 更新匹配状态
                    last_match = current_match

                    # 更新识别状态（如果没有匹配）
                    if not current_match:
                        self.root.after(0, lambda: self.text_result_label.config(
                            text=f"文本识别: 未匹配成功",
                            fg="black"))

                    process_time = time.time() - start_time
                    wait_time = max(0.05 - process_time, 0.001)
                    event.wait(timeout=wait_time)

                except Exception as e:
                    print(f"更新文本监控出错: {e}")
                    # 打印图像尺寸
                    try:
                        print(f"图像尺寸: {padded_img.shape}")
                    except:
                        print("图像尺寸未知")
                    time.sleep(0.1)

    def display_on_text_canvas(self, pil_img):
        """在文本Canvas上显示图像"""
        tk_img = ImageTk.PhotoImage(pil_img.resize(
            (self.TEXT_REGION['width'], self.TEXT_REGION['height'])))

        self.text_canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)
        self.text_canvas.image = tk_img  # 防止被垃圾回收

    def flash_text_canvas(self, color):
        """文本匹配成功视觉反馈"""
        self.text_canvas.config(bg=color)
        self.root.after(100, lambda: self.text_canvas.config(bg='white'))
    def perform_click(self, current_num):
        """在指定位置执行鼠标点击"""
        event = threading.Event()

        try:
            # 设置点击中状态，暂停自动刷新
            self.is_clicking = True
            # 更新状态显示
            self.status_label1.config(text=f"自动点击: 尝试以（{current_num}Hv＄）点击", fg="red")
            # 更新点击计数器
            self.click_count += 1
            self.click_count_label.config(text=f"点击: {self.click_count}次")
            # 使用pynput执行精确点击
            original_pos = self.mouse.position  # 保存原始位置

            # 移动鼠标到目标位置
            self.mouse.position = self.NUM_POSITION
              # 确保移动到位
            # 执行左键点击
            self.mouse.click(Button.left, 1)  # 单次点击
            event.wait(0.01)  # 确保移动到位

            # 移动鼠标到目标位置
            self.mouse.position = self.CLICK_POSITION
            event.wait(0.01)  # 确保移动到位

            # 执行左键点击
            self.mouse.click(Button.left, 1)  # 单次点击
            event.wait(0.01)  # 确保移动到位

            # 可选：返回原始位置（根据需求决定）
            self.mouse.position = original_pos
            event.wait(0.5)  # 确保移动到位



            if self.click_count >= self.MAX_ATTEMPTS:
                # 自动刷新暂停
                self.auto_refresh_running = False
                self.auto_refresh_label.config(text="F5启动自动刷新", bg="#2196F3")
                self.status_label2.config(text="自动刷新:已暂停", fg="gray")

                # 自动暂停自动采购
                self.click_paused = True
                self.toggle_button.config(text="允许点击", bg="#4CAF50")
                self.status_label1.config(text=f"已达最多尝试次数({self.MAX_ATTEMPTS})", fg="red")

                # 添加视觉反馈
                self.flash_canvas("red")

                # 显示提示信息
                messagebox.showinfo("提示", f"已达到最多尝试次数({self.MAX_ATTEMPTS})，计数归零，自动点击已暂停")

                self.click_count = 0
                self.click_count_label.config(text=f"点击: {self.click_count}次")
                self.success_count = 0
                self.success_count_label.config(text=f"成功: {self.success_count}次")
            if self.success_count >= self.MAX_SUCCESS:
                # 自动刷新暂停
                self.auto_refresh_running = False
                self.auto_refresh_label.config(text="F5启动自动刷新", bg="#2196F3")
                self.status_label2.config(text="自动刷新:已暂停", fg="gray")

                # 自动暂停自动采购
                self.click_paused = True
                self.toggle_button.config(text="允许点击", bg="#4CAF50")
                self.status_label1.config(text=f"已达最多成功次数({self.MAX_SUCCESS})", fg="red")

                # 添加视觉反馈
                self.flash_canvas("red")

                # 显示提示信息
                messagebox.showinfo("提示", f"已达最多成功次数({self.MAX_SUCCESS})，计数归零，自动点击已暂停")
                self.click_count = 0
                self.click_count_label.config(text=f"点击: {self.click_count}次")
                self.success_count = 0
                self.success_count_label.config(text=f"成功: {self.success_count}次")
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
    if not selector.monitor_region or not selector.click_region or not selector.num_region or not selector.text_region:
        messagebox.showwarning("警告", "请先选择监控区域、点击区域、数量区域和文本区域！")
        sys.exit()

    # 创建主窗口并传递配置
    root = tk.Tk()
    # 设置窗口位置（不改变尺寸）
    root.iconbitmap(resource_path('mouse.ico'))

    # 传递配置给主应用
    app = OverlayApp(root,
                     selector.threshold1_val,
                     selector.threshold2_val,
                     selector.max_attempts_val,
                     selector.max_success_val,
                     selector.monitor_region,
                     selector.click_region,
                     selector.num_region,
                     selector.text_region,
                     selector.shutdown_time_val,  # 传递关机时间
                     selector.auto_refresh_time_val,  # 传递自动刷新时间
                     selector.refresh_interval_steps)
    root.protocol("WM_DELETE_WINDOW", app.close_app)
    root.mainloop()
