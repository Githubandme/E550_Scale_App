import serial
import serial.tools.list_ports
import time
import threading
import logging
import sys
import ctypes
import subprocess
import re
import tkinter as tk
from tkinter import ttk, simpledialog
import hashlib
import json
import os
from typing import Optional, Callable

import requests
import winsound
import ttkbootstrap as tb  # 导入ttkbootstrap库
from tkinter import messagebox

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('weight_display_app.log', encoding='utf-8')
    ]
)
logger = logging.getLogger("E550_Scale_App")

# 增强的依赖检查：确保serial.tools模块可用
try:
    import serial
    import serial.tools.list_ports
except ImportError as e:
    serial = None
    logger.error("缺少必要的pyserial库或子模块: %s", e)
    logger.error("请重新安装pyserial库: pip install --force-reinstall pyserial")
    
    # 在启动前显示错误提示
    if messagebox is not None:
        messagebox.showerror(
            "缺少必要依赖", 
            "未检测到完整的pyserial库，秤重功能将无法使用。\n\n"
            "请重新安装pyserial:\n"
            "pip install --force-reinstall pyserial\n\n"
            "点击确定退出程序。"
        )
    sys.exit(1)  # 退出程序

# 设置文件路径和默认设置
SETTINGS_FILE = "settings.json"
DEFAULT_SETTINGS = {
    "device_no": "DEVICE_001",
    "api_domain": "api.example.com",
    "api_port": "80",
    "user_id": "USER_ID",
    "security_key": "SECURITY_KEY"
}
SETTINGS_PASSWORD = "password"  # 请在首次使用后修改此密码

def load_settings() -> dict:
    """从 settings.json 加载配置，若文件不存在则返回默认值。"""
    logger.info("加载设置文件 %s", SETTINGS_FILE)
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                logger.info("成功加载设置: %s", settings)
                return settings
        logger.info("未找到设置文件，使用默认设置")
        return DEFAULT_SETTINGS
    except Exception as e:
        logger.error("加载设置失败: %s，使用默认设置", e)
        return DEFAULT_SETTINGS

def save_settings(settings: dict) -> None:
    """将配置保存到 settings.json。"""
    logger.info("保存设置到 %s", SETTINGS_FILE)
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=4)
        logger.info("设置保存成功")
    except Exception as e:
        logger.error("保存设置失败: %s", e)

class SerialConnectionManager:
    """管理串口连接、数据接收和重置操作。"""
    def __init__(self, port: Optional[str] = None, data_callback: Optional[Callable[[str], None]] = None, 
                 status_callback: Optional[Callable[[str], None]] = None):
        logger.info("初始化 SerialConnectionManager，端口: %s", port)
        self.port = port
        self.serial_conn = None
        self.connection_status = "disconnected"
        self.data_received_buffer = ""
        self.running = False
        self.receive_thread = None
        self.primary_baudrate = 9600
        self.reset_baudrate = 19200
        self.data_callback = data_callback
        self.status_callback = status_callback
        
        if self.port is None:
            self.port = self.auto_detect_ch340_port()
            logger.info("自动检测端口: %s", self.port or "未找到")
    
    def set_port(self, port: str) -> None:
        """手动设置串口端口。"""
        logger.info("手动设置串口端口: %s", port)
        self.port = port
        self.update_status(f"端口已设置为 {port}")

    def auto_detect_ch340_port(self) -> Optional[str]:
        """自动检测 CH340 串口设备（Windows）。"""
        if not sys.platform.startswith('win'):
            logger.info("非Windows系统，跳过CH340检测")
            return None
        
        try:
            ports = list(serial.tools.list_ports.comports())
            for port in ports:
                if 'CH340' in port.description:
                    logger.info("检测到CH340设备: %s", port.device)
                    return port.device
                if hasattr(port, 'vid') and port.vid == 0x1A86 and port.pid == 0x7523:
                    logger.info("通过VID/PID检测到CH340: %s", port.device)
                    return port.device
            
            # 备用检测：WMI 或设备管理器
            ch340_ports = self.query_wmi_for_ch340()
            if ch340_ports:
                return ch340_ports[0]
            
            device_info = self.get_device_manager_info()
            if device_info:
                for info in device_info:
                    if 'CH340' in info or '1A86' in info:
                        match = re.search(r'\(COM(\d+)\)', info)
                        if match:
                            return f"COM{match.group(1)}"
            
            available_ports = self.list_available_ports()
            return available_ports[0] if available_ports else None
        except Exception as e:
            logger.error("CH340检测失败: %s", e)
            return None

    def query_wmi_for_ch340(self) -> list:
        """使用 WMI 查询 CH340 设备。"""
        try:
            import wmi
            c = wmi.WMI()
            return [item.DeviceID for item in c.Win32_SerialPort() if 'VID_1A86&PID_7523' in item.PNPDeviceID]
        except ImportError:
            logger.warning("未安装wmi库，跳过WMI查询")
            return []
        except Exception as e:
            logger.error("WMI查询失败: %s", e)
            return []

    def get_device_manager_info(self) -> list:
        """获取设备管理器中的串口信息。"""
        try:
            cmd = ['powershell.exe', 'Get-PnpDevice -Class Ports | Where-Object {$_.Status -eq "OK"} | Select-Object FriendlyName, DeviceID']
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=5, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                return [line.strip() for line in result.stdout.split('\n') if 'COM' in line]
        except Exception as e:
            logger.error("获取设备管理器信息失败: %s", e)
        return []

    def list_available_ports(self) -> list:
        """列出所有可用串口。"""
        try:
            return [port.device for port in serial.tools.list_ports.comports()]
        except Exception as e:
            logger.error("列出串口失败: %s", e)
            return []

    def update_status(self, status: str) -> None:
        """更新连接状态并调用回调。"""
        self.connection_status = status
        if self.status_callback:
            self.status_callback(status)
        logger.info("连接状态: %s", status)

    def connect(self) -> bool:
        """尝试以 9600 波特率连接，失败则尝试重置。"""
        if not self.port:
            self.update_status("未指定串口")
            return False
        
        if self._try_connect(self.primary_baudrate):
            return True
        return self._reset_and_connect()

    def _reset_and_connect(self) -> bool:
        """执行重置并重新连接。"""
        if self._perform_reset():
            return self._try_connect(self.primary_baudrate)
        return False

    def _perform_reset(self) -> bool:
        """使用 19200 波特率执行重置。"""
        if self._try_connect(self.reset_baudrate, is_reset=True):
            self._safe_close()
            time.sleep(1)
            return True
        return False

    def _try_connect(self, baudrate: int, is_reset: bool = False, attempts: int = 1) -> bool:
        """尝试连接串口。"""
        for attempt in range(1, attempts + 1):
            try:
                self._safe_close()
                if sys.platform == 'win32':
                    self._force_release_resources()
                    time.sleep(0.5)
                
                self.serial_conn = serial.Serial(
                    port=self.port,
                    baudrate=baudrate,
                    bytesize=serial.EIGHTBITS,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    timeout=0.5,
                    exclusive=True
                )
                
                if self.serial_conn.is_open:
                    status = f"connected @ {baudrate}bps{' (重置)' if is_reset else ''}"
                    self.update_status(status)
                    if not is_reset:
                        self.start_receiving()
                    return True
            except Exception as e:
                if serial is None:
                    self.update_status("connect_failed: pyserial库未安装")
                    return False
                elif serial is not None and isinstance(e, serial.SerialException):
                    self.update_status(f"connect_failed: {str(e)}")
                    if "PermissionError" in str(e) or "Access is denied" in str(e):
                        return False
                    time.sleep(0.5)
                else:
                    self.update_status(f"unexpected_error: {str(e)}")
                    time.sleep(0.5)
        return False

    def start_receiving(self) -> None:
        """启动数据接收线程。"""
        if self.running:
            return
        self.running = True
        self.receive_thread = threading.Thread(target=self._receive_data, daemon=True)
        self.receive_thread.start()
        logger.info("启动数据接收线程")

    def stop_receiving(self) -> None:
        """停止数据接收线程。"""
        if self.running:
            self.running = False
            if self.receive_thread and self.receive_thread.is_alive():
                self.receive_thread.join(timeout=1.0)
            logger.info("数据接收线程停止")

    def _receive_data(self) -> None:
        """接收串口数据并处理。"""
        try:
            while self.running and self.serial_conn and self.serial_conn.is_open:
                if self.serial_conn.in_waiting:
                    data = self.serial_conn.read(self.serial_conn.in_waiting).decode('ascii', errors='replace')
                    self.data_received_buffer += data
                    while '=' in self.data_received_buffer:
                        line, self.data_received_buffer = self.data_received_buffer.split('=', 1)
                        full_data_packet = line + '='
                        printable_data = ''.join(c for c in full_data_packet if c.isprintable() or c in ['\r', '\n', '='])
                        if printable_data.strip() and self.data_callback:
                            self.data_callback(printable_data.strip())
                time.sleep(0.05)
        except Exception as e:
            if serial is not None and isinstance(e, serial.SerialException):
                self.update_status(f"read_error: {str(e)}")
            else:
                self.update_status(f"receive_error: {str(e)}")
        finally:
            self.stop_receiving()

    def _safe_close(self) -> None:
        """安全关闭串口连接。"""
        self.stop_receiving()
        if self.serial_conn and self.serial_conn.is_open:
            try:
                self.serial_conn.dtr = False
                self.serial_conn.rts = False
                self.serial_conn.reset_input_buffer()
                self.serial_conn.reset_output_buffer()
                self.serial_conn.close()
                self.update_status("closed")
            except Exception as e:
                self.update_status(f"close_error: {str(e)}")
        if sys.platform == 'win32':
            self._force_release_resources()

    def _force_release_resources(self) -> None:
        """强制释放串口资源（Windows）。"""
        if sys.platform != 'win32':
            return
        try:
            subprocess.run(["mode", self.port], stdout=subprocess.PIPE, stderr=subprocess.PIPE, 
                          timeout=1, creationflags=subprocess.CREATE_NO_WINDOW)
            kernel32 = ctypes.windll.kernel32
            port_name = f"\\\\.\\{self.port}"
            handle = kernel32.CreateFileW(
                port_name, 0x80000000 | 0x40000000, 0, None, 3, 0x80, None)
            if handle != -1:
                kernel32.CloseHandle(handle)
        except Exception as e:
            logger.debug("释放资源失败: %s", e)

    def close(self) -> None:
        """关闭连接并执行重置。"""
        self._safe_close()
        self._perform_reset()

class WeightDisplayApp:
    """电子秤重量显示与上传 UI。"""
    def __init__(self, master: tb.Window):  # 修改为ttkbootstrap的Window
        self.master = master
        master.title("电子秤重量显示与上传")
        master.geometry("1400x800")
        master.resizable(False, False)
        # 使用ttkbootstrap主题
        self.style = tb.Style(theme="flatly")  # 使用flatly主题，也可选其他如cosmo, minty等
        self.serial_manager = None
        self.connected = False
        self.current_weight = "0.00 kg"
        self.raw_weight_value = 0.0
        self.upload_logs = []
        self.last_weight = None
        self.last_weight_time = None
        self.stable_duration = 1.5
        self.flash_state = False
        self.stability_check_id = None
        self.port_var = None

        self.settings = load_settings()
        self.api_base_url = f"http://{self.settings['api_domain']}:{self.settings['api_port']}"
        
        # 移除旧的颜色定义，使用ttkbootstrap主题颜色
        self.bg_color = self.style.colors.get('bg')  # 获取主题背景色
        self.btn_bg_color = self.style.colors.primary  # 获取主题主色
        
        self.font_family = "Arial"
        self.base_font_size = 10
        self.log_font_size = 10

        # 使用ttkbootstrap样式
        self._setup_ui()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.check_weight_stability()
        self.master.after(0, self.connect_serial)

    def _setup_ui(self) -> None:
        """初始化 UI 组件。"""
        # 使用ttkbootstrap内置样式，移除自定义样式
        # 创建菜单栏
        self.menu_bar = tb.Menu(self.master)
        self.master.config(menu=self.menu_bar)
        self.menu_bar.add_command(label="设置 (Settings)", command=self.open_settings)
        self.menu_bar.add_command(label="帮助 (Help)", command=self.show_help)
        self.menu_bar.add_command(label="退出 (Exit)", command=self.on_closing)

        self.main_frame = tb.Frame(self.master)
        self.main_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

        self.left_frame = tb.Frame(self.main_frame, width=500)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        self.left_frame.pack_propagate(False)

        self.input_frame = tb.Frame(self.left_frame)
        self.input_frame.pack(fill=tk.X, padx=40, pady=20)
        tb.Label(self.input_frame, text="扫描单号 (Scan No.):", 
                 font=(self.font_family, 12, "bold")).grid(
                     row=0, column=0, columnspan=2, sticky=tk.W, pady=12, padx=15)
        self.scan_no_entry = tb.Entry(self.input_frame, font=(self.font_family, 18), width=40)
        self.scan_no_entry.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=12, padx=15)
        self.scan_no_entry.focus_set()
        self.scan_no_entry.bind('<Return>', lambda event: self.initiate_upload_weight_thread())

        for i, label in enumerate(["包裹长 (Length, cm):", "包裹宽 (Width, cm):", "包裹高 (Height, cm):"], 2):
            tb.Label(self.input_frame, text=label, font=(self.font_family, 12)).grid(
                row=i, column=0, sticky=tk.W, pady=12, padx=15)
            entry = tb.Entry(self.input_frame, font=(self.font_family, 12), width=50)
            entry.grid(row=i, column=1, sticky=tk.EW, pady=12, padx=15)
            setattr(self, f"{'length' if i==2 else 'width' if i==3 else 'height'}_entry", entry)
        
        self.input_frame.grid_columnconfigure(1, weight=1)

        self.weight_frame = tb.Frame(self.left_frame)
        self.weight_frame.pack(pady=20)
        self.weight_label = tb.Label(self.weight_frame, text=self.current_weight, 
                                     font=(self.font_family, 48, "bold"))
        self.weight_label.pack(side=tk.LEFT, padx=(0, 10))
        self.status_indicator = tb.Canvas(self.weight_frame, width=20, height=20, highlightthickness=0)
        self.indicator_id = self.status_indicator.create_oval(0, 0, 20, 20, fill="gray")
        self.status_indicator.pack(side=tk.LEFT)

        self.status_label = tb.Label(self.left_frame, text="状态: 未连接", 
                                     font=(self.font_family, 14))
        self.status_label.pack(pady=10)
        self.upload_status_label = tb.Label(self.left_frame, text="上传状态: 等待操作", 
                                           font=(self.font_family, 12))
        self.upload_status_label.pack(pady=5)

        self.button_frame = tb.Frame(self.left_frame)
        self.button_frame.pack(pady=20, padx=40)
        self.toggle_button = tb.Button(self.button_frame, text="连接 (Connect)", 
                                        command=self.toggle_connection, 
                                        bootstyle="primary")  # 使用ttkbootstrap按钮样式
        self.toggle_button.pack(side=tk.LEFT, padx=25)
        self.upload_button = tb.Button(self.button_frame, text="上传重量 (Upload Weight)", 
                                        command=self.initiate_upload_weight_thread, 
                                        bootstyle="success", state=tk.DISABLED)  # 使用ttkbootstrap按钮样式
        self.upload_button.pack(side=tk.LEFT, padx=25)
        self.exit_button = tb.Button(self.button_frame, text="退出 (Exit)", 
                                     command=self.on_closing, bootstyle="danger")  # 使用ttkbootstrap按钮样式
        self.exit_button.pack(side=tk.LEFT, padx=15)

        self.log_frame = tb.Frame(self.main_frame, width=600)
        self.log_frame.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_frame.pack_propagate(False)
        tb.Label(self.log_frame, text="最近上传记录 (Recent Uploads)", 
                 font=(self.font_family, 12, "bold")).pack(anchor=tk.NW, padx=5, pady=5)
        self.log_text = tb.Text(self.log_frame, font=(self.font_family, self.log_font_size), 
                               width=50, height=15, wrap="word")
        self.log_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        self.log_text.bind("<1>", lambda event: self.log_text.focus_set())
        scrollbar = tb.Scrollbar(self.log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

    def process_received_data(self, data: str) -> None:
        """处理串口数据，格式 ab.cdef= 反转为 fedc.ba kg。"""
        logger.info("处理串口数据: %s", data)
        try:
            if not data.endswith('='):
                logger.warning("数据格式无效，未以'='结尾: %s", data)
                self.master.after(0, self.update_weight_display, "--- kg")
                self.raw_weight_value = None
                return

            data_stripped = data[:-1]
            if not re.fullmatch(r'\d{2}\.\d{4}', data_stripped):
                logger.warning("数据格式不匹配 'ab.cdef': %s", data_stripped)
                self.master.after(0, self.update_weight_display, "--- kg")
                self.raw_weight_value = None
                return

            reversed_data = data_stripped.replace('.', '')[::-1]
            weight_str = f"{reversed_data[:4]}.{reversed_data[4:]}"
            try:
                weight_kg = float(weight_str)
                self.raw_weight_value = weight_kg
                self.current_weight = f"{weight_kg:.2f} kg"
                logger.info("转换重量成功: %s (原始: %s)", self.current_weight, data_stripped)
                self.master.after(0, self.update_weight_display, self.current_weight)
            except ValueError:
                logger.error("转换浮点数失败: %s", weight_str)
                self.master.after(0, self.update_weight_display, "--- kg")
                self.raw_weight_value = None
        except Exception as e:
            logger.error("处理数据失败: %s", e)
            self.master.after(0, self.update_weight_display, "--- kg")
            self.raw_weight_value = None

    def check_weight_stability(self) -> None:
        """检查重量稳定性并更新指示灯。"""
        if not self.connected:
            self.status_indicator.itemconfig(self.indicator_id, fill="gray")
        elif self.current_weight == "--- kg":
            self.status_indicator.itemconfig(self.indicator_id, fill="yellow")
        else:
            current_time = time.time()
            if self.last_weight != self.raw_weight_value:
                self.last_weight = self.raw_weight_value
                self.last_weight_time = current_time
                self.flash_state = False

            time_elapsed = current_time - self.last_weight_time if self.last_weight_time else 0
            if time_elapsed < self.stable_duration:
                if self.raw_weight_value and self.raw_weight_value > 0:
                    self.flash_state = not self.flash_state
                    self.status_indicator.itemconfig(self.indicator_id, fill="green" if self.flash_state else "gray")
                else:
                    self.status_indicator.itemconfig(self.indicator_id, fill="blue")
            else:
                self.status_indicator.itemconfig(self.indicator_id, 
                    fill="blue" if self.raw_weight_value == 0 else "green" if self.raw_weight_value > 0 else "red")
        
        self.stability_check_id = self.master.after(100, self.check_weight_stability)

    def update_upload_log(self, log_entry: str) -> None:
        """更新上传记录，保留最近10条。"""
        self.upload_logs.append(log_entry)
        if len(self.upload_logs) > 10:
            self.upload_logs.pop(0)
        self.log_text.delete(1.0, tk.END)
        for entry in self.upload_logs:
            self.log_text.insert(tk.END, entry + "\n\n")
        self.log_text.see(tk.END)

    def open_settings(self) -> None:
        """打开设置窗口，需密码验证。"""
        password = simpledialog.askstring("输入密码", "请输入设置密码:", show='*', parent=self.master)
        if password != SETTINGS_PASSWORD:
            self.update_status_display("密码错误 (Incorrect Password)")
            return

        settings_window = tb.Toplevel(self.master)
        settings_window.title("设置 (Settings)")
        settings_window.geometry("500x400")
        settings_window.resizable(False, False)

        settings_frame = tb.Frame(settings_window)
        settings_frame.pack(pady=20, padx=20, fill=tk.BOTH)

        entries = {}
        for i, (key, label) in enumerate([
            ("device_no", "设备编号 (Device No.):"),
            ("api_domain", "API 域名 (API Domain):"),
            ("api_port", "API 端口 (API Port):"),
            ("user_id", "用户ID (User ID):"),
            ("security_key", "安全密钥 (Security Key):")
        ]):
            tb.Label(settings_frame, text=label, font=(self.font_family, 12)).grid(
                row=i, column=0, sticky=tk.W, pady=5, padx=5)
            entry = tb.Entry(settings_frame, font=(self.font_family, 12), 
                            width=50, show='*' if key == "security_key" else '')
            entry.grid(row=i, column=1, sticky=tk.EW, pady=5, padx=5)
            entry.insert(0, self.settings[key])
            entries[key] = entry
        settings_frame.grid_columnconfigure(1, weight=1)
        
        # 添加串口设置部分
        # 创建自定义样式并配置字体
        style_name = "SerialSettings.TLabelframe"
        self.style.configure(
            f"{style_name}.Label", 
            font=(self.font_family, 12)
        )
        serial_frame = tb.Labelframe(
            settings_window, 
            text="串口设置 (Serial Port)",
            bootstyle="primary",
            style=style_name  # 应用自定义样式
        )
        serial_frame.pack(pady=10, padx=20, fill=tk.X)
        
        # 获取可用串口列表
        available_ports = []
        if self.serial_manager:
            available_ports = self.serial_manager.list_available_ports()
        
        # 串口下拉框和输入框
        port_frame = tb.Frame(serial_frame)
        port_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tb.Label(port_frame, text="串口 (Port):", font=(self.font_family, 12)).grid(
            row=0, column=0, sticky=tk.W, padx=5)
        
        self.port_var = tk.StringVar()
        if self.serial_manager and self.serial_manager.port:
            self.port_var.set(self.serial_manager.port)
        elif available_ports:
            self.port_var.set(available_ports[0])
            
        port_combo = tb.Combobox(port_frame, textvariable=self.port_var, 
                                values=available_ports, width=20)
        port_combo.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # 添加手动输入框
        tb.Label(port_frame, text="或手动输入:", font=(self.font_family, 12)).grid(
            row=0, column=2, sticky=tk.W, padx=5)
        
        port_entry = tb.Entry(port_frame, font=(self.font_family, 12), width=15)
        port_entry.grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # 刷新按钮
        refresh_button = tb.Button(port_frame, text="刷新列表", 
                                   command=lambda: self.refresh_port_list(port_combo),
                                   bootstyle="secondary-outline")
        refresh_button.grid(row=0, column=4, padx=5)
        
        port_frame.grid_columnconfigure(1, weight=1)

        save_button = tb.Button(settings_window, text="保存 (Save)", 
                   command=lambda: self.save_settings_from_window(
                       *[entries[key].get().strip() for key in entries],
                       self.port_var.get() if not port_entry.get().strip() else port_entry.get().strip(),
                       settings_window),
                   bootstyle="primary")
        save_button.pack(pady=10)

    def save_settings_from_window(self, device_no: str, api_domain: str, api_port: str, 
                                 user_id: str, security_key: str, port: str, window: tk.Toplevel) -> None:
        """保存设置并更新配置。"""
        if not all([device_no, api_domain, api_port, user_id, security_key]):
            self.update_status_display("设置错误: 所有字段均为必填项")
            return
        try:
            int(api_port)
        except ValueError:
            self.update_status_display("设置错误: 端口号必须是数字")
            return

        self.settings = {
            "device_no": device_no, "api_domain": api_domain, "api_port": api_port,
            "user_id": user_id, "security_key": security_key
        }
        self.api_base_url = f"http://{api_domain}:{api_port}/api/device"
        save_settings(self.settings)
        
        # 更新串口设置
        if port and self.serial_manager:
            self.serial_manager.set_port(port)
            
        self.update_status_display("设置已保存 (Settings Saved)")
        window.destroy()

    def _generate_signature(self, user_id: str, security_key: str) -> str:
        """生成 API 的 MD5 签名。"""
        # 使用随机盐值替代特定域名，增强安全性
        salt = hashlib.sha256(security_key.encode()).hexdigest()[:16]
        data = user_id + security_key.upper() + salt
        return hashlib.md5(data.encode('utf-8')).hexdigest()

    def update_weight_display(self, weight_str: str) -> None:
        """更新重量显示。"""
        self.weight_label.config(text=weight_str)

    def update_status_display(self, status_str: str) -> None:
        """更新状态显示。"""
        self.status_label.config(text=f"状态: {status_str}")
        self.toggle_button.config(text="断开 (Disconnect)" if "connected" in status_str else "连接 (Connect)")
        # 更新上传按钮状态
        if "connected" in status_str:
            self.upload_button.config(state=tk.NORMAL, bootstyle="success")
        else:
            self.upload_button.config(state=tk.DISABLED)
        self.connected = "connected" in status_str

    def update_upload_status_display(self, status_str: str, is_error: bool = False) -> None:
        """更新上传状态显示。"""
        self.upload_status_label.config(text=f"上传状态: {status_str}", fg="red" if is_error else "#333333")

    def toggle_connection(self) -> None:
        """切换串口连接状态。"""
        if not self.connected:
            self.connect_serial()
        else:
            self.disconnect_serial()

    def connect_serial(self) -> None:
        """启动串口连接线程。"""
        self.toggle_button.config(state=tk.DISABLED)
        self.upload_button.config(state=tk.DISABLED)
        threading.Thread(target=self.connect_serial_thread, daemon=True).start()

    def connect_serial_thread(self) -> None:
        """执行串口连接操作。"""
        if self.serial_manager is None:
            if serial is None:
                self.master.after(0, self.update_status_display, "错误: pyserial库未安装")
                self.master.after(0, lambda: self.show_no_serial_warning("pyserial库未安装"))
                self.master.after(0, self.toggle_button.config, {'style': "Connect.TButton"})
                return
            self.serial_manager = SerialConnectionManager(
                data_callback=self.process_received_data,
                status_callback=lambda s: self.master.after(0, self.update_status_display, s)
            )
        
        self.master.after(0, self.update_status_display, "正在连接...")
        if not self.serial_manager.port:
            self.master.after(0, self.update_status_display, "连接失败：无可用串口")
            self.master.after(0, lambda: self.show_no_serial_warning("无可用串口"))
            self.master.after(0, self.toggle_button.config, {'style': "Connect.TButton"})
            return

        if not self.serial_manager.connect():
            self.master.after(0, self.update_status_display, "连接失败")
            self.master.after(0, self.update_weight_display, "--- kg")
            self.master.after(0, self.toggle_button.config, {'style': "Connect.TButton"})

    def disconnect_serial(self) -> None:
        """断开串口连接。"""
        if self.serial_manager and self.connected:
            self.serial_manager.close()
            self.connected = False
            self.master.after(0, self.update_status_display, "已断开 (Disconnected)")
            self.master.after(0, self.update_weight_display, "0.00 kg")
            self.master.after(0, lambda: self.update_upload_status_display("等待操作", False))
            self.master.after(0, self.toggle_button.config, {'style': "Connect.TButton"})
            self.upload_button.config(state=tk.DISABLED)
            self.raw_weight_value = 0.0
            self.last_weight = None
            self.last_weight_time = None

    def initiate_upload_weight_thread(self) -> None:
        """启动上传重量线程。"""
        self.upload_button.config(state=tk.DISABLED)
        self.master.after(0, lambda: self.update_upload_status_display("正在上传...", False))
        threading.Thread(target=self.upload_weight_to_its, daemon=True).start()

    def upload_weight_to_its(self) -> None:
        """上传重量和尺寸到 ITS 系统。"""
        device_no = self.settings["device_no"]
        scan_no = self.scan_no_entry.get().strip()
        weight = self.raw_weight_value
        length = self.length_entry.get().strip()
        width = self.width_entry.get().strip()
        height = self.height_entry.get().strip()
        log_base = f"单号:{scan_no} 重量:{weight:.3f}kg 长:{length or '-'} 宽:{width or '-'} 高:{height or '-'}"

        if not device_no or not scan_no or weight is None or weight <= 0:
            error_msg = (
                "错误: 设备编号未配置" if not device_no else
                "错误: 请输入扫描单号" if not scan_no else
                "错误: 重量无效"
            )
            self.master.after(0, lambda: self.update_upload_status_display(error_msg, True))
            self.master.after(0, lambda: self.update_upload_log(f"{log_base} 状态:{error_msg}"))
            if sys.platform == 'win32':
                if sys.platform == 'win32' and winsound is not None:
                    winsound.MessageBeep(winsound.MB_ICONHAND)
            self.master.after(0, lambda: self.upload_button.config(style="Upload.TButton"))
            return

        try:
            length = float(length) if length else None
            width = float(width) if width else None
            height = float(height) if height else None
        except ValueError:
            length = width = height = None

        api_url = f"{self.api_base_url}/inWarehouse"
        user_id = self.settings["user_id"]
        if user_id == "USER_ID" or self.settings["security_key"] == "SECURITY_KEY":
            error_msg = "错误: 请配置API信息"
            self.master.after(0, lambda: self.update_upload_status_display(error_msg, True))
            self.master.after(0, lambda: self.update_upload_log(f"{log_base} 状态:{error_msg}"))
            if sys.platform == 'win32':
                if sys.platform == 'win32' and winsound is not None:
                    winsound.MessageBeep(winsound.MB_ICONHAND)
            self.master.after(0, lambda: self.upload_button.config(style="Upload.TButton"))
            return

        headers = {
            "userId": user_id,
            "signature": self._generate_signature(user_id, self.settings["security_key"]),
            "Content-Type": "application/json"
        }
        payload = {
            "deviceNo": device_no, "scanNo": scan_no, "weight": float(f"{weight:.3f}"),
            "length": length, "width": width, "height": height, "pictureBase64": ""
        }

        try:
            if requests is None:
                raise ImportError("requests库未安装，无法进行上传")
            response = requests.post(api_url, headers=headers, json=payload, timeout=10)
            response.raise_for_status()
            response_data = response.json()
            if response_data.get("code") == 0:
                success_msg = f"上传成功! 单号: {response_data.get('data', {}).get('scanNo')}"
                self.master.after(0, lambda: self.update_upload_status_display(success_msg, False))
                self.master.after(0, lambda: self.update_upload_log(f"{log_base} 状态:上传成功"))
                if sys.platform == 'win32' and winsound is not None:
                    winsound.MessageBeep(winsound.MB_ICONASTERISK)
            else:
                error_msg = f"上传失败! 代码: {response_data.get('code')}, 消息: {response_data.get('msg')}"
                self.master.after(0, lambda: self.update_upload_status_display(error_msg, True))
                self.master.after(0, lambda: self.update_upload_log(f"{log_base} 状态:失败 ({response_data.get('msg')})"))
                if sys.platform == 'win32' and winsound is not None:
                    winsound.MessageBeep(winsound.MB_ICONASTERISK if '已入库' in response_data.get('msg', '') 
                                       else winsound.MB_ICONHAND)
        except ImportError as e:
            error_msg = f"上传错误: {str(e)}"
            self.master.after(0, lambda: self.update_upload_status_display(error_msg, True))
            self.master.after(0, lambda: self.update_upload_log(f"{log_base} 状态:{error_msg}"))
            if sys.platform == 'win32' and winsound is not None:
                winsound.MessageBeep(winsound.MB_ICONHAND)
        except Exception as e:
            if requests is not None and isinstance(e, requests.exceptions.RequestException):
                error_msg = f"网络错误: {str(e)}"
                self.master.after(0, lambda: self.update_upload_status_display(error_msg, True))
                self.master.after(0, lambda: self.update_upload_log(f"{log_base} 状态:{error_msg}"))
                if sys.platform == 'win32' and winsound is not None:
                    winsound.MessageBeep(winsound.MB_ICONHAND)
        finally:
            self.master.after(0, lambda: self.upload_button.config(style="Upload.TButton"))
            self.master.after(0, lambda: self.scan_no_entry.delete(0, tk.END))
            self.master.after(0, lambda: self.length_entry.delete(0, tk.END))
            self.master.after(0, lambda: self.width_entry.delete(0, tk.END))
            self.master.after(0, lambda: self.height_entry.delete(0, tk.END))
            self.master.after(0, lambda: self.scan_no_entry.focus_set())

    def show_no_serial_warning(self, reason: str) -> None:
        """显示无串口警告对话框。"""
        if messagebox is not None:
            messagebox.showwarning("连接失败", f"无法连接到电子称：{reason}\n\n您可以：\n1. 检查串口连接\n2. 在设置中指定串口\n3. 点击退出按钮关闭程序")
        else:
            print(f"警告: 无法连接到电子称：{reason}")

    def show_help(self) -> None:
        """显示帮助文档窗口。"""
        help_window = tb.Toplevel(self.master)
        help_window.title("电子称重量显示与上传系统 - 帮助文档")
        help_window.geometry("800x800")
        
        # 创建带滚动条的文本框
        help_frame = tb.Frame(help_window)
        help_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        help_text = tb.Text(help_frame, font=(self.font_family, 12), wrap=tk.WORD)
        scrollbar = tb.Scrollbar(help_frame, command=help_text.yview)
        help_text.config(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 帮助文档内容
        help_content = """电子称重量显示与上传系统
====================

功能概述
--------
本系统主要用于连接电子秤设备，读取重量数据，并将重量数据上传至ITS系统。

主要功能：
1. 自动连接电子秤串口设备
2. 实时显示重量数据
3. 记录包裹尺寸信息
4. 上传重量和尺寸数据到远程服务器
5. 显示最近上传记录

设备连接
--------
1. 连接电子秤：将电子秤通过USB线连接到电脑
2. 启动软件：软件会自动尝试连接电子秤
3. 手动连接：点击"连接"按钮可手动连接设备
4. 连接状态：状态指示灯显示当前连接状态
   - 灰色：未连接
   - 黄色：连接但数据异常
   - 蓝色：连接正常，当前重量为零
   - 绿色：连接正常，有物品在秤上
   - 红色：连接异常

称重与上传
----------
1. 称重步骤：
   - 确保设备已连接（状态指示灯为蓝色或绿色）
   - 将包裹放置在电子秤上
   - 等待重量稳定（指示灯变为绿色）
   
2. 上传数据：
   - 在"扫描单号"输入框中输入包裹单号
   - 可选：输入包裹的长、宽、高尺寸
   - 点击"上传重量"按钮或按回车键上传数据
   - 系统会显示上传状态和结果

系统设置
--------
点击"设置"菜单访问系统配置：
1. 设备编号：电子秤的唯一标识
2. API域名和端口：远程服务器地址
3. 用户ID和安全密钥：API访问凭证
4. 串口设置：手动选择或输入电子秤串口

注意事项
--------
1. 确保电子秤正确连接并通电
2. 上传前确认重量已稳定显示
3. 单号必须输入，否则无法上传
4. 如遇连接问题，可在设置中手动选择串口
5. 尺寸数据为可选项，不影响重量上传

常见问题
--------
Q: 无法连接电子秤？
A: 检查USB连接，重启软件，或在设置中手动选择正确的COM端口

Q: 重量显示不稳定？
A: 确保电子秤放置在平稳表面，避免振动

Q: 上传数据失败？
A: 检查网络连接，确认API配置正确，查看错误提示

Q: 如何退出程序？
A: 点击"退出"菜单或窗口右上角的关闭按钮"""
        
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)  # 设置为只读
        
        # 关闭按钮
        close_button = tb.Button(help_window, text="关闭", command=help_window.destroy, 
                                bootstyle="secondary")
        close_button.pack(pady=10)
    
    def on_closing(self) -> None:
        """处理窗口关闭事件。"""
        try:
            if self.connected:
                self.disconnect_serial()
            if self.serial_manager and serial is not None:
                self.serial_manager.close()
            if self.stability_check_id:
                self.master.after_cancel(self.stability_check_id)
        except Exception as e:
            print(f"关闭时发生错误: {e}")
        finally:
            self.master.destroy()

if __name__ == "__main__":
    # 检查serial库是否可用
    if serial is None:
        error_msg = "错误: 请安装 pyserial 库 (pip install pyserial)"
        print(error_msg)
        logging.error("pyserial 库缺失，程序无法正常工作")
        if messagebox is not None:
            messagebox.showerror("缺少必要库", 
                "未检测到 pyserial 库，秤重功能将无法使用。\n\n"
                "请安装 pyserial:\npip install pyserial\n\n"
                "点击确定退出程序。")
        sys.exit(1)  # 退出程序
    
    # 使用ttkbootstrap创建主窗口
    root = tb.Window(title="电子秤重量显示与上传", themename="flatly")
    app = WeightDisplayApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()