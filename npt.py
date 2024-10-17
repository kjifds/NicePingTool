import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, font
import subprocess
import threading
import os
from datetime import datetime
from PIL import ImageGrab


def format_duration(seconds):
    minutes, seconds = divmod(seconds, 60)
    if minutes > 0:
        return f"{minutes}分{seconds}秒"
    return f"{seconds}秒"


class NetworkMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("断网检测工具_v1.0.2")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")  # 设置背景色

        # 初始化状态和变量
        self.process = None
        self.running = False
        self.output_log = []
        self.start_time = None
        self.disconnection_start = None
        self.disconnection_count = 0
        self.target_address = "www.baidu.com"
        self.screenshot_enabled = tk.BooleanVar()
        self.save_log_enabled = tk.BooleanVar()

        # 创建控件
        self.create_widgets()

    def create_widgets(self):
        # 定义字体
        title_font = font.Font(family="微软雅黑", size=16, weight="bold")
        label_font = font.Font(family="微软雅黑", size=12)

        # Ping目标地址输入框及其按钮
        address_frame = ttk.Frame(self.root, padding="10")
        address_frame.pack(pady=5)

        ttk.Label(address_frame, text="Ping 目标地址:", font=label_font).pack(
            side=tk.LEFT, padx=5
        )
        self.address_entry = ttk.Entry(address_frame, width=30)
        self.address_entry.pack(side=tk.LEFT, padx=5)
        self.address_entry.insert(0, self.target_address)

        self.start_button = ttk.Button(
            address_frame, text="开始", command=self.start_monitoring, width=8
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        self.stop_button = ttk.Button(
            address_frame,
            text="停止",
            command=self.stop_monitoring,
            state=tk.DISABLED,
            width=8,
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # 是否启用截图和保存日志复选框
        options_frame = ttk.Frame(self.root, padding="10")
        options_frame.pack(pady=5)

        self.screenshot_checkbox = ttk.Checkbutton(
            options_frame, text="启用断网截图", variable=self.screenshot_enabled
        )
        self.screenshot_checkbox.pack(side=tk.LEFT, padx=5)

        self.save_log_checkbox = ttk.Checkbutton(
            options_frame, text="启用保存日志", variable=self.save_log_enabled
        )
        self.save_log_checkbox.pack(side=tk.LEFT, padx=5)

        # 输出框，显示实时ping命令结果
        self.output_text = scrolledtext.ScrolledText(
            self.root, height=15, width=90, bg="#000000", fg="#ffffff", font=label_font
        )
        self.output_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 创建红色加粗文本标签
        self.output_text.tag_config(
            "red_bold", foreground="red", font=("微软雅黑", 12, "bold")
        )

        # 断网间隔记录框
        self.interval_text = scrolledtext.ScrolledText(
            self.root, height=10, width=90, bg="#000000", fg="#ffffff", font=label_font
        )
        self.interval_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.interval_text.config(state=tk.DISABLED)  # 禁止用户输入

    def start_monitoring(self):
        self.target_address = self.address_entry.get()
        self.start_button.config(state=tk.DISABLED)  # 禁用“开始”按钮
        self.stop_button.config(state=tk.NORMAL)  # 启用“停止”按钮
        self.running = True
        self.output_log = []  # 清空日志
        self.start_time = datetime.now()  # 记录开始时间

        threading.Thread(target=self.run_ping_command).start()

    def stop_monitoring(self):
        self.running = False
        if self.process:
            self.process.terminate()  # 停止PowerShell进程

        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        messagebox.showinfo("结束", "监控已停止。")

        if self.save_log_enabled.get():
            self.save_log()  # 保存日志

    def run_ping_command(self):
        command = f'ping.exe -t {self.target_address} | ForEach {{ "{{0}} - {{1}}" -f (Get-Date),$_ }}'

        # 创建子进程时隐藏窗口
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        self.process = subprocess.Popen(
            ["powershell", "-Command", command],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            startupinfo=startupinfo,  # 隐藏PowerShell窗口
            creationflags=subprocess.CREATE_NO_WINDOW,  # 防止创建新窗口
        )

        while self.running:
            line = self.process.stdout.readline().strip()  # 获取实时ping输出
            if not line:
                continue

            # 检查是否是断网信息并设置对应的样式
            if "请求超时" in line or "无法访问目标主机" in line or "常见故障" in line:
                self.output_text.config(state=tk.NORMAL)  # 允许写入输出框
                self.output_text.insert(
                    tk.END, line + "\n", "red_bold"
                )  # 使用红色加粗标签
                self.output_text.see(tk.END)  # 自动滚动到底部
                self.output_text.config(state=tk.DISABLED)  # 禁止用户输入
            else:
                self.output_text.config(state=tk.NORMAL)  # 允许写入输出框
                self.output_text.insert(tk.END, line + "\n")  # 输出到文本框
                self.output_text.see(tk.END)  # 自动滚动到底部
                self.output_text.config(state=tk.DISABLED)  # 禁止用户输入

            self.output_log.append(line)  # 记录输出日志

            self.check_disconnection(line)  # 检查断网状态

    def check_disconnection(self, line):
        if (
            "请求超时" in line or "无法访问目标主机" in line or "常见故障" in line
        ):  # 检测到断网
            if not self.disconnection_start:  # 如果没有记录开始时间，则设置
                self.disconnection_start = datetime.now()
                self.disconnection_count += 1  # 增加断网次数
                self.log_disconnection_start()  # 记录断网开始时间
                if self.screenshot_enabled.get():
                    self.capture_screenshot()  # 执行断网截图功能
        else:  # 网络恢复
            if self.disconnection_start:  # 如果有开始时间，说明网络恢复
                self.log_disconnection_end()  # 记录断网结束时间
                self.disconnection_start = None  # 重置开始时间

    def log_disconnection_start(self):
        self.interval_text.config(state=tk.NORMAL)  # 允许写入间隔记录框
        self.interval_text.insert(
            tk.END,
            f"断网开始时间: {self.disconnection_start.strftime('%Y-%m-%d %H:%M:%S')}\n",
        )
        self.interval_text.config(state=tk.DISABLED)  # 禁止用户输入

    def log_disconnection_end(self):
        disconnection_end = datetime.now()  # 记录结束时间
        self.interval_text.config(state=tk.NORMAL)  # 允许写入间隔记录框
        self.interval_text.insert(
            tk.END,
            f"断网结束时间: {disconnection_end.strftime('%Y-%m-%d %H:%M:%S')}\n",
        )
        interval = disconnection_end - self.disconnection_start
        interval_seconds = interval.total_seconds()  # 转换为秒
        formatted_interval = format_duration(int(interval_seconds))  # 格式化为中文
        self.interval_text.insert(tk.END, f"断网间隔: {formatted_interval}\n")
        self.interval_text.insert(
            tk.END, f"断网次数: {self.disconnection_count}\n\n"
        )  # 显示断网次数
        self.interval_text.config(state=tk.DISABLED)  # 禁止用户输入

    def save_log(self):
        log_filename = f"长PING日志-{self.start_time.strftime('%Y%m%d-%H%M%S')}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.txt"
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        log_path = os.path.join(desktop_path, log_filename)

        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write(
                f"开始时间: {self.start_time}\n结束时间: {datetime.now()}\n\n"
            )
            for line in self.output_log:
                log_file.write(line + "\n")

        messagebox.showinfo("日志保存", f"日志已保存到: {log_path}")

    def capture_screenshot(self):
        screenshot_time = datetime.now().strftime("%Y%m%d-%H%M%S")
        screenshot_folder = os.path.join(
            os.path.expanduser("~"), "Desktop", "断网截图文件夹"
        )
        os.makedirs(screenshot_folder, exist_ok=True)

        screenshot_path = os.path.join(
            screenshot_folder, f"断网截图-{screenshot_time}.png"
        )
        screenshot = ImageGrab.grab()  # 获取屏幕截图
        screenshot.save(screenshot_path)  # 保存截图

        messagebox.showinfo("截图保存", f"截图已保存到: {screenshot_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = NetworkMonitorApp(root)
    root.mainloop()
