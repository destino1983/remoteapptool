import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
import subprocess
import json

class RemoteApp:
    def __init__(self, name, default_ip=""):
        self.name = name
        self.default_ip = default_ip

    def to_dict(self):
        return {"name": self.name, "default_ip": self.default_ip}

    @classmethod
    def from_dict(cls, data):
        return cls(data["name"], data["default_ip"])

class RDPWindow:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("RDP连接")
        self.remote_apps = []
        self.selected_app = None  # 用于存储当前选中的应用程序

        # 添加“创建应用程序”按钮
        self.create_button = ttk.Button(self.window, text="创建应用程序", command=self.create_remote_app)
        self.create_button.pack(side="top", fill="x", pady=10, padx=10)

        # 用于存储应用程序标签和连接按钮的框架
        self.app_frames = []

        # 使用ttk风格
        self.style = ttk.Style()
        # 尝试使用不同主题
        self.style.theme_use("clam")  # 你可以尝试其他主题，如"alt", "default", "classic", "vista", "xpnative"

        # 在窗口底部添加注释
        comment_text = ("注意: 创建应用程序时，请输入应用程序名称和IP地址，"
                        "用逗号分隔。例如：App1,192.168.1.1")
        self.comment_label = tk.Label(self.window, text=comment_text, wraplength=380, justify="left")
        self.comment_label.pack(side="bottom", pady=10, padx=10)

        # 加载应用程序信息
        self.load_apps()

    def save_apps(self):
        # 将应用程序信息保存到文件
        with open("apps_data.json", "w") as json_file:
            app_data = [app.to_dict() for app in self.remote_apps]
            json.dump(app_data, json_file)

    def load_apps(self):
        # 从文件加载应用程序信息
        try:
            with open("apps_data.json", "r") as json_file:
                app_data = json.load(json_file)
                self.remote_apps = [RemoteApp.from_dict(data) for data in app_data]
                self.update_labels()
        except FileNotFoundError:
            # 如果文件不存在，创建默认的应用程序实例
            app1 = RemoteApp("App1", default_ip="192.168.1.1")
            app2 = RemoteApp("App2", default_ip="192.168.1.2")
            self.remote_apps = [app1, app2]
            self.save_apps()  # 保存应用程序信息
            self.update_labels()

    def add_remote_app(self, remote_app):
        self.remote_apps.append(remote_app)
        self.update_labels()
        self.save_apps()  # 保存应用程序信息

    def create_remote_app(self):
        # 弹出对话框，输入应用程序信息
        input_window = tk.Toplevel(self.window)
        input_window.title("创建应用程序")

        # 为应用程序名称输入创建输入框
        name_label = ttk.Label(input_window, text="应用程序名称:")
        name_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        name_entry = ttk.Entry(input_window)
        name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        # 为IP地址输入创建输入框
        ip_label = ttk.Label(input_window, text="IP地址:")
        ip_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ip_entry = ttk.Entry(input_window)
        ip_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 创建一个保存信息的按钮
        save_button = ttk.Button(input_window, text="保存", command=lambda: self.on_create_save_button_click(name_entry.get(), ip_entry.get(), input_window))
        save_button.grid(row=2, columnspan=2, pady=10)

    def on_create_save_button_click(self, name, ip, input_window):
        if name and ip:
            new_app = RemoteApp(name, default_ip=ip)
            self.add_remote_app(new_app)
            input_window.destroy()
        else:
            messagebox.showerror("错误", "请输入应用程序名称和IP地址。")

    def connect_to_remote(self, app, ip_address, app_name):
        # 连接到远程主机的逻辑，这里使用mstsc.exe作为示例
        rdp_params = f"""
        alternate full address:s:{ip_address}
        alternate shell:s:rdpinit.exe
        full address:s:{ip_address}
        remoteapplicationmode:i:1
        remoteapplicationname:s:{app_name}
        remoteapplicationprogram:s:||{app_name}
        disableremoteappcapscheck:i:1
        drivestoredirect:s:*
        prompt for credentials:i:1
        promptcredentialonce:i:0
        redirectcomports:i:1
        span monitors:i:1
        use multimon:i:1
        """

        # 将RDP参数保存到文件
        with open("rdp_settings.rdp", "w") as rdp_file:
            rdp_file.write(rdp_params)

        # 使用保存的参数运行远程桌面连接工具
        subprocess.run(["mstsc.exe", "rdp_settings.rdp"], shell=True)

    def edit_connected_info(self):
        # 编辑已连接信息
        if self.selected_app:
            result = simpledialog.askstring("编辑已连接信息", "输入IP地址和应用程序名称（用逗号分隔）:",
                                             initialvalue=f"{self.selected_app.name},{self.selected_app.default_ip}")

            if result:
                try:
                    name, ip = result.split(',')
                    self.selected_app.name = name
                    self.selected_app.default_ip = ip
                    self.update_labels()
                    self.save_apps()  # 保存应用程序信息
                except ValueError:
                    messagebox.showerror("错误", "输入格式错误，请使用逗号分隔应用程序名称和IP地址。")

    def delete_remote_app(self):
        # 删除远程应用程序
        if self.selected_app in self.remote_apps:
            self.remote_apps.remove(self.selected_app)
            self.update_labels()
            self.save_apps()  # 保存应用程序信息

    def on_label_click(self, app):
        # 打开一个新窗口用于输入IP和应用程序名称
        input_window = tk.Toplevel(self.window)
        input_window.title(f"输入IP和应用程序名称 - {app.name}")

        # 为IP地址和远程应用程序名称输入创建两个输入框
        name_label = ttk.Label(input_window, text="应用程序名称:")
        name_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        name_entry = ttk.Entry(input_window)
        name_entry.insert(0, app.name)
        name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        ip_label = ttk.Label(input_window, text="IP地址:")
        ip_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ip_entry = ttk.Entry(input_window)
        ip_entry.insert(0, app.default_ip)
        ip_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 创建一个保存信息的按钮
        save_button = ttk.Button(input_window, text="保存", command=lambda: self.on_save_button_click(app, name_entry.get(), ip_entry.get(), input_window))
        save_button.grid(row=2, columnspan=2, pady=10)

    def on_save_button_click(self, app, name, ip, input_window):
        app.name = name
        app.default_ip = ip
        self.update_labels()
        input_window.destroy()
        self.save_apps()  # 保存应用程序信息

    def on_label_right_click(self, app, label_widget, event):
        # 右键菜单
        self.selected_app = app
        self.popup_menu = tk.Menu(self.window, tearoff=0)
        self.popup_menu.add_command(label="编辑", command=self.edit_connected_info)
        self.popup_menu.add_command(label="删除", command=self.delete_remote_app)
        self.popup_menu.post(event.x_root, event.y_root)

    def on_connect_button_click(self, app):
        # 连接按钮的点击事件
        self.connect_to_remote(app, app.default_ip, app.name)

    def update_labels(self):
        # 清除现有的所有标签和按钮
        for frame in self.app_frames:
            frame.destroy()

        # 创建新的标签和按钮
        for app in self.remote_apps:
            self.create_app_frame(app)

    def create_app_frame(self, app):
        # 创建一个框架包含应用程序标签和连接按钮
        frame = ttk.Frame(self.window, relief="solid", borderwidth=1)
        frame.pack(side="top", fill="x", pady=5, padx=10)  # 使用pack布局

        # 创建连接按钮
        connect_button = ttk.Button(frame, text="连接", command=lambda app=app: self.on_connect_button_click(app))
        connect_button.pack(side="left", padx=5, pady=5)

        # 创建应用程序标签
        label_text = f"{app.name} - {app.default_ip}"
        label = ttk.Label(frame, text=label_text, cursor="hand2")
        label.pack(side="left", padx=5, pady=5)

        # 绑定单击事件到标签
        label.bind("<Button-1>", lambda event, app=app: self.on_label_click(app))

        # 绑定右键事件到标签
        label.bind("<Button-3>", lambda event, app=app, label_widget=label: self.on_label_right_click(app, label_widget, event))

        # 存储应用程序实例引用在标签中，以便右键菜单能够访问
        label.app = app

        self.app_frames.append(frame)

# 创建RDP窗口实例
rdp_window = RDPWindow()

# 计算窗口居中的坐标
screen_width = rdp_window.window.winfo_screenwidth()
screen_height = rdp_window.window.winfo_screenheight()
x_coordinate = (screen_width - 400) // 2
y_coordinate = (screen_height - 300) // 2
rdp_window.window.geometry(f"400x300+{x_coordinate}+{y_coordinate}")

# 运行tkinter事件循环
rdp_window.window.mainloop()