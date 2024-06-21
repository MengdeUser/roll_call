# coding=utf-8
"""
三班出品，必属精品。版权归编者所有；
有意者加wx：Z18241291013或者QQ：1651473590；
备注说明来意。
"""
import os
import subprocess
import sys
import time
import json
import random
import win32gui
import win32con
import requests
import pyautogui
import webbrowser
import win32com.client
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox


__author__ = "梦与拾光遇"
# 软件的当前版本
current_version = "v0.1.8"  # 请替换为实际的当前版本号
new = 'https://www.123pan.com/s/jRAxjv-iGBWA.html'
web = 'https://mengdeuser.github.io/roll_call/'
hub = 'https://github.com/MengdeUser/roll_call'
names = ["李宇轩", "杨政皓", "高逸航", "高子航", "石博宇", "张高菲", "陈赞彭", "杨凯琪", "杜雨嘉诺", "马欣妍"]
kfz = 0

# 打开文件并读取内容
with open('_internal/res/students.json', 'r', encoding='utf-8') as file:
    names_list = json.load(file)


def set_window_transparency(window_handle, transparency):
    # 透明度
    win32gui.SetWindowLong(window_handle, win32con.GWL_EXSTYLE,
                           win32gui.GetWindowLong(window_handle, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
    win32gui.SetLayeredWindowAttributes(window_handle, 0, transparency * 255 // 100, win32con.LWA_ALPHA)


def setup():
    # GitHub仓库URL
    GITHUB_REPO_URL = "https://api.github.com/repos/MengdeUser/roll_call/releases/latest"
    # 本地软件版本
    LOCAL_VERSION = current_version  # 请替换为你的软件当前版本
    # 本地软件安装路径
    INSTALL_PATH = os.path.join(os.path.dirname(__file__), "setup.exe")  # 请替换为你的软件安装路径

    def get_latest_version():
        """从GitHub获取最新版本信息"""
        try:
            response = requests.get(GITHUB_REPO_URL)
            response.raise_for_status()
            data = response.json()
            return data['tag_name']
        except requests.exceptions.RequestException as e:
            messagebox.showinfo("提示", f"获取最新版本时出错: {e}")
            return None

    def download_file(progress_bar, progress_label, latest_version, root):
        """下载并安装最新版本"""
        # 这里需要根据实际情况编写下载和安装的代码
        # 以下代码仅为示例
        # 假设下载的安装包是setup.exe
        # 下载命令
        download_command = f"https://github.com/MengdeUser/roll_call/releases/download/{latest_version}/setup.exe"
        # 执行下载
        with requests.get(download_command, stream=True) as r:
            r.raise_for_status()
            total_length = int(r.headers.get('content-length', 1))
            with open(INSTALL_PATH, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:  # 过滤掉保持连接的chunk
                        f.write(chunk)
                        downloaded = f.tell()
                        # 更新进度条
                        progress_bar['value'] = (downloaded / total_length) * 100
                        progress_label['text'] = f"{downloaded / total_length * 100:.2f}%"
                        root.update_idletasks()  # 更新GUI
        root.destroy()
        wait = messagebox.askyesno("提示", "您是否需要启用新版本安装程序？")
        if wait:
            # 运行安装后的软件
            subprocess.run(f"{INSTALL_PATH}", shell=True)
        else:
            sys.exit(0)

    def main1(latest_version):
        root = tk.Tk()
        root.title("下载进度")
        root.iconbitmap('_internal\\res\\favicon.ico')
        root.attributes("-topmost", True)
        root.overrideredirect(True)
        width = 400
        height = 120
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
        root.resizable(False, False)

        showinfo = Label(text=f"下载和安装版本{latest_version}...", font=("汉仪文黑-85W", 15))
        showinfo.pack(pady=5)
        # 创建进度条
        progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
        progress_bar.pack(pady=10)
        # 创建进度标签
        progress_label = tk.Label(root, text="0%")
        progress_label.pack(pady=5)

        latest_version = get_latest_version()
        if latest_version > LOCAL_VERSION:
            download_file(progress_bar, progress_label, latest_version, root)
        else:
            # 这里添加预运行的代码
            root.destroy()
            main()

        root.mainloop()
    main1(latest_version=get_latest_version())


def main():
    # 定义一个标志文件的路径
    flag_file_path = '_internal/res/first_run_flag.txt'

    def is_first_run():
        # 检查标志文件是否存在
        if not os.path.exists(flag_file_path):
            # 如果文件不存在，创建文件并返回True表示是第一次运行
            with open(flag_file_path, 'w') as file:
                file.write('此程序已运行。')
            return True
        else:
            # 如果文件存在，返回False表示不是第一次运行
            return False

    # 使用函数
    if is_first_run():
        # 须知
        result = messagebox.askokcancel('用户须知', f'点名软件用户须知\n\n'
                                                f'一、引言欢迎您使用点名软件！\n为了确保您能够顺利、安全地使用我们的软件，并充分了解相关权益和义务，我们特此制定了这份用户须知。'
                                                f'\n请您仔细阅读并遵守以下规定，以保障您的权益并享受优质的软件服务。\n\n'
                                                f'二、软件使用授权使用：\n本软件仅供授权用户使用。未经授权，禁止复制、传播、出售或用于其他商业目的。'
                                                f'\n合规性：用户在使用软件时，应遵守国家法律法规，不得利用软件进行违法、违规活动。'
                                                f'\n软件更新：为了提供更好的服务和修复潜在问题，我们会不定期更新软件。请确保您的设备具有浏览器，以便接收更新。\n\n'
                                                f'三、内容与安全用户内容：\n用户在使用软件过程中产生的内容（如文件、数据、通信等）应遵守法律法规，不得包含违法、不良信息。\n\n'
                                                f'四、软件功能与服务功能使用：\n用户应按照软件提供的功能和界面指引使用，不得进行非法操作或破坏软件正常运行。'
                                                f'\n服务支持：我们会提供必要的技术支持和客户服务，但某些情况下可能需要用户自行解决问题或寻求第三方帮助。\n\n'
                                                f'五、责任与限制用户责任：\n用户应对自己的行为负责，如因使用软件导致的任何纠纷或损失，用户应自行承担相应责任。'
                                                f'\n服务限制：在某些情况下，我们可能会限制或终止用户的访问权限，包括但不限于违反规定、滥用服务或侵犯他人权益等。\n\n'
                                                f'六、法律声明与合规性法律管辖：\n本用户须知受中华人民共和国法律管辖。如有任何争议，应提交至有管辖权的人民法院解决。'
                                                f'\n合规性：我们将严格遵守相关法律法规，并努力保障用户合法权益。\n\n'
                                                f'七、接受与同意请您务必认真阅读并完全理解本用户须知的所有内容。\n\n'
                                                f'如您接受并同意遵守本用户须知的所有条款和条件，请点击确认按钮，继续使用点名软件。\n\n'
                                                f'感谢您的配合与支持，祝您使用愉快\n\n'
                                                f'三班出品，必属精品。\n版权归编者所有；\n有意者加wx：Z18241291013或者QQ：1651473590；\n备注说明来意。'
                                                f'建议与反馈\nQQ群：497499805')
        if result:
            go()
        else:
            # 指定要删除的文件路径
            file_path = '_internal/res/first_run_flag.txt'
            # 删除文件
            os.remove(file_path)
            sys.exit(0)
    else:
        go()


def go():
    # 创建WMI对象
    wmi = win32com.client.GetObject("winmgmts:")

    # 查询所有用户账户
    accounts = wmi.ExecQuery("SELECT * FROM Win32_UserAccount")

    # 遍历用户账户并打印信息
    for account in accounts:
        if account.Name == '16514':
            admin = messagebox.askyesno('确认', '你是否需要进入到开发者模式')
            if admin:
                welcome_Admin()
            else:
                second_window()
        else:
            second_window()
        # print(f"用户名: {account.Name}")
        # print(f"全名: {account.FullName}")


def welcome_Admin():
    def a():
        def gb():
            dm.destroy()
            splash.deiconify()

        splash.withdraw()
        global name_label

        def draw_name():
            global names, name_label
            if names:
                names = ["李宇轩", "杨政皓", "高逸航", "高子航", "石博宇", "张高菲", "陈赞彭", "杨凯琪", "杜雨嘉诺", "马欣妍"]
                chosen_name = random.choice(names)
                name_label.config(text=chosen_name, font=("汉仪文黑-85W", 50), width=6)
                names.remove(chosen_name)
            else:
                names = ["李宇轩", "杨政皓", "高逸航", "高子航", "石博宇", "张高菲", "陈赞彭", "杨凯琪", "杜雨嘉诺", "马欣妍"]

        names = ["李宇轩", "杨政皓", "高逸航", "高子航", "石博宇", "张高菲", "陈赞彭", "杨凯琪", "杜雨嘉诺", "马欣妍"]

        def set_window_position(dm, x, y):
            # 仅仅设置窗口的位置，不改变大小
            dm.geometry("+{}+{}".format(x, y))

        # 创建tkinter窗口
        dm = tk.Tk()
        dm.title("点名程序")
        dm.iconbitmap('_internal\\res\\favicon.ico')
        dm.attributes("-topmost", True)
        # 获取屏幕的宽度和高度
        screen_width = dm.winfo_screenwidth()
        screen_height = dm.winfo_screenheight()
        # 设置窗口的位置为屏幕左下角
        # 注意：Tkinter的坐标系统中，左上角是(0,0)，右下角坐标是(screen_width, screen_height)
        set_window_position(dm, screen_width // 3, screen_height // 3)
        dm.resizable(False, False)
        dm.protocol("WM_DELETE_WINDOW", gb)

        # 创建标签用于显示抽取的名字
        name_label = tk.Label(dm, font=("汉仪文黑-85W", 50), width=6)
        name_label.grid(row=0, column=1, padx=20, pady=5)

        # 创建按钮用于抽取名字
        draw_button = tk.Button(dm, text="抽取名字", command=draw_name)
        draw_button.grid(row=1, column=1, padx=20, pady=10)

        # 运行tkinter事件循环
        dm.mainloop()

    def set_window_position(splash, x, y):
        # 仅仅设置窗口的位置，不改变大小
        splash.geometry("+{}+{}".format(x, y))

    # 主程序
    splash = Tk()
    splash.title("点名系统")
    splash.iconbitmap('_internal\\res\\favicon.ico')
    splash.attributes("-topmost", True)
    # 获取屏幕的高度
    screen_height = splash.winfo_screenheight()
    # 设置窗口的位置为屏幕左下角
    # 注意：Tkinter的坐标系统中，左上角是(0,0)，右下角坐标是(screen_width, screen_height)
    set_window_position(splash, 10, 0+screen_height//2+screen_height//4+screen_height//12)
    splash.overrideredirect(True)
    splash.resizable(False, False)  # 禁止用户调整窗口的宽度和高度
    set_window_transparency(splash.winfo_id(), 70)

    button = Button(splash, text='点名', font=('汉仪文黑-85W', 25), command=a, width=4, height=1)
    button1 = Button(splash, text='X', font=('汉仪文黑-85W', 9), width=2, height=1)
    button2 = Button(splash, text='-', font=('汉仪文黑-85W', 9), width=2, height=1)
    button3 = Button(splash, text='□', font=('汉仪文黑-85W', 9), width=2, height=1)
    button4 = Button(splash, text='设置', font=('汉仪文黑-85W', 9), width=2, height=1)
    button.pack(side='left', fill="y")
    button1.pack(side='top')
    button2.pack(side='top')
    button3.pack(side='top')
    button4.pack(side='top')

    splash.mainloop()


def second_window():
    def a():
        def gb():
            dm.destroy()
            splash.deiconify()

        splash.withdraw()
        global name_label

        def draw_name():
            global names_list, name_label
            if names_list:
                chosen_name = random.choice(names_list)
                name_label.config(text=chosen_name, font=("汉仪文黑-85W", 50), width=6)
                names_list.remove(chosen_name)
            else:
                # 更新内容
                with open('_internal/res/students.json', 'r', encoding='utf-8') as file:
                    names_list = json.load(file)

        # 初始化名字列表
        with open('_internal/res/students.json', 'r', encoding='utf-8') as file:
            names_list = json.load(file)

        def set_window_position(dm, x, y):
            # 仅仅设置窗口的位置，不改变大小
            dm.geometry("+{}+{}".format(x, y))

        # 创建tkinter窗口
        dm = tk.Tk()
        dm.title("点名程序")
        dm.iconbitmap('_internal\\res\\favicon.ico')
        dm.attributes("-topmost", True)
        dm.resizable(False, False)
        dm.protocol("WM_DELETE_WINDOW", gb)
        # 获取屏幕的宽度和高度
        screen_width = dm.winfo_screenwidth()
        screen_height = dm.winfo_screenheight()
        # 设置窗口的位置为屏幕左下角
        # 注意：Tkinter的坐标系统中，左上角是(0,0)，右下角坐标是(screen_width, screen_height)
        set_window_position(dm, screen_width // 3, screen_height // 2)

        # 创建标签用于显示抽取的名字
        name_label = tk.Label(dm, font=("汉仪文黑-85W", 50), width=6)
        name_label.grid(row=0, column=1, padx=20, pady=5)

        # 创建按钮用于抽取名字
        draw_button = tk.Button(dm, text="抽取名字", command=draw_name)
        draw_button.grid(row=1, column=1, padx=20, pady=10)

        # 运行tkinter事件循环
        dm.mainloop()

    def s():
        # 设置窗口
        def hh():
            settings.destroy()
            splash.deiconify()

        splash.withdraw()
        settings = Tk()
        settings.title("设置")
        settings.iconbitmap('_internal\\res\\favicon.ico')
        settings.attributes("-topmost", True)
        width = 500
        height = 500
        screen_width = settings.winfo_screenwidth()
        screen_height = settings.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        settings.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
        settings.resizable(False, False)
        settings.protocol("WM_DELETE_WINDOW", hh)

        def edit_names():
            settings.withdraw()
            os.system("_internal\\res\\students.json")
            settings.deiconify()

        def gy():
            # 关于窗口
            def gx():
                xck.withdraw()

                # GitHub仓库URL
                GITHUB_REPO_URL = "https://api.github.com/repos/MengdeUser/roll_call/releases/latest"
                # 本地软件版本
                LOCAL_VERSION = current_version  # 请替换为你的软件当前版本
                # 本地软件安装路径
                INSTALL_PATH = os.path.join(os.path.dirname(__file__), "setup.exe")  # 请替换为你的软件安装路径

                def get_latest_version():
                    """从GitHub获取最新版本信息"""
                    try:
                        response = requests.get(GITHUB_REPO_URL)
                        response.raise_for_status()
                        data = response.json()
                        return data['tag_name']
                    except requests.exceptions.RequestException as e:
                        messagebox.showinfo("提示", f"获取最新版本时出错: {e}")
                        return None

                def download_file(progress_bar, progress_label, latest_version, root):
                    """下载并安装最新版本"""
                    # 这里需要根据实际情况编写下载和安装的代码
                    # 以下代码仅为示例
                    # 假设下载的安装包是setup.exe
                    # 下载命令
                    download_command = f"https://github.com/MengdeUser/roll_call/releases/download/{latest_version}/setup.exe"
                    # 执行下载
                    with requests.get(download_command, stream=True) as r:
                        r.raise_for_status()
                        total_length = int(r.headers.get('content-length', 1))
                        with open(INSTALL_PATH, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=8192):
                                if chunk:  # 过滤掉保持连接的chunk
                                    f.write(chunk)
                                    downloaded = f.tell()
                                    # 更新进度条
                                    progress_bar['value'] = (downloaded / total_length) * 100
                                    progress_label['text'] = f"{downloaded / total_length * 100:.2f}%"
                                    root.update_idletasks()  # 更新GUI
                    root.destroy()
                    wait = messagebox.askyesno("提示", "您是否需要启用新版本安装程序？")
                    if wait:
                        # 运行安装后的软件
                        subprocess.run(f"{INSTALL_PATH}", shell=True)
                    else:
                        sys.exit(0)

                def main1(latest_version):
                    root = tk.Tk()
                    root.title("下载进度")
                    root.iconbitmap('_internal\\res\\favicon.ico')
                    root.attributes("-topmost", True)
                    root.overrideredirect(True)
                    width = 400
                    height = 120
                    screen_width = root.winfo_screenwidth()
                    screen_height = root.winfo_screenheight()
                    x = (screen_width / 2) - (width / 2)
                    y = (screen_height / 2) - (height / 2)
                    root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
                    root.resizable(False, False)

                    showinfo = Label(root, text=f"下载和安装版本{latest_version}...", font=("汉仪文黑-85W", 15))
                    showinfo.pack(pady=5)
                    # 创建进度条
                    progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
                    progress_bar.pack(pady=10)
                    # 创建进度标签
                    progress_label = tk.Label(root, text="0%")
                    progress_label.pack(pady=5)

                    latest_version = get_latest_version()
                    if latest_version > LOCAL_VERSION:
                        download_file(progress_bar, progress_label, latest_version, root)
                    else:
                        # 这里添加预运行的代码
                        root.destroy()
                        messagebox.showinfo("提示", "已是最新版本！")
                        xck.deiconify()

                    root.mainloop()

                main1(latest_version=get_latest_version())

            def gw():
                # 打开网页
                webbrowser.open(web)

            def git():
                webbrowser.open(hub)

            def pp():
                settings.deiconify()
                xck.destroy()

            def kfc():
                global kfz
                kfz += 1
                if kfz == 14:
                    messagebox.showinfo("来自制作组弹窗", "您已进入开发者模式")
                elif kfz >= 15:
                    messagebox.showinfo("来自制作组弹窗", "别按了你已经进入开发者模式了")

            settings.withdraw()

            xck = Tk()

            xck.title("关于")
            xck.iconbitmap('_internal\\res\\favicon.ico')
            xck.attributes("-topmost", True)
            width = 500
            height = 500
            screen_width = xck.winfo_screenwidth()
            screen_height = xck.winfo_screenheight()
            x = (screen_width / 2) - (width / 2)
            y = (screen_height / 2) - (height / 2)
            xck.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
            xck.resizable(False, False)
            xck.protocol("WM_DELETE_WINDOW", pp)

            title = Label(xck, text='关于', font=('汉仪文黑-85W', 25))
            title.grid(row=0, column=0, padx=20, pady=5, sticky=W)
            title1 = Label(xck, text='版本信息：', font=('汉仪文黑-85W', 13))
            title1.grid(row=1, column=0, padx=20, pady=10, sticky=W)
            title2 = ttk.Button(xck, text='0.1.7', command=kfc)
            title2.grid(row=1, column=1, padx=0, pady=0, sticky=W)
            title3 = Label(xck, text='版本更新：', font=('汉仪文黑-85W', 13))
            bn = ttk.Button(xck, text='检查版本', command=gx)
            title3.grid(row=2, column=0, padx=20, pady=10, sticky=W)
            bn.grid(row=2, column=1, padx=0, pady=0, sticky=W)
            title4 = Label(xck, text='官网：', font=('汉仪文黑-85W', 13))
            bn1 = ttk.Button(xck, text='点名系统官网_老师助手', command=gw)
            bn2 = ttk.Button(xck, text='GitHub官网（需要打开加速器）', command=git)
            title4.grid(row=3, column=0, padx=20, pady=10, sticky=W)
            bn1.grid(row=3, column=1, padx=0, pady=0, sticky=W)
            bn2.grid(row=4, column=1, padx=0, pady=0, sticky=W)

            xck.mainloop()

        settings_button = ttk.Button(settings, text="设置点名名单", command=edit_names)
        settings_button.grid(row=1, column=1, padx=20, pady=10, sticky=W)
        settings_button1 = ttk.Button(settings, text="关于", command=gy)
        settings_button1.grid(row=2, column=1, padx=20, pady=10, sticky=W)

        settings.mainloop()

    def k():
        # 关闭
        sys.exit(0)

    def t():
        # 最小化
        splash.overrideredirect(False)
        pyautogui.keyDown('alt')
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.press('n')
        pyautogui.keyUp('alt')

    def d():
        # 还原
        splash.overrideredirect(True)
        splash.attributes("-topmost", True)

    def y():
        # 移动
        splash.overrideredirect(False)
        pyautogui.keyDown('alt')
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.press('m')
        pyautogui.keyUp('alt')

    def set_window_position(splash, x, y):
        # 仅仅设置窗口的位置，不改变大小
        splash.geometry("+{}+{}".format(x, y))

    # 主程序
    splash = Tk()
    splash.title("点名系统")
    splash.iconbitmap('_internal\\res\\favicon.ico')
    splash.attributes("-topmost", True)
    # 获取屏幕的高度
    screen_height = splash.winfo_screenheight()
    # 设置窗口的位置为屏幕左下角
    # 注意：Tkinter的坐标系统中，左上角是(0,0)，右下角坐标是(screen_width, screen_height)
    set_window_position(splash, 10, 0+screen_height//2+screen_height//4+screen_height//16)
    splash.overrideredirect(True)
    splash.resizable(False, False)  # 禁止用户调整窗口的宽度和高度
    set_window_transparency(splash.winfo_id(), 70)

    button = Button(splash, text='点名', font=('汉仪文黑-85W', 25), command=a, width=4, height=1)
    button1 = Button(splash, text='X', font=('汉仪文黑-85W', 8), command=k, width=2, height=1)
    button2 = Button(splash, text='-', font=('汉仪文黑-85W', 8), command=t, width=2, height=1)
    button3 = Button(splash, text='⿻', font=('汉仪文黑-85W', 8), command=d, width=2, height=1)
    button4 = Button(splash, text='设置', font=('汉仪文黑-85W', 8), command=s, width=4, height=1)
    button5 = Button(splash, text='↑\n↓', font=('汉仪文黑-85W', 9), command=y, width=1, height=3)
    button.pack(side='left', fill="y")
    button4.pack(side='bottom', fill='x')
    button5.pack(side='right', fill='y')
    button1.pack(side='top')
    button2.pack(side='top')
    button3.pack(side='top')

    splash.mainloop()


def is_process_running():
    lock_file = '_internal/my_program.lock'
    if os.path.exists(lock_file):
        return True
    else:
        with open(lock_file, 'w') as f:
            f.write('锁定')
        return False


if is_process_running():
    messagebox.showerror("警告", "程序已经在运行中")
else:
    try:
        # 你的程序代码
        setup()
    finally:
        os.remove('_internal\\my_program.lock')
