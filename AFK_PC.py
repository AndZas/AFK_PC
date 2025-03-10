import configparser
import ctypes
import os
import sys

import win32con
import win32gui

AFK_TIME = 120
ONLINE_CHECK_INTERVAL = 5
OFFLINE_CHECK_INTERVAL = 1
PAUSE_CHECK_INTERVAL = 1200
EXCEPTION_APPS = []
AUTO_START = 0

LAST_WINDOWS = []
PAUSED = False
AFK = False


class LastInputInfo(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_ulong)]


def load_settings():
    global AFK_TIME, ONLINE_CHECK_INTERVAL, OFFLINE_CHECK_INTERVAL, PAUSE_CHECK_INTERVAL, EXCEPTION_APPS, AUTO_START
    config = configparser.ConfigParser()
    config.read('config.ini')
    AFK_TIME = int(config['DEFAULT']['AFK_TIME'])
    ONLINE_CHECK_INTERVAL = int(config['DEFAULT']['ONLINE_CHECK_INTERVAL'])
    OFFLINE_CHECK_INTERVAL = float(config['DEFAULT']['OFFLINE_CHECK_INTERVAL'])
    PAUSE_CHECK_INTERVAL = int(config['DEFAULT']['PAUSE_CHECK_INTERVAL'])
    EXCEPTION_APPS = config['DEFAULT']['EXCEPTION_APPS'][1:-1].replace("'", '').split(', ')
    AUTO_START = int(config['DEFAULT']['AUTO_START'])


def save_settings():
    config = configparser.ConfigParser()
    config['DEFAULT'] = {'AFK_TIME': str(AFK_TIME),
                         'ONLINE_CHECK_INTERVAL': str(ONLINE_CHECK_INTERVAL),
                         'OFFLINE_CHECK_INTERVAL': str(OFFLINE_CHECK_INTERVAL),
                         'PAUSE_CHECK_INTERVAL': str(PAUSE_CHECK_INTERVAL),
                         'EXCEPTION_APPS': str(EXCEPTION_APPS),
                         'AUTO_START': str(AUTO_START)}
    with open('config.ini', 'w') as configfile:
        config.write(configfile)


def get_afk_time():
    last_input_info = LastInputInfo()
    last_input_info.cbSize = ctypes.sizeof(LastInputInfo)

    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(last_input_info)):
        return (ctypes.windll.kernel32.GetTickCount() - last_input_info.dwTime) / 1000.0
    return 0


def get_active_window_title():
    return win32gui.GetWindowText(win32gui.GetForegroundWindow())


def is_exception_application():
    if '' in EXCEPTION_APPS:
        EXCEPTION_APPS.remove('')
    return any(keyword in get_active_window_title() for keyword in EXCEPTION_APPS)


def is_watching_video():
    from re import split
    return (not any(keyword in split('[-—]', get_active_window_title())[0] for keyword in
                    ['YouTube', 'youtube', 'Twitch', 'twitch']) and
            any(keyword in get_active_window_title() for keyword in ['YouTube', 'youtube', 'Twitch', 'twitch']))


def is_window_maximize(_id):
    return win32gui.IsWindowVisible(_id) and win32gui.GetWindowText(_id) and not win32gui.IsIconic(_id)


def get_open_windows():
    windows = []
    win32gui.EnumWindows(lambda _id, _: windows.append(_id) if is_window_maximize(_id) else None, None)
    return windows


def minimize_windows():
    global LAST_WINDOWS
    LAST_WINDOWS = get_open_windows()
    for _id in get_open_windows():
        win32gui.ShowWindow(_id, win32con.SW_MINIMIZE)


def maximize_windows():
    for _id in reversed(LAST_WINDOWS):
        if win32gui.IsIconic(_id):
            win32gui.ShowWindow(_id, win32con.SW_RESTORE)


def get_absolute_path_to_exe_file():
    return os.path.abspath(sys.argv[0])


def add_to_startup():
    from win32com.client import Dispatch
    startup_folder = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    shortcut_path = os.path.join(startup_folder, "AFK_PC.lnk")
    if not os.path.exists(shortcut_path):
        exe_path = get_absolute_path_to_exe_file()
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = exe_path
        shortcut.WorkingDirectory = os.path.dirname(exe_path)
        shortcut.IconLocation = exe_path
        shortcut.save()


def remove_from_startup():
    startup_folder = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    shortcut_path = os.path.join(startup_folder, "AFK_PC.lnk")
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)


def settings():
    import customtkinter as ctk

    global AFK_TIME, ONLINE_CHECK_INTERVAL, OFFLINE_CHECK_INTERVAL, PAUSE_CHECK_INTERVAL, EXCEPTION_APPS, AUTO_START

    on_check_intervals = {"Очень редко": 30, "Редко": 10, "Часто": 5, "Довольно часто": 3, "Очень часто": 1}
    off_check_intervals = {"Очень медленно": 5, "Медленно": 3, "Быстро": 2,
                           "Довольно быстро": 1, "Очень быстро": 0.5}
    pause_check_intervals = {"Очень редко": 1200, "Редко": 600, "Часто": 300,
                             "Довольно часто": 100, "Очень часто": 50}

    def save_close():
        global AFK_TIME, ONLINE_CHECK_INTERVAL, OFFLINE_CHECK_INTERVAL, PAUSE_CHECK_INTERVAL, EXCEPTION_APPS, AUTO_START

        try:
            AFK_TIME = int(afk_time_entry.get())
            ONLINE_CHECK_INTERVAL = on_check_intervals[on_check_interval_cmb.get()]
            OFFLINE_CHECK_INTERVAL = off_check_intervals[off_check_interval_cmb.get()]
            PAUSE_CHECK_INTERVAL = pause_check_intervals[pause_check_interval_cmb.get()]
            EXCEPTION_APPS = exception_apps_widgets
            AUTO_START = int(auto_start_var.get())
            save_settings()

            if AUTO_START:
                add_to_startup()
            else:
                remove_from_startup()

        except ValueError:
            pass

        window.destroy()

    def add_exception_application(name=None):
        def delete_exception_application(_frame, _name):
            _frame.destroy()
            exception_apps_widgets.remove(_name)

        if name:
            frame = ctk.CTkFrame(exception_apps_frame, fg_color="#333333", corner_radius=8)
            frame.pack(fill='x', pady=10, padx=10)

            label = ctk.CTkLabel(frame, text=[name[:35] + "..." if len(name) > 35 else name][0])
            label.pack(side="left", padx=5, pady=5)

            delete_button = ctk.CTkButton(frame, text="X", width=28,
                                          command=lambda: delete_exception_application(frame, name))
            delete_button.pack(side="right", padx=5, pady=5)

        else:
            name = exception_apps_cmb.get()
            if name and name not in exception_apps_widgets:
                frame = ctk.CTkFrame(exception_apps_frame, fg_color="#333333", corner_radius=8)
                frame.pack(fill='x', pady=10, padx=10)

                label = ctk.CTkLabel(frame, text=[name[:35] + "..." if len(name) > 35 else name][0])
                label.pack(side="left", padx=5, pady=5)

                delete_button = ctk.CTkButton(frame, text="X", width=28,
                                              command=lambda: delete_exception_application(frame, name))
                delete_button.pack(side="right", padx=5, pady=5)

                exception_apps_widgets.append(name)

    open_applications_names = [str(win32gui.GetWindowText(i)) for i in get_open_windows()]
    if "Окно переполнения области задач." in open_applications_names:
        open_applications_names.remove('Окно переполнения области задач.')
    if "Program Manager" in open_applications_names:
        open_applications_names.remove('Program Manager')

    exception_apps_widgets = EXCEPTION_APPS

    window = ctk.CTk()
    window.title("Настройки")
    window.geometry("380x610")

    afk_time_label = ctk.CTkLabel(window, text="Время бездействия (сек):")
    afk_time_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')
    afk_time_entry = ctk.CTkEntry(window)
    afk_time_entry.grid(row=0, column=1, padx=10, pady=10, sticky='w')
    afk_time_entry.insert(0, str(AFK_TIME))

    on_check_interval_label = ctk.CTkLabel(window, text="Частота проверки на активность:")
    on_check_interval_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')
    on_check_interval_cmb = ctk.CTkComboBox(window, values=["Очень редко", "Редко", "Часто",
                                                            "Довольно часто", "Очень часто"])
    on_check_interval_cmb.grid(row=1, column=1, padx=10, pady=10, sticky='w')
    on_check_interval_cmb.set(list(on_check_intervals.keys())
                              [list(on_check_intervals.values()).index(ONLINE_CHECK_INTERVAL)])

    off_check_interval_label = ctk.CTkLabel(window, text="Скорость развертывания окон:")
    off_check_interval_label.grid(row=2, column=0, padx=10, pady=10, sticky='w')
    off_check_interval_cmb = ctk.CTkComboBox(window, values=["Очень медленно", "Медленно", "Быстро",
                                                             "Довольно быстро", "Очень быстро"])
    off_check_interval_cmb.grid(row=2, column=1, padx=10, pady=10, sticky='w')
    off_check_interval_cmb.set(list(off_check_intervals.keys())
                               [list(off_check_intervals.values()).index(OFFLINE_CHECK_INTERVAL)])

    pause_check_interval_label = ctk.CTkLabel(window, text="Частота проверки при паузе:")
    pause_check_interval_label.grid(row=3, column=0, padx=10, pady=10, sticky='w')
    pause_check_interval_cmb = ctk.CTkComboBox(window, values=["Очень редко", "Редко", "Часто",
                                                               "Довольно часто", "Очень часто"])
    pause_check_interval_cmb.grid(row=3, column=1, padx=10, pady=10, sticky='w')
    pause_check_interval_cmb.set(list(pause_check_intervals.keys())
                                 [list(pause_check_intervals.values()).index(PAUSE_CHECK_INTERVAL)])

    auto_start_var = ctk.BooleanVar(value=bool(AUTO_START))
    auto_start_cb = ctk.CTkCheckBox(window, text="Запуск с Windows", variable=auto_start_var)
    auto_start_cb.grid(row=4, column=0, padx=10, pady=10, sticky='w')

    exception_apps_label = ctk.CTkLabel(window, text="Приложения исключения:")
    exception_apps_label.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='w')

    exception_apps_cmb = ctk.CTkComboBox(window, values=open_applications_names)
    exception_apps_cmb.grid(row=6, column=0, padx=10, pady=10, sticky='w')

    exception_apps_button = ctk.CTkButton(window, text="Добавить", command=add_exception_application)
    exception_apps_button.grid(row=6, column=1, padx=10, pady=10, sticky='w')

    exception_apps_frame = ctk.CTkScrollableFrame(window)
    exception_apps_frame.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

    save_button = ctk.CTkButton(window, text="Сохранить", command=save_close)
    save_button.grid(row=8, column=0, padx=10, pady=10, sticky='w')

    quit_button = ctk.CTkButton(window, text="Отмена", command=window.destroy)
    quit_button.grid(row=8, column=1, padx=10, pady=10, sticky='w')

    if '' in exception_apps_widgets:
        exception_apps_widgets.remove('')
    for exc in exception_apps_widgets:
        add_exception_application(exc)

    window.mainloop()


def pause():
    global PAUSED
    PAUSED = not PAUSED


def exit_program(tray):
    tray.stop()
    sys.exit()


def tray_setup():
    from pystray import MenuItem, Icon
    from PIL import Image
    image = Image.new("RGB", (64, 64), (255, 0, 0))
    menu = (
        MenuItem('Настройки', settings),
        MenuItem('Пауза', pause),
        MenuItem('Выход', exit_program),
    )
    icon = Icon("AFK_PC", image, "AFK PC", menu)
    icon.run()


def mainloop():
    from time import sleep
    global PAUSED, AFK

    while True:
        if PAUSED:
            sleep(PAUSE_CHECK_INTERVAL)
            continue

        if get_afk_time() >= AFK_TIME and not is_watching_video() and not is_exception_application():
            if not AFK:
                minimize_windows()
                AFK = True
            sleep(OFFLINE_CHECK_INTERVAL)
        else:
            if AFK:
                maximize_windows()
                AFK = False
            sleep(ONLINE_CHECK_INTERVAL)


def main():
    from threading import Thread
    load_settings()
    Thread(target=mainloop, daemon=True).start()
    tray_setup()


if __name__ == '__main__':
    main()
