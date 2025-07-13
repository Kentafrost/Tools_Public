import pyautogui
import subprocess
from time import sleep
import win32gui
import win32con
from pathlib import Path
import win32process, psutil
import tkinter
from tkinter import filedialog

def select_directory():
    root = tkinter.Tk()
    root.withdraw()
    selected_dir = filedialog.askdirectory(title="フォルダを選択してください", initialdir=r"D:\app")
    print("Selected directory:", selected_dir)
    return selected_dir


def get_app_forground_name(hwnd, extra):
    if win32gui.IsWindowVisible(hwnd):
        if extra in win32gui.GetWindowText(hwnd):
            win32gui.SetForegroundWindow(hwnd)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)


def record():
    # 録画コマンド(Ctrl+R)
    pyautogui.click(1366, 401)
    pyautogui.keyDown("ctrl")
    pyautogui.keyDown("r")
    pyautogui.keyUp("r")
    pyautogui.keyUp("ctrl")

def active_window_process_name():
    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
    print(psutil.Process(pid[-1]).name())


if __name__ == '__main__':
    dir = select_directory()

    root = tkinter.Tk()
    root.withdraw()
    fTyp = [("","*")]
    tkinter.messagebox.showinfo('自動録画用tool', 'ディレクトリ内の録画するアプリを指定してください')
    app_path = filedialog.askopenfilename(filetypes=fTyp, initialdir=dir)

    app_fold_name = Path(app_path).parent.name

    if "製品版" in app_fold_name:
        app_title = app_fold_name.replace("製品版", "")
    elif "ver" in app_fold_name:
        app_title = app_fold_name.replace("ver", "")
    else:
        app_title = app_fold_name

    print(app_title)
    
    subprocess.run(app_path)  # Start the app executable
    sleep(5)  # Wait for the app window to appear

    sleep(3)  # Click on a window you like and wait 3 seconds
    active_window_process_name()

    # Wait for the window title to change (app window appears)
    windowTile = ""
    while True:
        newWindowTile = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if newWindowTile != windowTile:
            windowTile = newWindowTile
            print(windowTile)
            break

    # Send recording command twice
    record()
    record()
    sleep(5)
    pyautogui.moveTo(919, 400, 2)
    pyautogui.click(button="left", clicks=300, interval=5)  # click automatically