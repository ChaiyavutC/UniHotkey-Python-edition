##Hotkey
import keyboard as key
import win32gui
import win32con

import pyautogui

import myfunction

import tkinter as tk

target_mic_name = "Microphone (Razer Seiren Mini)"
volume = myfunction.system_mic_set_target(target_mic_name)
active_mic_mute = myfunction.system_mic_toggle_mute(0, volume)

def hotkey_system_mic_toggle_mute():
    global active_mic_mute
    active_mic_mute = myfunction.system_mic_toggle_mute(active_mic_mute, volume)
    if active_mic_mute == 1:
        mic_state.config(text = 'Muted')
    else:
        mic_state.config(text = 'Unmuted')
    return

def hotkey_system_move_active_win_to_next_monitor_or_move_to_nearest_area():
    ActiveWin = win32gui.GetForegroundWindow()
    myfunction.system_move_active_win_to_next_monitor_or_move_to_nearest_area(ActiveWin)
    return

def hotkey_minimize_active_win():
    print("System.win manager : Minimize active window")
    ActiveWin = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(ActiveWin, win32con.SW_MINIMIZE)
    return

def hotkey_move_active_win_along_mouse():
    print("System.win manager : Move active window along mouse")
    ActiveWin = win32gui.GetForegroundWindow()
    myfunction.system_move_active_win_along_mouse(ActiveWin)
    return

def hotkey_move_active_win_backward():
    print("System.win manager : Move active window to back")
    ActiveWin = win32gui.GetForegroundWindow()
    win32gui.SetWindowPos(ActiveWin, win32con.HWND_BOTTOM, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    return

def hotkey_toggle_on_top_active_window():
    print("System.win manager : Toggle On Top active window")
    ActiveWin = win32gui.GetForegroundWindow()
    window_styles = win32gui.GetWindowLong(ActiveWin, win32con.GWL_EXSTYLE)
    is_on_top = window_styles & win32con.WS_EX_TOPMOST

    # Toggle the window on top
    if is_on_top:
        win32gui.SetWindowPos(ActiveWin, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    else:
        win32gui.SetWindowPos(ActiveWin, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    return

def hotkey_system_vol_up():
    pyautogui.press('volumeup')
    print("System.vol : Increased vol")
    return

def hotkey_system_vol_down():
    pyautogui.press('volumedown')
    print("System.vol : Decreased vol")
    return

def hotkey_system_media_next():
    pyautogui.press('nexttrack')
    print("System.media : Next")
    return

def hotkey_system_media_pause():
    pyautogui.press('playpause')
    print("System.media : Next")
    return

def hotkey_system_media_stop():
    pyautogui.press('stop')
    print("System.media : Next")
    return

def hotkey_system_media_previous():
    pyautogui.press('prevtrack')
    print("System.media : Previous")
    return

#system_mic_toggle_mute
key.add_hotkey("alt+shift+z",hotkey_system_mic_toggle_mute)
#system_move_active_win_to_other_monitors
key.add_hotkey("alt+shift+r",hotkey_system_move_active_win_to_next_monitor_or_move_to_nearest_area)
#system_minimum_active_win
key.add_hotkey("alt+shift+e",hotkey_minimize_active_win)
#system_move_active_win_along_mouse
key.add_hotkey("alt+shift+q",hotkey_move_active_win_along_mouse)
#system_move_active_win_backward
key.add_hotkey("alt+shift+b",hotkey_move_active_win_backward)
#system_toggle_on_top_active_window
key.add_hotkey("alt+shift+a",hotkey_toggle_on_top_active_window)
#system_vol_up
key.add_hotkey("ctrl+F2",hotkey_system_vol_up)
#system_vol_down
key.add_hotkey("ctrl+F1",hotkey_system_vol_down)
#system_media_next
key.add_hotkey("ctrl+F5",hotkey_system_media_next)
#system_media_pause
key.add_hotkey("ctrl+F6",hotkey_system_media_pause)
#system_media_stop
key.add_hotkey("ctrl+F7",hotkey_system_media_stop)
#system_media_previous
key.add_hotkey("ctrl+F8",hotkey_system_media_previous)
#key.wait()

mainUI = tk.Tk()
mainUI.title('Universal v.0.0.0.3beta by ChaiyavutC')
mic_state = tk.Label(mainUI, text='Muted')
mic_state.pack()
mainUI.mainloop()
