##Hotkey
from ahk import AHK

import win32gui
import win32con

import myfunction

target_mic_name = "Microphone (Razer Seiren Mini)"
volume = myfunction.system_mic_set_target(target_mic_name)
active_mic_mute = myfunction.system_mic_toggle_mute(0, volume)

ahk = AHK()

def hotkey_system_mic_toggle_mute():
    global active_mic_mute
    active_mic_mute = myfunction.system_mic_toggle_mute(active_mic_mute, volume)
    return

def hotkey_system_minimize_active_win():
    ActiveWin = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(ActiveWin, win32con.SW_MINIMIZE)
    print("System.win manager : Minimize active window")
    return

def hotkey_system_move_active_win_to_next_monitor():
    ActiveWin = win32gui.GetForegroundWindow()
    myfunction.system_move_active_win_to_next_monitor(ActiveWin)
    print("System.win manager : Move active window to next monitor")
    return

def hotkey_system_vol_up():
    ahk.key_press("VOLUME_UP")
    print("System.vol : Increased vol")
    return
    
def hotkey_system_vol_down():
    ahk.key_press("VOLUME_Down")
    print("System.vol : Decreased down")
    return
    
#system_mic_toggle_mute
ahk.add_hotkey("alt+shift+z", callback=hotkey_system_mic_toggle_mute)
#system_move_active_win_to_other_monitors
ahk.add_hotkey("alt+shift+r", callback=hotkey_system_move_active_win_to_next_monitor)
#system_minimum_active_win
ahk.add_hotkey("alt+shift+e", callback=hotkey_system_minimize_active_win)
#system_vol_up
ahk.add_hotkey("alt+shift+up", callback=hotkey_system_vol_up)
#system_vol_down
ahk.add_hotkey("alt+shift+down", callback=hotkey_system_vol_down)
ahk.start_hotkeys()  # start the hotkey process thread
ahk.block_forever()  # not strictly needed in all scripts -- stops the script from exiting; sleep forever