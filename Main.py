from pynput.keyboard import Key, Listener
import win32api
import win32gui

def on_press(key):
    #print(f'{key} pressed')
    if key == Key.alt_gr:
        print('>>> I press RALT <<<')
        #Toggle Micorphone
        #https://stackoverflow.com/questions/50025927/how-mute-microphone-by-python
        ###win32api, win32gui
        WM_APPCOMMAND = 0x319
        APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000

        hwnd_active = win32gui.GetForegroundWindow()
        win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)

def on_release(key):
    #print(f'{key} released')

    if key == Key.esc:
        # Stop listener
        return False
    
# Collect events until released
with Listener(on_press=on_press,on_release=on_release) as listener:
    print('... other code ...')
    listener.join()