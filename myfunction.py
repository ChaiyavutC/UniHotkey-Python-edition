from __future__ import print_function
from ctypes import POINTER, cast
import comtypes

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IMMDeviceEnumerator, EDataFlow, ERole
from pycaw.constants import CLSID_MMDeviceEnumerator

import win32gui
import win32api
import win32con
import time

"""
This section is for Toggle Mute selected microphone result to whole system.
"""
#This class credit : https://github.com/AndreMiras/pycaw/issues/8#issuecomment-622122323
class MyAudioUtilities(AudioUtilities):
    @staticmethod
    def GetMicrophone(id=None):
        device_enumerator = comtypes.CoCreateInstance(CLSID_MMDeviceEnumerator, IMMDeviceEnumerator, comtypes.CLSCTX_INPROC_SERVER)
        if id is not None:
            microphones = device_enumerator.GetDevice(id)
        else:
            microphones = device_enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender.value, ERole.eMultimedia.value)
        return microphones

def system_mic_get_deivce_name():
    for dev in AudioUtilities.GetAllDevices():
        #find only available mic device
        if "Active" in str(dev.state) and "Mic" in dev.FriendlyName:
            print(dev.FriendlyName)
    return

def system_mic_set_target(target_mic_name):
    ### reset
    mixer_output = None

    ### select the same device name from all devices
    for device in MyAudioUtilities.GetAllDevices():
        if target_mic_name in str(device):
            mixer_output = device
            print("System.mic : Found target_mic_name")
            
    if mixer_output is None:
        print("System.mic : Not found target_mic_name")
            
    #devices = MyAudioUtilities.GetMicrophone(mixer_output.id)
    interface = MyAudioUtilities.GetMicrophone(mixer_output.id).Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return volume

def system_mic_toggle_mute(active_mic_mute, volume):

    ### Toggle mute the default device
    
    if active_mic_mute == 1:
        volume.SetMute(0, None)
        active_mic_mute = 0
        print("System.mic : Unmuted")
    elif active_mic_mute == 0:
        volume.SetMute(1, None)
        active_mic_mute = 1
        print("System.mic : Muted")
        if volume.GetMute() == 1:
            print("System.mic : Muted and Checked already muted")
    return active_mic_mute

"""
This section is for move the active windows to next monitor.
"""
def system_move_active_win_to_next_monitor(window_handle):
        # Check if the window is maximized
    window_placement = win32gui.GetWindowPlacement(window_handle)
    #; window_placement[1] refers to the showCmd field of the window_placement tuple. This field represents the current show command for the window, indicating its state (such as whether it is maximized, minimized, or in a normal state).
    if window_placement[1] == win32con.SW_MAXIMIZE:
        # Restore the window to its normal state
        win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)

    # Get the monitors info #>> monitor_info = [(<PyHANDLE:65537>, <PyHANDLE:0>, (0, 0, 1920, 1080)), (<PyHANDLE:65539>, <PyHANDLE:0>, (-1920, -120, 0, 1080))]
    monitor_info = win32api.EnumDisplayMonitors(None, None)
    # Get number of monitors
    num_monitors = len(monitor_info)

    # Find the current monitor index based on window position
    current_monitor_index = None
    #; win32gui.GetWindowRect(window_handle) is used to obtain the bounding rectangle of the window specified by the window_handle.
    window_rect = win32gui.GetWindowRect(window_handle)
    # get width and height from the bounding rectangle of the window.
    window_width = window_rect[2] - window_rect[0]
    window_height = window_rect[3] - window_rect[1]
    # get x, y of the window (x determine from top left corner of the screen to window_rect[0]) (x determine from top left corner of the screen to window_rect[1])
    window_x = window_rect[0] + window_width // 2
    window_y = window_rect[1] + window_height // 2
    
    for i, monitor in enumerate(monitor_info):
        #; monitor[2] refers to the third element of the monitor tuple, which contains the bounding rectangle coordinates of the monitor.
        monitor_left = monitor[2][0] #; represents the left coordinate of the monitor's bounding rectangle.
        monitor_top = monitor[2][1] #; represents the top coordinate of the monitor's bounding rectangle.
        monitor_right = monitor[2][2] #; represents the right coordinate of the monitor's bounding rectangle.
        monitor_bottom = monitor[2][3] #; represents the bottom coordinate of the monitor's bounding rectangle.
        #is used to check if the window's position falls within the bounding rectangle of a particular monitor. To simply, to get index of monitor which active window at.
        if window_x >= monitor_left and window_x < monitor_right and window_y >= monitor_top and window_y < monitor_bottom:
            current_monitor_index = i
            break

    # Increment monitor_index to cycle to the next monitor
    if current_monitor_index is not None:
        monitor_index = (current_monitor_index + 1) % num_monitors #; '% num_monitor' uses to cycle index of monitor in number of monitors.

    # Get the target monitor info
    target_monitor = monitor_info[monitor_index][2]
    target_monitor_left = target_monitor[0]
    target_monitor_top = target_monitor[1]
    target_monitor_right = target_monitor[2]
    target_monitor_bottom = target_monitor[3]

    # Calculate the new position on the target monitor
    new_x = target_monitor_left + ((target_monitor_right - target_monitor_left) - window_width) // 2
    new_y = target_monitor_top + ((target_monitor_bottom - target_monitor_top) - window_height) // 2

    # Set the new window position
    win32gui.SetWindowPos(window_handle, win32con.HWND_TOP, new_x, new_y, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)

    # Maximize the window again if it was originally maximized
    if window_placement[1] == win32con.SW_MAXIMIZE:
        time.sleep(0.5)
        win32gui.ShowWindow(window_handle, win32con.SW_MAXIMIZE)
    return