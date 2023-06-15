import function

##Keylogger
import keyboard

###microphone list
##https://pypi.org/project/pycaw/
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pycaw.constants import CLSID_MMDeviceEnumerator
### Get default device
interface = AudioUtilities.GetMicrophone().Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
### Mute the default device
cast(interface, POINTER(IAudioEndpointVolume)).SetMute(1, None)
active_mic_mute = 1

def action_right_alt():
    print('Toggle mute mic')
    #mic_default_toggle_mute()
    
keyboard.add_hotkey('Alt', action_right_alt, suppress=True, trigger_on_release=False)
keyboard.wait()