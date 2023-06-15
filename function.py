from __future__ import print_function
from ctypes import POINTER, cast
import comtypes

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IMMDeviceEnumerator, EDataFlow, ERole
from pycaw.constants import CLSID_MMDeviceEnumerator
active_mic_mute = 0

#https://github.com/AndreMiras/pycaw/issues/8#issuecomment-622122323
class MyAudioUtilities(AudioUtilities):
    @staticmethod
    def GetMicrophone(id=None):
        device_enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER)
        if id is not None:
            microphones = device_enumerator.GetDevice(id)
        else:
            microphones = device_enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender.value, ERole.eMultimedia.value)
        return microphones

def system_mic_get_deivce_name():
    devices = AudioUtilities.GetAllDevices()
    print(devices)
    for dev in devices:
        #find only available mic device
        if "Active" in str(dev.state) and "Mic" in dev.FriendlyName:
            print(dev.FriendlyName)
    return

def system_mic_toggle_mute(target_mic_name):
    ### reset
    mixer_output = None

    print(target_mic_name)
    ### get all devices
    devicelist = MyAudioUtilities.GetAllDevices()

    ### select the same  device name
    for device in devicelist:
        if target_mic_name in str(device):
            mixer_output = device

    ### get id
    devices = MyAudioUtilities.GetMicrophone(mixer_output.id)

    interface = devices.Activate(
        IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    
    ### Toggle mute the default device
    global active_mic_mute
    if active_mic_mute == 1:
        volume.SetMute(0, None)
        active_mic_mute = 0
    elif active_mic_mute == 0:
        volume.SetMute(1, None)
        active_mic_mute = 1
        if volume.GetMute() == 1:
            print("ok")
    return

system_mic_toggle_mute("Microphone (Razer Seiren Mini)")