"""
volume.py
====================================
Methods for getting and setting an audio endpoint volume
"""

from comtypes.client import CreateObject
from comtypes import GUID, IUnknown, CLSCTX_ALL, COMMETHOD,HRESULT
from ctypes import POINTER, pointer, cast, c_float, c_void_p
from ctypes.wintypes import DWORD, LPWSTR, LPCWSTR, UINT, BOOL

#Device class
class IMMDevice(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{D666063F-1587-4E43-81F1-B948E807363F}')
    _idlflags_ = []

#Endpoint class from endpointvolume.h
class IAudioEndpointVolume(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{5CDF2C82-841E-4546-9722-0CF74078229A}')
    _idlflags_ = []

#Enumerator class
class IMMDeviceEnumerator(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
    _idlflags_ = []

IMMDevice._methods_ = [
    COMMETHOD([], HRESULT, 'Activate',
              (['in'], POINTER(GUID), 'iid'),
              (['in'], DWORD, 'dwClsCtx'),
              (['in'], POINTER(DWORD), 'pActivationParams'),
              (['out'], POINTER(POINTER(IUnknown)), 'ppInterface')),
    COMMETHOD([], HRESULT, 'OpenPropertyStore',
              (['in'], DWORD, 'stgmAccess'),
              (['out'], POINTER(c_void_p), 'ppProperties')),
    COMMETHOD([], HRESULT, 'GetID',
              (['out'], POINTER(LPWSTR), 'ppstrID')),
    COMMETHOD([], HRESULT, 'GetState',
              (['out'], POINTER(DWORD), 'pdwState'))]


IMMDeviceEnumerator._methods_ = [
    COMMETHOD([], HRESULT, 'EnumAudioEndpoints',
              (['in'], DWORD, 'dataFlow'),
              (['in'], DWORD, 'dwStateMask'),
              (['out'], POINTER(c_void_p), 'ppDevices')),
    COMMETHOD([], HRESULT, 'GetDefaultAudioEndpoint',
              (['in'], DWORD, 'dataFlow'),
              (['in'], DWORD, 'role'),
              (['out'], POINTER(POINTER(IMMDevice)), 'ppEndpoint')),
    COMMETHOD([], HRESULT, 'GetDevice',
              (['in'], LPCWSTR, 'pwstrID'),
              (['out'], POINTER(POINTER(IMMDevice)), 'ppDevice')),
    COMMETHOD([], HRESULT, 'RegisterEndpointNotificationCallback',
              (['in'], c_void_p, 'pclient')),
    COMMETHOD([], HRESULT, 'UnregisterEndpointNotificationCallback',
              (['in'], c_void_p, 'pclient'))]


IAudioEndpointVolume._methods_ = [
    COMMETHOD([], HRESULT, 'RegisterControlChangeNotify',
              (['in'], c_void_p, 'pNotify')),
    COMMETHOD([], HRESULT, 'UnregisterControlChangeNotify',
              (['in'], c_void_p, 'pNotify')),
    COMMETHOD([], HRESULT, 'GetChannelCount',
              (['out'], POINTER(UINT), 'pnChannelCount')),
    COMMETHOD([], HRESULT, 'SetMasterVolumeLevel',
              (['in'], c_float, 'fLevelDB'),
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'SetMasterVolumeLevelScalar',
              (['in'], c_float, 'fLevel'),
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'GetMasterVolumeLevel',
              (['out'], POINTER(c_float), 'pfLevelDB')),
    COMMETHOD([], HRESULT, 'GetMasterVolumeLevelScalar',
              (['out'], POINTER(c_float), 'pfLevel')),
    COMMETHOD([], HRESULT, 'SetChannelVolumeLevel',
              (['in'], UINT, 'nChannel'),
              (['in'], c_float, 'fLevelDB'),
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'SetChannelVolumeLevelScalar',
              (['in'], UINT, 'nChannel'),
              (['in'], c_float, 'fLevel'),
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'GetChannelVolumeLevel',
              (['in'], UINT, 'nChannel'),
              (['out'], POINTER(c_float), 'pfLevelDB')),
    COMMETHOD([], HRESULT, 'GetChannelVolumeLevelScalar',
              (['in'], UINT, 'nChannel'),
              (['out'], POINTER(c_float), 'pfLevel')),
    COMMETHOD([], HRESULT, 'SetMute',
              (['in'], BOOL, 'bMute'),
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'GetMute',
              (['out'], POINTER(BOOL), 'pbMute')),
    COMMETHOD([], HRESULT, 'GetVolumeStepInfo',
              (['out'], POINTER(UINT), 'pnStep'),
              (['out'], POINTER(UINT), 'pnStepCount')),
    COMMETHOD([], HRESULT, 'VolumeStepUp',
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'VolumeStepDown',
              (['in'], POINTER(GUID), 'pguidEventContext')),
    COMMETHOD([], HRESULT, 'QueryHardwareSupport',
              (['out'], POINTER(DWORD), 'pdwHardwareSupportMask')),
    COMMETHOD([], HRESULT, 'GetVolumeRange',
              (['out'], POINTER(c_float), 'pflVolumeMindB'),
              (['out'], POINTER(c_float), 'pflVolumeMaxdB'),
              (['out'], POINTER(c_float), 'pflVolumeIncrementdB'))]


class AudioEndpoint:
    """Representation of an audio endpoint.

    Note: Should be instantiated as a context manager
    """
    def __init__(self):
        self.CLSID = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')

    def __enter__(self):
        self.enum = CreateObject(self.CLSID, interface = IMMDeviceEnumerator)
        RENDER_DIRECTION = 0
        DEVICE_ROLE = 1
        self.endpt_device = self.enum.GetDefaultAudioEndpoint(RENDER_DIRECTION, DEVICE_ROLE)
        iid = GUID('{5CDF2C82-841E-4546-9722-0CF74078229A}')
        refiid = pointer(iid)
        self.endpt = self.endpt_device.Activate(refiid, CLSCTX_ALL, None)
        self._endpt = cast(self.endpt, POINTER(IAudioEndpointVolume))

    def __exit__(self):
        self.enum.Release()
        self.endpt_device.Release()
        self.endpt.Release()
        
    def get_master_volume(self):
        """Returns master volume as proportion of full volume

        Returns
        -------
        float
            scalar between 0 and 1 indicating relative volume
        """
            
        return self._endpt.GetMasterVolumeLevelScalar()
    def set_master_volume(self, vol):
        """Sets master volume to a proportion of full volume

        Parameters
        ----------
        vol : float
            scalar between 0 and 1 indicating relative volume
        """
        vol_to_set = max(0.0, min(1.0, vol))
        self._endpt.SetMasterVolumeLevelScalar(vol_to_set, None)

    def get_mute(self):
        """Returns muted state of the audio endpoint

        Returns
        -------
        bool
            returns True if the endpoint is muted, False otherwise
        """
        return self._endpt.GetMute()
    def set_mute(self, state):
        """Sets muted state of the audio endpoint

        Parameters
        ----------
        state : bool
            set to True to mute the audio endpoint, False to unmute
        """
        self._endpt.SetMute(bool(state), None)

    def increment_master_volume(self):
        """Increments the master volume by a decibel amount

        See Also
        --------
        get_range : minimum, maximum and increment volumes in decibels"""
        self._endpt.VolumeStepUp(None)
    def decrement_master_volume(self):
        """Decrements the master volume by a decibel amount

        See Also
        --------
        get_range : minimum, maximum and increment volumes in decibels"""
        self._endpt.VolumeStepDown(None)

    def get_range(self):
        """Returns ranges (in dB) of the minimum and maximum volume,
        along with the increment

        Returns
        -------
        min : float
            minimum master volume (dB)
        max: float
            maximum master volume (dB)
        increment: float
            increment of master volume (dB)
        """
        return self._endpt.GetVolumeRange()    
