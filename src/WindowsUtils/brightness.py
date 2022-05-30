"""
brightness.py
====================================
Methods for getting and setting brightness of a display
"""

from ctypes import (c_int, c_long, WINFUNCTYPE, windll, POINTER, byref,
                    Structure, sizeof)
from ctypes.wintypes import (WCHAR, DWORD, HANDLE, RECT, BYTE, LPCWSTR)
import WindowsUtils.helpers.Wbem as Wbem

userlib = windll.user32
monitorlib = windll.Dxva2

class PHYSICAL_MONITOR(Structure):
    _fields_ = [("handle", HANDLE),
                ("desc", WCHAR * 128)]

class DISPLAY_DEVICE(Structure):
    _fields_ = [("cb", DWORD),
                ("DeviceName", WCHAR * 32),
                ("DeviceString", WCHAR * 128),
                ("StateFlags", DWORD),
                ("DeviceID", WCHAR * 128),
                ("DeviceKey", WCHAR * 128)]

class MONITOR_INFO(Structure):
    _fields_ = [("cbSize", DWORD),
                ("rcMonitor", RECT),
                ("rcWork", RECT),
                ("dwFlags", DWORD)]

class MONITOR_INFO_EX(Structure):
    _fields_ = [("info", MONITOR_INFO),
                ("szDevice", WCHAR * 32)]

_MONITORENUMPROC = WINFUNCTYPE(c_int, HANDLE, HANDLE, POINTER(RECT), POINTER(c_long))

def _get_man_code(man_code_nums):
    #Converts the EDID bits into a manufacturer code
    f,s,t = man_code_nums
    return chr(f) + chr(s) + chr(t)

class WmiMonitor:
    """Representation of a Laptop monitor"""
    def __init__(self, display_info, name):
        self.system_name = name
        self.displayed_name = display_info.DeviceString
        self._desc = "Wmi Monitor"
        server = Wbem.create_server("ROOT\\WMI")
        query_name = self.system_name.replace('\\','\\\\')
        query_string = f"SELECT * FROM WmiMonitorID WHERE InstanceName='{query_name}'"
        man_code_encoded = Wbem.get_property(server, query_string, "ManufacturerName")
        self.manufacturer_code = _get_man_code(man_code_encoded[:3])
        server.Release()

    def get_brightness(self):
        """Gets the brightness of this monitor

        Returns
        -------
        int
            integer between 0 and 100 representing the brightness
        """
        server = Wbem.create_server("ROOT\\WMI")
        query_name = self.system_name.replace('\\','\\\\')
        query_string = "Select * from WmiMonitorBrightness"
        query_string += f" where InstanceName='{query_name}'"
        val = Wbem.get_property(server, query_string, "CurrentBrightness")
        server.Release()
        return val

    def set_brightness(self, value):
        """Sets the brightness of this monitor

        Parameters
        -------
        value : int
            integer between 0 and 100 representing the brightness to be set
        """
        server = Wbem.create_server("ROOT\\WMI")
        query_name = self.system_name.replace('\\','\\\\')
        query_string = "Select * from WmiMonitorBrightnessMethods"
        query_string += f" where InstanceName='{query_name}'"
        Wbem.exec_method(server, query_string, "WmiMonitorBrightnessMethods",
                    "WmiSetBrightness",
                         ("Timeout", "0", 19),
                         ("Brightness", str(max(min(100, value),0)), 17))
        server.Release()


class VCPMonitor:
    """Representation of a desktop monitor"""
    def __init__(self, handle, physical_monitor, display_info, desc):
        self.displayed_name = display_info.DeviceString
        dev_id = display_info.DeviceID
        self.manufacturer_code = dev_id.split('#')[1][:3]
        self.system_name = dev_id.split('#')[2]
        self._display_handle = handle
        self._desc = desc
        self._pm = physical_monitor
        self._close()

    def get_brightness(self):
        """Gets the brightness of this monitor

        Returns
        -------
        int
            integer between 0 and 100 representing the brightness
        """
        self._open()
        brightness_code = BYTE(0x10)
        brightness = DWORD()
        for i in range(25):
            if monitorlib.GetVCPFeatureAndVCPFeatureReply(self._pm.handle,
                                                          brightness_code,
                                                          None,
                                                          byref(brightness),
                                                          None):
                self._close()            
                return brightness.value
        self._close()
        raise Exception("Couldn't access brightness for this monitor")

    def set_brightness(self, value):
        """Sets the brightness of this monitor

        Parameters
        -------
        value : int
            integer between 0 and 100 representing the brightness to be set
        """
        value = max(min(value,100),0)
        self._open()
        brightness_code = BYTE(0x10)
        for i in range(50):
            monitorlib.SetVCPFeature(self._pm.handle,
                                     brightness_code,
                                     DWORD(value))
                
        self._close()

    def _close(self):
        monitorlib.DestroyPhysicalMonitor(self._pm)
        del self._pm

    def _open(self):
        phys_mons = physical_monitors_from_hmonitor(self._display_handle)
        for pm in phys_mons:
            if pm.desc == self._desc:
                self._pm = pm
                return
        raise Exception("Couldn't retrieve the monitor")
                 
def _get_display_monitors():
    #Enumerates all of the display monitors
    monitors = []
    def _catch_monitor(hmon, dev_ctx, rect, data):
        monitors.append(HANDLE(hmon))
        return 1    #Continue enumerating
    callback_proc = _MONITORENUMPROC(_catch_monitor)
    if not userlib.EnumDisplayMonitors(None, None, callback_proc, None):
        raise Exception("Couldn't enumerate monitors")
    return monitors

def physical_monitors_from_hmonitor(hmon):
    #Returns an array of physical monitors given a display monitor handle
    num_phys_mons = DWORD()
    if not monitorlib.GetNumberOfPhysicalMonitorsFromHMONITOR(hmon,
                                                               byref(num_phys_mons)):
        raise Exception("Couldn't get number of physical monitors")
    phys_mons = (PHYSICAL_MONITOR * num_phys_mons.value)()
    if not monitorlib.GetPhysicalMonitorsFromHMONITOR(hmon,
                                                      num_phys_mons,
                                                      byref(phys_mons)):
        raise Exception("Couldn't get physical monitors")
    return phys_mons

def get_display_device(monitor_handle):
    #Gets information about a display device given its description
    #Handle should be a display monitor_handle
    monitor_info = MONITOR_INFO_EX()
    monitor_info.info.cbSize = sizeof(MONITOR_INFO_EX)
    if not userlib.GetMonitorInfoW(monitor_handle,
                                   byref(monitor_info)):
        raise Exception("Couldn't get monitor information")
    display_device = DISPLAY_DEVICE()
    display_device.cb = sizeof(DISPLAY_DEVICE)
    for i in range(8):
        if not userlib.EnumDisplayDevicesW(LPCWSTR(monitor_info.szDevice),
                                           DWORD(i),
                                           byref(display_device),
                                           DWORD(1)):
            continue
        else:
            return display_device
    raise Exception("Couldn't find the correct display adapter")

def _get_wmi_monitor_names():
    #Gets a list of all names of monitors accessible by WMI
    wmi_server = Wbem.create_server("ROOT\\WMI")
    obj_enum = wmi_server.ExecQuery("WQL", "Select * from WmiMonitorBrightness",
                                    48, None)
    obj, num_returned = obj_enum.Next(-1,1)
    monitor_names = []
    while obj:
        monitor_names.append(obj.Get("InstanceName",0)[0])
        obj, num_returned = obj_enum.Next(-1,1)
    return monitor_names

def get_monitors():
    """Returns a list of all monitor objects

    Returns
    -------
    list
        list of monitor objects.
        monitors could be of the WmiMonitor or VCPMonitor type
    See Also
    --------
    WmiMonitor : representation of a laptop monitor
    VCPMonitor : representation of a desktop monitor
    """
    Wbem._init_security()
    display_monitors = _get_display_monitors()
    mon_objects = []
    wmi_mons = _get_wmi_monitor_names()
    for dm_handle in display_monitors:
        dd_info = get_display_device(dm_handle)
        is_wmi = False
        for name in wmi_mons:
            cleaned_name = name.replace('_0','').split('\\')[2]
            if cleaned_name in dd_info.DeviceID.split('#')[2]:
                mon_objects.append(WmiMonitor(dd_info, name))
                is_wmi = True
                break
        if not is_wmi:
            phys_mons = physical_monitors_from_hmonitor(dm_handle)
            for pm in phys_mons:
                mon_objects.append(VCPMonitor(dm_handle, pm, dd_info, pm.desc))
    return mon_objects
