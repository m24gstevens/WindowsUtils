"""
power.py
====================================
Functions for configuring power plan settings
"""

from ctypes import (windll, byref, cast, c_wchar_p, POINTER, sizeof, c_ubyte,
                    create_string_buffer)
from ctypes.wintypes import DWORD, BYTE
from comtypes import GUID
from WindowsUtils.helpers.powertypes import *

powerlib = windll.PowrProf

class PowerScheme:
    """Represents one of the possible power schemes"""
    def __init__(self, guid):
        self._guid = guid
        dw_namesize = DWORD()
        powerlib.PowerReadFriendlyName(None,byref(guid),None,
                                  None,None,byref(dw_namesize))
        name_buffer = (c_ubyte * dw_namesize.value)()
        status = powerlib.PowerReadFriendlyName(None,byref(guid),None,
                                        None,name_buffer,byref(dw_namesize))
        if status == ERROR_MORE_DATA:
            raise Exception("Couldn't access the power scheme's name")
        wchar_buffer = cast(name_buffer, c_wchar_p)
        self.name = wchar_buffer.value

    @classmethod
    def _from_GUID_string(cls, guid_string):
        return cls(GUID(guid_string))

    def set_this_scheme(self):
        """Switches the active power scheme to this one.

        If the current scheme is already set, will update the settings changes"""
        status = powerlib.PowerSetActiveScheme(None, byref(self._guid))
        if status != ERROR_SUCCESS:
            raise Exception("Couldn't change the power scheme")

    def set_value(self, is_ac, subgroup_name, setting_name, value):
        """Sets a value of a particular power setting.

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        subgroup_name : str
            string name of the settings subgroup the specified setting lives in
            Should be one of the following, which will line up one of the groups
            * Parameter value - Header in documented file
            * 'None'          - Settings belong  to no subgroup
            * 'Disk'          - Hard disk
            * 'Desktop'       - Desktop background settings
            * 'Wireless'      - Wireless adapter settings
            * 'Sleep'         - Sleep
            * 'Buttons'       - Power buttons and lid
            * 'Processor'     - Processor power management
            * 'Display'       - Display
            * 'Energy Saver'  - Energy Saver settings
            * 'Battery'       - Battery
        setting_name : str
            string name of the particular setting, as documented in a generated
            settings text file
        value : int
            integer value to be passed to the setting, as documented in a
            generated settings text file


        See Also
        --------
        gen_power_settings : documents all possible settings
        """
        try:
            subgroup_guid_string, subgroup_settings = subgroups_dict[subgroup_name]
            setting_guid_string = subgroup_settings[setting_name]
        except KeyError:
            raise Exception("Couldn't find the specified setting or subgroup")
        subgroup_guid = GUID(subgroup_guid_string)
        setting_guid = GUID(setting_guid_string)
        write_value = powerlib.PowerWriteACValueIndex if is_ac else powerlib.PowerWriteDCValueIndex
        status = write_value(None,byref(self._guid),
                             byref(subgroup_guid),byref(setting_guid),
                             DWORD(value))
        if status != ERROR_SUCCESS:
            raise Exception("Couldn't change the power policy")
        currently_active = self.active_scheme()
        self.set_this_scheme()
        currently_active.set_this_scheme()

    def read_value(self, is_ac, subgroup_name, setting_name):
        """Gets a value of a particular power setting and its unit

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        subgroup_name : str
            string name of the settings subgroup the specified setting lives in
            Should be one of the following, which will line up one of the groups
            * Parameter value - Header in documented file
            * 'None'          - Settings belong  to no subgroup
            * 'Disk'          - Hard disk
            * 'Desktop'       - Desktop background settings
            * 'Wireless'      - Wireless adapter settings
            * 'Sleep'         - Sleep
            * 'Buttons'       - Power buttons and lid
            * 'Processor'     - Processor power management
            * 'Display'       - Display
            * 'Energy Saver'  - Energy Saver settings
            * 'Battery'       - Battery
        setting_name : str
            string name of the particular setting, as documented in a generated
            settings text file

        Returns
        -------
        int
            integer setting value as documented in a generated settings txt file
        str
            string representing the settings units


        See Also
        --------
        gen_power_settings : documents all possible settings
        """
        try:
            subgroup_guid_string, subgroup_settings = subgroups_dict[subgroup_name]
            setting_guid_string = subgroup_settings[setting_name]
        except KeyError:
            raise Exception("Couldn't find the specified setting or subgroup")
        subgroup_guid = GUID(subgroup_guid_string)
        setting_guid = GUID(setting_guid_string)
        read_it = powerlib.PowerReadACValue if is_ac else powerlib.PowerReadDCValue
        buffer_size = DWORD()
        read_it(None, byref(self._guid),
                   byref(subgroup_guid), byref(setting_guid),
                   None, None, byref(buffer_size))
        value_buffer = create_string_buffer(buffer_size.value)
        value_buffer = cast(value_buffer, POINTER(BYTE))
        status = read_it(None, byref(self._guid),
                            byref(subgroup_guid), byref(setting_guid),
                            None, value_buffer, byref(buffer_size))
        if status != ERROR_SUCCESS:
            raise Exception("Couldn't get the setting value")
        dw_value_p = cast(value_buffer, POINTER(DWORD))
        value = dw_value_p.contents.value
        """
        powerlib.PowerReadValueUnitsSpecifier(None, byref(subgroup_guid), byref(setting_guid),
                                              None, buffer_size)
        units_buffer = create_string_buffer(buffer_size.value)
        units_buffer = cast(units_buffer, POINTER(c_ubyte))
        status = powerlib.PowerReadValueUnitsSpecifier(None,
                                                       byref(subgroup_guid),
                                                       byref(setting_guid),
                                                       units_buffer,
                                                       buffer_size)
        if status != ERROR_SUCCESS:
            raise Exception("Couldn't get the setting units")
        units_wbuffer = cast(units_buffer, POINTER(c_wchar_p))"""
        return value, "units"

    @classmethod
    def MaxPowerSavings(cls):
        """Returns a PowerScheme instance representing the maximum
        power savings policy
        
        Returns
        -------
        PowerScheme
            object representing maximum power saving scheme
        """
        return cls._from_GUID_string(GUID_MAX_POWER_SAVINGS)

    @classmethod
    def MinPowerSavings(cls):
        """Returns a PowerScheme instance representing the minimum
        power savings policy
        
        Returns
        -------
        PowerScheme
            object representing minimum power saving scheme
        """
        return cls._from_GUID_string(GUID_MIN_POWER_SAVINGS)

    @classmethod
    def TypicalPowerSavings(cls):
        """Returns a PowerScheme instance representing the typical
        power savings policy

        Returns
        -------
        PowerScheme
            object representing typical power saving scheme
        """
        return cls._from_GUID_string(GUID_TYPICAL_POWER_SAVINGS)

    @classmethod
    def active_scheme(cls):
        """Returns a PowerScheme instance representing the active power
        plan

        Returns
        -------
        PowerScheme
            object representing active scheme
        """
        p_guid = POINTER(GUID)()
        status = powerlib.PowerGetActiveScheme(None, byref(p_guid))
        if status != ERROR_SUCCESS:
            raise Exception("Couldn't get the current power scheme")
        return PowerScheme(p_guid.contents)
        

    def sleep_after(self, is_ac, wait_time = 0):
        """Sets the time after which the computer will sleep

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        wait_time : int, optional
            number of seconds after which the computer will sleep.
            Defaults to stay awake indefinitely
        """
        self.set_value(is_ac, "Sleep", "Sleep after", max(wait_time, 0))

    def get_sleep(self, is_ac):
        """Gets the time after which the computer will sleep

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        int
            number of seconds after which the computer will sleep.
            zero value indicates the computer never sleeps
        """
        val, units = self.read_value(is_ac, "Sleep", "Sleep after")
        return val

    def display_off_after(self, is_ac, wait_time = 0):
        """Sets the time after which the computer will turn its display off

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        wait_time : int, optional
            number of seconds after which the computer will shut off display.
            Defaults to stay on indefinitely
        """
        self.set_value(is_ac, "Display", "Turn off display after", max(wait_time, 0))

    def get_display_off(self, is_ac):
        """Gets the time after which the computer will turn off display

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        int
            number of seconds after which the computer will turn off its display
            zero value indicates the computer never shuts off display
        """
        val, units = self.read_value(is_ac, "Display", "Turn off display after")
        return val
    
    def dim_display_after(self, is_ac, wait_time = 0):
        """Sets the time after which the computer will dim its display

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        wait_time : int, optional
            number of seconds after which the computer will dim display
            Defaults to stay awake indefinitely
        """
        self.set_value(is_ac, "Display", "Dim display after", max(wait_time, 0))

    def get_dim_display(self, is_ac):
        """Gets the time after which the computer will dim its display

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        int
            number of seconds after which the computer will dim its display
            zero value indicates the computer never dims
        """
        val, units = self.read_value(is_ac, "Display", "Dim display after")
        return val

    #def set_aggressive_energy_save(self, is_ac, resp):
        """Sets the aggressive energy saving mode.

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        resp : bool
            set to True to enable aggressive energy saving, False to disable
        """
        #self.set_value(is_ac, "Energy Saver", "Energy Saver Policy",
                       #1 if resp else 0)

    #def get_aggressive_energy_save(self, is_ac):
        """Gets the policy on aggressive energy saving

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        bool
            True indicates that aggressive energy saving is enabled
        """
        #val, units = self.read_value(is_ac, "Energy Saver", "Energy Saver Policy")
        #return bool(val)
    
    def set_low_battery_level(self, is_ac, level):
        """Sets the percentage at which the system switches to low battery mode

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        level : int
            % at which low battery mode will enable
        """
        self.set_value(is_ac, "Battery", "Low battery level",
                       min(max(level,0),100))
        
    def get_low_battery_level(self, is_ac):
        """Gets the low battery threshold

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        int
           percentage at which the computer goes into low battery mode
        """
        val, units = self.read_value(is_ac, "Battery", "Low battery level")
        return val

    def critical_battery_level(self, is_ac, level):
        """Sets the percentage at which the system switches to critical
        battery mode

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        level : int
            % at which critical battery mode will enable
        """
        self.set_value(is_ac, "Battery", "Critical battery level",
                       min(max(level,0),100))

    def get_critical_battery_level(self, is_ac):
        """Gets the critical battery threshold

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        int
           percentage at which the computer goes into low battery mode
        """
        val, units = self.read_value(is_ac, "Battery", "Critical battery level")
        return val

    def set_active_cooling(self, is_ac, resp):
        """Sets the active cooling policy.

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        resp : bool
            set to True to enable active cooling, False for passive
        """
        self.set_value(is_ac, "Processor", "System cooling policy",
                       0 if resp else 1)

    def get_active_cooling(self, is_ac):
        """Gets whether the cooling policy is active or passive

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        bool
            True indicates that active cooling is enabled
        """
        val, units = self.read_value(is_ac, "Processor", "System cooling policy")
        return bool(val)

    def set_maximum_processor_state(self, is_ac, level):
        """Sets the maximum percentage of processor resources to use

        Parameters
        ----------
        is_ac : bool
            set to True to alter an AC setting, otherwise will alter DC
        level : int
            percentage of processor resources that will be maximum limit
        """
        self.set_value(is_ac, "Processor", "Maximum processor state",
                       min(max(level, 5), 100))

    def get_maximum_processor_state(self, is_ac, level):
        """Gets the maximum percentage of processor resources used

        Parameters
        ----------
        is_ac : bool
            set to True to get an AC setting, otherwise will get DC

        Returns
        -------
        int
           maximum percentage of resouces used
        """
        val, units = self.read_value(is_ac, "Processor", "Maximum processor state")
        return val  
        

def get_all_schemes():
    """Returns a list of all accessable power schemes

    Returns
    -------
    list
        a list of accessable power schemes, encapsulated as PowerScheme objects
    """
    buffer = GUID()
    buffer_size = DWORD(sizeof(buffer))
    idx = 0
    schemes = []
    while True:
        status = powerlib.PowerEnumerate(None, None, None, DWORD(16),
                                         DWORD(idx),
                                         cast(byref(buffer), POINTER(c_ubyte)),
                                         byref(buffer_size))
        if status == ERROR_SUCCESS:
            schemes.append(PowerScheme(buffer))
            idx += 1
        else:
            break
    return schemes


def gen_power_settings():
    """Generates a text file documenting all possible power settings that
    could be changed on a profile"""
    import os
    os.system("powercfg /Q > PyPowerSettings.txt")
