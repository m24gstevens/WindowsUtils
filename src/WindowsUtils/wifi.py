"""
wifi.py
====================================
Functions for configuring a network session
"""

from ctypes import windll, cast, byref, POINTER, c_int, pointer
from ctypes.wintypes import LPWSTR, DWORD, HANDLE

def _array_from_list_struct(array_field, arr_item_class, num_items):
    #Casts an "empty array" found in structures to a proper array given
    #the number of objects allocated
    p_array = cast(byref(array_field), POINTER(arr_item_class * num_items))
    return p_array.contents

def process_exit_code(exit_code):
    #Gets last error message from the system
    if exit_code == ERROR_SUCCESS:
        return
    messagebuf = LPWSTR()
    size = windll.Kernel32.FormatMessageA(0x1300,
                                          None,
                                          exit_code,
                                          0,
                                          cast(byref(messagebuf), LPWSTR),
                                          0,
                                          None)
    #Get last error message into a buffer, in user/system default or english,
    #with the function allocating the buffer.
    raise Exception(messagebuf.value)

def simple_exit_code(exit_code, message):
    #Helper method
    if exit_code != ERROR_SUCCESS:
        raise Exception(message)                                                                     
                                
    
class AvailableNetwork:
    """Represents the properties of an available wifi network"""
    def __init__(self,struct):
        self.is_profile = bool(struct.strProfileName)
        self.name = struct.strProfileName if self.is_profile else struct.dot11Ssid.ucSSID.decode()
        self.security_enabled = bool(struct.bSecurityEnabled)
        self.is_connectable = bool(struct.bNetworkConnectable)
        self.signal_quality = (struct.wlanSignalQuality + 50) * 2
        self.dot11_ssid = pointer(struct.dot11Ssid)
        self.bss_type = struct.dot11BssType
        self.connection_mode = MODE_PROFILE if self.is_profile else MODE_DISC_SECURE


class WifiSession:
    """Represents a configuration session.

    Note: Should be instantiated as a context manager"""
    def __init__(self, client_ver=2):
        self.client_ver = client_ver

    def __enter__(self):
        #Opens a use handle to open a connection to the server
        use_ver = DWORD()
        use_handle = HANDLE()
        status = WAPI.WlanOpenHandle(DWORD(self.client_ver),
                                     None,
                                     byref(use_ver),
                                     byref(use_handle))
        if status != ERROR_SUCCESS:
            raise Exception("Failed to start a session")
        self._use_handle = use_handle
        self._allocated_objects = []

    def __exit__(self):
        #Closes a use handle and frees any memory allocated
        status = WAPI.WlanCloseHandle(self._use_handle, None)
        if status != ERROR_SUCCESS:
            raise Exception("Failed to close handle")
        del self._use_handle
        if getattr(self, 'interfaces', False):
            for interface in self.interfaces:
                interface._cleanup()
        for ptr in self._allocated_objects:
            WAPI.WlanFreeMemory(ptr)
        del self._allocated_objects

    def _allocated(self, *ptrs):
        #Stores allocated pointers to be freed later
        for ptr in ptrs:
            self._allocated_objects.append(ptr)

    def enum_interfaces(self):
        """Retrieves all interface instance objects

        Returns
        -------
        WlanInterface
            interface used to manage connection specific properties
        """
        iilp = POINTER(WLAN_INTERFACE_INFO_LIST)()
        status = WAPI.WlanEnumInterfaces(self._use_handle, None, byref(iilp))
        simple_exit_code(status, "Couldn't Access Interfaces")
        info_array = _array_from_list_struct(iilp.contents.InterfaceInfo,
                                             _WLAN_INTERFACE_INFO,
                                             iilp.contents.dwNumberOfItems)
        self._allocated(iilp)
        self.interfaces = [WlanInterface(itf, self._use_handle) for itf in list(info_array)]
        return self.interfaces

    def get_security_settings(self, opcode):
        """Gets security descriptor string associated with a particular object

        Parameters
        ----------
        opcode : int
            operation code of the setting to query.
            One of 17 SECURE_* values

        Returns
        -------
        str
            SDDL security descriptor string describing the queried setting"""
        security_descriptor = LPWSTR()
        granted_access = DWORD()
        settings_source = c_int()
        status = WAPI.WlanGetSecuritySettings(self._use_handle,
                                              c_int(opcode),
                                              byref(settings_source),
                                              byref(security_descriptor),
                                              byref(granted_access))
        simple_exit_code(status, "Invalid Opcode")
        access_list = []
        if granted_access.value == READ_ACCESS:
            access = 'read'
        elif granted_access.value == EXECUTE_ACCESS:
            access = 'execute'
        else:
            access = 'write'

        if settings_source == SET_BY_GROUP_POLICY:
            setter = 'group policy'
        else:
            setter = 'user'

        self._allocated(security_descriptor)
        return security_descriptor.value, setter, access

    def register_notifications(self, source = SOURCE_ALL, callback_func = basic_callback):
        """Registers for notifications on all wireless information.
        Results in all notifications being printed to stdout

        Should be called with no arguments: Optional arguments are for
        future versions"""
        #flags indicating notification source to be registered
        #A callback (python) function, taking arguments of
        #--pointer to a WLAN_NOTIFICATION_DATA structure containing notification
            #data
        #--pointer to context information, which for most cases should be ignored
        #By default, we register for all notifications, and provide a simple callback
        #printing out the type of ACM notification"""

        WlanNotificationCallback = _WlanNotificationCallback(callback_func)

        status = WAPI.WlanRegisterNotification(self._use_handle,
                                               c_int(source),
                                               c_int(1),    #Ignore duplicates
                                               WlanNotificationCallback,
                                               None,
                                               None,
                                               None)
        simple_exit_code(status, "Failed to register for notifications")

class WlanInterface:
    """Represents a particular wireless interface"""
    def __init__(self, interface_struct, handle):
        self._guid = interface_struct.GUID
        self.state = interface_struct.State
        self.description = interface_struct.Desc
        self._use_handle = handle
        self._allocated_objects = []

    def _allocated(self, *ptrs):
        #Stores allocated pointers to be freed later
        for ptr in ptrs:
            self._allocated_objects.append(ptr)

    def _cleanup(self):
        #Cleans up allocated resources
        for ptr in self._allocated_objects:
            WAPI.WlanFreeMemory(ptr)
        del self._allocated_objects

    def _query_interface(self, field):
        #Queries an interface for its given parameters. The field argument
        #must be string to use as a key in the query opcode dictionary.
        #Returns the data object
        try:
            opcode, datatype = query_opcode_dict[field]
        except Exception:
            raise Exception("Invalid query")
        data_pointer = POINTER(datatype)()
        data_size = DWORD()
        status = WAPI.WlanQueryInterface(self._use_handle,
                                         byref(self._guid),
                                         c_int(opcode),
                                         None,
                                         byref(data_size),
                                         byref(data_pointer),
                                         None)
        simple_exit_code(status, "Interface Query Failed")
        self._allocated(data_pointer)
        return data_pointer.contents
    
    def query_connection_attributes(self):
        """Returns information about the current connection

        Returns
        -------
        dict
            dictionary describing current connection state.
            keys:
            * 'Interface state' : str
                description of the interface's state
            * 'Connection mode' : str
                description of the current connection mode
            * 'Profile name' : str
                if connected to a profile, gives profile name
            * 'Security enabled' : bool
                True if security is enabled on this interface
        """
        cattrs = self._query_interface('current_connection')
        connection_dict = {}
        connection_dict['Interface state'] = state_descriptions[cattrs.isState]
        connection_dict['Connection mode'] = connection_modes[cattrs.wlanConnectionMode]
        connection_dict['Profile name'] = cattrs.strProfileName
        connection_dict['Security enabled'] = bool(cattrs.wlanSecurityAttributes.bSecurityEnabled)
        return connection_dict

    def _scan(self):
        #Requests a scan for available networks on an interface,
        #specified by its GUID
        status = WAPI.WlanScan(self._use_handle,
                               byref(self._guid),
                               None,
                               None,
                               None)
        simple_exit_code(status, "Couldn't request a scan")

    def get_available_network_list(self, adhoc = True, hidden = True):
        """Returns all available networks

        Returns
        -------
        list
            a list of AvailableNetwork instances representing the networks
        """
        self._scan()
        network_flags = (1 if adhoc else 0) | (2 if hidden else 0)
        network_list = POINTER(WLAN_AVAILABLE_NETWORK_LIST)()
        status = WAPI.WlanGetAvailableNetworkList(self._use_handle,
                                                  byref(self._guid),
                                                  c_int(network_flags),
                                                  None,
                                                  byref(network_list))
        simple_exit_code(status, "Failed in scan of available networks")
        self._allocated(network_list)
        network_array = _array_from_list_struct(network_list.contents.Network,
                                                _WLAN_AVAILABLE_NETWORK,
                                                network_list.contents.dwNumberOfItems)
        return [AvailableNetwork(nwk) for nwk in list(network_array)]

    def get_profile_list(self):
        """Returns a list of the network profiles stored by the user

        Returns
        -------
        list
            a list of names of profiles
        """
        profile_list = POINTER(WLAN_PROFILE_INFO_LIST)()
        status = WAPI.WlanGetProfileList(self._use_handle,
                                         byref(self._guid),
                                         None,
                                         byref(profile_list))
        simple_exit_code(status, "Failed to get a list of the profiles")
        self._allocated(profile_list)
        profile_array = _array_from_list_struct(profile_list.contents.ProfileInfo,
                                                _WLAN_PROFILE_INFO,
                                                profile_list.contents.dwNumberOfItems)
        return [prof.strProfileName for prof in list(profile_array)]

    def get_profile(self, profile_name):
        """Searches for a profile with the given name

        Parameters
        ----------
        profile_name : str
            a value in the list returned from get_profile_list

        Returns
        -------
        str
            an XML string representation of the profile
        """
        granted_access = DWORD()
        xml = LPWSTR()
        flags = DWORD(PROFILE_GROUP_POLICY)
        status = WAPI.WlanGetProfile(self._use_handle,
                                byref(self._guid),
                                LPCWSTR(profile_name),
                                None,
                                byref(xml),
                                byref(flags),
                                byref(granted_access))
        if status == ERROR_FILE_NOT_FOUND:
            raise Exception("Profile Not Found")
        elif status != ERROR_SUCCESS:
            raise Exception("Could not get profile")
        self._allocated(xml)
        if granted_access.value == READ_ACCESS:
            access = 'read'
        elif granted_access.value == EXECUTE_ACCESS: #Execute:
            access = 'execute'
        else:
            access = 'write'

        return xml.value, access

    def disconnect(self):
        """Disconnects from the current network"""
        status = WAPI.WlanDisconnect(self._use_handle,
                                     byref(self._guid),
                                     None)
        simple_exit_code(status, "Failed to disconnect")

    def connect(self, available_network):
        """Attempts a connection to a network without attempting authentication

        Due to lack of security, you should be connecting to networks with
        available profiles.
        
        Parameters
        ----------
        available_network : AvailableNetwork
            an AvailableNetwork instance we attempt connection to
        """
        conn_params = WLAN_CONNECTION_PARAMETERS()
        conn_params.wlanConnectionMode = available_network.connection_mode
        conn_params.strProfile = available_network.name
        conn_params.pDot11Ssid = available_network.dot11_ssid
        conn_params.pDesiredBssidList = None
        conn_params.dot11BssType = available_network.bss_type
        conn_params.dwFlags = HIDDEN_NETWORK
        status = windll.Wlanapi.WlanConnect(self._use_handle,
                                            byref(self._guid),
                                            byref(conn_params),
                                            None)
        simple_exit_code(status, "Attempt to connect failed")
