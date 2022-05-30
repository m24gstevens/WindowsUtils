from ctypes import (Structure, c_ulong, c_ushort, c_char, c_uint, c_int, c_long,
                    c_ulonglong, c_void_p, WINFUNCTYPE, POINTER)
from ctypes.wintypes import WCHAR, DWORD, BOOL, ULONG, LPCWSTR

ERROR_SUCCESS = 0
ERROR_FILE_NOT_FOUND = 2

class GUID(Structure):
    _fields_ = [("Data1", c_ulong),
                ("Data2", c_ushort),
                ("Data3", c_ushort),
                ("Data4", c_char * 8)]

#Wlan Interface State Enum
(WLAN_NOT_READY, WLAN_CONNECTED, WLAN_AD_HOC, WLAN_DISCONNECTING,
 WLAN_DISCONNECTED, WLAN_ASSOCIATING, WLAN_DISCOVERING, WLAN_AUTH) = range(8)
state_descriptions = {WLAN_NOT_READY: 'not ready',
                      WLAN_CONNECTED: 'connected',
                      WLAN_AD_HOC: 'ad-hoc network formed',
                      WLAN_DISCONNECTING: 'disconnecting',
                      WLAN_DISCONNECTED: 'disconnected',
                      WLAN_ASSOCIATING: 'associating',
                      WLAN_DISCOVERING: 'discovering',
                      WLAN_AUTH: 'authenticating'}


#Interface info (+ list) Structures
class _WLAN_INTERFACE_INFO(Structure):
    _fields_ = [("GUID", GUID),
                ("Desc", WCHAR * 256),
                ("State", c_uint)]

class WLAN_INTERFACE_INFO_LIST(Structure):
    _fields_ = [("dwNumberOfItems", DWORD),
                ("dwIndex", DWORD),
                ("InterfaceInfo", _WLAN_INTERFACE_INFO * 0)]

#Opcode Value Type Enum
(TYPE_QUERY_ONLY, SET_BY_GROUP_POLICY, SET_BY_USER, INVALID) = range(4)

#Wlan Radio State Structure and enums
(STATE_UNKNOWN, STATE_ON, STATE_OFF) = range(3)

class _WLAN_PHY_RADIO_STATE(Structure):
    _fields_ = [("dwPhyIndex", DWORD),
                ("dot11SoftwareRadioState", c_int),
                ("dot11HardwareRadioState", c_int)]

class WLAN_RADIO_STATE(Structure):
    _fields_ = [("dwNumberOfPhys", DWORD),
                ("PhyRadioState", _WLAN_PHY_RADIO_STATE * 64)]

#Auth-Cipher Pair (+ List) Structure

class _DOT11_AUTH_CIPHER_PAIR(Structure):
    _fields_ = [("AuthAlgoID", c_long),
                ("CipherAlgoID", c_long)]

class WLAN_AUTH_CIPHER_PAIR_LIST(Structure):
    _fields_ = [("dwNumberOfItems", DWORD),
                ("pAuthCipherPairList", _DOT11_AUTH_CIPHER_PAIR * 0)]

#Country Or Region (+ List) Structure

_DOT11_COUNTRY_OR_REGION_STRING = c_char * 3

class WLAN_COUNTRY_OR_REGION_STRING_LIST(Structure):
    _fields_ = [("dwNumberOfItems", DWORD),
                ("pCountryOrRegionStringList", _DOT11_COUNTRY_OR_REGION_STRING * 0)]
                     
#Wlan Statistics Structure

class _WLAN_MAC_FRAME_STATISTICS(Structure):
    _fields_ = [("ullTransmittedFrameCount", c_ulonglong),
                ("ullReceivedFrameCount", c_ulonglong),
                ("ullWEPExcludedCount", c_ulonglong),
                ("ullTKIPLocalMICFailures", c_ulonglong),
                ("ullTKIPReplays", c_ulonglong),
                ("ullTKIPICVErrorCount", c_ulonglong),
                ("ullCCMPReplays", c_ulonglong),
                ("ullCCMPDecryptErrors", c_ulonglong),
                ("ullWEPUndecryptableCount", c_ulonglong),
                ("ullWEPICVErrorCount", c_ulonglong),
                ("ullDecryptSuccessCount", c_ulonglong),
                ("ullDecryptFailureCount", c_ulonglong)]

class _WLAN_PHY_FRAME_STATISTICS(Structure):
    _fields_ = [("ullTransmittedFrameCount", c_ulonglong),
                ("ullMulticastTransmittedFrameCount", c_ulonglong),
                ("ullFailedCount", c_ulonglong),
                ("ullRetryCount", c_ulonglong),
                ("ullMultipleRetryCount", c_ulonglong),
                ("ullMaxTXLifetimeExceededCount", c_ulonglong),
                ("ullTransmittedFragmentCount", c_ulonglong),
                ("ullRTSSuccessCount", c_ulonglong),
                ("ullRTSFailureCount", c_ulonglong),
                ("ullACKFailureCount", c_ulonglong),
                ("ullReceivedFrameCount", c_ulonglong),
                ("ullMulticastReceivedFrameCount", c_ulonglong),
                ("ullPromiscuousReceivedFrameCount", c_ulonglong),
                ("ullMaxRXLifetimeExceededCount", c_ulonglong),
                ("ullFrameDuplicateCount", c_ulonglong),
                ("ullReceivedFragmentCount", c_ulonglong),
                ("ullPromiscuousReceivedFragmentCount", c_ulonglong),
                ("ullFCSErrorCount", c_ulonglong)]

class WLAN_STATISTICS(Structure):
    _fields_ = [("ullFourWayHandshakeFailures", c_ulonglong),
                ("ullTKIPCounterMeasuresInvoked", c_ulonglong),
                ("ullReserved", c_ulonglong),
                ("MacUcastCounters", _WLAN_MAC_FRAME_STATISTICS),
                ("MacMcastCounters", _WLAN_MAC_FRAME_STATISTICS),
                ("dwNumberOfPhys", DWORD),
                ("PhyCounters", _WLAN_PHY_FRAME_STATISTICS * 0)]

#Wlan Connection Mode enum
(MODE_PROFILE, MODE_TEMP_PROFILE, MODE_DISC_SECURE, MODE_DISC_UNSECURE,
 MODE_AUTO, MODE_INVALID) = range(6)
connection_modes = {
    MODE_PROFILE: 'Profile',
    MODE_TEMP_PROFILE: 'Temporary profile',
    MODE_DISC_SECURE: 'Secure Discovery',
    MODE_DISC_UNSECURE: 'Unsecure Discovery',
    MODE_AUTO: 'Automatic',
    MODE_INVALID: 'Invalid'
    }

#SSIDs
class _DOT11_SSID(Structure):
    _fields_ = [("uSSIDLength", c_ulong),
                ("ucSSID", c_char * 32)]

#Dot11 BSS Type Enum
(BSS_TYPE_INFRASTRUCTURE, BSS_TYPE_INDEPENDENT, BSS_TYPE_ANY) = range(1,4)
#Dot11 Phy Type Enum
PHY_TYPE_UNKNOWN = 0
(PHY_TYPE_ANY, PHY_TYPE_FHSS, PHY_TYPE_DSSS, PHY_TYPE_IRBASEBAND,
 PHY_TYPE_OFDM, PHY_TYPE_HRDSS, PHY_TYPE_ERP, PHY_TYPE_HT, PHY_TYPE_VHT) = range(9)
PHY_TYPE_IHV_START = 0x80000000
PHY_TYPE_IHV_END = 0xFFFFFFFF


#Dot11 Auth Algorithm Enum
(ALGO_80211_OPEN, ALGO_80211_SHARED_KEY, ALGO_WPA, ALGO_WPA_PSK,
 ALSO_WPA_NONE, ALGO_RSNA, ALGO_PSNA_PSK) = range(1,8)
ALGO_IHV_START = 0x80000000
ALGO_IHV_END = 0xFFFFFFFF

#Dot11 Cipher Algorithm Enum
C_ALGO_NONE = 0
C_ALGO_WEP40 = 1
C_ALGO_TKIP = 2
C_ALGO_CCMP = 4
C_ALGO_WEP104 = 5
C_ALGO_WPA_USE_GROUP = 0x100
C_ALGO_RSN_USE_GROUP = 0x100
C_ALGO_WEP = 0x101
C_ALGO_IHV_START = 0x80000000
C_ALGO_IHV_END = 0xFFFFFFFF

#Security and connection attribues for Connection Parameters
class WLAN_SECURITY_ATTRIBUTES(Structure):
    _fields_ = [("bSecurityEnabled", BOOL),
                ("bOneXEnabled", BOOL),
                ("dot11AuthAlgorithm", c_long),
                ("dot11CipherAlgorithm", c_long)]

class WLAN_ASSOCIATION_ATTRIBUTES(Structure):
    _fields_ = [("dot11Ssid", _DOT11_SSID),
                ("dot11BssType", c_uint),
                ("dot11Bssid", c_char * 6),
                ("dot11PhyType", c_long),
                ("uDot11PhyIndex", c_ulong),
                ("wlanSignalQuality", c_ulong),
                ("ulRxRate", c_ulong),
                ("ulTxRate", c_ulong)]


class WLAN_CONNECTION_ATTRIBUTES(Structure):
    _fields_ = [("isState", c_uint),
                ("wlanConnectionMode", c_uint),
                ("strProfileName", WCHAR * 256),
                ("wlanAssociationAttributes", WLAN_ASSOCIATION_ATTRIBUTES),
                ("wlanSecurityAttributes", WLAN_SECURITY_ATTRIBUTES)]

#Access masks
READ_ACCESS = 0x20001
EXECUTE_ACCESS = 0x20021
WRITE_ACCESS = 0x70023

#Query Interface Opcode Dict - Key:Value = 'Option':(code value, data out)
query_opcode_dict = {'autoconf_enabled':(1, BOOL),
                     'background_scan_enabled': (2, BOOL),
                     'radio_state': (4, WLAN_RADIO_STATE),
                     'bss_type': (5, c_int),
                     'interface_type': (6, c_int),
                     'current_connection': (7, WLAN_CONNECTION_ATTRIBUTES),
                     'channel_number': (8, c_ulong),
                     'supported_infrastructure_auth_cipher_pairs': (9, WLAN_AUTH_CIPHER_PAIR_LIST),
                     'supported_adhoc_auth_cipher_pairs': (10, WLAN_AUTH_CIPHER_PAIR_LIST),
                     'supported_country_or_region_list_string': (11, WLAN_COUNTRY_OR_REGION_STRING_LIST),
                     'media_streaming_mode': (3, BOOL),
                     'statistics': (0x100000101, WLAN_STATISTICS),
                     'current_operation_mode': (12, ULONG),
                     'supported_safe_mode': (13, BOOL),
                     'certified_safe_mode': (14, BOOL)}


#Available network (+ List) Structures

class _WLAN_AVAILABLE_NETWORK(Structure):
    _fields_ = [("strProfileName", WCHAR * 256),
                ("dot11Ssid", _DOT11_SSID),
                ("dot11BssType", c_uint),
                ("uNumberOfBssids", c_ulong),
                ("bNetworkConnectable", BOOL),
                ("wlanNotConnectableReason", DWORD),
                ("uNumberOfPhyTypes", c_ulong),
                ("dot11PhyTypes", c_long * 8),
                ("bMorePhyTypes", BOOL),
                ("wlanSignalQuality", c_ulong),
                ("bSecurityEnabled", BOOL),
                ("dot11DefaultAuthAlgorithm", c_long),
                ("dot11DefaultCipherAlgorithm", c_long),
                ("dwFlags", DWORD),
                ("dwReserved", DWORD)] 

class WLAN_AVAILABLE_NETWORK_LIST(Structure):
    _fields_ = [("dwNumberOfItems", DWORD),
                ("dwIndex", DWORD),
                ("Network", _WLAN_AVAILABLE_NETWORK * 0)]

#Notification information for the notification callback

class WLAN_NOTIFICATION_DATA(Structure):
    """See the WLAN_NOTIFICATION_DATA Struct on MSDN for more info"""
    _fields_ = [("NotificationSource", DWORD),
                ("NotificationCode", DWORD),
                ("InterfaceGUID", GUID),
                ("dwDataSize", DWORD),
                ("pData", c_void_p)]

#Notification source
SOURCE_NONE = 0x0
SOURCE_ONEX = 0x4
SOURCE_ACM = 0x8
SOURCE_MSM = 0x10
SOURCE_SECURITY = 0x20
SOURCE_IHV = 0x40
SOURCE_HNWK = 0x80
SOURCE_ALL = 0xFFFF


onex_codes = {
    1: "OneXNotificationTypeResultUpdated",
    2: "OneXNotificationTypeAuthRestarted",
    3: "EventInvalid"
    }

acm_codes = {
    1: "acm_autoconf_enabled",
    2: "acm_autoconf_disabled",
    3: "acm_background_scan_enabled",
    4: "acm_background_scan_disabled",
    5: "acm_bss_type_change",
    6: "acm_power_setting_change",
    7: "acm_scan_complete",
    8: "acm_scan_failed",
    9: "acm_connection_start",
    10: "acm_connection_complete",
    11: "acm_connection_attempt_fail",
    12: "acm_filter_list_change",
    13: "acm_interface_arrival",
    14: "acm_interface_removal",
    15: "acm_profile_change",
    16: "acm_profile_name_change",
    17: "acm_profiles_exhausted",
    18: "acm_network_not_available",
    19: "acm_network_available",
    20: "acm_network_disconnecting",
    21: "acm_disconnected",
    22: "acm_adhoc_network_state_change",
    23: "acm_profile_unblocked",
    24: "acm_screen_power_change",
    25: "acm_profile_blocked",
    26: "acm_scan_list_refresh"}

msm_codes = {
    1: "msm_associating",
    2: "msm_associated",
    3: "msm_authenticating",
    4: "msm_connected",
    5: "msm_roaming_start",
    6: "msm_roaming_end",
    7: "msm_radio_state_change",
    8: "msm_signal_quality_change",
    9: "msm_disassociating",
    10: "msm_disconnected",
    11: "msm_peer_join",
    12: "msm_peer_leave",
    13: "msm_adapter_removal",
    14: "msm_adapter_operation_mode_change"
    }

hnwk_codes = {
    0x1000: "hosted_network_state_change",
    0x1001: "hosted_network_peer_state_change",
    0x1002: "hosted_network_radio_state_change"
    }

source_dict_codes = {
    SOURCE_ONEX: onex_codes,
    SOURCE_ACM: acm_codes,
    SOURCE_MSM: msm_codes,
    SOURCE_HNWK: hnwk_codes
    }

    


#Notification Handling

_WlanNotificationCallback = WINFUNCTYPE(None, POINTER(WLAN_NOTIFICATION_DATA), c_void_p)

def basic_callback(pwlan_data, pvoid):
    src = pwlan_data.contents.NotificationSource
    code = pwlan_data.contents.NotificationCode
    if src in source_dict_codes.keys():
        print(source_dict_codes[src][code])


#Connection parameters Structure

class _NDIS_OBJECT_HEADER(Structure):
    _fields_ = [("Type", c_char),
                ("Revision", c_char),
                ("Size", c_ushort)]
    
class _DOT11_BSSID_LIST(Structure):
    _fields_ = [("Header", _NDIS_OBJECT_HEADER),
                ("uNumOfEntries", c_ulong),
                ("uTotalNumOfEntries", c_ulong),
                ("BSSIDs", (c_char * 6) * 1)]

class WLAN_CONNECTION_PARAMETERS(Structure):
    _fields_ = [("wlanConnectionMode", c_int),
                ("strProfile", LPCWSTR),
                ("pDot11Ssid", POINTER(_DOT11_SSID)),
                ("pDesiredBssidList", POINTER(_DOT11_BSSID_LIST)),
                ("dot11BssType", c_int),
                ("dwFlags", DWORD)]

#Connection flags
HIDDEN_NETWORK = 0x1
ADHOC_JOIN_ONLY = 0x2
IGNORE_PRIVACY_BIT = 0x4
EAPOL_PASSTHROUGH = 0x8
PERSIST_DISCOVERY_PROFILE = 0x10
PERSIST_DISCOVERY_PROFILE_CONNECTION_MODE_AUTO = 0x20
PERSIST_DISCOVERY_PROFILE_OVERWRITE_EXISTING = 0x40


#Profile info (+ List) Structures

class _WLAN_PROFILE_INFO(Structure):
    _fields_ = [("strProfileName", WCHAR * 256),
                ("dwFlags", DWORD)]

class WLAN_PROFILE_INFO_LIST(Structure):
    _fields_ = [("dwNumberOfItems", DWORD),
                ("dwIndex", DWORD),
                ("ProfileInfo", _WLAN_PROFILE_INFO * 0)]


#BSS Entry (+ List) Structures

class _WLAN_RATE_SET(Structure):
    _fields_ = [("uRateSetLength", c_ulong),
                ("usRateSet", c_ushort * 126)]

class _WLAN_BSS_ENTRY(Structure):
    _fields_ = [("dot11Ssid", _DOT11_SSID),
                ("uPhyID", c_ulong),
                ("dot11Bssid", c_char * 6),
                ("dot11BssType", c_uint),
                ("dot11BssPhyType", c_long),
                ("lRssi", c_long),
                ("iLinkQuality", c_ulong),
                ("bInRegDomain", c_char),
                ("usBeaconPeriod", c_ushort),
                ("ullTimestamp", c_ulonglong),
                ("ullHostTimestamp", c_ulonglong),
                ("isCapabilityInformation", c_ushort),
                ("ulChCenterFrequency", c_ulong),
                ("wlanRateSet", _WLAN_RATE_SET),
                ("ulIeOffset", c_ulong),
                ("ulIeSize", c_ulong)]

class WLAN_BSS_LIST(Structure):
    _fields_ = [("dwTotalSize", DWORD),
                ("dwNumberOfItems", DWORD),
                ("wlanBssEntries", _WLAN_BSS_ENTRY * 0)]
