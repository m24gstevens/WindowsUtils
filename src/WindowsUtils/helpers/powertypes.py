from ctypes import Structure, c_ulong, c_ushort, c_char, oledll, byref

ERROR_SUCCESS = 0
ERROR_MORE_DATA = 234

"""class GUID(Structure):
    _fields_ = [("Data1", c_ulong),
                ("Data2", c_ushort),
                ("Data3", c_ushort),
                ("Data4", c_char * 8)]
    def __init__(self, name = None):
        if name is not None:
            oledll.ole32.CLSIDFromString(name, byref(self))"""

GUID_MAX_POWER_SAVINGS = "{A1841308-3541-4FAB-BC81-F71556F20B4A}"
GUID_MIN_POWER_SAVINGS = "{8C5E7fDA-E8BF-4A96-9A85-A6E23A8C635C}"
GUID_TYPICAL_POWER_SAVINGS = "{381B4222-F694-41F0-9685-FF5BB260DF2E}"

sub_none_dict = {
    "Power plan type": "{245d8541-3943-4422-b025-13a784f679b7}",
    "Device idle policy": "{4faab71a-92e5-4726-b531-224559672d19}"
    }

sub_none_dict = {key : value.upper() for key, value in sub_none_dict.items()}

sub_disk_dict = {"Turn off hard disk after" : "{6738E2C4-E8A5-4A42-B16A-E040E769756E}"}

sub_desk_dict = {"Slide Show" : "{309DCE9B-BEF4-4119-9921-A851FB12F0F4}"}

sub_wire_dict = {"Power Saving Mode" : "{12BBEBE6-58D6-4636-95BB-3217EF867C1A}"}

sub_sleep_dict = {"Sleep after" : "{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}",
                  "System unattended sleep timeout" : "{7bc4a2f9-d8fc-4469-b07b-33eb785aaca0}",
                  "Allow hybrid sleep" : "{94ac6d29-73ce-41a6-809f-6363ba21b47e}",
                  "Hibernate after" : "{9d7815a6-7ee4-497e-8888-515a05f02364}",
                  "Allow Standby States" : "{abfc2519-3608-4c2a-94ea-171b0ed546ab}",
                  "Allow wake timers" : "{bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d}"
                  }
sub_sleep_dict = {key : value.upper() for key, value in sub_sleep_dict.items()}

sub_btns_dict = {"Lid close action" : "{5ca83367-6e45-459f-a27b-476b1d01c936}",
                 "Power button action" : "{7648efa3-dd9c-4e3e-b566-50f929386280}",
                 "Sleep button action" : "{96996bc0-ad50-47ec-923b-6f41874dd9eb}",
                 "Lid open action" : "{99ff10e7-23b1-4c07-a9d1-5c3206d741b4}",
                 "Start menu power button" : "{a7066653-8d6c-40a8-910e-a1f54b84c7e5}"
                 }
sub_btns_dict = {key : value.upper() for key, value in sub_btns_dict.items()}

sub_proc_dict = {"Processor performance increase threshold" : "{06cadf0e-64ed-448a-8927-ce7bf90eb35d}",
                 "Processor performance core parking min cores" : "{0cc5b647-c1df-4637-891a-dec35c318583}",
                 "Processor performance decrease threshold" : "{12a0ab44-fe28-4fa9-b3bd-4b64f44960a6}",
                 "Processor energy performance preference policy" : "{36687f9e-e3a5-4dbf-b1dc-15eb381c6863}",
                 "Allow Throttle States" : "{3b04d4fd-1cc7-4f23-ab1c-d1337819c4bb}",
                 "Processor performance decrease policy" : "{40fbefc7-2e9d-4d25-a185-0cfd8574bac6}",
                 "Processor performance boost policy" : "{45bcc044-d885-43e2-8605-ee0ec6e96b59}",
                 "Processor performance increase policy" : "{465e1f50-b610-473a-ab58-00d1077dc418}",
                 "Processor idle demote threshold" : "{4b92d758-5a24-4851-a470-815d78aee119}",
                 "Processor performance time check interval" : "{4d2b0152-7d5c-498b-88e2-34345392a2c5}",
                 "Maximum processor frequency" : "{75b0ae3f-bce0-45a7-8c89-c9611c25e100}",
                 "Minimum processor state" : "{893dee8e-2bef-41e0-89c6-b55d0929964c}",
                 "System cooling policy" : "{94d3a615-a899-4ac5-ae2b-e4d8f634367f}",
                 "Maximum processor state" : "{bc5038f7-23e0-4960-96da-33abaf5935ec}",
                 "Processor performance boost mode" : "{be337238-0d82-4146-a960-4f3749d470c7}"
                 }
sub_proc_dict = {key : value.upper() for key, value in sub_proc_dict.items()}

sub_disp_dict = {"Dim display after" : "{17aaa29b-8b43-4b94-aafe-35f64daaf1ee}",
                 "Turn off display after" : "{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}",
                 "Advanced Colour quality bias" : "{684c3e69-a4f7-4014-8754-d45179a56167}",
                 "Display brightness" : "{aded5e82-b909-4619-9949-f5d71dac0bcb}",
                 "Dimmed display brightness" : "{f1fbfde2-a960-4165-9f88-50667911ce96}",
                 "Enable adaptive brightness" : "{fbd9aa66-9553-4097-ba44-ed6e9d65eab8}"
                 }
sub_disp_dict = {key : value.upper for key, value in sub_disp_dict.items()}

sub_save_dict = {"Display brightness weight" : "{13d09884-f74e-474a-a852-b6bde8ad03a8}",
                 "Energy Saver Policy" : "{5c5bb349-ad29-4ee2-9d0b-2b25270f7a81}"
                 }
sub_save_dict = {key : value.upper for key, value in sub_save_dict.items()}

sub_btty_dict = {"Critical battery notification" : "{5dbb7c9f-38e9-40d2-9749-4f8a0e9f640f}",
                 "Critical battery action" : "{637ea02f-bbcb-4015-8e2c-a1c7b9c0b546}",
                 "Low battery level" : "{8183ba9a-e910-48da-8769-14ae6dc1170a}",
                 "Critical battery level" : "{9a66d8d7-4ff7-4ef9-b5a2-5a326ca2a469}",
                 "Low battery notification" : "{bcded951-187b-4d05-bccc-f7e51960c258}",
                 "Low battery action" : "{d8742dcb-3e6a-4b3c-b3fe-374623cdcf06}",
                 "Reserve battery level" : "{f3c5027d-cd16-4930-aa6b-90db844a8f00}"
                 }
sub_btty_dict = {key : value.upper for key, value in sub_btty_dict.items()}

(REG_NONE, REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD,
 REG_DWORD_LITTLE_ENDIAN, REG_DWORD_BIG_ENDIAN, REG_LINK,
 REG_MULTI_SZ, REG_RESOURCE_LIST, REG_FULL_RESOURCE_DESCRIPTOR,
 REG_RESOURCE_REQUIREMENTS_LIST, REG_QWORD) = range(13)

GUID_SUBGROUP_NONE = "{FEA3413E-7E05-4911-9A71-700331F1C294}"
GUID_SUBGROUP_DISK = "{0012EE47-9041-4B5D-9B77-535FBA8B1442}"
GUID_SUBGROUP_DESKTOP = "{0D7DBAE2-4294-402A-BA8E-26777E8488CD}"
GUID_SUBGROUP_WIRELESS = "{19CBB8FA-5279-450E-9FAC-8A3d5FEDD0C1}"
GUID_SUBGROUP_SLEEP = "{238C9FA8-0AAD-41ED-83F4-97BE242C8F20}"
GUID_SUBGROUP_BUTTONS = "{4F971E89-EEBD-4455-A8DE-9E59040E7347}"
GUID_SUBGROUP_PROCESSOR = "{54533251-82BE-4824-96C1-47B60B740D00}"
GUID_SUBGROUP_DISPLAY = "{7516B95F-F776-4464-8C53-06167F40CC99}"
GUID_SUBGROUP_ENERGY_SAVER = "{DE830923-A562-41AF-A086-E3A2C6BAD2DA}"
GUID_SUBGROUP_BATTERY = "{E73A048D-BF27-4F12-9731-8B2076E8891F}"

subgroups_dict = {
    "None": (GUID_SUBGROUP_NONE, sub_none_dict),
    "Disk": (GUID_SUBGROUP_DISK, sub_disk_dict),
    "Desktop": (GUID_SUBGROUP_DESKTOP, sub_desk_dict),
    "Wireless": (GUID_SUBGROUP_WIRELESS, sub_wire_dict),
    "Sleep": (GUID_SUBGROUP_SLEEP, sub_sleep_dict),
    "Buttons": (GUID_SUBGROUP_BUTTONS, sub_btns_dict),
    "Processor": (GUID_SUBGROUP_PROCESSOR, sub_proc_dict),
    "Display": (GUID_SUBGROUP_DISPLAY, sub_disp_dict),
    "Energy Saver": (GUID_SUBGROUP_ENERGY_SAVER, sub_save_dict),
    "Battery": (GUID_SUBGROUP_BATTERY, sub_btty_dict)
    }
