"""
firewall.py
====================================
Functions for configuring firewall settings and creating firewall rules

Note: For security reasons, functions will not work unless the python processes
are run with admin privelidges
"""

from ctypes import POINTER, c_long, c_short, c_int, oledll
from ctypes.wintypes import BYTE
from comtypes.client import CreateObject
from comtypes import BSTR, IUnknown, COMMETHOD, HRESULT, GUID
from comtypes.automation import VARIANT, IDispatch

VARIANT_TRUE = -1
VARIANT_FALSE = 0

class INetFwAuthorizedApplication(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{B5E64FFA-C2C5-444E-A301-FB5E00018050}")
    _idlflags_ = []

class INetFwAuthorizedApplications(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{644EFD52-CCF9-486C-97A2-39F352570B30}")
    _idlflags_ = []

class INetFwRules(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{9C4C6277-5027-441E-AFAE-CA1F542DA009}")
    _idlflags_ = []

class INetFwRule(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{AF230D27-BABA-4E42-ACED-F524F22CFCE2}")
    _idlflags_ = []
    _class_id_ = GUID("{2C5BC43E-3369-4C33-AB0C-BE9469677AF4}")
    
    

(IP_VERSION_V4, IP_VERSION_V6, IP_VERSION_ANY, IP_VERSION_MAX) = range(4)
(SCOPE_ALL, SCOPE_LOCAL_SUBNET, SCOPE_CUSTOM, SCOPE_MAX) = range(4)

INetFwAuthorizedApplications._methods_ = [
    COMMETHOD([], HRESULT, 'get_Count',
              (['out'], POINTER(c_long), 'count')),
    COMMETHOD([], HRESULT, 'Add',
              (['in'], POINTER(INetFwAuthorizedApplication), 'app')),
    COMMETHOD([], HRESULT, 'Remove',
              (['in'], BSTR, 'imageFileName')),
    COMMETHOD([], HRESULT, 'Item',
              (['in'], BSTR, 'imageFileName'),
              (['out'], POINTER(POINTER(INetFwAuthorizedApplication)), 'app')),
    COMMETHOD([], HRESULT, 'get__NewEnum',
             (['out'], POINTER(POINTER(IUnknown)), 'newEnum'))
    ]

INetFwAuthorizedApplication._methods_ = [
    COMMETHOD([], HRESULT, 'get_Name',
              (['out'], POINTER(BSTR), 'name')),
    COMMETHOD([], HRESULT, 'put_Name',
              (['in'], BSTR, 'name')),
    COMMETHOD([], HRESULT, 'get_ProcessImageFileName',
              (['out'], POINTER(BSTR), 'imageFileName')),
    COMMETHOD([], HRESULT, 'put_ProcessImageFileName',
              (['in'], BSTR, 'imageFileName')),
    COMMETHOD([], HRESULT, 'get_IpVersion',
              (['out'], POINTER(c_int), 'ipVersion')),
    COMMETHOD([], HRESULT, 'put_IpVersion',
              (['in'], c_int, 'ipVersion')),
    COMMETHOD([], HRESULT, 'get_Scope',
              (['out'], POINTER(c_int), 'scope')),
    COMMETHOD([], HRESULT, 'put_Scope',
              (['in'], c_int, 'scope')),
    COMMETHOD([], HRESULT, 'get_RemoteAddresses',
              (['out'], POINTER(BSTR), 'remoteAddrs')),
    COMMETHOD([], HRESULT, 'put_RemoteAddresses',
              (['in'], BSTR, 'remoteAddrs')),
    COMMETHOD([], HRESULT, 'get_Enabled',
              (['out'], POINTER(c_short), 'enabled')),
    COMMETHOD([], HRESULT, 'put_Enabled',
              (['in'], c_short, 'enabled'))
    ]

(RULE_DIR_IN, RULE_DIR_OUT, RULE_DIR_MAX) = range(1,4)
INetFwRules._methods_ = [
    COMMETHOD([], HRESULT, 'get_Count',
              (['out'], POINTER(c_long), 'count')),
    COMMETHOD([], HRESULT, 'Add',
              (['in'], POINTER(INetFwRule), 'rule')),
    COMMETHOD([], HRESULT, 'Remove',
              (['in'], BSTR, 'name')),
    COMMETHOD([], HRESULT, 'Item',
              (['in'], BSTR, 'name'),
              (['out'], POINTER(POINTER(INetFwRule)), 'rule')),
    COMMETHOD([], HRESULT, 'get__NewEnum',
              (['out'], POINTER(POINTER(IUnknown)), 'newEnum'))
    ]


INetFwRule._methods_ = [
    COMMETHOD([], HRESULT, 'get_Name',
              (['out'], POINTER(BSTR), 'name')),
    COMMETHOD([], HRESULT, 'put_Name',
              (['in'], BSTR, 'name')),
    COMMETHOD([], HRESULT, 'get_Description',
              (['out'], POINTER(BSTR), 'desc')),
    COMMETHOD([], HRESULT, 'put_Description',
              (['in'], BSTR, 'desc')),
    COMMETHOD([], HRESULT, 'get_ApplicationName',
              (['out'], POINTER(BSTR), 'imageFileName')),
    COMMETHOD([], HRESULT, 'put_ApplicationName',
              (['in'], BSTR, 'imageFileName')),
    COMMETHOD([], HRESULT, 'get_ServiceName',
              (['out'], POINTER(BSTR), 'serviceName')),
    COMMETHOD([], HRESULT, 'put_ServiceName',
              (['in'], BSTR, 'serviceName')),
    COMMETHOD([], HRESULT, 'get_Protocol',
              (['out'], POINTER(c_long), 'protocol')),
    COMMETHOD([], HRESULT, 'put_Protocol',
              (['in'], c_long, 'protocol')),
    COMMETHOD([], HRESULT, 'get_LocalPorts',
              (['out'], POINTER(BSTR), 'portNumbers')),
    COMMETHOD([], HRESULT, 'put_LocalPorts',
              (['in'], BSTR, 'portNumbers')),
    COMMETHOD([], HRESULT, 'get_RemotePorts',
              (['out'], POINTER(BSTR), 'portNumbers')),
    COMMETHOD([], HRESULT, 'put_RemotePorts',
              (['in'], BSTR, 'portNumbers')),
    COMMETHOD([], HRESULT, 'get_LocalAddresses',
              (['out'], POINTER(BSTR), 'localAddrs')),
    COMMETHOD([], HRESULT, 'put_LocalAddresses',
              (['in'], BSTR, 'localAddrs')),
    COMMETHOD([], HRESULT, 'get_RemoteAddresses',
              (['out'], POINTER(BSTR), 'remoteAddrs')),
    COMMETHOD([], HRESULT, 'put_RemoteAddresses',
              (['in'], BSTR, 'remoteAddrs')),
    COMMETHOD([], HRESULT, 'get_IcmpTypesAndCodes',
              (['out'], POINTER(BSTR), 'icmpTypesAndCodes')),
    COMMETHOD([], HRESULT, 'put_IcmpTypesAndCodes',
              (['in'], BSTR, 'icmpTypesAndCodes')),
    COMMETHOD([], HRESULT, 'get_Direction',
              (['out'], POINTER(c_int), 'dir')),
    COMMETHOD([], HRESULT, 'put_Direction',
              (['in'], c_int, 'dir')),
    COMMETHOD([], HRESULT, 'get_Interfaces',
              (['out'], POINTER(VARIANT), 'interfaces')),
    COMMETHOD([], HRESULT, 'put_Interfaces',
              (['in'], VARIANT, 'interfaces')),
    COMMETHOD([], HRESULT, 'get_InterfaceTypes',
              (['out'], POINTER(BSTR), 'interfaceTypes')),
    COMMETHOD([], HRESULT, 'put_InterfaceTypes',
              (['in'], BSTR, 'interfaceTypes')),
    COMMETHOD([], HRESULT, 'get_Enabled',
              (['out'], POINTER(c_short), 'enabled')),
    COMMETHOD([], HRESULT, 'put_Enabled',
              (['in'], c_short, 'enabled')),
    COMMETHOD([], HRESULT, 'get_Grouping',
              (['out'], POINTER(BSTR), 'context')),
    COMMETHOD([], HRESULT, 'put_Grouping',
              (['in'], BSTR, 'context')),
    COMMETHOD([], HRESULT, 'get_Profiles',
              (['out'], POINTER(c_long), 'profileTypesBitmask')),
    COMMETHOD([], HRESULT, 'put_Profiles',
              (['in'], c_long, 'profileTypesBitmask')),
    COMMETHOD([], HRESULT, 'get_EdgeTraversal',
              (['out'], POINTER(c_short), 'enabled')),
    COMMETHOD([], HRESULT, 'put_EdgeTraversal',
              (['in'], c_short, 'enabled')),
    COMMETHOD([], HRESULT, 'get_Action',
              (['out'], POINTER(c_int), 'action')),
    COMMETHOD([], HRESULT, 'put_Action',
              (['in'], c_int, 'action'))
    ]
    

class INetFwMgr(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{F7898AF5-CAC4-4632-A2EC-DA06E5111AF2}")
    _idlflags_ = []
    _class_id_ = GUID("{304CE942-6E39-40D8-943A-B913C40C9CD4}")

class INetFwPolicy(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{D46D2478-9AC9-4008-9DC7-5563CE5536CC}")
    _idlflags_ = []

class INetFwPolicy2(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{98325047-C671-4174-8D81-DEFCD3F03186}")
    _idlflags_ = []
    _class_id_ = GUID("{E2B3C97F-6AE1-41AC-817A-F6F92166D7DD}")

class INetFwProfile(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{174A0DDA-E9F9-449D-993B-21AB667CA456}")
    _idlflags_ = []
    
    

(PROFILE_DOMAIN, PROFILE_STANDARD, PROFILE_CURRENT, PROFILE_MAX) = range(4)
IP_PROTOCOL_TCP = 6
IP_PROTOCOL_UDP = 17
IP_PROTOCOL_ANY = 256

FW_PROFILE_DOMAIN = 0x1
FW_PROFILE_PRIVATE = 0x2
FW_PROFILE_PUBLIC = 0x4
NET_FW_PROFILE_ALL = 0x7FFFFFFF

INBOUND = 1
OUTBOUND = 2

(ACTION_BLOCK, ACTION_ALLOW, ACTION_MAX) = range(3)

INetFwMgr._methods_ = [
    COMMETHOD([], HRESULT, 'get_LocalPolicy',
              (['out'], POINTER(POINTER(INetFwPolicy)), 'localPolicy')),
    COMMETHOD([], HRESULT, 'get_CurrentProfileType',
              (['out'], POINTER(c_int), 'profileType')),
    COMMETHOD([], HRESULT, 'RestoreDefaults'),
    COMMETHOD([], HRESULT, 'IsPortAllowed',
              (['in'], BSTR, 'imageFileName'),
              (['in'], c_int, 'ipVersion'),
              (['in'], c_long, 'portNumber'),
              (['in'], BSTR, 'localAddress'),
              (['in'], c_int, 'ipProtocol'),
              (['out'], POINTER(VARIANT), 'allowed'),
              (['out'], POINTER(VARIANT), 'restricted')),
    COMMETHOD([], HRESULT, 'IsIcmpTypeAllowed',
              (['in'], c_int, 'ipVersion'),
              (['in'], BSTR, 'localAddress'),
              (['in'], BYTE, 'type'),
              (['out'], POINTER(VARIANT), 'allowed'),
              (['out'], POINTER(VARIANT), 'restricted'))
    ]

INetFwPolicy._methods_ = [
    COMMETHOD([], HRESULT, 'get_CurrentProfile',
              (['out'], POINTER(POINTER(INetFwProfile)), 'profile')),
    COMMETHOD([], HRESULT, 'getProfileByType',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(POINTER(INetFwProfile)), 'profile'))
    ]

INetFwProfile._methods_ = [
    COMMETHOD([], HRESULT, 'get_Type',
              (['out'], POINTER(c_int), 'type')),
    COMMETHOD([], HRESULT, 'get_FilewallEnabled',
              (['out'], POINTER(c_short), 'enabled')),
    COMMETHOD([], HRESULT, 'put_FirewallEnabled',
              (['in'], c_short, 'enabled')),
    COMMETHOD([], HRESULT, 'get_ExceptionsNotAllowed',
              (['out'], POINTER(c_short), 'notAllowed')),
    COMMETHOD([], HRESULT, 'put_ExceptionsNotAllowed',
              (['in'], c_short, 'notAllowed')),
    COMMETHOD([], HRESULT, 'get_NotificationsDisabled',
              (['out'], POINTER(c_short), 'disabled')),
    COMMETHOD([], HRESULT, 'put_NotificationsDisabled',
              (['in'], c_short, 'disabled')),
    COMMETHOD([], HRESULT, 'get_UnicastResponsesToMulticastBroadcastDisabled',
              (['out'], POINTER(c_short), 'disabled')),
    COMMETHOD([], HRESULT, 'put_UnicastResponsesToMulticastBroadcastDisabled',
              (['in'], c_short, 'disabled')),
    COMMETHOD([], HRESULT, 'get_RemoteAdminSettings',
              (['out'], POINTER(POINTER(IDispatch)), 'remoteAdminSettings')),
    COMMETHOD([], HRESULT, 'get_ImcpSettings',
              (['out'], POINTER(POINTER(IDispatch)), 'imcpSettings')),
    COMMETHOD([], HRESULT, 'get_GloballyOpenPorts',
              (['out'], POINTER(POINTER(IDispatch)), 'openPorts')),
    COMMETHOD([], HRESULT, 'get_Services',
              (['out'], POINTER(POINTER(IDispatch)), 'services')),
    COMMETHOD([], HRESULT, 'get_AuthorizedApplications',
              (['out'], POINTER(POINTER(INetFwAuthorizedApplications)), 'apps'))
    ]

INetFwPolicy2._methods_ = [
    COMMETHOD([], HRESULT, 'get_CurrentProfileTypes',
              (['out'], POINTER(c_long), 'profileTypesBitmask')),
    COMMETHOD([], HRESULT, 'get_FirewallEnabled',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(c_short), 'enabled')),
    COMMETHOD([], HRESULT, 'put_FirewallEnabled',
              (['in'], c_int, 'profileType'),
              (['in'], c_short, 'enabled')),
    COMMETHOD([], HRESULT, 'get_ExcludedInterfaces',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(VARIANT), 'interfaces')),
    COMMETHOD([], HRESULT, 'put_ExcludedInterfaces',
              (['in'], c_int, 'profileType'),
              (['in'], VARIANT, 'interfaces')),
    COMMETHOD([], HRESULT, 'get_BlockAllInboundTraffic',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(c_short), 'Block')),
    COMMETHOD([], HRESULT, 'put_BlockAllInboundTraffic',
              (['in'], c_int, 'profileType'),
              (['in'], c_short, 'Block')),
    COMMETHOD([], HRESULT, 'get_NotificationsDisabled',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(c_short), 'disabled')),
    COMMETHOD([], HRESULT, 'put_NotificationsDisabled',
              (['in'], c_int, 'profileType'),
              (['in'], c_short, 'disabled')),
    COMMETHOD([], HRESULT, 'get_UnicastResponsesToMulticastBroadcastDisabled',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(c_short), 'disabled')),
    COMMETHOD([], HRESULT, 'put_UnicastResponsesToMulticastBroadcastDisabled',
              (['in'], c_int, 'profileType'),
              (['in'], c_short, 'disabled')),
    COMMETHOD([], HRESULT, 'get_Rules',
              (['out'], POINTER(POINTER(INetFwRules)), 'rules')),
    COMMETHOD([], HRESULT, 'get_ServiceRestriction',
              (['out'], POINTER(POINTER(IDispatch)), 'ServiceRestriction')),
    COMMETHOD([], HRESULT, 'EnableRuleGroup',
              (['in'], c_long, 'profileTypesBitmask'),
              (['in'], BSTR, 'group'),
              (['in'], c_short, 'enable')),
    COMMETHOD([], HRESULT, 'IsRuleGroupEnabled',
              (['in'], c_long, 'profileTypesBitmask'),
              (['in'], BSTR, 'group'),
              (['out'], POINTER(c_short), 'enable')),
    COMMETHOD([], HRESULT, 'RestoreLocalFirewallDefaults'),
    COMMETHOD([], HRESULT, 'get_DefaultInboundAction',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(c_int), 'action')),
    COMMETHOD([], HRESULT, 'put_DefaultInboundAction',
              (['in'], c_int, 'profileType'),
              (['in'], c_int, 'action')),
    COMMETHOD([], HRESULT, 'get_DefaultOutboundAction',
              (['in'], c_int, 'profileType'),
              (['out'], POINTER(c_int), 'action')),
    COMMETHOD([], HRESULT, 'put_DefaultOutboundAction',
              (['in'], c_int, 'profileType'),
              (['in'], c_int, 'action')),
    COMMETHOD([], HRESULT, 'get_IsRuleGroupCurrentlyEnabled',
              (['in'], BSTR, 'group'),
              (['out'], POINTER(c_short), 'enabled')),
    COMMETHOD([], HRESULT, 'get_LocalPolicyModifyState',
              (['out'], POINTER(c_int), 'modifyState'))
    ]

def _init_settings():
    """oledll.ole32.CoInitializeSecurity(None, #COM does service
                                      -1,
                                      None,
                                      None,
                                      0,    #Default Authentication
                                      3,    #Default impersionation
                                      None,
                                      0,    #Nothing additional
                                      None)"""
    try:
        policy_itf = CreateObject(INetFwPolicy2._class_id_, interface = INetFwPolicy2)
    except Exception:
        raise Exception("Couldn't change settings")
    return policy_itf      
        
def get_firewall_enabled(profile_value):
    """Gets whether the firewall is enabled on a particular profile

    Parameters
    ----------
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles

    Returns
    -------
    bool
        returns True if the firewall is enabled
    """
    policy_itf = _init_settings()
    value = bool(policy_itf.get_FirewallEnabled(profile_value))
    return value

def put_firewall_enabled(is_enable, profile_value):
    """Sets whether the firewall is enabled on a particular profile

    Parameters
    ----------
    is_enable : bool
        whether the specified profile is to be enabled
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles
    """
    policy_itf = _init_settings()
    value = VARIANT_TRUE if is_enable else VARIANT_FALSE
    policy_itf.put_FirewallEnabled(profile_value, value)

def get_block_inbound(profile_value):
    """Gets whether the firewall blocks all inbound traffic on a particular
    profile

    Parameters
    ----------
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles

    Returns
    -------
    bool
        returns True if the firewall blocks all inbound traffic
    """
    policy_itf = _init_settings()
    value = bool(policy_itf.get_BlockAllInboundTraffic(profile_value))
    return value

def put_block_inbound(is_block, profile_value):
    """Sets whether the firewall blocks inbound traffic on a particular profile

    Parameters
    ----------
    is_block : bool
        whether the specified profile is to block inbound traffic by default
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles
    """
    policy_itf = _init_settings()
    value = VARIANT_TRUE if is_block else VARIANT_FALSE
    policy_itf.put_BlockAllInboundTraffic(profile_value, value)

def get_notifications_disabled(profile_value):
    """Gets whether the firewall is has disabled notifications on a particular
    profile

    Parameters
    ----------
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles

    Returns
    -------
    bool
        returns True if the firewall is has notifications disabled
    """
    policy_itf = _init_settings()
    value = bool(policy_itf.get_NotificationsDisabled(profile_value))
    return value

def put_notifications_disabled(is_disable, profile_value):
    """Sets whether the firewall is has disabled notifications on a particular
    profile

    Parameters
    ----------
    is_disable : bool
        whether the specified profile is to have notifications disabled
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles
    """
    policy_itf = _init_settings()
    value = VARIANT_TRUE if is_disable else VARIANT_FALSE
    policy_itf.put_NotificationsDisabled(profile_value, value)

def get_unicast_response_to_multicast(profile_value):
    """Gets whether the firewall is has disabled unicast responses to
    broadcasted messages on a speified profile

    Parameters
    ----------
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles

    Returns
    -------
    bool
        returns True if the firewall allows such unicast responses
    """
    policy_itf = _init_settings()
    value =  not bool(policy_itf.get_UnicastResponsesToMulticastBroadcastDisabled(profile_value))
    return value

def put_unicast_response_to_multicast(is_enable, profile_value):
    """Sets whether the firewall is allow unicast responses to broadcast
    messages on a particular profile

    Parameters
    ----------
    is_enable : bool
        whether the specified profile is to allow such responses
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles
    """
    policy_itf = _init_settings()
    value = VARIANT_FALSE if is_enable else VARIANT_TRUE
    policy_itf.put_UnicastResponsesToMulticastBroadcastDisabled(profile_value, value)

def get_default_inbound_block(profile_value):
    """Gets whether the firewall blocks inbound traffic by default on a
    particular profile

    Parameters
    ----------
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles

    Returns
    -------
    bool
        returns True if the firewall blocks by default
    """
    policy_itf = _init_settings()
    value =  not bool(policy_itf.get_DefaultInboundAction(profile_value))
    return value

def get_default_outbound_block(profile_value):
    """Gets whether the firewall blocks outbound traffic by default on a
    particular profile

    Parameters
    ----------
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles

    Returns
    -------
    bool
        returns True if the firewall blocks by default
    """
    policy_itf = _init_settings()
    value =  not bool(policy_itf.get_DefaultOutboundAction(profile_value))
    return value

def put_default_inbound_block(is_block, profile_value):
    """Sets whether the firewall blocks inbound by default on a particular
    profile

    Parameters
    ----------
    is_block : bool
        whether the specified profile should, by default, block inbound traffic
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles
    """
    policy_itf = _init_settings()
    value = ACTION_BLOCK if is_block else ACTION_ALLOW
    policy_itf.put_DefaultInboundAction(profile_value, value)

def put_default_outbound_block(is_block, profile_value):
    """Sets whether the firewall blocks outbound by default on a particular
    profile

    Parameters
    ----------
    is_block : bool
        whether the specified profile should, by default, block outbound traffic
    profile_value : int
        one of FW_PROFILE_DOMAIN, FW_PROFILE_PUBLIC, FW_PROFILE_PRIVATE to
        configure one of the domain, private or public profiles
    """
    policy_itf = _init_settings()
    value = ACTION_BLOCK if is_block else ACTION_ALLOW
    policy_itf.put_DefaultOutboundAction(profile_value, value)


def add_rule(name, description, protocol, is_allow, enable, 
             profiles = FW_PROFILE_PUBLIC, direction = None,
             grouping = "WindowsUtils",
             application = None, service = None, interface_types = None,
             local_port = None, remote_port = None):
    """Adds a rule into the firewall.

    Parameters
    ----------
    name : int
        name of the firewall rule
    description : str
        description of the firewall rule
    protocol : int
        one of IP_PROTOCOL_TCP, IP_PROTOCOL_UDP, or some other assigned IP
        protocol number
    is_allow : bool
        whether we will be allowing or disallowing traffic
    enable : bool
        whether the rule should be enabled after submitting
    profiles : int, optional
        bitmask of FW_PROFILE_PUBLIC/PRIVATE/DOMAIN values detailing profiles
        the rule will apply to
        Defaults to public only
    direction : int, optional
        one of INBOUND, OUTBOUND to specify rule direction.
        Defaults to both
    grouping : str, optional
        name of the group the rule will belong to.
        defaults to a WindowsUtils group
    application : str, optional
        system path, possibly including environment variables to be substituted,
        to a process executable to which the rule will apply
    service : str, optional
        name of service provided by application
    interface_types : str, optional
        one of "RemoteAccess", "Wireless", "LAN", "All" to specify an interface
        the rule will apply on.
        If using multiple, separate by a comma.
    local_port : int, optional
        port number on the local machine the rule will apply to
    remote_port : int, optional
         server's port number the rule will apply to
    """
    policy_itf = _init_settings()
    try:
        rules = policy_itf.get_Rules()
        rule = CreateObject(INetFwRule._class_id_, interface=INetFwRule)
    except Exception:
        raise Exception("Couldn't create a new rule")
    else:
        try:
            rule.put_Name(name)
            rule.put_Description(description)
            if not (application is None):
                rule.put_ApplicationName(application)
            if not (service is None):
                rule.put_ServiceName(service)
            rule.put_Protocol(protocol)
            if not (local_port is None):
                if 0 < local_port <= 65535:
                    rule.put_LocalPorts(str(local_port))
            if not (remote_port is None):
                if 0 < remote_port <= 65535:
                    rule.put_RemotePorts(str(remote_port))
            if not (direction is None):
                if direction in (INBOUND, OUTBOUND):
                    rule.put_Direction(direction)
            rule.put_Grouping(grouping)
            if interface_types:
                for s in interface_types.split('c'):
                    if not (s in ["LAN", "RemoteAccess", "Wireless", "All"]):
                        raise Exception("Invalid interface type")
                rule.put_InterfaceTypes(interface_types)
            rule.put_Profiles(profiles)
            rule.put_Action(ACTION_ALLOW if is_allow else ACTION_BLOCK)
            rule.put_Enabled(VARIANT_TRUE if enable else VARIANT_FALSE)
            rules.add(rule)
        except Exception:
            raise Exception("Invalid Rule Parameters")
            
        



    


