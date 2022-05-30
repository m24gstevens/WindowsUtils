from comtypes.client import CreateObject
from comtypes import *
from ctypes import *
from ctypes.wintypes import *
from comtypes.automation import VARIANT

BSTR = c_wchar_p

#Wbem classes

class IWbemLocator(IUnknown):
    _case_insensitive = True
    _iid_ = GUID("{DC12A687-737F-11CF-884D-00AA004B2E24}")
    _idlflags_ = []
    
class IWbemServices(IUnknown):
    _case_insensitive = True
    _iid_ = GUID("{9556DC99-828C-11CF-A37E-00AA003240C7}")
    _idlflags_ = []

class IWbemClassObject(IUnknown):
    _case_insensitive = True
    _iid_ = GUID("{DC12A681-737F-11CF-884D-00AA004B2E24}")
    _idlflags_ = []

class IEnumWbemClassObject(IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID("{027947E1-D731-11CE-A357-000000000001}")

IWbemLocator._methods_ = [
    COMMETHOD([], HRESULT, 'ConnectServer',
              (['in'], BSTR, 'strNetworkResource'),
              (['in'], BSTR, 'strUser'),
              (['in'], BSTR, 'strPassword'),
              (['in'], BSTR, 'strLocale'),
              (['in'], c_long, 'lSecurityFlags'),
              (['in'], BSTR, 'strAuthority'),
              (['in'], c_void_p, 'pCtx'),
              (['out'], POINTER(POINTER(IWbemServices)), 'ppNamespace'))
    ]

IWbemServices._methods_ = [
    COMMETHOD([], HRESULT, 'OpenNamespace',
              (['in'], BSTR, 'strNamespace'),
              (['in'], c_long, 'lFlags'),
              (['in'], POINTER(c_void_p),'ppWorkingNamespace'),
              (['in','out'], POINTER(c_void_p), 'ppResult')),
    COMMETHOD([], HRESULT, 'CancelAsyncCall',
              (['in'], c_void_p, 'pSink')),
    COMMETHOD([], HRESULT, 'QueryObjectSink',
              (['in'], c_long, 'lFlags'),
              (['out'], POINTER(c_void_p), 'ppResponsiveHandler')),
    COMMETHOD([], HRESULT, 'GetObject',
              (['in'], BSTR, 'strObjectPath'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in','out'], POINTER(POINTER(IWbemClassObject)), 'ppObject'),
              (['in','out'], POINTER(c_void_p), 'ppCallResult')),
    COMMETHOD([], HRESULT, 'GetObjectAsync',
              (['in'], BSTR, 'strObjectPath'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['out'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'PutClass',
              (['in'], POINTER(IWbemClassObject), 'pObject'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pContext'),
              (['in', 'out'], POINTER(c_void_p), 'ppCallResult')),
    COMMETHOD([], HRESULT, 'PutClassAsync',
              (['in'], POINTER(IWbemClassObject), 'pObject'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pContext'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'DeleteClass',
              (['in'], BSTR, 'strClass'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in', 'out'], POINTER(c_void_p), 'ppCallResult')),
    COMMETHOD([], HRESULT, 'DeleteClassAsync',
              (['in'], BSTR, 'strClass'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'CreateClassEnum',
              (['in'], POINTER(WCHAR), 'strSuperClass'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['out'], POINTER(POINTER(IEnumWbemClassObject)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'CreateClassEnumAsync',
              (['in'], BSTR, 'strSuperClass'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'PutInstance',
              (['in'], POINTER(IWbemClassObject), 'pInst'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in', 'out'], POINTER(c_void_p), 'ppCallResult')),
    COMMETHOD([], HRESULT, 'PutInstanceAsync',
              (['in'], POINTER(IWbemClassObject), 'pInst'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'DeleteInstance',
              (['in'], BSTR, 'strObjectPath'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in', 'out'], POINTER(c_void_p), 'ppCallResult')),
    COMMETHOD([], HRESULT, 'DeleteInstanceAsync',
              (['in'], BSTR, 'strObjectPath'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'CreateInstanceEnum',
              (['in'], BSTR, 'strFilter'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['out'], POINTER(POINTER(IEnumWbemClassObject)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'CreateInstanceEnumAsync',
              (['in'], BSTR, 'strFilter'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'ExecQuery',
              (['in'], BSTR, 'strQueryLanguage'),
              (['in'], BSTR, 'strQuery'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['out'], POINTER(POINTER(IEnumWbemClassObject)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'ExecQueryAsync',
              (['in'], BSTR, 'strQueryLanguage'),
              (['in'], BSTR, 'strQuery'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'ExecNotificationQuery',
              (['in'], BSTR, 'strQueryLanguage'),
              (['in'], BSTR, 'strQuery'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['out'], POINTER(POINTER(IEnumWbemClassObject)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'ExecNotificationQueryAsync',
              (['in'], BSTR, 'strQueryLanguage'),
              (['in'], BSTR, 'strQuery'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], c_void_p, 'pResponseHandler')),
    COMMETHOD([], HRESULT, 'ExecMethod',
              (['in'], BSTR, 'strObjectPath'),
              (['in'], BSTR, 'strMethodName'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], POINTER(IWbemClassObject), 'pInParams'),
              (['in', 'out'], POINTER(POINTER(IWbemClassObject)), 'ppOutParams'),
              (['in', 'out'], POINTER(c_void_p), 'ppCallResult')),
    COMMETHOD([], HRESULT, 'ExecMethodAsync',
              (['in'], BSTR, 'strObjectPath'),
              (['in'], BSTR, 'strMethodName'),
              (['in'], c_long, 'lFlags'),
              (['in'], c_void_p, 'pCtx'),
              (['in'], POINTER(IWbemClassObject), 'pInParams'),
              (['in'], c_void_p, 'pResponseHandler'))
    ]


IWbemClassObject._methods_ = [
    COMMETHOD([], HRESULT, 'GetQualifierSet',
              (['out'], POINTER(c_void_p), 'ppQualSet')),
    COMMETHOD([], HRESULT, 'Get',
              (['in'], LPCWSTR, 'wszName'),
              (['in'], c_long, 'lFlags'),
              (['in', 'out'], POINTER(VARIANT), 'pVal'),
              (['in', 'out'], POINTER(c_long), 'pType'),
              (['in', 'out'], POINTER(c_long), 'pFlavor')),
    COMMETHOD([], HRESULT, 'Put',
              (['in'], LPCWSTR, 'wszName'),
              (['in'], c_long, 'lFlags'),
              (['in'], POINTER(VARIANT), 'pVal'),
              (['in'], c_long, 'Type')),
    COMMETHOD([], HRESULT, 'Delete',
              (['in'], LPCWSTR, 'wszName')),
    COMMETHOD([], HRESULT, 'GetNames',
              (['in'], LPCWSTR, 'wszQualifierName'),
              (['in'], c_long, 'lFlags'),
              (['in'], POINTER(VARIANT), 'pQualifierVal'),
              (['out'], POINTER(c_void_p), 'pNames')),
    COMMETHOD([], HRESULT, 'BeginEnumeration',
              (['in'], c_long, 'lEnumFlags')),
    COMMETHOD([], HRESULT, 'Next',
              (['in'], c_long, 'lFlags'),
              (['in','out'], POINTER(BSTR), 'strName'),
              (['in','out'], POINTER(VARIANT), 'pVal'),
              (['in','out'], POINTER(c_long), 'pType'),
              (['in', 'out'], POINTER(c_long), 'plFlavor')),
    COMMETHOD([], HRESULT, 'EndEnumeration'),
    COMMETHOD([], HRESULT, 'GetPropertyQualifierSet',
              (['in'], LPCWSTR, 'wszProperty'),
              (['out'], POINTER(c_void_p), 'ppQualSet')),
    COMMETHOD([], HRESULT, 'Clone',
              (['out'], POINTER(POINTER(IWbemClassObject)), 'ppCopy')),
    COMMETHOD([], HRESULT, 'GetObjectText',
              (['in'], c_long, 'lFlags'),
              (['out'], POINTER(BSTR), 'pstrObjectText')),
    COMMETHOD([], HRESULT, 'SpawnDerivedClass',
              (['in'], c_long, 'lFlags'),
              (['out'], POINTER(POINTER(IWbemClassObject)), 'ppNewClass')),
    COMMETHOD([], HRESULT, 'SpawnInstance',
              (['in'], c_long, 'lFlags'),
              (['out'], POINTER(POINTER(IWbemClassObject)), 'ppNewInstance')),
    COMMETHOD([], HRESULT, 'CompareTo',
              (['in'], c_long, 'lFlags'),
              (['in'], POINTER(IWbemClassObject), 'pCompareTo')),
    COMMETHOD([], HRESULT, 'GetPropertyOrigin',
              (['in'], LPCWSTR, 'wszName'),
              (['out'], POINTER(BSTR), 'pstrClassName')),
    COMMETHOD([], HRESULT, 'InheritsFrom',
              (['in'], LPCWSTR, 'strAncestor')),
    COMMETHOD([], HRESULT, 'GetMethod',
              (['in'], LPCWSTR, 'wszName'),
              (['in'], c_long, 'lFlags'),
              (['out'], POINTER(POINTER(IWbemClassObject)), 'ppInSignature'),
              (['out'], POINTER(POINTER(IWbemClassObject)), 'ppOutSignature')),
    COMMETHOD([], HRESULT, 'PutMethod',
              (['in'], LPCWSTR, 'wszName'),
              (['in'], c_long, 'lFlags'),
              (['out'], POINTER(IWbemClassObject), 'pInSignature'),
              (['out'], POINTER(IWbemClassObject), 'pOutSignature')),
    COMMETHOD([], HRESULT, 'DeleteMethod',
              (['in'], LPCWSTR, 'wszName')),
    COMMETHOD([], HRESULT, 'BeginMethodEnumeration',
              (['in'], c_long, 'lEnumFlags')),
    COMMETHOD([], HRESULT, 'NextMethod',
              (['in'], c_long, 'lFlags'),
              (['in', 'out'], POINTER(BSTR), 'pstrName'),
              (['in','out'], POINTER(POINTER(IWbemClassObject)), 'ppInSignature'),
              (['in','out'], POINTER(POINTER(IWbemClassObject)), 'ppOutSignature')),
    COMMETHOD([], HRESULT, 'EndMethodEnumeration'),
    COMMETHOD([], HRESULT, 'GetMethodQualifierSet',
              (['in'], LPCWSTR, 'wszMethod'),
              (['out'], POINTER(c_void_p), 'ppQualSet')),
    COMMETHOD([], HRESULT, 'GetMethodOrigin',
              (['in'], LPCWSTR, 'wszMethodName'),
              (['out'], POINTER(BSTR), 'pstrClassName'))
    ]

IEnumWbemClassObject._methods_ = [
    COMMETHOD([], HRESULT, 'Reset'),
    COMMETHOD([], HRESULT, 'Next',
              (['in'], c_long, 'lTimeout'),
              (['in'], ULONG, 'uCount'),
              (['out'], POINTER(POINTER(IWbemClassObject)), 'apObjects'),
              (['out'], POINTER(ULONG), 'puReturned')),
    COMMETHOD([], HRESULT, 'NextAsync',
              (['in'], ULONG, 'uCount'),
              (['in'], c_void_p, 'pSink')),
    COMMETHOD([], HRESULT, 'Clone',
              (['out'], POINTER(POINTER(IEnumWbemClassObject)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'Skip',
              (['in'], c_long, 'lTimeout'),
              (['in'], ULONG, 'nCount'))
    ]

def _init_security():
    oledll.ole32.CoInitializeSecurity(None, #COM does service
                                      -1,
                                      None,
                                      None,
                                      0,    #Default Authentication
                                      3,    #Default impersionation
                                      None,
                                      0,    #Nothing additional
                                      None)

def create_server(namespace):
    """Connects to a WMI Server, outputting a reference to an
    IWbemServices interface"""
    CLSID = GUID("{4590F811-1D3A-11D0-891F-00AA004B2E24}")
    locator = CreateObject(CLSID, interface = IWbemLocator)
    services = locator.ConnectServer(namespace,
                                   None,
                                   None,
                                   None,
                                   0,
                                   None,
                                   None)
    locator.Release()
    return services

def get_property(server, query, property_name):
    """Gets the value of the given property from the first instance of
    an object returned by the query"""
    obj_enum = server.ExecQuery("WQL", query, 48, None)
    obj, num_returned = obj_enum.Next(-1, 1)
    if not obj:
        raise Exception("Query returned nothing")
    obj_enum.Release()
    value = obj.Get(property_name, 0)[0]
    obj.Release()
    return value

def exec_method(server, query, class_name, method_name, *params):
    """Gets a given object by a query, and executes a method on the first instance
    object returned
    Each parameter should be a tuple of form:
    (parameter string name, parameter string value, CIMTYPE enum data type)"""
    obj_enum = server.ExecQuery("WQL", query, 48, None)
    obj, num_returned = obj_enum.Next(-1,1)
    class_object, ignore = server.GetObject(class_name, 0, None)
    inparams, ignore = class_object.GetMethod(method_name, 0)
    instance = inparams.SpawnInstance(0)
    params_to_send = []
    for i in range(len(params)):
        strname, strvalue, cimtype = params[i]
        params_to_send.append(VARIANT())
        params_to_send[i]._set_value(strvalue)
        instance.Put(strname, 0, byref(params_to_send[i]), cimtype)
    path_var = obj.Get("__PATH",0)[0]
    server.ExecMethod(path_var, method_name, 0, None, instance)
    obj_enum.Release()
    obj.Release()
    class_object.Release()
    inparams.Release()
    instance.Release()
