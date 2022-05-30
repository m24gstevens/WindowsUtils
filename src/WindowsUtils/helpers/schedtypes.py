from ctypes import Structure, POINTER, c_long, c_int, c_short, c_double
from ctypes.wintypes import WORD, DWORD
from comtypes import BSTR, COMMETHOD, HRESULT, GUID, IUnknown
from comtypes.automation import VARIANT, IDispatch

VARIANT_TRUE = -1
VARIANT_FALSE = 0
S_OK = 0

class SYSTEMTIME(Structure):
    _fields_ = [
        ("wYear", WORD),
        ("wMonth", WORD),
        ("wDayOfWeek", WORD),
        ("wDay", WORD),
        ("wHour", WORD),
        ("wMinute", WORD),
        ("wSecond", WORD),
        ("wMilliseconds", WORD)
        ]

class ITaskService(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{2FABA4C7-4DA9-4013-9697-20CC3FD40F85}")
    _idlflags_ = []

class ITaskDefinition(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{F5BC8FC5-536D-4F77-B852-FBC1356FDEB6}")
    _idlflags_ = []

class ITaskFolder(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{8CFAC062-A080-4C15-9A88-AA7C2AF80DFC}")
    _idlflags_ = []

class IRegistrationInfo(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{416D8B73-CB41-4ea1-805C-9BE9A5AC4A74}")
    _idlflags_ = []

class IPrincipal(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{D98D51E5-C9B4-496a-A9C1-18980261CF0F}")
    _idlflags_ = []

class ITaskSettings(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{8FD4711D-2D02-4c8c-87E3-EFF699DE127E}")
    _idlflags_ = []

class ITriggerCollection(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{85DF5081-1B24-4F32-878A-D9D14DF4CB77}")
    _idlflags_ = []

class ITrigger(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{09941815-EA89-4B5B-89E0-2A773801FAC3}")
    _idlflags_ = []

class IRepetitionPattern(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{7FB9ACF1-26BE-400e-85B5-294B9C75DFD6}")
    _idlflags_ = []

class IIdleSettings(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{84594461-0053-4342-A8FD-088FABF11F32}")
    _idlflags_ = []
    

class IActionCollection(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{02820E19-7B98-4ed2-B2E8-FDCCCEFF619B}")
    _idlflags_ = []

class IAction(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{BAE54997-48B1-4cbe-9965-D6BE263EBEA4}")
    _idlflags_ = []

class IExecAction(IAction):
    _case_insensitive_ = True
    _iid_ = GUID("{4C3D624D-FD6B-49A3-B9B7-09CB3CD3F047}")
    _idlflags_ = []

#Currently no support for Email Actions

class IRegisteredTask(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{9C86F320-DEE3-4DD1-B972-A303F26B061E}")
    _idlflags_ = []

class IRegisteredTaskCollection(IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID("{86627EB4-42A7-41E4-A4D9-AC33A72F2D52}")
    _idlflags_ = []


ITaskService._methods_ = [
    COMMETHOD([], HRESULT, 'GetFolder',
              (['in'], BSTR, 'path'),
              (['out'], POINTER(POINTER(ITaskFolder)), 'ppFolder')),
    COMMETHOD([], HRESULT, 'GetRunningTasks',
              (['out'], POINTER(POINTER(IDispatch)), 'ppRunningTasks')),
    COMMETHOD([], HRESULT, 'NewTask',
              (['in'], DWORD, 'flags'),
              (['out'], POINTER(POINTER(ITaskDefinition)), 'ppDefinition')),
    COMMETHOD([], HRESULT, 'Connect',
              (['in'], VARIANT, 'serverName'),
              (['in'], VARIANT, 'user'),
              (['in'], VARIANT, 'domain'),
              (['in'], VARIANT, 'password')),
    COMMETHOD([], HRESULT, 'get_Connected',
              (['out'], POINTER(c_short), 'pConnected')),
    COMMETHOD([], HRESULT, 'get_TargetServer',
              (['out'], POINTER(BSTR), 'pServer')),
    COMMETHOD([], HRESULT, 'get_ConnectedUser',
              (['out'], POINTER(BSTR), 'pUser')),
    COMMETHOD([], HRESULT, 'get_ConnectedDomain',
              (['out'], POINTER(BSTR), 'pDomain')),
    COMMETHOD([], HRESULT, 'get_HighestVersion',
              (['out'], POINTER(DWORD), 'pVersion'))
    ]

ITaskDefinition._methods_ = [
    COMMETHOD([], HRESULT, 'get_RegistrationInfo',
              (['out'], POINTER(POINTER(IRegistrationInfo)), 'ppRegistrationInfo')),
    COMMETHOD([], HRESULT, 'put_RegistrationInfo',
              (['in'], POINTER(IRegistrationInfo), 'pRegistrationInfo')),
    COMMETHOD([], HRESULT, 'get_Triggers',
              (['out'], POINTER(POINTER(ITriggerCollection)), 'ppTriggers')),
    COMMETHOD([], HRESULT, 'put_Triggers',
              (['in'], POINTER(ITriggerCollection), 'pTriggers')),
    COMMETHOD([], HRESULT, 'get_Settings',
              (['out'], POINTER(POINTER(ITaskSettings)), 'ppSettings')),
    COMMETHOD([], HRESULT, 'put_Settings',
              (['in'], POINTER(ITaskSettings), 'pSettings')),
    COMMETHOD([], HRESULT, 'get_Data',
              (['out'], POINTER(BSTR), 'pData')),
    COMMETHOD([], HRESULT, 'put_Data',
              (['in'], BSTR, 'data')),
    COMMETHOD([], HRESULT, 'get_Principal',
              (['out'], POINTER(POINTER(IPrincipal)), 'ppPrincipal')),
    COMMETHOD([], HRESULT, 'put_Principal',
              (['in'], POINTER(IPrincipal), 'pPrincipal')),
    COMMETHOD([], HRESULT, 'get_Actions',
              (['out'], POINTER(POINTER(IActionCollection)),'ppActions')),
    COMMETHOD([], HRESULT, 'put_Actions',
              (['in'], POINTER(IActionCollection), 'pActions')),
    COMMETHOD([], HRESULT, 'get_XmlText',
              (['out'], POINTER(BSTR), 'pXml')),
    COMMETHOD([], HRESULT, 'put_XmlText',
              (['in'], BSTR, 'xml'))
    ]

ITaskFolder._methods_ = [
    COMMETHOD([], HRESULT, 'get_Name',
              (['out'], POINTER(BSTR), 'pName')),
    COMMETHOD([], HRESULT, 'get_Path',
              (['out'], POINTER(BSTR), 'pPath')),
    COMMETHOD([], HRESULT, 'GetFolder',
              (['in'], BSTR, 'path'),
              (['out'], POINTER(POINTER(ITaskFolder)), 'ppFolder')),
    COMMETHOD([], HRESULT, 'GetFolders',
              (['in'], c_long, 'flags'),
              (['out'], POINTER(POINTER(IDispatch)), 'ppFolders')),
    COMMETHOD([], HRESULT, 'CreateFolder',
              (['in'], BSTR, 'subFolderName'),
              (['in'], VARIANT, 'sddl'),
              (['out'], POINTER(POINTER(ITaskFolder)), 'ppFolder')),
    COMMETHOD([], HRESULT, 'DeleteFolder',
              (['in'], BSTR, 'subFolderName'),
              (['in'], c_long, 'flags')),
    COMMETHOD([], HRESULT, 'GetTask',
              (['in'], BSTR, 'path'),
              (['out'], POINTER(POINTER(IRegisteredTask)), 'ppTask')),
    COMMETHOD([], HRESULT, 'GetTasks',
              (['in'], c_long, 'flags'),
              (['out'], POINTER(POINTER(IRegisteredTaskCollection)), 'ppTasks')),
    COMMETHOD([], HRESULT, 'DeleteTask',
              (['in'], BSTR, 'name'),
              (['in'], c_long, 'flags')),
    COMMETHOD([], HRESULT, 'RegisterTask',
              (['in'], BSTR, 'path'),
              (['in'], BSTR, 'xmlText'),
              (['in'], c_long, 'flags'),
              (['in'], VARIANT, 'userId'),
              (['in'], VARIANT, 'password'),
              (['in'], c_int, 'logonType'),
              (['in'], VARIANT, 'sddl'),
              (['out'], POINTER(POINTER(IRegisteredTask)), 'ppTask')),
    COMMETHOD([], HRESULT, 'RegisterTaskDefinition',
              (['in'], BSTR, 'path'),
              (['in'], POINTER(ITaskDefinition), 'pDefinition'),
              (['in'], c_long, 'flags'),
              (['in'], VARIANT, 'userId'),
              (['in'], VARIANT, 'password'),
              (['in'], c_int, 'logonType'),
              (['in', 'optional'], VARIANT, 'sddl'),
              (['out'], POINTER(POINTER(IRegisteredTask)), 'ppTask')),
    COMMETHOD([], HRESULT, 'GetSecurityDescriptor',
              (['in'], c_long, 'securityInformation'),
              (['out'], POINTER(BSTR), 'pSddl')),
    COMMETHOD([], HRESULT, 'SetSecurityDescriptor',
              (['in'], BSTR, 'sddl'),
              (['in'], c_long, 'flags'))
    ]

IRegistrationInfo._methods_ = [
    COMMETHOD([], HRESULT, 'get_Description',
              (['out'], POINTER(BSTR), 'pDescription')),
    COMMETHOD([], HRESULT, 'put_Description',
              (['in'], BSTR, 'description')),
    COMMETHOD([], HRESULT, 'get_Author',
              (['out'], POINTER(BSTR), 'pAuthor')),
    COMMETHOD([], HRESULT, 'put_Author',
              (['in'], BSTR, 'author')),
    COMMETHOD([], HRESULT, 'get_Version',
              (['out'], POINTER(BSTR), 'pVersion')),
    COMMETHOD([], HRESULT, 'put_Version',
              (['in'], BSTR, 'version')),
    COMMETHOD([], HRESULT, 'get_Date',
              (['out'], POINTER(BSTR), 'pDate')),
    COMMETHOD([], HRESULT, 'put_Date',
              (['in'], BSTR, 'date')),
    COMMETHOD([], HRESULT, 'get_Documentation',
              (['out'], POINTER(BSTR), 'pDocumentation')),
    COMMETHOD([], HRESULT, 'put_Documentation',
              (['in'], BSTR, 'documentation')),
    COMMETHOD([], HRESULT, 'get_XmlText',
              (['out'], POINTER(BSTR), 'pText')),
    COMMETHOD([], HRESULT, 'put_XmlText',
              (['in'], BSTR, 'text')),
    COMMETHOD([], HRESULT, 'get_URI',
              (['out'], POINTER(BSTR), 'pUri')),
    COMMETHOD([], HRESULT, 'put_URI',
              (['in'], BSTR, 'uri')),
    COMMETHOD([], HRESULT, 'get_SecurityDescriptor',
              (['out'], POINTER(VARIANT), 'pSddl')),
    COMMETHOD([], HRESULT, 'put_SecurityDescriptor',
              (['in'], VARIANT, 'sddl')),
    COMMETHOD([], HRESULT, 'get_Source',
              (['out'], POINTER(BSTR), 'pSource')),
    COMMETHOD([], HRESULT, 'put_Source',
              (['in'], BSTR, 'source'))
    ]

RUNLEVEL_LOWEST = 0
RUNLEVEL_HIGHEST = 1
(LOGON_NONE, LOGON_PASSWORD, LOGON_S4U, LOGON_INTERACTIVE_TOKEN,
 LOGON_GROUP, LOGON_SERVICE_ACCOUNT, LOGON_INTERACTIVE_TOKEN_OR_PASSWORD) = range(7)


IPrincipal._methods_ = [
    COMMETHOD([], HRESULT, 'get_Id',
              (['out'], POINTER(BSTR), 'pId')),
    COMMETHOD([], HRESULT, 'put_Id',
              (['in'], BSTR, 'id')),
    COMMETHOD([], HRESULT, 'get_DisplayName',
              (['out'], POINTER(BSTR), 'pName')),
    COMMETHOD([], HRESULT, 'put_DisplayName',
              (['in'], BSTR, 'name')),
    COMMETHOD([], HRESULT, 'get_UserId',
              (['out'], POINTER(BSTR), 'pUser')),
    COMMETHOD([], HRESULT, 'put_UserId',
              (['in'], BSTR, 'user')),
    COMMETHOD([], HRESULT, 'get_LogonType',
              (['out'], POINTER(c_int), 'pLogon')),
    COMMETHOD([], HRESULT, 'put_LogonType',
              (['in'], c_int, 'logon')),
    COMMETHOD([], HRESULT, 'get_GroupId',
              (['out'], POINTER(BSTR), 'pGroup')),
    COMMETHOD([], HRESULT, 'put_GroupId',
              (['in'], BSTR, 'group')),
    COMMETHOD([], HRESULT, 'get_RunLevel',
              (['out'], POINTER(c_int), 'pRunLevel')),
    COMMETHOD([], HRESULT, 'put_RunLevel',
              (['in'], c_int, 'runLevel'))
    ]

(INSTANCES_PARALLEL, INSTANCES_QUEUE, INSTANCES_IGNORE_NEW, INSTANCES_STOP_EXISTING) = range(4)

ITaskSettings._methods_ = [
    COMMETHOD([], HRESULT, 'get_AllowDemandStart',
              (['out'], POINTER(c_short), 'pAllowDemandStart')),
    COMMETHOD([], HRESULT, 'put_AllowDemandStart',
              (['in'], c_short, 'allowDemandStart')),
    COMMETHOD([], HRESULT, 'get_RestartInterval',
              (['out'], POINTER(BSTR), 'pRestartInterval')),
    COMMETHOD([], HRESULT, 'put_RestartInterval',
              (['in'],BSTR, 'restartInterval')),
    COMMETHOD([], HRESULT, 'get_RestartCount',
              (['out'], POINTER(c_int), 'pRestartCount')),
    COMMETHOD([], HRESULT, 'put_RestartCount',
              (['in'], c_int, 'restartCount')),
    COMMETHOD([], HRESULT, 'get_MultipleInstances',
              (['out'], POINTER(c_int), 'pPolicy')),
    COMMETHOD([], HRESULT, 'put_MultipleInstances',
              (['in'], c_int, 'policy')),
    COMMETHOD([], HRESULT, 'get_DisallowStartIfOnBatteries',
              (['out'], POINTER(c_short), 'pDisallowStart')),
    COMMETHOD([], HRESULT, 'put_DisallowStartIfOnBatteries',
              (['in'], c_short, 'disallowStart')),
    COMMETHOD([], HRESULT, 'get_AllowHardTerminate',
              (['out'], POINTER(c_short), 'pAllowHardTerminate')),
    COMMETHOD([], HRESULT, 'put_AllowHardTerminate',
              (['in'], c_short, 'allowHardTerminate')),
    COMMETHOD([], HRESULT, 'get_StartWhenAvailable',
              (['out'], POINTER(c_short), 'pStartWhenAvailable')),
    COMMETHOD([], HRESULT, 'put_StartWhenAvailable',
              (['in'], c_short, 'startWhenAvailable')),
    COMMETHOD([], HRESULT, 'get_XmlText',
              (['out'], POINTER(BSTR), 'pText')),
    COMMETHOD([], HRESULT, 'put_XmlText',
              (['in'], BSTR, 'text')),
    COMMETHOD([], HRESULT, 'get_RunOnlyIfNetworkAvailable',
              (['out'], POINTER(c_short), 'pRunOnlyIfNetworkAvailable')),
    COMMETHOD([], HRESULT, 'put_RunOnlyIfNetworkAvailable',
              (['in'], c_short, 'runOnlyIfNetworkAvailable')),
    COMMETHOD([], HRESULT, 'get_ExecutionTimeLimit',
              (['out'], POINTER(BSTR), 'pExecutionTimeLimit')),
    COMMETHOD([], HRESULT, 'put_ExecutionTimeLimit',
              (['in'], BSTR, 'executionTimeLimit')),
    COMMETHOD([], HRESULT, 'get_Enabled',
              (['out'], POINTER(c_short), 'pEnabled')),
    COMMETHOD([], HRESULT, 'put_Enabled',
              (['in'], c_short, 'enabled')),
    COMMETHOD([], HRESULT, 'get_DeleteExpiredTaskAfter',
              (['out'], POINTER(BSTR), 'pExpirationDelay')),
    COMMETHOD([], HRESULT, 'put_DeleteExpiredTaskAfter',
              (['in'], BSTR, 'expirationDelay')),
    COMMETHOD([], HRESULT, 'get_Priority',
              (['out'], POINTER(c_int), 'pPriority')),
    COMMETHOD([], HRESULT, 'put_Priority',
              (['in'], c_int, 'priority')),
    COMMETHOD([], HRESULT, 'get_Compatibility',
              (['out'], POINTER(c_int), 'pCompatLevel')),
    COMMETHOD([], HRESULT, 'put_Compatibility',
              (['in'], c_int, 'compatLevel')),
    COMMETHOD([], HRESULT, 'get_Hidden',
              (['out'], POINTER(c_short), 'pHidden')),
    COMMETHOD([], HRESULT, 'put_Hidden',
              (['in'], c_short, 'hidden')),
    COMMETHOD([], HRESULT, 'get_IdleSettings',
              (['out'], POINTER(POINTER(IIdleSettings)), 'ppIdleSettings')),
    COMMETHOD([], HRESULT, 'put_IdleSettings',
              (['in'], POINTER(IIdleSettings), 'pIdleSettings')),
    COMMETHOD([], HRESULT, 'get_RunOnlyIfIdle',
              (['out'], POINTER(c_short), 'pRunOnlyIfIdle')),
    COMMETHOD([], HRESULT, 'put_RunOnlyIfIdle',
              (['in'], c_short, 'runOnlyIfIdle')),
    COMMETHOD([], HRESULT, 'get_WakeToRun',
              (['out'], POINTER(c_short), 'pWake')),
    COMMETHOD([], HRESULT, 'put_WakeToRun',
              (['in'], c_short, 'wake')),
    COMMETHOD([], HRESULT, 'get_NetworkSettings',
              (['out'], POINTER(POINTER(IDispatch)), 'ppNetworkSettings')),
    COMMETHOD([], HRESULT, 'put_NetworkSettings',
              (['in'], POINTER(IDispatch), 'pNetworkSettings'))
    ]

(TRIGGER_EVENT, TRIGGER_TIME, TRIGGER_DAILY, TRIGGER_WEEKLY,
 TRIGGER_MONTHLY, TRIGGER_MONTHLYDOW, TRIGGER_IDLE, TRIGGER_REGISTRATION,
 TRIGGER_BOOT, TRIGGER_LOGON) = range(10)
TRIGGER_SESSION_STATE_CHANGE = 11
TRIGGER_CUSTOM_TRIGGER = 12
ITriggerCollection._methods_ = [
    COMMETHOD([], HRESULT, 'get_Count',
              (['out'], POINTER(c_long), 'pCount')),
    COMMETHOD([], HRESULT, 'get_Item',
              (['in'], c_long, 'index'),
              (['out'], POINTER(POINTER(ITrigger)), 'ppTrigger')),
    COMMETHOD([], HRESULT, 'get__NewEnum',
              (['out'], POINTER(POINTER(IUnknown)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'Create',
              (['in'], c_int, 'type'),
              (['out'], POINTER(POINTER(ITrigger)), 'ppTrigger')),
    COMMETHOD([], HRESULT, 'Remove',
              (['in'], VARIANT, 'index')),
    COMMETHOD([], HRESULT, 'Clear')
    ]

ITrigger._methods_ = [
    COMMETHOD([], HRESULT, 'get_Type',
              (['out'], POINTER(c_int), 'pType')),
    COMMETHOD([], HRESULT, 'get_Id',
              (['out'], POINTER(BSTR), 'pId')),
    COMMETHOD([], HRESULT, 'put_Id',
              (['in'], BSTR, 'id')),
    COMMETHOD([], HRESULT, 'get_Repetition',
              (['out'], POINTER(POINTER(IRepetitionPattern)), 'ppRepeat')),
    COMMETHOD([], HRESULT, 'put_Repetition',
              (['in'], POINTER(IRepetitionPattern), 'pRepeat')),
    COMMETHOD([], HRESULT, 'get_ExecutionTimeLimit',
              (['out'], POINTER(BSTR), 'pTimeLimit')),
    COMMETHOD([], HRESULT, 'put_ExecutionTimeLimit',
              (['in'], BSTR, 'timeLimit')),
    COMMETHOD([], HRESULT, 'get_StartBoundary',
              (['out'], POINTER(BSTR), 'pStart')),
    COMMETHOD([], HRESULT, 'put_StartBoundary',
              (['in'], BSTR, 'start')),
    COMMETHOD([], HRESULT, 'get_EndBoundary',
              (['out'], POINTER(BSTR), 'pEnd')),
    COMMETHOD([], HRESULT, 'put_EndBoundary',
              (['in'], BSTR, 'send')),
    COMMETHOD([], HRESULT, 'get_Enabled',
              (['out'], POINTER(c_short), 'pEnabled')),
    COMMETHOD([], HRESULT, 'put_Enabled',
              (['in'], c_short, 'enabled'))
    ]

IRepetitionPattern._methods_ = [
    COMMETHOD([], HRESULT, 'get_Interval',
              (['out'], POINTER(BSTR), 'pInterval')),
    COMMETHOD([], HRESULT, 'put_Interval',
              (['in'],BSTR, 'interval')),
    COMMETHOD([], HRESULT, 'get_Duration',
              (['out'], POINTER(BSTR), 'pDuration')),
    COMMETHOD([], HRESULT, 'put_Duration',
              (['in'],BSTR, 'duration')),
    COMMETHOD([], HRESULT, 'get_StopAtDurationEnd',
              (['out'], POINTER(c_short), 'pStop')),
    COMMETHOD([], HRESULT, 'put_StopAtDurationEnd',
              (['in'], c_short, 'stop'))
    ]

IIdleSettings._methods_ = [
    COMMETHOD([], HRESULT, 'get_IdleDuration',
              (['out'], POINTER(BSTR), 'pDelay')),
    COMMETHOD([], HRESULT, 'put_IdleDuration',
              (['in'], BSTR, 'delay')),
    COMMETHOD([], HRESULT, 'get_WaitTimeout',
              (['out'], POINTER(BSTR), 'pTimeout')),
    COMMETHOD([], HRESULT, 'put_WaitTimeout',
              (['in'], BSTR, 'timeout')),
    COMMETHOD([], HRESULT, 'get_StopOnIdleEnd',
              (['out'], POINTER(c_short), 'pStop')),
    COMMETHOD([], HRESULT, 'put_StopOnIdleEnd',
              (['in'], c_short, 'stop')),
    COMMETHOD([], HRESULT, 'get_RestartOnIdle',
              (['out'], POINTER(c_short), 'pRestart')),
    COMMETHOD([], HRESULT, 'put_RestartOnIdle',
              (['in'], c_short, 'restart'))
    ]
    
    

ACTION_EXEC = 0
(ACTION_COM_HANDLER, ACTION_SEND_EMAIL, ACTION_SHOW_MESSAGE) = range(5,8)
IActionCollection._methods_ = [
    COMMETHOD([], HRESULT, 'get_Count',
              (['out'], POINTER(c_long), 'pCount')),
    COMMETHOD([], HRESULT, 'get_Item',
              (['in'], c_long, 'index'),
              (['out'], POINTER(POINTER(IAction)), 'ppTrigger')),
    COMMETHOD([], HRESULT, 'get__NewEnum',
              (['out'], POINTER(POINTER(IUnknown)), 'ppEnum')),
    COMMETHOD([], HRESULT, 'get_XmlText',
              (['out'], POINTER(BSTR), 'pText')),
    COMMETHOD([], HRESULT, 'put_XmlText',
              (['in'], BSTR, 'text')),
    COMMETHOD([], HRESULT, 'Create',
              (['in'], c_int, 'type'),
              (['out'], POINTER(POINTER(IAction)), 'ppAction')),
    COMMETHOD([], HRESULT, 'Remove',
              (['in'], VARIANT, 'index')),
    COMMETHOD([], HRESULT, 'Clear'),
    COMMETHOD([], HRESULT, 'get_Context',
              (['out'], POINTER(BSTR), 'pContext')),
    COMMETHOD([], HRESULT, 'put_Context',
              (['in'], BSTR, 'context'))
    ]

IAction._methods_ = [
    COMMETHOD([], HRESULT, 'get_Id',
              (['out'], POINTER(BSTR), 'pId')),
    COMMETHOD([], HRESULT, 'put_Id',
              (['in'], BSTR, 'id')),
    COMMETHOD([], HRESULT, 'get_Type',
              (['out'], POINTER(c_int), 'pType'))
    ]

IExecAction._methods_ = [
    COMMETHOD([], HRESULT, 'get_Path',
              (['out'], POINTER(BSTR), 'pPath')),
    COMMETHOD([], HRESULT, 'put_Path',
              (['in'], BSTR, 'path')),
    COMMETHOD([], HRESULT, 'get_Arguments',
              (['out'], POINTER(BSTR), 'pArgument')),
    COMMETHOD([], HRESULT, 'put_Arguments',
              (['in'], BSTR, 'argument')),
    COMMETHOD([], HRESULT, 'get_WorkingDirectory',
              (['out'], POINTER(BSTR), 'pWorkingDirectory')),
    COMMETHOD([], HRESULT, 'put_WorkingDirectory',
              (['in'], BSTR, 'workingDirectory'))
    ]

IRegisteredTaskCollection._methods_ = [
    COMMETHOD([], HRESULT, 'get_Count',
              (['out'], POINTER(c_long), 'pCount')),
    COMMETHOD([], HRESULT, 'get_Item',
              (['in'], c_long, 'index'),
              (['out'], POINTER(POINTER(IRegisteredTask)), 'ppRegisteredTask')),
    COMMETHOD([], HRESULT, 'get__NewEnum',
              (['out'], POINTER(POINTER(IUnknown)), 'ppEnum'))
    ]

CREATE_OR_UPDATE = 5
(STATE_UNKNOWN, STATE_DISABLED, STATE_QUEUED, STATE_READY, STATE_RUNNING) = range(5)
IRegisteredTask._methods_ = [
    COMMETHOD([], HRESULT, 'get_Name',
              (['out'], POINTER(BSTR), 'pName')),
    COMMETHOD([], HRESULT, 'get_Path',
              (['out'], POINTER(BSTR), 'pPath')),
    COMMETHOD([], HRESULT, 'get_State',
              (['out'], POINTER(c_int), 'pState')),
    COMMETHOD([], HRESULT, 'get_Enabled',
              (['out'], POINTER(c_short), 'pEnabled')),
    COMMETHOD([], HRESULT, 'put_Enabled',
              (['in'], c_short, 'enabled')),
    COMMETHOD([], HRESULT, 'Run',
              (['in'], VARIANT, 'params'),
              (['out'], POINTER(POINTER(IDispatch)), 'ppRunningTask')),
    COMMETHOD([], HRESULT, 'RunEx',
              (['in'], VARIANT, 'params'),
              (['in'], c_long, 'flags'),
              (['in'], c_long, 'sessionId'),
              (['in'], BSTR, 'user'),
              (['out'], POINTER(POINTER(IDispatch)), 'ppRunningTask')),
    COMMETHOD([], HRESULT, 'GetInstances',
              (['in'], c_long, 'flags'),
              (['out'], POINTER(POINTER(IDispatch)), 'ppRunningTasks')),
    COMMETHOD([], HRESULT, 'get_LastRunTime',
              (['out'], POINTER(c_double), 'pLastRunTime')),
    COMMETHOD([], HRESULT, 'get_LastTaskResult',
              (['out'], POINTER(c_long), 'pLastTaskResult')),
    COMMETHOD([], HRESULT, 'get_NumberOfMissedRuns',
              (['out'], POINTER(c_long), 'pNumberOfMissedRuns')),
    COMMETHOD([], HRESULT, 'get_NextRunTime',
              (['out'], POINTER(c_double), 'pNextRunTime')),
    COMMETHOD([], HRESULT, 'get_Definition',
              (['out'], POINTER(POINTER(ITaskDefinition)), 'ppDefinition')),
    COMMETHOD([], HRESULT, 'get_Xml',
              (['out'], POINTER(BSTR), 'pXml')),
    COMMETHOD([], HRESULT, 'GetSecurityDescriptor',
              (['in'], c_long, 'securityInformation'),
              (['out'], POINTER(BSTR), 'pSddl')),
    COMMETHOD([], HRESULT, 'SetSecurityDescriptor',
              (['in'], BSTR, 'sddl'),
              (['in'], c_long, 'flags')),
    COMMETHOD([], HRESULT, 'Stop',
              (['in'], c_long, 'flags')),
    COMMETHOD([], HRESULT, 'GetRunTimes',
              (['in'], POINTER(SYSTEMTIME), 'pstStart'),
              (['in'], POINTER(SYSTEMTIME), 'pstEnd'),
              (['in','out'], POINTER(DWORD), 'pCount'),
              (['out'], POINTER(SYSTEMTIME), 'pRunTimes'))
    ]
#A double represent the DATE type. Must be able to parse into correct format

class IIdleTrigger(ITrigger):
    _case_insensitive_ = True
    _iid_ = GUID("{D537D2B0-9FB3-4D34-9739-1FF5CE7B1EF3}")
    _idlflags_ = []

class IDailyTrigger(ITrigger):
    _case_insensitive_ = True
    _iid_ = GUID("{126C5CD8-B288-41D5-8DBF-E491446ADC5C}")
    _idlflags_ = []

IDailyTrigger._methods_ = [
    COMMETHOD([], HRESULT, 'get_DaysInterval',
              (['out'], POINTER(c_short), 'pDays')),
    COMMETHOD([], HRESULT, 'put_DaysInterval',
              (['in'], c_short, 'days')),
    COMMETHOD([], HRESULT, 'get_RandomDelay',
              (['out'], POINTER(BSTR), 'pRandomDelay')),
    COMMETHOD([], HRESULT, 'put_RandomDelay',
              (['in'], BSTR, 'randomDelay'))
    ]

class IWeeklyTrigger(ITrigger):
    _case_insensitive_ = True
    _iid_ = GUID("{5038FC98-82FF-436D-8728-A512A57C9DC1}")
    _idlflags_ = []

SUNDAY = 0x1
MONDAY = 0x2
TUESDAY = 0x4
WEDNESDAY = 0x8
THURSDAY = 0x10
FRIDAY = 0x20
SATURDAY = 0x40

IWeeklyTrigger._methods_ = [
    COMMETHOD([], HRESULT, 'get_DaysOfWeek',
              (['out'], POINTER(c_short), 'pDays')),
    COMMETHOD([], HRESULT, 'put_DaysOfWeek',
              (['in'], c_short, 'days')),
    COMMETHOD([], HRESULT, 'get_WeeksInterval',
              (['out'], POINTER(c_short), 'pWeeks')),
    COMMETHOD([], HRESULT, 'put_WeeksInterval',
              (['in'], c_short, 'weeks')),
    COMMETHOD([], HRESULT, 'get_RandomDelay',
              (['out'], POINTER(BSTR), 'pRandomDelay')),
    COMMETHOD([], HRESULT, 'put_RandomDelay',
              (['in'], BSTR, 'randomDelay'))
    ]

class ILogonTrigger(ITrigger):
    _case_insensitive_ = True
    _iid_ = GUID("{72DADE38-FAE4-4b3e-BAF4-5D009AF02B1C}")
    _idlflags_ = []

ILogonTrigger._methods_ = [
    COMMETHOD([], HRESULT, 'get_Delay',
              (['out'], POINTER(BSTR), 'pDelay')),
    COMMETHOD([], HRESULT, 'put_Delay',
              (['in'], BSTR, 'delay')),
    COMMETHOD([], HRESULT, 'get_UserID',
              (['out'], POINTER(BSTR), 'pUser')),
    COMMETHOD([], HRESULT, 'put_UserID',
              (['in'], BSTR, 'user'))
    ]


class ITimeTrigger(ITrigger):
    _case_insensitive_ = True
    _iid_ = GUID("{B45747E0-EBA7-4276-9F29-85C5BB300006}")
    _idlflags_ = []

ITimeTrigger._methods_ = [
    COMMETHOD([], HRESULT, 'get_RandomDelay',
              (['out'], POINTER(BSTR), 'pRandomDelay')),
    COMMETHOD([], HRESULT, 'put_RandomDelay',
              (['in'], BSTR, 'randomDelay'))
    ]
