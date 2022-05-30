"""
scanner.py
====================================
Function for scanning a file for malware
"""

from ctypes import (windll, POINTER, cast, c_void_p, Structure, memset,
                    byref, sizeof, c_int, c_size_t, c_char_p)
from ctypes.wintypes import BYTE, DWORD, LPCWSTR, HANDLE

kernellib = windll.Kernel32
amsilib = windll.Amsi
S_OK = 0
PROJECT_NAME = "WindowsUtils Scanner"
INVALID_HANDLE_VALUE = -1
INVALID_FILE_SIZE = 4294967295

RESULT_CLEAN = 0
RESULT_NOT_DETECTED = 1
RESULT_BLOCKED_BY_ADMIN_START = 0x4000
RESULT_BLOCKED_BY_ADMIN_END = 0x4FFF
RESULT_DETECTED = 32768


(AMSI_ATTR_APP_NAME, AMSI_ATTR_CONTENT_NAME, AMSI_ATTR_CONTENT_SIZE,
 AMSI_ATTR_CONTENT_ADDRESS, AMSI_ATTR_SESSION, AMSI_ATTR_REDIRECT_CHAIN_SIZE,
 AMSI_ATTR_REDIRECT_CHAIN_ADDRESS, AMSI_ATTR_ALL_SIZE,
 AMSI_ATTR_ALL_ADDRESS, AMSI_ATTR_QUIET) = range(10)

#Easier is to just scan buffers:

def get_scan_description(score):
    """Returns a description given a risk level returned from a scan_file

    Parameters
    ----------
    score : int
        risk score returned from a scan call

    Returns
    -------
    str
        description of the threat level
    """
    if score == RESULT_CLEAN:
        return "File considered clean"
    elif score == RESULT_NOT_DETECTED:
        return "No threats detected"
    elif score in range(RESULT_BLOCKED_BY_ADMIN_START, RESULT_BLOCKED_BY_ADMIN_END):
        return "Threats blocked by administrator"
    elif score == RESULT_DETECTED:
        return "File considered malware"
    else:
        return "N/A"

def summarize_scan_result(filename, file_length, is_malware, score):
    #Returns a summary of a scan result as a dictionary
    return {"Filename" : filename,
            "File size" : file_length,
            "Risk level": score,
            "Scan result": get_scan_description(score),
            "Is Malware": is_malware}

def get_file_info(filename):
    #Returns a file size dword and buufer containing file bytes
    GENERIC_READ = 0x80000000
    NO_SHARE_MODE = 0
    OPEN_EXISTING = 3
    FILE_ATTRIBUTES_NORMAL = 0x80
    CreateFile = kernellib.CreateFileW
    CreateFile.argtypes = [LPCWSTR, DWORD, DWORD, c_void_p, DWORD, DWORD, c_void_p]
    CreateFile.restype = HANDLE
    f_handle = CreateFile(filename,GENERIC_READ,NO_SHARE_MODE,
                           None,OPEN_EXISTING, FILE_ATTRIBUTES_NORMAL, None)
    if (not f_handle) or f_handle == INVALID_HANDLE_VALUE:
        raise Exception("Couldn't open the requested file. May be hidden, archive or a system file")
    f_handle = HANDLE(f_handle)
    filesize = kernellib.GetFileSize(f_handle, None)
    if (not filesize) or (filesize == INVALID_FILE_SIZE):
        raise Exception("Couldn't get file size")

    MEM_COMMIT = 0x1000
    PAGE_READWRITE = 0x4
    VirtualAlloc = kernellib.VirtualAlloc
    VirtualAlloc.argtypes = [c_void_p, c_size_t, DWORD, DWORD]
    VirtualAlloc.restype = c_void_p
    byte_buffer = VirtualAlloc(None, filesize, MEM_COMMIT, PAGE_READWRITE)
    byte_buffer = cast(byte_buffer, c_char_p)
    if not byte_buffer:
        raise Exception("Couldn't allocate memory for file")

    bytes_read = DWORD()
    if not kernellib.ReadFile(f_handle, byte_buffer, DWORD(filesize),
                              byref(bytes_read), None):
        raise Exception("Couldn't read file into memory")

    kernellib.CloseHandle(f_handle)
    return filesize, byte_buffer

class _HAMSICONTEXT(Structure):
    _fields_ = [("i", c_int)]
HAMSICONTEXT = POINTER(_HAMSICONTEXT)

class _HAMSISESSION(Structure):
    _fields_ = [("i", c_int)]
HAMSISESSION = POINTER(_HAMSISESSION)

def _scan(filename, raw_bytes, buffer_size):
    #Scans a buffer containing a file's bytes for malware

    #Returns a boolean indicating whether the buffer contains malware, and risk score
    amsi_context = HAMSICONTEXT()
    memset(byref(amsi_context), 0, sizeof(HAMSICONTEXT))
    amsi_session = HAMSISESSION()

    result = amsilib.AmsiInitialize(LPCWSTR(PROJECT_NAME),
                                    byref(amsi_context))
    if result != S_OK:
        raise Exception("Couldn't initialize the AMSI API")

    result = amsilib.AmsiOpenSession(amsi_context,
                                     byref(amsi_session))
    if result != S_OK:
        raise Exception("Coulndn't open an AMSI session")
    
    amsi_result = c_int()
    result = amsilib.AmsiScanBuffer(amsi_context,
                                    raw_bytes,
                                    DWORD(buffer_size),
                                    LPCWSTR(filename),
                                    amsi_session,
                                    byref(amsi_result))
    if result != S_OK:
        raise Exception("Couldn't scan the requested buffer.\nIs Real Time Protection enabled?")
    risk_level = amsi_result.value
    is_malware = risk_level >= RESULT_DETECTED

    amsilib.AmsiUninitialize(amsi_context)
    return is_malware, risk_level

def scan_file(filename):
    """Scans a specified file for malware.

    To not throw an exception, two settings should be enabled:
    * 'Real-time protection' in Windows Defender
    * 'Scan all downloaded files and attachments' in Local Group Policy Editor

    Parameters
    ----------
    filename : str
        absolute path to the file that should be scanned

    Returns
    -------
    dict
        Information about the scan. Has the following keys:
        'Filename' : str
            Same absolute path to file
        'File size' : int
            Length of the file in bytes
        'Risk level' : int
            One of RESULT_CLEAN, RESULT_NOT_DETECTED, RESULT_DETECTED,
            RESULT_BLOCKED_BY_ADMIN_START, RESULT_BLOCKED_BY_ADMIN_END
        'Scan result' : str
            Description of the risk level
        'Is Malware' : bool
            Determined result of whether scanned file is malware
    """
    if len(filename) == 0:
        raise Exception("Null filename")

    file_size, file_bytes = get_file_info(filename)
    is_malware, risk_level = _scan(filename, file_bytes, file_size)
    MEM_DECOMMIT = 0x4000
    if not kernellib.VirtualFree(file_bytes, DWORD(file_size),
                                 DWORD(MEM_DECOMMIT)):
        raise Exception("Couldn't Free Memory")
    return summarize_scan_result(filename, file_size, is_malware, risk_level)
    
