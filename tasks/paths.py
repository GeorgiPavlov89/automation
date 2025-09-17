# automation/tasks/paths.py
from __future__ import annotations
from typing import Optional
from pathlib import Path
import os, ctypes
from ctypes import wintypes

from automation.orchestrator import task

# GUID за Desktop (Known Folders)
# FOLDERID_Desktop = B4BFCC3A-DB2C-424C-B029-7FE99A87C641
class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]

FOLDERID_Desktop = GUID(
    0xB4BFCC3A, 0xDB2C, 0x424C,
    (ctypes.c_ubyte * 8)(0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41)
)

def _known_folder_path(fid: GUID) -> Optional[Path]:
    ppsz = wintypes.LPWSTR()
    shget = ctypes.windll.shell32.SHGetKnownFolderPath
    shget.argtypes = [ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(wintypes.LPWSTR)]
    shget.restype  = ctypes.HRESULT
    hr = shget(ctypes.byref(fid), 0, None, ctypes.byref(ppsz))
    if hr != 0:
        return None
    try:
        return Path(ppsz.value) if ppsz.value else None
    finally:
        ctypes.windll.ole32.CoTaskMemFree(ppsz)

def _desktop_from_onedrive() -> Optional[Path]:
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        od = os.environ.get(var)
        if od:
            p = Path(od) / "Desktop"
            if p.exists():
                return p
    return None

def _desktop_from_registry() -> Optional[Path]:
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        ) as key:
            val, _ = winreg.QueryValueEx(key, "Desktop")
        p = Path(os.path.expandvars(val))
        return p if p.exists() else None
    except Exception:
        return None

@task("get_desktop_dir")
def get_desktop_dir() -> str:
   
    
    p = _known_folder_path(FOLDERID_Desktop)
    if p and p.exists():
        return str(p)
    p = _desktop_from_onedrive()
    if p:
        return str(p)
    p = _desktop_from_registry()
    if p:
        return str(p)
    up = Path(os.environ.get("USERPROFILE", str(Path.home())))
    return str(up / "Desktop")
