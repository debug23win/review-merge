"""Check the built application opts out of Windows bitmap DPI virtualization."""
import ctypes, os, subprocess, time, uuid
from ctypes import wintypes
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'tests/.runs'/uuid.uuid4().hex
OUT.mkdir(parents=True)
if os.name!='nt':
    raise SystemExit('DPI test requires Windows')
query=ctypes.WinDLL('shcore').GetProcessDpiAwareness
query.argtypes=[wintypes.HANDLE,ctypes.POINTER(ctypes.c_int)]
query.restype=ctypes.c_long
image=OUT/'window.png'
p=subprocess.Popen([str(ROOT/'Свод_проверки.exe'),'--snapshot',str(image)],creationflags=0x08000000)
seen=[]
deadline=time.monotonic()+30
while p.poll() is None and time.monotonic()<deadline:
    value=ctypes.c_int()
    if query(wintypes.HANDLE(int(p._handle)),ctypes.byref(value))==0:seen.append(value.value)
    time.sleep(0.02)
assert p.wait(timeout=30)==0 and image.exists(),'Snapshot failed'
assert seen and max(seen)>=1,('Windows still scales a 96-DPI bitmap',seen)
print('DPI test passed: native process is DPI-aware; the window is rendered at the system scale')
print(image)
