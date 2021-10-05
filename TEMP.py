import sys
if hasattr(sys,'frozen') and (sys.frozen=="windows_exe" or sys.frozen=="console_exe"):
    texturePath="visual\\"
else:
    texturePath = "C:\Users\sasi\Desktop\learning python\basics" + "/"
del sys
