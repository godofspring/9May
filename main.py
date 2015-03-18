











__author__ = "Constantine Aitov"
__date__ = "$07.03.2015 12:52:15$"
__version__ = "1.0 Beta 5 RC"















import winreg 
import subprocess 
import time 
import win32api 
import win32gui 
 
#Путь к презентации
_path_to_present = 'c:\peremoga.ppt'


_our_programm = 'powerpoint'
toplist = []
winlist = []
Er = 0

try:
    import ntplib

    client = ntplib.NTPClient()
    response = client.request('pool.ntp.org')

    _year = int(time.strftime('%Y',time.localtime(response.tx_time)))
    _month = int(time.strftime('%m',time.localtime(response.tx_time)))
    _day = int(time.strftime('%d',time.localtime(response.tx_time)))
    _hour = int(time.strftime('%H',time.localtime(response.tx_time))) - 3
    _minutes = int(time.strftime('%M',time.localtime(response.tx_time)))
    _second = int(time.strftime('%S',time.localtime(response.tx_time)))
    _millisecond = int(0)
    _day_of_week = (int(time.strftime('%w',time.localtime(response.tx_time))))

    win32api.SetSystemTime(_year, _month, _day_of_week, _day, _hour,
    _minutes, _second, _millisecond) 
    
except:
    print ('Невозможно получить или установить точное время.' + chr(13) +
    'Вероятно нет соединения с интернетом, прав администратора' + chr(13) +
    'или не установлен модуль ntplib. Для установки выполните pip install ntplib')
    time.sleep (10)


def right_day ():
    _value = 0 
    _array_of_D_days = [91+i for i in range(40)] 
    _array_of_days_before = [1+j for j in range(90)]
    for _day1 in _array_of_D_days: 
        if int(time.strftime("%j", time.gmtime())) == _day1: _value = 1 
    for _day2 in _array_of_days_before: 
        if int(time.strftime("%j", time.gmtime())) == _day2: _value = 2 
    return int(_value)
    
def enum_callback(hwnd, results):
    winlist.append((hwnd, win32gui.GetWindowText(hwnd)))    
    
def get_registry_value(key, subkey, value):
    key = getattr(winreg, key)
    handle = winreg.OpenKey(key, subkey )
    (value, type) = winreg.QueryValueEx(handle, value)
    return value

windowsbit=cputype = get_registry_value("HKEY_LOCAL_MACHINE",
    "SYSTEM\\CurrentControlSet\Control\\Session Manager\\Environment",
    "PROCESSOR_ARCHITECTURE")

 
_pptpath = ''

if windowsbit == "AMD64":
    _pptpath = '"C:\Program Files (x86)\Microsoft Office\PowerPoint Viewer\pptview.exe" ' + _path_to_present +' /f'
else:
    _pptpath = '"C:\Program Files\Microsoft Office\PowerPoint Viewer\pptview.exe" ' + _path_to_present + ' /f'

_var_ = 1

while _var_ == 1 :
    if right_day() == 1:
        
        if int(time.strftime("%M", time.gmtime())) == int(0):
            proc = subprocess.Popen(_pptpath)
            try:
                win32gui.EnumWindows(enum_callback, toplist)
                _window = [(hwnd, title) for hwnd, title in winlist if _our_programm in title.lower()]
                _window = _window[0]
                win32gui.ShowWindow(_window[0], 5)
                win32gui.SetForegroundWindow(_window[0])
                win32gui.SetFocus(_window[0])
                Er = 0
            except:
                Er = 1
            try:
                outs, errs = proc.communicate(timeout=40)
            except:
                proc.kill()
                time.sleep(20.5)
                
        elif int(time.strftime("%M", time.gmtime())) == int(20):
            proc = subprocess.Popen(_pptpath)
            try:
                win32gui.EnumWindows(enum_callback, toplist)
                _window = [(hwnd, title) for hwnd, title in winlist if _our_programm in title.lower()]
                _window = _window[0]
                win32gui.ShowWindow(_window[0], 5)
                win32gui.SetForegroundWindow(_window[0])
                win32gui.SetFocus(_window[0])
                Er = 0
            except:
                Er = 1
            try:
                outs, errs = proc.communicate(timeout=40)
            except:
                proc.kill()
                time.sleep(20.5)
                
        elif int(time.strftime("%M", time.gmtime())) == int(40):
            proc = subprocess.Popen(_pptpath)
            try:
                win32gui.EnumWindows(enum_callback, toplist)
                _window = [(hwnd, title) for hwnd, title in winlist if _our_programm in title.lower()]
                _window = _window[0]
                win32gui.ShowWindow(_window[0], 5)
                win32gui.SetForegroundWindow(_window[0])
                win32gui.SetFocus(_window[0])
                Er = 0
            except: 
                Er = 1
            try:
                outs, errs = proc.communicate(timeout=40)
            except:
                proc.kill()
                time.sleep(20.5)
                
        time.sleep(0.01)
        
    elif right_day() == 2:
        sleep (3600)
    
    elif right_day() == 3:
        quit()
        
    
