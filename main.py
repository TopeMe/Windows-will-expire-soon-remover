import subprocess
import winreg

#Restarts Windows Explorer
def end_task_windows_explorer():
    try:
        subprocess.run('taskkill /f /im explorer.exe', check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while terminating Windows Explorer: {e}")
        return False

    return True


def start_windows_explorer():
    try:
        subprocess.Popen('explorer.exe')
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while starting Windows Explorer: {e}")
        return False

    return True

#Runs command as admin (slmgr -rearm)
def run_cmd_as_admin(command):
    try:
        subprocess.run(["powershell", "-Command", f"Start-Process cmd.exe -Verb RunAs -ArgumentList '/c {command}'"],
                       check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while running command as administrator: {e}")
        return False

    return True


#Enables no auto restart logged users
def enable_no_auto_restart():
    try:
        key_path = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
            winreg.SetValueEx(key, "NoAutoRebootWithLoggedOnUsers", 0, winreg.REG_DWORD, 1)
        return True
    except OSError as e:
        print(f"Error occurred while modifying the registry: {e}")
        return False

#Runs Functions step by step
if end_task_windows_explorer():
    print("Step one process successful.")
    if start_windows_explorer():
        print("Step two process successful.")
        if run_cmd_as_admin("slmgr -rearm"):
            print("Step three process successful.")
            if enable_no_auto_restart():
                print("Step four process successful.")
            else:
                print("Step four failed.Run as administrator")
        else:
            print("Step three failed.Run as administrator")
    else:
        print("Step three failed.Run as administrator")
else:
    print("Step one failed.Run as administrator")
