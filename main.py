import os
import subprocess
import winreg

def end_task_windows_explorer():
    # Terminate Windows Explorer process
    try:
        subprocess.run('taskkill /f /im explorer.exe', check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while terminating Windows Explorer: {e}")
        return False

    return True


def start_windows_explorer():
    # Start Windows Explorer process
    try:
        subprocess.Popen('explorer.exe')
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while starting Windows Explorer: {e}")
        return False

    return True

def run_cmd_as_admin(command):
    # Run command prompt as administrator
    try:
        subprocess.run(["powershell", "-Command", f"Start-Process cmd.exe -Verb RunAs -ArgumentList '/c {command}'"],
                       check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while running command as administrator: {e}")
        return False

    return True


def enable_no_auto_restart():
    try:
        # Open the Windows Registry key
        key_path = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
            # Set the value to enable "No auto-restart with logged on users for scheduled automatic updates installations"
            winreg.SetValueEx(key, "NoAutoRebootWithLoggedOnUsers", 0, winreg.REG_DWORD, 1)
        return True
    except OSError as e:
        print(f"Error occurred while modifying the registry: {e}")
        return False


# End Windows Explorer process
if end_task_windows_explorer():
    print("Step one process successful.")
    # Start Windows Explorer process
    if start_windows_explorer():
        print("Step two process successful.")
        # Run cmd as admin and execute slmgr -rearm
        if run_cmd_as_admin("slmgr -rearm"):
            print("Step three process successful.")

            # Enable "No auto-restart with logged on users for scheduled automatic updates installations"
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
