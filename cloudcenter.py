import os
import subprocess
import ctypes
from typing import Optional

def run_command(command: str) -> Optional[str]:
    """Runs a system command and returns the output."""
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Command '{command}' failed with error: {e}")
        return None

def check_admin() -> bool:
    """Checks if the script is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def restart_windows_search_service():
    """Restarts the Windows Search service."""
    print("Restarting Windows Search service...")
    run_command("net stop WSearch")
    run_command("net start WSearch")
    print("Windows Search service restarted.")

def rebuild_search_index():
    """Rebuilds the search index."""
    print("Rebuilding Windows Search Index...")
    command = (
        "powershell.exe -Command \"& {Start-Process PowerShell -Verb RunAs -ArgumentList 'Set-ExecutionPolicy Unrestricted -Scope Process -Force; "
        "Import-Module -Name Search -Force; "
        "Reset-WindowsSearchIndex; "
        "Set-ExecutionPolicy Restricted -Scope Process -Force'}\""
    )
    output = run_command(command)
    if output is not None:
        print("Search index rebuilt successfully.")
    else:
        print("Failed to rebuild search index.")

def repair_windows_search():
    """Performs all repair operations for Windows Search."""
    if not check_admin():
        print("Administrator privileges are required to run this script.")
        return

    print("Starting Windows Search repair...")
    restart_windows_search_service()
    rebuild_search_index()
    print("Windows Search repair completed.")

if __name__ == "__main__":
    repair_windows_search()