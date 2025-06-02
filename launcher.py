import os
import sys
import time
import logging
import argparse
import subprocess
import threading
import ctypes
import win32serviceutil
import win32service
import win32event
import servicemanager
import pywintypes

# Import our custom modules
from service_component import WingetUpdaterService, run_service, run_service_debug
from ui_component import run_tray_application

def is_admin():
    """Check if the current process has administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def run_as_admin(argv=None):
    """Re-run the program with admin rights"""
    if argv is None:
        argv = sys.argv
    if not is_admin():
        # Try to run the script as admin
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(['"{}"'.format(arg) for arg in argv]), None, 1
        )
        return True
    return False

def setup_logging(debug=False):
    """Set up logging for the launcher"""
    level = logging.DEBUG if debug else logging.INFO
    
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("winget_updater_launcher.log"),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger('WingetUpdaterLauncher')

def install_service():
    """Install the Windows service"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    
    if not is_admin():
        logger.error("Administrator privileges are required to install the service.")
        print("Error: Administrator privileges are required to install the service.")
        print("Please run the command prompt as Administrator and try again.")
        return False
    
    try:
        win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'install'])
        logger.info("Service installed successfully")
        print("Service installed successfully")
        return True
    except Exception as e:
        logger.error(f"Error installing service: {str(e)}")
        print(f"Error installing service: {str(e)}")
        return False

def uninstall_service():
    """Uninstall the Windows service"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    
    if not is_admin():
        logger.error("Administrator privileges are required to uninstall the service.")
        print("Error: Administrator privileges are required to uninstall the service.")
        print("Please run the command prompt as Administrator and try again.")
        return False
    
    try:
        win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'remove'])
        logger.info("Service uninstalled successfully")
        print("Service uninstalled successfully")
        return True
    except Exception as e:
        logger.error(f"Error uninstalling service: {str(e)}")
        print(f"Error uninstalling service: {str(e)}")
        return False

def start_service():
    """Start the Windows service"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    
    if not is_admin():
        logger.error("Administrator privileges are required to start the service.")
        print("Error: Administrator privileges are required to start the service.")
        print("Please run the command prompt as Administrator and try again.")
        return False
    
    try:
        win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'start'])
        logger.info("Service started successfully")
        print("Service started successfully")
        return True
    except Exception as e:
        logger.error(f"Error starting service: {str(e)}")
        print(f"Error starting service: {str(e)}")
        return False

def stop_service():
    """Stop the Windows service"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    
    if not is_admin():
        logger.error("Administrator privileges are required to stop the service.")
        print("Error: Administrator privileges are required to stop the service.")
        print("Please run the command prompt as Administrator and try again.")
        return False
    
    try:
        win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'stop'])
        logger.info("Service stopped successfully")
        print("Service stopped successfully")
        return True
    except Exception as e:
        logger.error(f"Error stopping service: {str(e)}")
        print(f"Error stopping service: {str(e)}")
        return False

def restart_service():
    """Restart the Windows service"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    
    if not is_admin():
        logger.error("Administrator privileges are required to restart the service.")
        print("Error: Administrator privileges are required to restart the service.")
        print("Please run the command prompt as Administrator and try again.")
        return False
    
    try:
        win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'restart'])
        logger.info("Service restarted successfully")
        print("Service restarted successfully")
        return True
    except Exception as e:
        logger.error(f"Error restarting service: {str(e)}")
        print(f"Error restarting service: {str(e)}")
        return False

def run_ui_only():
    """Run only the UI component without the service"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    logger.info("Starting UI component only")
    
    try:
        run_tray_application()
        return True
    except Exception as e:
        logger.error(f"Error running UI component: {str(e)}")
        print(f"Error running UI component: {str(e)}")
        return False

def run_service_only():
    """Run only the service component"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    logger.info("Starting service component only")
    
    if len(sys.argv) == 1:
        # If run by the service manager
        try:
            run_service()
            return True
        except Exception as e:
            logger.error(f"Error running service component: {str(e)}")
            return False
    else:
        # If run from command line
        try:
            win32serviceutil.HandleCommandLine(WingetUpdaterService)
            return True
        except Exception as e:
            logger.error(f"Error running service component: {str(e)}")
            print(f"Error running service component: {str(e)}")
            return False

def run_debug_mode():
    """Run both the service and UI components in debug mode"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    logger.info("Starting application in debug mode")
    
    # Start the service in a separate thread
    service_thread = threading.Thread(target=run_service_debug)
    service_thread.daemon = True
    service_thread.start()
    
    # Give the service time to start
    time.sleep(2)
    
    # Start the UI
    try:
        run_tray_application()
        return True
    except Exception as e:
        logger.error(f"Error running in debug mode: {str(e)}")
        print(f"Error running in debug mode: {str(e)}")
        return False

def run_standalone_mode():
    """Run both the service and UI components in standalone mode"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    logger.info("Starting application in standalone mode")
    
    # Check if the service is already running
    try:
        service_status = win32serviceutil.QueryServiceStatus(WingetUpdaterService._svc_name_)
        if service_status[1] == win32service.SERVICE_RUNNING:
            logger.info("Service is already running")
            # Just start the UI
            return run_ui_only()
    except pywintypes.error:
        # Service not installed or not running
        pass
    
    # Try to start the service if admin
    if is_admin():
        try:
            # Check if service is installed
            service_status = win32serviceutil.QueryServiceStatus(WingetUpdaterService._svc_name_)
            # Service exists, try to start it
            start_service()
            time.sleep(2)  # Give the service time to start
        except pywintypes.error:
            # Service not installed, run in debug mode
            return run_debug_mode()
    else:
        # Not admin, run in debug mode
        return run_debug_mode()
    
    # Start the UI
    return run_ui_only()

def autostart_setup(install=True):
    """Set up or remove autostart for the UI component"""
    logger = logging.getLogger('WingetUpdaterLauncher')
    
    try:
        import winreg
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        app_name = "WingetUpdater"
        exe_path = os.path.abspath(sys.argv[0])
        
        # Convert to .exe path if running from .py
        if exe_path.endswith('.py'):
            exe_path = f'"{sys.executable}" "{exe_path}" --ui'
        
        # Open the registry key
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        
        if install:
            # Add to autostart
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
            logger.info(f"Added {app_name} to autostart")
            print(f"Added {app_name} to autostart")
        else:
            # Remove from autostart
            try:
                winreg.DeleteValue(key, app_name)
                logger.info(f"Removed {app_name} from autostart")
                print(f"Removed {app_name} from autostart")
            except FileNotFoundError:
                logger.info(f"{app_name} was not in autostart")
                print(f"{app_name} was not in autostart")
        
        winreg.CloseKey(key)
        return True
    except Exception as e:
        logger.error(f"Error managing autostart: {str(e)}")
        print(f"Error managing autostart: {str(e)}")
        return False

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description="Winget Updater - Windows package update monitor")
    
    # Add argument group for service management
    service_group = parser.add_argument_group('Service Management')
    service_group.add_argument('--install', action='store_true', help='Install as a Windows service')
    service_group.add_argument('--uninstall', action='store_true', help='Uninstall the Windows service')
    service_group.add_argument('--start', action='store_true', help='Start the Windows service')
    service_group.add_argument('--stop', action='store_true', help='Stop the Windows service')
    service_group.add_argument('--restart', action='store_true', help='Restart the Windows service')
    
    # Add argument group for run modes
    run_group = parser.add_argument_group('Run Modes')
    run_group.add_argument('--service', action='store_true', help='Run as a service only (for service manager)')
    run_group.add_argument('--ui', action='store_true', help='Run the UI component only')
    run_group.add_argument('--standalone', action='store_true', help='Run both service and UI (default mode)')
    run_group.add_argument('--debug', action='store_true', help='Run in debug mode (service not installed)')
    
    # Add argument group for autostart
    autostart_group = parser.add_argument_group('Autostart Management')
    autostart_group.add_argument('--add-autostart', action='store_true', help='Add the application to autostart')
    autostart_group.add_argument('--remove-autostart', action='store_true', help='Remove the application from autostart')
    
    # Add other options
    parser.add_argument('--verbose', action='store_true', help='Enable verbose logging')
    
    return parser.parse_args()

def main():
    """Main entry point for the application"""
    # Parse command line arguments
    args = parse_arguments()
    
    # Setup logging
    logger = setup_logging(args.verbose or args.debug)
    logger.info("Winget Updater starting")
    
    # Handle autostart management
    if args.add_autostart:
        return autostart_setup(True)
        
    if args.remove_autostart:
        return autostart_setup(False)
    
    # Handle service management
    if args.install:
        return install_service()
        
    if args.uninstall:
        return uninstall_service()
        
    if args.start:
        return start_service()
        
    if args.stop:
        return stop_service()
        
    if args.restart:
        return restart_service()
    
    # Handle run modes
    if args.service:
        return run_service_only()
        
    if args.ui:
        return run_ui_only()
        
    if args.debug:
        return run_debug_mode()
    
    # Default to standalone mode
    return run_standalone_mode()

if __name__ == '__main__':
    try:
        exit_code = 0 if main() else 1
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\nApplication terminated by user.")
        sys.exit(0)
    except Exception as e:
        logger = logging.getLogger('WingetUpdaterLauncher')
        logger.error(f"Unhandled exception: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}")
        sys.exit(1)

