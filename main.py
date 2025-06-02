import os
import sys
import time
import argparse
import threading
import logging
import servicemanager
import win32service
import win32serviceutil
import win32event
import win32api

# Import our custom modules
from config_manager import ConfigManager
from update_checker import WingetUpdateChecker
from system_tray import SystemTrayIcon

class WingetUpdaterService(win32serviceutil.ServiceFramework):
    """Windows service for Winget Updater"""
    
    _svc_name_ = "WingetUpdaterService"
    _svc_display_name_ = "Winget Updater Service"
    _svc_description_ = "Monitors and notifies about available Winget updates"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = False
        self.system_tray = None
        self.tray_thread = None
        
        # Setup logging
        self._setup_logging()
        
        # Initialize config manager
        self.config_manager = ConfigManager()
    
    def _setup_logging(self):
        """Set up logging for the service"""
        log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'winget_updater_service.log')
        
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        self.logger = logging.getLogger('WingetUpdaterService')
    
    def SvcStop(self):
        """Stop the service"""
        self.logger.info('Service stop requested')
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        self.running = False
        
        # Stop the system tray if it's running
        if self.system_tray:
            try:
                self.system_tray.running = False
                if hasattr(self.system_tray, 'icon') and self.system_tray.icon:
                    self.system_tray.icon.stop()
            except Exception as e:
                self.logger.error(f"Error stopping system tray: {str(e)}")
    
    def SvcDoRun(self):
        """Run the service"""
        self.logger.info('Service starting')
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        self.running = True
        
        try:
            # Start the system tray in a separate thread
            self.tray_thread = threading.Thread(target=self._run_system_tray, daemon=True)
            self.tray_thread.start()
            
            # Main service loop
            while self.running:
                # Check if we need to stop
                if win32event.WaitForSingleObject(self.stop_event, 1000) == win32event.WAIT_OBJECT_0:
                    break
                    
                time.sleep(1)
                
        except Exception as e:
            self.logger.error(f"Service error: {str(e)}")
            self.running = False
            
        self.logger.info('Service stopped')
    
    def _run_system_tray(self):
        """Run the system tray application"""
        try:
            self.system_tray = SystemTrayIcon()
            self.system_tray.run()
        except Exception as e:
            self.logger.error(f"Error in system tray thread: {str(e)}")

def run_standalone():
    """Run the application as a standalone app (not as a service)"""
    print("Starting Winget Updater in standalone mode...")
    
    # Set up logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("winget_updater.log"),
            logging.StreamHandler()
        ]
    )
    
    logger = logging.getLogger('WingetUpdaterStandalone')
    logger.info("Starting application in standalone mode")
    
    try:
        # Create and run the system tray application
        system_tray = SystemTrayIcon()
        system_tray.run()
    except Exception as e:
        logger.error(f"Error in standalone mode: {str(e)}")
        print(f"Error: {str(e)}")
    
    logger.info("Application exited")

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
    
    # Add standalone mode
    parser.add_argument('--standalone', action='store_true', help='Run as a standalone application (not as a service)')
    
    # Add debug mode
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    
    return parser.parse_args()

def main():
    """Main entry point for the application"""
    args = parse_arguments()
    
    # Handle service installation/uninstallation
    if args.install:
        try:
            win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'install'])
            print("Service installed successfully")
        except Exception as e:
            print(f"Error installing service: {str(e)}")
        return
        
    if args.uninstall:
        try:
            win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'remove'])
            print("Service uninstalled successfully")
        except Exception as e:
            print(f"Error uninstalling service: {str(e)}")
        return
        
    if args.start:
        try:
            win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'start'])
            print("Service started successfully")
        except Exception as e:
            print(f"Error starting service: {str(e)}")
        return
        
    if args.stop:
        try:
            win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'stop'])
            print("Service stopped successfully")
        except Exception as e:
            print(f"Error stopping service: {str(e)}")
        return
        
    if args.restart:
        try:
            win32serviceutil.HandleCommandLine(WingetUpdaterService, argv=['', 'restart'])
            print("Service restarted successfully")
        except Exception as e:
            print(f"Error restarting service: {str(e)}")
        return
    
    # If no service commands, run in standalone mode or handle service events
    if args.standalone:
        run_standalone()
    else:
        # If run by the service manager, the script was called without arguments
        if len(sys.argv) == 1:
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(WingetUpdaterService)
            servicemanager.StartServiceCtrlDispatcher()
        else:
            # Otherwise, handle the command line
            win32serviceutil.HandleCommandLine(WingetUpdaterService)

if __name__ == '__main__':
    main()

