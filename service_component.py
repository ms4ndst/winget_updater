import os
import sys
import time
import logging
import threading
import win32serviceutil
import win32service
import win32event
import servicemanager
from datetime import datetime

# Import our custom modules
from config_manager import ConfigManager
from update_checker import WingetUpdateChecker
from ipc_handler import IPCServer

class WingetUpdaterService(win32serviceutil.ServiceFramework):
    """Windows service for Winget Updater"""
    
    _svc_name_ = "WingetUpdaterService"
    _svc_display_name_ = "Winget Updater Service"
    _svc_description_ = "Monitors and notifies about available Winget updates"
    
    def __init__(self, args=None, debug_mode=False):
        # Only initialize the service framework if not in debug mode
        if not debug_mode and args is not None:
            win32serviceutil.ServiceFramework.__init__(self, args)
            
        self.debug_mode = debug_mode
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = False
        
        # Setup logging
        self._setup_logging()
        
        # Initialize components
        self.config_manager = ConfigManager()
        self.update_checker = WingetUpdateChecker(self.config_manager)
        self.ipc_server = IPCServer()
        
        # Register IPC command handlers
        self._register_command_handlers()
        
        self.logger.info("Service initialized in {} mode".format("debug" if debug_mode else "service"))
    
    def _setup_logging(self):
        """Set up logging for the service"""
        log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'winget_updater_service.log')
        
        # Configure logging with both file and console output in debug mode
        if self.debug_mode:
            logging.basicConfig(
                level=logging.DEBUG,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file),
                    logging.StreamHandler()
                ]
            )
        else:
            logging.basicConfig(
                filename=log_file,
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
        
        self.logger = logging.getLogger('WingetUpdaterService')
    
    def _register_command_handlers(self):
        """Register handlers for IPC commands"""
        self.ipc_server.register_handler("check_updates", self._handle_check_updates)
        self.ipc_server.register_handler("get_status", self._handle_get_status)
        self.ipc_server.register_handler("get_updates", self._handle_get_updates)
        self.ipc_server.register_handler("get_last_check", self._handle_get_last_check)
        self.ipc_server.register_handler("save_settings", self._handle_save_settings)
        self.ipc_server.register_handler("get_settings", self._handle_get_settings)
    
    def SvcStop(self):
        """Stop the service"""
        self.logger.info('Service stop requested')
        
        # Only report status if not in debug mode
        if not self.debug_mode:
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            
        win32event.SetEvent(self.stop_event)
        self.running = False
        
        # Stop the IPC server
        self.ipc_server.stop()
    
    def SvcDoRun(self):
        """Run the service"""
        self.logger.info('Service starting')
        
        # Only report status if not in debug mode
        if not self.debug_mode:
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            
        self.running = True
        
        try:
            # Start the IPC server
            self.logger.info("Starting IPC server")
            self.ipc_server.start()
            self.logger.info("IPC server started successfully")
            
            # Perform initial update check
            self.logger.info("Performing initial update check")
            self._check_updates()
            
            # Start the scheduler thread
            self.logger.info("Starting scheduler thread")
            self.scheduler_thread = threading.Thread(target=self._run_scheduler)
            self.scheduler_thread.daemon = True
            self.scheduler_thread.start()
            
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
    
    def _run_scheduler(self):
        """Run the scheduler loop to check for updates at specified times"""
        self.logger.info("Starting scheduler thread")
        
        while self.running:
            try:
                # Check if automatic updates are enabled
                if self.config_manager.get_auto_check():
                    # Get current time
                    now = datetime.now()
                    current_time = now.strftime("%H:%M")
                    
                    # Get scheduled times
                    morning_check = self.config_manager.get_morning_check_time()
                    afternoon_check = self.config_manager.get_afternoon_check_time()
                    
                    # Check if it's time to run an update
                    if current_time == morning_check or current_time == afternoon_check:
                        self.logger.info(f"Scheduled update check at {current_time}")
                        self._check_updates()
                
                # Sleep for 30 seconds before checking again
                time.sleep(30)
                
            except Exception as e:
                self.logger.error(f"Error in scheduler thread: {str(e)}")
                time.sleep(60)  # Sleep longer on error
    
    def _check_updates(self):
        """Check for updates"""
        self.logger.info("Checking for updates")
        
        try:
            # Run the update check
            update_count = self.update_checker.check_updates()
            self.logger.info(f"Update check completed. Found {update_count} updates.")
            return update_count
        except Exception as e:
            self.logger.error(f"Error checking for updates: {str(e)}")
            return 0
    
    # IPC Command Handlers
    
    def _handle_check_updates(self, data):
        """Handle the check_updates command"""
        update_count = self._check_updates()
        
        return {
            "update_count": update_count,
            "success": True,
            "last_check": datetime.now().isoformat()
        }
    
    def _handle_get_status(self, data):
        """Handle the get_status command"""
        update_count = self.update_checker.get_update_count()
        last_check = self.update_checker.get_last_check_time()
        
        if last_check is None and self.config_manager.get_last_check():
            last_check = self.config_manager.get_last_check()
            
        return {
            "update_count": update_count,
            "last_check": last_check.isoformat() if last_check else None,
            "auto_check": self.config_manager.get_auto_check(),
            "morning_check": self.config_manager.get_morning_check_time(),
            "afternoon_check": self.config_manager.get_afternoon_check_time()
        }
    
    def _handle_get_updates(self, data):
        """Handle the get_updates command"""
        updates = self.update_checker.get_updates_list()
        
        return {
            "updates": updates,
            "count": len(updates)
        }
    
    def _handle_get_last_check(self, data):
        """Handle the get_last_check command"""
        last_check = self.update_checker.get_last_check_time()
        
        if last_check is None and self.config_manager.get_last_check():
            last_check = self.config_manager.get_last_check()
            
        return {
            "last_check": last_check.isoformat() if last_check else None
        }
    
    def _handle_save_settings(self, data):
        """Handle the save_settings command"""
        try:
            if "morning_check" in data:
                self.config_manager.set_morning_check_time(data["morning_check"])
                
            if "afternoon_check" in data:
                self.config_manager.set_afternoon_check_time(data["afternoon_check"])
                
            if "notify_on_updates" in data:
                self.config_manager.set_notify_on_updates(data["notify_on_updates"])
                
            if "auto_check" in data:
                self.config_manager.set_auto_check(data["auto_check"])
                
            self.logger.info("Settings updated via IPC")
            
            return {
                "success": True
            }
        except Exception as e:
            self.logger.error(f"Error saving settings: {str(e)}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _handle_get_settings(self, data):
        """Handle the get_settings command"""
        return {
            "morning_check": self.config_manager.get_morning_check_time(),
            "afternoon_check": self.config_manager.get_afternoon_check_time(),
            "notify_on_updates": self.config_manager.get_notify_on_updates(),
            "auto_check": self.config_manager.get_auto_check(),
            "last_check": self.config_manager.get_last_check().isoformat() if self.config_manager.get_last_check() else None
        }

def run_service():
    """Run the service"""
    servicemanager.Initialize()
    servicemanager.PrepareToHostSingle(WingetUpdaterService)
    servicemanager.StartServiceCtrlDispatcher()

def run_service_debug():
    """Run the service in debug mode (not as a service)"""
    print("Starting Winget Updater service in debug mode...")
    
    try:
        # Create service instance in debug mode
        service = WingetUpdaterService(debug_mode=True)
        print("Service instance created successfully")
        
        # Run the service
        service.SvcDoRun()
        print("Service running, press Ctrl+C to stop")
        
        # Keep running until KeyboardInterrupt
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nStopping service...")
            service.SvcStop()
            print("Service stopped")
            
        return True
    except Exception as e:
        print(f"Error running service in debug mode: {str(e)}")
        logging.error(f"Error running service in debug mode: {str(e)}", exc_info=True)
        return False

if __name__ == '__main__':
    if len(sys.argv) == 2 and sys.argv[1] == 'debug':
        run_service_debug()
    else:
        win32serviceutil.HandleCommandLine(WingetUpdaterService)

