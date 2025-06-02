import os
import configparser
from datetime import datetime
import logging

class ConfigManager:
    """Manages the configuration settings for the Winget Updater application"""
    
    def __init__(self, config_file='settings.ini'):
        # Define AppData directory for configuration
        self.config_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'WingetUpdater')
        
        # Create the directory if it doesn't exist
        if not os.path.exists(self.config_dir):
            try:
                os.makedirs(self.config_dir, exist_ok=True)
                logging.info(f"Created configuration directory: {self.config_dir}")
            except Exception as e:
                logging.error(f"Failed to create configuration directory: {str(e)}")
                # Fall back to current directory if AppData is not accessible
                self.config_dir = os.path.dirname(os.path.abspath(__file__))
                logging.warning(f"Using fallback directory: {self.config_dir}")
        
        # Use absolute path for config file
        self.config_file = os.path.join(self.config_dir, config_file)
        self.config = configparser.ConfigParser()
        
        # Set default values
        self.config['Settings'] = {
            'morning_check': '08:00',
            'afternoon_check': '16:00',
            'notify_on_updates': 'True',
            'last_check': '',
            'auto_check': 'True',
            'include_pinned_updates': 'False',
            'include_unknown_versions': 'False'
        }
        
        # Load existing config if it exists
        if os.path.exists(self.config_file):
            try:
                self.config.read(self.config_file)
                logging.info(f"Loaded configuration from {self.config_file}")
            except Exception as e:
                logging.error(f"Failed to load configuration: {str(e)}")
    
    def save_config(self):
        """Save the current configuration to the config file"""
        try:
            with open(self.config_file, 'w') as f:
                self.config.write(f)
            logging.info(f"Configuration saved to {self.config_file}")
        except Exception as e:
            logging.error(f"Failed to save configuration: {str(e)}")
    
    def get_morning_check_time(self):
        """Get the morning check time"""
        return self.config['Settings']['morning_check']
    
    def get_afternoon_check_time(self):
        """Get the afternoon check time"""
        return self.config['Settings']['afternoon_check']
    
    def set_morning_check_time(self, time_str):
        """Set the morning check time"""
        self.config['Settings']['morning_check'] = time_str
        self.save_config()
    
    def set_afternoon_check_time(self, time_str):
        """Set the afternoon check time"""
        self.config['Settings']['afternoon_check'] = time_str
        self.save_config()
    
    def get_notify_on_updates(self):
        """Get whether to notify on updates"""
        return self.config['Settings'].getboolean('notify_on_updates')
    
    def set_notify_on_updates(self, notify):
        """Set whether to notify on updates"""
        self.config['Settings']['notify_on_updates'] = str(notify)
        self.save_config()
    
    def get_last_check(self):
        """Get the last check time"""
        last_check = self.config['Settings']['last_check']
        if last_check:
            return datetime.fromisoformat(last_check)
        return None
    
    def set_last_check(self, check_time=None):
        """Set the last check time"""
        if check_time is None:
            check_time = datetime.now()
        self.config['Settings']['last_check'] = check_time.isoformat()
        self.save_config()
    
    def get_auto_check(self):
        """Get whether to automatically check for updates"""
        return self.config['Settings'].getboolean('auto_check')
    
    def set_auto_check(self, auto_check):
        """Set whether to automatically check for updates"""
        self.config['Settings']['auto_check'] = str(auto_check)
        self.save_config()

    def get_include_pinned_updates(self):
        """Get whether to include pinned updates in update checks"""
        # Ensure backward compatibility by returning False if the setting doesn't exist
        if 'include_pinned_updates' not in self.config['Settings']:
            self.config['Settings']['include_pinned_updates'] = 'False'
            self.save_config()
        return self.config['Settings'].getboolean('include_pinned_updates')
    
    def set_include_pinned_updates(self, include_pinned):
        """Set whether to include pinned updates in update checks"""
        self.config['Settings']['include_pinned_updates'] = str(include_pinned)
        self.save_config()
    
    def get_include_unknown_versions(self):
        """Get whether to include packages with unknown versions in update checks"""
        # Ensure backward compatibility by returning False if the setting doesn't exist
        if 'include_unknown_versions' not in self.config['Settings']:
            self.config['Settings']['include_unknown_versions'] = 'False'
            self.save_config()
        return self.config['Settings'].getboolean('include_unknown_versions')
    
    def set_include_unknown_versions(self, include_unknown):
        """Set whether to include packages with unknown versions in update checks"""
        self.config['Settings']['include_unknown_versions'] = str(include_unknown)
        self.save_config()

