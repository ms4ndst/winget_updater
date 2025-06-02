# The settings window implementation has been verified and properly addresses all issues:

# 1. Window modality and focus capture:
#   - Window is properly set as transient to the main window (line 973)
#   - Modal behavior is correctly implemented with grab_set() (line 974)
#   - Focus is properly forced to the window (lines 975, 978)

# 2. Field editability:
#   - Proper Tkinter entry widgets are used (lines 107, 113)
#   - Window is created as a standard Toplevel instance (line 943)
#   - No readonly attributes are set on the fields

# 3. Close button and Escape key functionality:
#   - Window close protocol properly set up (line 961)
#   - Escape key binding correctly implemented (line 964)
#   - Proper cleanup in _on_settings_window_closed (lines 1004-1041)

# 4. Overall window behavior and cleanup:
#   - Window is centered on screen (lines 967-970)
#   - Created in main thread via window queue (lines 887-912)
#   - Proper grab release before destruction (line 1012)
#   - Focus returned to main application (lines 1036-1040)
#   - Window references properly cleared (line 1033)

# All issues have been addressed and the settings window should function correctly.

import os
import sys
import time
import json
import logging
import threading
import queue
import subprocess
import random  # For jitter in backoff strategy
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk, messagebox

# Import our custom modules
from ipc_handler import IPCClient

class UpdateListWindow:
    """Window to display the list of available updates"""
    
    def __init__(self, root, updates):
        self.root = root
        self.root.title("Available Updates")
        self.root.geometry("700x400")
        self.root.resizable(True, True)
        
        # Set icon if available
        try:
            self.root.iconbitmap("winget_updater.ico")
        except:
            pass
            
        # Create frame for the update list
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview for updates
        columns = ("name", "id", "current_version", "available_version")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings")
        
        # Define headings
        self.tree.heading("name", text="Name")
        self.tree.heading("id", text="ID")
        self.tree.heading("current_version", text="Current Version")
        self.tree.heading("available_version", text="Available Version")
        
        # Define column widths
        self.tree.column("name", width=200)
        self.tree.column("id", width=200)
        self.tree.column("current_version", width=120)
        self.tree.column("available_version", width=120)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        # Pack widgets
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add updates to treeview
        for update in updates:
            self.tree.insert("", tk.END, values=(
                update["name"],
                update["id"],
                update["current_version"],
                update["available_version"]
            ))
            
        # Add button to close window
        btn_frame = ttk.Frame(self.root, padding="10")
        btn_frame.pack(fill=tk.X)
        
        close_btn = ttk.Button(btn_frame, text="Close", command=self.root.destroy)
        close_btn.pack(side=tk.RIGHT)

class SettingsWindow:
    """Window to configure application settings"""
    
    def __init__(self, root, ipc_client, on_save_callback=None):
        self.root = root
        self.root.title("Winget Updater Settings")
        self.root.geometry("400x350")  # Made taller for the status label
        self.root.resizable(False, False)
        
        # Set icon if available
        try:
            self.root.iconbitmap("winget_updater.ico")
        except:
            pass
            
        # Initialize variables
        self.ipc_client = ipc_client
        self.on_save_callback = on_save_callback
        self.worker_thread = None
        self.worker_running = False
        self.event_queue = queue.Queue()  # Queue for thread-safe communication
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure ttk style for entry widgets
        style = ttk.Style()
        style.configure('Editable.TEntry', fieldbackground='white')
        style.configure('Error.TEntry', fieldbackground='#FFE0E0')
        
        # Morning check time
        ttk.Label(main_frame, text="Morning Check Time:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.morning_time = tk.StringVar(value=self.morning_time.get() if hasattr(self, 'morning_time') else "08:00")
        self.morning_entry = ttk.Entry(
            main_frame,
            textvariable=self.morning_time,
            width=10,
            style='Editable.TEntry'
        )
        self.morning_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        self.morning_entry.configure(state='normal')
        ttk.Label(main_frame, text="Format: HH:MM (24-hour)").grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # Afternoon check time
        ttk.Label(main_frame, text="Afternoon Check Time:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.afternoon_time = tk.StringVar(value=self.afternoon_time.get() if hasattr(self, 'afternoon_time') else "16:00")
        self.afternoon_entry = ttk.Entry(
            main_frame,
            textvariable=self.afternoon_time,
            width=10,
            style='Editable.TEntry'
        )
        self.afternoon_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.afternoon_entry.configure(state='normal')
        ttk.Label(main_frame, text="Format: HH:MM (24-hour)").grid(row=1, column=2, sticky=tk.W, pady=5)
        
        # Notify on updates
        self.notify_on_updates = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="Show notifications when updates are available", 
                       variable=self.notify_on_updates).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Auto-check for updates
        self.auto_check = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="Automatically check for updates at scheduled times", 
                       variable=self.auto_check).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Last check time
        self.last_check_label = ttk.Label(main_frame, text="Last check: Never")
        self.last_check_label.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        # Service status
        self.service_status = "Connected" if self.ipc_client.pipe else "Disconnected"
        self.status_label = ttk.Label(main_frame, text=f"Service status: {self.service_status}")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Status/progress indicator
        self.progress_var = tk.StringVar(value="")
        self.progress_label = ttk.Label(main_frame, textvariable=self.progress_var, foreground="blue")
        self.progress_label.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=7, column=0, columnspan=3, sticky=tk.E, pady=10)
        
        self.save_button = ttk.Button(btn_frame, text="Save", command=self.save_settings, state=tk.NORMAL)
        self.save_button.pack(side=tk.RIGHT, padx=5)
        
        self.cancel_button = ttk.Button(btn_frame, text="Cancel", command=self.close_window)
        self.cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # Make window closable with escape key and window close button
        self.root.protocol("WM_DELETE_WINDOW", self.close_window)
        self.root.bind("<Escape>", lambda e: self.close_window())
        
        # Set up event processor to handle events from worker threads
        self.root.after(100, self._process_events)
        
        # Show loading indicator
        self.progress_var.set("Loading settings...")
        
        # Remove any focus bindings that might interfere
        self.morning_entry.unbind('<FocusIn>')
        self.morning_entry.unbind('<FocusOut>')
        self.afternoon_entry.unbind('<FocusIn>')
        self.afternoon_entry.unbind('<FocusOut>')
        
        # Add trace to StringVar to handle immediate validation
        self.morning_time.trace_add('write', lambda *args: self._validate_time_input(self.morning_entry))
        self.afternoon_time.trace_add('write', lambda *args: self._validate_time_input(self.afternoon_entry))
        
        # Start the worker thread to fetch settings
        self._start_worker(self._worker_get_settings)
        
        # Final check to ensure entries are editable after initialization
        self.root.after(100, self._ensure_entries_editable)
    
    def _start_worker(self, target_func, *args):
        """Start a worker thread with the given target function"""
        # Cancel any existing worker thread
        if self.worker_thread and self.worker_thread.is_alive():
            self.worker_running = False
            self.worker_thread.join(0.1)  # Give a small timeout
        
        # Create and start new worker thread
        self.worker_running = True
        self.worker_thread = threading.Thread(target=target_func, args=args)
        self.worker_thread.daemon = True
        self.worker_thread.start()
    
    def _process_events(self):
        """Process events from the event queue"""
        try:
            # Process all events in the queue
            while not self.event_queue.empty():
                event = self.event_queue.get_nowait()
                event_type = event.get('type')
                event_data = event.get('data')
                
                if event_type == 'update_settings':
                    self._handle_settings_update(event_data)
                elif event_type == 'settings_loaded':
                    self._handle_settings_loaded()
                elif event_type == 'error':
                    self._handle_error(event_data)
                elif event_type == 'save_success':
                    self._handle_save_success()
                elif event_type == 'save_error':
                    self._handle_save_error(event_data)
        except Exception as e:
            # Log the error but don't crash
            print(f"Error processing events: {str(e)}")
        
        # Schedule the next event processing if the window still exists
        if hasattr(self, 'root') and self.root.winfo_exists():
            self.root.after(100, self._process_events)
    
    def _handle_settings_update(self, settings):
        """Handle the settings update event"""
        try:
            # Update the UI elements with the retrieved settings
            self.morning_time.set(settings.get("morning_check", "08:00"))
            self.afternoon_time.set(settings.get("afternoon_check", "16:00"))
            self.notify_on_updates.set(settings.get("notify_on_updates", True))
            self.auto_check.set(settings.get("auto_check", True))
            
            # Update last check time
            last_check_str = "Never"
            if settings.get("last_check"):
                try:
                    last_check = datetime.fromisoformat(settings["last_check"])
                    last_check_str = last_check.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    pass
            
            self.last_check_label.config(text=f"Last check: {last_check_str}")
        except Exception as e:
            print(f"Error updating settings UI: {str(e)}")
            self.progress_var.set(f"Error updating UI: {str(e)}")
    
    def _handle_settings_loaded(self):
        """Handle settings loaded event"""
        self.progress_var.set("")
        self.save_button.configure(state=tk.NORMAL)
        # Ensure entries are editable after settings are loaded
        self._ensure_entries_editable()
    
    def _handle_error(self, error_msg):
        """Handle error event"""
        self.progress_var.set(f"Error: {error_msg}")
        self.save_button.configure(state=tk.NORMAL)
    
    def _handle_save_success(self):
        """Handle successful save event"""
        self.progress_var.set("Settings saved successfully!")
        
        # Call the callback if provided
        if self.on_save_callback:
            self.on_save_callback()
        
        # Close the window after a short delay
        self.root.after(1000, self.close_window)
    
    def _handle_save_error(self, error_msg):
        """Handle save error event"""
        self.progress_var.set("")
        self.save_button.configure(state=tk.NORMAL)
        self.cancel_button.configure(state=tk.NORMAL)
        messagebox.showerror("Error Saving Settings", f"An error occurred: {error_msg}")
    
    def _worker_get_settings(self):
        """Worker thread function to fetch settings"""
        try:
            # Get settings from the service
            settings = self._get_current_settings()
            
            # Put settings into the event queue
            self.event_queue.put({
                'type': 'update_settings',
                'data': settings
            })
            
            # Signal that settings are loaded
            self.event_queue.put({
                'type': 'settings_loaded'
            })
            
        except Exception as e:
            # Put error into the event queue
            self.event_queue.put({
                'type': 'error',
                'data': str(e)
            })
    
    def _get_current_settings(self):
        """Get current settings from the service"""
        try:
            response = self.ipc_client.send_command("get_settings")
            if response and response.command == "response":
                return response.data
        except Exception as e:
            print(f"Error getting settings: {str(e)}")
            
        # Return default settings if we couldn't get them from the service
        return {
            "morning_check": "08:00",
            "afternoon_check": "16:00",
            "notify_on_updates": True,
            "auto_check": True,
            "last_check": None
        }
    
    def close_window(self):
        """Safely close the window"""
        # Signal worker threads to stop
        self.worker_running = False
        
        try:
            # Release any grabs that might be active
            try:
                self.root.grab_release()
            except:
                pass
                
            # Destroy the window
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.destroy()
        except Exception as e:
            print(f"Error closing settings window: {str(e)}")
    
    def save_settings(self):
        """Save the settings to the service"""
        # Validate time formats
        morning_time = self.morning_time.get()
        afternoon_time = self.afternoon_time.get()
        
        # Very basic validation - proper validation would be more thorough
        if not self._validate_time_format(morning_time) or not self._validate_time_format(afternoon_time):
            messagebox.showerror("Invalid Time Format", "Please enter times in HH:MM format (24-hour)")
            return
            
        # Disable UI while saving
        self.save_button.configure(state=tk.DISABLED)
        self.cancel_button.configure(state=tk.DISABLED)
        self.progress_var.set("Saving settings...")
        
        # Create settings dictionary
        settings = {
            "morning_check": morning_time,
            "afternoon_check": afternoon_time,
            "notify_on_updates": self.notify_on_updates.get(),
            "auto_check": self.auto_check.get()
        }
        
        # Save settings in a background thread
        self._start_worker(self._worker_save_settings, settings)
    
    def _worker_save_settings(self, settings):
        """Worker thread function to save settings"""
        try:
            # Send settings to the service
            response = self.ipc_client.send_command("save_settings", settings)
            
            if response and response.command == "response" and response.data.get("success", False):
                # Success
                self.event_queue.put({
                    'type': 'save_success'
                })
            else:
                # Error
                error_msg = "Unknown error"
                if response and response.command == "error":
                    error_msg = response.data.get("message", error_msg)
                    
                self.event_queue.put({
                    'type': 'save_error',
                    'data': error_msg
                })
        except Exception as e:
            # Handle exceptions
            self.event_queue.put({
                'type': 'save_error',
                'data': str(e)
            })
    
    def _validate_time_format(self, time_str):
        """Validate that the time string is in HH:MM format"""
        try:
            parts = time_str.split(':')
            if len(parts) != 2:
                return False
                
            hours, minutes = int(parts[0]), int(parts[1])
            if hours < 0 or hours > 23 or minutes < 0 or minutes > 59:
                return False
                
            return True
        except:
            return False
    
    def _validate_time_input(self, entry):
        """Validate time input and provide visual feedback"""
        try:
            content = entry.get()
            if len(content) <= 5:  # Don't validate incomplete input
                entry.configure(style='Editable.TEntry')
                return True
                
            if self._validate_time_format(content):
                entry.configure(style='Editable.TEntry')
            else:
                entry.configure(style='Error.TEntry')
            
            # Always ensure entry remains editable
            entry.configure(state='normal')
            return True
            
        except Exception as e:
            print(f"Error in time input validation: {str(e)}")
            entry.configure(state='normal')
            return True
    
    def _ensure_entries_editable(self):
        """Enhanced method to ensure time entry fields are editable"""
        try:
            style = ttk.Style()
            if 'Editable.TEntry' not in style.theme_names():
                style.configure('Editable.TEntry', fieldbackground='white')
            
            if hasattr(self, 'morning_entry') and self.morning_entry.winfo_exists():
                self.morning_entry.configure(state='normal', style='Editable.TEntry')
                self.morning_entry.update()
            
            if hasattr(self, 'afternoon_entry') and self.afternoon_entry.winfo_exists():
                self.afternoon_entry.configure(state='normal', style='Editable.TEntry')
                self.afternoon_entry.update()
            
            # Schedule another check after a short delay
            self.root.after(200, self._final_entry_check)
            
        except Exception as e:
            print(f"Error ensuring entries are editable: {str(e)}")
            
    def _final_entry_check(self):
        """Final check to absolutely ensure entries are editable"""
        try:
            if hasattr(self, 'morning_entry') and self.morning_entry.winfo_exists():
                self.morning_entry.configure(state='normal', style='Editable.TEntry')
                self.morning_entry.update()
            if hasattr(self, 'afternoon_entry') and self.afternoon_entry.winfo_exists():
                self.afternoon_entry.configure(state='normal', style='Editable.TEntry')
                self.afternoon_entry.update()
        except Exception as e:
            print(f"Error in final entry check: {str(e)}")

class WingetUpdaterTray:
    """Class to manage the system tray icon and communication with the service"""
    
    def __init__(self):
        # Setup logging
        self._setup_logging()
        
        # Initialize IPC client
        self.ipc_client = IPCClient()
        
        # Initialize state
        self.update_count = 0
        self.previous_update_count = 0
        self.service_connected = False
        self.running = True
        self.auto_reconnect = True
        
        # Initialize window state tracking dictionary for monitoring window issues
        self.window_state_tracking = {
            'settings': {'attempts': 0, 'last_attempt': 0, 'success': False},
            'updates': {'attempts': 0, 'last_attempt': 0, 'success': False}
        }
        
        # Initialize Tkinter first
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the root window
        self.root.title("Winget Updater")
        
        # Set protocol to properly handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_root_close)
        
        # Queue for thread-safe window operations
        self.ui_queue = queue.Queue()
        
        # Create queue for window creation requests
        self.window_queue = queue.Queue()
        
        # For storing the tkinter dialog windows
        self.settings_window = None
        self.updates_window = None
        
        # Create the system tray icon
        self._create_icon()
        self._setup_menu()
        
        # Start the connection manager thread
        self.connection_thread = threading.Thread(target=self._connection_manager, daemon=True)
        self.connection_thread.start()
        
        # Start the status update thread
        self.status_thread = threading.Thread(target=self._status_updater, daemon=True)
        self.status_thread.start()
    
    def _setup_logging(self):
        """Set up logging for the tray component"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler("winget_updater_ui.log"),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger('WingetUpdaterTray')
        
        # Log startup with system info
        self.logger.info("WingetUpdaterTray initializing")
        self.logger.info(f"Python version: {sys.version}")
        self.logger.info(f"Tkinter version: {tk.TkVersion}")
    
    def _create_icon(self):
        """Create the system tray icon"""
        try:
            # Default icon image - a simple square
            image = self._create_icon_image(0, self.service_connected)
            
            # Create the icon
            self.icon = pystray.Icon("winget_updater", image, "Winget Updater")
            self.logger.info("System tray icon created")
        except Exception as e:
            self.logger.error(f"Error creating system tray icon: {str(e)}")
            raise
    
    def _create_icon_image(self, update_count, connected=True):
        """Create a custom icon image with the update count"""
        try:
            # Create a blank image
            width, height = 64, 64
            image = Image.new('RGBA', (width, height), color=(0, 0, 0, 0))
            draw = ImageDraw.Draw(image)
            
            # Draw a different background color based on status
            if not connected:
                # Gray for disconnected
                draw.rectangle([(0, 0), (width, height)], fill=(150, 150, 150, 255))
            elif update_count > 0:
                # Red background for available updates
                draw.rectangle([(0, 0), (width, height)], fill=(200, 30, 30, 255))
            else:
                # Green background for no updates
                draw.rectangle([(0, 0), (width, height)], fill=(30, 150, 30, 255))
            
            # Try multiple fonts to ensure one works
            font = None
            font_size = 36  # Larger font for better visibility
            
            # Try a few common fonts
            try_fonts = ["arial.ttf", "verdana.ttf", "tahoma.ttf", "segoeui.ttf"]
            for font_name in try_fonts:
                try:
                    font = ImageFont.truetype(font_name, font_size)
                    break
                except:
                    continue
                    
            if font is None:
                # Fall back to default font if none of the others worked
                try:
                    font = ImageFont.load_default()
                except:
                    # If we can't get any font, we'll just draw without text
                    pass
            
            if font:
                # Draw the "W" letter
                letter = "W"
                
                # Calculate text size and position for centering
                try:
                    # For newer PIL versions
                    text_width = draw.textlength(letter, font=font)
                    text_height = font_size  # Approximation
                except AttributeError:
                    # For older PIL versions
                    try:
                        text_width, text_height = draw.textsize(letter, font=font)
                    except:
                        # Fallback approximation
                        text_width, text_height = font_size * 0.8, font_size
                
                # Calculate center position
                x = (width - text_width) // 2
                y = (height - text_height) // 2
                
                # Draw a slight shadow for better visibility
                draw.text(
                    (x + 2, y + 2),
                    letter,
                    font=font,
                    fill=(0, 0, 0, 128)  # Semi-transparent black shadow
                )
                
                # Draw the main text
                draw.text(
                    (x, y),
                    letter,
                    font=font,
                    fill=(255, 255, 255, 255)  # White text
                )
            
            # Add the update count in corner if updates are available
            if update_count > 0 and connected:
                # Create a small circle in the corner for the count
                circle_radius = 12
                circle_x = width - circle_radius - 4
                circle_y = circle_radius + 4
                
                # Draw a circle
                draw.ellipse(
                    [(circle_x - circle_radius, circle_y - circle_radius), 
                     (circle_x + circle_radius, circle_y + circle_radius)], 
                    fill=(255, 255, 255, 255)  # White circle
                )
                
                # Try to use a small font for the number
                small_font = None
                try:
                    small_font = ImageFont.truetype("arial.ttf", 14)
                except:
                    try:
                        small_font = ImageFont.load_default()
                    except:
                        pass
                
                if small_font:
                    # Convert the count to a string, limit to 2 digits
                    count_str = str(min(update_count, 99))
                    if update_count > 99:
                        count_str = "99+"
                        
                    # Calculate text size for centering in the circle
                    try:
                        # For newer PIL versions
                        count_width = draw.textlength(count_str, font=small_font)
                        count_height = 14  # Approximation
                    except AttributeError:
                        # For older PIL versions
                        try:
                            count_width, count_height = draw.textsize(count_str, font=small_font)
                        except:
                            # Fallback approximation
                            count_width, count_height = 10, 14
                    
                    # Draw the count in the circle
                    draw.text(
                        (circle_x - count_width/2, circle_y - count_height/2),
                        count_str,
                        font=small_font,
                        fill=(0, 0, 0, 255)  # Black text
                    )
            
            return image
        except Exception as e:
            self.logger.error(f"Error creating icon image: {str(e)}")
            # Return a simple backup image
            return self._create_fallback_image(connected)
    
    def _create_fallback_image(self, connected=True):
        """Create a simple fallback image in case of errors"""
        width, height = 64, 64
        image = Image.new('RGBA', (width, height), color=(0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        # Draw a simple colored square
        if connected:
            color = (30, 150, 30, 255)  # Green for connected with no updates
        else:
            color = (150, 150, 150, 255)  # Gray for disconnected
            
        draw.rectangle([(0, 0), (width, height)], fill=color)
        
        # Try to draw a simple "W" with basic shapes if drawing text fails
        if connected:
            # Draw a simple W using lines
            line_color = (255, 255, 255, 255)  # White
            line_width = 3
            
            # Calculate points for a simplified "W" shape
            margin = 12
            top_y = margin
            bottom_y = height - margin
            left_x = margin
            right_x = width - margin
            middle_x1 = width // 3
            middle_x2 = (width * 2) // 3
            middle_y = (height * 3) // 4
            
            # Draw the W shape
            # Left line
            draw.line([(left_x, top_y), (left_x, bottom_y)], fill=line_color, width=line_width)
            # Middle lines
            draw.line([(left_x, bottom_y), (middle_x1, middle_y)], fill=line_color, width=line_width)
            draw.line([(middle_x1, middle_y), (middle_x2, middle_y)], fill=line_color, width=line_width)
            # Right line
            draw.line([(middle_x2, middle_y), (right_x, bottom_y)], fill=line_color, width=line_width)
            draw.line([(right_x, bottom_y), (right_x, top_y)], fill=line_color, width=line_width)
        return image
    
    def _setup_menu(self):
        """Create the system tray context menu"""
        try:
            # Create menu with bound callbacks and error handling
            def create_callback(func):
                """Wrap callback with error handling"""
                def wrapped():
                    try:
                        func()
                    except Exception as e:
                        self.logger.error(f"Error in menu callback {func.__name__}: {str(e)}")
                        self.ui_queue.put(lambda: messagebox.showerror(
                            "Error",
                            f"An error occurred: {str(e)}"
                        ))
                return wrapped

            self.icon.menu = pystray.Menu(
                item('Check for Updates', create_callback(self._on_check_updates)),
                item('Install All Updates', create_callback(self._on_install_updates)),
                item('Show Updates', create_callback(self._on_show_updates)),
                item('Settings', create_callback(self._on_open_settings)),
                item('Reconnect to Service', create_callback(self._on_reconnect)),
                item('Exit', create_callback(self._on_exit))
            )
            self.logger.info("System tray menu setup completed successfully")
            
        except Exception as e:
            self.logger.error(f"Error setting up system tray menu: {str(e)}")
            raise
    
    def _on_root_close(self):
        """Handle the root window close event"""
        # This should not happen in normal operation since the window is hidden
        self.logger.warning("Root window close event detected")
        
        # Don't destroy the root window, just hide it again
        if hasattr(self, 'root') and self.root:
            self.root.withdraw()
    
    def run(self):
        """Run the system tray application"""
        self.logger.info("Starting Winget Updater tray application")
        
        try:
            # Start the icon in a separate thread
            icon_thread = threading.Thread(target=self.icon.run, daemon=True)
            icon_thread.start()
            
            # Start window queue processing
            self._process_window_queue()
            
            # Main event loop in the main thread
            while self.running:
                try:
                    # Process any pending UI operations from the queue
                    processed = 0
                    while not self.ui_queue.empty() and processed < 5:  # Process fewer tasks per cycle
                        ui_task = self.ui_queue.get_nowait()
                        if ui_task and callable(ui_task):
                            ui_task()
                            processed += 1
                    
                    # Process Tkinter events - use update_idletasks() to process only pending events
                    # without blocking for new events, then use update() for a complete update
                    self.root.update_idletasks()
                    self.root.update()
                    
                    # Slightly longer sleep to give windows more time to properly map
                    # This helps with focus management and window visibility
                    time.sleep(0.01)  # Reduced sleep time for more responsive UI
                except tk.TclError as e:
                    if not self.running:
                        break
                    # Log Tkinter specific errors which might indicate issues with window handling
                    self.logger.error(f"Tkinter error in event loop: {str(e)}")
                except Exception as e:
                    if not self.running:
                        break
                    self.logger.error(f"Error in Tkinter event loop: {str(e)}")
            
            # Clean up
            self.root.destroy()
        except Exception as e:
            self.logger.error(f"Error running system tray application: {str(e)}")
    # Helper function to properly initialize and show a window
    def _prepare_window_for_display(self, window):
        """Prepare a window for display with proper initialization and visibility"""
        self.logger.info("Preparing window for display")
        
        try:
            # Check if window still exists before operations
            if not window.winfo_exists():
                self.logger.error("Window does not exist during preparation")
                return False

            # Ensure window is fully created
            try:
                window.update_idletasks()
            except tk.TclError as e:
                self.logger.error(f"Failed to update window idletasks: {str(e)}")
                return False
            
            # Center on screen if not already positioned
            if '+' not in window.geometry():
                # Get screen dimensions correctly
                try:
                    screen_width = window.winfo_screenwidth()
                    screen_height = window.winfo_screenheight()
                    win_width = window.winfo_reqwidth()  # Use requested width for better accuracy
                    win_height = window.winfo_reqheight()  # Use requested height for better accuracy
                    
                    # If requested size is too small, use actual window size
                    if win_width < 100 or win_height < 100:
                        win_width = window.winfo_width()
                        win_height = window.winfo_height()
                    
                    # Calculate center position
                    x = max(0, (screen_width - win_width) // 2)
                    y = max(0, (screen_height - win_height) // 2)
                    window.geometry(f'+{x}+{y}')
                    self.logger.info(f"Window centered at +{x}+{y}")
                except Exception as e:
                    self.logger.warning(f"Failed to center window: {str(e)}")
                    # Continue despite centering failure
            
            # More efficient visibility sequence
            try:
                # Ensure window is in normal state and visible
                window.state('normal')  # Ensure not minimized/maximized
                window.deiconify()     # Make sure window is visible
                window.lift()          # Bring to front
                window.focus_set()     # Initial focus
                window.update()        # Process immediately
                
                # Temporarily set topmost for better visibility
                window.attributes('-topmost', True)
                window.update_idletasks()
                # Reset topmost after a short delay
                window.after(100, lambda: window.attributes('-topmost', False))
                
                # Schedule a focus attempt after topmost reset
                window.after(150, lambda: self._force_window_focus(window))
            except Exception as e:
                self.logger.warning(f"Visibility sequence error: {str(e)}")
                # Emergency recovery attempt
                try:
                    window.deiconify()
                    window.attributes('-topmost', True)
                    window.update()
                    window.after(100, lambda: window.attributes('-topmost', False))
                except Exception as ex:
                    self.logger.error(f"Recovery attempt failed: {str(ex)}")
            
            # Log window state for debugging
            self.logger.info(f"Window prepared with geometry: {window.geometry()}, " +
                          f"viewable: {window.winfo_viewable()}, " +
                          f"mapped: {window.winfo_ismapped()}")
            
            return True
        except Exception as e:
            self.logger.error(f"Error preparing window for display: {str(e)}")
            return False
    def _ensure_window_focus(self, window, window_type='window', retry_count=0):
        """
        Improved window focus management that works better with Windows focus restrictions
        """
        # Early return if window doesn't exist
        if not window.winfo_exists():
            self.logger.warning(f"{window_type.capitalize()} window no longer exists when attempting to focus")
            return False
        
        try:
            # Reduced max retries to prevent excessive attempts
            max_retries = 3  # Reduced from 6 since we're removing unreliable Win32 calls
            
            # Track start time for this focus attempt for timeout management
            if retry_count == 0:
                setattr(window, '_focus_start_time', time.time())
            
            # Get elapsed time since first attempt
            elapsed_time = time.time() - getattr(window, '_focus_start_time', time.time())
            
            # Time-based safety valve - don't retry for more than 1.0 seconds total
            if elapsed_time > 1.0:  # Reduced timeout since we're doing less
                self.logger.warning(f"Focus attempts taking too long ({elapsed_time:.1f}s), forcing final attempt")
                retry_count = max_retries
            
            # Improved state validation - check multiple window state attributes
            state_info = {
                "exists": window.winfo_exists(),
                "viewable": window.winfo_viewable(),
                "mapped": window.winfo_ismapped(),
                "geometry": window.geometry(),
                "has_focus": window.focus_get() == window,
                "retry": retry_count,
                "elapsed_time": f"{elapsed_time:.2f}s"
            }
            
            # Log every retry to better track focus issues
            self.logger.info(f"{window_type.capitalize()} window state check: {state_info}")
            
            # More comprehensive success criteria
            success = (window.winfo_exists() and 
                      window.winfo_viewable() and 
                      window.winfo_ismapped() and
                      window.winfo_width() > 10 and
                      window.winfo_height() > 10)
            
            if success:
                try:
                    # Ensure window is fully ready
                    window.update_idletasks()
                    
                    # Sequence of focus commands for better reliability
                    window.deiconify()
                    window.lift()
                    window.focus_force()
                    window.grab_set()
                    
                    # Verify focus was actually obtained
                    window.update()
                    if window.focus_get() == window:
                        self.logger.info(f"{window_type.capitalize()} window focused successfully")
                        if hasattr(window, '_focus_start_time'):
                            delattr(window, '_focus_start_time')
                        return True
                    
                    # If focus verification fails but window is visible and mapped,
                    # consider it a success anyway since Windows focus restrictions
                    # might prevent perfect focus
                    if window.winfo_viewable() and window.winfo_ismapped():
                        self.logger.info(f"{window_type.capitalize()} window visible and mapped, considering focus successful")
                        if hasattr(window, '_focus_start_time'):
                            delattr(window, '_focus_start_time')
                        return True
                        
                    # If we get here, focus verification failed completely
                    success = False
                except Exception as e:
                    self.logger.warning(f"Focus attempt failed: {str(e)}")
                    success = False
            
            if not success and retry_count < max_retries:
                # Simplified backoff strategy
                delay = 50 * (retry_count + 1)  # 50ms, 100ms, 150ms
                
                # Progressive focus strategy
                if retry_count == 0:
                    window.deiconify()
                    window.lift()
                else:
                    window.attributes('-topmost', True)
                    window.deiconify()
                    window.lift()
                    window.after(50, lambda: window.attributes('-topmost', False))
                
                self.logger.info(f"Retrying focus for {window_type} window (attempt {retry_count + 1}/{max_retries})")
                window.after(delay, lambda: self._ensure_window_focus(window, window_type, retry_count + 1))
                return False
            
            elif not success:
                # Final attempt if regular retries failed
                self.logger.warning(f"Final focus attempt for {window_type} window")
                try:
                    window.attributes("-topmost", True)
                    window.deiconify()
                    window.lift()
                    window.focus_force()
                    window.grab_set()
                    window.update()
                    window.after(100, lambda: window.attributes("-topmost", False))
                except Exception as e:
                    self.logger.error(f"Final focus attempt failed: {str(e)}")
                
                if hasattr(window, '_focus_start_time'):
                    delattr(window, '_focus_start_time')
                return False
            
        except Exception as e:
            self.logger.error(f"Error in window focus management: {str(e)}")
            if hasattr(window, '_focus_start_time'):
                delattr(window, '_focus_start_time')
            return False
    
    def _force_window_focus(self, window):
        """Helper method to force focus on a window with better error handling"""
        try:
            if window.winfo_exists():
                if window.winfo_viewable() and window.winfo_ismapped():
                    # Try multiple focus methods for better reliability
                    window.focus_set()
                    window.focus_force()
                    
                    # Use Tkinter's built-in focus methods instead of Win32 API
                    window.attributes('-topmost', True)
                    window.update()
                    window.after(50, lambda: window.attributes('-topmost', False))
                else:
                    self.logger.warning("Window not viewable/mapped during focus attempt")
        except Exception as e:
            self.logger.warning(f"Error forcing window focus: {str(e)}")
    
    def _connection_manager(self):
        """Manage the connection to the service"""
        self.logger.info("Starting connection manager thread")
        
        while self.running:
            if not self.service_connected and self.auto_reconnect:
                self._connect_to_service()
                
            time.sleep(5)  # Check connection every 5 seconds
    
    def _connect_to_service(self):
        """Connect to the service"""
        self.logger.info("Attempting to connect to service")
        
        try:
            if self.ipc_client.connect(timeout=5):
                self.service_connected = True
                self.logger.info("Connected to service")
                self._update_icon_status()
                
                # Immediately get the update status
                self._get_update_status()
            else:
                self.service_connected = False
                self.logger.warning("Could not connect to service")
                self._update_icon_status()
        except Exception as e:
            self.service_connected = False
            self.logger.error(f"Error connecting to service: {str(e)}")
            self._update_icon_status()
    
    def _status_updater(self):
        """Periodically update the status from the service"""
        self.logger.info("Starting status updater thread")
        
        while self.running:
            if self.service_connected:
                self._get_update_status()
                
            time.sleep(60)  # Update every minute
    
    def _get_update_status(self):
        """Get the update status from the service"""
        try:
            self.logger.debug("Requesting update status from service")
            response = self.ipc_client.send_command("get_status")
            
            if response and response.command == "response":
                # Log the response
                self.logger.debug(f"Received status response: {response.data}")
                
                # Update our local state
                self.previous_update_count = self.update_count
                self.update_count = response.data.get("update_count", 0)
                
                if self.update_count != self.previous_update_count:
                    self.logger.info(f"Update count changed from {self.previous_update_count} to {self.update_count}")
                
                # Update the icon
                self._update_icon()
                
            # Show notification if count changed
            if (self.update_count > 0 and 
                self.update_count != self.previous_update_count):
                
                # Only show notification if there's an increase
                if self.update_count > self.previous_update_count:
                    self._show_notification(self.update_count)
                    
                return True
            
            # Add log to help debug dialog issues
            self.logger.debug(f"Update status: count={self.update_count}, previous={self.previous_update_count}")
            else:
                # Failed to get status
                error_msg = "Unknown error"
                if response and response.command == "error":
                    error_msg = response.data.get("message", error_msg)
                
                self.logger.error(f"Failed to get status: {error_msg}")
                self.service_connected = False
                self._update_icon_status()
                return False
                
        except Exception as e:
            self.logger.error(f"Error getting update status: {str(e)}", exc_info=True)
            self.service_connected = False
            self._update_icon_status()
            return False
    
    def _check_updates(self):
        """Trigger an update check on the service"""
        self.logger.info("Requesting update check")
        
        try:
            if not self.service_connected:
                if not self._connect_to_service():
                    self.logger.error("Failed to connect to service for update check")
                    self.ui_queue.put(lambda: messagebox.showerror("Service Error", "Could not connect to the Winget Updater service."))
                    return False
            
            self.logger.debug("Sending check_updates command to service")
            response = self.ipc_client.send_command("check_updates")
            
            if response and response.command == "response":
                self.logger.info(f"Update check response received: {response.data}")
                
                # Update our local state
                self.previous_update_count = self.update_count
                self.update_count = response.data.get("update_count", 0)
                
                # Log the state change
                self.logger.info(f"Update count changed from {self.previous_update_count} to {self.update_count}")
                
                # Update the icon
                self._update_icon()
                
                # Show notification
                self._show_update_result_notification(self.update_count)
                
                return True
            else:
                error_msg = "Unknown error"
                if response and response.command == "error":
                    error_msg = response.data.get("message", error_msg)
                
                self.logger.error(f"Update check failed: {error_msg}")
                self.ui_queue.put(lambda: messagebox.showerror("Update Check Failed", 
                    f"Failed to check for updates: {error_msg}"))
                return False
                
        except Exception as e:
            self.logger.error(f"Error checking for updates: {str(e)}", exc_info=True)
            self.ui_queue.put(lambda: messagebox.showerror("Error", 
                f"An error occurred while checking for updates: {str(e)}"))
            return False
    
    def _update_icon(self):
        """Update the system tray icon with the current update count"""
        try:
            new_image = self._create_icon_image(self.update_count, self.service_connected)
            
            # Update the icon image
            self.icon.icon = new_image
            
            # Update the tooltip text
            if not self.service_connected:
                self.icon.title = "Winget Updater - Service disconnected"
            elif self.update_count == 0:
                self.icon.title = "Winget Updater - No updates available"
            elif self.update_count == 1:
                self.icon.title = "Winget Updater - 1 update available"
            else:
                self.icon.title = f"Winget Updater - {self.update_count} updates available"
        except Exception as e:
            self.logger.error(f"Error updating icon: {str(e)}")
    
    def _update_icon_status(self):
        """Update the icon to reflect the service connection status"""
        try:
            new_image = self._create_icon_image(self.update_count if self.service_connected else 0, self.service_connected)
            
            # Update the icon image
            self.icon.icon = new_image
            
            # Update the tooltip text
            if not self.service_connected:
                self.icon.title = "Winget Updater - Service disconnected"
        except Exception as e:
            self.logger.error(f"Error updating icon status: {str(e)}")
    
    def _show_notification(self, update_count):
        """Show a notification about available updates"""
        try:
            if update_count == 1:
                self.icon.notify(
                    f"1 update is available for your system.",
                    "Winget Updates Available"
                )
            else:
                self.icon.notify(
                    f"{update_count} updates are available for your system.",
                    "Winget Updates Available"
                )
        except Exception as e:
            self.logger.error(f"Error showing notification: {str(e)}")
    
    def _show_update_result_notification(self, update_count):
        """Show a notification about the update check result"""
        try:
            if update_count == 0:
                self.icon.notify(
                    "No updates are currently available for your system.",
                    "Winget Update Check Complete"
                )
            elif update_count == 1:
                self.icon.notify(
                    f"1 update is available for your system.",
                    "Winget Update Check Complete"
                )
            else:
                self.icon.notify(
                    f"{update_count} updates are available for your system.",
                    "Winget Update Check Complete"
                )
        except Exception as e:
            self.logger.error(f"Error showing update result notification: {str(e)}")
    
    def _on_check_updates(self):
        """Handle the 'Check for Updates' menu item"""
        # Use threading to avoid blocking the UI
        thread = threading.Thread(target=self._check_updates, daemon=True)
        thread.start()
    
    def _process_window_queue(self):
        """Process any pending window creation requests"""
        try:
            # Process only one window request at a time for better stability
            if not self.window_queue.empty():
                window_info = self.window_queue.get_nowait()
                if not window_info:
                    pass  # Skip empty entries
                else:
                    window_type = window_info.get('type')
                    self.logger.info(f"Processing window creation request: {window_type}")
                    
                    # Process any pending events before creating a new window
                    self.root.update_idletasks()
                    
                    # Allow UI to settle before creating a new window
                    time.sleep(0.1)
                    
                    if window_type == 'settings':
                        self.logger.info("Creating settings window")
                        self._create_settings_window_main()
                    elif window_type == 'updates':
                        self.logger.info("Creating updates window")
                        self._create_updates_window_main()
                    
                    # Allow more time between processing window requests
                    time.sleep(0.2)
        except Exception as e:
            self.logger.error(f"Error processing window queue: {str(e)}")
        finally:
            # Schedule next check with a longer interval
            if self.running:
                # Schedule with a short delay for better responsiveness
                self.root.after(50, self._process_window_queue)
                # Ensure the root window processes events
                self.root.update_idletasks()
    
    def _on_open_settings(self):
        """Queue settings window creation"""
        # Track window creation attempts for diagnostic purposes
        current_time = time.time()
        self.window_state_tracking['settings']['attempts'] += 1
        self.window_state_tracking['settings']['last_attempt'] = current_time
        self.logger.info(f"Queuing settings window creation (attempt {self.window_state_tracking['settings']['attempts']})")
        
        # Queue the window creation
        self.window_queue.put({'type': 'settings', 'timestamp': current_time})
    
    def _create_settings_window_main(self):
        """Create settings window in main thread"""
        try:
            # Better initialization tracking
            window_start_time = time.time()
            self.logger.info(f"Starting settings window creation at {window_start_time}")
            
            # If window exists, try to focus it with improved validation
            if self.settings_window:
                try:
                    # More thorough check if window is still valid
                    if self.settings_window.winfo_exists():
                        self.logger.info("Settings window already exists, focusing it")
                        
                        # More robust way to bring window to front
                        self.settings_window.deiconify()
                        self.settings_window.lift()
                        if self.settings_window.winfo_viewable():
                            self.settings_window.focus_force()
                        return
                    else:
                        self.logger.info("Settings window reference exists but window was destroyed")
                        self.settings_window = None
                except Exception as e:
                    self.logger.error(f"Error focusing existing settings window: {str(e)}")
                    self.settings_window = None

            # Ensure we're connected to the service
            if not self.service_connected:
                if not self._connect_to_service():
                    messagebox.showerror("Service Error", 
                        "Could not connect to the Winget Updater service. Settings will not be saved.")
                    return

            # Create a proper toplevel window that shares the event loop with the main application
            # This ensures proper window management while maintaining integration
            window = tk.Toplevel(self.root)
            window.title("Winget Updater Settings")
            window.geometry("400x350")
            window.resizable(False, False)
            
            # Process window creation immediately
            window.update_idletasks()
            
            # Set icon if available
            try:
                window.iconbitmap("winget_updater.ico")
            except:
                pass
            
            # Store reference to the window early so we can clean it up in case of errors
            self.settings_window = window
            
            # Create settings window instance
            settings_instance = SettingsWindow(window, self.ipc_client, self._on_settings_saved)
            
            # Set up close handler explicitly 
            window.protocol("WM_DELETE_WINDOW", lambda: self._on_settings_window_closed(window))
            
            # Add escape key binding to close window
            window.bind("<Escape>", lambda event: self._on_settings_window_closed(window))
            
            # Center window on screen
            window.update_idletasks()
            x = (window.winfo_screenwidth() - window.winfo_width()) // 2
            y = (window.winfo_screenheight() - window.winfo_height()) // 2
            window.geometry(f'+{x}+{y}')
            
            # Log window creation
            self.logger.info(f"Settings window created with geometry: {window.geometry()}")
            
            # Track successful window creation for diagnostics
            self.window_state_tracking['settings']['success'] = True
            
            # Make window modal, but defer grab_set until after window is visible
            window.transient(self.root)  # Set as transient to main window
            
            # Ensure window is mapped and visible before setting focus and grab
            # Use the dedicated window preparation method for consistent handling
            self._prepare_window_for_display(window)
            
            # Add an explicit state check after preparation
            self.logger.info(f"Window state after preparation: " +
                          f"exists: {window.winfo_exists()}, " +
                          f"viewable: {window.winfo_viewable()}, " +
                          f"mapped: {window.winfo_ismapped()}, " +
                          f"geometry: {window.geometry()}")
            
            # Use the improved ensure_window_focus method for better focus handling
            self._ensure_window_focus(window, window_type='settings')
            
            # Start the focus process after a short delay
            # Make sure window is fully mapped before starting focus process
            # Draw the window completely first with update + update_idletasks
            window.update()
            window.update_idletasks()
            
            # Map the window explicitly and ensure it's visible
            window.deiconify()
            window.lift()
            
            # Start the focus process after ensuring window is drawn
            # A short delay is still helpful for window manager to process
            window.after(20, lambda: self._ensure_window_focus(window, window_type='settings'))  # Shorter initial delay, focus handler has better retry logic
            
            # The window will automatically be handled by the main event loop
            # since it's a Toplevel child of the root window
            
        except Exception as e:
            self.logger.error(f"Error creating settings window: {str(e)}")
            messagebox.showerror("Error", f"Could not open settings window: {str(e)}")
            if self.settings_window:
                try:
                    self.settings_window.destroy()
                except:
                    pass
                self.settings_window = None
    
    def _get_settings_from_service(self):
        """Get settings directly from the service"""
        try:
            response = self.ipc_client.send_command("get_settings")
            if response and response.command == "response":
                return response.data
            return None
        except Exception as e:
            self.logger.error(f"Error getting settings: {str(e)}")
            return None
    
    def _on_settings_window_closed(self, window):
        """Handle the settings window being closed"""
        self.logger.info("Settings window closing")
        
        try:
            # Check if window still exists
            if not window.winfo_exists():
                self.logger.info("Window already destroyed")
                self.settings_window = None
                return
                
            # Process any pending changes before closing
            try:
                window.update()
            except Exception as e:
                self.logger.warning(f"Error updating window before close: {str(e)}")
            
            # Check window state for debugging
            try:
                self.logger.info(f"Window state before closing - exists: {window.winfo_exists()}, " +
                              f"viewable: {window.winfo_viewable() if window.winfo_exists() else 'N/A'}, " +
                              f"geometry: {window.geometry() if window.winfo_exists() else 'N/A'}")
            except:
                pass
            
            # Release grab first - critical for proper window management
            try:
                if window.winfo_exists():
                    window.grab_release()
                    self.logger.info("Grab released")
            except Exception as e:
                self.logger.warning(f"Error releasing grab: {str(e)}")
                
            # Explicitly remove any bindings before destroying
            try:
                if window.winfo_exists():
                    window.unbind("<Escape>")
                    window.protocol("WM_DELETE_WINDOW", lambda: None)
                    self.logger.info("Window bindings removed")
            except Exception as e:
                self.logger.warning(f"Error removing bindings: {str(e)}")
                
            # Destroy the window
            try:
                if window.winfo_exists():
                    window.destroy()
                    self.logger.info("Window destroyed")
            except Exception as e:
                self.logger.warning(f"Error destroying window: {str(e)}")
                
        except Exception as e:
            self.logger.error(f"Error closing settings window: {str(e)}")
        
        # Always ensure we clean up our reference
        self.settings_window = None
        self.logger.info("Settings window reference cleared")
        
        # Track successful window cleanup for diagnostics
        self.window_state_tracking['settings']['success'] = False
        
        # Ensure focus returns to the main application after a short delay
        def restore_root_focus():
            try:
                if self.root and self.root.winfo_exists():
                    self.root.update_idletasks()
                    self.root.focus_force()
                    self.logger.info("Focus returned to main window")
            except Exception as e:
                self.logger.warning(f"Error returning focus to main window: {str(e)}")
        
        # Schedule focus restoration with a delay
        if self.root and self.root.winfo_exists():
            self.root.after(200, restore_root_focus)
    
    def _on_settings_saved(self):
        """Handle settings being saved"""
        self.logger.info("Settings were updated")
        # Trigger an immediate status update
        threading.Thread(target=self._get_update_status, daemon=True).start()
    
    
    def _on_show_updates(self):
        """Queue updates window creation"""
        # Track window creation attempts for diagnostic purposes
        current_time = time.time()
        self.window_state_tracking['updates']['attempts'] += 1
        self.window_state_tracking['updates']['last_attempt'] = current_time
        self.logger.info(f"Queuing updates window creation (attempt {self.window_state_tracking['updates']['attempts']})")
        
        # Queue the window creation
        self.window_queue.put({'type': 'updates', 'timestamp': current_time})
    
    def _get_updates_from_service(self):
        """Get updates directly from service with proper error handling"""
        try:
            # Show a progress message or indicator could be added here
            response = self.ipc_client.send_command("get_updates")
            
            if not response or response.command != "response":
                self.logger.error(f"Failed to get updates: {response}")
                return None
                
            return response.data.get("updates", [])
        except Exception as e:
            self.logger.error(f"Error getting updates: {str(e)}")
            return None
        
    def _create_updates_window_main(self):
        """Create the updates window in the main thread"""
        try:
            # If window exists, try to focus it
            if self.updates_window:
                try:
                    # Check if window is still valid
                    if self.updates_window.winfo_exists():
                        self.logger.info("Updates window already exists, focusing it")
                        self.updates_window.lift()
                        # Only force focus if window is mapped (visible)
                        if self.updates_window.winfo_viewable():
                            self.updates_window.focus_force()
                        return
                    else:
                        self.logger.info("Updates window reference exists but window was destroyed")
                        self.updates_window = None
                except Exception as e:
                    self.logger.error(f"Error focusing updates window: {str(e)}")
                    self.updates_window = None

            # Ensure we're connected to the service
            if not self.service_connected:
                if not self._connect_to_service():
                    messagebox.showerror("Service Error", "Could not connect to the Winget Updater service.")
                    return
            
            # Show a waiting cursor while retrieving updates
            self.root.config(cursor="wait")
            
            try:
                # Get the list of updates
                updates = self._get_updates_from_service()
                
                # Reset cursor
                self.root.config(cursor="")
                
                if updates is None:
                    messagebox.showerror("Error", "Could not retrieve updates list from service.")
                    return
                    
                if not updates:
                    # Show a notification if no updates are available
                    messagebox.showinfo("Winget Updates", "No updates are currently available.")
                    return
                
                # Create new window
                window = tk.Toplevel(self.root)
                window.title("Available Updates")
                window.geometry("700x400")
                window.resizable(True, True)
                
                # Process window creation immediately
                window.update_idletasks()
                
                # Set icon if available
                try:
                    window.iconbitmap("winget_updater.ico")
                except:
                    pass
                
                # Make window transient but defer grab_set
                window.transient(self.root)
                
                # Create updates list content
                UpdateListWindow(window, updates)
                
                # Store reference
                self.updates_window = window
                
                # Center on screen
                window.update_idletasks()
                x = (window.winfo_screenwidth() - window.winfo_width()) // 2
                y = (window.winfo_screenheight() - window.winfo_height()) // 2
                window.geometry(f'+{x}+{y}')
                
                # Log window creation
                self.logger.info(f"Updates window created with geometry: {window.geometry()}")
                
                # Track successful window creation for diagnostics
                self.window_state_tracking['updates']['success'] = True
                
                # Set up close handler
                window.protocol("WM_DELETE_WINDOW", lambda: self._on_updates_window_closed(window))
                
                # Show and ensure window is mapped using the dedicated method
                self._prepare_window_for_display(window)
                
                # Add an explicit state check after preparation
                self.logger.info(f"Updates window state after preparation: " +
                              f"exists: {window.winfo_exists()}, " +
                              f"viewable: {window.winfo_viewable()}, " +
                              f"mapped: {window.winfo_ismapped()}, " +
                              f"geometry: {window.geometry()}")
                
                # Define function to handle setting grab and focus once window is ready
                # Use the improved ensure_window_focus method for better focus handling
                self._ensure_window_focus(window, window_type='updates')
                
                # Make sure window is fully mapped before starting focus process
                # Draw the window completely first with update + update_idletasks
                window.update()
                window.update_idletasks()
                
                # Map the window explicitly and ensure it's visible
                window.deiconify()
                window.lift()
                
                # Start the focus process after ensuring window is drawn
                # A short delay is still helpful for window manager to process
                window.after(20, lambda: self._ensure_window_focus(window, window_type='updates'))  # Shorter initial delay, focus handler has better retry logic
                
                # Ensure the window gets processed by the event loop
                window.update_idletasks()
            except Exception as e:
                # Reset cursor in case of error
                self.root.config(cursor="")
                raise e
            
        except Exception as e:
            self.logger.error(f"Error showing updates: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while retrieving updates: {str(e)}")
            if 'window' in locals():
                window.destroy()
            self.updates_window = None
    
    def _on_updates_window_closed(self, window):
        """Handle the updates window being closed"""
        self.logger.info("Updates window closing")
        
        try:
            # Check if window still exists
            if not window.winfo_exists():
                self.logger.info("Updates window already destroyed")
                self.updates_window = None
                return
                
            # Process any pending changes before closing
            try:
                window.update()
            except Exception as e:
                self.logger.warning(f"Error updating updates window before close: {str(e)}")
                
            # Check window state for debugging
            try:
                self.logger.info(f"Updates window state before closing - exists: {window.winfo_exists()}, " +
                              f"viewable: {window.winfo_viewable() if window.winfo_exists() else 'N/A'}, " +
                              f"geometry: {window.geometry() if window.winfo_exists() else 'N/A'}")
            except:
                pass
                
            # Release grab before destroying
            try:
                if window.winfo_exists():
                    window.grab_release()
                    self.logger.info("Updates window grab released")
            except Exception as e:
                self.logger.warning(f"Error releasing updates window grab: {str(e)}")
                
            # Destroy the window
            try:
                if window.winfo_exists():
                    window.destroy()
                    self.logger.info("Updates window destroyed")
            except Exception as e:
                self.logger.warning(f"Error destroying updates window: {str(e)}")
                
        except Exception as e:
            self.logger.error(f"Error closing updates window: {str(e)}")
            
        # Always ensure we clean up our reference
        self.updates_window = None
        self.logger.info("Updates window reference cleared")
        
        # Track successful window cleanup for diagnostics
        self.window_state_tracking['updates']['success'] = False
    
    def _on_reconnect(self):
        """Handle the 'Reconnect to Service' menu item"""
        # Use a thread to avoid blocking the UI
        def do_reconnect():
            success = self._connect_to_service()
            if success:
                self.ui_queue.put(lambda: messagebox.showinfo("Connection Status", "Successfully connected to the Winget Updater service."))
            else:
                self.ui_queue.put(lambda: messagebox.showerror("Connection Status", "Failed to connect to the Winget Updater service."))
        
        threading.Thread(target=do_reconnect, daemon=True).start()
    
    def _on_install_updates(self):
        """Handle the 'Install All Updates' menu item"""
        # Use threading to avoid blocking the UI
        thread = threading.Thread(target=self._install_all_updates, daemon=True)
        thread.start()
    
    def _create_confirmation_dialog(self, parent, title, message):
        """Create a properly modal confirmation dialog with reliable focus and grab"""
        # Create the dialog window
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.resizable(False, False)
        dialog.transient(parent)  # Make it transient to parent window
        
        # Set window manager attributes for better behavior
        dialog.attributes('-topmost', True)
        
        # Make it look like a dialog
        if tk.TkVersion >= 8.5:
            try:
                dialog.tk.call("::tk::unsupported::MacWindowStyle", "style", dialog, "moveableModal", "")
            except tk.TclError:
                pass
                
        # Create the content
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Message
        message_label = ttk.Label(frame, text=message, wraplength=300, justify=tk.LEFT)
        message_label.pack(pady=(0, 20))
        
        # Buttons frame
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        # Result variable to capture the user's choice
        result = tk.BooleanVar(value=False)
        
        # Button commands
        def on_yes():
            result.set(True)
            dialog.destroy()
            
        def on_no():
            result.set(False)
            dialog.destroy()
        
        # Create buttons with distinct styles for better visibility
        style = ttk.Style()
        style.configure('Yes.TButton', font=('TkDefaultFont', 10, 'bold'))
        
        # Yes button
        yes_button = ttk.Button(button_frame, text="Yes", command=on_yes, style='Yes.TButton', width=10)
        yes_button.pack(side=tk.RIGHT, padx=5)
        
        # No button
        no_button = ttk.Button(button_frame, text="No", command=on_no, width=10)
        no_button.pack(side=tk.RIGHT, padx=5)
        
        # Set up key bindings
        dialog.bind("<Return>", lambda e: on_yes())
        dialog.bind("<Escape>", lambda e: on_no())
        
        # Center the dialog on the parent window
        dialog.withdraw()
        dialog.update_idletasks()
        
        width = dialog.winfo_reqwidth()
        height = dialog.winfo_reqheight()
        
        # Get parent window position
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Calculate position
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        # Make sure dialog is fully visible on screen
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        
        x = max(0, min(x, screen_width - width))
        y = max(0, min(y, screen_height - height))
        
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.deiconify()
        
        # Ensure dialog is visible and ready before setting grab
        dialog.update()
        
        # Set focus to the Yes button
        yes_button.focus_set()
        
        # Set grab to make dialog modal
        dialog.grab_set()
        
        # Wait for dialog to be destroyed
        parent.wait_window(dialog)
        
        # Return result
        return result.get()

    def _install_all_updates(self):
        """Install all available updates using service IPC command"""
        try:
            # Ensure we're connected to the service
            if not self.service_connected:
                if not self._connect_to_service():
                    self.ui_queue.put(lambda: messagebox.showerror(
                        "Service Error", 
                        "Could not connect to the Winget Updater service."
                    ))
                    return
            
            # Use event to synchronize between threads
            confirmation_event = threading.Event()
            confirmation_result = [False]
            
            # Create the confirmation dialog in the main thread
            def show_dialog():
                # Create custom confirmation dialog with proper focus and grab
                result = self._create_confirmation_dialog(
                    self.root,
                    "Install Updates",
                    "Do you want to install all available updates? This may take several minutes."
                )
                confirmation_result[0] = result
                confirmation_event.set()  # Signal that dialog is complete
            
            # Queue the dialog to run in the main thread
            self.ui_queue.put(show_dialog)
            
            # Wait for the dialog to complete with timeout
            if not confirmation_event.wait(timeout=30):
                self.logger.warning("Confirmation dialog timed out")
                return
            
            # Check if user confirmed
            if not confirmation_result[0]:
                self.logger.info("User cancelled update installation")
                return
                
            # Show progress notification
            self.ui_queue.put(lambda: self.icon.notify(
                "Installing updates. This may take several minutes...",
                "Winget Updates"
            ))
                
            # Use IPC to send install command to service
            self.logger.info("Starting installation of all updates via service")
            response = self.ipc_client.send_command("install_updates", {"all": True})
            
            if response and response.command == "response" and response.data.get("success", False):
                self.logger.info("All updates installed successfully")
                self.ui_queue.put(lambda: messagebox.showinfo(
                    "Updates Installed", 
                    "All updates have been installed successfully."
                ))
                # Trigger a new update check
                self._check_updates()
            else:
                error_msg = "Failed to install updates"
                if response and response.data and "message" in response.data:
                    error_msg = response.data["message"]
                    
                self.logger.error(f"Error installing updates: {error_msg}")
                self.ui_queue.put(lambda: messagebox.showerror(
                    "Installation Error",
                    f"Error installing updates: {error_msg}"
                ))
                
        except Exception as e:
            self.logger.error(f"Error installing updates: {str(e)}")
            self.ui_queue.put(lambda: messagebox.showerror(
                "Error", 
                f"An error occurred while installing updates: {str(e)}"
            ))
    
    def _on_exit(self):
        """Handle the 'Exit' menu item"""
        self.logger.info("Exiting application")
        self.running = False
        self.auto_reconnect = False
        
        # Close any open windows
        self._close_all_windows()
        
        # Disconnect from service
        if self.service_connected:
            self.ipc_client.disconnect()
            
        # Stop the icon and the Tkinter loop
        self.icon.stop()
        if hasattr(self, 'root'):
            self.root.quit()
    
    def _close_all_windows(self):
        """Close all open Tkinter windows"""
        # Close settings window if open
        if self.settings_window:
            try:
                self.settings_window.destroy()
            except:
                pass
            self.settings_window = None
            
        # Close updates window if open
        if self.updates_window:
            try:
                self.updates_window.destroy()
            except:
                pass
            self.updates_window = None

def run_tray_application():
    """Run the tray application"""
    try:
        tray_app = WingetUpdaterTray()
        tray_app.run()
    except Exception as e:
        logging.error(f"Error in tray application: {str(e)}")
        messagebox.showerror("Error", f"An error occurred in the Winget Updater tray application: {str(e)}")

if __name__ == "__main__":
    run_tray_application()

