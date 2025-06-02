"""
Window Manager for Winget Updater

This module provides centralized Tkinter window management to:
1. Maintain a single Tk root instance
2. Handle window ownership properly
3. Ensure windows are interactive and focused
4. Fix threading issues

It should be used for all Tkinter windows in the application.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import logging
import threading
import time
import queue
import os
import platform

class WindowManager:
    """
    Singleton class to manage all Tkinter windows in the application
    Ensures proper window creation, focus, and lifecycle management
    """
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(WindowManager, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
            
        self._initialized = True
        self.logger = logging.getLogger('WindowManager')
        self.logger.info("Initializing Window Manager")
        
        # Initialize Tkinter - will be done on first window creation
        self.root = None
        self.root_initialized = False
        
        # Dictionary to track all open windows
        self.windows = {}
        
        # Flag to track if the event processing thread is running
        self.event_thread_running = False
        self.event_thread = None
        
        # Command queue for thread-safe operations
        self.command_queue = queue.Queue()
        
        # Set process priority to high for better UI responsiveness on Windows
        if platform.system() == 'Windows':
            try:
                import psutil
                p = psutil.Process(os.getpid())
                p.nice(psutil.HIGH_PRIORITY_CLASS)
                self.logger.debug("Process priority set to high")
            except Exception as e:
                self.logger.debug(f"Could not set process priority: {str(e)}")
    
    def _initialize_root(self):
        """Initialize the root Tk instance if it doesn't exist"""
        if not self.root_initialized:
            self.logger.debug("Creating root Tk instance")
            
            # Create the root window
            self.root = tk.Tk()
            self.root.withdraw()  # Hide the root window
            
            # Set application name
            self.root.title("Winget Updater")
            
            # Configure Tk settings for better responsiveness
            if platform.system() == 'Windows':
                # Set DPI awareness on Windows
                try:
                    from ctypes import windll
                    windll.shcore.SetProcessDpiAwareness(1)
                    self.logger.debug("DPI awareness set")
                except Exception as e:
                    self.logger.debug(f"Could not set DPI awareness: {str(e)}")
            
            # Improve window drawing on Windows
            self.root.tk.call('tk', 'scaling', 1.0)
            
            # Handle window close explicitly
            self.root.protocol("WM_DELETE_WINDOW", self.shutdown)
            
            # Flag as initialized
            self.root_initialized = True
            
            # Start event processing
            self._start_event_processing()
            
            self.logger.debug("Root Tk instance created")
            
            # Add a small delay to ensure the root is fully initialized
            time.sleep(0.1)
    
    def _start_event_processing(self):
        """Start the event processing thread"""
        if not self.event_thread_running:
            self.event_thread_running = True
            self.event_thread = threading.Thread(
                target=self._process_events,
                daemon=True
            )
            self.event_thread.start()
            self.logger.debug("Event processing thread started")
    
    def _process_events(self):
        """Process Tkinter events in a separate thread"""
        try:
            self.logger.debug("Event processing thread started")
            event_count = 0
            last_log_time = time.time()
            
            while self.event_thread_running and self.root_initialized:
                try:
                    # Process events if the root exists
                    if self.root and self.root.winfo_exists():
                        # Process up to 100 events per cycle
                        for _ in range(100):
                            self.root.update_idletasks()
                            self.root.update()
                            event_count += 1
                            
                            # Process any commands in the queue
                            while not self.command_queue.empty():
                                try:
                                    cmd, args, kwargs = self.command_queue.get_nowait()
                                    if callable(cmd):
                                        cmd(*args, **kwargs)
                                    self.command_queue.task_done()
                                except queue.Empty:
                                    break
                                except Exception as e:
                                    self.logger.error(f"Error executing queued command: {str(e)}")
                    
                    # Check window visibility/focus
                    self._check_windows()
                    
                    # Log stats periodically
                    now = time.time()
                    if now - last_log_time > 60:  # Log every minute
                        self.logger.debug(f"Processed {event_count} events in the last minute")
                        event_count = 0
                        last_log_time = now
                    
                    # Sleep to prevent CPU overuse
                    time.sleep(0.01)  # Shorter sleep for better responsiveness
                except tk.TclError as e:
                    if "invalid command name" in str(e) or "application has been destroyed" in str(e):
                        self.logger.debug(f"Tk instance no longer valid: {str(e)}")
                        break
                    else:
                        self.logger.error(f"Tcl error in event processing: {str(e)}")
                        time.sleep(0.1)
                except Exception as e:
                    self.logger.error(f"Error in event processing: {str(e)}")
                    time.sleep(0.1)
        except Exception as e:
            self.logger.error(f"Event processing thread error: {str(e)}")
        finally:
            self.event_thread_running = False
            self.logger.debug("Event processing thread stopped")
    
    def _check_windows(self):
        """Check if windows need attention (focus, visibility)"""
        for window_id, window_info in list(self.windows.items()):
            window = window_info.get('window')
            if window:
                try:
                    # Skip if window doesn't exist anymore
                    if not window.winfo_exists():
                        self.logger.debug(f"Window {window_id} no longer exists, removing from tracking")
                        self.windows.pop(window_id, None)
                        continue
                    
                    # Check if window is minimized (iconified)
                    try:
                        if window.state() == 'iconic':
                            self.logger.debug(f"Window {window_id} is minimized, restoring")
                            window.deiconify()
                    except:
                        pass
                    
                    # Check if window needs focus
                    if window_info.get('needs_focus', False):
                        # Sequence of operations to ensure window is visible and focused
                        window.deiconify()
                        window.attributes('-topmost', True)
                        window.update()
                        window.lift()
                        window.focus_set()
                        window.focus_force()
                        window.grab_set()  # Set keyboard focus
                        
                        # Reset topmost after a short delay
                        def reset_topmost(win):
                            try:
                                if win.winfo_exists():
                                    win.attributes('-topmost', False)
                                    win.grab_release()
                            except:
                                pass
                                
                        window.after(1000, lambda w=window: reset_topmost(w))
                        window_info['needs_focus'] = False
                        self.logger.debug(f"Set focus for window {window_id}")
                        
                except tk.TclError as e:
                    if "invalid command name" in str(e) or "application has been destroyed" in str(e):
                        self.logger.debug(f"Window {window_id} was destroyed externally")
                    else:
                        self.logger.error(f"Tcl error checking window {window_id}: {str(e)}")
                    self.windows.pop(window_id, None)
                except Exception as e:
                    self.logger.error(f"Error checking window {window_id}: {str(e)}")
                    self.windows.pop(window_id, None)
    
    def _execute_in_main_thread(self, func, *args, **kwargs):
        """Execute a function in the main thread using the queue"""
        result = None
        exception = None
        done_event = threading.Event()
        
        # The function to run in the main thread
        def main_thread_func():
            nonlocal result, exception
            try:
                result = func(*args, **kwargs)
            except Exception as e:
                exception = e
            finally:
                done_event.set()
        
        # Queue the function
        self.command_queue.put((main_thread_func, [], {}))
        
        # Wait for the function to complete (with timeout)
        if not done_event.wait(timeout=5.0):
            raise TimeoutError("Operation timed out waiting for main thread")
            
        # Raise any exception that occurred
        if exception:
            raise exception
            
        return result
    
    def create_window(self, window_id, title, width=500, height=400, center=True,
                     resizable=True, topmost=True, icon=None):
        """
        Create a new Toplevel window
        
        Args:
            window_id: Unique identifier for the window
            title: Window title
            width: Window width
            height: Window height
            center: Whether to center the window on screen
            resizable: Whether the window is resizable
            topmost: Whether the window should be on top initially
            icon: Path to icon file
            
        Returns:
            The created Toplevel window
        """
        # This function needs to be thread-safe, so we'll use the main thread
        def _do_create_window():
            # Initialize root if needed
            self._initialize_root()
            
            # Check if window already exists
            if window_id in self.windows:
                try:
                    existing_window = self.windows[window_id]['window']
                    if existing_window.winfo_exists():
                        self.logger.debug(f"Window {window_id} already exists, bringing to front")
                        
                        # Make window visible with a sequence of operations for maximum reliability
                        existing_window.deiconify()
                        existing_window.attributes('-topmost', True)
                        existing_window.lift()
                        existing_window.focus_set()
                        existing_window.focus_force()
                        existing_window.update()
                        
                        # Set flag to keep checking focus
                        self.windows[window_id]['needs_focus'] = True
                        
                        # Reset topmost after a delay
                        existing_window.after(1000, lambda: existing_window.attributes('-topmost', False))
                        
                        return existing_window
                except tk.TclError:
                    # Window exists in our registry but is invalid - remove it
                    self.windows.pop(window_id, None)
                    
            # Create new window with exception handling
            try:
                self.logger.debug(f"Creating new window: {window_id}")
                new_window = tk.Toplevel(self.root)
                new_window.title(title)
                
                # Set initial position off-screen to avoid flicker during setup
                new_window.geometry(f"{width}x{height}-100-100")
                new_window.resizable(resizable, resizable)
                
                # Apply all settings before making visible
                self._configure_window(new_window, title, width, height, center, resizable, icon)
                
                # Make topmost and visible
                if topmost:
                    new_window.attributes('-topmost', True)
                    
                # Update to apply all settings
                new_window.update_idletasks()
                
                # Now make visible in the proper position
                if center:
                    self._center_window(new_window, width, height)
                
                # Store the window with focus flag
                self.windows[window_id] = {
                    'window': new_window,
                    'needs_focus': True,
                    'created_at': time.time()
                }
                
                # Set close handler
                new_window.protocol("WM_DELETE_WINDOW", lambda: self.close_window(window_id))
                
                # Set transient to root
                new_window.transient(self.root)
                
                # Final focus and deiconify
                new_window.deiconify()
                new_window.focus_force()
                new_window.lift()
                
                return new_window
                
            except Exception as e:
                self.logger.error(f"Error creating window {window_id}: {str(e)}")
                raise
        
        # Execute the window creation in the main thread
        try:
            # Get current thread ID
            current_thread_id = threading.get_ident()
            event_thread_id = self.event_thread.ident if self.event_thread else None
            
            # If we're already in the event thread, execute directly
            if current_thread_id == event_thread_id:
                return _do_create_window()
            else:
                # Otherwise execute via queue
                return self._execute_in_main_thread(_do_create_window)
                
        except Exception as e:
            self.logger.error(f"Failed to create window {window_id}: {str(e)}")
            raise
        
    def _configure_window(self, window, title, width, height, center, resizable, icon):
        """Apply standard configuration to a window"""
        # Set basic properties
        window.title(title)
        window.resizable(resizable, resizable)
        
        # Set icon if provided
        if icon and os.path.exists(icon):
            try:
                window.iconbitmap(icon)
            except Exception as e:
                self.logger.debug(f"Could not set window icon: {str(e)}")
                
        # Configure window appearance
        if platform.system() == "Windows":
            try:
                # Set window styles for better appearance on Windows
                from ctypes import windll
                GWL_STYLE = -16
                WS_MAXIMIZEBOX = 0x00010000
                WS_SIZEBOX = 0x00040000
                
                hwnd = windll.user32.GetParent(window.winfo_id())
                style = windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
                
                if not resizable:
                    style = style & ~WS_SIZEBOX & ~WS_MAXIMIZEBOX
                    windll.user32.SetWindowLongW(hwnd, GWL_STYLE, style)
            except:
                pass
    
    def _center_window(self, window, width, height):
        """Center a window on the screen"""
        # Get screen dimensions
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        
        # Calculate center position
        x = max(0, int(screen_width/2 - width/2))
        y = max(0, int(screen_height/2 - height/2))
        
        # Set the geometry
        window.geometry(f"{width}x{height}+{x}+{y}")
        
        # Update to apply the change
        window.update_idletasks()
    
    def close_window(self, window_id):
        """Close a window by its ID"""
        def _do_close_window():
            if window_id in self.windows:
                try:
                    window = self.windows[window_id]['window']
                    
                    # Check if window exists
                    if window.winfo_exists():
                        # Remove from tracking before destroying to avoid reentry issues
                        self.windows.pop(window_id, None)
                        
                        # Release any grabs
                        try:
                            window.grab_release()
                        except:
                            pass
                            
                        # Destroy the window
                        window.destroy()
                        
                        self.logger.debug(f"Window {window_id} closed")
                    else:
                        # Window doesn't exist, just remove from tracking
                        self.windows.pop(window_id, None)
                        self.logger.debug(f"Window {window_id} was already closed")
                        
                except tk.TclError as e:
                    if "invalid command name" in str(e) or "application has been destroyed" in str(e):
                        # Window was already destroyed
                        self.windows.pop(window_id, None)
                        self.logger.debug(f"Window {window_id} was already destroyed")
                    else:
                        self.logger.error(f"Tcl error closing window {window_id}: {str(e)}")
                        self.windows.pop(window_id, None)
                except Exception as e:
                    self.logger.error(f"Error closing window {window_id}: {str(e)}")
                    self.windows.pop(window_id, None)
        
        try:
            # Get current thread ID
            current_thread_id = threading.get_ident()
            event_thread_id = self.event_thread.ident if self.event_thread else None
            
            # If we're already in the event thread, execute directly
            if current_thread_id == event_thread_id:
                _do_close_window()
            else:
                # Queue the close operation
                self.command_queue.put((_do_close_window, [], {}))
        except Exception as e:
            self.logger.error(f"Failed to close window {window_id}: {str(e)}")
            # Try to remove from tracking
            if window_id in self.windows:
                self.windows.pop(window_id, None)
    
    def close_all_windows(self):
        """Close all open windows"""
        for window_id in list(self.windows.keys()):
            self.close_window(window_id)
    
    def shutdown(self):
        """Shutdown the window manager and cleanup resources"""
        self.logger.info("Shutting down Window Manager")
        
        # Set the shutdown flag
        self.event_thread_running = False
        
        # Clear the command queue
        while not self.command_queue.empty():
            try:
                self.command_queue.get_nowait()
                self.command_queue.task_done()
            except:
                pass
        
        # Close all windows
        window_ids = list(self.windows.keys())
        for window_id in window_ids:
            try:
                self.close_window(window_id)
            except:
                pass
        
        # Destroy root if it exists
        if self.root_initialized and self.root:
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass
            
        self.root = None
        self.root_initialized = False
        
        # Wait for event thread to finish (with timeout)
        if self.event_thread and self.event_thread.is_alive():
            try:
                self.logger.debug("Waiting for event thread to terminate...")
                self.event_thread.join(timeout=2.0)
                if self.event_thread.is_alive():
                    self.logger.warning("Event thread did not terminate within timeout")
                else:
                    self.logger.debug("Event thread terminated successfully")
            except Exception as e:
                self.logger.error(f"Error joining event thread: {str(e)}")
        
        self.logger.debug("Window Manager shutdown complete")

# Example usage:
if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(level=logging.DEBUG)
    
    # Create a test window
    window_manager = WindowManager()
    settings_window = window_manager.create_window(
        "settings", 
        "Settings Test", 
        width=450, 
        height=400
    )
    
    label = ttk.Label(settings_window, text="This is a test window")
    label.pack(padx=20, pady=20)
    
    button = ttk.Button(
        settings_window, 
        text="Close", 
        command=lambda: window_manager.close_window("settings")
    )
    button.pack(pady=10)
    
    # Keep script running
    try:
        while True:
            time.sleep(0.1)
    except KeyboardInterrupt:
        window_manager.shutdown()

