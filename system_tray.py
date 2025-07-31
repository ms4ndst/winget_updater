import os
import sys
import time
import asyncio
import threading
import logging
from PIL import Image, ImageDraw, ImageFont
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# Import our custom modules
from config_manager import ConfigManager
from update_checker import WingetUpdateChecker

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
    
    def __init__(self, root, config_manager, on_save_callback=None):
        self.root = root
        self.root.title("Winget Updater Settings")
        self.root.geometry("450x400")
        self.root.resizable(False, False)
        
        # Enable debug logging
        self.logger = logging.getLogger('SettingsWindow')
        
        # Enable more verbose logging for widget state issues
        self.logger.setLevel(logging.DEBUG)
        
        # Set icon if available
        try:
            self.root.iconbitmap("winget_updater.ico")
        except:
            pass
            
        self.config_manager = config_manager
        self.on_save_callback = on_save_callback
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Setup validation command
        vcmd = (self.root.register(self._validate_time_input), '%P')
        
        # Morning check time
        ttk.Label(main_frame, text="Morning Check Time:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.morning_time = tk.StringVar(value=self.config_manager.get_morning_check_time())
        # Create entry with validate and explicit state
        # Use standard tk.Entry with enhanced styling and binding
        self.morning_entry = tk.Entry(
            main_frame, 
            textvariable=self.morning_time, 
            width=10, 
            validate="key", 
            validatecommand=vcmd,
            state='normal',  # Set state directly in constructor
            relief="sunken",
            bd=1,
            font=("Arial", 10),
            bg="white",
            highlightthickness=1,
            highlightcolor="blue",
            insertwidth=2,
            insertbackground="blue"
        )
        self.morning_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        # Force enable the entry widget using multiple techniques
        self.morning_entry['state'] = 'normal'
        self.logger.debug(f"Morning entry initial state: {self.morning_entry['state']}")
        
        # Try alternate state setting methods
        try:
            self.morning_entry.config(state='normal')
            self.morning_entry.configure(state='normal')
        except Exception as e:
            self.logger.debug(f"Error setting morning entry state: {e}")
        
        # Add entry focus handlers
        # Add comprehensive event bindings
        self.morning_entry.bind("<FocusIn>", self._on_entry_focus_in)
        self.morning_entry.bind("<FocusOut>", self._on_entry_focus_out)
        self.morning_entry.bind("<Button-1>", self._on_entry_click)
        self.morning_entry.bind("<KeyRelease>", self._on_entry_key)
        # Add additional key bindings for better interaction
        self.morning_entry.bind("<KeyPress>", self._on_entry_keypress)
        self.morning_entry.bind("<Tab>", self._on_entry_tab)
        self.morning_entry.bind("<Return>", self._on_entry_return)
        # Add click and double-click bindings
        self.morning_entry.bind("<Double-Button-1>", self._on_entry_double_click)
        # Add right-click for context menu
        self.morning_entry.bind("<Button-3>", self._on_entry_right_click)
        ttk.Label(main_frame, text="Format: HH:MM (24-hour)").grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # Afternoon check time
        ttk.Label(main_frame, text="Afternoon Check Time:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.afternoon_time = tk.StringVar(value=self.config_manager.get_afternoon_check_time())
        # Create entry with validate and explicit state
        # Use standard tk.Entry with enhanced styling and binding
        self.afternoon_entry = tk.Entry(
            main_frame, 
            textvariable=self.afternoon_time, 
            width=10, 
            validate="key", 
            validatecommand=vcmd,
            state='normal',  # Set state directly in constructor
            relief="sunken",
            bd=1,
            font=("Arial", 10),
            bg="white",
            highlightthickness=1,
            highlightcolor="blue",
            insertwidth=2,
            insertbackground="blue"
        )
        self.afternoon_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        # Force enable the entry widget using multiple techniques
        self.afternoon_entry['state'] = 'normal'
        self.logger.debug(f"Afternoon entry initial state: {self.afternoon_entry['state']}")
        
        # Try alternate state setting methods
        try:
            self.afternoon_entry.config(state='normal')
            self.afternoon_entry.configure(state='normal')
        except Exception as e:
            self.logger.debug(f"Error setting afternoon entry state: {e}")
        
        # Add comprehensive event bindings
        self.afternoon_entry.bind("<FocusIn>", self._on_entry_focus_in)
        self.afternoon_entry.bind("<FocusOut>", self._on_entry_focus_out)
        self.afternoon_entry.bind("<Button-1>", self._on_entry_click)
        self.afternoon_entry.bind("<KeyRelease>", self._on_entry_key)
        # Add additional key bindings for better interaction
        self.afternoon_entry.bind("<KeyPress>", self._on_entry_keypress)
        self.afternoon_entry.bind("<Tab>", self._on_entry_tab)
        self.afternoon_entry.bind("<Return>", self._on_entry_return)
        # Add click and double-click bindings
        self.afternoon_entry.bind("<Double-Button-1>", self._on_entry_double_click)
        # Add right-click for context menu
        self.afternoon_entry.bind("<Button-3>", self._on_entry_right_click)
        ttk.Label(main_frame, text="Format: HH:MM (24-hour)").grid(row=1, column=2, sticky=tk.W, pady=5)
        
        # Add trace callbacks to monitor variable changes
        self.morning_time.trace_add("write", self._on_var_change)
        self.afternoon_time.trace_add("write", self._on_var_change)
        
        # Notify on updates
        self.notify_on_updates = tk.BooleanVar(value=self.config_manager.get_notify_on_updates())
        self.notify_cb = ttk.Checkbutton(
            main_frame, 
            text="Show notifications when updates are available", 
            variable=self.notify_on_updates,
            command=self._update_widget_states
        )
        self.notify_cb.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        # Force enable state
        self.notify_cb.state(['!disabled'])
        
        # Auto-check for updates
        self.auto_check = tk.BooleanVar(value=self.config_manager.get_auto_check())
        self.auto_cb = ttk.Checkbutton(
            main_frame, 
            text="Automatically check for updates at scheduled times", 
            variable=self.auto_check,
            command=self._update_widget_states
        )
        self.auto_cb.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        # Force enable state
        self.auto_cb.state(['!disabled'])
        
        # Add trace callbacks for checkboxes
        self.notify_on_updates.trace_add("write", self._on_var_change)
        self.auto_check.trace_add("write", self._on_var_change)
        
        # Update filter options
        ttk.Label(main_frame, text="Update Filter Options:", font=("", 9, "bold")).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(15, 5))
        
        # Include pinned packages
        self.include_pinned = tk.BooleanVar(value=self.config_manager.get_include_pinned_updates())
        self.pinned_cb = ttk.Checkbutton(
            main_frame, 
            text="Include pinned packages in update checks", 
            variable=self.include_pinned,
            command=self._update_widget_states
        )
        self.pinned_cb.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)
        # Force enable state
        self.pinned_cb.state(['!disabled'])
        self._create_tooltip(self.pinned_cb, "Show updates for packages that have been pinned using 'winget pin'")
        
        # Include unknown versions
        self.include_unknown = tk.BooleanVar(value=self.config_manager.get_include_unknown_versions())
        self.unknown_cb = ttk.Checkbutton(
            main_frame, 
            text="Include packages with unknown versions", 
            variable=self.include_unknown,
            command=self._update_widget_states
        )
        self.unknown_cb.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=5)
        # Force enable state
        self.unknown_cb.state(['!disabled'])
        self._create_tooltip(self.unknown_cb, "Show updates for packages where the current version is unknown or undetectable")
        
        # Add trace callbacks for new checkboxes
        self.include_pinned.trace_add("write", self._on_var_change)
        self.include_unknown.trace_add("write", self._on_var_change)
        
        # Last check time
        last_check = self.config_manager.get_last_check()
        last_check_str = "Never" if last_check is None else last_check.strftime("%Y-%m-%d %H:%M:%S")
        ttk.Label(main_frame, text=f"Last check: {last_check_str}").grid(row=7, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=8, column=0, columnspan=3, sticky=tk.E, pady=10)
        
        self.save_btn = ttk.Button(btn_frame, text="Save", command=self.save_settings)
        self.save_btn.pack(side=tk.RIGHT, padx=5)
        self.cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.root.destroy)
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Set up keyboard navigation
        self.root.bind("<Return>", lambda e: self.save_settings())
        self.root.bind("<Escape>", lambda e: self.root.destroy())
        
        # Focus cycle bindings (tab and shift-tab)
        self._setup_focus_cycle()
        
        # Force all widgets to be enabled
        self._force_enable_all_widgets()
        
        # Update widget states based on dependencies
        self._update_widget_states()
        
        # Set initial focus with explicit selection
        self.morning_entry.focus_set()
        self.morning_entry.selection_range(0, tk.END)  # Select all text
        
        # Add an explicit click handler to the window
        self.root.bind("<Button-1>", self._on_window_click)
        
        # Update widget states based on current settings
        self._update_widget_states()
        
        # Log initial variable values and widget states
        self.logger.info(f"Initial settings - morning: {self.morning_time.get()}, afternoon: {self.afternoon_time.get()}, "
                         f"notify: {self.notify_on_updates.get()}, auto: {self.auto_check.get()}, "
                         f"include_pinned: {self.include_pinned.get()}, include_unknown: {self.include_unknown.get()}")
                         
        # Schedule periodic state checks
        self._schedule_state_checks()
        
        # Force widget states after a delay
        self.root.after(500, self._force_enable_all_widgets)
        self.root.after(1000, self._force_enable_all_widgets)
        self.root.after(2000, self._force_enable_all_widgets)
    
    def _setup_focus_cycle(self):
        """Set up tab order for widgets"""
        # Define the widget order for tab navigation
        widgets = [
            self.morning_entry,
            self.afternoon_entry,
            self.notify_cb,
            self.auto_cb,
            self.pinned_cb,
            self.unknown_cb,
            self.save_btn,
            self.cancel_btn
        ]
        
        # Bind tab keys for each widget
        for i, widget in enumerate(widgets):
            # Next widget in cycle (Tab)
            next_idx = (i + 1) % len(widgets)
            widget.bind("<Tab>", lambda e, w=widgets[next_idx]: self._focus_widget(w))
            
            # Previous widget in cycle (Shift+Tab)
            prev_idx = (i - 1) % len(widgets)
            widget.bind("<Shift-Tab>", lambda e, w=widgets[prev_idx]: self._focus_widget(w))
    
    def _focus_widget(self, widget):
        """Focus a widget and ensure it's visible and enabled"""
        try:
            widget.focus_set()
            # For entry widgets, select all text
            if isinstance(widget, ttk.Entry):
                widget.selection_range(0, tk.END)
        except Exception as e:
            self.logger.error(f"Error focusing widget: {str(e)}")
        return "break"  # Prevent default tab behavior
    
    def _force_enable_all_widgets(self):
        """Force all widgets to be enabled"""
        try:
            # Force entry widgets to be enabled with multiple methods
            for entry in [self.morning_entry, self.afternoon_entry]:
                # Use multiple methods to force enable
                entry['state'] = 'normal'
                entry.config(state='normal')
                
                # Enhanced styling for better visibility
                entry.config(
                    background='white',
                    disabledbackground='#f0f0f0',
                    readonlybackground='#f5f5f5',
                    highlightthickness=1,
                    highlightcolor="gray",
                    highlightbackground="white",
                    insertwidth=2,
                    insertbackground="blue",
                    cursor="xterm"
                )
                
                # For ttk widgets, try state method
                try:
                    if hasattr(entry, 'state'):
                        entry.state(['!disabled'])
                except:
                    pass
            
            # Force checkbuttons to be enabled
            self.notify_cb.state(['!disabled'])
            self.auto_cb.state(['!disabled'])
            self.pinned_cb.state(['!disabled'])
            self.unknown_cb.state(['!disabled'])
            
            # Force buttons to be enabled
            self.save_btn['state'] = 'normal'
            self.cancel_btn['state'] = 'normal'
            
            # Update the root window
            self.root.update_idletasks()
            
            # Log current states for debugging
            self.logger.debug("--- Widget States After Force Enable ---")
            self.logger.debug(f"Morning entry state: {self.morning_entry['state']}")
            self.logger.debug(f"Afternoon entry state: {self.afternoon_entry['state']}")
            self.logger.debug(f"Auto check state: {self.auto_cb.instate(['disabled'])}")
            self.logger.debug(f"Include pinned state: {self.pinned_cb.instate(['disabled'])}")
            self.logger.debug(f"Include unknown state: {self.unknown_cb.instate(['disabled'])}")
            
        except Exception as e:
            self.logger.error(f"Error forcing widget states: {str(e)}")
    
    def _update_widget_states(self):
        """Update widget states based on dependencies"""
        try:
            # Force enable all widgets first
            self._force_enable_all_widgets()
            
            # Dependency: Time settings should only be enabled if auto_check is True
            if not self.auto_check.get():
                # Disable time entry widgets if auto-check is off
                self.morning_entry['state'] = 'disabled'
                self.afternoon_entry['state'] = 'disabled'
                
                # Change background to indicate disabled state
                self.morning_entry.config(background='#f0f0f0')
                self.afternoon_entry.config(background='#f0f0f0')
            else:
                # Ensure entry widgets are enabled
                self.morning_entry['state'] = 'normal'
                self.afternoon_entry['state'] = 'normal'
                
                # Normal background color
                self.morning_entry.config(background='white')
                self.afternoon_entry.config(background='white')
        except Exception as e:
            self.logger.error(f"Error updating widget states: {str(e)}")
            # If there's an error, force enable everything
            self._force_enable_all_widgets()
    
    def _validate_time_input(self, value):
        """Validate time input (allows partial input during typing)"""
        # Always allow empty string or partial input
        if not value or len(value) <= 5:
            return True
            
        # Check if the input is a valid time format
        try:
            # Check for special patterns that are valid during typing
            # e.g., "1:" or "12:"
            if ":" in value and len(value) < 5:
                parts = value.split(':')
                if len(parts) == 2 and parts[0].isdigit() and (not parts[1] or parts[1].isdigit()):
                    hours = int(parts[0])
                    if parts[1]:
                        minutes = int(parts[1])
                        if hours < 0 or hours > 23 or minutes < 0 or minutes > 59:
                            return False
                    else:
                        if hours < 0 or hours > 23:
                            return False
                    return True
            
            # Full format validation
            parts = value.split(':')
            if len(parts) != 2:
                return False
                
            hours, minutes = int(parts[0]), int(parts[1])
            if hours < 0 or hours > 23 or minutes < 0 or minutes > 59:
                return False
                
            return True
        except Exception as e:
            self.logger.debug(f"Time validation error: {str(e)}")
            return False
    
    def _on_var_change(self, *args):
        """Track variable changes for debugging"""
        try:
            var_name = args[0]
            if var_name == str(self.morning_time):
                self.logger.info(f"Morning time changed to: {self.morning_time.get()}")
            elif var_name == str(self.afternoon_time):
                self.logger.info(f"Afternoon time changed to: {self.afternoon_time.get()}")
            elif var_name == str(self.notify_on_updates):
                self.logger.info(f"Notify on updates changed to: {self.notify_on_updates.get()}")
            elif var_name == str(self.auto_check):
                self.logger.info(f"Auto check changed to: {self.auto_check.get()}")
                # Update widget states when auto_check changes
                self._update_widget_states()
            elif var_name == str(self.include_pinned):
                self.logger.info(f"Include pinned changed to: {self.include_pinned.get()}")
            elif var_name == str(self.include_unknown):
                self.logger.info(f"Include unknown changed to: {self.include_unknown.get()}")
        except Exception as e:
            self.logger.error(f"Error in variable change callback: {str(e)}")

    def _on_entry_focus_in(self, event):
        """Handle entry field focus in"""
        try:
            # Ensure the entry is enabled
            entry = event.widget
            entry['state'] = 'normal'
            
            # Enhance visual appearance to indicate focus
            entry.config(
                highlightthickness=2,
                highlightcolor="blue",
                highlightbackground="lightblue"
            )
            
            # Select all text
            entry.selection_range(0, tk.END)
            
            # Ensure proper cursor
            entry.config(cursor="xterm")
            
            # Log the event
            self.logger.debug(f"Entry focus in: {entry}")
        except Exception as e:
            self.logger.error(f"Error in entry focus in: {str(e)}")
    
    def _on_entry_focus_out(self, event):
        """Handle entry field focus out"""
        try:
            # Validate on focus out
            entry = event.widget
            value = None
            
            # Determine which entry and get its value
            if entry == self.morning_entry:
                value = self.morning_time.get()
                self.logger.debug(f"Morning time focus out: {value}")
            elif entry == self.afternoon_entry:
                value = self.afternoon_time.get()
                self.logger.debug(f"Afternoon time focus out: {value}")
            
            # Format the time value if needed
            if value and len(value) > 0:
                # Check if it needs to be padded
                if ":" in value:
                    parts = value.split(":")
                    if len(parts) == 2:
                        hours, minutes = parts
                        # Pad with zeros if needed
                        if hours.isdigit() and minutes.isdigit():
                            formatted = f"{int(hours):02d}:{int(minutes):02d}"
                            if entry == self.morning_entry:
                                self.morning_time.set(formatted)
                            elif entry == self.afternoon_entry:
                                self.afternoon_time.set(formatted)
                                
            # Always ensure the entry remains enabled
            if self.auto_check.get():
                entry['state'] = 'normal'
                
            # Reset highlight appearance
            entry.config(
                highlightthickness=1,
                highlightcolor="gray",
                highlightbackground="white"
            )
                
        except Exception as e:
            self.logger.error(f"Error in entry focus out: {str(e)}")
    
    def _on_entry_click(self, event):
        """Handle entry field click"""
        try:
            # Ensure the entry is enabled and focused
            entry = event.widget
            entry['state'] = 'normal'
            
            # Enhanced visual feedback
            entry.config(
                highlightthickness=2,
                highlightcolor="blue",
                highlightbackground="lightblue",
                cursor="xterm"
            )
            
            # Force focus and selection
            self.root.after(10, lambda: entry.focus_set())
            self.root.after(20, lambda: entry.selection_range(0, tk.END))
            
            # Move cursor to end if no selection
            if not entry.selection_present():
                entry.icursor(tk.END)
            
            # Log the event
            self.logger.debug(f"Entry clicked: {entry}")
        except Exception as e:
            self.logger.error(f"Error in entry click: {str(e)}")
    
    def _on_entry_key(self, event):
        """Handle entry key events"""
        try:
            # Always ensure the entry remains enabled after key press
            entry = event.widget
            entry['state'] = 'normal'
            
            # Handle special keys
            if event.keysym == 'Return':
                # Move to next field or save on Enter
                if entry == self.morning_entry:
                    self.afternoon_entry.focus_set()
                else:
                    self.save_settings()
            elif event.keysym == 'Escape':
                # Cancel on Escape
                self.root.destroy()
            elif event.keysym == 'Tab':
                # Let the normal tab handling work
                pass
            
            # Format the time as user types
            value = entry.get()
            if len(value) == 2 and value.isdigit() and ':' not in value:
                # Add colon after hours
                if entry == self.morning_entry:
                    self.morning_time.set(f"{value}:")
                    entry.icursor(tk.END)  # Move cursor to end
                elif entry == self.afternoon_entry:
                    self.afternoon_time.set(f"{value}:")
                    entry.icursor(tk.END)  # Move cursor to end
                    
        except Exception as e:
            self.logger.error(f"Error in entry key event: {str(e)}")
    
    def _on_entry_keypress(self, event):
        """Handle entry key press events"""
        try:
            # Always ensure the entry remains enabled
            entry = event.widget
            entry['state'] = 'normal'
            
            # Ensure visual feedback for typing
            entry.config(cursor="xterm")
            
        except Exception as e:
            self.logger.error(f"Error in entry keypress event: {str(e)}")
    
    def _on_entry_tab(self, event):
        """Handle tab key in entry fields"""
        try:
            # Move to next field
            entry = event.widget
            if entry == self.morning_entry:
                self.afternoon_entry.focus_set()
                self.afternoon_entry.selection_range(0, tk.END)
                return "break"  # Prevent default tab behavior
            elif entry == self.afternoon_entry:
                self.notify_cb.focus_set()
                return "break"  # Prevent default tab behavior
        except Exception as e:
            self.logger.error(f"Error in entry tab event: {str(e)}")
    
    def _on_entry_return(self, event):
        """Handle return key in entry fields"""
        try:
            # Move to next field or save
            entry = event.widget
            if entry == self.morning_entry:
                self.afternoon_entry.focus_set()
                self.afternoon_entry.selection_range(0, tk.END)
                return "break"  # Prevent default behavior
            else:
                self.save_settings()
                return "break"  # Prevent default behavior
        except Exception as e:
            self.logger.error(f"Error in entry return event: {str(e)}")
    
    def _on_entry_double_click(self, event):
        """Handle double click in entry fields"""
        try:
            # Select all text
            entry = event.widget
            entry.selection_range(0, tk.END)
            return "break"  # Prevent default behavior
        except Exception as e:
            self.logger.error(f"Error in entry double click event: {str(e)}")
    
    def _on_entry_right_click(self, event):
        """Handle right click in entry fields"""
        try:
            # Create a simple context menu
            entry = event.widget
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Cut", command=lambda: entry.event_generate("<<Cut>>"))
            menu.add_command(label="Copy", command=lambda: entry.event_generate("<<Copy>>"))
            menu.add_command(label="Paste", command=lambda: entry.event_generate("<<Paste>>"))
            menu.add_separator()
            menu.add_command(label="Select All", command=lambda: entry.selection_range(0, tk.END))
            
            # Display context menu
            menu.tk_popup(event.x_root, event.y_root)
            return "break"  # Prevent default behavior
        except Exception as e:
            self.logger.error(f"Error in entry right click event: {str(e)}")
            
    def _on_window_click(self, event):
        """Handle clicks anywhere in the window to ensure focus"""
        try:
            # Force enable entry widgets in case they've been disabled
            if self.auto_check.get():
                self.morning_entry['state'] = 'normal'
                self.afternoon_entry['state'] = 'normal'
        except Exception as e:
            self.logger.error(f"Error in window click handler: {str(e)}")
    
    def _schedule_state_checks(self):
        """Schedule periodic checks of widget states"""
        try:
            # Log current widget states
            self.logger.debug("--- Periodic Widget State Check ---")
            self.logger.debug(f"Morning entry state: {self.morning_entry['state']}")
            self.logger.debug(f"Afternoon entry state: {self.afternoon_entry['state']}")
            
            # If auto-check is enabled but entries are disabled, force enable them
            if self.auto_check.get():
                if self.morning_entry['state'] != 'normal' or self.afternoon_entry['state'] != 'normal':
                    self.logger.debug("Auto-correcting entry states...")
                    self.morning_entry['state'] = 'normal'
                    self.afternoon_entry['state'] = 'normal'
            
            # Schedule next check
            self.root.after(5000, self._schedule_state_checks)
        except Exception as e:
            self.logger.error(f"Error in scheduled state check: {str(e)}")

    def save_settings(self):
        """Save the settings to the configuration file"""
        self.logger.info("Saving settings...")
        try:
            # Validate time formats
            morning_time = self.morning_time.get()
            afternoon_time = self.afternoon_time.get()
            
            self.logger.info(f"Values to save - morning: {morning_time}, afternoon: {afternoon_time}, "
                            f"notify: {self.notify_on_updates.get()}, auto: {self.auto_check.get()}, "
                            f"include_pinned: {self.include_pinned.get()}, include_unknown: {self.include_unknown.get()}")
            
            # Very basic validation - proper validation would be more thorough
            if not self._validate_time_format(morning_time) or not self._validate_time_format(afternoon_time):
                self.logger.warning(f"Invalid time format - morning: {morning_time}, afternoon: {afternoon_time}")
                messagebox.showerror("Invalid Time Format", "Please enter times in HH:MM format (24-hour)")
                return
                
            # Save settings
            self.config_manager.set_morning_check_time(morning_time)
            self.config_manager.set_afternoon_check_time(afternoon_time)
            self.config_manager.set_notify_on_updates(self.notify_on_updates.get())
            self.config_manager.set_auto_check(self.auto_check.get())
            self.config_manager.set_include_pinned_updates(self.include_pinned.get())
            self.config_manager.set_include_unknown_versions(self.include_unknown.get())
            
            self.logger.info("Settings saved successfully")
            
            # Call the callback if provided
            if self.on_save_callback:
                self.on_save_callback()
                
            messagebox.showinfo("Settings Saved", "Your settings have been saved successfully.")
            self.root.destroy()
            
        except Exception as e:
            self.logger.error(f"Error saving settings: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while saving settings: {str(e)}")
    
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
        except Exception as e:
            self.logger.debug(f"Time validation failed: {time_str} - {str(e)}")
            return False
            
    def _create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            
            # Create a toplevel window
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            
            # Add a label with the tooltip text
            label = tk.Label(self.tooltip, text=text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"), padx=5, pady=2)
            label.pack(ipadx=1)
            
        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
                
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

class SystemTrayIcon:
    """Class to manage the system tray icon and its functionality"""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.update_checker = WingetUpdateChecker(self.config_manager)
        
        self._setup_logging()
        self._create_icon()
        self._setup_menu()
        
        # For thread synchronization
        self._window_lock = threading.Lock()
        
        # Track the previous update count to detect changes
        self.previous_update_count = 0
        
        # For storing the tkinter root windows
        self.settings_window = None
        self.updates_window = None
        
        # Flag to track if the application is running
        self.running = True
        
        # Start the scheduler thread
        self.scheduler_thread = threading.Thread(target=self._run_scheduler, daemon=True)
        self.scheduler_thread.start()
    
    def _setup_logging(self):
        """Set up logging for the system tray component"""
        self.logger = logging.getLogger('SystemTrayIcon')
    
    def _create_icon(self):
        """Create the system tray icon"""
        # Default icon image - a simple square
        image = self._create_icon_image(0)
        
        # Create the icon
        self.icon = pystray.Icon("winget_updater", image, "Winget Updater")
    
    def _create_icon_image(self, update_count):
        """Create a custom icon image with the update count"""
        # Create a blank image
        width, height = 64, 64
        image = Image.new('RGBA', (width, height), color=(0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        # Draw a different background color based on update availability
        if update_count > 0:
            # Green background for available updates
            draw.rectangle([(0, 0), (width, height)], fill=(0, 150, 0, 255))
        else:
            # Blue background for no updates
            draw.rectangle([(0, 0), (width, height)], fill=(0, 0, 150, 255))
        
        # Add the update count as text if updates are available
        if update_count > 0:
            try:
                # Try to use a TrueType font if available
                font = ImageFont.truetype("arial.ttf", 32)
            except:
                # Fall back to default font
                font = ImageFont.load_default()
            
            # Convert the count to a string, limit to 2 digits
            count_str = str(min(update_count, 99))
            if update_count > 99:
                count_str = "99+"
                
            # Calculate text size and position for centering
            text_width = draw.textlength(count_str, font=font)
            text_height = 32  # Approximate height
            
            # Draw the text
            draw.text(
                ((width - text_width) // 2, (height - text_height) // 2),
                count_str,
                font=font,
                fill=(255, 255, 255, 255)
            )
        
        return image
    
    def _setup_menu(self):
        """Create the system tray context menu"""
        self.icon.menu = pystray.Menu(
            item('Check for Updates', self._on_check_updates),
            item('Install All Updates', self._on_install_updates, enabled=lambda _: self.update_checker.get_update_count() > 0),
            item('Show Updates', self._on_show_updates),
            item('Settings', self._on_open_settings),
            item('Exit', self._on_exit)
        )
    
    def run(self):
        """Run the system tray application"""
        self.logger.info("Starting Winget Updater system tray application")
        
        # Initial update check
        self._check_updates()
        
        # Run the icon loop
        self.icon.run()
    
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
                # This gives a reasonable checking interval without excessive CPU usage
                time.sleep(30)
                
            except Exception as e:
                self.logger.error(f"Error in scheduler thread: {str(e)}")
                time.sleep(60)  # Sleep longer on error
    
    def _check_updates(self):
        """Check for updates and update the icon"""
        self.logger.info("Manually checking for updates")
        
        # Run the update check synchronously
        update_count = self.update_checker.check_updates()
        
        # Update the icon with the new count
        self._update_icon(update_count)
        
        # Show notification if count changed and notifications are enabled
        if (update_count > 0 and 
            update_count != self.previous_update_count and 
            self.config_manager.get_notify_on_updates()):
            self._show_notification(update_count)
        
        # Update the previous count
        self.previous_update_count = update_count
    
    def _update_icon(self, update_count):
        """Update the system tray icon with the current update count"""
        new_image = self._create_icon_image(update_count)
        
        # Update the icon image
        self.icon.icon = new_image
        
        # Update the tooltip text
        if update_count == 0:
            self.icon.title = "Winget Updater - No updates available"
        elif update_count == 1:
            self.icon.title = "Winget Updater - 1 update available"
        else:
            self.icon.title = f"Winget Updater - {update_count} updates available"
    
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
    
    def _on_check_updates(self):
        """Handle the 'Check for Updates' menu item"""
        self.logger.info("Update check requested via menu")
        
        try:
            # Show checking notification
            self.icon.notify(
                "Checking for updates...",
                "Winget Updater"
            )
            
            # Use threading to avoid blocking the UI
            thread = threading.Thread(target=self._check_updates, daemon=True)
            thread.start()
            
            self.logger.debug("Update check thread started")
            
        except Exception as e:
            self.logger.error(f"Error initiating update check: {str(e)}")
            # Use messagebox directly since we're still in the main thread
            messagebox.showerror(
                "Error",
                f"Failed to start update check: {str(e)}"
            )
    
    def _on_open_settings(self):
        """Handle the 'Settings' menu item in a new thread"""
        if self.settings_window and self.settings_window.winfo_exists():
            self.logger.info("Settings window already open, bringing to front")
            self._focus_existing_window(self.settings_window)
            return
        
        # Run in a separate thread to avoid blocking
        thread = threading.Thread(target=self._create_settings_window, daemon=True)
        thread.start()

    def _create_settings_window(self):
        """Create and run the settings window in a separate thread"""
        with self._window_lock:
            if self.settings_window and self.settings_window.winfo_exists():
                self._focus_existing_window(self.settings_window)
                return

            try:
                root = tk.Tk()
                self.settings_window = root

                settings = SettingsWindow(root, self.config_manager, self._on_settings_saved)
                
                root.protocol("WM_DELETE_WINDOW", lambda: self._on_window_closed(self.settings_window, "settings"))
                
                self._run_and_cleanup_window(root, "settings")
            except Exception as e:
                self.logger.error(f"Error creating settings window: {str(e)}")
                self.settings_window = None

    def _on_show_updates(self):
        """Handle the 'Show Updates' menu item in a new thread"""
        if self.updates_window and self.updates_window.winfo_exists():
            self.logger.info("Updates window already open, bringing to front")
            self._focus_existing_window(self.updates_window)
            return

        updates = self.update_checker.get_updates_list()
        if not updates:
            self.icon.notify("No updates are currently available.", "Winget Updates")
            return

        # Run in a separate thread to avoid blocking
        thread = threading.Thread(target=self._create_updates_window, args=(updates,), daemon=True)
        thread.start()

    def _create_updates_window(self, updates):
        """Create and run the updates window in a separate thread"""
        with self._window_lock:
            if self.updates_window and self.updates_window.winfo_exists():
                self._focus_existing_window(self.updates_window)
                return

            try:
                root = tk.Tk()
                self.updates_window = root

                UpdateListWindow(root, updates)

                root.protocol("WM_DELETE_WINDOW", lambda: self._on_window_closed(self.updates_window, "updates"))

                self._run_and_cleanup_window(root, "updates")
            except Exception as e:
                self.logger.error(f"Error creating updates window: {str(e)}")
                self.updates_window = None

    def _focus_existing_window(self, window):
        """Bring an existing window to the front and focus it"""
        try:
            window.deiconify()
            window.lift()
            window.attributes('-topmost', True)
            window.after(100, lambda: window.attributes('-topmost', False))
            window.focus_force()
        except Exception as e:
            self.logger.warning(f"Error focusing existing window: {str(e)}")

    def _on_window_closed(self, window, window_type):
        """Handle a window being closed"""
        with self._window_lock:
            try:
                window.destroy()
            except tk.TclError:
                pass  # Ignore errors if window is already destroyed
            finally:
                if window_type == "settings":
                    self.settings_window = None
                elif window_type == "updates":
                    self.updates_window = None
                self.logger.info(f"{window_type.capitalize()} window closed")

    def _run_and_cleanup_window(self, window, window_type):
        """Run the main loop for a window and ensure cleanup"""
        try:
            window.mainloop()
        finally:
            self._on_window_closed(window, window_type)
    
    def _on_settings_saved(self):
        """Handle settings being saved"""
        self.logger.info("Settings were updated")
        
        # Force a re-check of updates with the new settings
        try:
            self.logger.debug("Running update check with new settings")
            self._check_updates()
        except Exception as e:
            self.logger.error(f"Error checking updates after settings change: {str(e)}")
    
    def _create_confirmation_dialog(self, title, message):
        """Create a custom confirmation dialog that works properly on Windows"""
        # Create root window
        root = tk.Tk()
        root.title(title)
        root.resizable(False, False)
        root.attributes('-topmost', True)
        
        # Center on screen
        root.update_idletasks()
        width = 400
        height = 150
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Variable to store the result
        result = tk.BooleanVar(value=False)
        dialog_closed = threading.Event()
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message label
        message_label = ttk.Label(main_frame, text=message, wraplength=360, justify=tk.CENTER)
        message_label.pack(pady=(0, 20))
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # Button commands
        def on_yes():
            result.set(True)
            dialog_closed.set()
            root.quit()
            
        def on_no():
            result.set(False)
            dialog_closed.set()
            root.quit()
            
        def on_close():
            result.set(False)
            dialog_closed.set()
            root.quit()
        
        # Create buttons
        no_button = ttk.Button(button_frame, text="No", command=on_no, width=12)
        no_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        yes_button = ttk.Button(button_frame, text="Yes", command=on_yes, width=12)
        yes_button.pack(side=tk.RIGHT, padx=(5, 5))
        
        # Set up key bindings
        root.bind('<Return>', lambda e: on_yes())
        root.bind('<Escape>', lambda e: on_no())
        
        # Handle window close button
        root.protocol("WM_DELETE_WINDOW", on_close)
        
        # Focus the dialog
        root.focus_force()
        yes_button.focus_set()
        
        # Run the dialog
        try:
            root.mainloop()
        finally:
            try:
                root.destroy()
            except:
                pass
        
        return result.get()
    
    def _on_install_updates(self):
        """Handle the 'Install All Updates' menu item"""
        self.logger.info("Installation of all updates requested via menu")
        
        try:
            # Get the update count
            update_count = self.update_checker.get_update_count()
            
            if update_count == 0:
                self.icon.notify(
                    "No updates available to install.",
                    "Winget Updater"
                )
                return
            
            # Create custom confirmation dialog
            message = (f"Are you sure you want to install {update_count} updates?\n\n"
                      "This may take several minutes and require application restarts.")
            
            if not self._create_confirmation_dialog("Confirm Installation", message):
                self.logger.info("User cancelled installation")
                return
            
            # Show a notification that installation is starting
            self.icon.notify(
                "Starting installation of updates...",
                "Winget Updater"
            )
            
            # Start installation in a separate thread to avoid blocking the UI
            thread = threading.Thread(
                target=self._perform_installation,
                daemon=True
            )
            thread.start()
                
        except Exception as e:
            self.logger.error(f"Error initiating installation: {str(e)}")
            self.icon.notify(
                f"Failed to start installation: {str(e)}",
                "Winget Updater"
            )
    
    def _perform_installation(self):
        """Execute the installation process"""
        try:
            # Perform the installation
            success = self.update_checker.install_all_updates()
            
            if success:
                self.icon.notify(
                    "All updates were installed successfully.",
                    "Winget Updater"
                )
                # Refresh the update status after installation
                self._check_updates()
            else:
                self.icon.notify(
                    "Some updates failed to install. Check the log for details.",
                    "Winget Updater"
                )
                
        except Exception as e:
            self.logger.error(f"Error during installation: {str(e)}")
            self.icon.notify(
                f"Failed to install updates: {str(e)}",
                "Winget Updater"
            )
    
    def _on_exit(self):
        """Handle the 'Exit' menu item"""
        self.logger.info("Exiting application")
        self.running = False
        
        # Clean up open windows
        if self.settings_window:
            try:
                self.settings_window.destroy()
            except:
                pass
            self.settings_window = None
            
        if self.updates_window:
            try:
                self.updates_window.destroy()
            except:
                pass
            self.updates_window = None
                
        # Stop the system tray icon
        self.icon.stop()
        
# If this module is run directly, start the system tray application
if __name__ == "__main__":
    tray_app = SystemTrayIcon()
    tray_app.run()

