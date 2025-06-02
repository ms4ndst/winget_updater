import os
import json
import time
import socket
import logging
import threading
from datetime import datetime
import win32pipe, win32file, win32api, win32event, pywintypes

# Constants for the IPC
PIPE_NAME = r'\\.\pipe\WingetUpdaterPipe'
BUFFER_SIZE = 4096

class IPCMessage:
    """Represents a message in the IPC protocol"""
    def __init__(self, command, data=None):
        self.command = command
        self.data = data if data is not None else {}
        self.timestamp = datetime.now().isoformat()
    
    def to_json(self):
        """Convert message to JSON string"""
        return json.dumps({
            'command': self.command,
            'data': self.data,
            'timestamp': self.timestamp
        })
    
    @classmethod
    def from_json(cls, json_str):
        """Create message from JSON string"""
        try:
            data = json.loads(json_str)
            msg = cls(data['command'], data['data'])
            msg.timestamp = data['timestamp']
            return msg
        except (json.JSONDecodeError, KeyError) as e:
            logging.error(f"Error parsing IPC message: {str(e)}")
            return None

class IPCServer:
    """Server-side IPC handler using named pipes"""
    
    def __init__(self):
        self.running = False
        self.pipe = None
        self.command_handlers = {}
        self.logger = logging.getLogger('IPCServer')
    
    def register_handler(self, command, handler_func):
        """Register a handler function for a specific command"""
        self.command_handlers[command] = handler_func
    
    def start(self):
        """Start the IPC server"""
        self.running = True
        self.server_thread = threading.Thread(target=self._run_server)
        self.server_thread.daemon = True
        self.server_thread.start()
        self.logger.info("IPC server started")
    
    def stop(self):
        """Stop the IPC server"""
        self.running = False
        if self.pipe:
            try:
                win32file.CloseHandle(self.pipe)
            except:
                pass
        self.logger.info("IPC server stopped")
    
    def _run_server(self):
        """Main server loop"""
        while self.running:
            try:
                # Create the named pipe
                self.pipe = win32pipe.CreateNamedPipe(
                    PIPE_NAME,
                    win32pipe.PIPE_ACCESS_DUPLEX,
                    win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
                    win32pipe.PIPE_UNLIMITED_INSTANCES,
                    BUFFER_SIZE,
                    BUFFER_SIZE,
                    0,
                    None
                )
                
                # Wait for a client to connect
                self.logger.debug("Waiting for client connection...")
                win32pipe.ConnectNamedPipe(self.pipe, None)
                self.logger.debug("Client connected")
                
                # Process client requests
                while self.running:
                    try:
                        # Read request
                        result, data = win32file.ReadFile(self.pipe, BUFFER_SIZE)
                        message_json = data.decode('utf-8')
                        
                        # Parse message
                        message = IPCMessage.from_json(message_json)
                        if not message:
                            continue
                            
                        # Handle the command
                        self.logger.debug(f"Received command: {message.command}")
                        response = self._handle_command(message)
                        
                        # Send response
                        response_json = response.to_json()
                        win32file.WriteFile(self.pipe, response_json.encode('utf-8'))
                        
                    except pywintypes.error as e:
                        if e.winerror == 109:  # Broken pipe
                            self.logger.debug("Client disconnected")
                            break
                        else:
                            self.logger.error(f"Pipe error: {str(e)}")
                            break
                    except Exception as e:
                        self.logger.error(f"Error processing request: {str(e)}")
                        # Try to send error response
                        try:
                            error_response = IPCMessage("error", {"message": str(e)})
                            win32file.WriteFile(self.pipe, error_response.to_json().encode('utf-8'))
                        except:
                            pass
                        break
                
                # Close the pipe
                try:
                    # Only disconnect if the pipe is still connected
                    # Note: DisconnectNamedPipe is not always available
                    # so we handle it gracefully
                    if hasattr(win32file, 'DisconnectNamedPipe'):
                        win32file.DisconnectNamedPipe(self.pipe)
                except Exception as e:
                    self.logger.debug(f"Error disconnecting pipe: {str(e)}")
                
                try:
                    win32file.CloseHandle(self.pipe)
                except Exception as e:
                    self.logger.debug(f"Error closing pipe handle: {str(e)}")
                    
                self.pipe = None
                
            except Exception as e:
                self.logger.error(f"Server error: {str(e)}")
                time.sleep(1)  # Avoid rapid retries
    
    def _handle_command(self, message):
        """Handle a command message"""
        command = message.command
        
        # Check if we have a handler for this command
        if command in self.command_handlers:
            try:
                # Call the handler and get result
                result = self.command_handlers[command](message.data)
                return IPCMessage("response", result)
            except Exception as e:
                self.logger.error(f"Error in command handler for {command}: {str(e)}")
                return IPCMessage("error", {"message": str(e)})
        else:
            return IPCMessage("error", {"message": f"Unknown command: {command}"})

class IPCClient:
    """Client-side IPC handler using named pipes"""
    
    def __init__(self):
        self.pipe = None
        self.logger = logging.getLogger('IPCClient')
    
    def connect(self, timeout=10):
        """Connect to the IPC server"""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                # Try to open the pipe
                self.pipe = win32file.CreateFile(
                    PIPE_NAME,
                    win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                    0,
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None
                )
                
                # Set the read mode to message
                win32pipe.SetNamedPipeHandleState(
                    self.pipe, 
                    win32pipe.PIPE_READMODE_MESSAGE,
                    None,
                    None
                )
                
                self.logger.info("Connected to IPC server")
                return True
                
            except pywintypes.error as e:
                if e.winerror == 2:  # File not found
                    # Server isn't running yet, wait and retry
                    time.sleep(0.5)
                else:
                    self.logger.error(f"Error connecting to IPC server: {str(e)}")
                    return False
        
        self.logger.error(f"Timeout connecting to IPC server")
        return False
    
    def disconnect(self):
        """Disconnect from the IPC server"""
        if self.pipe:
            try:
                win32file.CloseHandle(self.pipe)
                self.pipe = None
                self.logger.info("Disconnected from IPC server")
            except Exception as e:
                self.logger.error(f"Error disconnecting from IPC server: {str(e)}")
    
    def send_command(self, command, data=None):
        """Send a command to the server and get the response"""
        if not self.pipe:
            if not self.connect():
                return None
        
        try:
            # Create and send the message
            message = IPCMessage(command, data)
            message_json = message.to_json()
            win32file.WriteFile(self.pipe, message_json.encode('utf-8'))
            
            # Read the response
            result, data = win32file.ReadFile(self.pipe, BUFFER_SIZE)
            response_json = data.decode('utf-8')
            
            # Parse and return the response
            response = IPCMessage.from_json(response_json)
            return response
            
        except pywintypes.error as e:
            self.logger.error(f"Error in IPC communication: {str(e)}")
            self.disconnect()
            return None
        except Exception as e:
            self.logger.error(f"Error sending command: {str(e)}")
            return None

