# Winget Updater

A Windows application that runs as a service and monitors available package updates using Microsoft's Winget package manager. The application displays a notification icon in the system tray, showing the number of available updates and providing notifications when new updates are available.

## DISCLAIMER
This application is made using Warp terminal AI agent. None of the code is done by me but is created using prompting as a test to see how well the AI agent could complete the task at hand. There are still some UI bugs like buttons not working as they should and I will try to fix this as son as possible but the application works more or less as it should.  The code is provided as-is and may contain bugs or issues. Use at your own risk.

## Features

- System tray icon showing the number of available updates
- Automated update checks at configurable times
- Notifications when new updates are available
- Detailed list of available updates
- Configurable settings via a GUI
- Ability to run as a Windows service or standalone application

## Prerequisites

- Windows 10 (1809+) or Windows 11
- [Python 3.8+](https://www.python.org/downloads/)
- [Microsoft's Winget](https://github.com/microsoft/winget-cli) package manager installed
- Administrator privileges (for service installation)

## Installation

### 1. Clone or Download the Repository

Download the files to your preferred location or clone the repository:

```
git clone https://github.com/yourusername/winget-updater.git
cd winget-updater
```

### 2. Install Required Python Packages

Install the required Python packages using pip:

```
pip install -r requirements.txt
```

### 3. Running the Application

#### As a Standalone Application

To run the application without installing it as a service:

```
python main.py --standalone
```

#### Using the VB Script Launcher (Recommended for Background Operation)

For a cleaner background experience without console windows, use the provided VB script:

```
launch_winget_updater.vbs
```

This script:
- Runs the Winget Updater in standalone mode
- Hides the console window for a cleaner experience
- Automatically sets the correct working directory
- Ensures proper environment variables are set

To use the VB script:
1. Double-click `launch_winget_updater.vbs` in File Explorer, or
2. Run from command line: `cscript launch_winget_updater.vbs` or `wscript launch_winget_updater.vbs`

**Note**: The VB script method is ideal for users who want to run the application in the background without seeing command prompt windows.

#### As a Windows Service

To install and run the application as a Windows service (requires Administrator privileges):

1. Install the service:
   ```
   python main.py --install
   ```

2. Start the service:
   ```
   python main.py --start
   ```

## Usage

### Service Management

The application provides several command-line options for service management:

- `--install`: Install the application as a Windows service
- `--uninstall`: Uninstall the Windows service
- `--start`: Start the Windows service
- `--stop`: Stop the Windows service
- `--restart`: Restart the Windows service
- `--standalone`: Run as a standalone application (not as a service)
- `--debug`: Enable debug logging

### System Tray Features

The system tray icon provides the following functionality:

- **Icon Color**: 
  - Blue: No updates available
  - Green: Updates available (with number overlay)

- **Context Menu**:
  - **Check for Updates**: Manually check for available Winget updates
  - **Settings**: Open the settings window
  - **Show Updates**: Display a list of available updates
  - **Exit**: Close the application

### Configuration Options

The settings window allows you to configure:

1. **Morning Check Time**: When to check for updates in the morning (format: HH:MM, 24-hour)
2. **Afternoon Check Time**: When to check for updates in the afternoon (format: HH:MM, 24-hour)
3. **Show Notifications**: Enable/disable notifications when updates are available
4. **Automatic Checking**: Enable/disable scheduled automatic checks

Settings are stored in a `settings.ini` file in the application directory.

## Troubleshooting

### Application Not Starting

- Verify that all dependencies are installed: `pip install -r requirements.txt`
- Check the log file (`winget_updater.log` or `winget_updater_service.log`) for error messages
- Ensure you have the required permissions (especially for service installation)

### Service Installation Issues

- Make sure you're running the command prompt as Administrator
- Check Windows Event Viewer for service-related errors
- Verify that the pywin32 package is correctly installed

### Winget Command Errors

- Ensure Winget is installed and accessible from the command line
- Try running `winget update` manually in a command prompt to verify it works
- Check if Windows App Installer is up to date

### No Updates Showing

- The application relies on Winget's output format, which may change with updates
- Check the log files for parsing errors
- Verify that Winget is correctly identifying updates

## Logs

Log files are stored in the application directory:

- `winget_updater.log`: Logs for standalone mode
- `winget_updater_service.log`: Logs for service mode

## License

[MIT License](LICENSE)

## Acknowledgements

- [Microsoft Winget](https://github.com/microsoft/winget-cli)
- [pystray](https://github.com/moses-palmer/pystray)
- [pywin32](https://github.com/mhammond/pywin32)

