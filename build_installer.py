import os
import sys
import subprocess
from pathlib import Path

def create_pyinstaller_spec():
    """Create PyInstaller spec file for the application"""
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('settings.ini', '.'),
        ('README.md', '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='WingetUpdater',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='winget_updater.ico',  # Will be created if not exists
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WingetUpdater',
)
"""
    with open('winget_updater.spec', 'w') as f:
        f.write(spec_content)
    print("Created PyInstaller spec file: winget_updater.spec")

def create_default_icon():
    """Create a default icon if none exists"""
    if not os.path.exists('winget_updater.ico'):
        try:
            from PIL import Image, ImageDraw
            
            # Create a 64x64 icon
            img = Image.new('RGBA', (64, 64), color=(0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # Draw a simple W with update arrow
            draw.rectangle([0, 0, 63, 63], fill=(30, 150, 30, 255))  # Green background
            
            # Draw a W
            draw.line([(10, 10), (10, 54)], fill=(255, 255, 255, 255), width=3)  # Left line
            draw.line([(10, 54), (32, 32)], fill=(255, 255, 255, 255), width=3)  # Bottom to middle
            draw.line([(32, 32), (54, 54)], fill=(255, 255, 255, 255), width=3)  # Middle to bottom
            draw.line([(54, 54), (54, 10)], fill=(255, 255, 255, 255), width=3)  # Right line
            
            # Save as ICO
            img.save('winget_updater.ico', format='ICO')
            print("Created default icon: winget_updater.ico")
        except ImportError:
            print("WARNING: Pillow (PIL) library not installed. Cannot create default icon.")
            print("Please install it with: pip install pillow")
            print("Or create your own icon file named 'winget_updater.ico'")

def create_inno_setup_script():
    """Create Inno Setup script for the installer"""
    script_content = """#define MyAppName "Winget Updater"
#define MyAppVersion "1.0"
#define MyAppPublisher "Visma IT"
#define MyAppExeName "WingetUpdater.exe"

[Setup]
AppId={{891E2B26-53C5-4371-A7BF-95A6AE3E944F}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=README.md
OutputDir=installer_output
OutputBaseFilename=WingetUpdater_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startup"; Description: "Start with Windows"; GroupDescription: "Windows Startup"

[Files]
Source: "dist\\WingetUpdater\\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"
Name: "{autodesktop}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userstartup}\\{#MyAppName}"; Filename: "{app}\\{#MyAppExeName}"; Tasks: startup

[Run]
Filename: "{app}\\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
"""
    os.makedirs('installer_output', exist_ok=True)
    with open('winget_updater.iss', 'w') as f:
        f.write(script_content)
    print("Created Inno Setup script: winget_updater.iss")

def check_requirements():
    """Check if all required packages are installed"""
    try:
        import PyInstaller
        print("PyInstaller is installed.")
    except ImportError:
        print("PyInstaller is not installed. Installing...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
            print("PyInstaller installed successfully.")
        except Exception as e:
            print(f"Failed to install PyInstaller: {e}")
            return False
    
    return True

def build_executable():
    """Build the executable using PyInstaller"""
    try:
        print("\nBuilding executable with PyInstaller...\n")
        subprocess.run(['pyinstaller', 'winget_updater.spec'], check=True)
        print("\nExecutable built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error building executable: {e}")
        return False

def main():
    print("Winget Updater - Build Script")
    print("=============================\n")
    
    # Check requirements
    if not check_requirements():
        print("Failed to satisfy requirements. Exiting.")
        return
    
    # Create default icon if needed
    create_default_icon()
    
    # Create PyInstaller spec
    create_pyinstaller_spec()
    
    # Build executable
    if not build_executable():
        print("Failed to build executable. Exiting.")
        return
    
    # Create Inno Setup script
    create_inno_setup_script()
    
    print("\nBuild process completed!")
    print("\nTo create the installer:")
    print("1. Download and install Inno Setup from: https://jrsoftware.org/isdl.php")
    print("2. Open winget_updater.iss with Inno Setup Compiler")
    print("3. Click Build > Compile")
    print("\nThe installer will be created in the 'installer_output' directory")

if __name__ == "__main__":
    main()

