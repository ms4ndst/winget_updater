import subprocess
import re
import json
import asyncio
import logging
import os
from datetime import datetime

class WingetUpdateChecker:
    """Class to check for available updates using Winget"""
    
    def __init__(self, config_manager=None):
        self.config_manager = config_manager
        self.available_updates = []
        self.update_count = 0
        self.last_check_time = None
        self.is_checking = False
        self.pinned_packages = set()  # Cache for pinned packages
        self._setup_logging()
    
    def _setup_logging(self):
        """Set up logging for the update checker"""
        # Only configure if not already configured
        if not logging.getLogger().handlers:
            logging.basicConfig(
                filename='winget_updater.log',
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
        self.logger = logging.getLogger('WingetUpdateChecker')
        
        # Make sure we have the re module for parsing
        import re
    
    async def check_updates_async(self):
        """Asynchronously check for updates using Winget"""
        if self.is_checking:
            self.logger.info("Update check already in progress, skipping")
            return self.update_count
            
        self.is_checking = True
        self.logger.info("Starting asynchronous update check")
        
        try:
            # Create a subprocess to run the winget update command
            process = await asyncio.create_subprocess_exec(
                'winget', 'update', 
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            
            # Capture the output
            stdout, stderr = await process.communicate()
            
            if process.returncode != 0:
                self.logger.error(f"Winget update command failed with return code {process.returncode}")
                self.logger.error(f"Error: {stderr.decode('utf-8')}")
                self.is_checking = False
                return 0
                
            # Process the output
            output = stdout.decode('utf-8')
            self._parse_winget_output(output)
            
            # Update the last check time in config if available
            if self.config_manager:
                self.config_manager.set_last_check()
                
            self.last_check_time = datetime.now()
            self.logger.info(f"Update check completed. Found {self.update_count} updates.")
            
            return self.update_count
            
        except Exception as e:
            self.logger.error(f"Error checking for updates: {str(e)}")
            return 0
        finally:
            self.is_checking = False
    
    def check_updates(self, force=False, include_pinned=None, include_unknown=None):
        """
        Synchronously check for updates using Winget
        
        Args:
            force: If True, force a new check even if one is in progress
            include_pinned: If True, include pinned packages in the results
            include_unknown: If True, include packages with unknown versions
        
        Returns:
            int: The number of available updates
        """
        if self.is_checking and not force:
            self.logger.info("Update check already in progress, skipping")
            return self.update_count
        
        # Get settings from config_manager if not explicitly provided
        if self.config_manager:
            if include_pinned is None:
                include_pinned = self.config_manager.get_include_pinned_updates()
            if include_unknown is None:
                include_unknown = self.config_manager.get_include_unknown_versions()
        
        # Log the filtering settings being used
        self.logger.debug(f"Update check settings - include_pinned: {include_pinned}, include_unknown: {include_unknown}")
            
        self.is_checking = True
        self.logger.info("Starting synchronous update check")
        
        # Always clear cached updates when starting a fresh check
        self.available_updates = []
        self.update_count = 0
        
        # Refresh the list of pinned packages
        self._refresh_pinned_packages()
        
        try:
            # Define base command with all necessary flags to show all updates
            base_command = [
                'winget', 'update',
                '--include-unknown',
                '--include-pinned',
                '--accept-source-agreements'
            ]
            
            # First try to use JSON format for more reliable parsing
            result = self._check_updates_json(base_command, include_pinned, include_unknown)
            if result is not None:
                return result
                
            # Fallback to text format if JSON fails
            self.logger.info("JSON format failed, falling back to text format")
            
            # Run the winget update command with all updates visible
            process = subprocess.run(
                base_command, 
                capture_output=True, 
                text=True, 
                check=False
            )
            
            if process.returncode != 0:
                self.logger.error(f"Winget update command failed with return code {process.returncode}")
                self.logger.error(f"Error: {process.stderr}")
                self.is_checking = False
                return 0
                
            # Process the output
            self._parse_winget_output(process.stdout, include_pinned, include_unknown)
            
            # Update the last check time in config if available
            if self.config_manager:
                self.config_manager.set_last_check()
                
            self.last_check_time = datetime.now()
            self.logger.info(f"Update check completed. Found {self.update_count} updates.")
            
            return self.update_count
            
        except Exception as e:
            self.logger.error(f"Error checking for updates: {str(e)}")
            return 0
        finally:
            self.is_checking = False
    
    def _check_updates_json(self, base_command=None, include_pinned=False, include_unknown=False):
        """
        Check for updates using Winget's JSON output format
        
        Args:
            base_command: Base command to use for checking updates
            include_pinned: If True, include pinned packages in the results
            include_unknown: If True, include packages with unknown versions
        """
        try:
            self.logger.debug("Attempting to use JSON output format")
            
            if base_command is None:
                base_command = ['winget', 'update', '--accept-source-agreements', '--include-unknown', '--include-pinned']
                
            # Try different command variants
            commands = [
                base_command + ['--format', 'json'],
                ['winget', 'update', '--format', 'json'],
                ['winget', 'upgrade', '--format', 'json'],  # Some versions use 'upgrade'
                ['winget', 'update', '--accept-source-agreements', '--include-unknown', '--include-pinned', '--source', 'winget', '--format', 'json']
            ]
            
            for cmd in commands:
                try:
                    self.logger.debug(f"Trying command: {' '.join(cmd)}")
                    process = subprocess.run(
                        cmd,
                        capture_output=True, 
                        text=True, 
                        check=False
                    )
                    
                    if process.returncode == 0:
                        self.logger.debug(f"Command {' '.join(cmd)} succeeded")
                        return self._parse_winget_json(process.stdout, include_pinned, include_unknown)
                        
                    self.logger.debug(f"Command {' '.join(cmd)} failed with code {process.returncode}")
                    if process.stderr:
                        self.logger.debug(f"Error output: {process.stderr[:200]}")
                    
                except Exception as e:
                    self.logger.debug(f"Command {' '.join(cmd)} failed: {str(e)}")
                    continue
            
            self.logger.warning("All JSON format attempts failed")
            return None
                
        except Exception as e:
            self.logger.warning(f"Error using JSON format: {str(e)}")
            return None
    
    def _parse_winget_json(self, output, include_pinned=False, include_unknown=False):
        """
        Parse JSON output from winget update command
        
        Args:
            output: JSON output from winget
            include_pinned: If True, include pinned packages in the results
            include_unknown: If True, include packages with unknown versions
        """
        try:
            # Reset the updates list
            self.available_updates = []
            
            # Parse JSON
            data = json.loads(output)
            self.logger.debug(f"Successfully parsed JSON output: {json.dumps(data, indent=2)[:200]}...")
            
            # Extract updates from the data structure
            # Structure varies by winget version, so try different approaches
            
            # Try to find a "Sources" or "Data" array
            if "Sources" in data:
                packages = []
                for source in data["Sources"]:
                    if "Packages" in source:
                        packages.extend(source["Packages"])
                        
                for pkg_id, pkg_info in packages.items():
                    if "Version" in pkg_info and "AvailableVersion" in pkg_info:
                        current_version = pkg_info.get("Version", "Unknown")
                        available_version = pkg_info.get("AvailableVersion", "Unknown")
                        
                        # Skip packages with unknown versions unless explicitly included
                        if (current_version == "Unknown" or not current_version) and not include_unknown:
                            self.logger.debug(f"Skipping package {pkg_id} due to unknown current version (set include_unknown=True to include)")
                            continue
                        elif current_version == "Unknown" or not current_version:
                            self.logger.debug(f"Including package {pkg_id} with unknown version as requested")
                            
                        # Skip if version comparison is unreliable
                        if not self._is_valid_version_comparison(current_version, available_version):
                            self.logger.debug(f"Skipping package {pkg_id} due to unreliable version comparison: {current_version} -> {available_version}")
                            continue
                            
                        # Skip pinned packages unless explicitly included
                        if not include_pinned and self._is_package_pinned(pkg_id):
                            self.logger.debug(f"Skipping pinned package {pkg_id} - set include_pinned=True to include it")
                            continue
                            
                        if current_version != available_version:
                            package_info = {
                                'name': pkg_info.get("Name", pkg_id),
                                'id': pkg_id,
                                'current_version': current_version,
                                'available_version': available_version
                            }
                            self.available_updates.append(package_info)
            
            # Alternative structure for newer winget versions
            elif "Data" in data:
                for item in data["Data"]:
                    if all(k in item for k in ["Name", "Id", "Version", "AvailableVersion"]):
                        current_version = item["Version"]
                        available_version = item["AvailableVersion"]
                        
                        # Skip packages with unknown versions unless explicitly included
                        if (current_version == "Unknown" or not current_version) and not include_unknown:
                            self.logger.debug(f"Skipping package {item['Id']} due to unknown current version (set include_unknown=True to include)")
                            continue
                        elif current_version == "Unknown" or not current_version:
                            self.logger.debug(f"Including package {item['Id']} with unknown version as requested")
                            
                        # Skip if version comparison is unreliable
                        if not self._is_valid_version_comparison(current_version, available_version):
                            self.logger.debug(f"Skipping package {item['Id']} due to unreliable version comparison: {current_version} -> {available_version}")
                            continue
                            
                        # Skip pinned packages unless explicitly included
                        if not include_pinned and self._is_package_pinned(item['Id']):
                            self.logger.debug(f"Skipping pinned package {item['Id']} - set include_pinned=True to include it")
                            continue
                        
                        if current_version != available_version:
                            package_info = {
                                'name': item["Name"],
                                'id': item["Id"],
                                'current_version': current_version,
                                'available_version': available_version
                            }
                            self.available_updates.append(package_info)
            
            self.update_count = len(self.available_updates)
            self.logger.info(f"Parsed winget JSON output, found {self.update_count} updates")
            
            # Update the last check time in config if available
            if self.config_manager:
                self.config_manager.set_last_check()
                
            self.last_check_time = datetime.now()
            
            return self.update_count
            
        except json.JSONDecodeError as e:
            self.logger.warning(f"Failed to parse JSON: {str(e)}")
            return None
        except Exception as e:
            self.logger.warning(f"Error processing JSON output: {str(e)}")
            return None
    
    def _parse_winget_output(self, output, include_pinned=False, include_unknown=False):
        """
        Parse the output from winget update command
        
        Args:
            output: Text output from winget
            include_pinned: If True, include pinned packages in the results
            include_unknown: If True, include packages with unknown versions
        """
        # Reset the updates list - always start fresh
        self.available_updates = []
        
        # Split the output into lines
        lines = output.strip().split('\n')
        
        # Debug output
        if len(lines) > 0:
            self.logger.debug(f"First few lines of output: {lines[:min(5, len(lines))]}")
        
        # Check for common no-updates patterns
        for line in lines:
            if 'No updates found.' in line or 'No available upgrades.' in line:
                self.logger.info("No updates available according to winget")
                self.update_count = 0
                return
                
        # Define section markers to identify different parts of the output
        section_markers = [
            "The following packages have an upgrade available, but require explicit targeting for upgrade:",
            "have version numbers that cannot be determined",
            "have pins that prevent upgrade",
        ]
        
        # Split output into sections
        sections = self._split_output_into_sections(lines, section_markers)
        self.logger.debug(f"Found {len(sections)} sections in winget output")
        
        # Process each section
        for section_idx, section_lines in enumerate(sections):
            if not section_lines:
                continue
                
            self.logger.debug(f"Processing section {section_idx} with {len(section_lines)} lines")
            self._process_output_section(section_lines, include_pinned, include_unknown)
            
        self.update_count = len(self.available_updates)
        self.logger.info(f"Parsed winget output, found {self.update_count} updates")
    
    def _split_output_into_sections(self, lines, section_markers):
        """Split winget output into sections based on section markers"""
        sections = []
        current_section = []
        in_first_section = True
        
        for line in lines:
            # Check if this line is a section marker
            is_marker = any(marker in line for marker in section_markers)
            
            # If we find a section marker, start a new section
            if is_marker:
                if current_section:
                    sections.append(current_section)
                current_section = [line]
                in_first_section = False
            # Or if we're in the first section and find a header line
            elif in_first_section and self._is_header_line(line):
                if current_section:
                    sections.append(current_section)
                current_section = [line]
            # Otherwise add to current section
            else:
                current_section.append(line)
        
        # Add the last section if it exists
        if current_section:
            sections.append(current_section)
            
        return sections
    
    def _is_header_line(self, line):
        """Check if a line is a header line"""
        # Only consider it a header if it's the exact standard header format
        standard_headers = [
            "Name                   Id                    Version     Available   Source",
            "Name  Id          Version  Available Source",
            "-----------------------------------------------------------------------"
        ]
        
        # Check for exact matches in standard headers
        if line.strip() in standard_headers:
            return True
            
        # More precise pattern matching for headers
        header_patterns = [
            # Standard English pattern with exact word matching
            lambda l: l.strip().startswith("Name") and " Id " in l and " Version " in l and " Available " in l,
            # Other standard header patterns
            lambda l: l.strip().startswith("Package") and " ID " in l and " Version " in l and " Available " in l
        ]
        
        return any(pattern(line) for pattern in header_patterns)
    
    def _should_skip_line(self, line):
        """Check if a line should be skipped"""
        return (not line.strip() or 
                all(c in '-' for c in line.strip()) or
                self._is_header_line(line) or
                any(x in line.lower() for x in [
                    'upgrades available',
                    'upgrade available',
                    'no updates found',
                    'package(s) have',
                    'prevent upgrade',
                    'explicit targeting'
                ]))
                
    def _process_output_section(self, section_lines, include_pinned=False, include_unknown=False):
        """
        Process a section of winget output
        
        Args:
            section_lines: Lines of text output from a section of winget output
            include_pinned: If True, include pinned packages in the results
            include_unknown: If True, include packages with unknown versions
        """
        # Skip empty sections
        if not section_lines:
            return
            
        for line in section_lines:
            # Skip lines we don't want to process
            if self._should_skip_line(line):
                continue

            try:
                # Split by multiple spaces but preserve name parts
                parts = [p.strip() for p in re.split(r'\s{2,}', line.strip())]
                
                if len(parts) >= 4:
                    # First try to extract version from name if present
                    name = parts[0]
                    id_str = parts[1]
                    version = parts[2]
                    available = parts[3]
                    
                    # Clean up version numbers
                    version_match = re.search(r'\d+(\.\d+)+', version)
                    if version_match:
                        version = version_match.group(0)
                        
                    available_match = re.search(r'\d+(\.\d+)+', available)
                    if available_match:
                        available = available_match.group(0)
                    
                    # Clean up name
                    name = re.sub(r'\s*\([^)]*\)', '', name)  # Remove parentheses
                    name = re.sub(r'\s+\d+(\.\d+)+$', '', name)  # Remove version number if present
                    name = name.strip()
                    
                    # Skip packages with unknown versions unless explicitly included
                    if (version.lower() == "unknown" or not version) and not include_unknown:
                        self.logger.debug(f"Skipping package {id_str} due to unknown current version (set include_unknown=True to include)")
                        continue
                    elif version.lower() == "unknown" or not version:
                        self.logger.debug(f"Including package {id_str} with unknown version as requested")
                        
                    # Skip if version comparison is unreliable
                    if not self._is_valid_version_comparison(version, available):
                        self.logger.debug(f"Skipping package {id_str} due to unreliable version comparison: {version} -> {available}")
                        continue
                        
                    # Skip pinned packages unless explicitly included
                    if not include_pinned and self._is_package_pinned(id_str):
                        self.logger.debug(f"Skipping pinned package {id_str} - set include_pinned=True to include it")
                        continue
                    
                    # Only add if we have a valid update
                    if all([name, id_str, version, available]) and version != available:
                        package_info = {
                            'name': name,
                            'id': id_str,
                            'current_version': version,
                            'available_version': available,
                            'source': 'winget'
                        }
                        self.available_updates.append(package_info)
                        self.logger.debug(f"Found update: {name} - {version} -> {available}")
                    
            except Exception as e:
                self.logger.debug(f"Could not parse line as update: {line.strip()} - Error: {str(e)}")
                continue
    
    def _refresh_pinned_packages(self):
        """Refresh the list of pinned packages"""
        self.pinned_packages = set()
        try:
            # Run winget pin list command
            process = subprocess.run(
                ['winget', 'pin', 'list'],
                capture_output=True,
                text=True,
                check=False
            )
            
            if process.returncode == 0:
                # Parse the output to find pinned package IDs
                lines = process.stdout.strip().split('\n')
                for line in lines:
                    # Skip header lines and separators
                    if not line.strip() or '---' in line or 'Name' in line and 'Id' in line:
                        continue
                        
                    # Split by multiple spaces to get package ID
                    parts = [p.strip() for p in re.split(r'\s{2,}', line.strip())]
                    if len(parts) >= 2:
                        self.pinned_packages.add(parts[1])  # The ID is typically the second column
                
                self.logger.debug(f"Found {len(self.pinned_packages)} pinned packages: {', '.join(self.pinned_packages)}")
        except Exception as e:
            self.logger.warning(f"Failed to get list of pinned packages: {str(e)}")
    
    def _is_package_pinned(self, package_id):
        """Check if a package is pinned"""
        return package_id in self.pinned_packages

    def get_updates_list(self, include_pinned=None, include_unknown=None):
        """
        Get the list of available updates
        
        Args:
            include_pinned: If True, include pinned packages in the results
            include_unknown: If True, include packages with unknown versions
        
        Returns:
            list: List of available updates
        """
        # Get settings from config_manager if not explicitly provided
        if self.config_manager:
            if include_pinned is None:
                include_pinned = self.config_manager.get_include_pinned_updates()
            if include_unknown is None:
                include_unknown = self.config_manager.get_include_unknown_versions()
        
        # Force a fresh check if no updates are cached
        if not self.available_updates:
            self.check_updates(force=True, include_pinned=include_pinned, include_unknown=include_unknown)
        return self.available_updates
    
    def get_update_count(self, include_pinned=None, include_unknown=None):
        """
        Get the count of available updates
        
        Args:
            include_pinned: If True, include pinned packages in the count
            include_unknown: If True, include packages with unknown versions
        
        Returns:
            int: Number of available updates
        """
        # Get settings from config_manager if not explicitly provided
        if self.config_manager:
            if include_pinned is None:
                include_pinned = self.config_manager.get_include_pinned_updates()
            if include_unknown is None:
                include_unknown = self.config_manager.get_include_unknown_versions()
                
        # Force a fresh check if no updates are cached
        if self.update_count == 0:
            self.check_updates(force=True, include_pinned=include_pinned, include_unknown=include_unknown)
        return self.update_count
    
    def _is_valid_version_comparison(self, current_version, available_version):
        """
        Validates that comparing the versions is reliable
        Returns True if the comparison is valid, False otherwise
        """
        # Check for empty or unknown versions
        if (not current_version or 
            not available_version or 
            current_version.lower() == "unknown" or 
            available_version.lower() == "unknown"):
            return False
            
        # Check for non-comparable versions (e.g., strings without numbers)
        has_numbers_current = any(char.isdigit() for char in current_version)
        has_numbers_available = any(char.isdigit() for char in available_version)
        
        if not has_numbers_current or not has_numbers_available:
            return False
            
        # Attempt to normalize version strings
        try:
            # Extract numeric parts for basic comparison
            current_nums = [int(part) for part in re.findall(r'\d+', current_version)]
            available_nums = [int(part) for part in re.findall(r'\d+', available_version)]
            
            # If both versions have numeric parts, consider it valid
            if current_nums and available_nums:
                return True
                
        except Exception:
            pass
            
        return False
        
    def get_last_check_time(self):
        """Get the time of the last update check"""
        return self.last_check_time

    def install_all_updates(self):
        """Install all available updates using Winget"""
        self.logger.info("Starting installation of all updates")
        
        # Force a fresh check to ensure we have current updates
        # Include pinned packages and unknown versions since we want to show everything that can be installed
        self.check_updates(force=True, include_pinned=True, include_unknown=True)
        
        if not self.available_updates:
            self.logger.info("No updates to install")
            return True
            
        # Store the list of updates we're about to install for verification
        updates_to_install = self.available_updates.copy()
        
        try:
            # Run winget upgrade --all command
            self.logger.info(f"Installing {len(self.available_updates)} updates")
            
            # Define installation command with all necessary flags
            command = [
                'winget', 'upgrade', '--all',
                '--accept-source-agreements',
                '--disable-interactivity'
            ]
            
            self.logger.debug(f"Running command: {' '.join(command)}")
            
            # Execute the installation command
            process = subprocess.run(
                command, 
                capture_output=True, 
                text=True,
                check=False
            )
            
            # Log output for debugging
            self.logger.debug(f"Installation stdout: {process.stdout[:500]}")
            if process.stderr:
                self.logger.debug(f"Installation stderr: {process.stderr[:500]}")
            
            # Check if the process was successful
            if process.returncode == 0:
                self.logger.info("Installation completed successfully")
                
                # Force a fresh check to verify updates were installed
                self.logger.info("Verifying updates were installed...")
                self.check_updates(force=True)
                
                # Check if any of the updates we tried to install are still pending
                still_pending = []
                for update in updates_to_install:
                    update_id = update['id']
                    for current_update in self.available_updates:
                        if current_update['id'] == update_id:
                            still_pending.append(update_id)
                            break
                
                if still_pending:
                    self.logger.warning(f"Some updates failed to install: {', '.join(still_pending)}")
                    return False
                else:
                    self.logger.info("All updates were successfully installed and verified")
                    return True
            else:
                self.logger.error(f"Installation failed with return code {process.returncode}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error during installation: {str(e)}")
            return False
