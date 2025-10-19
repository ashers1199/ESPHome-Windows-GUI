import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess
from subprocess import Popen, PIPE, TimeoutExpired
import threading
import os
import shutil
import sys
import serial.tools.list_ports
import requests
import json
import re
import time
import webbrowser
import socket
from packaging import version
from datetime import datetime
from zeroconf import Zeroconf, ServiceBrowser, ServiceListener
from collections import defaultdict
from datetime import datetime

# NEW IMPORTS for version management
import venv
from pathlib import Path
import glob
import hashlib

# Import ttkbootstrap
import ttkbootstrap as tb
from ttkbootstrap.constants import *

class ESPHomeListener(ServiceListener):
    def __init__(self):
        self.devices = []  # List of (name, ip) tuples

    def add_service(self, zeroconf, type, name):
        info = zeroconf.get_service_info(type, name)
        if info and info.addresses:
            ip = ".".join(str(b) for b in info.addresses[0])
            device_name = name.split('.')[0]  # Extract name from full mDNS name
            self.devices.append((device_name, ip))
            
def discover_esphome_devices():
    listener = ESPHomeListener()
    zeroconf = Zeroconf()

    ServiceBrowser(zeroconf, "_esphome._tcp.local.", listener)
    ServiceBrowser(zeroconf, "_esphomelib._tcp.local.", listener)

    time.sleep(3)
    zeroconf.close()
    return listener.devices  # Returns list of (name, ip)

# NEW - Enhanced file sync function
def sync_esphome_files(network_path, local_path, backup_path=None):
    """Sync YAML files and resources (images, fonts, etc.) and return list of synced files"""
    synced_files = []
    try:
        # Create local directory if it doesn't exist
        os.makedirs(local_path, exist_ok=True)
        
        # Define file patterns to sync
        sync_patterns = [
            "*.yaml",
            "*.yml", 
            "*.png", "*.jpg", "*.jpeg", "*.bmp", "*.gif",  # Images
            "*.ttf", "*.otf",  # Fonts
            "*.bin",  # Binary files
            "*.txt", "*.md",  # Text files
        ]
        
        # Sync files matching patterns
        for pattern in sync_patterns:
            pattern_path = os.path.join(network_path, pattern)
            for src_file in glob.glob(pattern_path):
                if os.path.isfile(src_file):
                    filename = os.path.basename(src_file)
                    dst = os.path.join(local_path, filename)
                    
                    # Create backup before overwriting (if backup path provided)
                    if backup_path and os.path.exists(dst):
                        if filename.lower().endswith((".yaml", ".yml")):
                            backup_file = create_backup(dst, backup_path, filename)
                            if backup_file:
                                print(f"Backed up: {backup_file}")

                    # Only copy if source is newer or destination doesn't exist
                    if not os.path.exists(dst) or os.path.getmtime(src_file) > os.path.getmtime(dst):
                        shutil.copy2(src_file, dst)
                        synced_files.append(filename)
                        print(f"Synced: {filename}")
        
        # Also sync subdirectories (for organized resources)
        for item in os.listdir(network_path):
            item_path = os.path.join(network_path, item)
            if os.path.isdir(item_path):
                local_subdir = os.path.join(local_path, item)
                os.makedirs(local_subdir, exist_ok=True)
                
                # Recursively sync subdirectory contents
                subdir_synced = sync_esphome_files(item_path, local_subdir, backup_path)
                synced_files.extend([f"{item}/{f}" for f in subdir_synced])
        
        return synced_files
    except Exception as e:
        print(f"Sync failed: {e}")
        return []

def sync_esphome_files_fast(network_path, local_path, backup_path=None, yaml_file=None):
    """Fast sync - only sync files needed for the current YAML"""
    synced_files = []
    try:
        os.makedirs(local_path, exist_ok=True)
        
        # If we have a specific YAML file, only sync referenced files
        if yaml_file and os.path.exists(yaml_file):
            # Get files referenced in this YAML
            referenced_files = get_referenced_files(yaml_file)
            referenced_files.append(os.path.basename(yaml_file))  # Always sync the main YAML
            
            # Also sync common directories that might be referenced
            common_dirs = ['images', 'fonts', 'binaries', 'scripts']
            for dir_name in common_dirs:
                dir_path = os.path.join(network_path, dir_name)
                if os.path.exists(dir_path):
                    # Add all files from common directories (they're usually small)
                    for file in os.listdir(dir_path):
                        if any(file.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.ttf', '.otf', '.bin']):
                            referenced_files.append(os.path.join(dir_name, file))
            
            # Remove duplicates
            referenced_files = list(set(referenced_files))
            
            for file_ref in referenced_files:
                # Handle files in subdirectories
                src = os.path.join(network_path, file_ref)
                dst = os.path.join(local_path, file_ref)
                
                # Create destination directory if needed
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                
                if os.path.exists(src) and os.path.isfile(src):
                    # Create backup if it's a YAML file and backup is enabled
                    if (backup_path and 
                        file_ref.lower().endswith(('.yaml', '.yml')) and 
                        os.path.exists(dst)):
                        backup_file = create_backup(dst, backup_path, os.path.basename(file_ref))
                        if backup_file:
                            print(f"Backed up: {backup_file}")
                    
                    # Only copy if source is newer or destination doesn't exist
                    if not os.path.exists(dst) or os.path.getmtime(src) > os.path.getmtime(dst):
                        shutil.copy2(src, dst)
                        synced_files.append(file_ref)
                        print(f"Synced: {file_ref}")
            
            return synced_files
        
        # Fallback to full sync if no specific YAML
        return sync_esphome_files(network_path, local_path, backup_path)
        
    except Exception as e:
        print(f"Fast sync failed: {e}")
        # Fallback to full sync
        return sync_esphome_files(network_path, local_path, backup_path)

def get_referenced_files(yaml_file):
    """Extract referenced files from YAML configuration"""
    referenced_files = []
    try:
        with open(yaml_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Patterns for common file references in ESPHome
        patterns = [
            r'image:\s*[\'"]([^\'"]+\.(?:png|jpg|jpeg|bmp|gif))[\'"]',
            r'font:\s*[\'"]([^\'"]+\.(?:ttf|otf))[\'"]',
            r'file:\s*[\'"]([^\'"]+\.(?:bin|txt))[\'"]',
            r'filename:\s*[\'"]([^\'"]+\.(?:bin|txt))[\'"]',
            r'source:\s*[\'"]([^\'"]+\.(?:bin|png|jpg|jpeg))[\'"]',
            r'uri:\s*file://([^\'"]+)',  # Local file URIs
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, content, re.IGNORECASE)
            referenced_files.extend(matches)
        
        # Also check for includes/substitutions that might reference other YAMLs
        include_patterns = [
            r'!include\s+[\'"]([^\'"]+\.(?:yaml|yml))[\'"]',
            r'substitutions:\s*!include\s+[\'"]([^\'"]+\.(?:yaml|yml))[\'"]',
        ]
        
        for pattern in include_patterns:
            matches = re.findall(pattern, content)
            referenced_files.extend(matches)
        
        return list(set(referenced_files))  # Remove duplicates
        
    except Exception as e:
        print(f"Error parsing referenced files: {e}")
        return []

# NEW - Backup functionality
def create_backup(file_path, backup_dir, original_name=None):
    """Create a timestamped backup of a YAML file inside its own subfolder"""
    try:
        os.makedirs(backup_dir, exist_ok=True)

        if original_name is None:
            original_name = os.path.basename(file_path)

        # Create subfolder for this file (e.g. backup_dir/light.yaml/)
        subdir = os.path.join(backup_dir, original_name)
        os.makedirs(subdir, exist_ok=True)

        # Create backup filename with original_name + timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{original_name}.{timestamp}"  # e.g. light.yaml.20251002_093344
        backup_path = os.path.join(subdir, backup_name)

        shutil.copy2(file_path, backup_path)
        return backup_path
    except Exception as e:
        print(f"Backup failed for {file_path}: {e}")
        return None

def cleanup_old_backups(backup_dir, max_backups=10):
    """
    Cleanup tiered backups in subdirectories per YAML file:
      - Keep last N most recent
      - Keep first of each day (suffix .daily)
      - Keep first of each week (suffix .weekly)
      - Keep first of each month (suffix .monthly)
      - Keep first of each year (suffix .yearly)
      - Always keep the oldest backup (suffix .baseline)
    """
    tags = {"daily", "weekly", "monthly", "yearly", "baseline"}

    try:
        if not os.path.exists(backup_dir):
            return

        for subdir in os.listdir(backup_dir):
            subpath = os.path.join(backup_dir, subdir)
            if not os.path.isdir(subpath):
                continue

            backups = []
            for file in os.listdir(subpath):
                parts = file.split(".")
                if len(parts) < 3:
                    continue
                # Handle files that may already have a tag suffix
                ts_part = parts[-1]
                if ts_part in tags and len(parts) >= 4:
                    ts_part = parts[-2]
                try:
                    dt = datetime.strptime(ts_part, "%Y%m%d_%H%M%S")
                except ValueError:
                    continue
                backups.append((os.path.join(subpath, file), dt))

            if not backups:
                continue

            backups.sort(key=lambda x: x[1], reverse=True)
            keep = {}

            # Keep rolling N
            for p, _ in backups[:max_backups]:
                keep[p] = None

            oldest_first = sorted(backups, key=lambda x: x[1])

            # Daily
            seen_days = set()
            for path, dt in oldest_first:
                day_key = dt.strftime("%Y-%m-%d")
                if day_key not in seen_days:
                    seen_days.add(day_key)
                    keep[path] = "daily"

            # Weekly
            seen_weeks = set()
            for path, dt in oldest_first:
                y, w, _ = dt.isocalendar()
                week_key = f"{y}-W{w:02d}"
                if week_key not in seen_weeks:
                    seen_weeks.add(week_key)
                    keep[path] = "weekly"

            # Monthly
            seen_months = set()
            for path, dt in oldest_first:
                month_key = dt.strftime("%Y-%m")
                if month_key not in seen_months:
                    seen_months.add(month_key)
                    keep[path] = "monthly"

            # Yearly
            seen_years = set()
            for path, dt in oldest_first:
                year_key = dt.strftime("%Y")
                if year_key not in seen_years:
                    seen_years.add(year_key)
                    keep[path] = "yearly"

            # Oldest baseline
            keep[oldest_first[0][0]] = "baseline"

            # Delete non-kept
            for path, _ in backups:
                if path not in keep:
                    try:
                        os.remove(path)
                    except Exception as e:
                        print(f"[{subdir}] Failed to delete {os.path.basename(path)}: {e}")

            # Rename kept anchors with suffix (strip old tags first)
            for path, tag in keep.items():
                if not tag:
                    continue
                dirname, fname = os.path.split(path)

                # Remove any old tag suffix
                for t in tags:
                    if fname.endswith(f".{t}"):
                        fname = fname[:-(len(t) + 1)]

                new_name = f"{fname}.{tag}"
                new_path = os.path.join(dirname, new_name)

                if new_path != path:  # only rename if different
                    try:
                        os.rename(path, new_path)
                    except Exception as e:
                        print(f"[{subdir}] Failed to rename {os.path.basename(path)}: {e}")

    except Exception as e:
        print(f"Backup cleanup failed: {e}")

def get_file_checksum(file_path):
    """Calculate MD5 checksum of a file"""
    try:
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception:
        return None

def is_ota_device_available(ip, port=3232, timeout=2):
    try:
        with socket.create_connection((ip, port), timeout=timeout):
            return True
    except Exception:
        return False

def get_device_esphome_version(yaml_path, ip, max_wait=10):
    """Run esphome logs briefly, parse ESPHome firmware version, then stop"""
    if not yaml_path or not ip:
        return None
        
    cmd = ["esphome", "logs", yaml_path, "--device", ip]
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True
        )
        start = time.time()
        for line in iter(proc.stdout.readline, ''):
            # Match the actual format from logs: "ESPHome version 2025.9.3"
            match = re.search(r"ESPHome version ([\d\.]+)", line)
            if match:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except:
                    proc.kill()
                return match.group(1)
            if time.time() - start > max_wait:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except:
                    proc.kill()
                break
        return None
    except Exception as e:
        print(f"Version check failed: {e}")
        return None
    
def get_device_info_with_progress(yaml_path, ip, progress_callback, max_wait=20):
    """Get device info using the reliable log parsing method"""
    cmd = ["esphome", "logs", yaml_path, "--device", ip]
    
    print(f"DEBUG: Starting device info collection")
    
    try:
        import os
        # === NEW IMPORTS ===
        import threading
        # === END NEW IMPORTS ===
        
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        
        if os.name == 'nt':
            env['PYTHONLEGACYWINDOWSSTDIO'] = '1'
        
        proc = subprocess.Popen(
            cmd, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT, 
            text=True,
            env=env,
            encoding='utf-8',
            errors='replace',
            bufsize=0  # Line buffered
        )

        print(f"DEBUG: Subprocess started with PID: {proc.pid}")
        
        start_time = time.time()
        line_count = 0
        
        found_info = {
            'firmware_version': False,
            'host_name': False, 
            'wifi_ssid': False,
            'local_mac': False,
            'wifi_signal': False,
            'chip': False,
            'frequency': False,
            'framework': False,
            'psram': False,
            'partition_table': False,
            'got_partition_data': False,
        }
        
        patterns = {
            'firmware_version': r'ESPHome version ([\d\.]+)',
            'host_name': r"Hostname:\s*'([^']+)'", 
            'wifi_ssid': r"SSID:\s*'([^']+)'",
            'local_mac': r"Local MAC:\s*([0-9A-Fa-f:]{17})",
            'wifi_signal': r"Signal strength:\s*(-?\d+)\s*dB",
            'chip': r'\[D\]\[debug:\d+\]: Chip:\s*Model=([^,]+)',
            'frequency': r'\[D\]\[debug:\d+\]: CPU Frequency:\s*(\d+)\s*MHz',
            'framework': r'\[D\]\[debug:\d+\]: Framework:\s*([^\s,]+)',
            'psram': r'\[C\]\[psram:\d+\]:\s*Size:\s*(\d+)\s*KB',
            'partition_table': r'Partition table:',
        }
        
        collected_values = {}
        partition_sizes = []
        last_line_time = time.time()
        
        # Initial progress
        progress_callback(10, "Connecting to device...")
        
        print(f"DEBUG: Entering main loop")
        
        # === SIMPLE HYBRID APPROACH ===
        while time.time() - start_time < max_wait:
            import threading
            
            def read_with_timeout(timeout=2.0):
                """Read a line with 2-second timeout"""
                result = [None]
                
                def read_line():
                    try:
                        result[0] = proc.stdout.readline()
                    except:
                        result[0] = None
                
                reader = threading.Thread(target=read_line, daemon=True)
                reader.start()
                reader.join(timeout)
                
                if reader.is_alive():
                    # readline() is blocking - no data available
                    return None
                else:
                    return result[0]
            
            line = read_with_timeout(2.0)  # 2 second timeout
            # === END NEW CODE ===
            
            if line:
                line_count += 1
                line = line.rstrip()
                last_line_time = time.time()
                
                print(f"DEBUG: Processing line {line_count}: {line[:80]}{'...' if len(line) > 80 else ''}")
                
                progress_percent = min(10 + (line_count / 2), 90)
                progress_callback(progress_percent, f"Reading logs... ({line_count} lines)")
                
                # Check for partition table
                if 'Partition table:' in line:
                    found_info['partition_table'] = True
                    progress_callback(70, "Reading partition table...")
                    print(f"DEBUG: ✓ Found partition table marker")
                
                # Collect partition sizes
                partition_match = re.search(r'\b(\w+)\s+\d+\s+\d+\s+0x[0-9A-Fa-f]+\s+0x([0-9A-Fa-f]+)', line)
                if partition_match:
                    try:
                        partition_name = partition_match.group(1)
                        size_hex = partition_match.group(2)
                        size_bytes = int(size_hex, 16)
                        partition_sizes.append(size_bytes)
                        found_info['got_partition_data'] = True
                        print(f"DEBUG: ✓ Found partition: {partition_name}, size={size_bytes} bytes (total partitions: {len(partition_sizes)})")
                    except ValueError:
                        print(f"DEBUG: ✗ Failed to parse partition size from: {line}")
                
                # Check other patterns
                for pattern_name, pattern in patterns.items():
                    if pattern_name != 'partition_table':
                        match = re.search(pattern, line)
                        if match and not found_info[pattern_name]:
                            found_info[pattern_name] = True
                            print(f"DEBUG: ✓ Found {pattern_name}: {match.group(1)}")
                            
                            if pattern_name in ['firmware_version', 'host_name', 'wifi_ssid', 'local_mac', 'chip', 'frequency', 'framework', 'wifi_signal']:
                                collected_values[pattern_name] = match.group(1)
                                progress_callback(progress_percent, f"Found {pattern_name}...")
                            
                            elif pattern_name == 'psram':
                                psram_kb = int(match.group(1))
                                if psram_kb >= 1024:
                                    psram_mb = psram_kb / 1024
                                    collected_values['psram_size'] = f"{psram_mb:.1f} MB"
                                else:
                                    collected_values['psram_size'] = f"{psram_kb} KB"
                                progress_callback(progress_percent, "Found PSRAM...")

                # Check for direct flash size reporting
                if 'Flash Chip:' in line and 'Size=' in line and collected_values.get('flash_size', 'N/A') == 'N/A':
                    flash_match = re.search(r'Size=(\d+kB)', line)
                    if flash_match:
                        flash_size_kb = flash_match.group(1)
                        flash_size_mb = int(flash_size_kb.replace('kB', '')) / 1024
                        collected_values['flash_size'] = f"{flash_size_mb:.1f} MB"
                        found_info['got_partition_data'] = True
                        progress_callback(80, "Found flash size...")
                        print(f"DEBUG: ✓ Found direct flash size: {flash_size_mb:.1f} MB")
            
            else:
                # No line available - check timeout
                time_since_last_line = time.time() - last_line_time
                print(f"DEBUG: No line available, time since last line: {time_since_last_line:.1f}s")
                
                essential_info = [
                    'firmware_version', 'host_name', 'wifi_ssid', 'local_mac',
                    'chip', 'frequency', 'framework', 'partition_table'
                ]
                has_all_essential = all(found_info[item] for item in essential_info)
                
                if has_all_essential and time_since_last_line > 3.0:
                    print(f"DEBUG: ✓ EXIT: Has all essentials + 3s timeout")
                    progress_callback(95, "No more output, finishing...")
                    break
                    
                if time.time() - start_time > max_wait - 2:
                    print(f"DEBUG: ⚠ FORCE EXIT: Overall timeout")
                    progress_callback(95, "Timeout reached, finishing...")
                    break
        
        # Clean up
        if proc.poll() is None:
            print(f"DEBUG: Process still running, terminating...")
            proc.terminate()
            try:
                proc.wait(timeout=2)
                print(f"DEBUG: Process terminated successfully")
            except:
                print(f"DEBUG: Process didn't terminate, killing...")
                proc.kill()
        
        # Final processing
        print(f"DEBUG: Final processing - partition_sizes={partition_sizes}")
        print(f"DEBUG: Final processing - collected_values keys: {list(collected_values.keys())}")
        
        if collected_values.get('flash_size', 'N/A') != 'N/A':
            print(f"DEBUG: Using direct flash size: {collected_values['flash_size']}")
        elif partition_sizes:
            total_bytes = sum(partition_sizes)
            common_sizes = {
                4194304: "4.0 MB",
                8388608: "8.0 MB",  
                16777216: "16.0 MB",
                33554432: "32.0 MB",
            }
            closest_size = min(common_sizes.keys(), key=lambda x: abs(x - total_bytes))
            if abs(total_bytes - closest_size) / closest_size < 0.55:
                collected_values['flash_size'] = common_sizes[closest_size]
            else:
                if total_bytes >= 1024 * 1024:
                    collected_values['flash_size'] = f"{total_bytes / (1024 * 1024):.1f} MB"
                else:
                    collected_values['flash_size'] = f"{total_bytes / 1024:.0f} KB"
            print(f"DEBUG: Calculated flash size from partitions: {collected_values['flash_size']} (total_bytes={total_bytes})")
        else:
            collected_values['flash_size'] = "N/A"
            print(f"DEBUG: No flash size data available")

        expected_keys = {
            'firmware_version': 'firmware_version',
            'host_name': 'host_name', 
            'wifi_ssid': 'wifi_ssid',
            'local_mac': 'local_mac',
            'wifi_signal': 'wifi_signal',
            'chip': 'chip',
            'frequency': 'frequency',
            'framework': 'framework',
            'psram_size': 'psram_size',
            'flash_size': 'flash_size'
        }
        
        final_result = {}
        for gui_key, data_key in expected_keys.items():
            final_result[gui_key] = collected_values.get(data_key, 'N/A')
            
        print(f"DEBUG: Final result: {final_result}")
        progress_callback(100, "Complete!")
        print(f"DEBUG: ✓ Device info collection completed successfully")
        return final_result
        
    except Exception as e:
        print(f"DEBUG: ✗ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        progress_callback(0, f"Error: {str(e)}")
        return None

class ModernESPHomeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ESPHome Studio - Professional ESPHome Manager")
        self.root.geometry("2100x1300")
        self.root.minsize(2100, 1400)

        # Force emoji font for colored emojis
        self.setup_emoji_font()

        # SIMPLE EMOJI TEST - add this temporarily
        self.simple_emoji_test()

        # Theme variables
        self.current_theme = "darkly"  # Default dark theme
        self.available_themes = {
            "Dark Themes": [
                ("Darkly", "darkly"),
                ("Cyborg", "cyborg"),
                ("Solar", "solar"),
                ("Superhero", "superhero"),
                ("Vapor", "vapor")
            ],
            "Light Themes": [
                ("Flatly", "flatly"),
                ("Pulse", "pulse"),
                ("Yeti", "yeti"),
                ("Sandstone", "sandstone"),
                ("Lumen", "lumen"),
                ("Journal", "journal"),
                ("Minty", "minty"),
                ("Litera", "litera"),
                ("Cosmo", "cosmo"),
                ("Morph", "morph"),
                ("Simplex", "simplex"),
                ("Cerculean", "cerculean")
            ]
        }
        
        # Initialize variables FIRST
        self.setup_variables()
        
        # Create main layout
        self.setup_main_layout()
        
        # Initialize components
        self.setup_menu()
        self.setup_status_bar()
        
        # Populate initial data
        self.scan_ports()
        self.get_current_versions()
        self.scan_esphome_versions()
        self.load_recent_files()

        # FULL SYNC AT STARTUP IN BACKGROUND
        self.startup_full_sync()

    def simple_emoji_test(self):
        """Windows 11 Color Emoji Diagnostic"""
        test_window = tb.Toplevel(self.root)
        test_window.title("Windows 11 Color Emoji Diagnostic")
        test_window.geometry("600x700")
        
        emojis = "🚀 📁 💾 📦 ⚙️ 🎯 📡 📋"
        
        # Windows 11 Specific Tests
        tb.Label(test_window, text="Windows 11 Color Emoji Diagnostic", font=("Arial", 12, "bold"), bootstyle="primary").pack(anchor='w', pady=(10, 5))
        
        # Test 1: Check if it's a theme issue
        tb.Label(test_window, text="1. Theme Compatibility Test", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        # Try different theme backgrounds
        backgrounds = [
            ("System Default", ""),
            ("Light Background", "#ffffff"), 
            ("Dark Background", "#1e1e1e"),
            ("Blue Background", "#0078d4"),
        ]
        
        for bg_name, bg_color in backgrounds:
            frame = tb.Frame(test_window)
            frame.pack(fill='x', pady=2)
            tb.Label(frame, text=f"{bg_name}:", width=20, anchor='w').pack(side=LEFT)
            label = tb.Label(frame, text=emojis, font=("Segoe UI Emoji", 14))
            label.pack(side=LEFT)
            if bg_color:
                label.configure(background=bg_color)
                if bg_color == "#1e1e1e":
                    label.configure(foreground="white")
        
        # Test 2: Windows 11 Font Fallback Test
        tb.Label(test_window, text="2. Windows 11 Font Stack Test", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        font_stacks = [
            "Segoe UI Emoji",
            "Segoe UI Emoji, Segoe UI Symbol",
            "Segoe UI Emoji, Arial",
            "Segoe UI Variable Display",  # Windows 11's new font
            "Segoe UI",  # Main Windows 11 font
        ]
        
        for font_stack in font_stacks:
            frame = tb.Frame(test_window)
            frame.pack(fill='x', pady=1)
            tb.Label(frame, text=f"{font_stack}:", width=30, anchor='w').pack(side=LEFT)
            tb.Label(frame, text=emojis, font=(font_stack, 12)).pack(side=LEFT)
        
        # Test 3: Windows 11 DPI/Scaling Test
        tb.Label(test_window, text="3. DPI/Scaling Test", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        try:
            from ctypes import windll
            # Get current DPI
            dpi = windll.user32.GetDpiForWindow(test_window.winfo_id())
            tb.Label(test_window, text=f"Current DPI: {dpi}", font=("Arial", 9)).pack(anchor='w')
            
            # Test different font weights (Windows 11 specific)
            weights = ["normal", "bold"]
            for weight in weights:
                frame = tb.Frame(test_window)
                frame.pack(fill='x', pady=1)
                tb.Label(frame, text=f"Weight {weight}:", width=15, anchor='w').pack(side=LEFT)
                tb.Label(frame, text=emojis, font=("Segoe UI Emoji", 12, weight)).pack(side=LEFT)
                
        except Exception as e:
            tb.Label(test_window, text=f"DPI check failed: {e}", font=("Arial", 9)).pack(anchor='w')
        
        # Test 4: Windows 11 Color Format Test
        tb.Label(test_window, text="4. Color Format Test", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        # Test if it's a color depth issue
        try:
            from ctypes import windll
            hdc = windll.user32.GetDC(0)
            color_depth = windll.gdi32.GetDeviceCaps(hdc, 12)  # BITSPIXEL
            tb.Label(test_window, text=f"Color Depth: {color_depth} bits per pixel", font=("Arial", 9)).pack(anchor='w')
            windll.user32.ReleaseDC(0, hdc)
        except:
            pass
        
        # Test 5: Windows 11 App Compatibility
        tb.Label(test_window, text="5. App Compatibility Test", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        tb.Label(test_window, text="If emojis show as black/white, this might be a Tkinter limitation", font=("Arial", 9)).pack(anchor='w')
        tb.Label(test_window, text="on Windows 11 with certain display configurations.", font=("Arial", 9)).pack(anchor='w')
        
        # Test 6: Browser Comparison
        tb.Label(test_window, text="6. Browser Comparison", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        import webbrowser
        def open_emoji_test():
            webbrowser.open("https://getemoji.com/")
        
        tb.Button(test_window, text="Open Browser Emoji Test", command=open_emoji_test, bootstyle="info").pack(anchor='w')
        tb.Label(test_window, text="Compare with browser - if browser shows colors but Tkinter doesn't,", font=("Arial", 9)).pack(anchor='w')
        tb.Label(test_window, text="it's a Tkinter/Win11 compatibility issue.", font=("Arial", 9)).pack(anchor='w')
        
        # Test 7: Final Workaround - Colored Symbols
        tb.Label(test_window, text="7. Colored Symbol Workaround", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        colored_symbols = [
            ("➤", "primary", "Compiler"),
            ("⬆", "success", "Upload"), 
            ("📦", "warning", "Versions"),
            ("⛁", "info", "Backup"),
            ("⚙", "secondary", "Tools"),
            ("📁", "info", "Files"),
            ("📊", "danger", "History"),
            ("📋", "dark", "Logs")
        ]
        
        for symbol, color, description in colored_symbols:
            frame = tb.Frame(test_window)
            frame.pack(fill='x', pady=1)
            tb.Label(frame, text=symbol, font=("Segoe UI Symbol", 14), bootstyle=color).pack(side=LEFT, padx=(0, 10))
            tb.Label(frame, text=description, font=("Arial", 10)).pack(side=LEFT)
        
        # System Info
        tb.Label(test_window, text="8. System Information", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 5))
        
        import platform
        info = f"""OS: Windows 11 {platform.version()}
    Architecture: {platform.architecture()[0]}
    Machine: {platform.machine()}
    Processor: {platform.processor()}"""
        
        tb.Label(test_window, text=info, font=("Consolas", 8), justify="left").pack(anchor='w')

    def startup_full_sync(self):
        """Perform full sync of all files at application startup"""
        def startup_sync_thread():
            self.status_var.set("Performing initial file sync...")
            self.log_message(">>> Smart syncing files...", "auto")
            self.log_text.see(tk.END)
            
            # Perform full sync (all files)
            synced_files = sync_esphome_files(
                r"\\192.168.4.76\config\esphome", 
                r"C:\esphome",
                None  # No backups during startup sync
            )
            
            self.last_sync_time = datetime.now().strftime("%H:%M:%S")
            
            if synced_files:
                self.sync_status_var.set(f"Sync: Startup ({len(synced_files)} files)")
                self.sync_indicator.configure(bootstyle="success")
                self.log_message(f">>> Startup sync completed: {len(synced_files)} files synchronized", "auto")
            else:
                self.sync_status_var.set("Sync: Startup (no changes)")
                self.sync_indicator.configure(bootstyle="success")
                self.log_message(">>> Startup sync completed: No changes needed", "auto")
            
            self.status_var.set("Ready")
            self.log_text.see(tk.END)
        
        # Start sync in background - don't block UI
        threading.Thread(target=startup_sync_thread, daemon=True).start()

    def setup_emoji_font(self):
        """Simple Windows color emoji fix"""
        try:
            # Try to use Windows Segoe UI Emoji font
            import platform
            if platform.system() == "Windows":
                # This font should give you color emojis in Windows 10/11
                self.emoji_font = ("Segoe UI Emoji", 10)
            else:
                self.emoji_font = None
        except:
            self.emoji_font = None

    def setup_main_layout(self):
        """Create the main tabbed interface"""
        # Main container
        main_container = tb.Frame(self.root, padding=10)
        main_container.pack(fill=BOTH, expand=True)
        
        # Create notebook for tabs
        self.notebook = tb.Notebook(main_container, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=True)
        
        # Create tabs
        self.compiler_tab = tb.Frame(self.notebook, padding=10)
        self.versions_tab = tb.Frame(self.notebook, padding=10)
        self.backup_tab = tb.Frame(self.notebook, padding=10)
        self.tools_tab = tb.Frame(self.notebook, padding=10)
        
        self.notebook.add(self.compiler_tab, text="🚀 Compiler")
        self.notebook.add(self.versions_tab, text="📦 Versions")
        self.notebook.add(self.backup_tab, text="💾 Backup")
        self.notebook.add(self.tools_tab, text="⚙️ Tools")
        
        # Setup each tab
        self.setup_compiler_tab()
        self.setup_versions_tab()
        self.setup_backup_tab()
        self.setup_tools_tab()

    def setup_compiler_tab(self):
        """Setup the compiler tab with modern layout"""
        # Main container
        main_container = tb.Frame(self.compiler_tab)
        main_container.pack(fill=BOTH, expand=True)
        
        # Left panel - Configuration with fixed width (wider)
        left_panel = tb.Frame(main_container, width=700)
        left_panel.pack(side=LEFT, fill=Y, padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # Right panel - Log output
        right_panel = tb.Frame(main_container)
        right_panel.pack(side=RIGHT, fill=BOTH, expand=True)
        
        # LEFT PANEL - Configuration sections
        self.setup_file_section(left_panel)
        self.setup_version_section(left_panel)
        self.setup_upload_section(left_panel)
        self.setup_actions_section(left_panel)
        self.setup_build_history_section(left_panel)

        # RIGHT PANEL - Log and progress
        self.setup_log_section(right_panel)

    def setup_file_section(self, parent):
        """File selection section with recent files dropdown"""
        frame = tb.Labelframe(parent, text="📁 File Configuration", padding=10, bootstyle="info")
        frame.pack(fill=X, pady=(0, 10))
        
        # File selection with better layout
        tb.Label(frame, text="YAML File to compile:", bootstyle="info").grid(row=0, column=0, sticky=W, pady=(0, 5))
        
        file_container = tb.Frame(frame)
        file_container.grid(row=1, column=0, columnspan=2, sticky=EW, pady=(0, 10))
        file_container.columnconfigure(0, weight=1)
        
        self.file_path = tk.StringVar()
        
        # Create a frame for the entry and dropdown button
        entry_dropdown_frame = tb.Frame(file_container)
        entry_dropdown_frame.grid(row=0, column=0, sticky=EW, padx=(0, 5))
        entry_dropdown_frame.columnconfigure(0, weight=1)
        
        self.file_entry = tb.Entry(entry_dropdown_frame, textvariable=self.file_path, width=40)
        self.file_entry.grid(row=0, column=0, sticky=EW)
        
        # Create dropdown button for recent files
        self.recent_files_btn = tb.Menubutton(entry_dropdown_frame, text="▼", width=2, bootstyle="outline-primary")
        self.recent_files_btn.grid(row=0, column=1, sticky=EW, padx=(2, 0))
        
        # Create dropdown menu
        self.recent_files_menu = tk.Menu(self.recent_files_btn, tearoff=0)
        self.recent_files_btn.configure(menu=self.recent_files_menu)
        
        # Update the recent files dropdown
        self.update_recent_files_dropdown()
        
        self.browse_btn = tb.Button(file_container, text="Browse", command=self.browse_file, width=12)
        self.browse_btn.grid(row=0, column=1)
        
        # Sync status with more info
        sync_frame = tb.Frame(frame)
        sync_frame.grid(row=2, column=0, columnspan=2, sticky=EW, pady=5)
        
        self.sync_status_var = tk.StringVar(value="Sync: Unknown")
        self.sync_status_label = tb.Label(sync_frame, textvariable=self.sync_status_var, bootstyle="secondary")
        self.sync_status_label.pack(side=LEFT)
        
        self.sync_indicator = tb.Label(sync_frame, text="●", bootstyle="secondary", font=('Arial', 24))
        self.sync_indicator.pack(side=RIGHT)
        
        # Auto-sync info
        self.auto_sync_var = tk.StringVar(value="Auto-sync: On compile/upload")
        tb.Label(frame, textvariable=self.auto_sync_var, bootstyle="info").grid(row=3, column=0, sticky=W)

    def setup_upload_section(self, parent):
        """Upload configuration section"""
        frame = tb.Labelframe(parent, text="⬆️ Upload Configuration", padding=10, bootstyle="info")
        frame.pack(fill=X, pady=(0, 10))
        
        # Upload mode
        mode_frame = tb.Frame(frame)
        mode_frame.grid(row=0, column=0, columnspan=2, sticky=EW, pady=(0, 10))
        
        tb.Label(mode_frame, text="Upload Mode:", bootstyle="info").pack(side=LEFT)
        
        self.upload_mode_var = tk.StringVar(value="COM")
        mode_container = tb.Frame(mode_frame)
        mode_container.pack(side=LEFT, padx=(10, 0))
        
        self.com_btn = tb.Button(mode_container, text="COM", command=lambda: self.set_upload_mode("COM"), width=8, bootstyle="outline-toolbutton")
        self.com_btn.pack(side=LEFT, padx=(0, 5))
        
        self.ota_btn = tb.Button(mode_container, text="OTA", command=lambda: self.set_upload_mode("OTA"), width=8, bootstyle="outline-toolbutton")
        self.ota_btn.pack(side=LEFT)
        
        self.update_toggle_styles()
        
        # COM Port selection
        self.com_frame = tb.Frame(frame)
        self.com_frame.grid(row=1, column=0, columnspan=2, sticky=EW, pady=(0, 10))
        
        tb.Label(self.com_frame, text="COM Port:").grid(row=0, column=0, sticky=W)
        
        port_container = tb.Frame(self.com_frame)
        port_container.grid(row=0, column=1, sticky=EW, padx=(15, 0))
        port_container.columnconfigure(0, weight=1)
        
        self.port_var = tk.StringVar()
        self.port_combo = tb.Combobox(port_container, textvariable=self.port_var, state="readonly")
        self.port_combo.grid(row=0, column=0, sticky=EW, padx=(0, 5))
        
        self.scan_btn = tb.Button(port_container, text="Scan", command=self.scan_ports, width=6)
        self.scan_btn.grid(row=0, column=1)
        
        # OTA Configuration
        self.ota_frame = tb.Frame(frame)
        self.ota_frame.grid(row=2, column=0, columnspan=2, sticky=EW)
        
        tb.Label(self.ota_frame, text="OTA Device:").grid(row=0, column=0, sticky=W)
        
        ota_container = tb.Frame(self.ota_frame)
        ota_container.grid(row=0, column=1, sticky=EW, padx=(10, 0))
        ota_container.columnconfigure(0, weight=1)
        
        self.ota_ip_var = tk.StringVar()
        self.ip_combo = tb.Combobox(ota_container, textvariable=self.ota_ip_var, state="readonly", width=40)
        self.ip_combo.grid(row=0, column=0, sticky=EW, padx=(0, 5))
        
        self.scan_ips_btn = tb.Button(ota_container, text="Scan IPs", command=self.scan_ips, width=8)
        self.scan_ips_btn.grid(row=0, column=1)
        
        self.update_ota_visibility()

    def setup_actions_section(self, parent):
        """Action buttons section"""
        frame = tb.Labelframe(parent, text="🎯 Actions", padding=10, bootstyle="info")
        frame.pack(fill=X, pady=(0, 10))

        # If we have emoji font, apply it to the frame's label
        if hasattr(self, 'emoji_font') and self.emoji_font:
            try:
                # This is a hack to access the labelframe's label
                for child in frame.winfo_children():
                    if isinstance(child, tk.Label):
                        child.configure(font=self.emoji_font)
            except:
                pass  # If it fails, just use default

        # Create a custom label with emoji font for the title
        title_label = tb.Label(frame, text="🎯 Actions2", font=("Segoe UI Emoji", 10, "bold"), bootstyle="info")
        title_label.place(x=10, y=5)  # Position it where the title would be

        # Main actions - Row 1
        action_frame = tb.Frame(frame)
        action_frame.pack(fill=X, pady=5)
        
        self.compile_btn = tb.Button(action_frame, text="Compile", command=self.compile, bootstyle="outline-primary")
        self.compile_btn.pack(side=LEFT, padx=(0, 5))
        
        self.upload_btn = tb.Button(action_frame, text="Upload", command=self.upload, bootstyle="outline-primary")
        self.upload_btn.pack(side=LEFT, padx=(0, 5))
        
        self.compile_upload_btn = tb.Button(action_frame, text="Compile & Upload", command=self.compile_and_upload, bootstyle="success")
        self.compile_upload_btn.pack(side=LEFT, padx=(0, 5))
        
        self.stop_btn = tb.Button(action_frame, text="⏹️ Stop", command=self.stop_process, bootstyle="danger")
        self.stop_btn.pack(side=LEFT, padx=(0, 5))
        
        # Utility actions - Row 2
        util_frame = tb.Frame(frame)
        util_frame.pack(fill=X, pady=10)
        
        # Sync buttons with dropdown
        sync_menu = tk.Menu(util_frame, tearoff=0)
        sync_menu.add_command(label="Smart Sync (Current Project)", command=self.smart_sync_before_compile)
        sync_menu.add_command(label="Full Sync (All Files)", command=self.manual_full_sync)
        
        self.sync_btn = tb.Button(util_frame, text="Sync Files", command=self.smart_sync_before_compile, bootstyle="info")
        self.sync_btn.pack(side=LEFT, padx=(15, 5))
        
        # Right-click for sync options
        def show_sync_menu(event):
            sync_menu.post(event.x_root, event.y_root)
        
        self.sync_btn.bind("<Button-3>", show_sync_menu)  # Right-click
        
        self.backup_btn = tb.Button(util_frame, text="Backup File", command=self.backup_current_file, bootstyle="info")
        self.backup_btn.pack(side=LEFT, padx=(0, 5))
        
        self.clean_btn = tb.Button(util_frame, text="Clean Build", command=self.clean_build, bootstyle="warning")
        self.clean_btn.pack(side=LEFT, padx=(0, 5))
        
        self.clear_btn = tb.Button(util_frame, text="Clear Log", command=self.clear_log, bootstyle="secondary")
        self.clear_btn.pack(side=LEFT)

    def setup_build_history_section(self, parent):
        """Build history section to display last compilation stats"""
        frame = tb.Labelframe(parent, text="📊 Build History", padding=10, bootstyle="info")
        frame.pack(fill=X, pady=(0, 10))

        # Create a grid layout for the history
        data_grid = tb.Frame(frame)
        data_grid.pack(fill=X)

        # Device Info Section - Left
        info_frame_l = tb.Labelframe(data_grid, text="Device Info", padding=5, bootstyle="primary")
        info_frame_l.grid(row=0, column=0, sticky=NSEW, padx=(0, 5), pady=2)
        data_grid.columnconfigure(0, weight=1)

        # Device Info Section - Right
        info_frame_r = tb.Labelframe(data_grid, text="Device Info", padding=5, bootstyle="primary")
        info_frame_r.grid(row=0, column=1, sticky=NSEW, padx=(5, 0), pady=2)
        data_grid.columnconfigure(1, weight=1)

        # Device info variables - Left frame
        self.firmware_version_var = tk.StringVar(value="N/A")
        self.host_name_var = tk.StringVar(value="N/A")
        self.wifi_ssid_var = tk.StringVar(value="N/A")
        self.local_mac_var = tk.StringVar(value="N/A")
        self.psram_size_var = tk.StringVar(value="N/A")

        # CHANGED: Single line with bold label and normal value
        firmware_frame = tb.Frame(info_frame_l)
        firmware_frame.pack(fill=X, pady=1)
        tb.Label(firmware_frame, text="Firmware:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(firmware_frame, textvariable=self.firmware_version_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        host_frame = tb.Frame(info_frame_l)
        host_frame.pack(fill=X, pady=1)
        tb.Label(host_frame, text="Host:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(host_frame, textvariable=self.host_name_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        ssid_frame = tb.Frame(info_frame_l)
        ssid_frame.pack(fill=X, pady=1)
        tb.Label(ssid_frame, text="SSID:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(ssid_frame, textvariable=self.wifi_ssid_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        mac_frame = tb.Frame(info_frame_l)
        mac_frame.pack(fill=X, pady=1)
        tb.Label(mac_frame, text="MAC:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(mac_frame, textvariable=self.local_mac_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        psram_frame = tb.Frame(info_frame_l)
        psram_frame.pack(fill=X, pady=1)
        tb.Label(psram_frame, text="PSRAM:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(psram_frame, textvariable=self.psram_size_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)

        # Device info variables - Right frame
        self.wifi_signal_var = tk.StringVar(value="N/A")
        self.chip_var = tk.StringVar(value="N/A")
        self.frequency_var = tk.StringVar(value="N/A")
        self.framework_var = tk.StringVar(value="N/A")
        self.flash_size_var = tk.StringVar(value="N/A")
        
        # CHANGED: Single line with bold label and normal value
        signal_frame = tb.Frame(info_frame_r)
        signal_frame.pack(fill=X, pady=1)
        tb.Label(signal_frame, text="WiFi Signal:", font=('Arial', 8, 'bold'), width=12, anchor='w').pack(side=LEFT)
        tb.Label(signal_frame, textvariable=self.wifi_signal_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        chip_frame = tb.Frame(info_frame_r)
        chip_frame.pack(fill=X, pady=1)
        tb.Label(chip_frame, text="Chip:", font=('Arial', 8, 'bold'), width=12, anchor='w').pack(side=LEFT)
        tb.Label(chip_frame, textvariable=self.chip_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        freq_frame = tb.Frame(info_frame_r)
        freq_frame.pack(fill=X, pady=1)
        tb.Label(freq_frame, text="Frequency:", font=('Arial', 8, 'bold'), width=12, anchor='w').pack(side=LEFT)
        tb.Label(freq_frame, textvariable=self.frequency_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        framework_frame = tb.Frame(info_frame_r)
        framework_frame.pack(fill=X, pady=1)
        tb.Label(framework_frame, text="Framework:", font=('Arial', 8, 'bold'), width=12, anchor='w').pack(side=LEFT)
        tb.Label(framework_frame, textvariable=self.framework_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        flash_frame = tb.Frame(info_frame_r)
        flash_frame.pack(fill=X, pady=1)
        tb.Label(flash_frame, text="Flash:", font=('Arial', 8, 'bold'), width=12, anchor='w').pack(side=LEFT)
        tb.Label(flash_frame, textvariable=self.flash_size_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        # Initialize build history variables
        self.build_history = {
            'last_compile': {
                'timestamp': 'Never',
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            },
            'last_upload': {
                'timestamp': 'Never', 
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            }
        }

        # Create a grid layout for the history
        history_grid = tb.Frame(frame)
        history_grid.pack(fill=X)
        
        # Last Compile Section
        compile_frame = tb.Labelframe(history_grid, text="Last Compile", padding=5, bootstyle="primary")
        compile_frame.grid(row=0, column=0, sticky=NSEW, padx=(0, 5), pady=2)
        history_grid.columnconfigure(0, weight=1)
        
        # Last Upload Section  
        upload_frame = tb.Labelframe(history_grid, text="Last Upload", padding=5, bootstyle="primary")
        upload_frame.grid(row=0, column=1, sticky=NSEW, padx=(5, 0), pady=2)
        history_grid.columnconfigure(1, weight=1)
        
        # Compile history labels
        self.compile_time_var = tk.StringVar(value="N/A")
        self.compile_duration_var = tk.StringVar(value="N/A")
        self.compile_size_var = tk.StringVar(value="N/A")
        self.compile_flash_var = tk.StringVar(value="N/A")
        self.compile_ram_var = tk.StringVar(value="N/A")
        self.compile_version_var = tk.StringVar(value="N/A")
        self.compile_file_var = tk.StringVar(value="N/A")
        
        # CHANGED: Single line with bold label and normal value
        compile_time_frame = tb.Frame(compile_frame)
        compile_time_frame.pack(fill=X, pady=1)
        tb.Label(compile_time_frame, text="Time:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_time_frame, textvariable=self.compile_time_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        compile_duration_frame = tb.Frame(compile_frame)
        compile_duration_frame.pack(fill=X, pady=1)
        tb.Label(compile_duration_frame, text="Duration:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_duration_frame, textvariable=self.compile_duration_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        compile_size_frame = tb.Frame(compile_frame)
        compile_size_frame.pack(fill=X, pady=1)
        tb.Label(compile_size_frame, text="Size:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_size_frame, textvariable=self.compile_size_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        compile_flash_frame = tb.Frame(compile_frame)
        compile_flash_frame.pack(fill=X, pady=1)
        tb.Label(compile_flash_frame, text="Flash:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_flash_frame, textvariable=self.compile_flash_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        compile_ram_frame = tb.Frame(compile_frame)
        compile_ram_frame.pack(fill=X, pady=1)
        tb.Label(compile_ram_frame, text="RAM:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_ram_frame, textvariable=self.compile_ram_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        compile_version_frame = tb.Frame(compile_frame)
        compile_version_frame.pack(fill=X, pady=1)
        tb.Label(compile_version_frame, text="Version:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_version_frame, textvariable=self.compile_version_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        compile_file_frame = tb.Frame(compile_frame)
        compile_file_frame.pack(fill=X, pady=1)
        tb.Label(compile_file_frame, text="File:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(compile_file_frame, textvariable=self.compile_file_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        # Upload history labels
        self.upload_time_var = tk.StringVar(value="N/A")
        self.upload_duration_var = tk.StringVar(value="N/A")
        self.upload_size_var = tk.StringVar(value="N/A")
        self.upload_flash_var = tk.StringVar(value="N/A")
        self.upload_ram_var = tk.StringVar(value="N/A")
        self.upload_version_var = tk.StringVar(value="N/A")
        self.upload_file_var = tk.StringVar(value="N/A")
        
        # CHANGED: Single line with bold label and normal value
        upload_time_frame = tb.Frame(upload_frame)
        upload_time_frame.pack(fill=X, pady=1)
        tb.Label(upload_time_frame, text="Time:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_time_frame, textvariable=self.upload_time_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        upload_duration_frame = tb.Frame(upload_frame)
        upload_duration_frame.pack(fill=X, pady=1)
        tb.Label(upload_duration_frame, text="Duration:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_duration_frame, textvariable=self.upload_duration_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        upload_size_frame = tb.Frame(upload_frame)
        upload_size_frame.pack(fill=X, pady=1)
        tb.Label(upload_size_frame, text="Size:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_size_frame, textvariable=self.upload_size_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        upload_flash_frame = tb.Frame(upload_frame)
        upload_flash_frame.pack(fill=X, pady=1)
        tb.Label(upload_flash_frame, text="Flash:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_flash_frame, textvariable=self.upload_flash_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        upload_ram_frame = tb.Frame(upload_frame)
        upload_ram_frame.pack(fill=X, pady=1)
        tb.Label(upload_ram_frame, text="RAM:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_ram_frame, textvariable=self.upload_ram_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        upload_version_frame = tb.Frame(upload_frame)
        upload_version_frame.pack(fill=X, pady=1)
        tb.Label(upload_version_frame, text="Version:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_version_frame, textvariable=self.upload_version_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        upload_file_frame = tb.Frame(upload_frame)
        upload_file_frame.pack(fill=X, pady=1)
        tb.Label(upload_file_frame, text="File:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(upload_file_frame, textvariable=self.upload_file_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        

        # === UPDATED: Both buttons on same row ===
        button_frame = tb.Frame(frame)
        button_frame.pack(fill=X, pady=(10, 0))
        
        # Left side: Refresh button
        self.refresh_device_btn = tb.Button(
            button_frame, 
            text="🔄 Refresh Device Info", 
            command=self.start_device_info_check,
            bootstyle="success",
            width=20
        )
        self.refresh_device_btn.pack(side=LEFT, padx=(0, 10))
        
        # Middle: Clear history button
        clear_history_btn = tb.Button(
            button_frame, 
            text="Clear History", 
            command=self.clear_build_history, 
            bootstyle="secondary", 
            width=15
        )
        clear_history_btn.pack(side=LEFT)
        
        # Right side: Status label
        self.device_check_status = tb.Label(
            button_frame, 
            text="Ready", 
            bootstyle="info"
        )
        self.device_check_status.pack(side=RIGHT)
        # === END UPDATED CODE ===

        # # Clear history button
        # clear_history_btn = tb.Button(frame, text="Clear History", command=self.clear_build_history, 
        #                             bootstyle="secondary", width=15)
        # clear_history_btn.pack(pady=(5, 0))

    def setup_log_section(self, parent):
        """Log output section with forced colors for better visibility"""
        frame = tb.Labelframe(parent, text="📋 Output Log", padding=10, bootstyle="primary")
        frame.pack(fill=BOTH, expand=True)
        
        # TOP LINE: Main status and progress
        top_line = tb.Frame(frame)
        top_line.pack(fill=X, pady=(0, 5))
        
        # Phase label (main process phase)
        self.phase_label = tb.Label(top_line, text="Ready", bootstyle="primary", font=('Arial', 10, 'bold'))
        self.phase_label.pack(side=LEFT)
        
        # Timer
        self.timer_label = tb.Label(top_line, textvariable=self.timer_var, bootstyle="secondary", font=('Arial', 10))
        self.timer_label.pack(side=LEFT, padx=(15, 0))
        
        # Error indicator
        self.error_indicator = tb.Label(top_line, text="●", bootstyle="success", font=('Arial', 16))
        self.error_indicator.pack(side=LEFT, padx=(15, 0))
        
        # Firmware size (right side)
        self.firmware_size_var = tk.StringVar(value="Firmware: N/A")
        tb.Label(top_line, textvariable=self.firmware_size_var, bootstyle="info", font=('Arial', 9)).pack(side=RIGHT)
        
        # BOTTOM LINE: Detailed status and spinner
        bottom_line = tb.Frame(frame)
        bottom_line.pack(fill=X, pady=(5, 10))
        
        # Process status with spinner (left side)
        self.process_status_var = tk.StringVar(value="Idle")
        self.process_status = tb.Label(bottom_line, textvariable=self.process_status_var, bootstyle="info", font=('Arial', 9))
        self.process_status.pack(side=LEFT)
        
        # Version info (right side)
        self.inst_ver_var = tk.StringVar(value=" ")
        inst_ver_label = tb.Label(bottom_line, textvariable=self.inst_ver_var, bootstyle="secondary", font=('Arial', 8))
        inst_ver_label.pack(side=RIGHT)
        
        # Progress bar (full width below both lines)
        self.progress_bar = tb.Progressbar(frame, orient=HORIZONTAL, mode="determinate", 
                                        bootstyle="info-striped")
        self.progress_bar.pack(fill=X, pady=(0, 10))
        
        # Log text area - USE THEME-RESISTANT COLORS
        self.log_text = scrolledtext.ScrolledText(
            frame, 
            wrap=WORD, 
            font=('Consolas', 9),
            background='#1e1e1e',  # Dark background
            foreground='#ffffff',   # White default text
            insertbackground='white',  # White cursor
            relief='solid',         # Add border to make it clear this is separate
            borderwidth=1
        )
        self.log_text.pack(fill=BOTH, expand=True)
        
        # CONFIGURE ALL TAGS IN ONE PLACE - NO DUPLICATES!
        self.log_text.tag_configure("script", foreground="#e9a4fe")     # Bright Cyan - Script messages
        self.log_text.tag_configure("esphome", foreground="#ffffff")    # White - ESPHome output
        self.log_text.tag_configure("command", foreground="#ffff00")    # Bright Yellow - Commands
        self.log_text.tag_configure("success", foreground="#90ee90")    # Light Green - Success messages
        self.log_text.tag_configure("warning", foreground="#ffa500")    # Orange - Warning messages  
        self.log_text.tag_configure("error", foreground="#ff6b6b")      # Bright Red - Error messages
        self.log_text.tag_configure("debug", foreground="#a0a0a0")      # Gray - Debug messages        

    def detect_common_errors(self, output):
        """Detect and highlight common compilation errors"""
        error_patterns = {
            "fill error": (r'unexpected "\*fill\*"', 
                        "Memory allocation error. Try cleaning build or adjusting flash settings."),
            "memory full": (r'region.*overflowed', 
                        "Firmware too large for device. Reduce features or use larger flash chip."),
            "wifi credentials": (r'WiFi credential', 
                            "Check WiFi SSID/password in YAML configuration."),
            "compilation failed": (r'Compilation failed', 
                                "Check YAML syntax and dependencies."),
        }
        
        for error_name, (pattern, suggestion) in error_patterns.items():
            if re.search(pattern, output, re.IGNORECASE):
                self.log_message( f">>> DETECTED: {error_name.upper()}", "auto")
                self.log_message( f">>> SUGGESTION: {suggestion}", "auto")
                return True
        return False

    def setup_version_section(self, parent):
        """Version selection section"""
        frame = tb.Labelframe(parent, text="📦 ESPHome Version", padding=10, bootstyle="info")
        frame.pack(fill=X, pady=(0, 10))
        
        # Version selection
        tb.Label(frame, text="Active Version:", bootstyle="info").grid(row=0, column=0, sticky=W, pady=(0, 5))
        
        version_container = tb.Frame(frame)
        version_container.grid(row=1, column=0, columnspan=2, sticky=EW)
        version_container.columnconfigure(0, weight=1)
        
        self.current_esphome_version = tk.StringVar(value="Default")
        self.version_combo = tb.Combobox(version_container, textvariable=self.current_esphome_version, state="readonly")
        self.version_combo.grid(row=0, column=0, sticky=EW)
        
        # Version info
        self.version_info_var = tk.StringVar(value="System default version")
        tb.Label(frame, textvariable=self.version_info_var, bootstyle="info").grid(row=2, column=0, sticky=W)

    def setup_versions_tab(self):
        """Setup the versions management tab"""
        main_frame = tb.Frame(self.versions_tab, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Title
        tb.Label(main_frame, text="ESPHome Version Manager", 
                font=('Arial', 16, 'bold'), bootstyle="primary").pack(pady=10)
        
        # Current version info
        current_frame = tb.Labelframe(main_frame, text="Current Version", padding=10, bootstyle="primary")
        current_frame.pack(fill=X, pady=10)
        
        tb.Label(current_frame, text="Active Version:").grid(row=0, column=0, sticky=W, padx=5, pady=2)
        current_version_label = tb.Label(current_frame, textvariable=self.current_esphome_version, font=('Arial', 10, 'bold'), bootstyle="info")
        current_version_label.grid(row=0, column=1, sticky=W, padx=5, pady=2)
        
        # Quick actions
        action_frame = tb.Labelframe(main_frame, text="Quick Actions", padding=10, bootstyle="primary")
        action_frame.pack(fill=X, pady=10)
        
        button_frame = tb.Frame(action_frame)
        button_frame.pack(fill=X)
        
        tb.Button(button_frame, text="Open Version Manager", 
                command=self.open_version_manager, 
                bootstyle="primary").pack(side=LEFT, padx=5)
        
        tb.Button(button_frame, text="Install New Version", 
                command=self.install_version_dialog, bootstyle="success").pack(side=LEFT, padx=5)
        
        tb.Button(button_frame, text="Refresh Version List", 
                command=self.scan_esphome_versions, bootstyle="info").pack(side=LEFT, padx=5)
        
        # Info
        info_frame = tb.Labelframe(main_frame, text="Information", padding=10, bootstyle="primary")
        info_frame.pack(fill=X, pady=10)
        
        info_text = """Manage multiple ESPHome versions for different projects.

    • Use different versions for compatibility
    • Install specific versions for testing
    • Switch between versions easily
    • All versions are isolated in virtual environments"""
        
        tb.Label(info_frame, text=info_text, justify=LEFT, bootstyle="secondary").pack(anchor=W)

    def setup_backup_tab(self):
        """Setup the backup management tab"""
        main_frame = tb.Frame(self.backup_tab, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        tb.Label(main_frame, text="Backup Management", 
                font=('Arial', 16, 'bold'), bootstyle="primary").pack(pady=10)
        
        tb.Button(main_frame, text="Open Backup Manager", 
                command=self.manage_backups, 
                bootstyle="primary").pack(pady=20)

    def setup_tools_tab(self):
        """Setup the tools tab"""
        main_frame = tb.Frame(self.tools_tab, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        tb.Label(main_frame, text="Tools & Utilities", 
                font=('Arial', 16, 'bold'), bootstyle="primary").pack(pady=10)
        
        # Simple button layout
        buttons = [
            ("Setup Context Menu", self.setup_context_menu, "info"),
            ("Check for Updates", self.check_updates, "info"),
            ("Update ESPHome", self.update_esphome, "success"),
            ("Scan COM Ports", self.scan_ports, "secondary"),
            ("Scan OTA Devices", self.scan_ips, "secondary")
        ]
        
        for text, command, style in buttons:
            tb.Button(main_frame, text=text, command=command, bootstyle=style, width=20).pack(pady=5)

    def setup_menu(self):
        """Setup the main menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        self.menu_bar = menubar
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open YAML", command=self.browse_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Check for Updates", command=self.check_updates)
        tools_menu.add_command(label="Update ESPHome", command=self.update_esphome)
        tools_menu.add_separator()
        tools_menu.add_command(label="Setup Context Menu", command=self.setup_context_menu)
        
        # Theme menu
        theme_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Theme", menu=theme_menu)
        
        # Dark themes submenu
        dark_themes_menu = tk.Menu(theme_menu, tearoff=0)
        theme_menu.add_cascade(label="Dark Themes", menu=dark_themes_menu)
        for theme_name, theme_id in self.available_themes["Dark Themes"]:
            dark_themes_menu.add_command(
                label=theme_name,
                command=lambda t=theme_id: self.change_theme(t)
            )
        
        # Light themes submenu
        light_themes_menu = tk.Menu(theme_menu, tearoff=0)
        theme_menu.add_cascade(label="Light Themes", menu=light_themes_menu)
        for theme_name, theme_id in self.available_themes["Light Themes"]:
            light_themes_menu.add_command(
                label=theme_name,
                command=lambda t=theme_id: self.change_theme(t)
            )
        
        theme_menu.add_separator()
        theme_menu.add_command(label="Toggle Light/Dark", command=self.toggle_light_dark)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def setup_status_bar(self):
        """Setup the status bar"""
        status_frame = tb.Frame(self.root)
        status_frame.pack(fill=X, side=BOTTOM)
        
        self.status_var = tk.StringVar(value="Ready")
        status_bar = tb.Label(status_frame, textvariable=self.status_var, relief=SUNKEN, anchor=W, bootstyle="info")
        status_bar.pack(fill=X)
        
        self.update_status = tk.StringVar(value="Update: Not checked")
        update_label = tb.Label(status_frame, textvariable=self.update_status, relief=SUNKEN, anchor=E, bootstyle="info")
        update_label.pack(fill=X, side=RIGHT)

    def setup_variables(self):
        """Initialize all the variables"""
        # Initialize all required variables first
        self.main_frame = None
        self.update_frame = None
        self.file_label = None
        self.port_label = None
        self.log_label = None
        self.update_label = None
        self.firmware_size_label = None
        self.version_label = None
        self.menu_bar = None

        # Recent files management
        self.recent_files = []  # List to store recent file paths
        self.max_recent_files = 3  # Maximum number of recent files to show

        # Process spinner
        self.spinner_running = False
        self.spinner_steps = ["●", "◎", "○", "◎"]
        self.spinner_index = 0

        # Thread management
        self.compile_thread = None
        self.device_info_thread = None
        self.is_running = False

        # Status variables
        self.last_upload_progress = 0
        self.firmware_size = "N/A"
        self.firmware_max_size = "N/A"
        self.firmware_percentage = "N/A"
        self.last_sync_time = None
        self.synced_files = []
        self.esphome_versions = {}
        self.versions_base_path = Path("C:/esphome_versions")
        self.backup_base_path = Path("C:/esphome_backups")
        self.last_backup_time = None
        self.backup_enabled = tk.BooleanVar(value=True)
        self.max_backups = tk.IntVar(value=10)
        self.ip_list_var = tk.StringVar()
        self.timer_var = tk.StringVar(value="00:00")
        self.timer_running = False
        self.phase_label = None

        self.backup_status_var = tk.StringVar(value="Backup: Not created")
        self.auto_sync_var = tk.StringVar(value="Auto-sync: On compile/upload")

        # Add these for process control
        self.current_process = None
        self.is_running = False

        # Status bar variables
        self.status_var = tk.StringVar(value="Ready")
        self.update_status = tk.StringVar(value="Update: Not checked")
        self.last_update_check = "Never"
        
        # Version information
        self.current_versions = {
            'esphome': None, 'python': None, 'pyserial': None, 'platformio': None
        }

    def protect_log_colors(self):
        """Re-apply forced colors to log text widget after theme changes"""
        # Re-apply the dark theme colors to resist ttkbootstrap theme changes
        self.log_text.configure(
            background='#1e1e1e',
            foreground='#ffffff',
            insertbackground='white'
        )

    def start_device_info_check(self):
        """Start manual device info check"""
        if not self.file_path.get() or not self.ota_ip_var.get():
            messagebox.showwarning("Warning", "Please select a YAML file and ensure OTA device is selected")
            return
        
        # Disable start button, enable stop button
        self.refresh_device_btn.configure(state="disabled")
        self.device_check_status.configure(text="Checking device...", bootstyle="warning")
        
        # Start the device info check
        self.get_device_info_for_selected_file()

    def change_theme(self, theme_name):
        """Change the application theme"""
        try:
            self.root.style.theme_use(theme_name)
            self.current_theme = theme_name
            self.status_var.set(f"Theme changed to {theme_name}")
            
            # Update theme-dependent elements
            self.update_toggle_styles()
            
            # PROTECT THE LOG COLORS - re-apply dark theme to log
            self.protect_log_colors()

        except Exception as e:
            messagebox.showerror("Theme Error", f"Could not load theme '{theme_name}': {e}")

    def toggle_light_dark(self):
        """Toggle between light and dark theme families"""
        # Get current theme family
        current_is_dark = any(self.current_theme == theme_id for _, theme_id in self.available_themes["Dark Themes"])
        
        if current_is_dark:
            # Switch to first light theme
            new_theme = self.available_themes["Light Themes"][0][1]
        else:
            # Switch to first dark theme
            new_theme = self.available_themes["Dark Themes"][0][1]
        
        self.change_theme(new_theme)

    def browse_file(self):
        """Browse for YAML file and populate device info"""
        filename = filedialog.askopenfilename(
            title="Select ESPHome YAML file",
            filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            # ADD THIS LINE to add the file to recent files
            self.add_to_recent_files(filename)
            self.status_var.set(f"Selected: {os.path.basename(self.file_path.get())}")
            self.scan_ips()  # Auto-trigger IP scan
            self.check_sync_status()  # Check sync status for the selected file
            # Ensure we have the latest version list before getting device info
            self.scan_esphome_versions()
            self.device_check_status.configure(text="Auto-checking device...", bootstyle="warning")
            
            # Get device information when file is selected
            self.get_device_info_for_selected_file()

    def get_device_info_for_selected_file(self):
        """Get device information for the selected YAML file with progress"""
        yaml_path = self.file_path.get()
        ip = self.ota_ip_var.get().strip()
        
        # === FIX: Reset stop flag at start ===
        self.device_info_stop_requested = False

        if not yaml_path or not ip:
            self.update_device_info_display({
                'firmware_version': 'N/A',
                'host_name': 'N/A', 
                'wifi_ssid': 'N/A',
                'local_mac': 'N/A',
                'wifi_signal': 'N/A',
                'chip': 'N/A',
                'frequency': 'N/A',
                'framework': 'N/A',
                'psram_size': 'N/A',
                'flash_size': 'N/A'
            })
            # === NEW CODE: Reset buttons if no file/IP ===
            self.refresh_device_btn.configure(state="normal")
            self.device_check_status.configure(text="No file/device", bootstyle="secondary")
            # === END NEW CODE ===

            return
            
        def progress_callback(progress, status):
            """Update progress bar and status from worker thread"""
            self.root.after(0, lambda: self.update_progress(progress))
            self.root.after(0, lambda: self.status_var.set(status))
            self.root.after(0, lambda: self.update_phase_label(status))
            # === NEW CODE: Update status label ===
            self.root.after(0, lambda: self.device_check_status.configure(text=status, bootstyle="info"))
            # === END NEW CODE ===

        def device_info_thread():
            self.root.after(0, lambda: self.update_progress(0))
            device_info = get_device_info_with_progress(yaml_path, ip, progress_callback)
            
            if device_info:
                self.root.after(0, lambda: self.update_device_info_display(device_info))
                self.root.after(0, lambda: self.status_var.set("Device information complete"))
                self.root.after(0, lambda: self.update_phase_label("Done"))
                self.root.after(0, lambda: self.device_check_status.configure(text="Complete", bootstyle="success"))
            else:
                self.root.after(0, lambda: self.update_device_info_display({
                    'firmware_version': 'Error',
                    'host_name': 'Error',
                    'wifi_ssid': 'Error',
                    'local_mac': 'Error', 
                    'wifi_signal': 'Error',
                    'chip': 'Error',
                    'frequency': 'Error',
                    'framework': 'Error',
                    'psram_size': 'Error',
                    'flash_size': 'Error'
                }))
                self.root.after(0, lambda: self.status_var.set("Failed to get device information"))
                self.root.after(0, lambda: self.update_phase_label("Error"))
                self.root.after(0, lambda: self.device_check_status.configure(text="Failed", bootstyle="danger"))
            # === NEW CODE: Reset buttons when done ===
            self.root.after(0, lambda: self.refresh_device_btn.configure(state="normal"))
            # === END NEW CODE ===


            # Reset progress after a short delay
            self.root.after(2000, lambda: self.update_progress(0))
                    
        threading.Thread(target=device_info_thread, daemon=True).start()

    def update_device_info_display(self, device_info):
        """Update the device info display with retrieved information"""
        # Update left frame
        self.firmware_version_var.set(f"{device_info['firmware_version']}")
        self.host_name_var.set(f"{device_info['host_name']}")
        self.wifi_ssid_var.set(f"{device_info['wifi_ssid']}")
        self.local_mac_var.set(f"{device_info['local_mac']}")
        self.psram_size_var.set(f"{device_info.get('psram_size', 'N/A')}")
        
        # Update right frame
        self.wifi_signal_var.set(f"{device_info['wifi_signal']}")
        self.chip_var.set(f"{device_info['chip']}")
        self.frequency_var.set(f"{device_info['frequency']}")
        self.framework_var.set(f"{device_info['framework']}")
        self.flash_size_var.set(f"{device_info.get('flash_size', 'N/A')}")

        # Auto-select the exact matching ESPHome version
        if device_info['firmware_version'] != 'N/A' and device_info['firmware_version'] != 'Error':
            self.auto_select_esphome_version(device_info['firmware_version'])

    def update_build_history(self, action_type, duration, firmware_size=None, flash_usage=None, ram_usage=None):
        """Update build history with new compilation/upload stats"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        filename = os.path.basename(self.file_path.get()) if self.file_path.get() else "N/A"
        version = self.current_esphome_version.get()
        
        history_data = {
            'timestamp': timestamp,
            'duration': f"{duration:.1f}s",
            'firmware_size': firmware_size or "N/A",
            'flash_usage': flash_usage or "N/A",
            'ram_usage': ram_usage or "N/A",
            'version': version,
            'file': filename
        }
        
        self.build_history[f'last_{action_type}'] = history_data
        self.update_history_display(action_type)

    def update_history_display(self, action_type):
        """Update the UI with build history data"""
        history = self.build_history[f'last_{action_type}']
        
        if action_type == 'compile':
            self.compile_time_var.set(f"{history['timestamp']}")
            self.compile_duration_var.set(f"{history['duration']}")
            self.compile_size_var.set(f"{history['firmware_size']}")
            self.compile_flash_var.set(f"{history['flash_usage']}")
            self.compile_ram_var.set(f"{history['ram_usage']}")
            self.compile_version_var.set(f"{history['version']}")
            self.compile_file_var.set(f"{history['file']}")
        else:  # upload
            self.upload_time_var.set(f"{history['timestamp']}")
            self.upload_duration_var.set(f" {history['duration']}")
            self.upload_size_var.set(f"{history['firmware_size']}")
            self.upload_flash_var.set(f"{history['flash_usage']}")
            self.upload_ram_var.set(f"{history['ram_usage']}")
            self.upload_version_var.set(f"{history['version']}")
            self.upload_file_var.set(f"{history['file']}")

    def clear_build_history(self):
        """Clear all build history"""
        self.build_history = {
            'last_compile': {
                'timestamp': 'Never',
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            },
            'last_upload': {
                'timestamp': 'Never',
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            }
        }
        self.update_history_display('compile')
        self.update_history_display('upload')
        self.update_device_info_display({
            'firmware_version': 'N/A',
            'host_name': 'N/A', 
            'wifi_ssid': 'N/A',
            'local_mac': 'N/A',
            'wifi_signal': 'N/A',
            'chip': 'N/A',
            'frequency': 'N/A',
            'framework': 'N/A',
            'psram_size': 'N/A',
            'flash_size': 'N/A'
        })
        self.status_var.set("Build history cleared")

    def auto_select_esphome_version(self, firmware_version):
        """Automatically select the exact matching ESPHome version if available"""
        if not firmware_version or firmware_version == 'N/A' or firmware_version == 'Error':
            return False
            
        # Clean up the version string
        clean_version = firmware_version.strip()
        
        # Look for exact match only (case-sensitive)
        for version_name, version_info in self.esphome_versions.items():
            if version_name == "Default":
                continue  # Skip the default system version
                
            if version_info["version"] == clean_version:
                self.current_esphome_version.set(version_name)
                self.status_var.set(f"Auto-selected ESPHome version: {version_name} ({clean_version})")
                self.log_message( f">>> Auto-selected matching ESPHome version: {version_name}", "auto")
                return True
        
        # No exact match found
        self.status_var.set(f"No exact ESPHome version match for device firmware {clean_version}")
        self.log_message( f">>> No exact ESPHome version match found for device firmware {clean_version}", "auto")
        return False

    def start_process_spinner(self, process_name):
        """Start a text spinner for process feedback - slower and smoother"""
        self.spinner_running = True
        self.spinner_steps = ["●", "◎", "○"]  # Simpler, fewer steps
        self.spinner_index = 0
        self.current_process_name = process_name
        
        def update_spinner():
            if hasattr(self, 'spinner_running') and self.spinner_running:
                spinner_char = self.spinner_steps[self.spinner_index]
                # Only update the spinner character, keep process name static
                self.process_status_var.set(f"{spinner_char} {self.current_process_name}")
                self.spinner_index = (self.spinner_index + 1) % len(self.spinner_steps)
                self.root.after(800, update_spinner)  # Slower: 800ms instead of 500ms
        
        update_spinner()

    def update_process_status(self, status_text):
        """Update the process status without changing the spinner"""
        if hasattr(self, 'spinner_running') and self.spinner_running:
            self.current_process_name = status_text
            # Keep current spinner character, just update text
            spinner_char = self.spinner_steps[self.spinner_index]
            self.process_status_var.set(f"{spinner_char} {status_text}")

    def stop_process_spinner(self):
        """Stop the process spinner"""
        self.spinner_running = False
        self.process_status_var.set("Idle")

    def set_upload_mode(self, mode):
        self.upload_mode_var.set(mode)
        self.update_toggle_styles()
        self.update_ota_visibility()

    def update_toggle_styles(self):
        if self.upload_mode_var.get() == "COM":
            self.com_btn.configure(bootstyle="primary")
            self.ota_btn.configure(bootstyle="outline-primary")
        else:
            self.com_btn.configure(bootstyle="outline-primary")
            self.ota_btn.configure(bootstyle="primary")

    def update_ota_visibility(self):
        if self.upload_mode_var.get() == "OTA":
            self.ota_frame.grid()
            self.com_frame.grid_remove()
        else:
            self.ota_frame.grid_remove()
            self.com_frame.grid()

    def scan_ips(self):
        self.status_var.set("Scanning for OTA devices...")
        devices = discover_esphome_devices()
        devices.sort(key=lambda x: x[0])  # Sort by name

        display_list = [f"{name} ({ip})" for name, ip in devices]
        ip_only_list = [ip for _, ip in devices]

        self.ip_combo['values'] = display_list

        yaml_name = os.path.splitext(os.path.basename(self.file_path.get()))[0] if self.file_path.get() else ""

        # Try to match YAML name after populating dropdown
        for name, ip in devices:
            if name == yaml_name:
                self.ota_ip_var.set(ip)
                break
        else:
            # Fallback to first device if no match
            if ip_only_list:
                self.ota_ip_var.set(ip_only_list[0])

        if ip_only_list:
            self.status_var.set(f"Found {len(ip_only_list)} OTA device(s)")
        else:
            self.status_var.set("No OTA devices found")

    def scan_ports(self):
        """Scan for available COM ports"""
        self.status_var.set("Scanning for COM ports...")
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            self.port_combo.set(ports[0])
            self.status_var.set(f"Found {len(ports)} COM port(s)")
        else:
            self.status_var.set("No COM ports found")

    def stop_process(self):
        """Stop the current running process immediately"""
        if hasattr(self, 'current_process') and self.current_process:
            try:
                self.is_running = False
                self._force_kill_process()
                self.status_var.set("Process forcefully stopped")
                self.update_phase_label("Stopped")
                self.log_message( ">>> Process forcefully terminated", "auto")
                self.log_text.see(tk.END)
                self.update_progress(0)
                self.stop_timer()
                self.error_indicator.configure(bootstyle="warning")

            except Exception as e:
                self.log_message( f">>> Error stopping process: {str(e)}", "auto")

    def clean_build(self):
        """Clean the build directory to fix compilation issues"""
        if not self.validate_file():
            return
            
        def clean_thread():
            self.status_var.set("Cleaning build...")
            yaml_path = self.file_path.get()
            build_dir = os.path.join(os.path.dirname(yaml_path), ".esphome", "build")
            
            if os.path.exists(build_dir):
                try:
                    shutil.rmtree(build_dir)
                    self.log_message( ">>> Build directory cleaned successfully", "auto")
                    self.status_var.set("Build directory cleaned")
                except Exception as e:
                    self.log_message( f">>> Error cleaning build: {str(e)}", "auto")
                    self.status_var.set("Error cleaning build")
            else:
                self.log_message( ">>> No build directory found", "auto")
                self.status_var.set("No build directory found")
        
        threading.Thread(target=clean_thread, daemon=True).start()

    def backup_current_file(self):
        """Create a backup of the current YAML file"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
            
        def backup_thread():
            self.status_var.set("Creating backup...")
            backup_path = create_backup(
                self.file_path.get(), 
                self.backup_base_path,
                os.path.basename(self.file_path.get())
            )
            
            if backup_path:
                self.last_backup_time = datetime.now().strftime("%H:%M:%S")
                self.backup_status_var.set(f"Backup: {self.last_backup_time}")
                self.log_message( f">>> Backup created: {backup_path}", "auto")
                self.status_var.set("Backup created successfully")
                
                # Clean up old backups
                cleanup_old_backups(self.backup_base_path, self.max_backups.get())
            else:
                self.log_message( ">>> Backup failed", "auto")
                self.status_var.set("Backup failed")
        
        threading.Thread(target=backup_thread, daemon=True).start()

    def auto_backup_file(self, file_path):
        """Automatically backup file if backup is enabled"""
        if self.backup_enabled.get() and file_path and os.path.exists(file_path):
            backup_path = create_backup(
                file_path,
                self.backup_base_path,
                os.path.basename(file_path)
            )
            if backup_path:
                self.last_backup_time = datetime.now().strftime("%H:%M:%S")
                self.backup_status_var.set(f"Backup: Auto ({self.last_backup_time})")
                self.log_message( f">>> Auto-backup created: {os.path.basename(backup_path)}", "auto")
            
            # Clean up old backups
            cleanup_old_backups(self.backup_base_path, self.max_backups.get())

    def manage_backups(self):
        """Open backup management window"""
        backup_window = tb.Toplevel(self.root)
        backup_window.title("Backup Management")
        backup_window.geometry("700x500")
        backup_window.transient(self.root)
        backup_window.grab_set()
        
        # Create frames
        main_frame = tb.Frame(backup_window, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        # Settings frame
        settings_frame = tb.Labelframe(main_frame, text="Backup Settings", padding="10", bootstyle="primary")
        settings_frame.pack(fill=X, pady=5)
        
        tb.Checkbutton(settings_frame, text="Enable automatic backups", 
                       variable=self.backup_enabled, bootstyle="primary-round-toggle").pack(anchor=W, pady=2)
        
        tb.Label(settings_frame, text="Max backups to keep:").pack(anchor=W, pady=2)
        max_backup_spin = tb.Spinbox(settings_frame, from_=1, to=50, width=5, 
                                     textvariable=self.max_backups)
        max_backup_spin.pack(anchor=W, pady=2)
        
        tb.Label(settings_frame, text=f"Backup location: {self.backup_base_path}", bootstyle="info").pack(anchor=W, pady=2)
        
        # Button frame
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=X, pady=5)
        
        tb.Button(button_frame, text="Refresh", 
                  command=lambda: self.populate_backup_list(tree), bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Create Backup", 
                  command=self.backup_current_file, bootstyle="success").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Cleanup Old", 
                  command=lambda: self.cleanup_backups(tree), bootstyle="warning").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Close", 
                  command=backup_window.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        
        # Treeview for backups
        tree_frame = tb.Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=5)
        
        columns = ("Filename", "Date", "Size", "Original")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        # Define headings
        tree.heading("Filename", text="Backup File")
        tree.heading("Date", text="Backup Date")
        tree.heading("Size", text="Size")
        tree.heading("Original", text="Original File")
        
        # Define columns
        tree.column("Filename", width=200)
        tree.column("Date", width=150)
        tree.column("Size", width=100)
        tree.column("Original", width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Populate the list
        self.populate_backup_list(tree)

    def populate_backup_list(self, tree):
        """Populate backup manager treeview"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Scan backup directory
        if not self.backup_base_path.exists():
            return
            
        try:
            for subdir in self.backup_base_path.iterdir():
                if subdir.is_dir():
                    for backup_file in subdir.iterdir():
                        if backup_file.is_file():
                            stat = backup_file.stat()
                            file_time = datetime.fromtimestamp(stat.st_mtime)
                            
                            tree.insert("", "end", values=(
                                backup_file.name,
                                file_time.strftime("%Y-%m-%d %H:%M:%S"),
                                f"{stat.st_size / 1024:.1f} KB",
                                subdir.name
                            ))
        except Exception as e:
            print(f"Error populating backup list: {e}")

    def restore_backup(self, tree):
        """Restore selected backup"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a backup to restore")
            return
        
        item = tree.item(selection[0])
        backup_name = item['values'][0]
        original_name = item['values'][3]
        
        backup_path = self.backup_base_path / backup_name
        
        if not backup_path.exists():
            messagebox.showerror("Error", "Backup file not found")
            return
        
        # Ask for restore location
        restore_path = filedialog.asksaveasfilename(
            title="Restore backup as...",
            initialfile=original_name,
            filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")]
        )
        
        if restore_path:
            try:
                shutil.copy2(backup_path, restore_path)
                messagebox.showinfo("Success", f"Backup restored to:\n{restore_path}")
                # Update current file if it matches
                if self.file_path.get() and os.path.basename(self.file_path.get()) == original_name:
                    self.file_path.set(restore_path)
                    self.check_sync_status()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to restore backup: {str(e)}")

    def delete_backup(self, tree):
        """Delete selected backup"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a backup to delete")
            return
        
        item = tree.item(selection[0])
        backup_name = item['values'][0]
        
        if not messagebox.askyesno("Confirm Delete", f"Delete backup '{backup_name}'?"):
            return
        
        try:
            backup_path = self.backup_base_path / backup_name
            backup_path.unlink()
            self.populate_backup_list(tree)
            messagebox.showinfo("Success", "Backup deleted")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete backup: {str(e)}")

    def cleanup_backups(self, tree):
        """Clean up old backups"""
        if messagebox.askyesno("Confirm Cleanup", 
                              f"Remove old backups, keeping only {self.max_backups.get()} most recent?"):
            cleanup_old_backups(self.backup_base_path, self.max_backups.get())
            self.populate_backup_list(tree)
            messagebox.showinfo("Success", "Old backups cleaned up")

    def sync_all_files(self):
        """Enhanced sync that includes all file types"""
        def sync_thread():
            self.status_var.set("Syncing all files...")
            # Use enhanced sync function
            synced_files = sync_esphome_files(
                r"\\192.168.4.76\config\esphome", 
                r"C:\esphome",
                self.backup_base_path if self.backup_enabled.get() else None
            )
            self.last_sync_time = datetime.now().strftime("%H:%M:%S")
            
            if synced_files:
                self.synced_files = synced_files
                current_file = os.path.basename(self.file_path.get()) if self.file_path.get() else None
                
                if current_file in synced_files:
                    self.sync_status_var.set(f"Sync: Just synced ({self.last_sync_time})")
                    self.sync_indicator.configure(bootstyle="success")
                    self.status_var.set(f"Synced {len(synced_files)} file(s) including current")
                else:
                    self.sync_status_var.set(f"Sync: Synced {len(synced_files)} file(s)")
                    self.sync_indicator.configure(bootstyle="info")
                    self.status_var.set(f"Synced {len(synced_files)} file(s)")
                    
                # Log detailed sync information
                self.log_message( f">>> Sync completed at {self.last_sync_time}", "auto")
                self.log_message( f">>> Files synced: {', '.join(synced_files)}", "auto")
            else:
                self.sync_status_var.set("Sync: No files needed syncing")
                self.sync_indicator.configure(bootstyle="success")
                self.status_var.set("No files needed syncing")
            
            # Re-check sync status for current file
            self.check_sync_status()
        
        threading.Thread(target=sync_thread, daemon=True).start()

    def manual_sync(self):
        """Manual sync of all files"""
        self.sync_all_files()

    def check_sync_status(self):
        """Check if the current YAML file needs syncing"""
        if not self.file_path.get():
            self.sync_status_var.set("Sync: No file selected")
            self.sync_indicator.configure(bootstyle="secondary")
            return False
        
        filename = os.path.basename(self.file_path.get())
        network_path = r"\\192.168.4.76\config\esphome"
        local_path = r"C:\esphome"
        
        src = os.path.join(network_path, filename)
        dst = os.path.join(local_path, filename)
        
        try:
            if not os.path.exists(src):
                self.sync_status_var.set("Sync: Network file not found")
                self.sync_indicator.configure(bootstyle="warning")
                return False
            
            if not os.path.exists(dst):
                self.sync_status_var.set("Sync: Needs sync (local missing)")
                self.sync_indicator.configure(bootstyle="danger")
                return True
            
            # Check if source is newer using checksum for better accuracy
            src_checksum = get_file_checksum(src)
            dst_checksum = get_file_checksum(dst)
            
            if src_checksum and dst_checksum and src_checksum != dst_checksum:
                self.sync_status_var.set("Sync: Needs sync (content changed)")
                self.sync_indicator.configure(bootstyle="danger")
                return True
            elif os.path.getmtime(src) > os.path.getmtime(dst):
                self.sync_status_var.set("Sync: Needs sync (newer version)")
                self.sync_indicator.configure(bootstyle="danger")
                return True
            else:
                self.sync_status_var.set("Sync: Up to date")
                self.sync_status_label.configure(bootstyle="info")
                self.sync_indicator.configure(bootstyle="success")
                return False
                
        except Exception as e:
            self.sync_status_var.set(f"Sync: Error checking")
            self.sync_indicator.configure(bootstyle="warning")
            return False

    def scan_esphome_versions(self):
        """Scan for available ESPHome versions"""
        self.esphome_versions = {"Default": {"path": "esphome", "version": "System Default"}}
        
        # Create versions directory if it doesn't exist
        self.versions_base_path.mkdir(exist_ok=True)
        
        # Scan for virtual environments
        if self.versions_base_path.exists():
            for version_dir in self.versions_base_path.iterdir():
                if version_dir.is_dir():
                    # Check if it's a valid virtual environment
                    scripts_dir = version_dir / "Scripts"
                    esphome_exe = scripts_dir / "esphome.exe"
                    
                    if esphome_exe.exists():
                        # Get version info
                        try:
                            result = subprocess.run(
                                [str(esphome_exe), "version"], 
                                capture_output=True, 
                                text=True, 
                                timeout=10
                            )
                            if result.returncode == 0:
                                # Extract version from output
                                version_match = re.search(r'Version:\s*([\d.]+)', result.stdout)
                                if version_match:
                                    version_str = version_match.group(1)
                                    self.esphome_versions[version_dir.name] = {
                                        "path": str(esphome_exe),
                                        "version": version_str
                                    }
                        except (subprocess.TimeoutExpired, subprocess.CalledProcessError):
                            # If version check fails, still add it but mark as unknown
                            self.esphome_versions[version_dir.name] = {
                                "path": str(esphome_exe),
                                "version": "Unknown"
                            }
        
        # Update combo box
        version_names = list(self.esphome_versions.keys())
        self.version_combo['values'] = version_names
        
        # Set default selection
        if "Default" in version_names:
            self.current_esphome_version.set("Default")
        elif version_names:
            self.current_esphome_version.set(version_names[0])

    def get_esphome_command(self):
        """Get the correct ESPHome command based on selected version"""
        selected = self.current_esphome_version.get()
        if selected == "Default" or selected not in self.esphome_versions:
            return "esphome"
        else:
            return f'"{self.esphome_versions[selected]["path"]}"'

    def open_version_manager(self):
        """Open version manager window"""
        version_window = tb.Toplevel(self.root)
        version_window.title("ESPHome Version Manager")
        version_window.geometry("700x500")
        version_window.transient(self.root)
        version_window.grab_set()
        
        # Create frames
        main_frame = tb.Frame(version_window, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        # Button frame
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=X, pady=5)
        
        tb.Button(button_frame, text="Refresh", 
                  command=lambda: self.populate_version_list(tree), bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Install New", 
                  command=self.install_version_dialog, bootstyle="success").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Remove Selected", 
                  command=lambda: self.remove_version(tree), bootstyle="danger").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Set as Active", 
                  command=lambda: self.set_default_version(tree), bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Close", 
                  command=version_window.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        
        # Treeview for versions
        tree_frame = tb.Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=5)
        
        columns = ("Name", "Version", "Path", "Status")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        # Define headings
        tree.heading("Name", text="Environment Name")
        tree.heading("Version", text="ESPHome Version")
        tree.heading("Path", text="Path")
        tree.heading("Status", text="Status")
        
        # Define columns
        tree.column("Name", width=150)
        tree.column("Version", width=120)
        tree.column("Path", width=250)
        tree.column("Status", width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Populate the list
        self.populate_version_list(tree)

    def populate_version_list(self, tree):
        """Populate version manager treeview"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Re-scan versions
        self.scan_esphome_versions()
        
        # Add versions to treeview
        for name, info in self.esphome_versions.items():
            status = "System" if name == "Default" else "Virtual Env"
            if name == self.current_esphome_version.get():
                status += " (Active)"
            
            tree.insert("", "end", values=(
                name,
                info["version"],
                info["path"],
                status
            ))

    def install_version_dialog(self):
        """Dialog to install a specific ESPHome version"""
        install_window = tb.Toplevel(self.root)
        install_window.title("Install ESPHome Version")
        install_window.geometry("500x350")
        install_window.transient(self.root)
        install_window.grab_set()
        
        main_frame = tb.Frame(install_window, padding="20")
        main_frame.pack(fill=BOTH, expand=True)
        
        # Environment name
        tb.Label(main_frame, text="Environment Name:").pack(anchor=W, pady=5)
        env_name_var = tk.StringVar()
        env_entry = tb.Entry(main_frame, textvariable=env_name_var)
        env_entry.pack(fill=X, pady=5)
        
        # Version selection
        tb.Label(main_frame, text="ESPHome Version:").pack(anchor=W, pady=5)
        version_var = tk.StringVar()
        version_entry = tb.Entry(main_frame, textvariable=version_var)
        version_entry.pack(fill=X, pady=5)
        
        # Common versions
        tb.Label(main_frame, text="Common Versions:", bootstyle="primary").pack(anchor=W, pady=5)
        common_frame = tb.Frame(main_frame)
        common_frame.pack(fill=X, pady=5)
        
        common_versions = ["2024.12.9", "2025.6.6", "2025.9.0", "latest"]
        for ver in common_versions:
            tb.Button(common_frame, text=ver, 
                      command=lambda v=ver: version_var.set(v),
                      bootstyle="outline-primary").pack(side=LEFT, padx=2)
        
        # Progress and log
        progress_var = tk.StringVar(value="Ready to install")
        tb.Label(main_frame, textvariable=progress_var, bootstyle="info").pack(anchor=W, pady=10)
        
        log_text = scrolledtext.ScrolledText(main_frame, height=8)
        log_text.pack(fill=BOTH, expand=True, pady=5)
        
        # Buttons
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=X, pady=10)
        
        def install_version():
            env_name = env_name_var.get().strip()
            version = version_var.get().strip()
            
            if not env_name or not version:
                messagebox.showerror("Error", "Please provide both environment name and version")
                return
            
            # Check if name already exists
            if env_name in self.esphome_versions:
                if not messagebox.askyesno("Overwrite?", f"Environment '{env_name}' already exists. Overwrite?"):
                    return
            
            def install_thread():
                try:
                    progress_var.set("Creating virtual environment...")
                    log_text.insert(tk.END, f"Creating environment: {env_name}\n")
                    
                    env_path = self.versions_base_path / env_name
                    
                    # Remove existing if overwriting
                    if env_path.exists():
                        shutil.rmtree(env_path)
                    
                    # Create virtual environment
                    venv.create(env_path, with_pip=True)
                    
                    progress_var.set("Installing ESPHome...")
                    log_text.insert(tk.END, f"Installing ESPHome {version}...\n")
                    log_text.see(tk.END)
                    install_window.update_idletasks()
                    
                    # Install ESPHome
                    pip_exe = env_path / "Scripts" / "pip.exe"
                    esphome_spec = f"esphome=={version}" if version != "latest" else "esphome"
                    
                    result = subprocess.run([
                        str(pip_exe), "install", esphome_spec
                    ], capture_output=True, text=True)
                    
                    if result.returncode == 0:
                        progress_var.set("Installation completed successfully!")
                        log_text.insert(tk.END, "Installation completed successfully!\n")
                        log_text.insert(tk.END, result.stdout)
                        
                        # Refresh version list
                        self.scan_esphome_versions()
                        messagebox.showinfo("Success", f"ESPHome {version} installed as '{env_name}'")
                    else:
                        progress_var.set("Installation failed!")
                        log_text.insert(tk.END, f"Installation failed!\n{result.stderr}")
                        
                except Exception as e:
                    progress_var.set(f"Error: {str(e)}")
                    log_text.insert(tk.END, f"Error: {str(e)}\n")
                
                log_text.see(tk.END)
            
            threading.Thread(target=install_thread, daemon=True).start()
        
        tb.Button(button_frame, text="Install", command=install_version, bootstyle="success").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Cancel", command=install_window.destroy, bootstyle="secondary").pack(side=LEFT, padx=5)
        
        # Pre-fill with suggested name
        version_var.set("latest")
        env_name_var.set(f"esphome_latest_{datetime.now().strftime('%Y%m%d')}")

    def remove_version(self, tree):
        """Remove selected ESPHome version"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a version to remove")
            return
        
        item = tree.item(selection[0])
        env_name = item['values'][0]
        
        if env_name == "Default":
            messagebox.showerror("Error", "Cannot remove system default ESPHome")
            return
        
        if not messagebox.askyesno("Confirm Removal", f"Remove ESPHome environment '{env_name}'?"):
            return
        
        try:
            env_path = self.versions_base_path / env_name
            if env_path.exists():
                shutil.rmtree(env_path)
            
            # Refresh list
            self.populate_version_list(tree)
            self.scan_esphome_versions()
            
            # Reset to default if removing current version
            if self.current_esphome_version.get() == env_name:
                self.current_esphome_version.set("Default")
            
            messagebox.showinfo("Success", f"Removed environment '{env_name}'")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove environment: {str(e)}")

    def set_default_version(self, tree):
        """Set selected version as default"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a version to set as active")
            return
        
        item = tree.item(selection[0])
        env_name = item['values'][0]
        
        self.current_esphome_version.set(env_name)
        self.populate_version_list(tree)  # Refresh to show new active status
        messagebox.showinfo("Success", f"Set '{env_name}' as active version")

    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.root.update_idletasks()

    def open_com_manager(self):
        """Open a COM port manager window"""
        com_window = tb.Toplevel(self.root)
        com_window.title("COM Port Manager")
        com_window.geometry("500x400")
        com_window.transient(self.root)
        com_window.grab_set()
        
        # Create frames
        main_frame = tb.Frame(com_window, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        # Button frame
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=X, pady=5)
        
        tb.Button(button_frame, text="Refresh", command=lambda: self.populate_com_list(tree), bootstyle="info").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Use Selected", 
                  command=lambda: self.use_selected_port(tree, com_window), bootstyle="success").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Close", command=com_window.destroy, bootstyle="secondary").pack(side=RIGHT, padx=5)
        
        # Treeview for COM ports
        tree_frame = tb.Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=5)
        
        columns = ("Port", "Description", "Hardware ID")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        # Define headings
        tree.heading("Port", text="Port")
        tree.heading("Description", text="Description")
        tree.heading("Hardware ID", text="Hardware ID")
        
        # Define columns
        tree.column("Port", width=100)
        tree.column("Description", width=200)
        tree.column("Hardware ID", width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Populate the list
        self.populate_com_list(tree)

    def populate_com_list(self, tree):
        """Populate the treeview with COM port information"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Get COM ports
        ports = serial.tools.list_ports.comports()
        
        # Add ports to treeview
        for port in ports:
            tree.insert("", "end", values=(port.device, port.description, port.hwid))

    def use_selected_port(self, tree, window):
        """Use the selected COM port from the manager"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a COM port first")
            return
        
        item = tree.item(selection[0])
        port_name = item['values'][0]
        self.port_var.set(port_name)
        window.destroy()
        self.status_var.set(f"Selected port: {port_name}")

    def get_current_versions(self):
        """Get current versions of installed components"""
        try:
            # Get ESPHome version
            result = subprocess.run(['esphome', 'version'], capture_output=True, text=True)
            if result.returncode == 0:
                match = re.search(r'Version:\s*([\d.]+)', result.stdout)
                if match:
                    self.current_versions['esphome'] = match.group(1)
            
            # Get Python version
            self.current_versions['python'] = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
            
            # Get PySerial version
            try:
                import serial
                self.current_versions['pyserial'] = serial.__version__
            except ImportError:
                self.current_versions['pyserial'] = "Not installed"
            
            # Get PlatformIO version - fixed to extract just the version number
            try:
                result = subprocess.run(['pio', '--version'], capture_output=True, text=True)
                if result.returncode == 0:
                    output = result.stdout.strip()
                    # Extract just the version number from PlatformIO output
                    version_match = re.search(r'(\d+\.\d+\.\d+)', output)
                    if version_match:
                        self.current_versions['platformio'] = version_match.group(1)
                    else:
                        self.current_versions['platformio'] = output  # Fallback to full output
            except (FileNotFoundError, subprocess.CalledProcessError):
                self.current_versions['platformio'] = "Not found"
                
        except Exception as e:
            self.log_message( f">>> Error getting versions: {str(e)}", "auto")

    def check_updates(self):
        """Check for updates to all components"""
        def update_check_thread():
            self.status_var.set("Checking for updates...")
            self.update_status.set("Update status: Checking...")
            
            # Get current versions if not already done
            if not any(self.current_versions.values()):
                self.get_current_versions()
            
            # Check for updates
            update_info = {}
            
            # Check ESPHome
            try:
                response = requests.get('https://pypi.org/pypi/esphome/json', timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    latest_version = data['info']['version']
                    current_version = self.current_versions['esphome'] or '0'
                    update_info['esphome'] = {
                        'current': current_version,
                        'latest': latest_version,
                        'update_available': version.parse(current_version) < version.parse(latest_version)
                    }
            except Exception as e:
                update_info['esphome'] = {'error': str(e)}
            
            # Check Python (we'll just show the current version as PyPI doesn't have Python updates)
            update_info['python'] = {
                'current': self.current_versions['python'],
                'latest': f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
                'update_available': False  # We don't check for Python updates via PyPI
            }
            
            # Check PySerial
            try:
                response = requests.get('https://pypi.org/pypi/pyserial/json', timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    latest_version = data['info']['version']
                    current_version = self.current_versions['pyserial'] or '0'
                    update_info['pyserial'] = {
                        'current': current_version,
                        'latest': latest_version,
                        'update_available': version.parse(current_version) < version.parse(latest_version)
                    }
            except Exception as e:
                update_info['pyserial'] = {'error': str(e)}
            
            # Check PlatformIO
            try:
                response = requests.get('https://pypi.org/pypi/platformio/json', timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    latest_version = data['info']['version']
                    current_version = self.current_versions['platformio'] or '0'
                    
                    # Handle the case where PlatformIO version might be a full string
                    try:
                        # Try to extract just the version number if it's a longer string
                        version_match = re.search(r'(\d+\.\d+\.\d+)', current_version)
                        if version_match:
                            current_version = version_match.group(1)
                        
                        update_available = version.parse(current_version) < version.parse(latest_version)
                    except:
                        # If version parsing fails, assume no update available
                        update_available = False
                    
                    update_info['platformio'] = {
                        'current': self.current_versions['platformio'],
                        'latest': latest_version,
                        'update_available': update_available
                    }
            except Exception as e:
                update_info['platformio'] = {'error': str(e)}
            
            # Update the UI with results
            self.last_update_check = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.show_update_results(update_info)
            
        threading.Thread(target=update_check_thread, daemon=True).start()

    def show_update_results(self, update_info):
        """Show update results in a new window"""
        # Create update window
        update_window = tb.Toplevel(self.root)
        update_window.title("Update Check Results")
        update_window.geometry("600x400")
        update_window.transient(self.root)
        
        # Create frames
        main_frame = tb.Frame(update_window, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        # Last checked label
        last_checked = tb.Label(main_frame, text=f"Last checked: {self.last_update_check}", bootstyle="info")
        last_checked.pack(anchor=W, pady=5)
        
        # Treeview for update results
        tree_frame = tb.Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=5)
        
        columns = ("Component", "Current Version", "Latest Version", "Status")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        # Define headings
        tree.heading("Component", text="Component")
        tree.heading("Current Version", text="Current Version")
        tree.heading("Latest Version", text="Latest Version")
        tree.heading("Status", text="Status")
        
        # Define columns
        tree.column("Component", width=120)
        tree.column("Current Version", width=120)
        tree.column("Latest Version", width=120)
        tree.column("Status", width=120)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Populate the tree
        for component, info in update_info.items():
            if 'error' in info:
                tree.insert("", "end", values=(
                    component.capitalize(),
                    "N/A",
                    "N/A",
                    f"Error: {info['error']}"
                ))
            else:
                status = "Update available" if info['update_available'] else "Up to date"
                tree.insert("", "end", values=(
                    component.capitalize(),
                    info['current'] or "Not found",
                    info['latest'],
                    status
                ))
        
        # Button frame
        button_frame = tb.Frame(main_frame)
        button_frame.pack(fill=X, pady=10)
        
        tb.Button(button_frame, text="Update All", 
                  command=lambda: self.update_all(update_info, update_window), 
                  bootstyle="success").pack(side=LEFT, padx=5)
        tb.Button(button_frame, text="Close", 
                  command=update_window.destroy, 
                  bootstyle="secondary").pack(side=RIGHT, padx=5)
        
        # Update status in main window
        any_updates = any('update_available' in info and info['update_available'] for info in update_info.values())
        status_text = "Updates available!" if any_updates else "All components up to date"
        self.update_status.set(f"Update status: {status_text} (Checked: {self.last_update_check})")
        self.status_var.set("Update check completed")

    def update_all(self, update_info, window):
        """Update all components that need updating"""
        needs_update = []
        
        for component, info in update_info.items():
            if 'update_available' in info and info['update_available']:
                needs_update.append(component)
        
        if not needs_update:
            messagebox.showinfo("No Updates", "All components are already up to date.")
            return
        
        # Confirm update
        confirm = messagebox.askyesno(
            "Confirm Update", 
            f"The following components will be updated:\n{', '.join(needs_update)}\n\nContinue?"
        )
        
        if not confirm:
            return
        
        # Perform updates
        def update_thread():
            for component in needs_update:
                self.status_var.set(f"Updating {component}...")
                try:
                    if component == 'esphome':
                        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "esphome"], check=True)
                    elif component == 'pyserial':
                        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pyserial"], check=True)
                    elif component == 'platformio':
                        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "platformio"], check=True)
                    
                    self.log_message( f">>> Successfully updated {component}", "auto")
                except subprocess.CalledProcessError as e:
                    self.log_message( f">>> Error updating {component}: {str(e)}", "auto")
            
            self.status_var.set("Update completed")
            messagebox.showinfo("Update Complete", "All selected components have been updated.")
            window.destroy()
            self.check_updates()  # Refresh the update status
        
        threading.Thread(target=update_thread, daemon=True).start()

    def update_esphome(self):
        """Update ESPHome specifically"""
        confirm = messagebox.askyesno(
            "Update ESPHome", 
            "This will update ESPHome to the latest version. Continue?"
        )
        
        if not confirm:
            return
        
        def update_thread():
            self.status_var.set("Updating ESPHome...")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "esphome"], check=True)
                self.log_message( ">>> Successfully updated ESPHome", "auto")
                self.status_var.set("ESPHome updated successfully")
                messagebox.showinfo("Update Complete", "ESPHome has been updated to the latest version.")
                self.check_updates()  # Refresh the update status
            except subprocess.CalledProcessError as e:
                self.log_message( f">>> Error updating ESPHome: {str(e)}", "auto")
                self.status_var.set("Error updating ESPHome")
        
        threading.Thread(target=update_thread, daemon=True).start()

    def show_about(self):
        """Show about dialog"""
        about_text = f"""
        ESPHome Compiler GUI with Version Manager
        
        Version: 1.2.0
        Python: {sys.version}
        
        This application provides a graphical interface for compiling
        and uploading ESPHome configurations with multi-version support.
        
        Components:
        - ESPHome: {self.current_versions['esphome'] or 'Not found'}
        - Python: {self.current_versions['python'] or 'Not found'}
        - PySerial: {self.current_versions['pyserial'] or 'Not found'}
        - PlatformIO: {self.current_versions['platformio'] or 'Not found'}
        
        Last update check: {self.last_update_check or 'Never'}
        Last YAML sync: {self.last_sync_time or 'Never'}
        Last backup: {self.last_backup_time or 'Never'}
        
        Active ESPHome Version: {self.current_esphome_version.get()}
        Auto backup: {'Enabled' if self.backup_enabled.get() else 'Disabled'}
        """
        
        messagebox.showinfo("About ESPHome Compiler GUI", about_text)

    def bytes_to_mb(self, bytes_str):
        """Convert bytes string to MB format"""
        try:
            bytes_val = int(bytes_str)
            mb_val = bytes_val / (1024 * 1024)
            return f"{mb_val:.2f}MB"
        except (ValueError, TypeError):
            return "N/A"

    def extract_ram_usage_from_log(self):
        """Extract RAM usage percentage from the log text"""
        log_content = self.log_text.get(1.0, tk.END)
        ram_pattern = r'RAM:\s*\[.*\]\s*([\d.]+)%'
        ram_match = re.search(ram_pattern, log_content)
        if ram_match:
            return f"{ram_match.group(1)}%"
        return "N/A"

    def extract_firmware_size(self, output):
        """Extract firmware size information from the output"""
        # Try multiple patterns to catch different ESPHome output formats
        
        # Pattern for the specific format you showed:
        # Flash: [====      ]  38.2% (used 3100650 bytes from 8126464 bytes)
        flash_pattern = r'Flash:\s*\[.*\]\s*([\d.]+)%\s*\(used\s*(\d+)\s*bytes\s*from\s*(\d+)\s*bytes\)'
        flash_match = re.search(flash_pattern, output)
        
        if flash_match:
            percentage = flash_match.group(1)
            used_bytes = flash_match.group(2)
            total_bytes = flash_match.group(3)
            
            used_mb = self.bytes_to_mb(used_bytes)
            total_mb = self.bytes_to_mb(total_bytes)
            
            self.firmware_size = used_mb
            self.firmware_max_size = total_mb
            self.firmware_percentage = percentage
            self.firmware_size_var.set(f"Firmware size: {used_mb}/{total_mb} ({percentage}%)")
            return True
        
        # Also try RAM pattern for completeness
        ram_pattern = r'RAM:\s*\[.*\]\s*([\d.]+)%\s*\(used\s*(\d+)\s*bytes\s*from\s*(\d+)\s*bytes\)'
        ram_match = re.search(ram_pattern, output)
        
        if ram_match:
            # We could display RAM usage too if wanted
            pass
        
        # Try other patterns as fallback
        other_patterns = [
            # Pattern 1: "Flash size: 1.23MB/4.00MB"
            r'Flash size:\s*([\d.]+)([KM]B)/([\d.]+)([KM]B)',
            # Pattern 2: "Total: 1.23MB"
            r'Total:\s*([\d.]+)([KM]B)',
            # Pattern 3: "Total size: 1.23MB"
            r'Total size:\s*([\d.]+)([KM]B)',
            # Pattern 4: "Firmware size: 1.23MB"
            r'Firmware size:\s*([\d.]+)([KM]B)',
            # Pattern 5: "Size: 1.23MB"
            r'Size:\s*([\d.]+)([KM]B)',
        ]
        
        for pattern in other_patterns:
            match = re.search(pattern, output)
            if match:
                if len(match.groups()) >= 4:
                    # Pattern with both current and max size
                    self.firmware_size = f"{match.group(1)}{match.group(2)}"
                    self.firmware_max_size = f"{match.group(3)}{match.group(4)}"
                    self.firmware_size_var.set(f"Firmware size: {self.firmware_size}/{self.firmware_max_size}")
                else:
                    # Pattern with only current size
                    self.firmware_size = f"{match.group(1)}{match.group(2)}"
                    self.firmware_size_var.set(f"Firmware size: {self.firmware_size}")
                return True
        return False

    def upload(self):
        """Upload firmware only"""
        if not self.validate_file():
            return
            
        def upload_thread():
            start_time = time.time()
            self.status_var.set("Uploading...")
            port = self.port_var.get().strip()
            if port:
                command = f'esphome upload --device {port} "{self.file_path.get()}"'
            else:
                command = f'esphome upload "{self.file_path.get()}"'
            success = self.run_command(command, time.time(), 60)  # Shorter estimated time for upload only
            end_time = time.time()
            duration = end_time - start_time
            
            if success:
                self.status_var.set("Upload completed")
                self.update_phase_label("Done")
                self.error_indicator.configure(bootstyle="success")
                
                # Update build history for upload - NEW
                self.update_build_history(
                    action_type='upload',
                    duration=duration,
                    firmware_size=self.firmware_size,
                    flash_usage=f"{self.firmware_percentage}%" if self.firmware_percentage != "N/A" else "N/A",
                    ram_usage=self.extract_ram_usage_from_log()
                )
            else:
                self.status_var.set("Upload failed")
                self.update_phase_label("Upload failed")
                self.error_indicator.configure(bootstyle="danger")
            
        threading.Thread(target=upload_thread, daemon=True).start()

    def should_filter_upload_progress(self, output):
        """Check if the output line should be filtered (upload progress spam)"""
        # Look for upload progress patterns - various formats ESPHome might use
        upload_patterns = [
            r'\[\s*\d+%\]',  # [ 10%]
            r'\d+%\s*\|',    # 10% |
            r'Uploading.*\d+%',  # Uploading... 10%
            r'\d+\.\d+%\s*',     # 10.5%
        ]
        
        for pattern in upload_patterns:
            if re.search(pattern, output):
                # Extract the percentage number
                percent_match = re.search(r'(\d+)(?:\.\d+)?%', output)
                if percent_match:
                    current_percent = int(percent_match.group(1))
                    # Only show every 10% or when it changes significantly
                    if current_percent >= self.last_upload_progress + 10 or current_percent == 100:
                        self.last_upload_progress = current_percent
                        return False  # Don't filter this one - show it
                    return True  # Filter this one - don't show it
        return False  # Don't filter other lines

    def smart_sync_before_compile(self):
        """Smart sync before compilation - only syncs current YAML and referenced files"""
        current_file = self.file_path.get() if self.file_path.get() else None
        
        if not current_file:
            return []  # No file selected, nothing to sync
        
        # Use a thread-safe way to get the result
        result_container = []
        
        def smart_sync_thread():
            self.status_var.set("Smart syncing current project...")
            
            synced_files = sync_esphome_files_fast(
                r"\\192.168.4.76\config\esphome", 
                r"C:\esphome",
                self.backup_base_path if self.backup_enabled.get() else None,
                current_file  # Only sync files needed for this YAML
            )
            
            self.last_sync_time = datetime.now().strftime("%H:%M:%S")
            result_container.extend(synced_files)  # Store result
            
            if synced_files:
                current_filename = os.path.basename(current_file)
                if current_filename in synced_files:
                    self.sync_status_var.set(f"Sync: Smart ({self.last_sync_time})")
                    self.sync_indicator.configure(bootstyle="success")
                    self.status_var.set(f"Smart synced {len(synced_files)} file(s)")
                else:
                    self.sync_status_var.set(f"Sync: Smart {len(synced_files)} files")
                    self.sync_indicator.configure(bootstyle="info")
                    self.status_var.set(f"Smart synced {len(synced_files)} file(s)")
            else:
                self.status_var.set("Smart sync: No files needed syncing")
        
        # Run sync and wait for completion (this should be fast now)
        sync_thread = threading.Thread(target=smart_sync_thread, daemon=True)
        sync_thread.start()
        sync_thread.join(timeout=15)  # Wait max 15 seconds for sync
        
        return result_container

    def manual_full_sync(self):
        """Manual full sync - use when you want to update all resources"""
        def full_sync_thread():
            self.status_var.set("Performing full manual sync...")
            self.log_message( ">>> Starting full manual sync...", "auto")
            
            synced_files = sync_esphome_files(
                r"\\192.168.4.76\config\esphome", 
                r"C:\esphome",
                None  # No backups during full sync
            )
            
            self.last_sync_time = datetime.now().strftime("%H:%M:%S")
            
            if synced_files:
                self.sync_status_var.set(f"Sync: Full ({len(synced_files)} files)")
                self.sync_indicator.configure(bootstyle="success")
                self.log_message( f">>> Full sync completed: {len(synced_files)} files synchronized", "auto")
            else:
                self.sync_status_var.set("Sync: Full (no changes)")
                self.sync_indicator.configure(bootstyle="success")
                self.log_message( ">>> Full sync completed: No changes needed", "auto")
            
            self.status_var.set("Full sync completed")
            self.log_text.see(tk.END)
        
        threading.Thread(target=full_sync_thread, daemon=True).start()

    def update_recent_files_dropdown(self):
        """Update the recent files dropdown menu"""
        # Clear the current menu
        self.recent_files_menu.delete(0, 'end')
        
        if not self.recent_files:
            self.recent_files_menu.add_command(label="No recent files", state="disabled")
            self.recent_files_btn.configure(state="disabled")
        else:
            self.recent_files_btn.configure(state="normal")
            for i, file_path in enumerate(self.recent_files):
                # Show just the filename in the menu but store full path
                filename = os.path.basename(file_path)
                self.recent_files_menu.add_command(
                    label=f"{i+1}. {filename}",
                    command=lambda path=file_path: self.select_recent_file(path)
                )
            # Add separator and clear option
            self.recent_files_menu.add_separator()
            self.recent_files_menu.add_command(
                label="Clear Recent Files",
                command=self.clear_recent_files
            )

    def add_to_recent_files(self, file_path):
        """Add a file to recent files list"""
        if file_path in self.recent_files:
            # Remove if already exists (will be re-added at top)
            self.recent_files.remove(file_path)
        
        # Add to beginning of list
        self.recent_files.insert(0, file_path)
        
        # Keep only the most recent files
        if len(self.recent_files) > self.max_recent_files:
            self.recent_files = self.recent_files[:self.max_recent_files]
        
        # Update the dropdown
        self.update_recent_files_dropdown()

    def select_recent_file(self, file_path):
        """Select a file from the recent files list"""
        if os.path.exists(file_path):
            self.file_path.set(file_path)
            self.status_var.set(f"Selected: {os.path.basename(file_path)}")
            self.scan_ips()  # Auto-trigger IP scan
            self.check_sync_status()  # Check sync status for the selected file
            self.get_device_info_for_selected_file()  # Get device information
        else:
            # File no longer exists, remove from recent files
            if file_path in self.recent_files:
                self.recent_files.remove(file_path)
                self.update_recent_files_dropdown()
            messagebox.showwarning("File Not Found", f"The file no longer exists:\n{file_path}")

    def clear_recent_files(self):
        """Clear all recent files"""
        self.recent_files = []
        self.update_recent_files_dropdown()
        self.status_var.set("Recent files cleared")

    def save_recent_files(self):
        """Save recent files to a configuration file"""
        try:
            config_dir = os.path.expanduser("~/.esphome_studio")
            os.makedirs(config_dir, exist_ok=True)
            config_file = os.path.join(config_dir, "recent_files.json")
            
            with open(config_file, 'w') as f:
                json.dump(self.recent_files, f)
        except Exception as e:
            print(f"Could not save recent files: {e}")

    def load_recent_files(self):
        """Load recent files from configuration file"""
        try:
            config_file = os.path.expanduser("~/.esphome_studio/recent_files.json")
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    loaded_files = json.load(f)
                    # Only keep files that still exist
                    self.recent_files = [f for f in loaded_files if os.path.exists(f)]
                    self.update_recent_files_dropdown()
        except Exception as e:
            print(f"Could not load recent files: {e}")

    def on_closing(self):
        """Save recent files when application closes"""
        self.save_recent_files()
        self.root.quit()

    def compile(self):
        """Compile firmware with smart sync"""
        # Smart sync before compile (waits for completion since it should be fast)
        synced_files = self.smart_sync_before_compile()
        
        # Auto-backup current file if enabled
        if self.file_path.get() and os.path.exists(self.file_path.get()):
            self.auto_backup_file(self.file_path.get())
        
        if not self.validate_file():
            return
              
        def compile_thread():
            # Start unified timer
            start_time = time.time()
            estimated_total = 120  # Adjust as needed
            self.start_timer()

            # Reset firmware size display
            self.firmware_size_var.set("Firmware size: N/A")
            self.firmware_size = "N/A"
            self.firmware_max_size = "N/A"
            self.firmware_percentage = "N/A"

            self.status_var.set("Compiling...")
            command = f'esphome compile "{self.file_path.get()}"'
            success = self.run_command(command, start_time, estimated_total)
            end_time = time.time()
            duration = end_time - start_time
            self.stop_timer()

            # If compilation failed, show error in status
            if not success:
                self.status_var.set("Compilation failed")
                self.update_phase_label("Compile failed")
                self.error_indicator.configure(bootstyle="danger")
            else:
                self.status_var.set("Compilation completed")
                self.update_phase_label("Done")
                self.error_indicator.configure(bootstyle="success")
                
                # Update build history
                self.update_build_history(
                    action_type='compile',
                    duration=duration,
                    firmware_size=self.firmware_size,
                    flash_usage=f"{self.firmware_percentage}%" if self.firmware_percentage != "N/A" else "N/A",
                    ram_usage=self.extract_ram_usage_from_log()
                )

        threading.Thread(target=compile_thread, daemon=True).start()

    def compile_and_upload(self):
        """Compile and upload firmware with smart sync"""
        self.inst_ver_var.set(" ")
        
        # IMMEDIATE VISUAL FEEDBACK
        self.status_var.set("Starting compile & upload process...")
        self.update_phase_label("Starting...")
        self.update_progress(5)
        self.error_indicator.configure(bootstyle="info")
        self.start_process_spinner("Initializing...")
        self.log_message( ">>> 🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪", "auto")
        self.log_message( ">>> 🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪", "auto")
        self.log_message( ">>> 🟪🟪🟪🟪🟪🟪🟪  Starting compile & upload process...  🟪🟪🟪🟪🟪🟪🟪", "auto")
        self.log_message( ">>> 🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪", "auto")
        self.log_message( ">>> 🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪🟪", "auto")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        if not self.validate_file():
            self.stop_process_spinner()
            return
        
        def compile_upload_thread():
            try:
                # Set running flag - ONLY ONCE at the start
                self.is_running = True
                self.log_message( ">>> DEBUG: Thread started, is_running = True", "auto")
                
                # Check if user stopped already
                if not self.is_running:
                    self.stop_process_spinner()
                    return
                
                # Smart sync
                self.status_var.set("Smart syncing files...")
                self.update_phase_label("Syncing...")
                self.update_progress(10)
                self.update_process_status("Syncing files...")
                
                current_file = self.file_path.get() if self.file_path.get() else None
                if current_file:
                    synced_files = sync_esphome_files_fast(
                        r"\\192.168.4.76\config\esphome", 
                        r"C:\esphome",
                        self.backup_base_path if self.backup_enabled.get() else None,
                        current_file
                    )
                    self.last_sync_time = datetime.now().strftime("%H:%M:%S")
                    self.log_message( f">>> SYNC UPDATE: Synced the following: {synced_files}", "auto")
                
                if not self.is_running:
                    self.log_message( ">>> DEBUG: Stopped after sync", "auto")
                    self.stop_process_spinner()
                    return
                
                # Auto-backup
                if self.file_path.get() and os.path.exists(self.file_path.get()):
                    self.status_var.set("Creating backup...")
                    self.update_phase_label("Backing up...")
                    self.update_progress(15)
                    self.update_process_status("Creating backup...")
                    self.auto_backup_file(self.file_path.get())
                
                if not self.is_running:
                    self.log_message( ">>> DEBUG: Stopped after backup", "auto")
                    self.stop_process_spinner()
                    return
                
                # COMPILE PHASE
                self.status_var.set("Compiling...")
                self.update_phase_label("Compiling...")
                self.update_progress(20)
                self.update_process_status("Compiling firmware...")
                self.log_message( ">>> Starting compilation...", "auto")
                self.log_text.see(tk.END)

                # Reset firmware size display
                self.firmware_size_var.set("Firmware size: N/A")

                yaml_path = self.file_path.get()
                start_time = time.time()
                estimated_total = 120
                self.start_timer()
                
                # Compile phase
                compile_command = f'esphome compile "{yaml_path}"'
                self.log_message( f">>> DEBUG: Running compile command: {compile_command}", "auto")
                compile_success = self.run_command(compile_command, start_time, estimated_total)

                self.log_message( f">>> DEBUG: Compile success result: {compile_success}", "auto")
                self.log_message( f">>> DEBUG: Is running after compile: {self.is_running}", "auto")
                
                # CHECK is_running BEFORE proceeding to upload
                if not self.is_running:
                    self.log_message( ">>> DEBUG: Process was stopped during compile", "auto")
                    self.stop_process_spinner()
                    return

                if not compile_success:
                    self.log_message( ">>> DEBUG: Compile failed, checking if we should continue anyway...", "auto")
                    # Check if firmware actually exists despite the failure flag
                    firmware_bin_path = os.path.join(os.path.dirname(yaml_path), ".esphome", "build", os.path.splitext(os.path.basename(yaml_path))[0], "firmware.bin")
                    firmware_exists = os.path.exists(firmware_bin_path)
                    self.log_message( f">>> DEBUG: Firmware exists: {firmware_exists} at {firmware_bin_path}", "auto")
                    
                    if not firmware_exists:
                        self.status_var.set("Compile failed")
                        self.update_phase_label("Compile failed")
                        self.error_indicator.configure(bootstyle="danger")
                        self.stop_process_spinner()
                        return
                    else:
                        self.log_message( ">>> DEBUG: Firmware exists, forcing continue to upload", "auto")

                # UPLOAD PHASE
                self.log_message( ">>> DEBUG: Starting upload phase", "auto")
                self.status_var.set("Uploading...")
                self.update_phase_label("Uploading...")
                self.update_progress(70)
                self.update_process_status("Uploading to device...")
                self.log_message( ">>> Starting upload...", "auto")
                self.log_text.see(tk.END)

                mode = self.upload_mode_var.get()
                if mode == "COM":
                    port = self.port_var.get().strip()
                    upload_command = f'esphome upload "{yaml_path}"'
                    if port:
                        upload_command = f'esphome upload --device {port} "{yaml_path}"'
                    self.log_message( f">>> DEBUG: COM upload command: {upload_command}", "auto")
                else:  # OTA mode
                    ip = self.ota_ip_var.get().strip()
                    self.log_message( f">>> DEBUG: OTA IP: {ip}", "auto")
                    if not is_ota_device_available(ip):
                        self.status_var.set("OTA device not reachable")
                        self.update_phase_label("OTA check failed")
                        self.error_indicator.configure(bootstyle="warning")
                        self.stop_process_spinner()
                        return
                    upload_command = f'esphome upload --device {ip} "{yaml_path}"'
                    self.log_message( f">>> DEBUG: OTA upload command: {upload_command}", "auto")

                self.log_text.see(tk.END)

                upload_success = self.run_command(upload_command, start_time, estimated_total)
                end_time = time.time()
                duration = end_time - start_time
                self.stop_timer()

                self.log_message( f">>> DEBUG: Upload success result: {upload_success}", "auto")

                if upload_success and self.is_running:
                    self.status_var.set("Upload completed successfully")
                    self.update_phase_label("Done")
                    self.error_indicator.configure(bootstyle="success")
                    self.update_process_status("Complete")
                    self.update_progress(100)  # FORCE 100% at the end
                    
                    # Update build history
                    self.update_build_history(
                        action_type='upload',
                        duration=duration,
                        firmware_size=self.firmware_size,
                        flash_usage=f"{self.firmware_percentage}%" if self.firmware_percentage != "N/A" else "N/A",
                        ram_usage=self.extract_ram_usage_from_log()
                    )
                    
                    self.log_message( ">>> Upload completed successfully", "auto")
                else:
                    self.status_var.set("Upload failed")
                    self.update_phase_label("Upload failed")
                    self.error_indicator.configure(bootstyle="danger")
                    self.update_process_status("Upload failed")
                    self.update_progress(100)  # FORCE 100% even on failure

                # Only set is_running to False when we're completely done
                self.is_running = False
                self.stop_process_spinner()
                
            except Exception as e:
                self.log_message( f">>> Thread error: {str(e)}", "auto")
                self.log_text.see(tk.END)
                self.is_running = False
                self.stop_process_spinner()
        
        # Start the main thread
        self.compile_thread = threading.Thread(target=compile_upload_thread, daemon=True)
        self.compile_thread.start()

    def upload(self):
        """Upload firmware only"""
        if not self.validate_file():
            return
            
        def upload_thread():
            start_time = time.time()
            self.status_var.set("Uploading...")
            port = self.port_var.get().strip()
            if port:
                command = f'esphome upload --device {port} "{self.file_path.get()}"'
            else:
                command = f'esphome upload "{self.file_path.get()}"'
            success = self.run_command(command, time.time(), 60)  # Shorter estimated time for upload only
            end_time = time.time()
            duration = end_time - start_time
            
            if success:
                self.status_var.set("Upload completed")
                self.update_phase_label("Done")
                self.error_indicator.configure(bootstyle="success")
                
                # Update build history for upload
                self.update_build_history(
                    action_type='upload',
                    duration=duration,
                    firmware_size=self.firmware_size,
                    flash_usage=f"{self.firmware_percentage}%" if self.firmware_percentage != "N/A" else "N/A",
                    ram_usage=self.extract_ram_usage_from_log()
                )
            else:
                self.status_var.set("Upload failed")
                self.update_phase_label("Upload failed")
                self.error_indicator.configure(bootstyle="danger")
            
        threading.Thread(target=upload_thread, daemon=True).start()

    def run_command(self, command, start_time, estimated_total):
        """Run a command with the selected ESPHome version"""
        # Replace 'esphome' with the correct path/command
        esphome_cmd = self.get_esphome_command()
        
        # Update the command to use the selected version
        if command.startswith('esphome '):
            command = command.replace('esphome ', f'{esphome_cmd} ', 1)
        
        try:
            self.log_message( f">>> Using: {self.current_esphome_version.get()} ({esphome_cmd})", "auto")
            self.log_message( f">>> {command}", "auto")
            self.log_text.see(tk.END)

            # Reset upload progress tracking
            self.last_upload_progress = 0
            # DON'T set is_running here - it should already be set by the calling method
            # self.is_running = True  # REMOVE THIS LINE

            # Set environment
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            
            if sys.platform == 'win32':
                env['PYTHONUTF8'] = '1'

            self.current_process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                shell=True,
                text=True,
                bufsize=1,
                universal_newlines=True,
                env=env,
                encoding='utf-8',
                errors='replace'
            )

            # Track compilation success
            compilation_successful = False
            saw_firmware_bin = False
            upload_complete = False
            
            # Simple output reading
            while True:
                # CHECK FOR STOP REQUEST
                if not self.is_running:
                    self.current_process.terminate()
                    self.log_message( ">>> Process stopped by user", "auto")
                    break
                    
                output = self.current_process.stdout.readline()
                if output == '' and self.current_process.poll() is not None:
                    break
                    
                if output:
                    output = output.strip()
                    
                    # Extract firmware size information
                    self.extract_firmware_size(output)
                    
                    # IMPROVED SUCCESS DETECTION
                    if "[SUCCESS]" in output:
                        self.log_message( ">>> DEBUG: Detected [SUCCESS] - compilation complete!", "auto")
                        compilation_successful = True
                        self.update_progress(100)  # Force progress to 100%
                    elif "Successfully compiled program." in output:
                        self.log_message( ">>> DEBUG: Detected compilation success message!", "auto")
                        compilation_successful = True
                        self.update_progress(100)  # Force progress to 100%
                    elif "Building" in output and "firmware.bin" in output:
                        self.update_phase_label("Building...")
                        self.update_process_status("Building firmware...")
                        saw_firmware_bin = True
                    elif "Linking" in output and "firmware.elf" in output:
                        self.update_phase_label("Linking...")
                        self.update_process_status("Linking firmware...")
                    elif "Uploading" in output:
                        self.update_phase_label("Uploading...")
                        self.update_process_status("Uploading to device...")
                    elif "Successfully uploaded" in output or "OTA successful" in output:
                        self.update_phase_label("Done")
                        self.update_process_status("Upload complete")
                        self.update_progress(100)  # Force progress to 100%
                        upload_complete = True
                        return True
                    elif "Compilation failed" in output:
                        self.update_phase_label("Failed")
                        self.update_process_status("Compilation failed")
                        return False
                    
                    # DETECT UPLOAD COMPLETION
                    if "Upload took" in output and "waiting for result" in output:
                        self.update_progress(95)  # Almost done
                    elif "Done..." in output and not upload_complete:
                        self.update_progress(99)  # Very close to done
                    
                    # IGNORE esp_idf_size warnings completely
                    if "esp_idf_size:" in output:
                        self.log_message( ">>> (Ignoring size tool warning)", "auto")
                        continue
                    
                    # Filter upload progress spam but update progress
                    if self.should_filter_upload_progress(output):
                        percent_match = re.search(r'(\d+)(?:\.\d+)?%', output)
                        if percent_match:
                            current_percent = int(percent_match.group(1))
                            self.update_progress(current_percent)
                        continue
                    
                    self.log_message( output + "", "auto")
                    self.log_text.see(tk.END)
                    self.root.update_idletasks()

                    # Update progress based on time as fallback
                    elapsed = time.time() - start_time
                    progress = min(99, (elapsed / estimated_total) * 100)  # Cap at 99% until complete
                    
                    percent_match = re.search(r'(\d+)(?:\.\d+)?%', output)
                    if percent_match:
                        progress = int(percent_match.group(1))
                    
                    self.update_progress(progress)

            return_code = self.current_process.poll()
            
            # FORCE 100% PROGRESS if we detected success
            if compilation_successful or saw_firmware_bin or upload_complete:
                self.update_progress(100)
                self.log_message( ">>> Compilation successful - proceeding to upload", "auto")
                return True
            
            self.update_progress(100)  # Still set to 100% even on failure for clean UI
            self.log_message( f">>> Compilation may have failed - return code: {return_code}", "auto")
            return return_code == 0

        except Exception as e:
            self.log_message( f">>> Error: {str(e)}", "auto")
            self.status_var.set(f"Error: {str(e)}")
            self.update_phase_label("Error")
            self.error_indicator.configure(bootstyle="danger")
            self.update_progress(100)  # Set to 100% even on error
            return False

    
    def _read_available_output(self, stream):
        """Read available output without blocking"""
        import msvcrt  # Windows-specific
        try:
            # Check if data is available
            if msvcrt.kbhit() or True:  # Always try to read for now
                line = stream.readline()
                return line if line else None
        except:
            # Fallback: try to read anyway
            try:
                line = stream.readline()
                return line if line else None
            except:
                return None
        return None

    def _force_kill_process(self):
        """Force kill the process tree on Windows"""
        try:
            if os.name == 'nt':  # Windows
                # Use taskkill to terminate the entire process tree
                subprocess.run(f'taskkill /F /T /PID {self.current_process.pid}', 
                            shell=True, capture_output=True, timeout=5)
            else:
                self.current_process.kill()
        except Exception as e:
            self.log_message( f">>> Warning: {str(e)}", "auto")

    def update_phase_label(self, text):
        """Update the phase label with logging"""
        if hasattr(self, "phase_label") and self.phase_label is not None:
            try:
                print(f"PHASE CHANGE: {text}")  # Debug logging
                self.phase_label.config(text=text)
            except Exception as e:
                print(f"Warning: could not update phase label: {e}")

    def update_process_status(self, status_text):
        """Update the process status without changing the spinner"""
        if hasattr(self, 'spinner_running') and self.spinner_running:
            self.current_process_name = status_text
            spinner_char = self.spinner_steps[self.spinner_index % len(self.spinner_steps)]
            print(f"STATUS CHANGE: {status_text}")  # Debug logging
            self.process_status_var.set(f"{spinner_char} {status_text}")

    def validate_file(self):
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a YAML file first")
            return False
        if not os.path.isfile(file_path):
            messagebox.showerror("Error", "The selected file does not exist")
            return False
        if not file_path.lower().endswith(('.yaml', '.yml')):
            messagebox.showwarning("Warning", "The selected file doesn't appear to be a YAML file")
        return True

    def log_message(self, message, message_type="auto"):
        """
        Add a message to the log with colored formatting
        Auto-detects message type if set to "auto"
        """
        if message_type == "auto":
            # Auto-detect message type based on content
            message_lower = message.lower()
            
            if any(cmd in message_lower for cmd in [">>>", "debug:", "starting", "completed", "successfully"]):
                message_type = "script"
            elif "error" in message_lower or "failed" in message_lower:
                message_type = "error"
            elif "warning" in message_lower:
                message_type = "warning"
            elif "success" in message_lower:
                message_type = "success"
            elif any(cmd in message_lower for cmd in ["command:", "running:", "compile", "upload"]):
                message_type = "command"
            elif "esphome" in message_lower:
                message_type = "esphome"
            else:
                message_type = "esphome"  # Default for ESPHome output
        
        self.log_text.insert(tk.END, message + "\n", message_type)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        # Reset firmware size display when clearing log
        self.firmware_size_var.set("Firmware size: N/A")
        self.inst_ver_var.set(" ")
        self.timer_var.set("00:00")
        self.timer_running = False
        self.error_indicator.configure(bootstyle="secondary")

    def start_timer(self):
        self.timer_start = time.time()
        self.timer_running = True
        self.update_timer()

    def stop_timer(self):
        self.timer_running = False

    def update_timer(self):
        if self.timer_running:
            elapsed = int(time.time() - self.timer_start)
            mins, secs = divmod(elapsed, 60)
            self.timer_var.set(f"{mins:02}:{secs:02}")
            self.root.after(1000, self.update_timer)


    def toggle_theme(self):
        """Toggle between light and dark themes"""
        if self.current_theme == "darkly":
            self.current_theme = "morph"  # Light theme
        else:
            self.current_theme = "darkly"  # Dark theme
            
        self.root.style.theme_use(self.current_theme)

    def setup_context_menu(self):
        """Setup Windows context menu integration"""
        try:
            messagebox.showinfo("Context Menu", 
                              "Context menu setup requires administrator privileges.\n\n"
                              "This feature would add 'Open with ESPHome Studio' to the\n"
                              "right-click context menu for YAML files.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to setup context menu: {e}")

    def stop_process(self):
        """Stop the current running process immediately - enhanced version"""
        try:
            self.is_running = False
            self.spinner_running = False  # Stop spinner immediately
            
            # Stop any subprocess
            if hasattr(self, 'current_process') and self.current_process:
                self._force_kill_process()
            
            # Stop device info collection if running
            if hasattr(self, 'device_info_thread') and self.device_info_thread:
                # We can't easily stop the device info thread, but we can mark it for stop
                self.is_running = False
            
            # Stop timers
            self.stop_timer()
            
            # Update UI immediately
            self.status_var.set("Process stopped by user")
            self.update_phase_label("Stopped")
            self.update_progress(0)
            self.process_status_var.set("Stopped")
            self.error_indicator.configure(bootstyle="warning")
            
            # Add to log
            self.log_message( ">>> Process stopped by user", "auto")
            self.log_text.see(tk.END)
            
            # Force UI update
            self.root.update_idletasks()
            
        except Exception as e:
            self.log_message( f">>> Error stopping process: {str(e)}", "auto")
            self.log_text.see(tk.END)

    def clean_build(self):
        """Clean the build directory to fix compilation issues"""
        if not self.validate_file():
            return
            
        def clean_thread():
            self.status_var.set("Cleaning build...")
            yaml_path = self.file_path.get()
            build_dir = os.path.join(os.path.dirname(yaml_path), ".esphome", "build")
            
            if os.path.exists(build_dir):
                try:
                    shutil.rmtree(build_dir)
                    self.log_message( ">>> Build directory cleaned successfully", "auto")
                    self.status_var.set("Build directory cleaned")
                except Exception as e:
                    self.log_message( f">>> Error cleaning build: {str(e)}", "auto")
                    self.status_var.set("Error cleaning build")
            else:
                self.log_message( ">>> No build directory found", "auto")
                self.status_var.set("No build directory found")
        
        threading.Thread(target=clean_thread, daemon=True).start()

    def backup_current_file(self):
        """Create a backup of the current YAML file"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
            
        def backup_thread():
            self.status_var.set("Creating backup...")
            backup_path = create_backup(
                self.file_path.get(), 
                self.backup_base_path,
                os.path.basename(self.file_path.get())
            )
            
            if backup_path:
                self.last_backup_time = datetime.now().strftime("%H:%M:%S")
                self.backup_status_var.set(f"Backup: {self.last_backup_time}")
                self.log_message( f">>> Backup created: {backup_path}", "auto")
                self.status_var.set("Backup created successfully")
                
                # Clean up old backups
                cleanup_old_backups(self.backup_base_path, self.max_backups.get())
            else:
                self.log_message( ">>> Backup failed", "auto")
                self.status_var.set("Backup failed")
        
        threading.Thread(target=backup_thread, daemon=True).start()

def main():
    # Create root window with ttkbootstrap theme
    root = tb.Window(themename="darkly")
    
    # Set window icon if available
    try:
        root.iconbitmap("esphome.ico")  # You might want to add an icon
    except:
        pass
    
    # Create and run the application
    app = ModernESPHomeGUI(root)
    
    root.protocol("WM_DELETE_WINDOW", app.on_closing)

    # Start the GUI
    root.mainloop()

if __name__ == "__main__":
    main()