# ADD:
#       schedule upload at time, compile earlier
#       long term storage of device data, notice device change
#       batch processing
# This is an esphome installation manager. It can compile and install different versions of esphome and includes device data reading and storing and a backup system.
# It is designed to keep the local pc's version of the yaml files synced with the HA server side. The HA server is the truth. Yaml files from the local pc are never copied to the HA server.

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess
from subprocess import Popen, PIPE, TimeoutExpired
import threading
from threading import Lock
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
from datetime import datetime, timedelta
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

import schedule
import tempfile
import zipfile
import sqlite3

class ESPHomeListener(ServiceListener):
    def __init__(self):
        self.devices = []  # List of (name, ip) tuples

    def add_service(self, zeroconf, type, name):
        info = zeroconf.get_service_info(type, name)
        if info and info.addresses:
            ip = ".".join(str(b) for b in info.addresses[0])
            device_name = name.split('.')[0]  # Extract name from full mDNS name
            self.devices.append((device_name, ip))

class ESPHomeDataManager:
    def __init__(self):
        self.data_dir = Path.home() / ".esphome_studio"
        self.data_dir.mkdir(exist_ok=True)
        self.devices_file = self.data_dir / "devices.json"
        self.history_file = self.data_dir / "upload_history.json"
        
    def get_yaml_key(self, yaml_path):
        """Create a unique key for a YAML file"""
        return str(Path(yaml_path).resolve())
    
    def load_devices_data(self):
        """Load all devices data"""
        try:
            if self.devices_file.exists():
                with open(self.devices_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading devices data: {e}")
        return {}
    
    def load_history_data(self):
        """Load all upload history data"""
        try:
            if self.history_file.exists():
                with open(self.history_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading history data: {e}")
        return {}
    
    def save_device_info(self, yaml_path, device_info):
        """Save device information for a YAML file"""
        try:
            data = self.load_devices_data()
            key = self.get_yaml_key(yaml_path)
            data[key] = {
                'device_info': device_info,
                'last_updated': datetime.now().isoformat(),
                'yaml_file': os.path.basename(yaml_path)
            }
            with open(self.devices_file, 'w') as f:
                json.dump(data, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving device info: {e}")
            return False
    
    def save_upload_history(self, yaml_path, upload_history):
        """Save upload history for a YAML file"""
        try:
            data = self.load_history_data()
            key = self.get_yaml_key(yaml_path)
            data[key] = {
                'upload_history': upload_history,
                'last_updated': datetime.now().isoformat(),
                'yaml_file': os.path.basename(yaml_path)
            }
            with open(self.history_file, 'w') as f:
                json.dump(data, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving upload history: {e}")
            return False
    
    def get_device_info(self, yaml_path):
        """Get stored device information for a YAML file"""
        data = self.load_devices_data()
        key = self.get_yaml_key(yaml_path)
        return data.get(key, {}).get('device_info', None)
    
    def get_upload_history(self, yaml_path):
        """Get stored upload history for a YAML file"""
        data = self.load_history_data()
        key = self.get_yaml_key(yaml_path)
        return data.get(key, {}).get('upload_history', None)
    
    def delete_device_info(self, yaml_path):
        """Delete device information for a YAML file"""
        try:
            data = self.load_devices_data()
            key = self.get_yaml_key(yaml_path)
            if key in data:
                del data[key]
                with open(self.devices_file, 'w') as f:
                    json.dump(data, f, indent=2)
                return True
        except Exception as e:
            print(f"Error deleting device info: {e}")
        return False
    
    def delete_upload_history(self, yaml_path):
        """Delete upload history for a YAML file"""
        try:
            data = self.load_history_data()
            key = self.get_yaml_key(yaml_path)
            if key in data:
                del data[key]
                with open(self.history_file, 'w') as f:
                    json.dump(data, f, indent=2)
                return True
        except Exception as e:
            print(f"Error deleting upload history: {e}")
        return False
    
    def delete_all_data(self):
        """Delete all stored data"""
        try:
            if self.devices_file.exists():
                self.devices_file.unlink()
            if self.history_file.exists():
                self.history_file.unlink()
            return True
        except Exception as e:
            print(f"Error deleting all data: {e}")
        return False
    
    def get_stats(self):
        """Get statistics about stored data"""
        devices_data = self.load_devices_data()
        history_data = self.load_history_data()
        
        return {
            'total_devices': len(devices_data),
            'total_histories': len(history_data),
            'devices_size': self.devices_file.stat().st_size if self.devices_file.exists() else 0,
            'history_size': self.history_file.stat().st_size if self.history_file.exists() else 0,
            'data_dir': str(self.data_dir)
        }

class DelayedUploadManager:
    def __init__(self, data_dir):
        self.data_dir = Path(data_dir) / "delayed_uploads"
        self.data_dir.mkdir(exist_ok=True)
        self.db_file = self.data_dir / "uploads.db"
        self.compile_cache_dir = self.data_dir / "compile_cache"
        self.compile_cache_dir.mkdir(exist_ok=True)
        self.lock = Lock()
        self.init_database()
        
    def init_database(self):
        """Initialize the SQLite database for delayed uploads"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS delayed_uploads (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    yaml_path TEXT NOT NULL,
                    yaml_filename TEXT NOT NULL,
                    compiled_firmware_path TEXT,
                    target_device TEXT NOT NULL,
                    upload_mode TEXT NOT NULL,
                    scheduled_time DATETIME NOT NULL,
                    status TEXT NOT NULL,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    esphome_version TEXT,
                    device_info TEXT,
                    upload_history TEXT,
                    compile_mode TEXT NOT NULL DEFAULT 'at_upload',  -- NEW: 'at_schedule' or 'at_upload'
                    compile_status TEXT,  -- NEW: 'pending', 'success', 'failed'
                    compile_output TEXT,  -- NEW: Store compilation output for debugging
                    last_compile_attempt DATETIME
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS batch_groups (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    description TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS batch_uploads (
                    batch_id INTEGER,
                    upload_id INTEGER,
                    FOREIGN KEY (batch_id) REFERENCES batch_groups (id),
                    FOREIGN KEY (upload_id) REFERENCES delayed_uploads (id)
                )
            ''')
            conn.commit()
            conn.close()
    
    def store_upload_job(self, yaml_path, target_device, upload_mode, scheduled_time, 
                        esphome_version=None, device_info=None, upload_history=None,
                        compile_mode='at_upload'):  # NEW: Added compile_mode parameter
        """Store a delayed upload job with compile mode option"""
        try:
            with self.lock:
                # Create a unique identifier for this job
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                yaml_filename = os.path.basename(yaml_path)
                job_id = f"{os.path.splitext(yaml_filename)[0]}_{timestamp}"
                
                # Store the YAML file and dependencies
                job_dir = self.compile_cache_dir / job_id
                job_dir.mkdir(exist_ok=True)
                
                # Copy YAML file and referenced files
                self._copy_yaml_and_dependencies(yaml_path, job_dir)
                
                compiled_firmware_path = None
                compile_status = 'pending'  # All start as pending, user compiles via Compile Now button
                compile_output = None
                last_compile_attempt = None
                
                # NOTE: We don't compile here anymore - that was blocking the UI
                # User should click "Compile Now" button to trigger compilation
                
                # Store in database
                conn = sqlite3.connect(self.db_file)
                cursor = conn.cursor()
                
                cursor.execute('''
                    INSERT INTO delayed_uploads 
                    (yaml_path, yaml_filename, compiled_firmware_path, target_device, 
                     upload_mode, scheduled_time, status, esphome_version, device_info, 
                     upload_history, compile_mode, compile_status, compile_output, last_compile_attempt)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    yaml_path, yaml_filename, compiled_firmware_path, target_device,
                    upload_mode, scheduled_time, 'scheduled', 
                    esphome_version, 
                    json.dumps(device_info) if device_info else None,
                    json.dumps(upload_history) if upload_history else None,
                    compile_mode,  # NEW
                    compile_status,  # NEW
                    compile_output,  # NEW
                    last_compile_attempt  # NEW
                ))
                
                upload_id = cursor.lastrowid
                conn.commit()
                conn.close()
                
                return upload_id
                
        except Exception as e:
            print(f"Error storing upload job: {e}")
            return None

    def _update_compile_status(self, job_id, status, output=None):
        """Update compile status for a job"""
        # This can be used for real-time status updates if needed
        pass
    
    def _copy_yaml_and_dependencies(self, yaml_path, target_dir):
        """Copy YAML file and all referenced files"""
        try:
            # Copy main YAML file
            shutil.copy2(yaml_path, target_dir / os.path.basename(yaml_path))
            
            # Copy referenced files (images, fonts, includes, etc.)
            referenced_files = get_referenced_files(yaml_path)
            yaml_dir = os.path.dirname(yaml_path)
            
            for file_ref in referenced_files:
                src_path = os.path.join(yaml_dir, file_ref)
                if os.path.exists(src_path):
                    # Create target subdirectory if needed
                    target_file_path = target_dir / file_ref
                    target_file_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src_path, target_file_path)
                    
        except Exception as e:
            print(f"Error copying dependencies: {e}")
    
    def _compile_firmware(self, yaml_path, target_dir, esphome_version=None):
        """Compile firmware and return (path, success, output)"""
        try:
            # Use the appropriate ESPHome command
            if esphome_version and esphome_version != "Default":
                esphome_cmd = self._get_esphome_command(esphome_version)
            else:
                esphome_cmd = "esphome"
            
            compile_command = f'{esphome_cmd} compile "{yaml_path}"'
            
            # Run compilation with output capture
            result = subprocess.run(
                compile_command,
                shell=True,
                capture_output=True,
                text=True,
                cwd=os.path.dirname(yaml_path),
                timeout=300  # 5 minute timeout
            )
            
            output = result.stdout + result.stderr
            success = result.returncode == 0
            
            firmware_path = None
            if success:
                # Find the compiled firmware
                project_name = os.path.splitext(os.path.basename(yaml_path))[0]
                firmware_path = os.path.join(
                    os.path.dirname(yaml_path),
                    ".esphome",
                    "build",
                    project_name,
                    ".pioenvs",
                    project_name,
                    "firmware.bin"
                )
                
                if os.path.exists(firmware_path):
                    # Copy firmware to cache
                    cached_firmware = target_dir / "firmware.bin"
                    shutil.copy2(firmware_path, cached_firmware)
                    firmware_path = cached_firmware
            
            return firmware_path, success, output
            
        except subprocess.TimeoutExpired:
            return None, False, "Compilation timed out after 5 minutes"
        except Exception as e:
            return None, False, f"Compilation error: {str(e)}"

    def compile_pending_firmware(self, upload_id=None):
        """Compile firmware for pending uploads"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if upload_id:
                # Compile specific upload
                cursor.execute('''
                    SELECT * FROM delayed_uploads 
                    WHERE id = ? AND status = 'scheduled' AND compile_mode = 'at_upload'
                ''', (upload_id,))
            else:
                # Compile all pending uploads that need compilation
                cursor.execute('''
                    SELECT * FROM delayed_uploads 
                    WHERE status = 'scheduled' AND compile_mode = 'at_upload'
                    AND (compile_status IS NULL OR compile_status = 'pending' OR compile_status = 'failed')
                ''')
            
            uploads = []
            for row in cursor.fetchall():
                uploads.append({
                    'id': row[0],
                    'yaml_path': row[1],
                    'yaml_filename': row[2],
                    'compiled_firmware_path': row[3],
                    'target_device': row[4],
                    'upload_mode': row[5],
                    'scheduled_time': row[6],
                    'status': row[7],
                    'esphome_version': row[9],
                    'compile_mode': row[12],
                    'compile_status': row[13]
                })
            
            conn.close()
            
            results = []
            for upload in uploads:
                result = self._compile_single_upload(upload)
                results.append(result)
            
            return results

    def _compile_single_upload(self, upload):
        """Compile a single upload"""
        try:
            # Find the job directory
            job_dir = self.compile_cache_dir / f"{os.path.splitext(upload['yaml_filename'])[0]}_*"
            job_dirs = list(self.compile_cache_dir.glob(f"{os.path.splitext(upload['yaml_filename'])[0]}_*"))
            
            if not job_dirs:
                return {'id': upload['id'], 'success': False, 'error': 'Job directory not found'}
            
            job_dir = job_dirs[0]
            
            # Update status to compiling
            self._update_upload_compile_status(upload['id'], 'compiling', None)
            
            # Compile the firmware
            firmware_path, success, output = self._compile_firmware(
                upload['yaml_path'], 
                job_dir, 
                upload['esphome_version']
            )
            
            # Update database
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if success:
                cursor.execute('''
                    UPDATE delayed_uploads 
                    SET compiled_firmware_path = ?, compile_status = ?, 
                        compile_output = ?, last_compile_attempt = ?
                    WHERE id = ?
                ''', (str(firmware_path) if firmware_path else None, 'success', 
                      output, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), upload['id']))
            else:
                cursor.execute('''
                    UPDATE delayed_uploads 
                    SET compile_status = ?, compile_output = ?, last_compile_attempt = ?
                    WHERE id = ?
                ''', ('failed', output, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), upload['id']))
            
            conn.commit()
            conn.close()
            
            return {
                'id': upload['id'],
                'success': success,
                'output': output,
                'firmware_path': str(firmware_path) if firmware_path else None
            }
            
        except Exception as e:
            return {'id': upload['id'], 'success': False, 'error': str(e)}

    def _update_upload_compile_status(self, upload_id, status, output, firmware_path=None):
        """Update compile status in database"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if firmware_path:
                cursor.execute('''
                    UPDATE delayed_uploads 
                    SET compile_status = ?, compile_output = ?, last_compile_attempt = ?, compiled_firmware_path = ?
                    WHERE id = ?
                ''', (status, output, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), firmware_path, upload_id))
            else:
                cursor.execute('''
                    UPDATE delayed_uploads 
                    SET compile_status = ?, compile_output = ?, last_compile_attempt = ?
                    WHERE id = ?
                ''', (status, output, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), upload_id))
            
            conn.commit()
            conn.close()

    def get_upload_compile_status(self, upload_id):
        """Get compile status for an upload"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('SELECT compile_status, compile_output, last_compile_attempt FROM delayed_uploads WHERE id = ?', (upload_id,))
            result = cursor.fetchone()
            conn.close()
            
            if result:
                return {
                    'compile_status': result[0],
                    'compile_output': result[1],
                    'last_compile_attempt': result[2]
                }
            return None

    
    def _get_esphome_command(self, version_name):
        """Get ESPHome command for specific version"""
        # This would need to integrate with your version management system
        # For now, using system default
        return "esphome"
    
    def get_pending_uploads(self):
        """Get all pending uploads"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT * FROM delayed_uploads 
                WHERE status = 'scheduled' AND scheduled_time <= datetime('now')
                ORDER BY scheduled_time
            ''')
            
            uploads = []
            for row in cursor.fetchall():
                uploads.append({
                    'id': row[0],
                    'yaml_path': row[1],
                    'yaml_filename': row[2],
                    'compiled_firmware_path': row[3],
                    'target_device': row[4],
                    'upload_mode': row[5],
                    'scheduled_time': row[6],
                    'status': row[7],
                    'created_at': row[8],
                    'esphome_version': row[9],
                    'device_info': json.loads(row[10]) if row[10] else None,
                    'upload_history': json.loads(row[11]) if row[11] else None
                })
            
            conn.close()
            return uploads
    
    def get_scheduled_uploads(self, include_completed=False):
        """Get all scheduled uploads"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if include_completed:
                cursor.execute('''
                    SELECT * FROM delayed_uploads 
                    ORDER BY scheduled_time
                ''')
            else:
                cursor.execute('''
                    SELECT * FROM delayed_uploads 
                    WHERE status IN ('scheduled', 'processing')
                    ORDER BY scheduled_time
                ''')
            
            uploads = []
            for row in cursor.fetchall():
                uploads.append({
                    'id': row[0],
                    'yaml_path': row[1],
                    'yaml_filename': row[2],
                    'compiled_firmware_path': row[3],
                    'target_device': row[4],
                    'upload_mode': row[5],
                    'scheduled_time': row[6],
                    'status': row[7],
                    'created_at': row[8],
                    'esphome_version': row[9],
                    'device_info': json.loads(row[10]) if row[10] else None,
                    'upload_history': json.loads(row[11]) if row[11] else None,
                    'compile_mode': row[12] if len(row) > 12 else 'at_upload',
                    'compile_status': row[13] if len(row) > 13 else 'pending'
                })
            
            conn.close()
            return uploads
    
    def update_upload_status(self, upload_id, status):
        """Update upload status"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                UPDATE delayed_uploads 
                SET status = ? 
                WHERE id = ?
            ''', (status, upload_id))
            
            conn.commit()
            conn.close()
    
    def delete_upload(self, upload_id):
        """Delete a scheduled upload"""
        with self.lock:
            # First get the upload to clean up files
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('SELECT compiled_firmware_path FROM delayed_uploads WHERE id = ?', (upload_id,))
            result = cursor.fetchone()
            
            if result and result[0]:
                firmware_path = Path(result[0])
                # Delete the cached files
                job_dir = firmware_path.parent
                if job_dir.exists() and job_dir != self.compile_cache_dir:
                    shutil.rmtree(job_dir)
            
            # Delete from database
            cursor.execute('DELETE FROM delayed_uploads WHERE id = ?', (upload_id,))
            conn.commit()
            conn.close()
    
    def create_batch_group(self, name, description=None):
        """Create a batch group for multiple uploads"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO batch_groups (name, description)
                VALUES (?, ?)
            ''', (name, description))
            
            batch_id = cursor.lastrowid
            conn.commit()
            conn.close()
            return batch_id
    
    def add_to_batch(self, batch_id, upload_id):
        """Add upload to batch group"""
        with self.lock:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO batch_uploads (batch_id, upload_id)
                VALUES (?, ?)
            ''', (batch_id, upload_id))
            
            conn.commit()
            conn.close()

class UploadScheduler:
    def __init__(self, delayed_upload_manager, gui_callback=None):
        self.manager = delayed_upload_manager
        self.gui_callback = gui_callback
        self.running = False
        self.scheduler_thread = None
        
    def start(self):
        """Start the scheduler"""
        self.running = True
        self.scheduler_thread = threading.Thread(target=self._run_scheduler, daemon=True)
        self.scheduler_thread.start()
        
    def stop(self):
        """Stop the scheduler"""
        self.running = False
        if self.scheduler_thread:
            self.scheduler_thread.join(timeout=5)
            
    def _run_scheduler(self):
        """Main scheduler loop"""
        while self.running:
            try:
                # Check for pending uploads
                pending_uploads = self.manager.get_pending_uploads()
                
                for upload in pending_uploads:
                    self._process_upload(upload)
                    
                # Sleep for a bit before checking again
                time.sleep(30)  # Check every 30 seconds
                
            except Exception as e:
                print(f"Scheduler error: {e}")
                time.sleep(60)  # Sleep longer on error
    
    def _process_upload(self, upload):
        """Process a single upload"""
        try:
            # Update status to processing
            self.manager.update_upload_status(upload['id'], 'processing')
            
            if self.gui_callback:
                self.gui_callback(f"Processing scheduled upload: {upload['yaml_filename']}")
            
            # NEW: Compile now if needed
            if upload['compile_mode'] == 'at_upload':
                if self.gui_callback:
                    self.gui_callback(f"Compiling firmware for: {upload['yaml_filename']}")
                
                compile_result = self.manager._compile_single_upload(upload)
                
                if not compile_result['success']:
                    self.manager.update_upload_status(upload['id'], 'failed')
                    if self.gui_callback:
                        self.gui_callback(f"Compilation failed for: {upload['yaml_filename']}")
                    return
            
            # Perform the upload
            success = self._perform_upload(upload)
            
            # Update status
            new_status = 'completed' if success else 'failed'
            self.manager.update_upload_status(upload['id'], new_status)
            
            if self.gui_callback:
                status_msg = "completed" if success else "failed"
                self.gui_callback(f"Scheduled upload {status_msg}: {upload['yaml_filename']}")
                
        except Exception as e:
            print(f"Error processing upload {upload['id']}: {e}")
            self.manager.update_upload_status(upload['id'], 'failed')
    
    def _perform_upload(self, upload):
        """Perform the actual upload"""
        try:
            # Use the correct esphome version if specified
            esphome_version = upload.get('esphome_version', None)
            if esphome_version and esphome_version != "System Default":
                esphome_cmd = f'esphome_{esphome_version}'
            else:
                esphome_cmd = 'esphome'
            
            if upload['upload_mode'] == 'OTA':
                # OTA upload
                command = f'{esphome_cmd} upload --device {upload["target_device"]} "{upload["yaml_path"]}"'
            else:
                # COM upload
                command = f'{esphome_cmd} upload --device {upload["target_device"]} "{upload["yaml_path"]}"'
            
            # Use the stored firmware if available
            firmware_path = upload.get('compiled_firmware_path')
            if firmware_path and os.path.exists(firmware_path):
                # For OTA, we might need to use a different approach since esphome upload expects YAML
                # We'll use the standard command for now
                pass
            
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=300  # 5 minute timeout
            )
            
            return result.returncode == 0
            
        except Exception as e:
            print(f"Upload error: {e}")
            return False

###############################
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
            bufsize=0,  # Line buffered
            creationflags=subprocess.HIGH_PRIORITY_CLASS
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
            'firmware_version': r'ESPHome version *?([\d\.]+)',   #'firmware_version': r'ESPHome version ([\d\.]+)',
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
                #time_since_last_line = time.time() - last_line_time
                #print(f"DEBUG: No line available, time since last line: {time_since_last_line:.1f}s")
                
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
        self.data_manager = ESPHomeDataManager()

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
        self.load_settings()
        self.initialize_history_display()

        # FULL SYNC AT STARTUP IN BACKGROUND
        self.startup_full_sync()

    def startup_full_sync(self):
        """Perform full sync of all files at application startup"""
        def startup_sync_thread():
            self.status_var.set("Performing initial file sync...")
            self.log_message(">>> Smart syncing files...", "auto")
            self.log_text.see(tk.END)
            
            # Perform full sync (all files) using configurable paths
            synced_files = sync_esphome_files(
                self.sync_source_path.get(), 
                self.sync_local_path.get(),
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
        # self.notebook.add(self.delayed_upload_tab, text="⏰ Delayed Upload")
        
        # Setup each tab
        self.setup_compiler_tab()
        self.setup_versions_tab()
        self.setup_backup_tab()
        self.setup_tools_tab()
        self.setup_delayed_upload_tab()

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

        self.schedule_btn_main = tb.Button(action_frame, text="Schedule Upload", command=self.open_schedule_dialog, bootstyle="info")
        self.schedule_btn_main.pack(side=LEFT, padx=(0, 5))
        
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
        """Build history section to display last two upload stats"""
        frame = tb.Labelframe(parent, text="📊 Upload History", padding=10, bootstyle="info")  # Changed title
        frame.pack(fill=X, pady=(0, 10))

        # Create a grid layout for the history
        data_grid = tb.Frame(frame)
        data_grid.pack(fill=X)

        # Device Info Section - Left (unchanged)
        info_frame_l = tb.Labelframe(data_grid, text="Device Info", padding=5, bootstyle="primary")
        info_frame_l.grid(row=0, column=0, sticky=NSEW, padx=(0, 5), pady=2)
        data_grid.columnconfigure(0, weight=1)

        # Device Info Section - Right (unchanged)
        info_frame_r = tb.Labelframe(data_grid, text="Device Info", padding=5, bootstyle="primary")
        info_frame_r.grid(row=0, column=1, sticky=NSEW, padx=(5, 0), pady=2)
        data_grid.columnconfigure(1, weight=1)

        # Device info variables - Left frame (unchanged)
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

        # Device info variables - Right frame (unchanged)
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
        
        # Initialize upload history variables - CHANGED TO TWO UPLOAD HISTORIES
        self.build_history = {
            'last_upload': {
                'timestamp': 'Never',
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            },
            'previous_upload': {  # NEW: Second most recent upload
                'timestamp': 'Never',
                'duration': 'N/A', 
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            }
        }

        # Create a grid layout for the history - CHANGED LABELS
        history_grid = tb.Frame(frame)
        history_grid.pack(fill=X)
        
        # Last Upload Section - CHANGED FROM "Last Compile"
        last_upload_frame = tb.Labelframe(history_grid, text="Last Upload", padding=5, bootstyle="primary")
        last_upload_frame.grid(row=0, column=0, sticky=NSEW, padx=(0, 5), pady=2)
        history_grid.columnconfigure(0, weight=1)
        
        # Previous Upload Section - CHANGED FROM "Last Upload"  
        previous_upload_frame = tb.Labelframe(history_grid, text="Previous Upload", padding=5, bootstyle="primary")
        previous_upload_frame.grid(row=0, column=1, sticky=NSEW, padx=(5, 0), pady=2)
        history_grid.columnconfigure(1, weight=1)
        
        # Last upload history labels - CHANGED VARIABLE NAMES
        self.last_upload_time_var = tk.StringVar(value="N/A")
        self.last_upload_duration_var = tk.StringVar(value="N/A")
        self.last_upload_size_var = tk.StringVar(value="N/A")
        self.last_upload_flash_var = tk.StringVar(value="N/A")
        self.last_upload_ram_var = tk.StringVar(value="N/A")
        self.last_upload_version_var = tk.StringVar(value="N/A")
        self.last_upload_file_var = tk.StringVar(value="N/A")
        
        # CHANGED: Single line with bold label and normal value - LAST UPLOAD
        last_upload_time_frame = tb.Frame(last_upload_frame)
        last_upload_time_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_time_frame, text="Time:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_time_frame, textvariable=self.last_upload_time_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        last_upload_duration_frame = tb.Frame(last_upload_frame)
        last_upload_duration_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_duration_frame, text="Duration:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_duration_frame, textvariable=self.last_upload_duration_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        last_upload_size_frame = tb.Frame(last_upload_frame)
        last_upload_size_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_size_frame, text="Size:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_size_frame, textvariable=self.last_upload_size_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        last_upload_flash_frame = tb.Frame(last_upload_frame)
        last_upload_flash_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_flash_frame, text="Flash:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_flash_frame, textvariable=self.last_upload_flash_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        last_upload_ram_frame = tb.Frame(last_upload_frame)
        last_upload_ram_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_ram_frame, text="RAM:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_ram_frame, textvariable=self.last_upload_ram_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        last_upload_version_frame = tb.Frame(last_upload_frame)
        last_upload_version_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_version_frame, text="Version:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_version_frame, textvariable=self.last_upload_version_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        last_upload_file_frame = tb.Frame(last_upload_frame)
        last_upload_file_frame.pack(fill=X, pady=1)
        tb.Label(last_upload_file_frame, text="File:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(last_upload_file_frame, textvariable=self.last_upload_file_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        # Previous upload history labels - NEW VARIABLES
        self.previous_upload_time_var = tk.StringVar(value="N/A")
        self.previous_upload_duration_var = tk.StringVar(value="N/A")
        self.previous_upload_size_var = tk.StringVar(value="N/A")
        self.previous_upload_flash_var = tk.StringVar(value="N/A")
        self.previous_upload_ram_var = tk.StringVar(value="N/A")
        self.previous_upload_version_var = tk.StringVar(value="N/A")
        self.previous_upload_file_var = tk.StringVar(value="N/A")
        
        # CHANGED: Single line with bold label and normal value - PREVIOUS UPLOAD
        previous_upload_time_frame = tb.Frame(previous_upload_frame)
        previous_upload_time_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_time_frame, text="Time:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_time_frame, textvariable=self.previous_upload_time_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        previous_upload_duration_frame = tb.Frame(previous_upload_frame)
        previous_upload_duration_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_duration_frame, text="Duration:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_duration_frame, textvariable=self.previous_upload_duration_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        previous_upload_size_frame = tb.Frame(previous_upload_frame)
        previous_upload_size_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_size_frame, text="Size:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_size_frame, textvariable=self.previous_upload_size_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        previous_upload_flash_frame = tb.Frame(previous_upload_frame)
        previous_upload_flash_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_flash_frame, text="Flash:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_flash_frame, textvariable=self.previous_upload_flash_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        previous_upload_ram_frame = tb.Frame(previous_upload_frame)
        previous_upload_ram_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_ram_frame, text="RAM:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_ram_frame, textvariable=self.previous_upload_ram_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        previous_upload_version_frame = tb.Frame(previous_upload_frame)
        previous_upload_version_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_version_frame, text="Version:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_version_frame, textvariable=self.previous_upload_version_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)
        
        previous_upload_file_frame = tb.Frame(previous_upload_frame)
        previous_upload_file_frame.pack(fill=X, pady=1)
        tb.Label(previous_upload_file_frame, text="File:", font=('Arial', 8, 'bold'), width=10, anchor='w').pack(side=LEFT)
        tb.Label(previous_upload_file_frame, textvariable=self.previous_upload_file_var, font=('Arial', 8)).pack(side=LEFT, fill=X, expand=True)

        # Button frame
        button_frame = tb.Frame(frame)
        button_frame.pack(fill=X, pady=(10, 0))

        # Left side: Refresh button with right-click menu for force refresh
        self.refresh_device_btn = tb.Button(
            button_frame, 
            text="🔄 Refresh Device Info", 
            command=self.start_device_info_check,  # Normal refresh on left-click
            bootstyle="success",
            width=20
        )
        self.refresh_device_btn.pack(side=LEFT, padx=(0, 10))

        # Create right-click menu for force refresh
        refresh_context_menu = tk.Menu(self.refresh_device_btn, tearoff=0)
        refresh_context_menu.add_command(
            label="Force Refresh (No Confirmation)", 
            command=self.force_refresh_device_info
        )

        # Bind right-click to show context menu
        def show_refresh_context_menu(event):
            refresh_context_menu.post(event.x_root, event.y_root)

        self.refresh_device_btn.bind("<Button-3>", show_refresh_context_menu)  # Button-3 is right-click

        # Middle: Clear history button
        clear_history_btn = tb.Button(
            button_frame, 
            text="Clear History", 
            command=self.clear_build_history, 
            bootstyle="secondary", 
            width=15
        )
        clear_history_btn.pack(side=LEFT, padx=(0, 10))

        # Right side: Status label
        self.device_check_status = tb.Label(
            button_frame, 
            text="Ready", 
            bootstyle="info"
        )
        self.device_check_status.pack(side=RIGHT)

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

    def setup_version_section(self, parent):
        """Version selection section with mismatch warning"""
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
        
        # NEW: Version mismatch warning - initially hidden
        self.version_warning_var = tk.StringVar(value="")
        self.version_warning_label = tb.Label(
            frame, 
            textvariable=self.version_warning_var, 
            bootstyle="danger", 
            font=('Arial', 9, 'bold'),
            wraplength=650  # Allow text to wrap
        )
        self.version_warning_label.grid(row=3, column=0, sticky=W, pady=(5, 0))
        
        # Initially hide the warning
        self.hide_version_warning()

    def show_version_warning(self, message):
        """Show a warning message in the version section"""
        self.version_warning_var.set(message)
        self.version_warning_label.configure(bootstyle="danger")
        self.version_warning_label.grid()  # Ensure it's visible

    def hide_version_warning(self):
        """Hide the version warning message"""
        self.version_warning_var.set("")
        self.version_warning_label.grid_remove()  # Remove from layout

    def show_version_success(self, message):
        """Show a success message in the version section"""
        self.version_warning_var.set(message)
        self.version_warning_label.configure(bootstyle="success")
        self.version_warning_label.grid()  # Ensure it's visible

    def check_version_mismatch(self):
        """Check if stored device info firmware matches last upload firmware"""
        if not self.file_path.get():
            self.hide_version_warning()
            self.device_check_status.configure(text="Ready", bootstyle="info")
            return False
        
        yaml_path = self.file_path.get()
        stored_info = self.data_manager.get_device_info(yaml_path)
        upload_history = self.data_manager.get_upload_history(yaml_path)
        
        # If no stored data, hide warning and return
        if not stored_info or not upload_history:
            self.hide_version_warning()
            self.device_check_status.configure(text="Ready", bootstyle="info")
            return False
        
        last_upload = upload_history.get('last_upload', {})
        stored_version = stored_info.get('firmware_version', 'N/A')
        upload_version = last_upload.get('version', 'N/A')
        
        # Check if versions don't match and both are valid
        if (stored_version != upload_version and 
            stored_version not in ['N/A', 'Error', 'Checking...', 'Refreshing...'] and
            upload_version not in ['N/A', 'Error']):
            
            warning_msg = f"⚠️ Version mismatch: Device reports v{stored_version}, but last upload was v{upload_version}"
            self.show_version_warning(warning_msg)
            self.device_check_status.configure(text="Version mismatch", bootstyle="warning")
            return True
        else:
            # Versions match or one is unavailable
            if stored_version not in ['N/A', 'Error', 'Checking...', 'Refreshing...'] and upload_version not in ['N/A', 'Error']:
                success_msg = f"✅ Version match: Device v{stored_version} matches last upload"
                self.show_version_success(success_msg)
                self.device_check_status.configure(text="Versions match", bootstyle="success")
            else:
                self.hide_version_warning()
                self.device_check_status.configure(text="Ready", bootstyle="info")
            return False

    def setup_versions_tab(self):
        """Setup the versions management tab with integrated functionality"""
        main_frame = tb.Frame(self.versions_tab, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Title
        tb.Label(main_frame, text="ESPHome Version Manager", 
                font=('Arial', 16, 'bold'), bootstyle="primary").pack(pady=(0, 20))
        
        # Create a paned window for resizable sections
        paned = tb.PanedWindow(main_frame, orient=HORIZONTAL, bootstyle="primary")
        paned.pack(fill=BOTH, expand=True, pady=10)
        
        # Left pane - Version List
        left_frame = tb.Frame(paned, padding=10)
        paned.add(left_frame, weight=2)
        
        # Right pane - Install New Version
        right_frame = tb.Frame(paned, padding=10)
        paned.add(right_frame, weight=1)
        
        # LEFT PANE: Version List and Management
        list_frame = tb.Labelframe(left_frame, text="Installed Versions", padding=15, bootstyle="primary")
        list_frame.pack(fill=BOTH, expand=True)
        
        # Treeview for versions
        tree_frame = tb.Frame(list_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        columns = ("Name", "Version", "Path", "Status")
        self.versions_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        
        # Define headings
        self.versions_tree.heading("Name", text="Environment Name")
        self.versions_tree.heading("Version", text="ESPHome Version")
        self.versions_tree.heading("Path", text="Path")
        self.versions_tree.heading("Status", text="Status")
        
        # Define columns
        self.versions_tree.column("Name", width=150)
        self.versions_tree.column("Version", width=120)
        self.versions_tree.column("Path", width=200)
        self.versions_tree.column("Status", width=100)
        
        # Add scrollbar
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.versions_tree.yview)
        self.versions_tree.configure(yscroll=tree_scrollbar.set)
        
        self.versions_tree.pack(side=LEFT, fill=BOTH, expand=True)
        tree_scrollbar.pack(side=RIGHT, fill=Y)
        
        # Version management buttons
        button_frame = tb.Frame(list_frame)
        button_frame.pack(fill=X, pady=5)
        
        tb.Button(button_frame, text="Refresh List", 
                command=self.refresh_versions_list, 
                bootstyle="info").pack(side=LEFT, padx=5)
        
        tb.Button(button_frame, text="Remove Selected", 
                command=self.remove_selected_version, 
                bootstyle="danger").pack(side=LEFT, padx=5)
        
        tb.Button(button_frame, text="Set as Active", 
                command=self.set_active_version, 
                bootstyle="success").pack(side=LEFT, padx=5)
        
        # Current version info
        current_frame = tb.Labelframe(left_frame, text="Current Version", padding=10, bootstyle="success")
        current_frame.pack(fill=X, pady=(10, 0))
        
        current_info = tb.Frame(current_frame)
        current_info.pack(fill=X)
        
        tb.Label(current_info, text="Active Version:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=W, padx=5, pady=2)
        self.current_version_label = tb.Label(current_info, textvariable=self.current_esphome_version, font=('Arial', 10), bootstyle="info")
        self.current_version_label.grid(row=0, column=1, sticky=W, padx=5, pady=2)
        
        # RIGHT PANE: Install New Version
        install_frame = tb.Labelframe(right_frame, text="Install New Version", padding=15, bootstyle="primary")
        install_frame.pack(fill=BOTH, expand=True)
        
        # Environment name
        tb.Label(install_frame, text="Environment Name:", bootstyle="primary").pack(anchor=W, pady=(5, 0))
        self.env_name_var = tk.StringVar()
        env_entry = tb.Entry(install_frame, textvariable=self.env_name_var)
        env_entry.pack(fill=X, pady=5)
        
        # Version selection
        tb.Label(install_frame, text="ESPHome Version:", bootstyle="primary").pack(anchor=W, pady=(10, 0))
        self.install_version_var = tk.StringVar(value="latest")
        version_entry = tb.Entry(install_frame, textvariable=self.install_version_var)
        version_entry.pack(fill=X, pady=5)
        
        # Common versions
        tb.Label(install_frame, text="Common Versions:", bootstyle="primary").pack(anchor=W, pady=(10, 0))
        common_frame = tb.Frame(install_frame)
        common_frame.pack(fill=X, pady=5)
        
        common_versions = ["2024.12.9", "2025.6.6", "2025.9.0", "latest"]
        for ver in common_versions:
            tb.Button(common_frame, text=ver, 
                    command=lambda v=ver: self.install_version_var.set(v),
                    bootstyle="outline-primary", width=8).pack(side=LEFT, padx=2)
        
        # Install button
        self.install_btn = tb.Button(install_frame, text="Install Version", 
                                    command=self.install_version, 
                                    bootstyle="success")
        self.install_btn.pack(pady=10)
        
        # Progress and log
        progress_frame = tb.Frame(install_frame)
        progress_frame.pack(fill=X, pady=5)
        
        self.install_progress_var = tk.StringVar(value="Ready")
        tb.Label(progress_frame, textvariable=self.install_progress_var, bootstyle="info").pack(anchor=W)
        
        self.install_log = scrolledtext.ScrolledText(install_frame, height=12, font=('Consolas', 8))
        self.install_log.pack(fill=BOTH, expand=True, pady=5)
        
        # Info section
        info_frame = tb.Labelframe(right_frame, text="Information", padding=10, bootstyle="info")
        info_frame.pack(fill=X, pady=(10, 0))
        
        info_text = """Versions are installed in:
    C:\\esphome_versions\\

    • Each version is isolated in its own virtual environment
    • Switch between versions for compatibility testing
    • Install specific versions for different projects
    • All versions are completely independent"""
        
        tb.Label(info_frame, text=info_text, justify=LEFT, bootstyle="secondary", font=('Arial', 9)).pack(anchor=W)
        
        # Initial population
        self.refresh_versions_list()
        
        # Pre-fill with suggested name
        self.env_name_var.set(f"esphome_latest_{datetime.now().strftime('%Y%m%d')}")

    def refresh_versions_list(self):
        """Refresh the versions treeview"""
        # Clear existing items
        for item in self.versions_tree.get_children():
            self.versions_tree.delete(item)
        
        # Re-scan versions
        self.scan_esphome_versions()
        
        # Add versions to treeview
        for name, info in self.esphome_versions.items():
            status = "System" if name == "Default" else "Virtual Env"
            if name == self.current_esphome_version.get():
                status += " (Active)"
            
            self.versions_tree.insert("", "end", values=(
                name,
                info["version"],
                info["path"],
                status
            ))

    def setup_backup_tab(self):
        """Setup the backup management tab with integrated functionality"""
        main_frame = tb.Frame(self.backup_tab, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Title
        tb.Label(main_frame, text="Backup Management", 
                font=('Arial', 16, 'bold'), bootstyle="primary").pack(pady=(0, 20))
        
        # Configure Treeview font and larger row height for bigger twisties
        self.tree_font = ('Arial', 7)
        self.tree_heading_font = ('Arial', 11, 'bold')
        
        style = ttk.Style()
        style.configure("Treeview", font=self.tree_font, rowheight=30)  # Increased row height
        style.configure("Treeview.Heading", font=self.tree_heading_font)
        
        # Create a paned window for resizable sections
        paned = tb.PanedWindow(main_frame, orient=HORIZONTAL, bootstyle="primary")
        paned.pack(fill=BOTH, expand=True, pady=10)
        
        # Left pane - Backup Settings and Controls
        left_frame = tb.Frame(paned, padding=10)
        paned.add(left_frame, weight=1)
        
        # Right pane - Backup Tree
        right_frame = tb.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)
        
        # LEFT PANE: Backup Settings and Controls
        settings_frame = tb.Labelframe(left_frame, text="Backup Settings", padding=15, bootstyle="primary")
        settings_frame.pack(fill=BOTH, pady=(0, 10))
        
        # Backup settings
        tb.Checkbutton(settings_frame, text="Enable automatic backups", 
                    variable=self.backup_enabled, bootstyle="primary-round-toggle").pack(anchor=W, pady=5)
        
        tb.Label(settings_frame, text="Max backups to keep:").pack(anchor=W, pady=(10, 0))
        max_backup_spin = tb.Spinbox(settings_frame, from_=1, to=50, width=10, 
                                    textvariable=self.max_backups)
        max_backup_spin.pack(anchor=W, pady=5)
        
        tb.Label(settings_frame, text=f"Backup location:", bootstyle="info").pack(anchor=W, pady=(10, 0))
        tb.Label(settings_frame, text=f"{self.backup_base_path}", bootstyle="secondary", 
                font=('Arial', 8)).pack(anchor=W, pady=2)
        
        # Backup status
        status_frame = tb.Labelframe(left_frame, text="Backup Status", padding=15, bootstyle="info")
        status_frame.pack(fill=X, pady=10)
        
        self.backup_status_var = tk.StringVar(value="No backup created yet")
        tb.Label(status_frame, textvariable=self.backup_status_var, bootstyle="info").pack(anchor=W, pady=2)
        
        self.total_size_var = tk.StringVar(value="Total size: Calculating...")
        tb.Label(status_frame, textvariable=self.total_size_var, bootstyle="success", 
                font=('Arial', 10, 'bold')).pack(anchor=W, pady=2)
        
        last_backup_var = tk.StringVar(value="Last backup: Never")
        tb.Label(status_frame, textvariable=last_backup_var, bootstyle="secondary", font=('Arial', 9)).pack(anchor=W, pady=2)
        
        # Update last backup time display
        if hasattr(self, 'last_backup_time') and self.last_backup_time:
            last_backup_var.set(f"Last backup: {self.last_backup_time}")
        
        # Quick Actions
        actions_frame = tb.Labelframe(left_frame, text="Quick Actions", padding=15, bootstyle="success")
        actions_frame.pack(fill=X, pady=10)
        
        tb.Button(actions_frame, text="Backup Current File", 
                command=self.backup_current_file, 
                bootstyle="success", width=20).pack(fill=X, pady=5)
        
        tb.Button(actions_frame, text="Create Full Backup", 
                command=self.create_full_backup, 
                bootstyle="info", width=20).pack(fill=X, pady=5)
        
        tb.Button(actions_frame, text="Cleanup Old Backups", 
                command=lambda: self.cleanup_backups(self.backup_tree), 
                bootstyle="warning", width=20).pack(fill=X, pady=5)
        
        # Information
        info_frame = tb.Labelframe(left_frame, text="Information", padding=15, bootstyle="secondary")
        info_frame.pack(fill=BOTH, expand=True)
        
        info_text = """Backup Strategy:
    • Tiered backup system
    • Keep last 10 most recent
    • Keep first of each day
    • Keep first of each week  
    • Keep first of each month
    • Keep first of each year
    • Always keep the oldest

    Automatic backups occur:
    • Before compilation
    • Before upload"""
        
        tb.Label(info_frame, text=info_text, justify=LEFT, bootstyle="secondary", 
                font=('Arial', 9)).pack(anchor=W)
        
        # RIGHT PANE: Backup Tree
        list_frame = tb.Labelframe(right_frame, text="Backup Files (Tree View)", padding=15, bootstyle="primary")
        list_frame.pack(fill=BOTH, expand=True)
        
        # Button frame for backup tree
        tree_buttons_frame = tb.Frame(list_frame)
        tree_buttons_frame.pack(fill=X, pady=(0, 10))
        
        tb.Button(tree_buttons_frame, text="Refresh Tree", 
                command=lambda: self.populate_backup_tree(self.backup_tree), 
                bootstyle="info").pack(side=LEFT, padx=5)
        
        tb.Button(tree_buttons_frame, text="Expand All", 
                command=lambda: self.expand_all_tree_items(self.backup_tree), 
                bootstyle="outline-info").pack(side=LEFT, padx=5)
        
        tb.Button(tree_buttons_frame, text="Collapse All", 
                command=lambda: self.collapse_all_tree_items(self.backup_tree), 
                bootstyle="outline-info").pack(side=LEFT, padx=5)
        
        tb.Button(tree_buttons_frame, text="Restore Selected", 
                command=lambda: self.restore_backup_tree_item(self.backup_tree), 
                bootstyle="success").pack(side=LEFT, padx=5)
        
        tb.Button(tree_buttons_frame, text="Delete Selected", 
                command=lambda: self.delete_backup_tree_item(self.backup_tree), 
                bootstyle="danger").pack(side=LEFT, padx=5)
        
        # Multi-select info label
        self.multi_select_info = tb.Label(tree_buttons_frame, text="Use Ctrl+Click or Shift+Click for multiple selection", 
                                        bootstyle="secondary", font=('Arial', 8))
        self.multi_select_info.pack(side=RIGHT, padx=5)
        
        # Treeview for backups - CHANGED TO EXTENDED SELECTION
        tree_frame = tb.Frame(list_frame)
        tree_frame.pack(fill=BOTH, expand=True)
        
        columns = ("Name", "Date", "Size", "Type", "FullPath")
        self.backup_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", selectmode="extended")  # Changed to extended
        
        # Define headings with sorting
        self.backup_tree.heading("#0", text="Backup Structure", command=lambda: self.tree_sort(self.backup_tree, "#0", False))
        self.backup_tree.heading("Name", text="File Name", command=lambda: self.tree_sort(self.backup_tree, "Name", False))
        self.backup_tree.heading("Date", text="Backup Date", command=lambda: self.tree_sort(self.backup_tree, "Date", False))
        self.backup_tree.heading("Size", text="Size", command=lambda: self.tree_sort(self.backup_tree, "Size", False))
        self.backup_tree.heading("Type", text="Backup Type", command=lambda: self.tree_sort(self.backup_tree, "Type", False))
        self.backup_tree.heading("FullPath", text="Path")
        
        # Define columns
        self.backup_tree.column("#0", width=200, minwidth=150)
        self.backup_tree.column("Name", width=300, minwidth=100)
        self.backup_tree.column("Date", width=70, minwidth=50)
        self.backup_tree.column("Size", width=10, minwidth=10)
        self.backup_tree.column("Type", width=50, minwidth=20)
        self.backup_tree.column("FullPath", width=0, stretch=False)  # Hidden column
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=self.backup_tree.yview)
        self.backup_tree.configure(yscroll=scrollbar.set)
        
        self.backup_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Double-click to restore
        self.backup_tree.bind("<Double-1>", lambda e: self.restore_backup_tree_item(self.backup_tree))
        
        # Context menu for backup tree
        self.backup_context_menu = tk.Menu(self.backup_tree, tearoff=0)
        self.backup_context_menu.add_command(label="Restore Backup", 
                                            command=lambda: self.restore_backup_tree_item(self.backup_tree))
        self.backup_context_menu.add_command(label="Delete Selected Backups", 
                                        command=lambda: self.delete_backup_tree_item(self.backup_tree))
        self.backup_context_menu.add_separator()
        self.backup_context_menu.add_command(label="Select All", 
                                        command=lambda: self.select_all_tree_items(self.backup_tree))
        self.backup_context_menu.add_command(label="Clear Selection", 
                                        command=lambda: self.backup_tree.selection_remove(self.backup_tree.selection()))
        self.backup_context_menu.add_separator()
        self.backup_context_menu.add_command(label="Open Backup Location", 
                                        command=self.open_backup_location)
        self.backup_context_menu.add_command(label="Open File Location", 
                                        command=lambda: self.open_file_location(self.backup_tree))
        
        self.backup_tree.bind("<Button-3>", self.show_backup_context_menu)
        
        # Initial population
        self.populate_backup_tree(self.backup_tree)

    def populate_backup_tree(self, tree):
        """Populate backup treeview with folder structure"""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Scan backup directory
        if not self.backup_base_path.exists():
            self.update_backup_total_size()
            return
        
        try:
            total_backups = 0
            total_size = 0
            yaml_folders = {}
            
            # First pass: collect all YAML folders and their backups
            for subdir in self.backup_base_path.iterdir():
                if subdir.is_dir():
                    original_filename = subdir.name
                    folder_size = 0
                    folder_backups = []
                    
                    for backup_file in subdir.iterdir():
                        if backup_file.is_file():
                            stat = backup_file.stat()
                            file_time = datetime.fromtimestamp(stat.st_mtime)
                            
                            # Determine backup type from filename and get emoji
                            backup_type = "Recent"
                            type_emoji = "🟢"
                            if backup_file.name.endswith('.daily'):
                                backup_type = "Daily"
                                type_emoji = "🔵"
                            elif backup_file.name.endswith('.weekly'):
                                backup_type = "Weekly" 
                                type_emoji = "🟣"
                            elif backup_file.name.endswith('.monthly'):
                                backup_type = "Monthly"
                                type_emoji = "🟠"
                            elif backup_file.name.endswith('.yearly'):
                                backup_type = "Yearly"
                                type_emoji = "🟡"
                            elif backup_file.name.endswith('.baseline'):
                                backup_type = "Baseline"
                                type_emoji = "⚫"
                            
                            folder_backups.append({
                                'name': backup_file.name,
                                'date': file_time,
                                'size': stat.st_size,
                                'type': backup_type,
                                'type_emoji': type_emoji,
                                'path': backup_file
                            })
                            
                            folder_size += stat.st_size
                            total_backups += 1
                            total_size += stat.st_size
                    
                    yaml_folders[original_filename] = {
                        'size': folder_size,
                        'backups': folder_backups
                    }
            
            # Second pass: build the tree structure with folder icons
            for yaml_file, folder_data in sorted(yaml_folders.items()):
                folder_size_str = self.format_size(folder_data['size'])
                backup_count = len(folder_data['backups'])
                
                # Add folder with icon and count
                folder_text = f"📁 {yaml_file} ({backup_count} backups)"
                folder_item = tree.insert("", "end", text=folder_text, 
                                        values=("", "", folder_size_str, "Folder", ""))
                
                # Add backup files as children with emoji indicators
                for backup in sorted(folder_data['backups'], key=lambda x: x['date'], reverse=True):
                    size_str = self.format_size(backup['size'])
                    date_str = backup['date'].strftime("%Y-%m-%d %H:%M:%S")
                    
                    tree.insert(folder_item, "end", text="", 
                            values=(backup['name'], date_str, size_str, 
                                    f"{backup['type_emoji']} {backup['type']}", str(backup['path'])))
            
            # Update status
            if hasattr(self, 'backup_status_var'):
                size_mb = total_size / (1024 * 1024)
                self.backup_status_var.set(f"Backups: {total_backups} files in {len(yaml_folders)} YAML files")
            
            # Update total size display
            self.update_backup_total_size()
            
            # Auto-expand all folders initially
            self.expand_all_tree_items(tree)
            
        except Exception as e:
            print(f"Error populating backup tree: {e}")
            self.update_backup_total_size()

    def tree_sort(self, tree, column, reverse):
        """Sort tree contents when a column header is clicked"""
        # Get all items from the tree
        items = [(tree.set(item, column), item) for item in tree.get_children('')]
        
        # Separate folders and files for proper sorting
        folders = []
        files = []
        
        for value, item in items:
            if tree.get_children(item):  # This is a folder (has children)
                folders.append((value, item))
            else:
                files.append((value, item))
        
        # Sort based on column type
        if column == "Size":
            # Convert size strings to bytes for proper numeric sorting
            def size_to_bytes(size_str):
                if not size_str or size_str == "":
                    return 0
                size_str = size_str.upper()
                if "KB" in size_str:
                    return float(size_str.replace("KB", "").strip()) * 1024
                elif "MB" in size_str:
                    return float(size_str.replace("MB", "").strip()) * 1024 * 1024
                elif "GB" in size_str:
                    return float(size_str.replace("GB", "").strip()) * 1024 * 1024 * 1024
                else:
                    try:
                        return float(size_str)
                    except ValueError:
                        return 0
            
            folders.sort(key=lambda x: size_to_bytes(x[0]), reverse=reverse)
            files.sort(key=lambda x: size_to_bytes(x[0]), reverse=reverse)
        elif column == "Date":
            # Sort by date
            def parse_date(date_str):
                if not date_str:
                    return datetime.min
                try:
                    return datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    return datetime.min
            
            folders.sort(key=lambda x: parse_date(x[0]), reverse=reverse)
            files.sort(key=lambda x: parse_date(x[0]), reverse=reverse)
        else:
            # Alphabetical sort for other columns
            folders.sort(key=lambda x: x[0].lower() if x[0] else "", reverse=reverse)
            files.sort(key=lambda x: x[0].lower() if x[0] else "", reverse=reverse)
        
        # Rebuild the tree with sorted items
        for item in folders + files:
            tree.move(item[1], '', 'end')
        
        # Reverse sort next time
        tree.heading(column, command=lambda: self.tree_sort(tree, column, not reverse))

    def get_selected_backup_paths(self, tree):
        """Get the full paths of all selected backup files"""
        selections = tree.selection()
        backup_paths = []
        
        for item in selections:
            item_data = tree.item(item)
            values = item_data['values']
            
            # Check if it's a backup file (has FullPath) and not a folder
            if len(values) > 4 and values[4]:  # FullPath exists and is not empty
                backup_path = Path(values[4])
                if backup_path.exists():
                    backup_paths.append(backup_path)
        
        return backup_paths

    def restore_backup_tree_item(self, tree):
        """Restore selected backup from tree (single restore only)"""
        selections = tree.selection()
        if not selections:
            messagebox.showwarning("No Selection", "Please select a backup file to restore")
            return
        
        # For restore, only allow single selection
        if len(selections) > 1:
            messagebox.showwarning("Multiple Selection", "Please select only one backup file to restore")
            return
        
        item = tree.item(selections[0])
        values = item['values']
        
        # Check if it's a backup file (has FullPath) and not a folder
        if len(values) > 4 and values[4]:  # FullPath exists and is not empty
            backup_path = Path(values[4])
        else:
            messagebox.showwarning("Invalid Selection", "Please select a backup file to restore (not a folder)")
            return
        
        if not backup_path.exists():
            messagebox.showerror("Error", "Backup file not found")
            return
        
        # Extract original filename from parent folder name
        original_name = backup_path.parent.name
        
        # Ask for restore location
        restore_path = filedialog.asksaveasfilename(
            title="Restore backup as...",
            initialfile=original_name,
            initialdir=r"C:\esphome",
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
                    
                self.log_message(f">>> Restored backup: {original_name}", "auto")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to restore backup: {str(e)}")

    def delete_backup_tree_item(self, tree):
        """Delete selected backup(s) from tree - NOW SUPPORTS MULTIPLE SELECTION"""
        backup_paths = self.get_selected_backup_paths(tree)
        if not backup_paths:
            messagebox.showwarning("No Selection", "Please select backup file(s) to delete (not folders)")
            return
        
        if len(backup_paths) == 1:
            message = f"Delete backup '{backup_paths[0].name}'?"
        else:
            message = f"Delete {len(backup_paths)} selected backups?"
        
        if not messagebox.askyesno("Confirm Delete", message):
            return
        
        deleted_count = 0
        errors = []
        
        for backup_path in backup_paths:
            try:
                backup_path.unlink()
                deleted_count += 1
            except Exception as e:
                errors.append(f"{backup_path.name}: {str(e)}")
        
        if deleted_count > 0:
            self.populate_backup_tree(tree)
            if errors:
                messagebox.showwarning("Partial Success", 
                                    f"Deleted {deleted_count} backup(s), but {len(errors)} failed:\n" + "\n".join(errors))
            else:
                if len(backup_paths) == 1:
                    messagebox.showinfo("Success", "Backup deleted")
                else:
                    messagebox.showinfo("Success", f"Deleted {deleted_count} backups")
            
            self.log_message(f">>> Deleted {deleted_count} backup(s)", "auto")
        else:
            messagebox.showerror("Error", "Failed to delete any backups:\n" + "\n".join(errors))

    def select_all_tree_items(self, tree):
        """Select all items in the tree (files only, not folders)"""
        all_items = tree.get_children()
        file_items = []
        
        for item in all_items:
            # Check if this is a file (no children) or get all file children of folders
            if not tree.get_children(item):
                file_items.append(item)
            else:
                # Add all children of folders
                file_items.extend(tree.get_children(item))
        
        tree.selection_set(file_items)

    def show_backup_context_menu(self, event):
        """Show context menu for backup tree"""
        item = self.backup_tree.identify_row(event.y)
        if item:
            # Only select the item under cursor if it's not already selected
            if item not in self.backup_tree.selection():
                self.backup_tree.selection_set([item])
            self.backup_context_menu.post(event.x_root, event.y_root)

    def expand_all_tree_items(self, tree):
        """Expand all items in the tree"""
        for item in tree.get_children():
            tree.item(item, open=True)

    def collapse_all_tree_items(self, tree):
        """Collapse all items in the tree"""
        for item in tree.get_children():
            tree.item(item, open=False)

    def open_file_location(self, tree):
        """Open the location of the selected backup file"""
        backup_path = self.get_selected_backup_path(tree)
        if not backup_path:
            messagebox.showwarning("No Selection", "Please select a backup file to open location")
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(backup_path.parent)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", backup_path.parent])
            else:  # Linux
                subprocess.run(["xdg-open", backup_path.parent])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file location: {e}")

    def get_folder_size(self, folder_path):
        """Calculate total size of a folder in bytes"""
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(folder_path):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                    except (OSError, FileNotFoundError):
                        continue
        except (OSError, FileNotFoundError):
            return 0
        return total_size

    def format_size(self, size_bytes):
        """Convert bytes to human readable format"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        if i == 0:
            return f"{size_bytes:.0f} {size_names[i]}"
        else:
            return f"{size_bytes:.2f} {size_names[i]}"

    def update_backup_total_size(self):
        """Update the total backup folder size display"""
        if hasattr(self, 'backup_base_path') and self.backup_base_path.exists():
            total_size = self.get_folder_size(self.backup_base_path)
            formatted_size = self.format_size(total_size)
            self.total_size_var.set(f"Total size: {formatted_size}")
        else:
            self.total_size_var.set("Total size: 0 B")

    def show_backup_context_menu(self, event):
        """Show context menu for backup tree"""
        item = self.backup_tree.identify_row(event.y)
        if item:
            self.backup_tree.selection_set(item)
            self.backup_context_menu.post(event.x_root, event.y_root)

    def open_backup_location(self):
        """Open the backup directory in file explorer"""
        try:
            if sys.platform == "win32":
                os.startfile(self.backup_base_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", self.backup_base_path])
            else:  # Linux
                subprocess.run(["xdg-open", self.backup_base_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open backup location: {e}")

    def create_full_backup(self):
        """Create a backup of all YAML files"""
        def full_backup_thread():
            self.status_var.set("Creating full backup...")
            
            # Find all YAML files in the esphome directory
            esphome_dir = r"C:\esphome"
            yaml_files = []
            
            for root, dirs, files in os.walk(esphome_dir):
                for file in files:
                    if file.lower().endswith(('.yaml', '.yml')):
                        yaml_files.append(os.path.join(root, file))
            
            backed_up_files = []
            for yaml_file in yaml_files:
                backup_path = create_backup(
                    yaml_file,
                    self.backup_base_path,
                    os.path.basename(yaml_file)
                )
                if backup_path:
                    backed_up_files.append(os.path.basename(yaml_file))
            
            if backed_up_files:
                self.last_backup_time = datetime.now().strftime("%H:%M:%S")
                self.backup_status_var.set(f"Full backup: {len(backed_up_files)} files")
                self.status_var.set(f"Full backup created: {len(backed_up_files)} files")
                self.log_message(f">>> Full backup created: {len(backed_up_files)} files", "auto")
                
                # Clean up old backups
                cleanup_old_backups(self.backup_base_path, self.max_backups.get())
                
                # Refresh the backup list
                self.populate_backup_list(self.backup_tree)
                self.update_backup_total_size()
            else:
                self.status_var.set("No files needed backup")
        
        threading.Thread(target=full_backup_thread, daemon=True).start()

    def cleanup_backups(self, tree):
        """Clean up old backups"""
        if messagebox.askyesno("Confirm Cleanup", 
                            f"Remove old backups, keeping only {self.max_backups.get()} most recent per file?"):
            cleanup_old_backups(self.backup_base_path, self.max_backups.get())
            self.populate_backup_tree(tree)
            self.update_backup_total_size()
            messagebox.showinfo("Success", "Old backups cleaned up")
            self.log_message(">>> Cleaned up old backups", "auto")

    def setup_tools_tab(self):
        """Setup the tools tab with data management"""
        main_frame = tb.Frame(self.tools_tab, padding=20)
        main_frame.pack(fill=BOTH, expand=True)
        
        tb.Label(main_frame, text="Tools & Data Management", 
                font=('Arial', 16, 'bold'), bootstyle="primary").pack(pady=(0, 20))
        
        # Create paned window for better layout
        paned = tb.PanedWindow(main_frame, orient=HORIZONTAL, bootstyle="primary")
        paned.pack(fill=BOTH, expand=True, pady=10)
        
        # Left pane - Utilities
        left_frame = tb.Frame(paned, padding=10)
        paned.add(left_frame, weight=1)
        
        # Right pane - Data Management
        right_frame = tb.Frame(paned, padding=10)
        paned.add(right_frame, weight=1)
        
        # LEFT PANE: Utilities
        utils_frame = tb.Labelframe(left_frame, text="Utilities", padding=15, bootstyle="primary")
        utils_frame.pack(fill=BOTH, expand=True)
        
        utilities = [
            ("Setup Context Menu", self.setup_context_menu, "info"),
            ("Check for Updates", self.check_updates, "info"),
            ("Update ESPHome", self.update_esphome, "success"),
            ("Scan COM Ports", self.scan_ports, "secondary"),
            ("Scan OTA Devices", self.scan_ips, "secondary"),
            ("Clean Build Directory", self.clean_build, "warning"),
        ]
        
        for text, command, style in utilities:
            tb.Button(utils_frame, text=text, command=command, bootstyle=style, width=20).pack(pady=5)
        
        # Settings section (sync paths)
        settings_frame = tb.Labelframe(left_frame, text="Sync Settings", padding=15, bootstyle="info")
        settings_frame.pack(fill=X, pady=(10, 0))
        
        # Source path
        tb.Label(settings_frame, text="Network Source Path:", bootstyle="info").pack(anchor=W)
        source_entry = tb.Entry(settings_frame, textvariable=self.sync_source_path, width=35)
        source_entry.pack(fill=X, pady=(2, 8))
        
        # Local path
        tb.Label(settings_frame, text="Local Destination Path:", bootstyle="info").pack(anchor=W)
        local_entry = tb.Entry(settings_frame, textvariable=self.sync_local_path, width=35)
        local_entry.pack(fill=X, pady=(2, 8))
        
        # Save button
        tb.Button(settings_frame, text="Save Settings", 
                command=self.save_settings, bootstyle="success", width=15).pack(pady=5)

        
        # RIGHT PANE: Data Management
        data_frame = tb.Labelframe(right_frame, text="Data Management", padding=15, bootstyle="primary")
        data_frame.pack(fill=BOTH, expand=True)
        
        # Current file info
        current_file_frame = tb.Labelframe(data_frame, text="Current File", padding=10, bootstyle="info")
        current_file_frame.pack(fill=X, pady=(0, 10))
        
        self.current_file_label = tb.Label(current_file_frame, text="No file selected", bootstyle="secondary")
        self.current_file_label.pack(anchor=W, pady=2)
        
        self.storage_status_label = tb.Label(current_file_frame, text="No data stored", bootstyle="secondary")
        self.storage_status_label.pack(anchor=W, pady=2)
        
        # Update current file info
        if self.file_path.get():
            self.update_storage_status()
        
        # Data actions for current file
        current_file_actions = tb.Frame(data_frame)
        current_file_actions.pack(fill=X, pady=5)
        
        tb.Button(current_file_actions, text="Refresh Stored Data", 
                command=self.refresh_stored_data, 
                bootstyle="info", width=18).pack(side=LEFT, padx=2)
        
        tb.Button(current_file_actions, text="Delete Device Info", 
                command=self.delete_current_device_info, 
                bootstyle="warning", width=18).pack(side=LEFT, padx=2)
        
        tb.Button(current_file_actions, text="Delete Upload History", 
                command=self.delete_current_upload_history, 
                bootstyle="warning", width=18).pack(side=LEFT, padx=2)
        
        # Global data management
        global_data_frame = tb.Labelframe(data_frame, text="Global Data Management", padding=10, bootstyle="warning")
        global_data_frame.pack(fill=X, pady=10)
        
        # Statistics
        stats = self.data_manager.get_stats()
        stats_text = f"""Storage Statistics:
    • Devices: {stats['total_devices']}
    • Upload Histories: {stats['total_histories']}
    • Data Location: {stats['data_dir']}"""
        
        tb.Label(global_data_frame, text=stats_text, justify=LEFT, 
                bootstyle="secondary", font=('Arial', 9)).pack(anchor=W, pady=5)
        
        # Global actions
        global_actions = tb.Frame(global_data_frame)
        global_actions.pack(fill=X, pady=5)
        
        tb.Button(global_actions, text="View All Stored Data", 
                command=self.view_all_stored_data, 
                bootstyle="info", width=18).pack(side=LEFT, padx=2)
        
        tb.Button(global_actions, text="Export All Data", 
                command=self.export_all_data, 
                bootstyle="success", width=18).pack(side=LEFT, padx=2)
        
        tb.Button(global_actions, text="Delete All Data", 
                command=self.delete_all_stored_data, 
                bootstyle="danger", width=18).pack(side=LEFT, padx=2)
        
        # Information
        info_frame = tb.Labelframe(data_frame, text="Information", padding=10, bootstyle="secondary")
        info_frame.pack(fill=BOTH, expand=True)
        
        info_text = """Data Storage:
    • Device information is automatically saved
    • Upload history is saved after each upload
    • Data persists between application sessions
    • Fast loading of previously gathered information
    • Manual refresh available when needed"""
        
        tb.Label(info_frame, text=info_text, justify=LEFT, 
                bootstyle="secondary", font=('Arial', 9)).pack(anchor=W)

    def update_storage_status(self):
        """Update the storage status display"""
        if not self.file_path.get():
            self.current_file_label.config(text="No file selected")
            self.storage_status_label.config(text="No data stored")
            return
        
        yaml_path = self.file_path.get()
        has_device_info = self.data_manager.get_device_info(yaml_path) is not None
        has_upload_history = self.data_manager.get_upload_history(yaml_path) is not None
        
        self.current_file_label.config(text=f"File: {os.path.basename(yaml_path)}")
        
        status_parts = []
        if has_device_info:
            status_parts.append("Device Info")
        if has_upload_history:
            status_parts.append("Upload History")
        
        if status_parts:
            self.storage_status_label.config(text=f"Stored: {', '.join(status_parts)}", bootstyle="success")
        else:
            self.storage_status_label.config(text="No data stored", bootstyle="secondary")

    def refresh_stored_data(self):
        """Refresh stored data for current file"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
        
        # Force refresh device info
        self.device_check_status.configure(text="Refreshing...", bootstyle="warning")
        self.get_device_info_for_selected_file()

    def delete_current_device_info(self):
        """Delete device info for current file"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
        
        if messagebox.askyesno("Confirm Delete", "Delete stored device information for this file?"):
            if self.data_manager.delete_device_info(self.file_path.get()):
                messagebox.showinfo("Success", "Device information deleted")
                self.update_storage_status()
                # Clear the display
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
            else:
                messagebox.showerror("Error", "Failed to delete device information")

    def delete_current_upload_history(self):
        """Delete upload history for current file"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
        
        if messagebox.askyesno("Confirm Delete", "Delete stored upload history for this file?"):
            if self.data_manager.delete_upload_history(self.file_path.get()):
                messagebox.showinfo("Success", "Upload history deleted")
                self.update_storage_status()
                # Clear the history display
                self.clear_build_history()
            else:
                messagebox.showerror("Error", "Failed to delete upload history")

    def view_all_stored_data(self):
        """View all stored data in a new window"""
        data_window = tb.Toplevel(self.root)
        data_window.title("All Stored Data")
        data_window.geometry("800x600")
        data_window.transient(self.root)
        
        main_frame = tb.Frame(data_window, padding=10)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Create notebook for different data types
        notebook = tb.Notebook(main_frame)
        notebook.pack(fill=BOTH, expand=True)
        
        # Devices tab
        devices_frame = tb.Frame(notebook, padding=10)
        notebook.add(devices_frame, text="Devices")
        
        devices_data = self.data_manager.load_devices_data()
        if devices_data:
            tree_frame = tb.Frame(devices_frame)
            tree_frame.pack(fill=BOTH, expand=True)
            
            columns = ("YAML File", "Last Updated", "Host Name", "Firmware", "Chip")
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
            
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=120)
            
            for key, data in devices_data.items():
                device_info = data.get('device_info', {})
                tree.insert("", "end", values=(
                    data.get('yaml_file', 'Unknown'),
                    data.get('last_updated', 'Unknown')[:16],  # Shorten timestamp
                    device_info.get('host_name', 'N/A'),
                    device_info.get('firmware_version', 'N/A'),
                    device_info.get('chip', 'N/A')
                ))
            
            scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            tree.pack(side=LEFT, fill=BOTH, expand=True)
            scrollbar.pack(side=RIGHT, fill=Y)
        else:
            tb.Label(devices_frame, text="No device data stored", bootstyle="secondary").pack(pady=20)
        
        # History tab
        history_frame = tb.Frame(notebook, padding=10)
        notebook.add(history_frame, text="Upload History")
        
        history_data = self.data_manager.load_history_data()
        if history_data:
            tree_frame = tb.Frame(history_frame)
            tree_frame.pack(fill=BOTH, expand=True)
            
            columns = ("YAML File", "Last Updated", "Last Upload", "Firmware Size")
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
            
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=120)
            
            for key, data in history_data.items():
                upload_history = data.get('upload_history', {})
                last_upload = upload_history.get('last_upload', {})
                tree.insert("", "end", values=(
                    data.get('yaml_file', 'Unknown'),
                    data.get('last_updated', 'Unknown')[:16],
                    last_upload.get('timestamp', 'Never'),
                    last_upload.get('firmware_size', 'N/A')
                ))
            
            scrollbar = ttk.Scrollbar(tree_frame, orient=VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            tree.pack(side=LEFT, fill=BOTH, expand=True)
            scrollbar.pack(side=RIGHT, fill=Y)
        else:
            tb.Label(history_frame, text="No upload history stored", bootstyle="secondary").pack(pady=20)

    def export_all_data(self):
        """Export all stored data to a zip file"""
        try:
            export_file = filedialog.asksaveasfilename(
                title="Export data as...",
                defaultextension=".zip",
                filetypes=[("ZIP files", "*.zip")]
            )
            
            if export_file:
                import zipfile
                with zipfile.ZipFile(export_file, 'w') as zipf:
                    if self.data_manager.devices_file.exists():
                        zipf.write(self.data_manager.devices_file, "devices.json")
                    if self.data_manager.history_file.exists():
                        zipf.write(self.data_manager.history_file, "upload_history.json")
                
                messagebox.showinfo("Success", f"Data exported to:\n{export_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {e}")

    def delete_all_stored_data(self):
        """Delete all stored data"""
        if messagebox.askyesno("Confirm Delete", 
                            "Delete ALL stored device information and upload history?\nThis action cannot be undone!"):
            if self.data_manager.delete_all_data():
                messagebox.showinfo("Success", "All stored data deleted")
                self.update_storage_status()
                # Clear current displays
                self.clear_build_history()
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
            else:
                messagebox.showerror("Error", "Failed to delete all data")

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

        # Sync path configuration (configurable in Tools tab)
        self.sync_source_path = tk.StringVar(value=r"\\192.168.4.76\config\esphome")
        self.sync_local_path = tk.StringVar(value=r"C:\esphome")


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
        """Start manual device info check - only refreshes if no stored data exists"""
        if not self.file_path.get() or not self.ota_ip_var.get():
            messagebox.showwarning("Warning", "Please select a YAML file and ensure OTA device is selected")
            return
        
        yaml_path = self.file_path.get()
        
        # Check if we already have device info stored
        stored_info = self.data_manager.get_device_info(yaml_path)
        if stored_info:
            # Check if firmware version matches last upload
            upload_history = self.data_manager.get_upload_history(yaml_path)
            if upload_history and upload_history.get('last_upload', {}).get('version') != 'N/A':
                last_upload_version = upload_history['last_upload']['version']
                stored_version = stored_info.get('firmware_version', 'N/A')
                
                if stored_version != last_upload_version and last_upload_version != 'N/A':
                    # Firmware versions don't match - device info is outdated
                    response = messagebox.askyesno(
                        "Device Info Outdated", 
                        f"Stored device info (v{stored_version}) doesn't match last upload (v{last_upload_version}).\n\nRefresh device info now?"
                    )
                    if not response:
                        return
                    # Continue to refresh
                else:
                    # Versions match, no need to refresh
                    messagebox.showinfo("Info", "Device information is already up to date with last upload.")
                    return
            else:
                # No upload history, ask if they want to refresh anyway
                response = messagebox.askyesno(
                    "Refresh Device Info", 
                    "Device information already exists. Refresh anyway?"
                )
                if not response:
                    return
        
        # Disable start button during collection
        self.refresh_device_btn.configure(state="disabled")
        self.device_check_status.configure(text="Checking device...", bootstyle="warning")
        
        # Clear current device info display to show we're refreshing
        self.update_device_info_display({
            'firmware_version': 'Refreshing...',
            'host_name': 'Refreshing...', 
            'wifi_ssid': 'Refreshing...',
            'local_mac': 'Refreshing...',
            'wifi_signal': 'Refreshing...',
            'chip': 'Refreshing...',
            'frequency': 'Refreshing...',
            'framework': 'Refreshing...',
            'psram_size': 'Refreshing...',
            'flash_size': 'Refreshing...'
        })
        
        # Force a fresh device info collection
        self.get_device_info_for_selected_file(force_refresh=True)

    def stop_device_info_check(self):
        """Stop the device info collection"""
        self.device_info_stop_requested = True
        self.refresh_device_btn.configure(state="normal")
        self.device_check_status.configure(text="Stopped", bootstyle="secondary")
        self.status_var.set("Device info collection stopped")
        self.update_progress(0)

    def force_refresh_device_info(self):
        """Force refresh device info without confirmation dialog"""
        if not self.file_path.get() or not self.ota_ip_var.get():
            messagebox.showwarning("Warning", "Please select a YAML file and ensure OTA device is selected")
            return
        
        # Disable start button during collection
        self.refresh_device_btn.configure(state="disabled")
        self.device_check_status.configure(text="Force refreshing...", bootstyle="warning")
        
        # Clear current device info display to show we're refreshing
        self.update_device_info_display({
            'firmware_version': 'Refreshing...',
            'host_name': 'Refreshing...', 
            'wifi_ssid': 'Refreshing...',
            'local_mac': 'Refreshing...',
            'wifi_signal': 'Refreshing...',
            'chip': 'Refreshing...',
            'frequency': 'Refreshing...',
            'framework': 'Refreshing...',
            'psram_size': 'Refreshing...',
            'flash_size': 'Refreshing...'
        })
        
        # Hide any version warnings during refresh
        self.hide_version_warning()
        
        # Log that we're doing a force refresh
        self.log_message(">>> Force refreshing device information...", "auto")
        
        # Force a fresh device info collection
        self.get_device_info_for_selected_file(force_refresh=True)

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
            # Clear any existing version warnings first
            self.hide_version_warning()
            
            self.file_path.set(filename)
            # Add the file to recent files
            self.add_to_recent_files(filename)
            self.status_var.set(f"Selected: {os.path.basename(self.file_path.get())}")

            # Update schedule tab file display
            if hasattr(self, 'schedule_file_var'):
                self.schedule_file_var.set(os.path.basename(filename))

            # Clear display and load stored data for new file
            self.load_stored_data(filename)
        
            # Update storage status in tools tab
            self.update_storage_status()

            self.scan_ips()  # Auto-trigger IP scan
            self.check_sync_status()  # Check sync status for the selected file
            # Ensure we have the latest version list before getting device info
            self.scan_esphome_versions()
            
            # Only auto-check device if we have an IP and no stored data
            stored_info = self.data_manager.get_device_info(filename)
            if not stored_info and self.ota_ip_var.get():
                self.device_check_status.configure(text="Auto-checking device...", bootstyle="warning")
                self.get_device_info_for_selected_file(force_refresh=True)

    def load_stored_data(self, yaml_path):
        """Load stored device info and upload history for a YAML file - clear display if no data"""
        # Clear display first to avoid showing old data
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
        
        # Clear history display
        self.clear_build_history()
        
        # Hide version warnings
        self.hide_version_warning()
        self.device_check_status.configure(text="Ready", bootstyle="info")
        
        # Now try to load stored data
        device_info = self.data_manager.get_device_info(yaml_path)
        if device_info:
            self.log_message(">>> Loaded device information from storage", "auto")
            self.update_device_info_display(device_info)
        else:
            self.log_message(">>> No stored device information found", "auto")
            # Auto-start device info collection if we have an IP
            if self.ota_ip_var.get():
                self.device_check_status.configure(text="Auto-checking device...", bootstyle="warning")
                self.get_device_info_for_selected_file()
        
        # Load upload history
        upload_history = self.data_manager.get_upload_history(yaml_path)
        if upload_history:
            self.log_message(">>> Loaded upload history from storage", "auto")
            self.build_history = upload_history
            self.update_history_display()
        else:
            self.log_message(">>> No stored upload history found", "auto")
            # Initialize with empty history
            self.build_history = {
                'last_upload': {
                    'timestamp': 'Never',
                    'duration': 'N/A',
                    'firmware_size': 'N/A',
                    'flash_usage': 'N/A',
                    'ram_usage': 'N/A',
                    'version': 'N/A',
                    'file': 'N/A'
                },
                'previous_upload': {
                    'timestamp': 'Never',
                    'duration': 'N/A',
                    'firmware_size': 'N/A',
                    'flash_usage': 'N/A',
                    'ram_usage': 'N/A',
                    'version': 'N/A',
                    'file': 'N/A'
                }
            }
            self.update_history_display()

    def initialize_history_display(self):
        """Initialize the history display with stored data"""
        if self.file_path.get():
            # Try to load upload history for current file
            upload_history = self.data_manager.get_upload_history(self.file_path.get())
            if upload_history:
                self.build_history = upload_history
                self.update_history_display()
            else:
                # Initialize with empty history if none exists
                self.build_history = {
                    'last_upload': {
                        'timestamp': 'Never',
                        'duration': 'N/A',
                        'firmware_size': 'N/A',
                        'flash_usage': 'N/A',
                        'ram_usage': 'N/A',
                        'version': 'N/A',
                        'file': 'N/A'
                    },
                    'previous_upload': {
                        'timestamp': 'Never',
                        'duration': 'N/A',
                        'firmware_size': 'N/A',
                        'flash_usage': 'N/A',
                        'ram_usage': 'N/A',
                        'version': 'N/A',
                        'file': 'N/A'
                    }
                }
                self.update_history_display()

    def get_device_info_for_selected_file(self, force_refresh=False):
        """Get device information for the selected YAML file with progress"""
        yaml_path = self.file_path.get()
        ip = self.ota_ip_var.get().strip()

        # If we're not forcing a refresh, try to load from storage first
        if not force_refresh:
            stored_info = self.data_manager.get_device_info(yaml_path)
            stored_history = self.data_manager.get_upload_history(yaml_path)
            
            # Check if firmware versions match between stored device info and last upload
            if stored_info and stored_history:
                device_fw = stored_info.get('firmware_version', 'N/A')
                upload_fw = stored_history.get('last_upload', {}).get('version', 'N/A')
                
                # If versions don't match, mark device info as potentially outdated
                if device_fw != upload_fw and upload_fw != 'N/A' and device_fw != 'N/A':
                    self.device_check_status.configure(
                        text="Device info may be outdated", 
                        bootstyle="warning"
                    )
            
            if stored_info and not force_refresh:
                self.log_message(">>> Loaded device info from storage", "auto")
                self.update_device_info_display(stored_info)
                if stored_history:
                    self.log_message(">>> Loaded upload history from storage", "auto")
                    self.set_build_history(stored_history)
                    self.update_history_display()
                return
        
        # If no stored data or forcing refresh, proceed with collection
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
            self.refresh_device_btn.configure(state="normal")
            self.device_check_status.configure(text="No file/device", bootstyle="secondary")
            return
            
        def progress_callback(progress, status):
            """Update progress bar and status from worker thread"""
            self.root.after(0, lambda: self.update_progress(progress))
            self.root.after(0, lambda: self.status_var.set(status))
            self.root.after(0, lambda: self.update_phase_label(status))
            self.root.after(0, lambda: self.device_check_status.configure(text=status, bootstyle="info"))

        def device_info_thread():
            self.root.after(0, lambda: self.update_progress(0))
            device_info = get_device_info_with_progress(yaml_path, ip, progress_callback)
            
            if device_info and not self.device_info_stop_requested:
                # Save to storage
                self.data_manager.save_device_info(yaml_path, device_info)
                self.root.after(0, lambda: self.update_device_info_display(device_info))
                self.root.after(0, lambda: self.status_var.set("Device information complete"))
                self.root.after(0, lambda: self.update_phase_label("Done"))
                self.root.after(0, lambda: self.device_check_status.configure(text="Complete", bootstyle="success"))
                
                # Auto-select matching ESPHome version
                if device_info['firmware_version'] != 'N/A' and device_info['firmware_version'] != 'Error':
                    self.auto_select_esphome_version(device_info['firmware_version'])
            else:
                if not self.device_info_stop_requested:
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
            
            # Reset buttons when done
            self.root.after(0, lambda: self.refresh_device_btn.configure(state="normal"))

            # Reset progress after a short delay
            self.root.after(2000, lambda: self.update_progress(0))
                        
        threading.Thread(target=device_info_thread, daemon=True).start()

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
            self.refresh_device_btn.configure(state="normal")
            self.device_check_status.configure(text="No file/device", bootstyle="secondary")
            return
            
        def progress_callback(progress, status):
            """Update progress bar and status from worker thread"""
            self.root.after(0, lambda: self.update_progress(progress))
            self.root.after(0, lambda: self.status_var.set(status))
            self.root.after(0, lambda: self.update_phase_label(status))
            self.root.after(0, lambda: self.device_check_status.configure(text=status, bootstyle="info"))

        def device_info_thread():
            self.root.after(0, lambda: self.update_progress(0))
            device_info = get_device_info_with_progress(yaml_path, ip, progress_callback)
            
            if device_info and not self.device_info_stop_requested:
                # Save to storage
                self.data_manager.save_device_info(yaml_path, device_info)
                self.root.after(0, lambda: self.update_device_info_display(device_info))
                self.root.after(0, lambda: self.status_var.set("Device information complete"))
                self.root.after(0, lambda: self.update_phase_label("Done"))
                self.root.after(0, lambda: self.device_check_status.configure(text="Complete", bootstyle="success"))
                
                # Auto-select matching ESPHome version
                if device_info['firmware_version'] != 'N/A' and device_info['firmware_version'] != 'Error':
                    self.auto_select_esphome_version(device_info['firmware_version'])
            else:
                if not self.device_info_stop_requested:
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
            
            # Reset buttons when done
            self.root.after(0, lambda: self.refresh_device_btn.configure(state="normal"))

            # Reset progress after a short delay
            self.root.after(2000, lambda: self.update_progress(0))
                        
        threading.Thread(target=device_info_thread, daemon=True).start()

    def stop_device_info_check(self):
        """Stop the device info collection"""
        self.device_info_stop_requested = True
        self.refresh_device_btn.configure(state="normal")
        self.device_check_status.configure(text="Stopped", bootstyle="secondary")
        self.status_var.set("Device info collection stopped")
        self.update_progress(0)

    def set_build_history(self, history_data):
        """Set the complete build history from stored data"""
        self.build_history = history_data
        
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
        
        # Check for version mismatch after updating device info
        self.check_version_mismatch()

    def update_build_history(self, action_type, duration, firmware_size=None, flash_usage=None, ram_usage=None):
        """Update build history with new upload stats - now tracks two most recent uploads"""
        if action_type != 'upload':
            return  # Only track uploads now
            
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
        
        # Shift the current last_upload to previous_upload
        if self.build_history['last_upload']['timestamp'] != 'Never':
            self.build_history['previous_upload'] = self.build_history['last_upload'].copy()
        
        # Set the new upload as last_upload
        self.build_history['last_upload'] = history_data

        # Save to storage
        if self.file_path.get():
            self.data_manager.save_upload_history(self.file_path.get(), self.build_history)

        # Update the display
        self.update_history_display()
        
        # Check for version mismatch after updating upload history
        self.check_version_mismatch()

    def update_history_display(self,):
        """Update the UI with both upload histories"""
        # Update last upload display
        last_upload = self.build_history['last_upload']
        self.last_upload_time_var.set(f"{last_upload['timestamp']}")
        self.last_upload_duration_var.set(f"{last_upload['duration']}")
        self.last_upload_size_var.set(f"{last_upload['firmware_size']}")
        self.last_upload_flash_var.set(f"{last_upload['flash_usage']}")
        self.last_upload_ram_var.set(f"{last_upload['ram_usage']}")
        self.last_upload_version_var.set(f"{last_upload['version']}")
        self.last_upload_file_var.set(f"{last_upload['file']}")
        
        # Update previous upload display
        previous_upload = self.build_history['previous_upload']
        self.previous_upload_time_var.set(f"{previous_upload['timestamp']}")
        self.previous_upload_duration_var.set(f"{previous_upload['duration']}")
        self.previous_upload_size_var.set(f"{previous_upload['firmware_size']}")
        self.previous_upload_flash_var.set(f"{previous_upload['flash_usage']}")
        self.previous_upload_ram_var.set(f"{previous_upload['ram_usage']}")
        self.previous_upload_version_var.set(f"{previous_upload['version']}")
        self.previous_upload_file_var.set(f"{previous_upload['file']}")

    def clear_build_history(self):
        """Clear all build history"""
        self.build_history = {
            'last_upload': {
                'timestamp': 'Never',
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            },
            'previous_upload': {
                'timestamp': 'Never',
                'duration': 'N/A',
                'firmware_size': 'N/A',
                'flash_usage': 'N/A',
                'ram_usage': 'N/A',
                'version': 'N/A',
                'file': 'N/A'
            }
        }
        self.update_history_display()
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
        self.status_var.set("Upload history cleared")

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
        """Stop the current running process immediately - including device info collection"""
        # Stop any device info collection first
        if hasattr(self, 'device_info_stop_requested'):
            self.device_info_stop_requested = True
            self.refresh_device_btn.configure(state="normal")
            self.device_check_status.configure(text="Stopped", bootstyle="secondary")
        
        # Then stop any compilation/upload processes
        if hasattr(self, 'current_process') and self.current_process:
            try:
                self.is_running = False
                self._force_kill_process()
                self.status_var.set("Process forcefully stopped")
                self.update_phase_label("Stopped")
                self.log_message(">>> Process forcefully terminated", "auto")
                self.log_text.see(tk.END)
                self.update_progress(0)
                self.stop_timer()
                self.error_indicator.configure(bootstyle="warning")

            except Exception as e:
                self.log_message(f">>> Error stopping process: {str(e)}", "auto")
        else:
            self.status_var.set("No process running")

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

    def sync_all_files(self):
        """Enhanced sync that includes all file types"""
        def sync_thread():
            self.status_var.set("Syncing all files...")
            # Use enhanced sync function with configurable paths
            synced_files = sync_esphome_files(
                self.sync_source_path.get(), 
                self.sync_local_path.get(),
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

    def remove_selected_version(self):
        """Remove selected version from the treeview"""
        selection = self.versions_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a version to remove")
            return
        
        item = self.versions_tree.item(selection[0])
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
            self.refresh_versions_list()
            
            # Reset to default if removing current version
            if self.current_esphome_version.get() == env_name:
                self.current_esphome_version.set("Default")
            
            messagebox.showinfo("Success", f"Removed environment '{env_name}'")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove environment: {str(e)}")

    def install_version(self):
        """Install a new ESPHome version"""
        env_name = self.env_name_var.get().strip()
        version = self.install_version_var.get().strip()
        
        if not env_name or not version:
            messagebox.showerror("Error", "Please provide both environment name and version")
            return
        
        # Check if name already exists
        if env_name in self.esphome_versions:
            if not messagebox.askyesno("Overwrite?", f"Environment '{env_name}' already exists. Overwrite?"):
                return
        
        # Disable install button during installation
        self.install_btn.configure(state="disabled")
        self.install_progress_var.set("Starting installation...")
        self.install_log.delete(1.0, tk.END)
        
        def install_thread():
            try:
                self.install_progress_var.set("Creating virtual environment...")
                self.install_log.insert(tk.END, f"Creating environment: {env_name}\n")
                self.install_log.see(tk.END)
                self.root.update_idletasks()
                
                env_path = self.versions_base_path / env_name
                
                # Remove existing if overwriting
                if env_path.exists():
                    shutil.rmtree(env_path)
                
                # Create virtual environment
                venv.create(env_path, with_pip=True)
                
                self.install_progress_var.set("Installing ESPHome...")
                self.install_log.insert(tk.END, f"Installing ESPHome {version}...\n")
                self.install_log.see(tk.END)
                self.root.update_idletasks()
                
                # Install ESPHome
                pip_exe = env_path / "Scripts" / "pip.exe"
                esphome_spec = f"esphome=={version}" if version != "latest" else "esphome"
                
                result = subprocess.run([
                    str(pip_exe), "install", esphome_spec
                ], capture_output=True, text=True, timeout=120)
                
                if result.returncode == 0:
                    self.install_progress_var.set("Installation completed successfully!")
                    self.install_log.insert(tk.END, "Installation completed successfully!\n")
                    self.install_log.insert(tk.END, result.stdout)
                    
                    # Refresh version list and select new version
                    self.scan_esphome_versions()
                    self.current_esphome_version.set(env_name)
                    self.refresh_versions_list()
                    
                    self.install_log.insert(tk.END, f"\n✅ Successfully installed ESPHome {version} as '{env_name}'\n")
                    
                else:
                    self.install_progress_var.set("Installation failed!")
                    self.install_log.insert(tk.END, f"Installation failed!\n{result.stderr}\n")
                    
            except subprocess.TimeoutExpired:
                self.install_progress_var.set("Installation timed out!")
                self.install_log.insert(tk.END, "Installation timed out after 2 minutes.\n")
            except Exception as e:
                self.install_progress_var.set(f"Error: {str(e)}")
                self.install_log.insert(tk.END, f"Error: {str(e)}\n")
            
            finally:
                # Re-enable install button
                self.root.after(0, lambda: self.install_btn.configure(state="normal"))
                self.install_log.see(tk.END)
        
        # Start installation in thread
        threading.Thread(target=install_thread, daemon=True).start()
        
    def set_active_version(self):
        """Set selected version as active"""
        selection = self.versions_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a version to set as active")
            return
        
        item = self.versions_tree.item(selection[0])
        env_name = item['values'][0]
        
        self.current_esphome_version.set(env_name)
        self.refresh_versions_list()
        messagebox.showinfo("Success", f"Set '{env_name}' as active version")

    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.root.update_idletasks()

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
                self.sync_source_path.get(), 
                self.sync_local_path.get(),
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
                self.sync_source_path.get(), 
                self.sync_local_path.get(),
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
            # Clear any existing version warnings first
            self.hide_version_warning()
            
            self.file_path.set(file_path)
            self.status_var.set(f"Selected: {os.path.basename(file_path)}")
            
            # Clear display and load stored data for new file
            self.load_stored_data(file_path)
            
            # Update storage status in tools tab
            self.update_storage_status()
            
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

    def save_settings(self):
        """Save application settings to a configuration file"""
        try:
            config_dir = os.path.expanduser("~/.esphome_studio")
            os.makedirs(config_dir, exist_ok=True)
            config_file = os.path.join(config_dir, "settings.json")
            
            settings = {
                'sync_source_path': self.sync_source_path.get(),
                'sync_local_path': self.sync_local_path.get(),
                'backup_enabled': self.backup_enabled.get(),
                'max_backups': self.max_backups.get(),
            }
            
            with open(config_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            print(f"Could not save settings: {e}")

    def load_settings(self):
        """Load application settings from configuration file"""
        try:
            config_file = os.path.expanduser("~/.esphome_studio/settings.json")
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    settings = json.load(f)
                    
                    if 'sync_source_path' in settings:
                        self.sync_source_path.set(settings['sync_source_path'])
                    if 'sync_local_path' in settings:
                        self.sync_local_path.set(settings['sync_local_path'])
                    if 'backup_enabled' in settings:
                        self.backup_enabled.set(settings['backup_enabled'])
                    if 'max_backups' in settings:
                        self.max_backups.set(settings['max_backups'])
        except Exception as e:
            print(f"Could not load settings: {e}")

    def on_closing(self):
        """Save settings and recent files when application closes"""
        self.save_recent_files()
        self.save_settings()
        self.root.quit()


    def compile(self):
        """Compile firmware only"""
        self.inst_ver_var.set(" ")
        
        # IMMEDIATE VISUAL FEEDBACK
        self.status_var.set("Starting compilation process...")
        self.update_phase_label("Starting...")
        self.update_progress(5)
        self.error_indicator.configure(bootstyle="info")
        self.start_process_spinner("Initializing...")
        self.log_message(">>> Starting compilation process...", "auto")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        if not self.validate_file():
            self.stop_process_spinner()
            return
        
        def compile_thread():
            try:
                # Set running flag
                self.is_running = True
                self.log_message(">>> Compile thread started", "auto")
                
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
                    if synced_files:
                        self.log_message(f">>> Synced files: {len(synced_files)}", "auto")
                
                if not self.is_running:
                    self.log_message(">>> Stopped after sync", "auto")
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
                    self.log_message(">>> Stopped after backup", "auto")
                    self.stop_process_spinner()
                    return
                
                # COMPILE PHASE
                self.status_var.set("Compiling...")
                self.update_phase_label("Compiling...")
                self.update_progress(20)
                self.update_process_status("Compiling firmware...")
                self.log_message(">>> Starting compilation...", "auto")
                self.log_text.see(tk.END)

                # Reset firmware size display
                self.firmware_size_var.set("Firmware size: N/A")

                yaml_path = self.file_path.get()
                start_time = time.time()
                estimated_total = 120
                self.start_timer()
                
                # Compile phase - USE CORRECT ESPHome COMMAND
                esphome_cmd = self.get_esphome_command()
                compile_command = f'{esphome_cmd} compile "{yaml_path}"'
                self.log_message(f">>> Running compile command: {compile_command}", "auto")
                compile_success = self.run_command(compile_command, start_time, estimated_total)

                end_time = time.time()
                duration = end_time - start_time
                self.stop_timer()

                # CHECK is_running BEFORE updating UI
                if not self.is_running:
                    self.log_message(">>> Process was stopped during compile", "auto")
                    self.stop_process_spinner()
                    return

                if compile_success:
                    self.status_var.set("Compilation completed")
                    self.update_phase_label("Done")
                    self.error_indicator.configure(bootstyle="success")
                    self.update_process_status("Complete")
                    self.update_progress(100)
                    
                    # Update build history
                    self.update_build_history(
                        action_type='compile',
                        duration=duration,
                        firmware_size=self.firmware_size,
                        flash_usage=f"{self.firmware_percentage}%" if self.firmware_percentage != "N/A" else "N/A",
                        ram_usage=self.extract_ram_usage_from_log()
                    )
                    
                    self.log_message(">>> Compilation completed successfully", "auto")
                else:
                    self.status_var.set("Compilation failed")
                    self.update_phase_label("Compile failed")
                    self.error_indicator.configure(bootstyle="danger")
                    self.update_process_status("Compile failed")
                    self.update_progress(100)

                # Reset running flag
                self.is_running = False
                self.stop_process_spinner()
                
            except Exception as e:
                self.log_message(f">>> Thread error: {str(e)}", "auto")
                self.log_text.see(tk.END)
                self.is_running = False
                self.stop_process_spinner()
        
        # Start the main thread
        self.compile_thread = threading.Thread(target=compile_thread, daemon=True)
        self.compile_thread.start()

    def analyze_firmware(self):
        """Analyze the last compiled binary to determine build size using ESPHome tools"""
        yaml_path = self.file_path.get()
        if not yaml_path: 
            return

        project_name = os.path.splitext(os.path.basename(yaml_path))[0]
        base_dir = os.path.dirname(yaml_path)
        
        # Paths to build artifacts
        build_dir = os.path.join(base_dir, ".esphome", "build", project_name)
        firmware_elf = os.path.join(build_dir, ".pioenvs", project_name, "firmware.elf")
        firmware_bin = os.path.join(build_dir, ".pioenvs", project_name, "firmware.bin")
        build_log_path = os.path.join(build_dir, "log", "build.log")
        
        if not os.path.exists(firmware_elf):
            self.log_message(">>> Error: No compiled firmware found. Please compile first.", "error")
            return

        self.log_message(f">>> Analyzing firmware size for {project_name}...", "command")
        
        def analyze_thread():
            try:
                # METHOD 1: PRIORITY - Try to get exact numbers from build log
                build_log_available = False
                if os.path.exists(build_log_path):
                    self.log_message(">>> Reading exact size information from build log...", "info")
                    try:
                        with open(build_log_path, 'r', encoding='utf-8', errors='ignore') as f:
                            log_content = f.read()
                        
                        # SPECIFIC PATTERNS FOR ESPHome SIZE OUTPUT
                        flash_pattern = r'Flash:\s*\[[^\]]+\]\s*([\d.]+)%\s*\(used\s*(\d+)\s*bytes\s*from\s*(\d+)\s*bytes\)'
                        ram_pattern = r'RAM:\s*\[[^\]]+\]\s*([\d.]+)%\s*\(used\s*(\d+)\s*bytes\s*from\s*(\d+)\s*bytes\)'
                        
                        flash_matches = re.findall(flash_pattern, log_content)
                        ram_matches = re.findall(ram_pattern, log_content)
                        
                        if flash_matches or ram_matches:
                            build_log_available = True
                            self.log_message(">>> EXACT SIZE INFORMATION FROM COMPILATION:", "success")
                            
                            if flash_matches:
                                # Take the last match (most recent)
                                flash_percent, flash_used, flash_total = flash_matches[-1]
                                flash_used_mb = int(flash_used) / (1024 * 1024)
                                flash_total_mb = int(flash_total) / (1024 * 1024)
                                self.log_message(">>> FLASH USAGE (from compilation):", "success")
                                self.log_message(f">>>   Usage: {flash_percent}%", "info")
                                self.log_message(f">>>   Used: {int(flash_used):,} bytes ({flash_used_mb:.2f} MB)", "info")
                                self.log_message(f">>>   Total: {int(flash_total):,} bytes ({flash_total_mb:.2f} MB)", "info")
                                self.log_message(f">>>   Free: {int(flash_total) - int(flash_used):,} bytes ({(flash_total_mb - flash_used_mb):.2f} MB)", "info")
                            
                            if ram_matches:
                                # Take the last match (most recent)
                                ram_percent, ram_used, ram_total = ram_matches[-1]
                                ram_used_kb = int(ram_used) / 1024
                                ram_total_kb = int(ram_total) / 1024
                                self.log_message(">>> RAM USAGE (from compilation):", "success")
                                self.log_message(f">>>   Usage: {ram_percent}%", "info")
                                self.log_message(f">>>   Used: {int(ram_used):,} bytes ({ram_used_kb:.1f} KB)", "info")
                                self.log_message(f">>>   Total: {int(ram_total):,} bytes ({ram_total_kb:.1f} KB)", "info")
                                self.log_message(f">>>   Free: {int(ram_total) - int(ram_used):,} bytes ({(ram_total_kb - ram_used_kb):.1f} KB)", "info")
                        
                    except Exception as e:
                        self.log_message(f">>> Could not read build log: {str(e)}", "warning")
                
                # METHOD 2: FALLBACK - Use file sizes if build log not available or incomplete
                if not build_log_available:
                    self.log_message(">>> No detailed compilation data found, using file size analysis...", "warning")
                    self.log_message(">>> APPROXIMATE SIZE INFORMATION (from file sizes):", "success")
                    
                    if os.path.exists(firmware_bin):
                        bin_size = os.path.getsize(firmware_bin)
                        bin_kb = bin_size / 1024
                        bin_mb = bin_kb / 1024
                        self.log_message(f">>> Binary (.bin) size: {bin_size:,} bytes ({bin_kb:,.1f} KB / {bin_mb:.2f} MB)", "info")
                        
                        # Common ESP32 flash sizes for comparison
                        flash_sizes = {
                            "4MB": 4 * 1024 * 1024,
                            "8MB": 8 * 1024 * 1024, 
                            "16MB": 16 * 1024 * 1024
                        }
                        
                        for size_name, size_bytes in flash_sizes.items():
                            usage_percent = (bin_size / size_bytes) * 100
                            self.log_message(f">>>   Would use {usage_percent:.1f}% of {size_name} flash", "info")
                    
                    if os.path.exists(firmware_elf):
                        elf_size = os.path.getsize(firmware_elf)
                        elf_kb = elf_size / 1024
                        elf_mb = elf_kb / 1024
                        self.log_message(f">>> ELF (.elf) size: {elf_size:,} bytes ({elf_kb:,.1f} KB / {elf_mb:.2f} MB)", "info")
                        self.log_message(">>> Note: ELF size includes debug symbols, not actual flash usage", "warning")
                    
                    self.log_message(">>> For exact flash/RAM usage, please check compilation log", "warning")
                
                # Always show file sizes for reference, even when we have build log data
                if build_log_available and os.path.exists(firmware_bin):
                    self.log_message(">>> FILE SIZE REFERENCE:", "success")
                    bin_size = os.path.getsize(firmware_bin)
                    bin_kb = bin_size / 1024
                    bin_mb = bin_kb / 1024
                    self.log_message(f">>> Binary file size: {bin_size:,} bytes ({bin_kb:,.1f} KB / {bin_mb:.2f} MB)", "info")
                
                self.log_message(">>> Firmware analysis completed", "success")
                    
            except Exception as e:
                self.log_message(f">>> Error during firmware analysis: {str(e)}", "error")
        
        # Start analysis in a separate thread
        threading.Thread(target=analyze_thread, daemon=True).start()

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
                    self.log_message( f">> SYNC UPDATE: Synced the following: {synced_files}", "auto")
                
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
                    self.update_progress(100)
                    
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
                    self.update_progress(100)

                self.is_running = False
                self.stop_process_spinner()
                
            except Exception as e:
                self.log_message( f">>> Thread error: {str(e)}", "auto")
                self.log_text.see(tk.END)
                self.is_running = False
                self.stop_process_spinner()
        
        self.compile_thread = threading.Thread(target=compile_upload_thread, daemon=True)
        self.compile_thread.start()

    def upload(self):
        """Upload firmware only"""
        self.inst_ver_var.set(" ")
        
        # IMMEDIATE VISUAL FEEDBACK
        self.status_var.set("Starting upload process...")
        self.update_phase_label("Starting...")
        self.update_progress(5)
        self.error_indicator.configure(bootstyle="info")
        self.start_process_spinner("Initializing...")
        self.log_message(">>> Starting upload process...", "auto")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        if not self.validate_file():
            self.stop_process_spinner()
            return
        
        def upload_thread():
            try:
                # Set running flag
                self.is_running = True
                self.log_message(">>> Upload thread started", "auto")
                
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
                    if synced_files:
                        self.log_message(f">>> Synced files: {len(synced_files)}", "auto")
                
                if not self.is_running:
                    self.log_message(">>> Stopped after sync", "auto")
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
                    self.log_message(">>> Stopped after backup", "auto")
                    self.stop_process_spinner()
                    return
                
                # UPLOAD PHASE
                self.status_var.set("Uploading...")
                self.update_phase_label("Uploading...")
                self.update_progress(20)
                self.update_process_status("Uploading to device...")
                self.log_message(">>> Starting upload...", "auto")
                self.log_text.see(tk.END)

                yaml_path = self.file_path.get()
                start_time = time.time()
                estimated_total = 60  # Shorter for upload only
                self.start_timer()
                
                # Build upload command based on mode
                esphome_cmd = self.get_esphome_command()
                mode = self.upload_mode_var.get()
                
                if mode == "COM":
                    port = self.port_var.get().strip()
                    upload_command = f'{esphome_cmd} upload "{yaml_path}"'
                    if port:
                        upload_command = f'{esphome_cmd} upload --device {port} "{yaml_path}"'
                    self.log_message(f">>> COM upload command: {upload_command}", "auto")
                else:  # OTA mode
                    ip = self.ota_ip_var.get().strip()
                    self.log_message(f">>> OTA IP: {ip}", "auto")
                    if not is_ota_device_available(ip):
                        self.status_var.set("OTA device not reachable")
                        self.update_phase_label("OTA check failed")
                        self.error_indicator.configure(bootstyle="warning")
                        self.stop_process_spinner()
                        return
                    upload_command = f'{esphome_cmd} upload --device {ip} "{yaml_path}"'
                    self.log_message(f">>> OTA upload command: {upload_command}", "auto")

                upload_success = self.run_command(upload_command, start_time, estimated_total)
                end_time = time.time()
                duration = end_time - start_time
                self.stop_timer()

                # CHECK is_running BEFORE updating UI
                if not self.is_running:
                    self.log_message(">>> Process was stopped during upload", "auto")
                    self.stop_process_spinner()
                    return

                if upload_success:
                    self.status_var.set("Upload completed successfully")
                    self.update_phase_label("Done")
                    self.error_indicator.configure(bootstyle="success")
                    self.update_process_status("Complete")
                    self.update_progress(100)
                    
                    # Update build history
                    self.update_build_history(
                        action_type='upload',
                        duration=duration,
                        firmware_size=self.firmware_size,
                        flash_usage=f"{self.firmware_percentage}%" if self.firmware_percentage != "N/A" else "N/A",
                        ram_usage=self.extract_ram_usage_from_log()
                    )
                    
                    self.log_message(">>> Upload completed successfully", "auto")
                else:
                    self.status_var.set("Upload failed")
                    self.update_phase_label("Upload failed")
                    self.error_indicator.configure(bootstyle="danger")
                    self.update_process_status("Upload failed")
                    self.update_progress(100)

                # Reset running flag
                self.is_running = False
                self.stop_process_spinner()
                
            except Exception as e:
                self.log_message(f">>> Thread error: {str(e)}", "auto")
                self.log_text.see(tk.END)
                self.is_running = False
                self.stop_process_spinner()
        
        # Start the main thread
        self.upload_thread = threading.Thread(target=upload_thread, daemon=True)
        self.upload_thread.start()


    def setup_delayed_upload_tab(self):
        """Setup the delayed upload tab"""
        self.delayed_upload_tab = tb.Frame(self.notebook, padding=10)
        self.notebook.add(self.delayed_upload_tab, text="⏰ Delayed Upload")
        
        # Initialize managers
        self.delayed_upload_manager = DelayedUploadManager(self.data_manager.data_dir)
        self.upload_scheduler = UploadScheduler(self.delayed_upload_manager, self.delayed_upload_callback)
        self.upload_scheduler.start()
        
        main_frame = tb.Frame(self.delayed_upload_tab)
        main_frame.pack(fill=BOTH, expand=True)
        
        # Create paned window for resizable sections
        paned = tb.PanedWindow(main_frame, orient=HORIZONTAL, bootstyle="primary")
        paned.pack(fill=BOTH, expand=True, pady=10)
        
        # Left pane - Schedule new upload
        left_frame = tb.Frame(paned, padding=10)
        paned.add(left_frame, weight=1)
        
        # Right pane - Scheduled uploads list
        right_frame = tb.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)
        
        self.setup_schedule_upload_section(left_frame)
        self.setup_scheduled_uploads_section(right_frame)
        
        # Load initial data
        self.refresh_scheduled_uploads_list()

    def setup_schedule_upload_section(self, parent):
        """Setup the schedule upload section"""
        frame = tb.Labelframe(parent, text="Schedule New Upload", padding=15, bootstyle="primary")
        frame.pack(fill=BOTH, expand=True)
        
        # Current file info
        current_file_frame = tb.Frame(frame)
        current_file_frame.pack(fill=X, pady=5)
        
        tb.Label(current_file_frame, text="Current File:", bootstyle="primary", font=('Arial', 10, 'bold')).pack(anchor=W)
        self.schedule_file_var = tk.StringVar(value="No file selected")
        tb.Label(current_file_frame, textvariable=self.schedule_file_var, bootstyle="secondary").pack(anchor=W, pady=2)
        
        # Schedule time
        schedule_time_frame = tb.Frame(frame)
        schedule_time_frame.pack(fill=X, pady=10)
        
        tb.Label(schedule_time_frame, text="Schedule Time:", bootstyle="primary").grid(row=0, column=0, sticky=W, pady=5)
        
        # Date and time entry
        datetime_frame = tb.Frame(schedule_time_frame)
        datetime_frame.grid(row=1, column=0, columnspan=2, sticky=EW, pady=5)
        
        # Date
        self.schedule_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        tb.Label(datetime_frame, text="Date:").pack(side=LEFT, padx=(0, 5))
        self.schedule_date_entry = tb.Entry(datetime_frame, textvariable=self.schedule_date_var, width=12)
        self.schedule_date_entry.pack(side=LEFT, padx=(0, 10))
        
        # Time
        self.schedule_time_var = tk.StringVar(value=(datetime.now() + timedelta(hours=1)).strftime("%H:%M"))
        tb.Label(datetime_frame, text="Time:").pack(side=LEFT, padx=(0, 5))
        self.schedule_time_entry = tb.Entry(datetime_frame, textvariable=self.schedule_time_var, width=8)
        self.schedule_time_entry.pack(side=LEFT)
        
        # Quick schedule buttons
        quick_schedule_frame = tb.Frame(schedule_time_frame)
        quick_schedule_frame.grid(row=2, column=0, columnspan=2, sticky=EW, pady=5)
        
        quick_times = [
            ("+1 hour", timedelta(hours=1)),
            ("+2 hours", timedelta(hours=2)),
            ("+6 hours", timedelta(hours=6)),
            ("Tomorrow", timedelta(days=1)),
        ]
        
        for text, delta in quick_times:
            tb.Button(quick_schedule_frame, text=text, bootstyle="outline-primary", width=10,
                    command=lambda d=delta: self.set_quick_schedule(d)).pack(side=LEFT, padx=2)
        
        # Target device
        device_frame = tb.Frame(frame)
        device_frame.pack(fill=X, pady=10)
        
        tb.Label(device_frame, text="Target Device:", bootstyle="primary").grid(row=0, column=0, sticky=W, pady=5)
        
        self.schedule_device_var = tk.StringVar()
        device_combo = tb.Combobox(device_frame, textvariable=self.schedule_device_var, state="readonly")
        device_combo.grid(row=1, column=0, sticky=EW, pady=5)
        
        # Update device list when mode changes
        def update_device_list(*args):
            if self.upload_mode_var.get() == "OTA":
                devices = [f"{name} ({ip})" for name, ip in discover_esphome_devices()]
                device_combo['values'] = devices
                if devices:
                    device_combo.set(devices[0])
            else:
                ports = [port.device for port in serial.tools.list_ports.comports()]
                device_combo['values'] = ports
                if ports:
                    device_combo.set(ports[0])
        
        self.upload_mode_var.trace('w', update_device_list)
        update_device_list()  # Initial update
        
        # Schedule button
        button_frame = tb.Frame(frame)
        button_frame.pack(fill=X, pady=10)
        
        self.schedule_btn = tb.Button(button_frame, text="Schedule Upload", 
                                    command=self.schedule_upload, bootstyle="success", width=15)
        self.schedule_btn.pack(pady=5)
        
        # Batch scheduling
        batch_frame = tb.Labelframe(frame, text="Batch Scheduling", padding=10, bootstyle="info")
        batch_frame.pack(fill=X, pady=10)
        
        tb.Label(batch_frame, text="Batch Name:", bootstyle="info").pack(anchor=W, pady=2)
        self.batch_name_var = tk.StringVar(value=f"Batch_{datetime.now().strftime('%Y%m%d_%H%M')}")
        batch_name_entry = tb.Entry(batch_frame, textvariable=self.batch_name_var)
        batch_name_entry.pack(fill=X, pady=5)
        
        tb.Button(batch_frame, text="Add Multiple Files to Batch", 
                command=self.open_batch_scheduler, bootstyle="info").pack(pady=5)

        # NEW: Compile mode selection
        compile_mode_frame = tb.Frame(frame)
        compile_mode_frame.pack(fill=X, pady=10)
        
        tb.Label(compile_mode_frame, text="Compile Mode:", bootstyle="primary").grid(row=0, column=0, sticky=W, pady=5)
        
        self.compile_mode_var = tk.StringVar(value="at_upload")
        
        compile_mode_container = tb.Frame(compile_mode_frame)
        compile_mode_container.grid(row=1, column=0, columnspan=2, sticky=EW, pady=5)
        
        # Compile mode options with explanations
        modes_frame = tb.Frame(compile_mode_container)
        modes_frame.pack(fill=X)
        
        # Option 1: Compile at upload time
        mode1_frame = tb.Frame(modes_frame)
        mode1_frame.pack(fill=X, pady=2)
        
        tb.Radiobutton(mode1_frame, text="Compile at upload time", 
                    variable=self.compile_mode_var, value="at_upload",
                    bootstyle="primary").pack(side=LEFT)
        
        help1 = tb.Label(mode1_frame, text="ⓘ", bootstyle="info", cursor="hand2")
        help1.pack(side=LEFT, padx=5)
        help1.bind("<Button-1>", lambda e: messagebox.showinfo("Compile at Upload", 
            "Firmware will be compiled when the upload runs.\n\n"
            "✓ Uses latest source code\n"
            "✗ May fail if compilation environment changes\n"
            "✗ You won't know about compile errors until upload time"))
        
        # Option 2: Compile at schedule time  
        mode2_frame = tb.Frame(modes_frame)
        mode2_frame.pack(fill=X, pady=2)
        
        tb.Radiobutton(mode2_frame, text="Compile immediately", 
                    variable=self.compile_mode_var, value="at_schedule",
                    bootstyle="primary").pack(side=LEFT)
        
        help2 = tb.Label(mode2_frame, text="ⓘ", bootstyle="info", cursor="hand2")
        help2.pack(side=LEFT, padx=5)
        help2.bind("<Button-1>", lambda e: messagebox.showinfo("Compile Immediately", 
            "Firmware will be compiled now and stored for later upload.\n\n"
            "✓ Catch compile errors while you're awake\n"  
            "✓ Guaranteed working binary for upload\n"
            "✗ Uses current source code (may be outdated at upload time)"))

    def setup_scheduled_uploads_section(self, parent):
        """Setup the scheduled uploads list section"""
        frame = tb.Labelframe(parent, text="Scheduled Uploads", padding=15, bootstyle="primary")
        frame.pack(fill=BOTH, expand=True)
        
        # Toolbar - ADD COMPILE BUTTONS
        toolbar = tb.Frame(frame)
        toolbar.pack(fill=X, pady=(0, 10))
        
        tb.Button(toolbar, text="Refresh", command=self.refresh_scheduled_uploads_list, 
                bootstyle="outline-primary").pack(side=LEFT, padx=2)
        tb.Button(toolbar, text="Batch Schedule", command=self.open_batch_scheduler, 
                bootstyle="info").pack(side=LEFT, padx=2)  # NEW - Batch scheduler button
        tb.Button(toolbar, text="Compile Now", command=self.compile_selected_uploads, 
                bootstyle="outline-warning").pack(side=LEFT, padx=2)  # NEW
        tb.Button(toolbar, text="Run Now", command=self.run_selected_upload_now, 
                bootstyle="outline-success").pack(side=LEFT, padx=2)
        tb.Button(toolbar, text="Delete", command=self.delete_selected_upload, 
                bootstyle="outline-danger").pack(side=LEFT, padx=2)
        tb.Button(toolbar, text="Process All Due", command=self.process_all_due_uploads, 
                bootstyle="success").pack(side=LEFT, padx=2)
        
        # Uploads list - ADD COMPILE STATUS COLUMN
        list_frame = tb.Frame(frame)
        list_frame.pack(fill=BOTH, expand=True)
        
        columns = ("ID", "File", "Device", "Scheduled", "Compile Mode", "Compile Status", "Status")  # UPDATED
        self.uploads_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        
        # Define headings
        for col in columns:
            self.uploads_tree.heading(col, text=col)
        
        # Define columns
        self.uploads_tree.column("ID", width=50)
        self.uploads_tree.column("File", width=150)
        self.uploads_tree.column("Device", width=120)
        self.uploads_tree.column("Scheduled", width=120)
        self.uploads_tree.column("Compile Mode", width=100)  # NEW
        self.uploads_tree.column("Compile Status", width=100)  # NEW
        self.uploads_tree.column("Status", width=80)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.uploads_tree.yview)
        self.uploads_tree.configure(yscroll=scrollbar.set)
        
        self.uploads_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # Context menu
        self.uploads_context_menu = tk.Menu(self.uploads_tree, tearoff=0)
        self.uploads_context_menu.add_command(label="Compile Now", command=self.compile_selected_uploads)
        self.uploads_context_menu.add_command(label="Run Now", command=self.run_selected_upload_now)
        self.uploads_context_menu.add_command(label="Edit Schedule", command=self.edit_upload_schedule)
        self.uploads_context_menu.add_separator()
        self.uploads_context_menu.add_command(label="Delete", command=self.delete_selected_upload)
        
        self.uploads_tree.bind("<Button-3>", self.show_uploads_context_menu)

    def delayed_upload_callback(self, message):
        """Callback for scheduler status updates"""
        self.log_message(f">>> {message}", "auto")

    def set_quick_schedule(self, time_delta):
        """Set a quick schedule time"""
        new_time = datetime.now() + time_delta
        self.schedule_date_var.set(new_time.strftime("%Y-%m-%d"))
        self.schedule_time_var.set(new_time.strftime("%H:%M"))

    def schedule_upload(self):
        """Schedule a new upload"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
        
        # Validate schedule time
        try:
            schedule_datetime = datetime.strptime(
                f"{self.schedule_date_var.get()} {self.schedule_time_var.get()}", 
                "%Y-%m-%d %H:%M"
            )
            
            if schedule_datetime <= datetime.now():
                messagebox.showwarning("Invalid Time", "Scheduled time must be in the future")
                return
                
        except ValueError:
            messagebox.showwarning("Invalid Format", "Please use YYYY-MM-DD for date and HH:MM for time")
            return
        
        if not self.schedule_device_var.get():
            messagebox.showwarning("No Device", "Please select a target device")
            return
        
        # Get device info and upload history for storage
        device_info = self.data_manager.get_device_info(self.file_path.get())
        upload_history = self.data_manager.get_upload_history(self.file_path.get())
        
        # NEW: Show compilation progress if compiling immediately
        if self.compile_mode_var.get() == 'at_schedule':
            self.status_var.set("Compiling firmware for scheduled upload...")
            self.log_message(f">>> Compiling firmware for scheduled upload: {os.path.basename(self.file_path.get())}", "auto")
        
        # Store the upload job
        upload_id = self.delayed_upload_manager.store_upload_job(
            yaml_path=self.file_path.get(),
            target_device=self.schedule_device_var.get(),
            upload_mode=self.upload_mode_var.get(),
            scheduled_time=schedule_datetime.strftime("%Y-%m-%d %H:%M:%S"),
            esphome_version=self.current_esphome_version.get(),
            device_info=device_info,
            upload_history=upload_history,
            compile_mode=self.compile_mode_var.get()  # NEW: Pass compile mode
        )
        
        if upload_id:
            if self.compile_mode_var.get() == 'at_schedule':
                # Check if compilation was successful
                status = self.delayed_upload_manager.get_upload_compile_status(upload_id)
                if status and status['compile_status'] == 'success':
                    messagebox.showinfo("Success", 
                        f"Upload scheduled for {schedule_datetime.strftime('%Y-%m-%d %H:%M')}\n"
                        f"Firmware compiled successfully!")
                    self.log_message(">>> Firmware compiled successfully for scheduled upload", "auto")
                else:
                    messagebox.showerror("Compilation Failed", 
                        f"Failed to compile firmware for scheduled upload!\n\n"
                        f"Check the scheduled uploads list for details.")
                    self.log_message(">>> Firmware compilation failed for scheduled upload", "error")
            else:
                messagebox.showinfo("Success", 
                    f"Upload scheduled for {schedule_datetime.strftime('%Y-%m-%d %H:%M')}\n"
                    f"Firmware will be compiled at upload time.")
            
            self.log_message(f">>> Upload scheduled: {os.path.basename(self.file_path.get())} for {schedule_datetime.strftime('%Y-%m-%d %H:%M')}", "auto")
            self.refresh_scheduled_uploads_list()
        else:
            messagebox.showerror("Error", "Failed to schedule upload")

    def refresh_scheduled_uploads_list(self):
        """Refresh the scheduled uploads list"""
        # Clear existing items
        for item in self.uploads_tree.get_children():
            self.uploads_tree.delete(item)
        
        # Load scheduled uploads
        uploads = self.delayed_upload_manager.get_scheduled_uploads()
        
        for upload in uploads:
            scheduled_time = datetime.strptime(upload['scheduled_time'], "%Y-%m-%d %H:%M:%S")
            
            # Determine compile status display
            compile_status = upload.get('compile_status', 'N/A')
            if compile_status == 'success':
                compile_display = "✅ Success"
            elif compile_status == 'failed':
                compile_display = "❌ Failed"
            elif compile_status == 'compiling':
                compile_display = "🔄 Compiling"
            else:
                compile_display = "⏳ Pending"
            
            # Determine compile mode display
            compile_mode = upload.get('compile_mode', 'at_upload')
            if compile_mode == 'at_schedule':
                mode_display = "Immediate"
            else:
                mode_display = "At Upload"
            
            self.uploads_tree.insert("", "end", values=(
                upload['id'],
                upload['yaml_filename'],
                upload['target_device'],
                scheduled_time.strftime("%m/%d %H:%M"),
                mode_display,  # NEW
                compile_display,  # NEW
                upload['status']
            ))

    def run_selected_upload_now(self):
        """Run selected upload immediately"""
        selection = self.uploads_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an upload to run")
            return
        
        for item in selection:
            upload_id = self.uploads_tree.item(item)['values'][0]
            # Update schedule time to now
            self.delayed_upload_manager.update_upload_status(upload_id, 'scheduled')
            # The scheduler will pick it up on next check
        
        messagebox.showinfo("Success", "Selected upload(s) will run shortly")
        self.refresh_scheduled_uploads_list()

    def delete_selected_upload(self):
        """Delete selected upload(s)"""
        selection = self.uploads_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select upload(s) to delete")
            return
        
        if messagebox.askyesno("Confirm Delete", f"Delete {len(selection)} scheduled upload(s)?"):
            for item in selection:
                upload_id = self.uploads_tree.item(item)['values'][0]
                self.delayed_upload_manager.delete_upload(upload_id)
            
            self.refresh_scheduled_uploads_list()
            self.log_message(f">>> Deleted {len(selection)} scheduled upload(s)", "auto")

    def process_all_due_uploads(self):
        """Process all due uploads immediately"""
        pending = self.delayed_upload_manager.get_pending_uploads()
        if not pending:
            messagebox.showinfo("No Pending", "No pending uploads to process")
            return
        
        if messagebox.askyesno("Confirm", f"Process {len(pending)} pending upload(s) now?"):
            # The scheduler will pick them up on next check
            messagebox.showinfo("Started", f"Processing {len(pending)} upload(s)...")
            self.log_message(f">>> Processing {len(pending)} pending upload(s)", "auto")

    def show_uploads_context_menu(self, event):
        """Show context menu for uploads tree"""
        item = self.uploads_tree.identify_row(event.y)
        if item:
            self.uploads_tree.selection_set(item)
            self.uploads_context_menu.post(event.x_root, event.y_root)

    def edit_upload_schedule(self):
        """Edit schedule for selected upload"""
        selection = self.uploads_tree.selection()
        if not selection or len(selection) > 1:
            messagebox.showwarning("Selection", "Please select exactly one upload to edit")
            return
        
        # Get the selected upload info
        item = selection[0]
        values = self.uploads_tree.item(item)['values']
        upload_id = values[0]
        current_file = values[1]
        current_time = values[3]
        compile_status = values[5] if len(values) > 5 else "N/A"
        
        # Create edit dialog
        dialog = tb.Toplevel(self.root)
        dialog.title(f"Edit Schedule - {current_file}")
        dialog.geometry("400x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Current schedule info
        info_frame = tb.Labelframe(dialog, text="Current Schedule", padding=10, bootstyle="info")
        info_frame.pack(fill=X, padx=15, pady=10)
        
        tb.Label(info_frame, text=f"File: {current_file}", bootstyle="info").pack(anchor=W)
        tb.Label(info_frame, text=f"Current Time: {current_time}", bootstyle="info").pack(anchor=W)
        tb.Label(info_frame, text=f"Compile Status: {compile_status}", bootstyle="info").pack(anchor=W)
        
        # New time selection
        time_frame = tb.Labelframe(dialog, text="New Scheduled Time", padding=10, bootstyle="primary")
        time_frame.pack(fill=X, padx=15, pady=10)
        
        # Date entry
        date_inner = tb.Frame(time_frame)
        date_inner.pack(fill=X, pady=5)
        tb.Label(date_inner, text="Date (YYYY-MM-DD):", width=16).pack(side=LEFT)
        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        tb.Entry(date_inner, textvariable=date_var, width=15).pack(side=LEFT, padx=5)
        
        # Time entry
        time_inner = tb.Frame(time_frame)
        time_inner.pack(fill=X, pady=5)
        tb.Label(time_inner, text="Time (HH:MM):", width=16).pack(side=LEFT)
        time_var = tk.StringVar(value="02:00")
        tb.Entry(time_inner, textvariable=time_var, width=15).pack(side=LEFT, padx=5)
        
        # Quick time buttons
        quick_frame = tb.Frame(time_frame)
        quick_frame.pack(fill=X, pady=5)
        tb.Label(quick_frame, text="Quick:", width=16).pack(side=LEFT)
        tb.Button(quick_frame, text="Tonight 2AM", bootstyle="outline-info",
                command=lambda: [date_var.set((datetime.now() + timedelta(days=0 if datetime.now().hour < 2 else 1)).strftime("%Y-%m-%d")), time_var.set("02:00")]).pack(side=LEFT, padx=2)
        tb.Button(quick_frame, text="+1 Hour", bootstyle="outline-info",
                command=lambda: [date_var.set((datetime.now() + timedelta(hours=1)).strftime("%Y-%m-%d")), time_var.set((datetime.now() + timedelta(hours=1)).strftime("%H:%M"))]).pack(side=LEFT, padx=2)
        
        # View compile output button (if failed)
        if compile_status == 'failed':
            output_frame = tb.Frame(dialog)
            output_frame.pack(fill=X, padx=15, pady=5)
            tb.Button(output_frame, text="View Compile Output", bootstyle="warning",
                    command=lambda: self.view_compile_output(upload_id)).pack()
        
        # Action buttons
        btn_frame = tb.Frame(dialog)
        btn_frame.pack(fill=X, padx=15, pady=15)
        
        def save_new_schedule():
            try:
                new_datetime = datetime.strptime(f"{date_var.get()} {time_var.get()}", "%Y-%m-%d %H:%M")
                new_time_str = new_datetime.strftime("%Y-%m-%d %H:%M:%S")
                
                # Update in database
                import sqlite3
                conn = sqlite3.connect(self.delayed_upload_manager.db_file)
                cursor = conn.cursor()
                cursor.execute('UPDATE delayed_uploads SET scheduled_time = ? WHERE id = ?', (new_time_str, upload_id))
                conn.commit()
                conn.close()
                
                self.refresh_scheduled_uploads_list()
                self.log_message(f">>> Updated schedule for {current_file} to {new_time_str}", "auto")
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Invalid Date/Time", f"Please enter valid date (YYYY-MM-DD) and time (HH:MM)")
        
        tb.Button(btn_frame, text="Save", command=save_new_schedule, bootstyle="success", width=12).pack(side=LEFT, padx=5)
        tb.Button(btn_frame, text="Cancel", command=dialog.destroy, bootstyle="secondary", width=12).pack(side=LEFT, padx=5)
    
    def view_compile_output(self, upload_id):
        """View the compile output for a failed upload"""
        status = self.delayed_upload_manager.get_upload_compile_status(upload_id)
        if not status:
            messagebox.showinfo("No Output", "No compile output available")
            return
        
        # Create output window
        output_win = tb.Toplevel(self.root)
        output_win.title("Compile Output")
        output_win.geometry("800x500")
        
        tb.Label(output_win, text=f"Compile Status: {status.get('compile_status', 'N/A')}", 
                font=('Arial', 12, 'bold'), bootstyle="warning").pack(pady=10)
        
        text = scrolledtext.ScrolledText(output_win, wrap=tk.WORD, font=('Consolas', 9))
        text.pack(fill=BOTH, expand=True, padx=10, pady=5)
        text.insert(tk.END, status.get('compile_output', 'No output available'))
        text.config(state='disabled')
        
        tb.Button(output_win, text="Close", command=output_win.destroy, bootstyle="secondary").pack(pady=10)

    def open_batch_scheduler(self):
        """Open batch scheduler dialog for scheduling multiple files at once"""
        # Create batch scheduler dialog
        dialog = tb.Toplevel(self.root)
        dialog.title("Batch Scheduler - Schedule Multiple Uploads")
        dialog.geometry("850x750")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Header
        tb.Label(dialog, text="Schedule Multiple Files", font=('Arial', 14, 'bold'),
                bootstyle="primary").pack(pady=10)
        
        # File selection frame
        file_frame = tb.Labelframe(dialog, text="Select YAML Files", padding=10, bootstyle="primary")
        file_frame.pack(fill=BOTH, expand=True, padx=15, pady=5)
        
        # Get list of YAML files from local path
        yaml_files = []
        local_path = self.sync_local_path.get()
        if os.path.exists(local_path):
            for f in os.listdir(local_path):
                if f.endswith('.yaml') or f.endswith('.yml'):
                    yaml_files.append(f)
            yaml_files.sort()
        
        # Listbox with multi-select
        list_frame = tb.Frame(file_frame)
        list_frame.pack(fill=BOTH, expand=True)
        
        file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, font=('Consolas', 10), height=12)
        file_scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL, command=file_listbox.yview)
        file_listbox.configure(yscrollcommand=file_scrollbar.set)
        
        for f in yaml_files:
            file_listbox.insert(tk.END, f)
        
        file_listbox.pack(side=LEFT, fill=BOTH, expand=True)
        file_scrollbar.pack(side=RIGHT, fill=Y)
        
        # Selection buttons
        sel_frame = tb.Frame(file_frame)
        sel_frame.pack(fill=X, pady=5)
        tb.Button(sel_frame, text="Select All", bootstyle="outline-info",
                command=lambda: file_listbox.select_set(0, tk.END)).pack(side=LEFT, padx=2)
        tb.Button(sel_frame, text="Clear Selection", bootstyle="outline-secondary",
                command=lambda: file_listbox.selection_clear(0, tk.END)).pack(side=LEFT, padx=2)
        
        # Schedule settings
        settings_frame = tb.Labelframe(dialog, text="Schedule Settings", padding=10, bootstyle="info")
        settings_frame.pack(fill=X, padx=15, pady=10)
        
        # ESPHome Version selector
        version_row = tb.Frame(settings_frame)
        version_row.pack(fill=X, pady=3)
        tb.Label(version_row, text="ESPHome Version:", width=15).pack(side=LEFT)
        
        # Get available versions from the main GUI
        version_list = list(self.esphome_versions.keys()) if hasattr(self, 'esphome_versions') else []
        if not version_list:
            version_list = ["System Default"]
        
        batch_version_var = tk.StringVar(value=version_list[0] if version_list else "System Default")
        version_combo = ttk.Combobox(version_row, textvariable=batch_version_var, values=version_list, width=25)
        version_combo.pack(side=LEFT, padx=5)
        
        # Device/Target
        device_row = tb.Frame(settings_frame)
        device_row.pack(fill=X, pady=3)
        tb.Label(device_row, text="Target:", width=15).pack(side=LEFT)
        target_var = tk.StringVar(value="Auto-detect from YAML")
        tb.Entry(device_row, textvariable=target_var, width=30).pack(side=LEFT, padx=5)
        
        # Upload mode
        mode_row = tb.Frame(settings_frame)
        mode_row.pack(fill=X, pady=3)
        tb.Label(mode_row, text="Upload Mode:", width=15).pack(side=LEFT)
        mode_var = tk.StringVar(value="OTA")
        tb.Radiobutton(mode_row, text="OTA", variable=mode_var, value="OTA", bootstyle="info").pack(side=LEFT, padx=5)
        tb.Radiobutton(mode_row, text="Serial", variable=mode_var, value="Serial", bootstyle="info").pack(side=LEFT, padx=5)
        
        # Compile mode (for scheduler - when upload time arrives)
        compile_row = tb.Frame(settings_frame)
        compile_row.pack(fill=X, pady=3)
        tb.Label(compile_row, text="At Upload Time:", width=15).pack(side=LEFT)
        compile_var = tk.StringVar(value="at_upload")
        tb.Radiobutton(compile_row, text="Use Pre-compiled", variable=compile_var, value="at_schedule", bootstyle="warning").pack(side=LEFT, padx=5)
        tb.Radiobutton(compile_row, text="Compile Fresh", variable=compile_var, value="at_upload", bootstyle="info").pack(side=LEFT, padx=5)
        
        # Help text
        help_label = tb.Label(settings_frame, 
            text="💡 Tip: Use 'Compile Now' button after scheduling to compile during working hours and catch errors.",
            font=('Arial', 8), bootstyle="secondary", wraplength=400)
        help_label.pack(fill=X, pady=5)
        
        # Schedule time
        time_row = tb.Frame(settings_frame)
        time_row.pack(fill=X, pady=3)
        tb.Label(time_row, text="Schedule Time:", width=15).pack(side=LEFT)
        
        # Default to tonight at 2 AM
        tonight = datetime.now()
        if tonight.hour >= 2:
            tonight = tonight + timedelta(days=1)
        tonight = tonight.replace(hour=2, minute=0, second=0, microsecond=0)
        
        date_var = tk.StringVar(value=tonight.strftime("%Y-%m-%d"))
        time_var = tk.StringVar(value="02:00")
        tb.Entry(time_row, textvariable=date_var, width=12).pack(side=LEFT, padx=2)
        tb.Entry(time_row, textvariable=time_var, width=8).pack(side=LEFT, padx=2)
        
        # Quick buttons
        tb.Button(time_row, text="Tonight 2AM", bootstyle="outline-info",
                command=lambda: [date_var.set(tonight.strftime("%Y-%m-%d")), time_var.set("02:00")]).pack(side=LEFT, padx=5)
        tb.Button(time_row, text="Tomorrow 2AM", bootstyle="outline-info",
                command=lambda: [date_var.set((datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")), time_var.set("02:00")]).pack(side=LEFT, padx=5)
        
        # Stagger option
        stagger_row = tb.Frame(settings_frame)
        stagger_row.pack(fill=X, pady=3)
        tb.Label(stagger_row, text="Stagger uploads:", width=15).pack(side=LEFT)
        stagger_var = tk.IntVar(value=5)
        tb.Spinbox(stagger_row, from_=0, to=60, textvariable=stagger_var, width=5).pack(side=LEFT, padx=2)
        tb.Label(stagger_row, text="minutes apart").pack(side=LEFT, padx=2)
        
        # Action buttons
        btn_frame = tb.Frame(dialog)
        btn_frame.pack(fill=X, padx=15, pady=15)
        
        def schedule_batch():
            selection = file_listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select at least one YAML file")
                return
            
            try:
                base_time = datetime.strptime(f"{date_var.get()} {time_var.get()}", "%Y-%m-%d %H:%M")
                stagger_minutes = stagger_var.get()
                
                scheduled_count = 0
                for i, idx in enumerate(selection):
                    yaml_file = yaml_files[idx]
                    yaml_path = os.path.join(local_path, yaml_file)
                    
                    # Calculate staggered time
                    scheduled_time = base_time + timedelta(minutes=i * stagger_minutes)
                    scheduled_time_str = scheduled_time.strftime("%Y-%m-%d %H:%M:%S")
                    
                    # Determine target device (use filename as hostname for OTA)
                    device_name = os.path.splitext(yaml_file)[0]
                    target = device_name if mode_var.get() == "OTA" else target_var.get()
                    
                    # Store the upload job
                    self.delayed_upload_manager.store_upload_job(
                        yaml_path=yaml_path,
                        target_device=target,
                        upload_mode=mode_var.get(),
                        scheduled_time=scheduled_time_str,
                        esphome_version=batch_version_var.get(),
                        compile_mode=compile_var.get()
                    )
                    scheduled_count += 1
                
                self.refresh_scheduled_uploads_list()
                self.log_message(f">>> Scheduled {scheduled_count} uploads starting at {date_var.get()} {time_var.get()}", "auto")
                messagebox.showinfo("Batch Scheduled", f"Successfully scheduled {scheduled_count} uploads")
                dialog.destroy()
                
            except ValueError as e:
                messagebox.showerror("Invalid Date/Time", "Please enter valid date (YYYY-MM-DD) and time (HH:MM)")
        
        tb.Button(btn_frame, text="Schedule Selected", command=schedule_batch, 
                bootstyle="success", width=18).pack(side=LEFT, padx=5)
        tb.Button(btn_frame, text="Cancel", command=dialog.destroy, 
                bootstyle="secondary", width=12).pack(side=LEFT, padx=5)

    def open_schedule_dialog(self):
        """Open schedule dialog from main actions"""
        if not self.file_path.get():
            messagebox.showwarning("No File", "Please select a YAML file first")
            return
        
        # Switch to delayed upload tab
        self.notebook.select(self.delayed_upload_tab)
        
        # Update the file display in schedule tab
        if hasattr(self, 'schedule_file_var'):
            self.schedule_file_var.set(os.path.basename(self.file_path.get()))

    def on_closing(self):
        """Save recent files when application closes"""
        # Stop the upload scheduler
        if hasattr(self, 'upload_scheduler'):
            self.upload_scheduler.stop()
        
        self.save_recent_files()
        self.root.quit()

    def compile_selected_uploads(self):
        """Compile selected uploads now with live output"""
        selection = self.uploads_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select upload(s) to compile")
            return
        
        # Get all selected upload IDs and info
        uploads_to_compile = []
        for item in selection:
            values = self.uploads_tree.item(item)['values']
            upload_id = values[0]
            yaml_file = values[1]
            uploads_to_compile.append({'id': upload_id, 'file': yaml_file})
        
        if not uploads_to_compile:
            messagebox.showinfo("No Selection", "Please select uploads to compile")
            return
        
        # Compile one at a time with live output
        def compile_thread():
            total = len(uploads_to_compile)
            
            for idx, upload_info in enumerate(uploads_to_compile):
                upload_id = upload_info['id']
                yaml_file = upload_info['file']
                
                self.log_message(f"\n>>> ============================================", "auto")
                self.log_message(f">>> Compiling [{idx+1}/{total}]: {yaml_file}", "auto")
                self.log_message(f">>> ============================================", "auto")
                self.status_var.set(f"Compiling {idx+1}/{total}: {yaml_file}")
                
                # Get the upload info from database
                try:
                    conn = sqlite3.connect(self.delayed_upload_manager.db_file)
                    cursor = conn.cursor()
                    cursor.execute('SELECT yaml_path, esphome_version FROM delayed_uploads WHERE id = ?', (upload_id,))
                    result = cursor.fetchone()
                    conn.close()
                    
                    if not result:
                        self.log_message(f">>> ❌ Upload ID {upload_id} not found in database", "error")
                        continue
                    
                    yaml_path = result[0]
                    esphome_version = result[1]
                    
                    # Build the compile command (use same logic as get_esphome_command)
                    if esphome_version and esphome_version not in ["System Default", "Default", None, ""] and esphome_version in self.esphome_versions:
                        esphome_cmd = f'"{self.esphome_versions[esphome_version]["path"]}"'
                    else:
                        esphome_cmd = 'esphome'
                    
                    compile_command = f'{esphome_cmd} compile "{yaml_path}"'
                    self.log_message(f">>> Running: {compile_command}", "auto")
                    
                    # Update status to compiling
                    self.delayed_upload_manager._update_upload_compile_status(upload_id, 'compiling', None)
                    self.root.after(0, self.refresh_scheduled_uploads_list)
                    
                    # Run with live output streaming
                    process = subprocess.Popen(
                        compile_command,
                        shell=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        text=True,
                        cwd=os.path.dirname(yaml_path),
                        bufsize=1
                    )
                    
                    output_lines = []
                    for line in iter(process.stdout.readline, ''):
                        if line:
                            line_stripped = line.rstrip()
                            output_lines.append(line_stripped)
                            # Show in log
                            self.log_message(line_stripped, "auto")
                            self.log_text.see(tk.END)
                    
                    process.wait()
                    success = process.returncode == 0
                    full_output = '\n'.join(output_lines)
                    
                    # Update database
                    if success:
                        # Find the compiled firmware path
                        project_name = os.path.splitext(os.path.basename(yaml_path))[0]
                        firmware_path = os.path.join(
                            os.path.dirname(yaml_path),
                            ".esphome", "build", project_name, ".pioenvs", project_name, "firmware.bin"
                        )
                        
                        if os.path.exists(firmware_path):
                            self.delayed_upload_manager._update_upload_compile_status(upload_id, 'success', full_output, firmware_path)
                            self.log_message(f">>> ✅ Compilation successful for: {yaml_file}", "success")
                            self.log_message(f">>> 📦 Firmware: {firmware_path}", "auto")
                        else:
                            self.delayed_upload_manager._update_upload_compile_status(upload_id, 'success', full_output)
                            self.log_message(f">>> ✅ Compilation successful for: {yaml_file}", "success")
                            self.log_message(f">>> ⚠️ Firmware not found at expected path", "auto")
                    else:
                        self.delayed_upload_manager._update_upload_compile_status(upload_id, 'failed', full_output)
                        self.log_message(f">>> ❌ Compilation FAILED for: {yaml_file}", "error")
                    
                    # Refresh list after each compile
                    self.root.after(0, self.refresh_scheduled_uploads_list)
                    
                except Exception as e:
                    self.log_message(f">>> ❌ Error compiling {yaml_file}: {str(e)}", "error")
                    self.delayed_upload_manager._update_upload_compile_status(upload_id, 'failed', str(e))
            
            self.log_message(f"\n>>> ============================================", "auto")
            self.log_message(f">>> Compilation batch complete ({total} files)", "auto")
            self.log_message(f">>> ============================================\n", "auto")
            self.status_var.set("Compilation batch completed")
            self.root.after(0, self.refresh_scheduled_uploads_list)
        
        # Switch to the compiler tab to see the log
        self.notebook.select(self.compiler_tab)
        
        threading.Thread(target=compile_thread, daemon=True).start()
        self.log_message(">>> Starting batch compilation...", "auto")


    def run_command(self, command, start_time, estimated_total):
        """Run a command with the selected ESPHome version"""
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
            is_overwriting = False

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
                errors='replace',
                creationflags=subprocess.HIGH_PRIORITY_CLASS
            )

            # Track compilation success
            compilation_successful = False
            saw_firmware_bin = False
            upload_complete = False
            
            # Simple output reading
            while True:
                # CHECK FOR STOP REQUEST
                if not self.is_running:
                    if self.current_process:
                        self.current_process.terminate()
                    self.log_message( ">>> Process stopped by user", "auto")
                    break
                    
                # Read line
                output = self.current_process.stdout.readline()
                
                # Check if process is finished
                if output == '' and self.current_process.poll() is not None:
                    break
                    
                if output:
                    output = output.strip()
                    if not output: continue 
                    
                    # Extract firmware size (if applicable to your output)
                    self.extract_firmware_size(output)

                    # ---------------------------------------------------------
                    # 1. OPTIMIZED UPLOAD DETECTION
                    # ---------------------------------------------------------
                    # We strictly look for the pattern you found works best
                    if "Uploading: [" in output:
                        
                        # Update Progress Bar
                        percent_match = re.search(r'(\d+)%', output)
                        if percent_match:
                            current_percent = int(percent_match.group(1))
                            self.update_progress(current_percent)
                            self.process_status_var.set(f"Uploading... {current_percent}%")

                        # DELETE PREVIOUS LINE BEFORE WRITING NEW ONE
                        if is_overwriting:
                            try:
                                # Delete from start of last line to end of last line
                                self.log_text.delete("end-2l", "end-1l")
                            except:
                                pass
                        
                        # Write the new line
                        self.log_message(output, "auto")
                        
                        # Mark that the last thing we did was a progress bar
                        is_overwriting = True
                        
                        # Update UI to show change
                        self.root.update_idletasks()
                        continue 

                    # ---------------------------------------------------------
                    # 2. NORMAL LOGGING (NO FLICKER)
                    # ---------------------------------------------------------
                    
                    # IMPORTANT: If we just finished a block of progress bars, 
                    # we do NOT want to delete the last one (the 100% one).
                    # We simply turn off the flag and print the new line normally.
                    is_overwriting = False

                    # Status Updates based on keywords
                    if "[SUCCESS]" in output or "Successfully compiled" in output:
                        compilation_successful = True
                        self.update_progress(100)
                    elif "Building" in output and "firmware.bin" in output:
                        self.update_phase_label("Building...")
                        saw_firmware_bin = True
                    elif "Linking" in output and "firmware.elf" in output:
                        self.update_phase_label("Linking...")
                    elif "Successfully uploaded" in output:
                        self.update_phase_label("Done")
                        self.update_process_status("Upload complete")
                        self.update_progress(100)
                        upload_complete = True

                    # Print the normal line
                    self.log_message(output, "auto")
                    self.log_text.see(tk.END)
                    
                    # IGNORE esp_idf_size warnings
                    if "esp_idf_size:" not in output:
                         self.root.update_idletasks()
            # ---------------------------------------------------------
            # END OF LOOP
            # ---------------------------------------------------------
            
            return_code = self.current_process.poll()
            
            if compilation_successful or saw_firmware_bin or upload_complete:
                self.update_progress(100)
                return True
            
            self.update_progress(100)
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
            elif "success" in message_lower or ">> " in message_lower:
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