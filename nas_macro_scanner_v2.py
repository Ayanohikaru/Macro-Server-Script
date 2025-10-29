#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Synopsis:
---------
Memory-efficient  Macro Scanner with enhanced logging for Office files.

Description:
-----------
Scans Microsoft Office macro-enabled files for hardcoded UNC paths (\\server\share) 
and mapped drives (Z:\folder). Features:
- Directory-level discovery logs
- Memory-efficient CSV caching
- VBA macro content scanning
- Detailed progress tracking
- Safe file handling

Change Log:
----------
Version  Author          Date         Changes
-------  -------------   ----------   -----------------------------------------
2.0.0    Naomi Tran     2025-10-30   Added memory-efficient processing, improved logging
1.0.0    Naomi Tran     2025-10-30   Initial version with core scanning logic
"""

import os
import re
import csv
import sys
import time
import logging
import datetime
import threading
import tempfile
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Tuple, Set
import colorama
from colorama import Fore, Style

# Initialize colorama for Windows color support
colorama.init()

# Try to import oletools for VBA scanning
try:
    from oletools.olevba import VBA_Parser
    OLETOOLS_AVAILABLE = True
except ImportError:
    OLETOOLS_AVAILABLE = False
    print(f"{Fore.YELLOW}Warning: oletools not installed. VBA macro scanning will be disabled.{Style.RESET_ALL}")
    print("To enable VBA scanning, install oletools: pip install oletools")

# =============================================================================
# Global Configuration Variables
# =============================================================================

# Input file containing s to scan (located on Desktop)
INPUT_FILE = os.path.join(os.path.expanduser('~'), 'OneDrive - nab','Desktop','shares.txt')

# Output directory name (created on Desktop)
OUTPUT_DIR_NAME = 'MetaScanner_Output'

# Number of days to skip already-scanned shares
DAYS_THRESHOLD = 7

# Batch size for processing files (adjust based on available memory)
BATCH_SIZE = 100

# Supported Office macro-enabled file extensions
ALLOWED_EXTENSIONS = [
    # Word macro-enabled documents
    '.docm', '.dotm',
    # Excel macro-enabled workbooks
    '.xlsm', '.xltm', '.xlam',
    # PowerPoint macro-enabled presentations
    '.pptm', '.potm', '.ppsm', '.ppam'
]

# Regex pattern for UNC paths (*.aur.national.com.au domain)
UNC_PATTERN = r'\\\\[a-zA-Z0-9,._-]+\.aur\.national\.com\.au\\[a-zA-Z0-9$_\-\\]+'

# Regex pattern for specific drive paths (aur.national.com.au)
DRIVE_PATTERN = r'\\\\aur\.national\.com\.au\\[a-zA-Z0-9$_\-\\]+'

# Compile regex patterns for better performance
UNC_REGEX = re.compile(UNC_PATTERN)
DRIVE_REGEX = re.compile(DRIVE_PATTERN)

# =============================================================================
# File Processing Statistics
# =============================================================================
class ScanStats:
    def __init__(self):
        self._lock = threading.Lock()
        self.total_scanned = 0
        self.with_hardcoded_paths = 0
        self.skipped_encrypted = 0
        self.skipped_permission = 0
        self.skipped_corrupted = 0
        self.skipped_recent = 0
        self.folders_scanned = 0
        self.start_time = datetime.datetime.now()

    def increment(self, stat_name: str) -> None:
        """Thread-safe increment of statistics."""
        with self._lock:
            setattr(self, stat_name, getattr(self, stat_name) + 1)

    def get_elapsed_time(self) -> str:
        """Returns formatted elapsed time since scan started."""
        elapsed = datetime.datetime.now() - self.start_time
        hours, remainder = divmod(elapsed.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# Global statistics object
stats = ScanStats()

# =============================================================================
# Logging Configuration
# =============================================================================
def setup_logging(output_dir: str) -> logging.Logger:
    """
    Configure logging with both file and console handlers.
    
    Args:
        output_dir: Directory where log file will be created
    
    Returns:
        Configured logger instance
    """
    timestamp = datetime.datetime.now().strftime('%Y%m%d-%H%M')
    log_file = os.path.join(output_dir, f'MacroScan-{timestamp}.log')
    
    logger = logging.getLogger('NASMacroScanner')
    logger.setLevel(logging.INFO)
    
    # File handler with immediate flush
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    
    # Console handler with immediate flush
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# =============================================================================
# CSV Handling Functions
# =============================================================================
def create_temp_csv(output_dir: str) -> str:
    """Create a temporary CSV file for result caching."""
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='w', 
                                          newline='', encoding='utf-8', dir=output_dir)
    writer = csv.writer(temp_file)
    writer.writerow(['FilePath', 'Status', 'Last Modified', 'Type', 'FoundString'])
    temp_file.flush()
    return temp_file.name

def append_results(temp_csv: str, results: List[tuple]) -> None:
    """Append scan results to temporary CSV."""
    with open(temp_csv, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(results)
        f.flush()

# =============================================================================
# File Processing Functions
# =============================================================================
def scan_vba_macros(file_path: str) -> List[Tuple[str, str]]:
    """
    Extract and scan VBA macro code for hardcoded paths.
    
    Args:
        file_path: Path to the Office file to scan
    
    Returns:
        List of tuples containing (found_path, 'Macro') for paths found in macro code
    """
    if not OLETOOLS_AVAILABLE:
        return []
    
    try:
        vba_parser = VBA_Parser(file_path)
        results = []
        
        if vba_parser.detect_vba_macros():
            # Extract all macro source code
            for (_, _, _, vba_code) in vba_parser.extract_macros():
                if vba_code is not None:
                    # Search for paths in the VBA code
                    unc_paths = set(UNC_REGEX.findall(vba_code))
                    drive_paths = set(DRIVE_REGEX.findall(vba_code))
                    
                    # Add all found paths with source type 'Macro'
                    for path in (unc_paths | drive_paths):
                        results.append((path, 'Macro'))
        
        vba_parser.close()
        return results
        
    except Exception as e:
        logging.warning(f"Error scanning macros in {file_path}: {str(e)}")
        return []

def scan_file(file_path: str) -> Tuple[str, List[Tuple[str, str]]]:
    """
    Scan a single file for hardcoded paths in both content and macros.
    
    Args:
        file_path: Path to the file to scan
    
    Returns:
        Tuple containing file path and list of (path, source_type) tuples
    """
    try:
        found_paths = []
        
        # Scan file content (XML/text content)
        with open(file_path, 'rb') as f:
            content = f.read().decode('utf-8', errors='ignore')
        
        # Find all matches in content using our compiled regex patterns
        unc_paths = set(UNC_REGEX.findall(content))
        drive_paths = set(DRIVE_REGEX.findall(content))
        
        # Add content-based findings
        for path in (unc_paths | drive_paths):
            found_paths.append((path, 'Content'))
        
        # Add macro-based findings if available
        found_paths.extend(scan_vba_macros(file_path))
        
        if found_paths:
            stats.increment('with_hardcoded_paths')
        stats.increment('total_scanned')
        
        return file_path, found_paths
        
    except PermissionError:
        stats.increment('skipped_permission')
        return file_path, [('ERROR: Permission Denied', '')]
    except UnicodeDecodeError:
        stats.increment('skipped_encrypted')
        return file_path, [('ERROR: Possibly Encrypted', '')]
    except Exception as e:
        stats.increment('skipped_corrupted')
        return file_path, [(f'ERROR: Corrupted or Invalid - {str(e)}', '')]

def should_skip_share(share_path: str, output_dir: str, logger: logging.Logger) -> bool:
    """
    Check if a share should be skipped based on DAYS_THRESHOLD.
    
    Args:
        share_path: Path to the 
        output_dir: Directory containing previous scan results
        logger: Logger instance for recording decisions
        
    Returns:
        bool: True if share should be skipped, False otherwise
    """
    # Get the last two segments of the share path
    share_segments = [seg for seg in share_path.strip('\\').split('\\') if seg]
    share_identifier = '-'.join(share_segments[-2:] if len(share_segments) >= 2 else share_segments)
    
    # Look for existing CSV files for this share
    for file in os.listdir(output_dir):
        if file.endswith('.csv') and share_identifier in file:
            file_path = os.path.join(output_dir, file)
            file_age = datetime.datetime.now() - datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
            
            if file_age.days < DAYS_THRESHOLD:
                logger.info(f"Skipping share {share_path} — already scanned within threshold window ({file_age.days} days ago)")
                stats.increment('skipped_recent')
                return True
    
    return False

def log_scan_outcome(log_path: str, share_path: str, status: str) -> None:
    """
    Log scan outcome to the appropriate log file.
    
    Args:
        log_path: Path to the log file (success or failure)
        share_path: Path to the network share
        status: Status message to log
    """
    with open(log_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([share_path, 
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        status])

def process_share(share_path: str, output_dir: str, logger: logging.Logger) -> None:
    """
    Process all macro-enabled files in a network share with memory-efficient batching.
    
    Args:
        share_path: Path to the network share
        output_dir: Directory where output files will be saved
        logger: Logger instance for recording progress
    """
    # Get paths to global scan logs
    success_log = os.path.join(output_dir, "scan_success_log.csv")
    failure_log = os.path.join(output_dir, "scan_failure_log.csv")
    
    # Check if share is accessible
    if not os.path.exists(share_path):
        msg = f"Share path not found or inaccessible"
        logger.warning(f"{Fore.YELLOW}{msg}: {share_path}{Style.RESET_ALL}")
        log_scan_outcome(failure_log, share_path, msg)
        stats.increment('skipped_permission')
        return
    
    # Check if we should skip this share based on DAYS_THRESHOLD
    if should_skip_share(share_path, output_dir, logger):
        return

    share_start_time = datetime.datetime.now()
    logger.info(f"Started scanning share: {share_path}")
    
    # Create temporary CSV for results
    temp_csv = create_temp_csv(output_dir)
    file_batch = []
    
    # Wrap entire share processing in try/except
    # Wrap entire share processing in try/except
    try:
        # Create temporary CSV for results
        temp_csv = create_temp_csv(output_dir)
        
        # Walk through the share
        for root, dirs, files in os.walk(share_path):
            try:
                stats.increment('folders_scanned')
                
                # Log directory information
                macro_files = [f for f in files if any(f.lower().endswith(ext) for ext in ALLOWED_EXTENSIONS)]
                logger.info(f"Scanning folder: {root} ({len(files)} files, {len(dirs)} subfolders)")
            except PermissionError as pe:
                logger.error(f"{Fore.RED}Permission denied accessing folder: {root}{Style.RESET_ALL}")
                continue
            except Exception as e:
                logger.error(f"{Fore.RED}Error accessing folder {root}: {str(e)}{Style.RESET_ALL}")
                continue
            
            if not macro_files:
                logger.info("No matching files in this folder, moving deeper…")
                continue
            
            # Log progress every 10 folders
            if stats.folders_scanned % 10 == 0:
                elapsed = stats.get_elapsed_time()
                logger.info(f"Reading subfolder ({stats.folders_scanned}/ongoing): {root}")
                logger.info(f"Elapsed time: {elapsed}")
            
            # Process files in current directory
            for file in macro_files:
                full_path = os.path.join(root, file)
                try:
                    last_modified = datetime.datetime.fromtimestamp(
                        os.path.getmtime(full_path)
                    ).strftime('%Y-%m-%d %H:%M:%S')
                    
                    # Determine macro type
                    file_ext = os.path.splitext(file)[1].lower()
                    if file_ext in ['.docm', '.dotm']:
                        macro_type = 'Word Macro'
                    elif file_ext in ['.xlsm', '.xltm', '.xlam']:
                        macro_type = 'Excel Macro'
                    else:
                        macro_type = 'PowerPoint Macro'
                    
                    # Scan file
                    _, found_paths = scan_file(full_path)
                    
                    # Write results for this file
                    results = []
                    for path, source_type in found_paths:
                        if isinstance(path, str) and path.startswith('ERROR:'):
                            results.append([full_path, path, last_modified, macro_type, ''])
                        else:
                            results.append([full_path, 'Found', last_modified, 
                                         f"{macro_type} ({source_type})", path])
                    
                    if results:
                        append_results(temp_csv, results)
                    
                except Exception as e:
                    logger.error(f"Error processing {full_path}: {str(e)}")
        
        # Generate final CSV with proper naming
        share_parts = share_path.strip('\\').split('\\')
        share_identifier = '-'.join(share_parts[-2:] if len(share_parts) >= 2 else share_parts)
        final_csv = os.path.join(output_dir, 
                               f"{share_identifier}-MacroScan-{datetime.datetime.now().strftime('%Y%m%d')}.csv")
        
        # Sort and clean up
        with open(temp_csv, 'r', newline='', encoding='utf-8') as temp, \
             open(final_csv, 'w', newline='', encoding='utf-8') as final:
            reader = csv.reader(temp)
            writer = csv.writer(final)
            
            # Copy header
            header = next(reader)
            writer.writerow(header)
            
            # Sort and write data
            data = sorted(reader, key=lambda x: x[0])  # Sort by filepath
            writer.writerows(data)
        
        # Clean up
        os.remove(temp_csv)
        logger.info("Cleaned up temporary cache files")
        
        # Log completion with duration
        duration = datetime.datetime.now() - share_start_time
        logger.info(f"{Fore.GREEN}✅ Successfully completed scan for '{share_path}' (Duration: {duration}){Style.RESET_ALL}")
        
        # Log successful completion
        log_scan_outcome(success_log, share_path, "Success")
        
    except PermissionError as pe:
        error_msg = f"Failed: PermissionError at root"
        logger.error(f"{Fore.RED}Error scanning '{share_path}': {str(pe)}{Style.RESET_ALL}")
        log_scan_outcome(failure_log, share_path, error_msg)
        if os.path.exists(temp_csv):
            try:
                os.remove(temp_csv)
                logger.info("Cleaned up temporary files after error")
            except:
                pass
    
    except Exception as e:
        error_msg = f"Failed: {str(e)}"
        logger.error(f"{Fore.RED}Error scanning '{share_path}': {str(e)}{Style.RESET_ALL}")
        log_scan_outcome(failure_log, share_path, error_msg)
        if os.path.exists(temp_csv):
            try:
                os.remove(temp_csv)
                logger.info("Cleaned up temporary files after error")
            except:
                pass

def init_scan_logs(output_dir: str) -> Tuple[str, str]:
    """
    Initialize the global success and failure log files.
    
    Args:
        output_dir: Directory where log files will be created
        
    Returns:
        Tuple containing paths to success and failure logs
    """
    success_log = os.path.join(output_dir, "scan_success_log.csv")
    failure_log = os.path.join(output_dir, "scan_failure_log.csv")
    
    # Initialize log headers if missing
    for log_path in [success_log, failure_log]:
        if not os.path.exists(log_path):
            with open(log_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['SharePath', 'ScanDate', 'Status'])
    
    return success_log, failure_log

def main():
    """Main execution function that handles setup, user input, and coordinates scanning."""
    print(f"\n{Fore.CYAN}=== Network Share Macro Scanner ==={Style.RESET_ALL}\n")
    
    # Ensure Desktop folders exist
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    output_dir = os.path.join(desktop_path, OUTPUT_DIR_NAME)
    os.makedirs(output_dir, exist_ok=True)
    
    # Initialize global scan logs
    SCAN_SUCCESS_LOG, SCAN_FAILURE_LOG = init_scan_logs(output_dir)
    
    # Setup logging
    logger = setup_logging(output_dir)
    logger.info("Starting  Macro Scanner...")
    
    # Verify input file exists
    if not os.path.exists(INPUT_FILE):
        logger.error(f"Input file not found: {INPUT_FILE}")
        print(f"\n{Fore.RED}Error: shares.txt not found on Desktop!{Style.RESET_ALL}")
        return
    
    # Get thread count from user
    try:
        thread_input = input("\nPlease enter how many CPU threads to use (default = 1): ").strip()
        thread_count = max(1, int(thread_input)) if thread_input else 1
    except ValueError:
        thread_count = 1
    logger.info(f"Using {thread_count} threads for processing")
    
    # Read shares from input file
    try:
        with open(INPUT_FILE, 'r') as f:
            shares = [line.strip() for line in f if line.strip()]
    except Exception as e:
        logger.error(f"Error reading shares.txt: {str(e)}")
        return
    
    # Process each share
    with ThreadPoolExecutor(max_workers=thread_count) as executor:
        futures = []
        for share in shares:
            future = executor.submit(process_share, share, output_dir, logger)
            futures.append(future)
        
        # Wait for all tasks to complete
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                logger.error(f"Error in thread: {str(e)}")
    
    # Print summary
    summary = f"""
{Fore.GREEN}✅ Script Complete{Style.RESET_ALL}
{Fore.CYAN}Summary:{Style.RESET_ALL}
Total scanned: {stats.total_scanned}
With hardcoded paths: {stats.with_hardcoded_paths}
Folders processed: {stats.folders_scanned}
Skipped – Recent Scan: {stats.skipped_recent}
Skipped – Encrypted: {stats.skipped_encrypted}
Skipped – Permission Denied: {stats.skipped_permission}
Skipped – Corrupted: {stats.skipped_corrupted}
Total runtime: {stats.get_elapsed_time()}
"""
    print(summary)
    logger.info(summary)
    
    print(f"\n{Fore.YELLOW}Output files saved to: {output_dir}{Style.RESET_ALL}")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{Fore.RED}Script interrupted by user.{Style.RESET_ALL}")
    except Exception as e:
        print(f"\n{Fore.RED}Unexpected error: {str(e)}{Style.RESET_ALL}")
    finally:
        colorama.deinit()