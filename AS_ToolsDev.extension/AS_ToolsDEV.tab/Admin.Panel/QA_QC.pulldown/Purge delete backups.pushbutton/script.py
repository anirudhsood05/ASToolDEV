#! python2
# -*- coding: utf-8 -*-
"""
Revit Backup File Cleanup
Recursively scans directories for .####.rvt and .####.rfa backup files and deletes them.
"""

# -*- coding: utf-8 -*-
__title__ = "Purge Delete Backups"
__doc__ = "Deletes Revit project files and family backups. The files usually ending with 00??.rvt will be removed from the selected folder. Will recursively go through folders and sub-folders to delete any backup files."
__author__ = "Deyan Nenov"


import System.IO
import enum
import re
from pyrevit import forms


class SIZE_UNIT(enum.Enum):
    """Enumeration for file size units"""
    BYTES = 1
    KB = 2
    MB = 3
    GB = 4


def convert_unit(size_in_bytes, unit):
    """
    Convert file size from bytes to specified unit
    
    Args:
        size_in_bytes: Size in bytes
        unit: Target SIZE_UNIT enum value
    
    Returns:
        Converted size as float
    """
    if unit == SIZE_UNIT.KB:
        return size_in_bytes / 1024.0
    elif unit == SIZE_UNIT.MB:
        return size_in_bytes / (1024.0 * 1024.0)
    elif unit == SIZE_UNIT.GB:
        return size_in_bytes / (1024.0 * 1024.0 * 1024.0)
    else:
        return float(size_in_bytes)


def scan_directory_recursive(directory_path, backup_pattern):
    """
    Recursively scan directory for backup files matching pattern
    
    Args:
        directory_path: Root directory to scan
        backup_pattern: Compiled regex pattern for backup files
    
    Returns:
        Tuple of (list of FileInfo objects, total size in bytes, error count)
    """
    found_files = []
    total_size = 0
    error_count = 0
    
    try:
        dir_info = System.IO.DirectoryInfo(directory_path)
        
        # Process files in current directory
        try:
            for file_info in dir_info.EnumerateFiles():
                try:
                    if backup_pattern.match(file_info.Name):
                        found_files.append(file_info)
                        total_size += file_info.Length
                except Exception as e:
                    error_count += 1
                    print("  Warning: Could not process file '{0}': {1}".format(
                        file_info.Name if hasattr(file_info, 'Name') else 'unknown',
                        str(e)
                    ))
        except Exception as e:
            error_count += 1
            print("  Warning: Could not enumerate files in '{0}': {1}".format(
                directory_path, str(e)
            ))
        
        # Recurse into subdirectories
        try:
            for subdir_info in dir_info.EnumerateDirectories():
                try:
                    sub_files, sub_size, sub_errors = scan_directory_recursive(
                        subdir_info.FullName, 
                        backup_pattern
                    )
                    found_files.extend(sub_files)
                    total_size += sub_size
                    error_count += sub_errors
                except Exception as e:
                    error_count += 1
                    print("  Warning: Could not process subdirectory '{0}': {1}".format(
                        subdir_info.Name if hasattr(subdir_info, 'Name') else 'unknown',
                        str(e)
                    ))
        except Exception as e:
            error_count += 1
            print("  Warning: Could not enumerate subdirectories in '{0}': {1}".format(
                directory_path, str(e)
            ))
            
    except System.UnauthorizedAccessException:
        error_count += 1
        print("  Access denied: {0}".format(directory_path))
    except System.IO.DirectoryNotFoundException:
        error_count += 1
        print("  Directory not found: {0}".format(directory_path))
    except Exception as e:
        error_count += 1
        print("  Error accessing '{0}': {1}".format(directory_path, str(e)))
    
    return found_files, total_size, error_count


def delete_backup_files(files_to_delete):
    """
    Delete list of backup files
    
    Args:
        files_to_delete: List of FileInfo objects to delete
    
    Returns:
        Tuple of (deleted count, failed count, total size deleted)
    """
    deleted_count = 0
    failed_count = 0
    deleted_size = 0
    
    for file_info in files_to_delete:
        try:
            file_size = file_info.Length
            file_info.Delete()
            deleted_count += 1
            deleted_size += file_size
        except System.UnauthorizedAccessException:
            failed_count += 1
            print("  Access denied: {0}".format(file_info.FullName))
        except System.IO.FileNotFoundException:
            failed_count += 1
            print("  File not found (may have been deleted): {0}".format(file_info.FullName))
        except System.IO.IOException as io_ex:
            failed_count += 1
            print("  IO error deleting '{0}': {1}".format(file_info.FullName, str(io_ex)))
        except Exception as e:
            failed_count += 1
            print("  Failed to delete '{0}': {1}".format(file_info.FullName, str(e)))
    
    return deleted_count, failed_count, deleted_size


def main():
    """Main workflow for backup file cleanup"""
    
    print("Revit Backup File Cleanup")
    print("-" * 50)
    
    # Prompt user to select directory
    selected_directory = forms.pick_folder(
        title="Select parent folder to purge backup files"
    )
    
    if not selected_directory:
        print("No folder selected. Operation cancelled.")
        return
    
    # Validate directory exists
    if not System.IO.Directory.Exists(selected_directory):
        print("ERROR: Selected directory does not exist: {0}".format(selected_directory))
        return
    
    print("\nScanning directory: {0}".format(selected_directory))
    print("Looking for backup files matching pattern: *.####.rvt or *.####.rfa\n")
    
    # Compile backup file pattern (4-digit suffix before extension)
    backup_pattern = re.compile(r'^.+?\.\d{4}\.(rfa|rvt)$', re.IGNORECASE)
    
    # Scan for backup files
    files_found, total_size, scan_errors = scan_directory_recursive(
        selected_directory,
        backup_pattern
    )
    
    # Report scan results
    if scan_errors > 0:
        print("\n{0} warning(s) occurred during scan (see above).\n".format(scan_errors))
    
    if not files_found or len(files_found) == 0:
        print("No backup files found.")
        return
    
    # Display summary and confirm deletion
    size_mb = convert_unit(total_size, SIZE_UNIT.MB)
    
    print("Scan complete:")
    print("  - Files found: {0}".format(len(files_found)))
    print("  - Total size: {0:.2f} MB".format(size_mb))
    
    confirmation_message = (
        'Found {0} backup file(s) totaling {1:.2f} MB.\n\n'
        'Do you want to DELETE these files permanently?'
    ).format(len(files_found), size_mb)
    
    user_confirmed = forms.alert(
        confirmation_message,
        ok=False,
        yes=True,
        no=True
    )
    
    if not user_confirmed:
        print("\nDeletion cancelled by user.")
        return
    
    # Delete files
    print("\nDeleting backup files...")
    
    deleted_count, failed_count, deleted_size = delete_backup_files(files_found)
    
    # Report deletion results
    print("\n" + "=" * 50)
    print("Deletion complete:")
    print("  - Successfully deleted: {0} file(s) ({1:.2f} MB)".format(
        deleted_count,
        convert_unit(deleted_size, SIZE_UNIT.MB)
    ))
    
    if failed_count > 0:
        print("  - Failed to delete: {0} file(s)".format(failed_count))
        print("    (See errors above for details)")
    
    print("=" * 50)


# Script entry point
if __name__ == '__main__':
    try:
        main()
        print("\nDone.")
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user (Ctrl+C).")
    except Exception as e:
        print("\n\nFATAL ERROR: {0}".format(str(e)))
        print("Script terminated unexpectedly.")