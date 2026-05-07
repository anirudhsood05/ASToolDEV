#! python2
# -*- coding: utf-8 -*-

import System.IO
import enum
import re
from pyrevit import forms

files = 0
lengths = 0
files_to_delete = []

class SIZE_UNIT(enum.Enum):
    BYTES = 1
    KB = 2
    MB = 3
    GB = 4

def convert_unit(size_in_bytes, unit):
    """Convert the size from bytes to other units like KB, MB or GB"""
    if unit == SIZE_UNIT.KB:
        return size_in_bytes/1024
    elif unit == SIZE_UNIT.MB:
        return size_in_bytes/(1024*1024)
    elif unit == SIZE_UNIT.GB:
        return size_in_bytes/(1024*1024*1024)
    else:
        return size_in_bytes

def DeleteRecursive(_dir):
    """Recursively scan directories for backup files"""
    try:
        dir_info = System.IO.DirectoryInfo(_dir)
        
        # Process files in current directory
        global files, lengths
        for file in dir_info.EnumerateFiles():
            if re.match(r'^.+?\.\d{4}\.rfa$', file.Name) or \
               re.match(r'^.+?\.\d{4}\.rvt$', file.Name):
                lengths += file.Length
                files += 1
                files_to_delete.append(file)
        
        # Recurse into subdirectories
        for subdir in dir_info.EnumerateDirectories():
            DeleteRecursive(subdir.FullName)
            
    except System.UnauthorizedAccessException:
        print("Access denied: " + _dir)
    except Exception as e:
        print("Error accessing: " + _dir + " - " + str(e))

def Delete():
    """Main deletion workflow"""
    global files, lengths, files_to_delete
    
    # Reset globals
    files = 0
    lengths = 0
    files_to_delete = []
    
    directory = forms.pick_folder("Select parent folder to purge backup files")
    
    if not directory:
        print("No folder selected. Exiting.")
        return
    
    # Scan for backup files
    print("Scanning for backup files...")
    DeleteRecursive(directory)
    
    if files == 0:
        print("No backup files found.")
        return
    
    # Confirm deletion
    size_mb = convert_unit(lengths, SIZE_UNIT.MB)
    message = 'Found {0} backup files totaling {1:.2f} MB.\n\nDelete these files?'.format(
        files, size_mb
    )
    
    if not forms.alert(message, ok=False, yes=True, no=True):
        print("Deletion cancelled by user.")
        return
    
    # Delete files
    deleted = 0
    failed = 0
    for file in files_to_delete:
        try:
            file.Delete()
            deleted += 1
        except Exception as e:
            failed += 1
            print("Failed to delete: " + file.FullName + " - " + str(e))
    
    print("\nDeletion complete:")
    print("  - Deleted: {0} files ({1:.2f} MB)".format(deleted, convert_unit(lengths, SIZE_UNIT.MB)))
    if failed > 0:
        print("  - Failed: {0} files".format(failed))

# Run the script
Delete()
print("\nDone.")