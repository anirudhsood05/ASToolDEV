# -*- coding: utf-8 -*-
"""
Family Reload Tool - Fixed Interface Implementation
Reloads families from a folder, overwriting existing families in the project.
"""
import os
import re
from pyrevit.framework import clr
from pyrevit import forms, script, revit, DB

# Import .NET interface
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import IFamilyLoadOptions

# Import existing file finder
from file_utils import FileFinder

logger = script.get_logger()
output = script.get_output()

# ============================================================================
# FIXED FAMILY RELOAD OPTIONS CLASS
# ============================================================================

class FamilyReloadOptions(IFamilyLoadOptions):
    """
    Proper .NET interface implementation for forced family reloading
    Implements IFamilyLoadOptions interface correctly
    """
    
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        """
        Called when family already exists in project
        Returns True to overwrite existing family
        """
        # Set the out parameter to True to overwrite parameter values
        overwriteParameterValues.Value = True
        return True  # Always overwrite existing families
    
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        """
        Called when shared family is found
        Returns True to continue loading
        """
        # Set the out parameter to True to overwrite parameter values
        overwriteParameterValues.Value = True
        return True

class FamilyLoaderEnhanced(object):
    """
    Enhanced family loader with proper reload capability
    """
    
    def __init__(self, path):
        """
        Parameters
        ----------
        path : str
            Absolute path to family .rfa file
        """
        self.path = path
        self.name = os.path.basename(path).replace(".rfa", "")
        self._existing_family = None
        self._check_existing_family()
    
    def _check_existing_family(self):
        """Check if family already exists in project"""
        try:
            collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
            for family in collector:
                if family.Name == self.name:
                    self._existing_family = family
                    break
        except Exception as e:
            logger.debug("Error checking existing family: {}".format(str(e)))
    
    @property
    def is_loaded(self):
        """Returns True if family is already loaded in project"""
        return self._existing_family is not None
    
    @property
    def existing_family(self):
        """Returns existing family element if loaded"""
        return self._existing_family
    
    def get_family_info(self):
        """
        Returns detailed information about family
        
        Returns
        -------
        dict
            Dictionary containing family information
        """
        info = {
            'name': self.name,
            'path': self.path,
            'is_loaded': self.is_loaded,
            'symbol_count': 0,
            'status': 'New' if not self.is_loaded else 'Existing'
        }
        
        if self.is_loaded and self._existing_family:
            try:
                symbol_ids = self._existing_family.GetFamilySymbolIds()
                info['symbol_count'] = len(symbol_ids)
            except:
                info['symbol_count'] = 0
        
        return info
    
    def reload_family(self):
        """
        Reload family with forced overwrite using proper interface
        
        Returns
        -------
        tuple
            (success: bool, message: str, family: Family or None)
        """
        try:
            # Method 1: Try with custom load options
            try:
                load_options = FamilyReloadOptions()
                ret_ref = clr.Reference[DB.Family]()
                success = revit.doc.LoadFamily(self.path, load_options, ret_ref)
                
                if success and ret_ref.Value:
                    loaded_family = ret_ref.Value
                    self._existing_family = loaded_family
                    
                    symbol_count = len(loaded_family.GetFamilySymbolIds())
                    
                    if self.is_loaded:
                        message = "Successfully reloaded family '{}' with {} types".format(
                            self.name, symbol_count)
                    else:
                        message = "Successfully loaded new family '{}' with {} types".format(
                            self.name, symbol_count)
                    
                    return True, message, loaded_family
                    
            except Exception as method1_error:
                logger.debug("Method 1 failed: {}".format(str(method1_error)))
                
                # Method 2: Fallback to simple LoadFamily
                try:
                    success = revit.doc.LoadFamily(self.path)
                    
                    if success:
                        # Find the loaded family
                        collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
                        for family in collector:
                            if family.Name == self.name:
                                loaded_family = family
                                self._existing_family = loaded_family
                                
                                symbol_count = len(loaded_family.GetFamilySymbolIds())
                                message = "Successfully loaded family '{}' with {} types (fallback method)".format(
                                    self.name, symbol_count)
                                
                                return True, message, loaded_family
                        
                        # If we get here, family was loaded but we can't find it
                        message = "Family '{}' was loaded but could not be located in project".format(self.name)
                        return True, message, None
                        
                except Exception as method2_error:
                    logger.debug("Method 2 failed: {}".format(str(method2_error)))
                    raise method2_error
            
            # If we get here, both methods failed
            return False, "Failed to load family '{}' - no successful method found".format(self.name), None
                
        except Exception as e:
            error_msg = "Error loading family '{}': {}".format(self.name, str(e))
            logger.error(error_msg)
            return False, error_msg, None
    
    def __str__(self):
        return self.name
    
    def __repr__(self):
        return "<FamilyLoaderEnhanced: {} (Loaded: {})>".format(self.name, self.is_loaded)

# ============================================================================
# MAIN SCRIPT FUNCTIONS
# ============================================================================

def main():
    """Main execution function"""
    
    # Print tool header
    output.print_md("# Family Reload Tool")
    output.print_md("Reloads families from a folder, overwriting existing families in the project.")
    output.print_md("---")
    
    # Get directory with families
    directory = forms.pick_folder("Select folder containing families to reload")
    logger.debug('Selected folder: {}'.format(directory))
    
    if directory is None:
        logger.debug('No directory selected. Exiting.')
        script.exit()
    
    output.print_md("**Selected Folder:** `{}`".format(directory))
    
    # Find family files in directory
    try:
        finder = FileFinder(directory)
        finder.search('*.rfa')
        
        # Exclude backup files (pattern: file.0001.rfa, file.0002.rfa, etc.)
        backup_pattern = r'^.*\.\d{4}\.rfa$'
        finder.exclude_by_pattern(backup_pattern)
        
        paths = list(finder.paths)  # Convert to list for consistent ordering
        
    except Exception as e:
        forms.alert("Error searching for family files: {}".format(str(e)), title="Search Error")
        script.exit()
    
    if not paths:
        forms.alert("No family files found in selected directory.", title="No Families Found")
        script.exit()
    
    output.print_md("**Found {} family files**".format(len(paths)))
    
    # Create dictionary for path lookup
    path_dict = {}
    family_info_dict = {}
    
    # Analyze families and their current status
    output.print_md("## Analyzing Families...")
    
    analysis_count = 0
    for path in paths:
        try:
            relative_path = os.path.relpath(path, directory)
            path_dict[relative_path] = path
            
            # Create family loader and get info
            family_loader = FamilyLoaderEnhanced(path)
            family_info = family_loader.get_family_info()
            family_info_dict[relative_path] = family_info
            
            analysis_count += 1
            
            logger.debug('Found family: {} (Status: {})'.format(
                family_info['name'], family_info['status']))
                
        except Exception as e:
            logger.debug('Error analyzing family {}: {}'.format(path, str(e)))
            continue
    
    output.print_md("**Successfully analyzed {} families**".format(analysis_count))
    
    # Prepare family selection options with status indicators
    family_select_options = []
    for relative_path in sorted(path_dict.keys(), key=lambda x: (x.count(os.sep), x)):
        info = family_info_dict[relative_path]
        status_indicator = "[RELOAD]" if info['is_loaded'] else "[NEW]"
        display_name = "{} {}".format(status_indicator, relative_path)
        family_select_options.append({
            'display': display_name,
            'path': relative_path,
            'info': info
        })
    
    # Show preview of families to be processed
    output.print_md("### Family Status Preview:")
    existing_count = sum(1 for info in family_info_dict.values() if info['is_loaded'])
    new_count = len(family_info_dict) - existing_count
    
    output.print_md("- **Existing families to reload:** {}".format(existing_count))
    output.print_md("- **New families to load:** {}".format(new_count))
    
    # User input -> Select families to reload
    selected_options = forms.SelectFromList.show(
        [opt['display'] for opt in family_select_options],
        title="Select Families to Reload/Load",
        width=600,
        button_name="Reload Selected Families",
        multiselect=True
    )
    
    if selected_options is None:
        logger.debug('No families selected. Exiting.')
        script.exit()
    
    # Map selected display names back to paths
    selected_families = []
    for selected_display in selected_options:
        for opt in family_select_options:
            if opt['display'] == selected_display:
                selected_families.append(opt['path'])
                break
    
    logger.debug('Selected {} families for processing'.format(len(selected_families)))
    
    # Process families
    process_families(selected_families, path_dict, family_info_dict, directory)

def process_families(selected_families, path_dict, family_info_dict, directory):
    """
    Process selected families for reloading
    """
    
    # Results tracking
    results = {
        'reloaded': [],
        'newly_loaded': [],
        'failed': [],
        'total_processed': 0
    }
    
    output.print_md("## Processing Families...")
    
    # Process families with progress bar
    max_value = len(selected_families)
    with forms.ProgressBar(title='Processing Family {value} of {max_value}',
                          cancellable=True, step=1) as pb:
        
        for count, family_path in enumerate(selected_families, 1):
            if pb.cancelled:
                output.print_md("**Operation cancelled by user.**")
                break
                
            pb.update_progress(count, max_value)
            results['total_processed'] += 1
            
            # Get family information
            family_info = family_info_dict[family_path]
            absolute_path = path_dict[family_path]
            
            output.print_md("### Processing: `{}`".format(family_info['name']))
            
            # Create family loader
            family_loader = FamilyLoaderEnhanced(absolute_path)
            
            # Perform reload within transaction
            with revit.Transaction('Reload Family: {}'.format(family_info['name'])):
                success, message, loaded_family = family_loader.reload_family()
                
                if success:
                    if family_info['is_loaded']:
                        results['reloaded'].append({
                            'name': family_info['name'],
                            'path': family_path,
                            'message': message
                        })
                        output.print_md("SUCCESS - RELOADED: {}".format(message))
                    else:
                        results['newly_loaded'].append({
                            'name': family_info['name'],
                            'path': family_path,
                            'message': message
                        })
                        output.print_md("SUCCESS - NEW: {}".format(message))
                else:
                    results['failed'].append({
                        'name': family_info['name'],
                        'path': family_path,
                        'error': message
                    })
                    output.print_md("FAILED: {}".format(message))
                
                logger.debug('Processed family: {} - Success: {}'.format(
                    family_info['name'], success))
    
    # Generate final report
    generate_final_report(results)

def generate_final_report(results):
    """Generate comprehensive final report"""
    
    output.print_md("---")
    output.print_md("# Family Reload Report")
    
    # Summary statistics
    total_processed = results['total_processed']
    reloaded_count = len(results['reloaded'])
    newly_loaded_count = len(results['newly_loaded'])
    failed_count = len(results['failed'])
    success_count = reloaded_count + newly_loaded_count
    
    output.print_md("## Summary")
    output.print_md("- **Total Processed:** {}".format(total_processed))
    output.print_md("- **Successfully Reloaded:** {}".format(reloaded_count))
    output.print_md("- **Newly Loaded:** {}".format(newly_loaded_count))
    output.print_md("- **Failed:** {}".format(failed_count))
    output.print_md("- **Success Rate:** {:.1f}%".format(
        (success_count / float(total_processed) * 100) if total_processed > 0 else 0))
    
    # Detailed results
    if results['reloaded']:
        output.print_md("## Reloaded Families")
        for item in results['reloaded']:
            output.print_md("- **{}** - {}".format(item['name'], item['message']))
    
    if results['newly_loaded']:
        output.print_md("## Newly Loaded Families")
        for item in results['newly_loaded']:
            output.print_md("- **{}** - {}".format(item['name'], item['message']))
    
    if results['failed']:
        output.print_md("## Failed Operations")
        for item in results['failed']:
            output.print_md("- **{}** - {}".format(item['name'], item['error']))
    
    # Final message
    if failed_count == 0:
        output.print_md("**All operations completed successfully!**")
    elif success_count > 0:
        output.print_md("**Completed with {} successes and {} failures.**".format(
            success_count, failed_count))
    else:
        output.print_md("**All operations failed. Please check error messages above.**")

# ============================================================================
# SCRIPT EXECUTION
# ============================================================================

# Run main function
if __name__ == '__main__':
    main()