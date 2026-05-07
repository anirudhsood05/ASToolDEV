# -*- coding: utf-8 -*-
"""Batch reads workset assignments and writes to AUK_WorksetCheck parameter."""

__title__ = 'Workset to\nParameter'
__doc__ = 'Reads workset assignments from all model elements and writes to AUK_WorksetCheck parameter.'

# Import libraries
from pyrevit import script, forms, revit, DB
import time

# Get document and output
doc = revit.doc
output = script.get_output()

def validate_environment():
    """Validate model is workshared and parameter exists"""
    validation_results = {
        'is_valid': True,
        'warnings': [],
        'errors': []
    }
    
    # Check if model is workshared
    if not doc.IsWorkshared:
        validation_results['errors'].append("Model is not workshared. This tool requires a workshared model.")
        validation_results['is_valid'] = False
        return validation_results
    
    # Check if any worksets exist
    try:
        worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
        if not worksets:
            validation_results['warnings'].append("No user worksets found in model.")
    except Exception as e:
        validation_results['errors'].append("Could not access worksets: {}".format(str(e)))
        validation_results['is_valid'] = False
    
    return validation_results

def get_workset_name(element):
    """Get workset name for an element"""
    try:
        workset_id = element.WorksetId
        if workset_id and workset_id != DB.ElementId.InvalidElementId:
            workset = doc.GetWorksetTable().GetWorkset(workset_id)
            return workset.Name
        else:
            return "No Workset"
    except Exception as e:
        return "Error: {}".format(str(e))

def has_auk_workset_parameter(element):
    """Check if element has the AUK_WorksetCheck parameter"""
    try:
        param = element.LookupParameter("AUK_WorksetCheck")
        return param is not None and not param.IsReadOnly
    except:
        return False

def set_workset_parameter(element, workset_name):
    """Set the AUK_WorksetCheck parameter value"""
    try:
        param = element.LookupParameter("AUK_WorksetCheck")
        if param and not param.IsReadOnly:
            if param.StorageType == DB.StorageType.String:
                param.Set(workset_name)
                return True, "Success"
            else:
                return False, "Parameter is not text type"
        else:
            return False, "Parameter not found or read-only"
    except Exception as e:
        return False, "Error: {}".format(str(e))

def is_valid_model_element(element):
    """Check if element should be processed"""
    try:
        # Skip invalid elements
        if not element or element.Id == DB.ElementId.InvalidElementId:
            return False
        
        # Skip element types
        if hasattr(element, 'GetType') and 'Type' in element.GetType().Name:
            return False
        
        # Skip curtain panels and grids as requested
        if element.Category:
            category_name = element.Category.Name
            if category_name in ["Curtain Panels", "Curtain Wall Mullions", "Curtain Grids"]:
                return False
        
        # Must have workset assignment capability
        if not hasattr(element, 'WorksetId'):
            return False
        
        # Must have the target parameter
        if not has_auk_workset_parameter(element):
            return False
        
        return True
        
    except Exception as e:
        return False

def collect_model_elements():
    """Collect all valid model elements"""
    try:
        output.print_md("Collecting all model elements (including hidden)...")
        
        # Get all elements from model regardless of visibility
        collector = DB.FilteredElementCollector(doc)
        all_elements = collector.WhereElementIsNotElementType().ToElements()
        
        output.print_md("Found {} total elements in model".format(len(all_elements)))
        
        # Filter to valid model elements
        valid_elements = []
        elements_with_param = 0
        curtain_excluded = 0
        
        for element in all_elements:
            if is_valid_model_element(element):
                valid_elements.append(element)
            elif has_auk_workset_parameter(element):
                elements_with_param += 1
                if element.Category and element.Category.Name in ["Curtain Panels", "Curtain Wall Mullions", "Curtain Grids"]:
                    curtain_excluded += 1
        
        # Report filtering results
        output.print_md("- **Valid elements for processing**: {}".format(len(valid_elements)))
        output.print_md("- **Elements with parameter but excluded**: {}".format(elements_with_param - len(valid_elements)))
        output.print_md("- **Curtain sub-elements excluded**: {}".format(curtain_excluded))
        
        return valid_elements
        
    except Exception as e:
        output.print_md("ERROR collecting elements: {}".format(str(e)))
        return []

def process_elements_batch(elements, batch_size=100):
    """Process elements in batches with progress reporting"""
    total_elements = len(elements)
    successful_count = 0
    failed_count = 0
    results = {
        'successful': [],
        'failed': [],
        'errors': []
    }
    
    output.print_md("## Processing {} elements in batches of {}...".format(total_elements, batch_size))
    
    # Process in batches
    for i in range(0, total_elements, batch_size):
        batch = elements[i:i + batch_size]
        batch_num = (i // batch_size) + 1
        total_batches = (total_elements + batch_size - 1) // batch_size
        batch_name = "Workset Parameter Update - Batch {} of {}".format(batch_num, total_batches)
        
        try:
            with revit.Transaction(batch_name):
                for j, element in enumerate(batch):
                    try:
                        # Get workset name
                        workset_name = get_workset_name(element)
                        
                        # Set parameter
                        success, message = set_workset_parameter(element, workset_name)
                        
                        if success:
                            successful_count += 1
                            results['successful'].append({
                                'element_id': element.Id.IntegerValue,
                                'workset': workset_name,
                                'category': element.Category.Name if element.Category else 'Unknown'
                            })
                        else:
                            failed_count += 1
                            results['failed'].append({
                                'element_id': element.Id.IntegerValue,
                                'error': message
                            })
                            
                    except Exception as elem_error:
                        failed_count += 1
                        error_msg = "Element {}: {}".format(element.Id.IntegerValue, str(elem_error))
                        results['errors'].append(error_msg)
                        continue
                
                # Progress update every batch
                processed = min(i + batch_size, total_elements)
                progress_percent = int((processed / float(total_elements)) * 100)
                output.print_md("**Batch {} completed** - Progress: {}% ({}/{} elements)".format(
                    batch_num, progress_percent, processed, total_elements))
                
        except Exception as batch_error:
            output.print_md("ERROR in batch {}: {}".format(batch_num, str(batch_error)))
            continue
    
    return results, successful_count, failed_count

def generate_summary_report(results, successful_count, failed_count, total_count):
    """Generate comprehensive results report"""
    output.print_md("# Workset Parameter Update Results")
    output.print_md("## Summary")
    output.print_md("- **Total Elements Processed**: {}".format(total_count))
    output.print_md("- **Successful Updates**: {}".format(successful_count))
    output.print_md("- **Failed Updates**: {}".format(failed_count))
    
    if successful_count > 0:
        success_rate = (successful_count / float(total_count)) * 100
        output.print_md("- **Success Rate**: {:.1f}%".format(success_rate))
    

    
    # Show errors if any
    if results['failed'] or results['errors']:
        output.print_md("## Issues Encountered")
        
        for failure in results['failed'][:5]:  # Show first 5 failures
            output.print_md("- Element {}: {}".format(failure['element_id'], failure['error']))
        
        for error in results['errors'][:5]:  # Show first 5 errors
            output.print_md("- {}".format(error))
        
        if len(results['failed']) + len(results['errors']) > 10:
            output.print_md("... and more issues (see full log)")

def main():
    """Main execution function"""
    # Validate environment
    validation = validate_environment()
    if not validation['is_valid']:
        error_msg = "Cannot run tool:\n" + "\n".join(validation['errors'])
        forms.alert(error_msg, title="Environment Error")
        return
    
    if validation['warnings']:
        warning_msg = "Warnings:\n" + "\n".join(validation['warnings']) + "\n\nContinue anyway?"
        if not forms.alert(warning_msg, title="Warnings", yes=True, no=True):
            return
    
    # Collect elements
    output.print_md("# Workset to Parameter Tool")
    output.print_md("Collecting model elements...")
    
    elements = collect_model_elements()
    
    if not elements:
        forms.alert("No valid elements found with AUK_WorksetCheck parameter.", title="No Elements")
        return
    
    # Confirm with user
    confirm_msg = "Found {} elements with AUK_WorksetCheck parameter.\n\n".format(len(elements))
    confirm_msg += "This will update the parameter with current workset assignments.\n\n"
    confirm_msg += "Continue with the operation?"
    
    if not forms.alert(confirm_msg, title="Confirm Operation", yes=True, no=True):
        return
    
    # Process elements
    start_time = time.time()
    results, successful_count, failed_count = process_elements_batch(elements)
    end_time = time.time()
    
    # Generate report
    generate_summary_report(results, successful_count, failed_count, len(elements))
    
    # Show completion message
    duration = end_time - start_time
    output.print_md("## Operation completed in {:.1f} seconds".format(duration))
    
    if failed_count == 0:
        forms.alert("Successfully updated {} elements!".format(successful_count), title="Success")
    else:
        message = "Operation completed:\n"
        message += "- {} successful updates\n".format(successful_count)
        message += "- {} failed updates\n\n".format(failed_count)
        message += "Check output window for details."
        forms.alert(message, title="Partial Success")

# Run the tool
if __name__ == '__main__':
    main()