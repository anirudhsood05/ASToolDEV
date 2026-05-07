# -*- coding: utf-8 -*-
"""Batch reads workset assignments and writes to AUK_WorksetCheck parameter."""

__title__ = 'Workset Name\nto Parameter'
__doc__ = 'Reads workset assignments from all model elements and writes to AUK_WorksetCheck parameter.'

# Import libraries
from pyrevit import script, forms, revit, DB
import wpf
import clr
from System import Windows
from System.Windows import Controls
from System.IO import StringReader
import time

# Get document and output
doc = revit.doc
output = script.get_output()

def get_text_parameters_from_elements():
    """Get all text instance parameters available on model elements"""
    try:
        output.print_md("Scanning model for text parameters...")
        
        # Get a broader sample of elements
        collector = DB.FilteredElementCollector(doc)
        all_elements = collector.WhereElementIsNotElementType().ToElements()
        
        # Sample different types of elements to find more parameters
        sample_size = min(500, len(all_elements))
        sample_elements = all_elements[:sample_size]
        
        parameter_info = {}
        
        for element in sample_elements:
            try:
                for param in element.Parameters:
                    try:
                        param_name = param.Definition.Name
                        
                        # Check if it's a text parameter (less restrictive)
                        is_text_param = (
                            param.StorageType == DB.StorageType.String and
                            not param.IsReadOnly
                        )
                        
                        if is_text_param:
                            if param_name not in parameter_info:
                                parameter_info[param_name] = {
                                    'count': 0,
                                    'is_shared': param.IsShared,
                                    'param_group': param.Definition.ParameterGroup.ToString() if hasattr(param.Definition, 'ParameterGroup') else 'Unknown'
                                }
                            parameter_info[param_name]['count'] += 1
                    except:
                        continue
            except:
                continue
        
        # Return parameters found on multiple elements
        common_params = []
        for param_name, info in parameter_info.items():
            if info['count'] >= 3:  # Lowered threshold - parameter appears on at least 3 elements
                common_params.append(param_name)
        
        output.print_md("Found {} text parameters across {} elements".format(len(common_params), sample_size))
        
        # If no parameters found, try a different approach
        if not common_params:
            output.print_md("No common parameters found. Trying alternative detection...")
            return get_text_parameters_alternative()
        
        return sorted(common_params)
        
    except Exception as e:
        output.print_md("Error finding parameters: {}".format(str(e)))
        return get_text_parameters_alternative()

def get_text_parameters_alternative():
    """Alternative method to find text parameters"""
    try:
        # Try to find parameters from specific element categories
        categories_to_check = [
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors, 
            DB.BuiltInCategory.OST_Doors,
            DB.BuiltInCategory.OST_Windows,
            DB.BuiltInCategory.OST_GenericModel
        ]
        
        found_params = set()
        
        for category in categories_to_check:
            try:
                collector = DB.FilteredElementCollector(doc)
                elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()
                
                if elements:
                    # Check first few elements of this category
                    for element in elements[:5]:
                        try:
                            for param in element.Parameters:
                                if (param.StorageType == DB.StorageType.String and 
                                    not param.IsReadOnly):
                                    found_params.add(param.Definition.Name)
                        except:
                            continue
            except:
                continue
        
        output.print_md("Alternative method found {} parameters".format(len(found_params)))
        return sorted(list(found_params))
        
    except Exception as e:
        output.print_md("Alternative detection failed: {}".format(str(e)))
        # Return some common parameter names as fallback
        return ["Comments", "Mark", "AUK_WorksetCheck", "Description"]

# Define XAML for parameter selection dialog
xaml_content = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AUK Workset Parameter Tool" Width="450" Height="400" 
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" 
        MinWidth="400" MinHeight="350"
        FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Area -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="AUK Workset Parameter Tool" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- Instructions -->
        <GroupBox Grid.Row="2" Header="Instructions" Padding="8" Background="White">
            <TextBlock TextWrapping="Wrap">
                This tool will read the workset assignment from all model elements and write the workset name to your selected text parameter. Choose a text-based instance parameter below.
            </TextBlock>
        </GroupBox>

        <!-- Parameter Selection -->
        <GroupBox Grid.Row="3" Header="Select Target Parameter" Padding="8" Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Available text instance parameters:" FontWeight="Medium"/>
                
                <ListBox Grid.Row="2" Name="UI_parameter_list" 
                         MinHeight="120" MaxHeight="200"
                         SelectionMode="Single">
                </ListBox>
                
                <TextBlock Grid.Row="3" Name="UI_parameter_info" 
                           Text="Select a parameter from the list above" 
                           FontStyle="Italic" Foreground="Gray" 
                           Margin="0,8,0,0"/>
            </Grid>
        </GroupBox>

        <!-- Element Count Info -->
        <TextBlock Grid.Row="4" Name="UI_element_count" 
                   Text="Checking element count..." 
                   Margin="0,8,0,8" FontWeight="Medium"/>

        <!-- Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="UI_process_btn" Content="Process Elements" 
                    Width="120" Height="25" Margin="0,0,8,0" IsEnabled="False"/>
            <Button Name="UI_cancel_btn" Content="Cancel" 
                    Width="80" Height="25"/>
        </StackPanel>
    </Grid>
</Window>
"""

class ParameterSelectionWindow(forms.WPFWindow):
    """UI for selecting target parameter"""
    
    def __init__(self):
        self.selected_parameter = None
        self.valid_elements = []
        
        # Load XAML
        xaml_stream = StringReader(xaml_content)
        wpf.LoadComponent(self, xaml_stream)
        
        # Initialize
        self.initialize()
        self.connect_events()
    
    def initialize(self):
        """Initialize the window"""
        # Get available parameters
        self.available_parameters = get_text_parameters_from_elements()
        
        # Populate parameter list
        for param_name in self.available_parameters:
            self.UI_parameter_list.Items.Add(param_name)
        
        # Select default if available
        if "AUK_WorksetCheck" in self.available_parameters:
            for i, item in enumerate(self.UI_parameter_list.Items):
                if item == "AUK_WorksetCheck":
                    self.UI_parameter_list.SelectedIndex = i
                    break
        elif self.available_parameters:
            self.UI_parameter_list.SelectedIndex = 0
        
        # Update element count
        self.update_element_count()
    
    def connect_events(self):
        """Connect UI events"""
        self.UI_parameter_list.SelectionChanged += self.on_parameter_selected
        self.UI_process_btn.Click += self.on_process_clicked
        self.UI_cancel_btn.Click += self.on_cancel_clicked
    
    def on_parameter_selected(self, sender, args):
        """Handle parameter selection"""
        if self.UI_parameter_list.SelectedItem:
            self.selected_parameter = str(self.UI_parameter_list.SelectedItem)
            self.UI_parameter_info.Text = "Selected: {}".format(self.selected_parameter)
            
            # Update element count for selected parameter
            self.update_element_count()
            
            # Enable process button
            self.UI_process_btn.IsEnabled = True
        else:
            self.UI_process_btn.IsEnabled = False
    
    def update_element_count(self):
        """Update element count display"""
        if hasattr(self, 'selected_parameter') and self.selected_parameter:
            # Count elements with selected parameter
            count = self.count_elements_with_parameter(self.selected_parameter)
            self.UI_element_count.Text = "Elements with '{}': {}".format(
                self.selected_parameter, count)
        else:
            self.UI_element_count.Text = "Select a parameter to see element count"
    
    def count_elements_with_parameter(self, param_name):
        """Count elements that have the specified parameter"""
        try:
            collector = DB.FilteredElementCollector(doc)
            all_elements = collector.WhereElementIsNotElementType().ToElements()
            
            count = 0
            valid_elements = []
            
            for element in all_elements:
                if self.element_has_parameter(element, param_name):
                    count += 1
                    valid_elements.append(element)
            
            self.valid_elements = valid_elements
            return count
            
        except Exception as e:
            return 0
    
    def element_has_parameter(self, element, param_name):
        """Check if element has the specified parameter and is valid for processing"""
        try:
            # Skip invalid elements
            if not element or element.Id == DB.ElementId.InvalidElementId:
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
            param = element.LookupParameter(param_name)
            return param is not None and not param.IsReadOnly and param.StorageType == DB.StorageType.String
            
        except:
            return False
    
    def on_process_clicked(self, sender, args):
        """Handle process button click"""
        if self.selected_parameter and self.valid_elements:
            self.DialogResult = True
            self.Close()
        else:
            forms.alert("Please select a parameter first.", title="No Parameter Selected")
    
    def on_cancel_clicked(self, sender, args):
        """Handle cancel button click"""
        self.DialogResult = False
        self.Close()

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

def has_target_parameter(element, param_name):
    """Check if element has the target parameter"""
    try:
        param = element.LookupParameter(param_name)
        return param is not None and not param.IsReadOnly and param.StorageType == DB.StorageType.String
    except:
        return False

def set_workset_parameter(element, workset_name, param_name):
    """Set the target parameter value"""
    try:
        param = element.LookupParameter(param_name)
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

def is_valid_model_element(element, param_name):
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
        if not has_target_parameter(element, param_name):
            return False
        
        return True
        
    except Exception as e:
        return False

def collect_model_elements(param_name):
    """Collect all valid model elements for the specified parameter"""
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
            if is_valid_model_element(element, param_name):
                valid_elements.append(element)
            elif has_target_parameter(element, param_name):
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

def process_elements_batch(elements, param_name, batch_size=100):
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
                        success, message = set_workset_parameter(element, workset_name, param_name)
                        
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
    
    # Show parameter selection dialog
    dialog = ParameterSelectionWindow()
    result = dialog.ShowDialog()
    
    if not result or not dialog.selected_parameter:
        return  # User cancelled
    
    selected_param = dialog.selected_parameter
    valid_elements = dialog.valid_elements
    
    # Start processing
    output.print_md("# Workset to Parameter Tool")
    output.print_md("**Selected parameter**: {}".format(selected_param))
    output.print_md("**Elements to process**: {}".format(len(valid_elements)))
    
    if not valid_elements:
        forms.alert("No valid elements found with '{}' parameter.".format(selected_param), title="No Elements")
        return
    
    # Confirm with user
    confirm_msg = "Ready to update {} elements.\n\n".format(len(valid_elements))
    confirm_msg += "Parameter: '{}'\n".format(selected_param)
    confirm_msg += "Action: Write workset names to this parameter\n\n"
    confirm_msg += "Continue with the operation?"
    
    if not forms.alert(confirm_msg, title="Confirm Operation", yes=True, no=True):
        return
    
    # Process elements
    start_time = time.time()
    results, successful_count, failed_count = process_elements_batch(valid_elements, selected_param)
    end_time = time.time()
    
    # Generate report
    generate_summary_report(results, successful_count, failed_count, len(valid_elements))
    
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