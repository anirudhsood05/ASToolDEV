# -*- coding: utf-8 -*-
"""Batch reads phase created and demolished properties and writes to user-selected text parameters."""

__title__ = 'Phase Properties\nto Parameters'
__doc__ = 'Reads Phase Created and Phase Demolished from all model elements and writes to selected text parameters.'

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
            if info['count'] >= 3:  # Parameter appears on at least 3 elements
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
        return ["Comments", "Mark", "AUK_PhaseCreated", "AUK_PhaseDemolished", "Description"]

# Define XAML for dual parameter selection dialog
xaml_content = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AUK Phase Parameter Tool" Width="500" Height="600" 
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" 
        MinWidth="450" MinHeight="550"
        FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Area -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="AUK Phase Parameter Tool" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- Instructions -->
        <GroupBox Grid.Row="2" Header="Instructions" Padding="8" Background="White">
            <TextBlock TextWrapping="Wrap">
                This tool will read the Phase Created and Phase Demolished properties from all model elements and write them to your selected text parameters. Choose two different text-based instance parameters below.
            </TextBlock>
        </GroupBox>

        <!-- Parameter Selection -->
        <GroupBox Grid.Row="4" Header="Select Target Parameters" Padding="8" Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="100"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="100"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Available text instance parameters:" FontWeight="Medium"/>
                
                <TextBlock Grid.Row="2" Text="Select parameter for Phase Created:" FontWeight="Medium" Foreground="DarkGreen"/>
                
                <ListBox Grid.Row="3" Name="UI_created_parameter_list" 
                         SelectionMode="Single">
                </ListBox>
                
                <TextBlock Grid.Row="5" Text="Select parameter for Phase Demolished:" FontWeight="Medium" Foreground="DarkRed"/>
                
                <ListBox Grid.Row="6" Name="UI_demolished_parameter_list" 
                         SelectionMode="Single">
                </ListBox>
                
                <StackPanel Grid.Row="8" Orientation="Vertical">
                    <TextBlock Name="UI_created_parameter_info" 
                               Text="Created Phase parameter: (none selected)" 
                               FontStyle="Italic" Foreground="Gray" Margin="0,4,0,2"/>
                    <TextBlock Name="UI_demolished_parameter_info" 
                               Text="Demolished Phase parameter: (none selected)" 
                               FontStyle="Italic" Foreground="Gray" Margin="0,2,0,4"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- Element Count Info -->
        <TextBlock Grid.Row="5" Name="UI_element_count" 
                   Text="Checking element count..." 
                   Margin="0,8,0,8" FontWeight="Medium"/>

        <!-- Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="UI_process_btn" Content="Process Elements" 
                    Width="120" Height="25" Margin="0,0,8,0" IsEnabled="False"/>
            <Button Name="UI_cancel_btn" Content="Cancel" 
                    Width="80" Height="25"/>
        </StackPanel>
    </Grid>
</Window>
"""

class PhaseParameterSelectionWindow(forms.WPFWindow):
    """UI for selecting both phase parameters"""
    
    def __init__(self):
        self.selected_created_parameter = None
        self.selected_demolished_parameter = None
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
        
        # Populate both parameter lists
        for param_name in self.available_parameters:
            self.UI_created_parameter_list.Items.Add(param_name)
            self.UI_demolished_parameter_list.Items.Add(param_name)
        
        # Set defaults if available
        self.set_default_selections()
        
        # Update element count
        self.update_element_count()
    
    def set_default_selections(self):
        """Set default parameter selections"""
        # Look for common phase parameter names
        created_defaults = ["AUK_PhaseCreated", "Phase Created", "Created Phase", "Comments"]
        demolished_defaults = ["AUK_PhaseDemolished", "Phase Demolished", "Demolished Phase", "Mark"]
        
        # Set created parameter default
        for default_name in created_defaults:
            if default_name in self.available_parameters:
                for i, item in enumerate(self.UI_created_parameter_list.Items):
                    if item == default_name:
                        self.UI_created_parameter_list.SelectedIndex = i
                        break
                break
        
        # Set demolished parameter default
        for default_name in demolished_defaults:
            if default_name in self.available_parameters:
                for i, item in enumerate(self.UI_demolished_parameter_list.Items):
                    if item == default_name:
                        self.UI_demolished_parameter_list.SelectedIndex = i
                        break
                break
    
    def connect_events(self):
        """Connect UI events"""
        self.UI_created_parameter_list.SelectionChanged += self.on_created_parameter_selected
        self.UI_demolished_parameter_list.SelectionChanged += self.on_demolished_parameter_selected
        self.UI_process_btn.Click += self.on_process_clicked
        self.UI_cancel_btn.Click += self.on_cancel_clicked
    
    def on_created_parameter_selected(self, sender, args):
        """Handle created phase parameter selection"""
        if self.UI_created_parameter_list.SelectedItem:
            self.selected_created_parameter = str(self.UI_created_parameter_list.SelectedItem)
            self.UI_created_parameter_info.Text = "Created Phase parameter: {}".format(self.selected_created_parameter)
            self.update_process_button()
    
    def on_demolished_parameter_selected(self, sender, args):
        """Handle demolished phase parameter selection"""
        if self.UI_demolished_parameter_list.SelectedItem:
            self.selected_demolished_parameter = str(self.UI_demolished_parameter_list.SelectedItem)
            self.UI_demolished_parameter_info.Text = "Demolished Phase parameter: {}".format(self.selected_demolished_parameter)
            self.update_process_button()
    
    def update_process_button(self):
        """Enable process button only when both parameters are selected"""
        self.UI_process_btn.IsEnabled = (
            self.selected_created_parameter is not None and 
            self.selected_demolished_parameter is not None
        )
        
        if self.UI_process_btn.IsEnabled:
            self.update_element_count()
    
    def update_element_count(self):
        """Update element count display"""
        if (hasattr(self, 'selected_created_parameter') and self.selected_created_parameter and
            hasattr(self, 'selected_demolished_parameter') and self.selected_demolished_parameter):
            # Count elements with both selected parameters
            count = self.count_elements_with_both_parameters()
            self.UI_element_count.Text = "Elements with both parameters: {}".format(count)
        else:
            self.UI_element_count.Text = "Select both parameters to see element count"
    
    def count_elements_with_both_parameters(self):
        """Count elements that have both specified parameters"""
        try:
            # Use the enhanced collection method
            valid_elements = collect_phase_elements(
                self.selected_created_parameter, 
                self.selected_demolished_parameter
            )
            
            self.valid_elements = valid_elements
            return len(valid_elements)
            
        except Exception as e:
            output.print_md("Error counting elements: {}".format(str(e)))
            return 0
    
    def element_has_both_parameters(self, element):
        """Check if element has both specified parameters and is valid for processing"""
        try:
            # Skip invalid elements
            if not element or element.Id == DB.ElementId.InvalidElementId:
                return False
            
            # Skip element types
            if hasattr(element, 'GetType') and 'Type' in element.GetType().Name:
                return False
            
            # Must have phase properties (most model elements do)
            if not hasattr(element, 'CreatedPhaseId') or not hasattr(element, 'DemolishedPhaseId'):
                return False
            
            # Must have both target parameters
            created_param = element.LookupParameter(self.selected_created_parameter)
            demolished_param = element.LookupParameter(self.selected_demolished_parameter)
            
            created_valid = (created_param is not None and 
                            not created_param.IsReadOnly and 
                            created_param.StorageType == DB.StorageType.String)
            
            demolished_valid = (demolished_param is not None and 
                               not demolished_param.IsReadOnly and 
                               demolished_param.StorageType == DB.StorageType.String)
            
            return created_valid and demolished_valid
            
        except:
            return False
    
    def on_process_clicked(self, sender, args):
        """Handle process button click"""
        if (self.selected_created_parameter and self.selected_demolished_parameter and 
            self.valid_elements):
            
            # Check for same parameter selection
            if self.selected_created_parameter == self.selected_demolished_parameter:
                forms.alert("Please select different parameters for Created and Demolished phases.", 
                           title="Same Parameter Selected")
                return
            
            self.DialogResult = True
            self.Close()
        else:
            forms.alert("Please select both parameters first.", title="Parameters Not Selected")
    
    def on_cancel_clicked(self, sender, args):
        """Handle cancel button click"""
        self.DialogResult = False
        self.Close()

def validate_environment():
    """Validate model has phases and show phase information"""
    validation_results = {
        'is_valid': True,
        'warnings': [],
        'errors': []
    }
    
    # Check if any phases exist and list them
    try:
        phases = doc.Phases
        if not phases or phases.Size == 0:
            validation_results['errors'].append("No phases found in model.")
            validation_results['is_valid'] = False
        else:
            phase_info = []
            output.print_md("## Available Phases in Model:")
            for i, phase in enumerate(phases):
                phase_info.append("Phase {}: '{}' (ID: {})".format(
                    i + 1, phase.Name, phase.Id.IntegerValue))
                output.print_md("- **{}** (ID: {})".format(phase.Name, phase.Id.IntegerValue))
            
            validation_results['warnings'].append("Found {} phases in model".format(phases.Size))
            
            # Test with a sample element to see phase access
            try:
                collector = DB.FilteredElementCollector(doc)
                sample_elements = collector.WhereElementIsNotElementType().ToElements()[:3]
                
                output.print_md("## Sample Element Phase Analysis:")
                for elem in sample_elements:
                    try:
                        created_id = elem.CreatedPhaseId
                        demolished_id = elem.DemolishedPhaseId
                        
                        output.print_md("- Element {} ({}):".format(
                            elem.Id.IntegerValue, 
                            elem.Category.Name if elem.Category else "No Category"))
                        output.print_md("  - CreatedPhaseId: {} ({})".format(
                            created_id.IntegerValue if created_id else "None",
                            "Valid" if created_id and created_id != DB.ElementId.InvalidElementId else "Invalid"))
                        output.print_md("  - DemolishedPhaseId: {} ({})".format(
                            demolished_id.IntegerValue if demolished_id else "None",
                            "Valid" if demolished_id and demolished_id != DB.ElementId.InvalidElementId else "Invalid/Not Demolished"))
                        
                        # Try to resolve phase names
                        if created_id and created_id != DB.ElementId.InvalidElementId:
                            created_phase = doc.GetElement(created_id)
                            if created_phase:
                                output.print_md("  - Created Phase Name: '{}'".format(created_phase.Name))
                            else:
                                output.print_md("  - Created Phase: Element not found")
                        else:
                            output.print_md("  - Created Phase: Invalid ID")
                            
                    except Exception as elem_error:
                        output.print_md("  - Error accessing phase properties: {}".format(str(elem_error)))
                        
            except Exception as sample_error:
                output.print_md("Could not analyze sample elements: {}".format(str(sample_error)))
                
    except Exception as e:
        validation_results['errors'].append("Could not access phases: {}".format(str(e)))
        validation_results['is_valid'] = False
    
    return validation_results

def get_phase_created_name(element):
    """Get the name of the phase when element was created"""
    try:
        # Debug output for first few elements
        if hasattr(get_phase_created_name, 'debug_count'):
            get_phase_created_name.debug_count += 1
        else:
            get_phase_created_name.debug_count = 1
        
        if get_phase_created_name.debug_count <= 3:
            output.print_md("DEBUG - Element {}: Checking phase properties...".format(element.Id.IntegerValue))
        
        created_phase_id = element.CreatedPhaseId
        
        if get_phase_created_name.debug_count <= 3:
            output.print_md("DEBUG - Element {}: CreatedPhaseId = {}".format(
                element.Id.IntegerValue, 
                created_phase_id.IntegerValue if created_phase_id else "None"))
        
        # Check if we have a valid phase ID (not -1)
        if (created_phase_id and 
            created_phase_id != DB.ElementId.InvalidElementId and 
            created_phase_id.IntegerValue != -1):
            
            phase_element = doc.GetElement(created_phase_id)
            if phase_element and hasattr(phase_element, 'Name'):
                phase_name = phase_element.Name
                if get_phase_created_name.debug_count <= 3:
                    output.print_md("DEBUG - Element {}: Found phase name = '{}'".format(
                        element.Id.IntegerValue, phase_name))
                return phase_name
            else:
                if get_phase_created_name.debug_count <= 3:
                    output.print_md("DEBUG - Element {}: Phase element not found for ID {}".format(
                        element.Id.IntegerValue, created_phase_id.IntegerValue))
        
        # For elements with -1 or invalid phase IDs, use fallback
        if get_phase_created_name.debug_count <= 3:
            output.print_md("DEBUG - Element {}: Using fallback phase".format(element.Id.IntegerValue))
        
        # Try to get first phase as fallback
        try:
            phases = doc.Phases
            if phases and phases.Size > 0:
                # Use first phase as default for elements without explicit phase
                first_phase = phases[0]
                return "Default: {}".format(first_phase.Name)
        except:
            pass
        
        return "No Phase Assigned"
    except Exception as e:
        return "Error: {}".format(str(e))

def get_phase_demolished_name(element):
    """Get the name of the phase when element was demolished"""
    try:
        demolished_phase_id = element.DemolishedPhaseId
        
        # Check for valid demolished phase (not -1 and not invalid)
        if (demolished_phase_id and 
            demolished_phase_id != DB.ElementId.InvalidElementId and
            demolished_phase_id.IntegerValue != -1):
            
            phase_element = doc.GetElement(demolished_phase_id)
            if phase_element and hasattr(phase_element, 'Name'):
                return phase_element.Name
            else:
                return "Unknown Demolished Phase (ID: {})".format(demolished_phase_id.IntegerValue)
        
        # -1 or invalid ID means not demolished
        return "Not Demolished"
    except Exception as e:
        return "Error: {}".format(str(e))

def set_phase_parameters(element, created_phase_name, demolished_phase_name, created_param_name, demolished_param_name):
    """Set both phase parameter values"""
    results = {'created': False, 'demolished': False, 'messages': []}
    
    # Set Created Phase parameter
    try:
        created_param = element.LookupParameter(created_param_name)
        if created_param and not created_param.IsReadOnly:
            if created_param.StorageType == DB.StorageType.String:
                created_param.Set(created_phase_name)
                results['created'] = True
                results['messages'].append("Created phase set successfully")
            else:
                results['messages'].append("Created phase parameter is not text type")
        else:
            results['messages'].append("Created phase parameter not found or read-only")
    except Exception as e:
        results['messages'].append("Created phase error: {}".format(str(e)))
    
    # Set Demolished Phase parameter
    try:
        demolished_param = element.LookupParameter(demolished_param_name)
        if demolished_param and not demolished_param.IsReadOnly:
            if demolished_param.StorageType == DB.StorageType.String:
                demolished_param.Set(demolished_phase_name)
                results['demolished'] = True
                results['messages'].append("Demolished phase set successfully")
            else:
                results['messages'].append("Demolished phase parameter is not text type")
        else:
            results['messages'].append("Demolished phase parameter not found or read-only")
    except Exception as e:
        results['messages'].append("Demolished phase error: {}".format(str(e)))
    
    return results

def is_valid_phase_element(element, created_param_name, demolished_param_name):
    """Check if element should be processed for phase parameters"""
    try:
        # Skip invalid elements
        if not element or element.Id == DB.ElementId.InvalidElementId:
            return False
        
        # Skip element types
        if hasattr(element, 'GetType') and 'Type' in element.GetType().Name:
            return False
        
        # Focus on building elements that actually have meaningful phase information
        if element.Category:
            category_name = element.Category.Name
            
            # Skip system/annotation categories that don't have meaningful phases
            excluded_categories = [
                "Materials", "Worksets", "Multi-Category", "Lines", 
                "Annotations", "Analytical Nodes", "Analytical Links",
                "Sheets", "Views", "Schedules", "Legends", "Tags",
                "Detail Items", "Model Groups", "Phases", "Levels",
                "Grids", "Reference Planes", "Scope Boxes"
            ]
            if category_name in excluded_categories:
                return False
            
            # Focus on building elements that should have phase information
            preferred_categories = [
                "Walls", "Floors", "Roofs", "Ceilings", "Doors", "Windows",
                "Stairs", "Railings", "Columns", "Structural Framing",
                "Structural Foundations", "Generic Models", "Furniture",
                "Plumbing Fixtures", "Electrical Fixtures", "Mechanical Equipment",
                "Specialty Equipment", "Casework", "Entourage", "Planting",
                "Site", "Parking", "Roads", "Topography"
            ]
            
            # If category is not in preferred list, check if it's likely a building element
            if category_name not in preferred_categories:
                # Additional check - if it has "OST_" in the category, it might be a building element
                # But we'll be more restrictive to avoid system elements
                if not any(keyword in category_name.lower() for keyword in 
                          ["wall", "floor", "door", "window", "roof", "ceiling", "stair", "column"]):
                    return False
        else:
            # No category - likely a system element
            return False
        
        # Must have phase properties
        try:
            # Try to access phase properties
            created_id = element.CreatedPhaseId
            demolished_id = element.DemolishedPhaseId
            # If we can access them without error, consider it valid
        except:
            # If we can't access phase properties, skip this element
            return False
        
        # Must have both target parameters
        created_param = element.LookupParameter(created_param_name)
        demolished_param = element.LookupParameter(demolished_param_name)
        
        created_valid = (created_param is not None and 
                        not created_param.IsReadOnly and 
                        created_param.StorageType == DB.StorageType.String)
        
        demolished_valid = (demolished_param is not None and 
                           not demolished_param.IsReadOnly and 
                           demolished_param.StorageType == DB.StorageType.String)
        
        return created_valid and demolished_valid
        
    except Exception as e:
        return False

def collect_phase_elements(created_param_name, demolished_param_name):
    """Collect all valid model elements for phase parameter processing"""
    try:
        output.print_md("Collecting building elements with valid phase information...")
        
        # Based on the analysis, focus on categories that actually have elements
        # and prioritize those with valid phase assignments
        building_elements = []
        category_counts = {}
        
        # First priority: Elements with valid phase assignments (not -1)
        building_categories = [
            (DB.BuiltInCategory.OST_Walls, "Walls"),
            (DB.BuiltInCategory.OST_Floors, "Floors"),
            (DB.BuiltInCategory.OST_Doors, "Doors"),
            (DB.BuiltInCategory.OST_Windows, "Windows"),
            (DB.BuiltInCategory.OST_Roofs, "Roofs"),
            (DB.BuiltInCategory.OST_Ceilings, "Ceilings"),
            (DB.BuiltInCategory.OST_Columns, "Columns"),
            (DB.BuiltInCategory.OST_StructuralFraming, "Structural Framing"),
            (DB.BuiltInCategory.OST_StructuralFoundation, "Structural Foundations"),
            (DB.BuiltInCategory.OST_GenericModel, "Generic Models"),
            (DB.BuiltInCategory.OST_Furniture, "Furniture"),
            (DB.BuiltInCategory.OST_PlumbingFixtures, "Plumbing Fixtures"),
            (DB.BuiltInCategory.OST_ElectricalFixtures, "Electrical Fixtures"),
            (DB.BuiltInCategory.OST_MechanicalEquipment, "Mechanical Equipment"),
            (DB.BuiltInCategory.OST_SpecialityEquipment, "Specialty Equipment"),
            (DB.BuiltInCategory.OST_Casework, "Casework"),
            (DB.BuiltInCategory.OST_Stairs, "Stairs"),
            (DB.BuiltInCategory.OST_Railings, "Railings")
        ]
        
        valid_elements_found = 0
        elements_with_valid_phases = 0
        
        # Check each building category
        for built_in_cat, cat_display_name in building_categories:
            try:
                collector = DB.FilteredElementCollector(doc)
                elements = collector.OfCategory(built_in_cat).WhereElementIsNotElementType().ToElements()
                
                if elements:
                    category_total = len(elements)
                    category_valid = 0
                    category_with_phases = 0
                    
                    for element in elements:
                        # Check if element has the required parameters
                        if element_has_required_parameters(element, created_param_name, demolished_param_name):
                            building_elements.append(element)
                            category_valid += 1
                            valid_elements_found += 1
                            
                            # Check if it has valid phase assignments
                            try:
                                created_id = element.CreatedPhaseId
                                if (created_id and created_id.IntegerValue != -1 and 
                                    created_id != DB.ElementId.InvalidElementId):
                                    category_with_phases += 1
                                    elements_with_valid_phases += 1
                            except:
                                pass
                    
                    if category_valid > 0:
                        category_counts[cat_display_name] = {
                            'total': category_total,
                            'valid': category_valid,
                            'with_phases': category_with_phases
                        }
                        
            except Exception as cat_error:
                output.print_md("Error checking {}: {}".format(cat_display_name, str(cat_error)))
                continue
        
        # Report what we found
        output.print_md("## Building Elements Analysis:")
        if category_counts:
            for cat_name, counts in sorted(category_counts.items()):
                output.print_md("- **{}**: {} valid elements (with parameters) out of {} total, {} with valid phases".format(
                    cat_name, counts['valid'], counts['total'], counts['with_phases']))
        else:
            output.print_md("- No building elements found with required parameters")
        
        output.print_md("- **Total valid elements**: {}".format(valid_elements_found))
        output.print_md("- **Elements with valid phase assignments**: {}".format(elements_with_valid_phases))
        
        # If we found elements, show some samples
        if building_elements:
            output.print_md("## Sample Elements Found:")
            sample_count = min(5, len(building_elements))
            for i in range(sample_count):
                element = building_elements[i]
                cat_name = element.Category.Name if element.Category else "No Category"
                
                # Check phase assignment
                try:
                    created_id = element.CreatedPhaseId
                    phase_status = "Valid phase" if (created_id and created_id.IntegerValue != -1) else "Default phase"
                except:
                    phase_status = "Phase unknown"
                
                output.print_md("- Element {} ({}): {}".format(
                    element.Id.IntegerValue, cat_name, phase_status))
        
        return building_elements
        
    except Exception as e:
        output.print_md("ERROR collecting elements: {}".format(str(e)))
        return []

def element_has_required_parameters(element, created_param_name, demolished_param_name):
    """Check if element has both required parameters and can be processed"""
    try:
        # Skip invalid elements
        if not element or element.Id == DB.ElementId.InvalidElementId:
            return False
        
        # Must have both target parameters
        created_param = element.LookupParameter(created_param_name)
        demolished_param = element.LookupParameter(demolished_param_name)
        
        created_valid = (created_param is not None and 
                        not created_param.IsReadOnly and 
                        created_param.StorageType == DB.StorageType.String)
        
        demolished_valid = (demolished_param is not None and 
                           not demolished_param.IsReadOnly and 
                           demolished_param.StorageType == DB.StorageType.String)
        
        return created_valid and demolished_valid
        
    except Exception as e:
        return False

def process_phase_elements_batch(elements, created_param_name, demolished_param_name, batch_size=100):
    """Process elements for phase parameter updates"""
    total_elements = len(elements)
    successful_count = 0
    failed_count = 0
    results = {
        'successful': [],
        'failed': [],
        'errors': []
    }
    
    output.print_md("## Processing {} elements for phase parameters...".format(total_elements))
    
    # Process in batches
    for i in range(0, total_elements, batch_size):
        batch = elements[i:i + batch_size]
        batch_num = (i // batch_size) + 1
        total_batches = (total_elements + batch_size - 1) // batch_size
        batch_name = "Phase Parameter Update - Batch {} of {}".format(batch_num, total_batches)
        
        try:
            with revit.Transaction(batch_name):
                for element in batch:
                    try:
                        # Get phase names
                        created_phase_name = get_phase_created_name(element)
                        demolished_phase_name = get_phase_demolished_name(element)
                        
                        # Set parameters
                        param_results = set_phase_parameters(
                            element, created_phase_name, demolished_phase_name,
                            created_param_name, demolished_param_name
                        )
                        
                        if param_results['created'] and param_results['demolished']:
                            successful_count += 1
                            results['successful'].append({
                                'element_id': element.Id.IntegerValue,
                                'created_phase': created_phase_name,
                                'demolished_phase': demolished_phase_name,
                                'category': element.Category.Name if element.Category else 'Unknown'
                            })
                        else:
                            failed_count += 1
                            results['failed'].append({
                                'element_id': element.Id.IntegerValue,
                                'errors': param_results['messages']
                            })
                            
                    except Exception as elem_error:
                        failed_count += 1
                        error_msg = "Element {}: {}".format(element.Id.IntegerValue, str(elem_error))
                        results['errors'].append(error_msg)
                        continue
                
                # Progress update
                processed = min(i + batch_size, total_elements)
                progress_percent = int((processed / float(total_elements)) * 100)
                output.print_md("**Batch {} completed** - Progress: {}% ({}/{} elements)".format(
                    batch_num, progress_percent, processed, total_elements))
                
        except Exception as batch_error:
            output.print_md("ERROR in batch {}: {}".format(batch_num, str(batch_error)))
            continue
    
    return results, successful_count, failed_count

def generate_phase_summary_report(results, successful_count, failed_count, total_count, created_param, demolished_param):
    """Generate comprehensive results report for phase processing"""
    output.print_md("# Phase Parameter Update Results")
    output.print_md("## Summary")
    output.print_md("- **Total Elements Processed**: {}".format(total_count))
    output.print_md("- **Successful Updates**: {}".format(successful_count))
    output.print_md("- **Failed Updates**: {}".format(failed_count))
    output.print_md("- **Created Phase Parameter**: {}".format(created_param))
    output.print_md("- **Demolished Phase Parameter**: {}".format(demolished_param))
    
    if successful_count > 0:
        success_rate = (successful_count / float(total_count)) * 100
        output.print_md("- **Success Rate**: {:.1f}%".format(success_rate))
    
    # Show sample successful results
    if results['successful']:
        output.print_md("## Sample Successful Updates")
        for result in results['successful'][:5]:  # Show first 5
            output.print_md("- Element {} ({}): Created='{}', Demolished='{}'".format(
                result['element_id'], 
                result['category'],
                result['created_phase'], 
                result['demolished_phase']
            ))
    
    # Show errors if any
    if results['failed'] or results['errors']:
        output.print_md("## Issues Encountered")
        
        for failure in results['failed'][:5]:  # Show first 5 failures
            output.print_md("- Element {}: {}".format(failure['element_id'], "; ".join(failure['errors'])))
        
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
        warning_msg = "Project Info:\n" + "\n".join(validation['warnings']) + "\n\nContinue with phase parameter tool?"
        if not forms.alert(warning_msg, title="Phase Information", yes=True, no=True):
            return
    
    # Show parameter selection dialog
    dialog = PhaseParameterSelectionWindow()
    result = dialog.ShowDialog()
    
    if not result or not dialog.selected_created_parameter or not dialog.selected_demolished_parameter:
        return  # User cancelled
    
    created_param = dialog.selected_created_parameter
    demolished_param = dialog.selected_demolished_parameter
    
    # Use the enhanced collection method instead of dialog's cached elements
    output.print_md("# Phase Properties to Parameters Tool")
    output.print_md("**Created Phase parameter**: {}".format(created_param))
    output.print_md("**Demolished Phase parameter**: {}".format(demolished_param))
    
    # Collect valid elements with detailed reporting
    valid_elements = collect_phase_elements(created_param, demolished_param)
    
    output.print_md("**Elements to process**: {}".format(len(valid_elements)))
    
    if not valid_elements:
        forms.alert("No valid building elements found with both '{}' and '{}' parameters.".format(
            created_param, demolished_param), title="No Elements")
        return
    
    # Show sample of what will be processed
    output.print_md("## Sample Elements to be Processed:")
    sample_count = min(5, len(valid_elements))
    for i in range(sample_count):
        element = valid_elements[i]
        category_name = element.Category.Name if element.Category else "No Category"
        output.print_md("- Element {} ({}): Ready for phase parameter update".format(
            element.Id.IntegerValue, category_name))
    
    if len(valid_elements) > sample_count:
        output.print_md("... and {} more elements".format(len(valid_elements) - sample_count))
    
    # Confirm with user
    confirm_msg = "Ready to update {} elements.\n\n".format(len(valid_elements))
    confirm_msg += "Created Phase -> '{}'\n".format(created_param)
    confirm_msg += "Demolished Phase -> '{}'\n\n".format(demolished_param)
    confirm_msg += "This will read phase properties and write to the selected parameters.\n\n"
    confirm_msg += "Continue with the operation?"
    
    if not forms.alert(confirm_msg, title="Confirm Phase Parameter Update", yes=True, no=True):
        return
    
    # Process elements
    start_time = time.time()
    results, successful_count, failed_count = process_phase_elements_batch(
        valid_elements, created_param, demolished_param)
    end_time = time.time()
    
    # Generate report
    generate_phase_summary_report(results, successful_count, failed_count, 
                                 len(valid_elements), created_param, demolished_param)
    
    # Show completion message
    duration = end_time - start_time
    output.print_md("## Operation completed in {:.1f} seconds".format(duration))
    
    if failed_count == 0:
        forms.alert("Successfully updated {} elements with phase information!".format(successful_count), 
                   title="Success")
    else:
        message = "Phase parameter update completed:\n"
        message += "- {} successful updates\n".format(successful_count)
        message += "- {} failed updates\n\n".format(failed_count)
        message += "Check output window for detailed results."
        forms.alert(message, title="Partial Success")

# Run the tool
if __name__ == '__main__':
    main()