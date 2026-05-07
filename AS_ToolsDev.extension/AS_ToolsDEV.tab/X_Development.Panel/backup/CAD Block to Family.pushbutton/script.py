# -*- coding: utf-8 -*-
"""CAD Block to Family Placement Tool - Final Fixed Version
Simplified version with robust family type handling."""

# Import libraries
from pyrevit import script, forms, revit, DB
import wpf
import clr
import System
from System.IO import StringReader
from System.Collections.Generic import List
from System.Windows.Threading import DispatcherPriority

# XAML (same as before but with improved status display)
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK CAD Block to Family (Fixed)" Width="450" Height="400" ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
    FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="AUK CAD Block to Family (Fixed)" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- CAD Selection -->
        <GroupBox Grid.Row="2" Header="1. CAD Link" Padding="5" Margin="0,0,0,8" Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Column="0" Content="CAD Link:" FontWeight="Medium" VerticalAlignment="Center"/>
                <ComboBox Grid.Column="1" x:Name="UI_cad_links" Height="22" Margin="8,0,0,0"/>
            </Grid>
        </GroupBox>

        <!-- Family Selection -->
        <GroupBox Grid.Row="3" Header="2. Family" Padding="5" Margin="0,0,0,8" Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Label Grid.Row="0" Grid.Column="0" Content="Family:" FontWeight="Medium" VerticalAlignment="Center"/>
                <ComboBox Grid.Row="0" Grid.Column="1" x:Name="UI_families" Height="22" Margin="8,4,8,4"/>
                <Button Grid.Row="0" Grid.Column="2" x:Name="UI_refresh_types" Content="Refresh" 
                        Width="60" Height="22"/>

                <Label Grid.Row="1" Grid.Column="0" Content="Type:" FontWeight="Medium" VerticalAlignment="Center"/>
                <ComboBox Grid.Row="1" Grid.Column="1" x:Name="UI_family_types" Height="22" Margin="8,4,8,4"/>
                <TextBlock Grid.Row="1" Grid.Column="2" x:Name="UI_type_count" Text="" 
                          FontSize="10" VerticalAlignment="Center"/>
            </Grid>
        </GroupBox>

        <!-- Host Selection -->
        <GroupBox Grid.Row="4" Header="3. Host Element" Padding="5" Margin="0,0,0,8" Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Column="0" Content="Host:" FontWeight="Medium" VerticalAlignment="Center"/>
                <TextBox Grid.Column="1" x:Name="UI_host_element" Height="22" 
                         Margin="8,0,8,0" IsReadOnly="True" 
                         Text="Click 'Select' to choose host"/>
                <Button Grid.Column="2" x:Name="UI_select_host" Content="Select" 
                        Width="70" Height="22"/>
            </Grid>
        </GroupBox>

        <!-- Results -->
        <GroupBox Grid.Row="5" Header="Status" Padding="5" Background="White">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBlock x:Name="UI_status" TextWrapping="Wrap" 
                           Text="Select CAD link, family, and host element to proceed..."/>
            </ScrollViewer>
        </GroupBox>

        <!-- Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="UI_test_btn" Content="Test Placement" 
                    Width="90" Height="25" Margin="0,0,8,0"/>
            <Button x:Name="UI_close_btn" Content="Close" 
                    Width="90" Height="25"/>
        </StackPanel>
    </Grid>
</Window>"""

class FinalCADPlacementWindow(forms.WPFWindow):
    def __init__(self):
        # Load XAML
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        # Initialize properties
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.selected_cad_link = None
        self.selected_host = None
        self.family_symbols_cache = {}  # Cache family symbols for debugging
        
        # Initialize and connect events
        self.initialize()
        self.connect_events()
    
    def initialize(self):
        """Initialize UI elements"""
        self.populate_cad_links()
        self.populate_families()
    
        # Disable test button initially
        self.UI_test_btn.IsEnabled = False
    
        # QUICK FIX: Manually trigger the CAD link selection after a short delay
        # This ensures the selected_cad_link property is set correctly
        if self.UI_cad_links.Items.Count > 0 and self.UI_cad_links.SelectedIndex >= 0:
            self.on_cad_link_changed(None, None)
    
        # Also trigger family selection if one is already selected
        if self.UI_families.Items.Count > 0 and self.UI_families.SelectedIndex >= 0:
            self.on_family_changed(None, None)
    
        # Final status check
        self.check_ready_status()
        
        # Disable test button initially
        self.UI_test_btn.IsEnabled = False
    
    def connect_events(self):
        """Connect UI events"""
        self.UI_cad_links.SelectionChanged += self.on_cad_link_changed
        self.UI_families.SelectionChanged += self.on_family_changed
        self.UI_refresh_types.Click += self.refresh_family_types
        self.UI_select_host.Click += self.select_host_element
        self.UI_test_btn.Click += self.test_placement
        self.UI_close_btn.Click += self.close_window
    
    def populate_cad_links(self):
        """Populate CAD links dropdown"""
        self.UI_cad_links.Items.Clear()
        
        try:
            # Get all CAD link instances
            active_view = self.doc.ActiveView
            cad_links = DB.FilteredElementCollector(self.doc, active_view.Id)\
                          .OfClass(DB.ImportInstance)\
                          .WhereElementIsNotElementType()\
                          .ToElements()
            
            if not cad_links:
                self.UI_cad_links.Items.Add("No CAD links found")
                self.update_status("No CAD links found in current view")
                return
            
            for i, link in enumerate(cad_links):
                try:
                    link_type = self.doc.GetElement(link.GetTypeId())
                    type_name = "CAD Link"
                    if link_type and hasattr(link_type, 'Name') and link_type.Name:
                        type_name = link_type.Name
                    
                    display_name = "{} - {}".format(type_name, link.Id)
                    self.UI_cad_links.Items.Add(display_name)
                except Exception as e:
                    display_name = "CAD Link - {}".format(link.Id)
                    self.UI_cad_links.Items.Add(display_name)
            
            if self.UI_cad_links.Items.Count > 0:
                self.UI_cad_links.SelectedIndex = 0
                # Manually trigger the selection change to set the selected_cad_link
                self.on_cad_link_changed(None, None)
                self.update_status("Found {} CAD link(s)".format(len(cad_links)))
                
        except Exception as e:
            self.UI_cad_links.Items.Add("Error loading CAD links")
            self.update_status("Error loading CAD links: {}".format(str(e)))
    
    def populate_families(self):
        """Populate families dropdown"""
        self.UI_families.Items.Clear()
        
        try:
            # Get all families
            families = DB.FilteredElementCollector(self.doc)\
                         .OfClass(DB.Family)\
                         .ToElements()
            
            # Add all families for testing
            family_list = []
            for family in families:
                if not family.IsInPlace:  # Exclude in-place families
                    family_list.append(family)
            
            # Sort and add to dropdown
            sorted_families = sorted(family_list, key=lambda f: f.Name)
            for family in sorted_families:
                self.UI_families.Items.Add(family.Name)
            
            if self.UI_families.Items.Count > 0:
                self.UI_families.SelectedIndex = 0
                self.update_status("Found {} family/families".format(len(sorted_families)))
            else:
                self.UI_families.Items.Add("No families found")
                self.update_status("No suitable families found")
                
        except Exception as e:
            self.UI_families.Items.Add("Error loading families")
            self.update_status("Error loading families: {}".format(str(e)))
    
    def on_cad_link_changed(self, sender, args):
        """Handle CAD link selection"""
        if self.UI_cad_links.SelectedItem is None:
            return
        
        try:
            selected_text = str(self.UI_cad_links.SelectedItem)
            if "No CAD links found" in selected_text or "Error" in selected_text:
                return
            
            # Extract element ID
            element_id_str = selected_text.split(" - ")[-1]
            element_id = DB.ElementId(int(element_id_str))
            self.selected_cad_link = self.doc.GetElement(element_id)
            
            self.update_status("Selected CAD link: {}".format(self.selected_cad_link.Id))
            self.check_ready_status()
            
        except Exception as e:
            self.update_status("Error selecting CAD link: {}".format(str(e)))
    
    def on_family_changed(self, sender, args):
        """Handle family selection"""
        if self.UI_families.SelectedItem is None:
            return
        
        # Clear types first
        self.UI_family_types.Items.Clear()
        self.UI_type_count.Text = ""
        
        # Add "Loading..." message
        self.UI_family_types.Items.Add("Loading types...")
        self.UI_family_types.IsEnabled = False
        
        # Populate types immediately
        self.populate_family_types()
        
        self.check_ready_status()
    
    def refresh_family_types(self, sender, args):
        """Manual refresh of family types"""
        self.on_family_changed(None, None)
    
    def populate_family_types(self):
        """Populate family types with extensive debugging"""
        
        # Re-enable dropdown
        self.UI_family_types.IsEnabled = True
        
        # Clear and start fresh
        self.UI_family_types.Items.Clear()
        
        if not self.UI_families.SelectedItem:
            return
        
        family_name = str(self.UI_families.SelectedItem)
        self.update_status("Loading types for family: {}...".format(family_name))
        
        try:
            # Method 1: Direct approach from all FamilySymbols
            self.update_status("Trying Method 1: Direct FamilySymbol collection...")
            
            # Get ALL family symbols in the project
            all_symbols = DB.FilteredElementCollector(self.doc)\
                         .OfClass(DB.FamilySymbol)\
                         .ToElements()
            
            self.update_status("Found {} total FamilySymbol elements".format(len(all_symbols)))
            
            # Find symbols for our family
            found_symbols = []
            for symbol in all_symbols:
                try:
                    # Check if this symbol belongs to our family
                    if symbol.Family and symbol.Family.Name == family_name:
                        # Try to get a usable name
                        symbol_name = self.get_symbol_name_robust(symbol)
                        if symbol_name:
                            found_symbols.append((symbol, symbol_name))
                            self.update_status("Found symbol: {} (ID: {})".format(symbol_name, symbol.Id))
                except Exception as e:
                    continue
            
            # Add found symbols to dropdown
            if found_symbols:
                for symbol, symbol_name in found_symbols:
                    self.UI_family_types.Items.Add(symbol_name)
                    # Cache the symbol for later retrieval
                    self.family_symbols_cache[symbol_name] = symbol
                
                # Select first item
                self.UI_family_types.SelectedIndex = 0
                self.UI_type_count.Text = "({} types)".format(len(found_symbols))
                self.update_status("✓ Found {} type(s) for family {}".format(len(found_symbols), family_name))
            else:
                # Try Method 2: Through Family.GetFamilySymbolIds()
                self.update_status("Method 1 failed, trying Method 2...")
                success = self.populate_types_method2(family_name)
                
                if not success:
                    # Last resort
                    self.UI_family_types.Items.Add("No accessible types")
                    self.UI_type_count.Text = "(0 types)"
                    self.update_status("✗ No accessible types found for family {}".format(family_name))
            
        except Exception as e:
            self.UI_family_types.Items.Add("Error loading types")
            self.UI_type_count.Text = "(Error)"
            self.update_status("Error in populate_family_types: {}".format(str(e)))
            script.get_logger().error("Family types error: {}".format(str(e)))
        
        # Update ready status
        self.check_ready_status()
    
    def populate_types_method2(self, family_name):
        """Alternative method using Family.GetFamilySymbolIds()"""
        try:
            # Find the family object
            families = DB.FilteredElementCollector(self.doc)\
                         .OfClass(DB.Family)\
                         .ToElements()
            
            selected_family = None
            for family in families:
                if family.Name == family_name:
                    selected_family = family
                    break
            
            if not selected_family:
                self.update_status("Family object not found: {}".format(family_name))
                return False
            
            # Get symbol IDs
            symbol_ids = selected_family.GetFamilySymbolIds()
            self.update_status("GetFamilySymbolIds() returned {} IDs".format(len(list(symbol_ids))))
            
            # Process each symbol ID
            found_count = 0
            for symbol_id in symbol_ids:
                try:
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        symbol_name = self.get_symbol_name_robust(symbol)
                        if symbol_name:
                            self.UI_family_types.Items.Add(symbol_name)
                            self.family_symbols_cache[symbol_name] = symbol
                            found_count += 1
                            self.update_status("Method 2 found: {} (ID: {})".format(symbol_name, symbol_id))
                except Exception as e:
                    self.update_status("Error processing symbol ID {}: {}".format(symbol_id, str(e)))
                    continue
            
            if found_count > 0:
                self.UI_family_types.SelectedIndex = 0
                self.UI_type_count.Text = "({} types)".format(found_count)
                self.update_status("✓ Method 2 found {} type(s)".format(found_count))
                return True
            else:
                self.update_status("✗ Method 2 found no accessible types")
                return False
                
        except Exception as e:
            self.update_status("Method 2 error: {}".format(str(e)))
            return False
    
    def get_symbol_name_robust(self, symbol):
        """Get symbol name using multiple fallback methods"""
        # Method 1: Try Name property directly
        try:
            if hasattr(symbol, 'Name'):
                name = symbol.Name
                if name and name.strip():
                    return name
        except:
            pass
        
        # Method 2: Try built-in parameter SYMBOL_NAME_PARAM
        try:
            name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if name_param and name_param.HasValue:
                name = name_param.AsString()
                if name and name.strip():
                    return name
        except:
            pass
        
        # Method 3: Try ALL_MODEL_TYPE_NAME parameter
        try:
            type_param = symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if type_param and type_param.HasValue:
                name = type_param.AsString()
                if name and name.strip():
                    return name
        except:
            pass
        
        # Method 4: Try element parameters by name
        try:
            for param in symbol.Parameters:
                if param.Definition.Name in ["Type Name", "Name", "Symbol Name"]:
                    if param.StorageType == DB.StorageType.String and param.HasValue:
                        name = param.AsString()
                        if name and name.strip():
                            return name
        except:
            pass
        
        # Method 5: Generate name from Family + ID
        try:
            family_name = symbol.Family.Name if symbol.Family else "Unknown"
            return "{}_Type_{}".format(family_name, symbol.Id)
        except:
            pass
        
        # Last resort: Just use ID
        try:
            return "Type_{}".format(symbol.Id)
        except:
            return None
    
    def select_host_element(self, sender, args):
        """Select host element"""
        try:
            self.Hide()
            
            # Use pyRevit's selection method
            selection = revit.pick_element("Select host element (ceiling/wall/floor)")
            
            if selection:
                self.selected_host = selection
                element_info = "{} (ID: {})".format(
                    selection.Category.Name if selection.Category else "Element",
                    selection.Id
                )
                self.UI_host_element.Text = element_info
                self.update_status("Selected host: {}".format(element_info))
                self.check_ready_status()
            
        except Exception as e:
            self.update_status("Error selecting host: {}".format(str(e)))
        finally:
            self.Show()
    
    def check_ready_status(self):
        """Check if all elements are selected"""
        # Check CAD link
        cad_ready = (self.selected_cad_link is not None and 
                    self.UI_cad_links.SelectedItem is not None and
                    "No CAD links found" not in str(self.UI_cad_links.SelectedItem) and
                    "Error" not in str(self.UI_cad_links.SelectedItem))
        
        # Check family
        family_ready = (self.UI_families.SelectedItem is not None and 
                       str(self.UI_families.SelectedItem) not in ["No families found", "Error loading families"])
        
        # Check family type
        type_ready = (self.UI_family_types.SelectedItem is not None and 
                     str(self.UI_family_types.SelectedItem) not in ["No accessible types", "Error loading types", "Loading types..."])
        
        # Check host
        host_ready = self.selected_host is not None
        
        all_ready = cad_ready and family_ready and type_ready and host_ready
        self.UI_test_btn.IsEnabled = all_ready
        
        # Debug the status check
        debug_info = []
        debug_info.append("CAD ready: {} (selected_cad_link: {}, dropdown: {})".format(
            cad_ready, self.selected_cad_link is not None, str(self.UI_cad_links.SelectedItem)))
        debug_info.append("Family ready: {} ({})".format(family_ready, str(self.UI_families.SelectedItem)))
        debug_info.append("Type ready: {} ({})".format(type_ready, str(self.UI_family_types.SelectedItem)))
        debug_info.append("Host ready: {} ({})".format(host_ready, self.selected_host is not None))
        
        script.get_logger().info("\n".join(debug_info))
        
        if all_ready:
            self.update_status("✓ Ready for test placement")
        else:
            missing = []
            if not cad_ready:
                missing.append("CAD link")
            if not family_ready:
                missing.append("family")
            if not type_ready:
                missing.append("family type")
            if not host_ready:
                missing.append("host element")
            
            self.update_status("Missing: {}".format(", ".join(missing)))
    
    def test_placement(self, sender, args):
        """Test placement functionality"""
        try:
            # Get selected family type from cache
            type_name = str(self.UI_family_types.SelectedItem)
            family_type = self.family_symbols_cache.get(type_name)
            
            if not family_type:
                self.update_status("Error: Could not get family type from cache")
                return
            
            # Simple test: try to place one instance at origin
            test_point = DB.XYZ(0, 0, 0)
            
            # Find a face on the host element
            host_face = self.find_simple_host_face()
            if not host_face:
                self.update_status("Error: Could not find suitable host face")
                return
            
            # Test placement
            with revit.Transaction("Test Family Placement"):
                # Activate family if needed
                if not family_type.IsActive:
                    family_type.Activate()
                    self.update_status("Activated family type: {}".format(type_name))
                
                # Create instance
                instance = self.doc.Create.NewFamilyInstance(
                    test_point,
                    family_type,
                    host_face,
                    DB.StructuralType.NonStructural
                )
                
                if instance:
                    self.update_status("✓ Test placement successful! Instance ID: {}".format(instance.Id))
                    self.update_status("Family: {} | Type: {} | Host: {}".format(
                        family_type.Family.Name, type_name, self.selected_host.Category.Name))
                else:
                    self.update_status("✗ Test placement failed - instance not created")
                    
        except Exception as e:
            self.update_status("Test placement error: {}".format(str(e)))
            script.get_logger().error("Test placement error: {}".format(str(e)))
    
    def find_simple_host_face(self):
        """Find a simple host face for testing"""
        try:
            if not self.selected_host:
                return None
            
            # Get host geometry
            options = DB.Options()
            geometry = self.selected_host.get_Geometry(options)
            
            # Find first suitable face
            for geo_obj in geometry:
                if isinstance(geo_obj, DB.Solid):
                    for face in geo_obj.Faces:
                        if isinstance(face, DB.PlanarFace):
                            # Check if face is reasonably large
                            if face.Area > 1.0:  # At least 1 square foot
                                return face
            
            return None
            
        except Exception as e:
            script.get_logger().error("Error finding host face: {}".format(str(e)))
            return None
    
    def update_status(self, message):
        """Update status text"""
        self.UI_status.Text = message
    
    def close_window(self, sender, args):
        """Close window"""
        self.Close()

# Run the tool
if __name__ == '__main__':
    if not revit.doc:
        forms.alert("No Revit document open", title="Error")
    else:
        FinalCADPlacementWindow().ShowDialog()