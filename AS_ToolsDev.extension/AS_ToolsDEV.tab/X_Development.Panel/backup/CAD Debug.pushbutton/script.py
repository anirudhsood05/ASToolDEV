# -*- coding: utf-8 -*-
"""CAD Block to Family Placement Tool - Debug Version
Simplified version for troubleshooting CAD link issues."""

# Import libraries
from pyrevit import script, forms, revit, DB
import wpf
import clr
from System.IO import StringReader

# Simple XAML for testing
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK CAD Debug Tool" Width="400" Height="300" ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
    FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="CAD Link Debug Tool" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- CAD Selection -->
        <GroupBox Grid.Row="2" Header="CAD Link Information" Padding="5" Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <Button Grid.Row="0" x:Name="UI_analyze_btn" Content="Analyze CAD Links" 
                        Height="25" Margin="0,5"/>
                <Button Grid.Row="1" x:Name="UI_test_selection_btn" Content="Test Selection" 
                        Height="25" Margin="0,5"/>
            </Grid>
        </GroupBox>

        <!-- Results -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto">
            <TextBlock x:Name="UI_results" TextWrapping="Wrap" 
                       Background="White" Padding="8" Margin="0,8"
                       Text="Click 'Analyze CAD Links' to start..."/>
        </ScrollViewer>

        <!-- Close Button -->
        <Button Grid.Row="4" x:Name="UI_close_btn" Content="Close" 
                Width="90" Height="25" HorizontalAlignment="Right"/>
    </Grid>
</Window>"""

class CADDebugWindow(forms.WPFWindow):
    def __init__(self):
        # Load XAML
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        # Initialize
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        
        # Connect events
        self.UI_analyze_btn.Click += self.analyze_cad_links
        self.UI_test_selection_btn.Click += self.test_selection
        self.UI_close_btn.Click += self.close_window
    
    def analyze_cad_links(self, sender, args):
        """Analyze CAD links with detailed error reporting"""
        results = ["=== CAD LINK ANALYSIS ===\n"]
        
        try:
            # Get active view
            active_view = self.doc.ActiveView
            results.append("Active View: {}".format(active_view.Name))
            results.append("View Type: {}".format(active_view.ViewType))
            
            # Get all elements in view
            all_elements = DB.FilteredElementCollector(self.doc, active_view.Id)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
            results.append("Total elements in view: {}".format(len(all_elements)))
            
            # Filter for ImportInstance (CAD links)
            import_instances = []
            for elem in all_elements:
                if isinstance(elem, DB.ImportInstance):
                    import_instances.append(elem)
            
            results.append("ImportInstance elements found: {}".format(len(import_instances)))
            
            # Analyze each CAD link
            if import_instances:
                results.append("\n--- CAD LINK DETAILS ---")
                for i, link in enumerate(import_instances):
                    results.append("\nCAD Link {}:".format(i + 1))
                    results.append("  - Element ID: {}".format(link.Id))
                    results.append("  - Category: {}".format(link.Category.Name if link.Category else "None"))
                    
                    # Try to get type information
                    try:
                        link_type = self.doc.GetElement(link.GetTypeId())
                        if link_type:
                            # Safely get name
                            type_name = "Unknown"
                            if hasattr(link_type, 'Name'):
                                type_name = link_type.Name or "Unnamed"
                            results.append("  - Type Name: {}".format(type_name))
                            results.append("  - Type ID: {}".format(link_type.Id))
                        else:
                            results.append("  - Type: Not found")
                    except Exception as e:
                        results.append("  - Type Error: {}".format(str(e)))
                    
                    # Check if it's visible
                    try:
                        visibility = link.IsHidden(active_view)
                        results.append("  - Hidden in view: {}".format(visibility))
                    except:
                        results.append("  - Visibility: Unknown")
                    
                    # Try to get geometry
                    try:
                        options = DB.Options()
                        geometry = link.get_Geometry(options)
                        if geometry:
                            geo_count = len(list(geometry))
                            results.append("  - Geometry objects: {}".format(geo_count))
                        else:
                            results.append("  - Geometry: None")
                    except Exception as e:
                        results.append("  - Geometry Error: {}".format(str(e)))
            
            else:
                results.append("\nNo CAD links found in current view.")
                results.append("\nTroubleshooting:")
                results.append("- Check if CAD file is linked (not imported)")
                results.append("- Verify CAD link is visible in current view")
                results.append("- Try different view (plan, elevation, 3D)")
            
            # Test alternative collection method
            results.append("\n--- ALTERNATIVE COLLECTION METHOD ---")
            try:
                cad_links_alt = DB.FilteredElementCollector(self.doc)\
                               .OfClass(DB.ImportInstance)\
                               .ToElements()
                results.append("Total CAD links in model: {}".format(len(cad_links_alt)))
                
                # Check which are in current view
                in_view_count = 0
                for link in cad_links_alt:
                    if link.OwnerViewId == active_view.Id or link.OwnerViewId == DB.ElementId.InvalidElementId:
                        in_view_count += 1
                results.append("CAD links visible in current view: {}".format(in_view_count))
                
            except Exception as e:
                results.append("Alternative method error: {}".format(str(e)))
            
        except Exception as e:
            results.append("CRITICAL ERROR: {}".format(str(e)))
            import traceback
            results.append("\nFull traceback:")
            results.append(traceback.format_exc())
        
        # Display results
        self.UI_results.Text = "\n".join(results)
    
    def test_selection(self, sender, args):
        """Test manual CAD link selection"""
        results = ["=== MANUAL SELECTION TEST ===\n"]
        
        try:
            # Hide window temporarily
            self.Hide()
            
            # Prompt user to select CAD link
            results.append("Please select a CAD link in the model...")
            
            # Use PickObject to select
            from Autodesk.Revit.UI.Selection import ObjectType
            selection = self.uidoc.Selection.PickObject(ObjectType.Element, "Select a CAD link")
            
            if selection:
                element = self.doc.GetElement(selection.ElementId)
                results.append("Selected element ID: {}".format(element.Id))
                results.append("Element type: {}".format(type(element).__name__))
                results.append("Category: {}".format(element.Category.Name if element.Category else "None"))
                
                if isinstance(element, DB.ImportInstance):
                    results.append("✓ Successfully selected CAD link!")
                    
                    # Get type information
                    link_type = self.doc.GetElement(element.GetTypeId())
                    if link_type and hasattr(link_type, 'Name'):
                        results.append("Type name: {}".format(link_type.Name))
                    
                    # Test geometry access
                    try:
                        geometry = element.get_Geometry(DB.Options())
                        if geometry:
                            results.append("✓ Geometry accessible")
                            
                            # Try to get first geometry instance
                            for geo in geometry:
                                if isinstance(geo, DB.GeometryInstance):
                                    results.append("✓ Found GeometryInstance")
                                    break
                        else:
                            results.append("✗ No geometry found")
                    except Exception as e:
                        results.append("✗ Geometry error: {}".format(str(e)))
                else:
                    results.append("✗ Selected element is not a CAD link")
                    results.append("Selected: {}".format(type(element).__name__))
            
        except Exception as e:
            results.append("Selection error: {}".format(str(e)))
        finally:
            # Show window again
            self.Show()
        
        # Display results
        self.UI_results.Text = "\n".join(results)
    
    def close_window(self, sender, args):
        """Close the window"""
        self.Close()

# Run the debug tool
if __name__ == '__main__':
    # Check if Revit document is available
    if not revit.doc:
        forms.alert("No Revit document open", title="Error")
    else:
        # Run the debug tool
        CADDebugWindow().ShowDialog()