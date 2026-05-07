# -*- coding: utf-8 -*-
"""CAD Block Detection Test Tool
Quick test tool to verify CAD block detection is working."""

from pyrevit import script, forms, revit, DB
import wpf
from System.IO import StringReader

# Simple test XAML
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="CAD Block Detection Test" Width="500" Height="400" ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen" FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="CAD Block Detection Test" 
                   FontWeight="Bold" FontSize="14" HorizontalAlignment="Center" Margin="0,0,0,12"/>

        <GroupBox Grid.Row="1" Header="CAD Link Selection" Padding="5" Background="White">
            <ComboBox x:Name="UI_cad_links" Height="22" Margin="0,4"/>
        </GroupBox>

        <GroupBox Grid.Row="2" Header="Layer Analysis" Padding="5" Background="White" Margin="0,8,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <ComboBox Grid.Column="0" x:Name="UI_layers" Height="22" Margin="0,4,8,4"/>
                <Button Grid.Column="1" x:Name="UI_test_layer" Content="Test Layer" 
                        Width="80" Height="22"/>
            </Grid>
        </GroupBox>

        <GroupBox Grid.Row="3" Header="Results" Padding="5" Background="White">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBlock x:Name="UI_results" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10"/>
            </ScrollViewer>
        </GroupBox>

        <Button Grid.Row="4" x:Name="UI_close_btn" Content="Close" 
                Width="90" Height="25" HorizontalAlignment="Right" Margin="0,8,0,0"/>
    </Grid>
</Window>"""

class CADBlockTestWindow(forms.WPFWindow):
    def __init__(self):
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        self.doc = revit.doc
        self.selected_cad_link = None
        self.results = []
        
        self.initialize()
        self.connect_events()
    
    def initialize(self):
        """Initialize the tool"""
        self.populate_cad_links()
        self.add_result("Tool initialized. Select a CAD link to continue.")
    
    def connect_events(self):
        """Connect events"""
        self.UI_cad_links.SelectionChanged += self.on_cad_link_changed
        self.UI_test_layer.Click += self.test_layer
        self.UI_close_btn.Click += self.close_window
    
    def populate_cad_links(self):
        """Populate CAD links"""
        self.UI_cad_links.Items.Clear()
        
        try:
            active_view = self.doc.ActiveView
            cad_links = DB.FilteredElementCollector(self.doc, active_view.Id)\
                          .OfClass(DB.ImportInstance)\
                          .WhereElementIsNotElementType()\
                          .ToElements()
            
            for link in cad_links:
                try:
                    link_type = self.doc.GetElement(link.GetTypeId())
                    type_name = "CAD Link"
                    if link_type and hasattr(link_type, 'Name') and link_type.Name:
                        type_name = link_type.Name
                    
                    display_name = "{} - {}".format(type_name, link.Id)
                    self.UI_cad_links.Items.Add(display_name)
                except:
                    display_name = "CAD Link - {}".format(link.Id)
                    self.UI_cad_links.Items.Add(display_name)
            
            if self.UI_cad_links.Items.Count > 0:
                self.UI_cad_links.SelectedIndex = 0
                self.add_result("Found {} CAD link(s)".format(len(cad_links)))
            else:
                self.add_result("No CAD links found in current view")
            
        except Exception as e:
            self.add_result("Error loading CAD links: {}".format(str(e)))
    
    def on_cad_link_changed(self, sender, args):
        """Handle CAD link selection"""
        if self.UI_cad_links.SelectedItem is None:
            return
        
        selected_text = str(self.UI_cad_links.SelectedItem)
        element_id_str = selected_text.split(" - ")[-1]
        element_id = DB.ElementId(int(element_id_str))
        self.selected_cad_link = self.doc.GetElement(element_id)
        
        self.add_result("Selected CAD link: {}".format(selected_text))
        self.populate_layers()
    
    def populate_layers(self):
        """Populate layers from selected CAD link"""
        self.UI_layers.Items.Clear()
        
        if not self.selected_cad_link:
            return
        
        try:
            self.add_result("Analyzing CAD layers...")
            
            # Get geometry
            options = DB.Options()
            geometry = self.selected_cad_link.get_Geometry(options)
            
            # Extract layers
            layers = set()
            geometry_count = 0
            instance_count = 0
            
            for geo_obj in geometry:
                geometry_count += 1
                if isinstance(geo_obj, DB.GeometryInstance):
                    instance_count += 1
                    nested_geo = geo_obj.GetInstanceGeometry()
                    for nested in nested_geo:
                        if hasattr(nested, 'GraphicsStyleId') and nested.GraphicsStyleId:
                            style = self.doc.GetElement(nested.GraphicsStyleId)
                            if style and hasattr(style, 'GraphicsStyleCategory') and style.GraphicsStyleCategory:
                                layers.add(style.GraphicsStyleCategory.Name)
                else:
                    if hasattr(geo_obj, 'GraphicsStyleId') and geo_obj.GraphicsStyleId:
                        style = self.doc.GetElement(geo_obj.GraphicsStyleId)
                        if style and hasattr(style, 'GraphicsStyleCategory') and style.GraphicsStyleCategory:
                            layers.add(style.GraphicsStyleCategory.Name)
            
            # Add layers to dropdown
            sorted_layers = sorted(list(layers))
            for layer in sorted_layers:
                self.UI_layers.Items.Add(layer)
            
            self.add_result("Analysis complete:")
            self.add_result("- Total geometry objects: {}".format(geometry_count))
            self.add_result("- Geometry instances: {}".format(instance_count))
            self.add_result("- Unique layers found: {}".format(len(sorted_layers)))
            
            if sorted_layers:
                self.UI_layers.SelectedIndex = 0
                self.add_result("Layers: {}".format(", ".join(sorted_layers)))
            else:
                self.add_result("No layers found in CAD geometry")
            
        except Exception as e:
            self.add_result("Error analyzing layers: {}".format(str(e)))
    
    def test_layer(self, sender, args):
        """Test selected layer for potential blocks"""
        if not self.selected_cad_link or not self.UI_layers.SelectedItem:
            self.add_result("Select CAD link and layer first")
            return
        
        layer_name = str(self.UI_layers.SelectedItem)
        self.add_result("\n=== TESTING LAYER: {} ===".format(layer_name))
        
        try:
            # Get CAD geometry and transform
            options = DB.Options()
            geometry = self.selected_cad_link.get_Geometry(options)
            cad_transform = self.selected_cad_link.GetTransform()
            
            # Track findings
            instances_on_layer = 0
            geometry_on_layer = 0
            potential_blocks = []
            
            # Analyze geometry
            for geo_obj in geometry:
                if isinstance(geo_obj, DB.GeometryInstance):
                    # Check if instance is on layer
                    nested_geo = geo_obj.GetInstanceGeometry()
                    instance_on_layer = False
                    
                    for nested in nested_geo:
                        if self._is_on_layer(nested, layer_name):
                            instance_on_layer = True
                            break
                    
                    if instance_on_layer:
                        instances_on_layer += 1
                        # Get instance location
                        transform = geo_obj.Transform
                        world_transform = cad_transform.Multiply(transform)
                        location = world_transform.Origin
                        
                        potential_blocks.append({
                            'type': 'Instance',
                            'location': location,
                            'transform': transform
                        })
                        
                        self.add_result("Instance #{}: Location X:{:.2f}, Y:{:.2f}, Z:{:.2f}".format(
                            instances_on_layer, location.X, location.Y, location.Z))
                else:
                    # Check individual geometry
                    if self._is_on_layer(geo_obj, layer_name):
                        geometry_on_layer += 1
                        
                        # Try to get center point
                        center = self._get_geometry_center(geo_obj, cad_transform)
                        if center:
                            potential_blocks.append({
                                'type': 'Geometry',
                                'location': center
                            })
                            
                            self.add_result("Geometry #{}: Center X:{:.2f}, Y:{:.2f}, Z:{:.2f}".format(
                                geometry_on_layer, center.X, center.Y, center.Z))
            
            # Summary
            self.add_result("\nSUMMARY:")
            self.add_result("- Instances on layer '{}': {}".format(layer_name, instances_on_layer))
            self.add_result("- Individual geometry on layer: {}".format(geometry_on_layer))
            self.add_result("- Total potential block locations: {}".format(len(potential_blocks)))
            
            if len(potential_blocks) > 0:
                self.add_result("✓ Found {} potential placement location(s)".format(len(potential_blocks)))
            else:
                self.add_result("✗ No placement locations found on this layer")
                self.add_result("Try a different layer or check the CAD file structure")
            
        except Exception as e:
            self.add_result("Error testing layer: {}".format(str(e)))
    
    def _is_on_layer(self, geometry, layer_name):
        """Check if geometry is on specified layer"""
        try:
            if hasattr(geometry, 'GraphicsStyleId') and geometry.GraphicsStyleId:
                style = self.doc.GetElement(geometry.GraphicsStyleId)
                if style and hasattr(style, 'GraphicsStyleCategory') and style.GraphicsStyleCategory:
                    return style.GraphicsStyleCategory.Name == layer_name
        except:
            pass
        return False
    
    def _get_geometry_center(self, geometry, transform):
        """Get center point of geometry"""
        try:
            if hasattr(geometry, 'GetBoundingBox'):
                bbox = geometry.GetBoundingBox()
                if bbox:
                    center = (bbox.Min + bbox.Max) * 0.5
                    return transform.OfPoint(center)
            elif hasattr(geometry, 'StartPoint') and hasattr(geometry, 'EndPoint'):
                midpoint = (geometry.StartPoint + geometry.EndPoint) * 0.5
                return transform.OfPoint(midpoint)
        except:
            pass
        return None
    
    def add_result(self, message):
        """Add result to display"""
        self.results.append(message)
        self.UI_results.Text = "\n".join(self.results)
    
    def close_window(self, sender, args):
        """Close window"""
        self.Close()

# Run the test tool
if __name__ == '__main__':
    if not revit.doc:
        forms.alert("No Revit document open", title="Error")
    else:
        CADBlockTestWindow().ShowDialog()