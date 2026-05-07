# -*- coding: utf-8 -*-
"""Family Type Debug Tool
Specifically for testing family type detection issues."""

from pyrevit import script, forms, revit, DB
import wpf
from System.IO import StringReader

# Simple debug XAML
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Family Type Debug" Width="500" Height="400" ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
    FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Family Type Debug Tool" 
                   FontWeight="Bold" FontSize="14" HorizontalAlignment="Center" Margin="0,0,0,12"/>

        <GroupBox Grid.Row="1" Header="Family Selection" Padding="5" Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Column="0" Content="Family:" FontWeight="Medium" VerticalAlignment="Center"/>
                <ComboBox Grid.Column="1" x:Name="UI_families" Height="22" Margin="8,0,8,0"/>
                <Button Grid.Column="2" x:Name="UI_analyze_btn" Content="Analyze" Width="70" Height="22"/>
            </Grid>
        </GroupBox>

        <GroupBox Grid.Row="2" Header="Found Types" Padding="5" Background="White" Margin="0,8,0,8">
            <ListBox x:Name="UI_family_types" Height="60" Margin="0,4"/>
        </GroupBox>

        <GroupBox Grid.Row="3" Header="Debug Information" Padding="5" Background="White">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBlock x:Name="UI_debug_info" TextWrapping="Wrap" FontFamily="Consolas" FontSize="10"/>
            </ScrollViewer>
        </GroupBox>

        <Button Grid.Row="4" x:Name="UI_close_btn" Content="Close" 
                Width="90" Height="25" HorizontalAlignment="Right" Margin="0,8,0,0"/>
    </Grid>
</Window>"""

class FamilyTypeDebugWindow(forms.WPFWindow):
    def __init__(self):
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        self.doc = revit.doc
        self.debug_info = []
        
        self.initialize()
        self.connect_events()
    
    def initialize(self):
        """Initialize the tool"""
        self.populate_families()
        self.update_debug("Tool initialized")
    
    def connect_events(self):
        """Connect events"""
        self.UI_analyze_btn.Click += self.analyze_family
        self.UI_close_btn.Click += self.close_window
    
    def populate_families(self):
        """Populate families list"""
        self.UI_families.Items.Clear()
        
        try:
            families = DB.FilteredElementCollector(self.doc)\
                         .OfClass(DB.Family)\
                         .ToElements()
            
            # Get families that are not in-place
            family_list = [f for f in families if not f.IsInPlace]
            sorted_families = sorted(family_list, key=lambda f: f.Name)
            
            for family in sorted_families:
                self.UI_families.Items.Add(family.Name)
            
            if self.UI_families.Items.Count > 0:
                self.UI_families.SelectedIndex = 0
            
            self.update_debug("Found {} families".format(len(sorted_families)))
            
        except Exception as e:
            self.update_debug("Error loading families: {}".format(str(e)))
    
    def analyze_family(self, sender, args):
        """Analyze selected family in detail"""
        if not self.UI_families.SelectedItem:
            self.update_debug("No family selected")
            return
        
        family_name = str(self.UI_families.SelectedItem)
        self.debug_info = ["=== ANALYZING FAMILY: {} ===".format(family_name)]
        
        try:
            # Find the family
            families = DB.FilteredElementCollector(self.doc)\
                         .OfClass(DB.Family)\
                         .ToElements()
            
            selected_family = None
            for family in families:
                if family.Name == family_name:
                    selected_family = family
                    break
            
            if not selected_family:
                self.update_debug("Family not found!")
                return
            
            self.debug_info.append("Family ID: {}".format(selected_family.Id))
            self.debug_info.append("Is In-Place: {}".format(selected_family.IsInPlace))
            
            # Method 1: GetFamilySymbolIds()
            self.debug_info.append("\n--- METHOD 1: GetFamilySymbolIds() ---")
            try:
                symbol_ids = selected_family.GetFamilySymbolIds()
                self.debug_info.append("Return type: {}".format(type(symbol_ids).__name__))
                self.debug_info.append("Has Count property: {}".format(hasattr(symbol_ids, 'Count')))
                self.debug_info.append("Has ToArray method: {}".format(hasattr(symbol_ids, 'ToArray')))
                self.debug_info.append("Has ToList method: {}".format(hasattr(symbol_ids, 'ToList')))
                
                # Try getting count
                try:
                    count = symbol_ids.Count if hasattr(symbol_ids, 'Count') else len(symbol_ids)
                    self.debug_info.append("Count: {}".format(count))
                except Exception as e:
                    self.debug_info.append("Error getting count: {}".format(str(e)))
                
                # Try different conversion methods
                symbol_list = []
                
                # Try method 1: Direct iteration
                try:
                    self.debug_info.append("Trying direct iteration...")
                    for symbol_id in symbol_ids:
                        symbol_list.append(symbol_id)
                    self.debug_info.append("✓ Direct iteration successful: {} symbols".format(len(symbol_list)))
                except Exception as e:
                    self.debug_info.append("✗ Direct iteration failed: {}".format(str(e)))
                    
                    # Try method 2: list() conversion
                    try:
                        self.debug_info.append("Trying list() conversion...")
                        symbol_list = list(symbol_ids)
                        self.debug_info.append("✓ list() conversion successful: {} symbols".format(len(symbol_list)))
                    except Exception as e2:
                        self.debug_info.append("✗ list() conversion failed: {}".format(str(e2)))
                        
                        # Try method 3: ToArray
                        try:
                            self.debug_info.append("Trying ToArray() method...")
                            symbol_list = list(symbol_ids.ToArray())
                            self.debug_info.append("✓ ToArray() successful: {} symbols".format(len(symbol_list)))
                        except Exception as e3:
                            self.debug_info.append("✗ ToArray() failed: {}".format(str(e3)))
                
                # Clear previous types and populate
                self.UI_family_types.Items.Clear()
                
                # Process symbols
                if symbol_list:
                    for i, symbol_id in enumerate(symbol_list):
                        try:
                            symbol = self.doc.GetElement(symbol_id)
                            if symbol:
                                self.UI_family_types.Items.Add(symbol.Name)
                                self.debug_info.append("Type {}: {} (ID: {})".format(i+1, symbol.Name, symbol_id))
                            else:
                                self.debug_info.append("Type {}: Symbol not found for ID {}".format(i+1, symbol_id))
                        except Exception as e:
                            self.debug_info.append("Error getting symbol {}: {}".format(symbol_id, str(e)))
                
            except Exception as e:
                self.debug_info.append("ERROR in GetFamilySymbolIds(): {}".format(str(e)))
            
            # Method 2: Alternative - Find all symbols by family
            self.debug_info.append("\n--- METHOD 2: Find all FamilySymbols ---")
            try:
                all_symbols = DB.FilteredElementCollector(self.doc)\
                             .OfClass(DB.FamilySymbol)\
                             .ToElements()
                
                family_symbols = [s for s in all_symbols if s.Family.Id == selected_family.Id]
                self.debug_info.append("Found {} symbols via alternative method".format(len(family_symbols)))
                
                for i, symbol in enumerate(family_symbols):
                    self.debug_info.append("Alt Type {}: {} (ID: {})".format(i+1, symbol.Name, symbol.Id))
                    if symbol.Name not in [str(item) for item in self.UI_family_types.Items]:
                        self.UI_family_types.Items.Add(symbol.Name)
                
            except Exception as e:
                self.debug_info.append("ERROR in alternative method: {}".format(str(e)))
            
            # Method 3: Check family categories and placement
            self.debug_info.append("\n--- METHOD 3: Family Properties ---")
            try:
                category = selected_family.FamilyCategory
                if category:
                    self.debug_info.append("Category: {}".format(category.Name))
                    self.debug_info.append("Category ID: {}".format(category.Id))
                else:
                    self.debug_info.append("No category assigned")
                
                # Check if family can be host-based
                if hasattr(selected_family, 'FamilyPlacementType'):
                    self.debug_info.append("Placement Type: {}".format(selected_family.FamilyPlacementType))
                
            except Exception as e:
                self.debug_info.append("ERROR checking family properties: {}".format(str(e)))
            
            self.update_debug_display()
            
        except Exception as e:
            self.update_debug("CRITICAL ERROR: {}".format(str(e)))
    
    def update_debug(self, message):
        """Add debug message"""
        self.debug_info.append(message)
        self.update_debug_display()
    
    def update_debug_display(self):
        """Update debug display"""
        self.UI_debug_info.Text = "\n".join(self.debug_info)
    
    def close_window(self, sender, args):
        """Close window"""
        self.Close()

# Run the debug tool
if __name__ == '__main__':
    if not revit.doc:
        forms.alert("No Revit document open", title="Error")
    else:
        FamilyTypeDebugWindow().ShowDialog()