# -*- coding: utf-8 -*-
"""CAD Block to Family Placer - Ultra Simple Version
Absolute minimum to avoid crashes."""

from pyrevit import script, forms, revit, DB
import wpf
from System.IO import StringReader

# Minimal XAML
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="CAD Block Placer - Ultra Simple" Width="450" Height="300" 
    WindowStartupLocation="CenterScreen" FontFamily="Arial" FontSize="12">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Ultra Simple CAD Block Placer" 
                   FontWeight="Bold" FontSize="14" HorizontalAlignment="Center" Margin="0,0,0,12"/>

        <!-- Family Selection Only -->
        <GroupBox Grid.Row="1" Header="Select Family" Margin="0,0,0,8">
            <Grid Margin="8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <ComboBox Grid.Column="0" x:Name="UI_family_combo" Height="22"/>
                <ComboBox Grid.Column="2" x:Name="UI_type_combo" Height="22"/>
            </Grid>
        </GroupBox>

        <!-- Level Selection -->
        <GroupBox Grid.Row="2" Header="Select Level" Margin="0,0,0,8">
            <ComboBox x:Name="UI_level_combo" Height="22" Margin="8"/>
        </GroupBox>

        <!-- Simple Options -->
        <GroupBox Grid.Row="3" Header="Options" Margin="0,0,0,8">
            <StackPanel Margin="8">
                <TextBox x:Name="UI_layer_textbox" Height="22" 
                         Text="Enter CAD layer name (optional)"/>
                <CheckBox x:Name="UI_place_origin_chk" Content="Place single family at origin (test)" 
                          IsChecked="True" Margin="0,4,0,0"/>
            </StackPanel>
        </GroupBox>

        <!-- Results -->
        <ScrollViewer Grid.Row="4" VerticalScrollBarVisibility="Auto" Margin="0,0,0,8">
            <TextBlock x:Name="UI_results" FontFamily="Consolas" FontSize="10" TextWrapping="Wrap"/>
        </ScrollViewer>

        <!-- Single Action Button -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="UI_place_btn" Content="Place Family at Origin" Width="150" Height="30" 
                    FontWeight="Bold" Margin="0,0,8,0"/>
            <Button x:Name="UI_close_btn" Content="Close" Width="80" Height="30"/>
        </StackPanel>
    </Grid>
</Window>"""

class UltraSimpleCADPlacer(forms.WPFWindow):
    def __init__(self):
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.families = {}
        self.levels = []
        
        # Simple initialization without error-prone operations
        self.simple_init()
        self.connect_events()
    
    def simple_init(self):
        """Ultra simple initialization."""
        self.log("Initializing...")
        
        # Load families - minimal error handling
        try:
            symbols = DB.FilteredElementCollector(self.doc)\
                       .OfClass(DB.FamilySymbol)\
                       .ToElements()
            
            # Just get the first few families to avoid issues
            count = 0
            for symbol in symbols:
                if count >= 10:  # Limit to first 10 families
                    break
                    
                try:
                    if symbol and symbol.Family and not symbol.Family.IsInPlace:
                        family_name = symbol.Family.Name
                        if family_name and family_name not in self.families:
                            self.families[family_name] = [symbol]
                            self.UI_family_combo.Items.Add(family_name)
                            count += 1
                except:
                    continue
            
            self.log("Loaded {} families".format(len(self.families)))
            
        except Exception as e:
            self.log("Error loading families: {}".format(str(e)))
        
        # Load levels - simple
        try:
            levels = DB.FilteredElementCollector(self.doc)\
                      .OfClass(DB.Level)\
                      .WhereElementIsNotElementType()\
                      .ToElements()
            
            for level in levels:
                self.UI_level_combo.Items.Add(level.Name)
                self.levels.append(level)
            
            if levels:
                self.UI_level_combo.SelectedIndex = 0
                
            self.log("Loaded {} levels".format(len(levels)))
            
        except Exception as e:
            self.log("Error loading levels: {}".format(str(e)))
    
    def connect_events(self):
        """Connect minimal events."""
        self.UI_family_combo.SelectionChanged += self.on_family_changed
        self.UI_place_btn.Click += self.place_at_origin
        self.UI_close_btn.Click += self.close_window
    
    def on_family_changed(self, sender, args):
        """Handle family change - simplified."""
        try:
            self.UI_type_combo.Items.Clear()
            
            if self.UI_family_combo.SelectedIndex >= 0:
                family_name = str(self.UI_family_combo.SelectedItem)
                
                if family_name in self.families:
                    symbols = self.families[family_name]
                    
                    # Just use "Type 1" for simplicity
                    self.UI_type_combo.Items.Add("Type 1")
                    self.UI_type_combo.SelectedIndex = 0
                    
                    self.log("Selected family: {}".format(family_name))
                    
        except Exception as e:
            self.log("Error selecting family: {}".format(str(e)))
    
    def place_at_origin(self, sender, args):
        """Place a single family at origin - most basic placement possible."""
        self.log("\n=== PLACING FAMILY AT ORIGIN ===")
        
        # Basic validation
        if self.UI_family_combo.SelectedIndex < 0:
            self.log("✗ No family selected")
            return
            
        if self.UI_level_combo.SelectedIndex < 0:
            self.log("✗ No level selected")  
            return
        
        try:
            # Get family symbol (first one only)
            family_name = str(self.UI_family_combo.SelectedItem)
            symbol = self.families[family_name][0]  # Just use first symbol
            
            # Get level
            level = self.levels[self.UI_level_combo.SelectedIndex]
            
            self.log("Using family: {}".format(family_name))
            self.log("Using level: {}".format(level.Name))
            
            # Activate if needed
            if not symbol.IsActive:
                symbol.Activate()
                self.doc.Regenerate()
                self.log("Activated family symbol")
            
            # Place at origin
            origin = DB.XYZ(0, 0, 0)
            
            # Transaction
            with DB.Transaction(self.doc, "Place Family at Origin") as t:
                t.Start()
                
                instance = self.doc.Create.NewFamilyInstance(
                    origin,
                    symbol,
                    level,
                    DB.Structure.StructuralType.NonStructural
                )
                
                t.Commit()
                
                self.log("✓ SUCCESS: Placed family at origin")
                self.log("Instance ID: {}".format(instance.Id))
                
                forms.alert("Family placed successfully at origin!", title="Success")
        
        except Exception as e:
            self.log("✗ Error placing family: {}".format(str(e)))
            import traceback
            self.log("Details: {}".format(traceback.format_exc()))
    
    def log(self, message):
        """Simple logging."""
        current = self.UI_results.Text
        self.UI_results.Text = current + ("\n" if current else "") + message
    
    def close_window(self, sender, args):
        """Close window."""
        self.Close()

# Run
if __name__ == '__main__':
    if not revit.doc:
        forms.alert("Open a Revit document first.", title="No Document")
    else:
        UltraSimpleCADPlacer().ShowDialog()