# -*- coding: utf-8 -*-
"""
DRAWING PACKAGE AUTOMATION TOOL
===============================

Author: Anirudh Sood
Compatible: Revit 2023-2025, Python 2.7
Enhanced with consolidated WPF UI and CSV-driven configuration
"""

__title__ = "Drawing Package\nAutomation"
__author__ = "Anirudh Sood"

# ================================================================================================
# IMPORTS
# ================================================================================================
from pyrevit import revit, DB, forms, script
from pyrevit.revit.db import query
from pychilizer import database
import csv
import os
import sys
import wpf
from System.Windows import Window
from System.IO import StringReader

# ================================================================================================
# XAML UI DEFINITION
# ================================================================================================

xaml_file = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Drawing Package Automation" Width="550" Height="580" 
    ShowInTaskbar="False" WindowStartupLocation="CenterScreen" 
    ResizeMode="NoResize" FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Area -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="Drawing Package Automation" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- Main Configuration Section -->
        <GroupBox Grid.Row="2" Header="Package Configuration" Padding="8" Background="White">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Batch Mode Checkbox -->
                <CheckBox Grid.Row="0" Grid.ColumnSpan="2" x:Name="UI_batch_mode_check" 
                          Content="Batch Mode: Generate all series marked 'AutoGenerate' in CSV" 
                          FontWeight="SemiBold"
                          ToolTip="Generates multiple series in one run based on CSV AutoGenerate column"/>

                <!-- Drawing Series -->
                <Label Grid.Row="2" Grid.Column="0" Content="Drawing Series:" 
                       FontWeight="Medium" VerticalAlignment="Center"/>
                <ComboBox Grid.Row="2" Grid.Column="1" x:Name="UI_series_combo" 
                          Height="24" VerticalAlignment="Center"/>

                <!-- Package Name -->
                <Label Grid.Row="4" Grid.Column="0" Content="Package Name:" 
                       FontWeight="Medium" VerticalAlignment="Center"/>
                <TextBox Grid.Row="4" Grid.Column="1" x:Name="UI_package_text" 
                         Height="24" VerticalAlignment="Center"
                         ToolTip="Enter package name (e.g., PARTITIONS)"/>
            </Grid>
        </GroupBox>

        <!-- Legend Options Section -->
        <GroupBox Grid.Row="4" Header="Legend Options (Optional)" Padding="8" Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Add Legend Checkbox -->
                <CheckBox Grid.Row="0" x:Name="UI_add_legend_check" 
                          Content="Add legend to all sheets" 
                          VerticalAlignment="Center"
                          FontWeight="Medium"/>

                <!-- Legend Selection -->
                <Grid Grid.Row="2" x:Name="UI_legend_grid">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="120"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <Label Grid.Column="0" Content="Select Legend:" 
                           FontWeight="Medium" VerticalAlignment="Center"/>
                    <ComboBox Grid.Column="1" x:Name="UI_legend_combo" 
                              Height="24" VerticalAlignment="Center"
                              IsEnabled="False"/>
                </Grid>
            </Grid>
        </GroupBox>

        <!-- Info Panel - Dynamic based on series selection -->
        <GroupBox Grid.Row="5" Header="Configuration Guide" Padding="8" 
                  Background="#F8F8F8" Margin="0,8,0,0">
            <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="120">
                <TextBlock x:Name="UI_info_text" TextWrapping="Wrap" 
                           FontSize="11" Foreground="#333333"/>
            </ScrollViewer>
        </GroupBox>

        <!-- Button Area -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="UI_generate_btn" Content="Generate Package" 
                    Width="130" Height="28" Margin="0,0,8,0"
                    FontWeight="SemiBold"/>
            <Button x:Name="UI_cancel_btn" Content="Cancel" 
                    Width="90" Height="28"/>
        </StackPanel>
    </Grid>
</Window>
"""

# ================================================================================================
# CONFIGURATION
# ================================================================================================

class PackageConfig:
    """Configuration for drawing package automation"""
    
    # Path to CSV configuration file - relative to script location
    @staticmethod
    def get_csv_path():
        """Get CSV path relative to script location"""
        script_dir = os.path.dirname(__file__)
        csv_filename = "drawing_series_config_enhanced.csv"
        return os.path.join(script_dir, csv_filename)
    
    # Titleblock
    TITLEBLOCK_FAMILY_NAME = "AUK_Titleblock_WorkingA1"
    
    # View placement (in feet) - coordinates from bottom-left of sheet
    VIEW_CENTER_X = 1.5  # 1.5 feet from left edge
    VIEW_CENTER_Y = 1.0  # 1.0 feet from bottom edge
    
    # Legend placement - uses auto-centering (0,0,0)
    # Legends are placed at sheet center, then manually adjusted in Revit if needed
    
    # Element types to tag (fallback if CSV not available)
    ELEMENTS_TO_TAG = {
        DB.BuiltInCategory.OST_Walls: ["AUK_Wall_Partition_"],
    }
    
    # Tag settings
    TAG_OFFSET = 0.2  # Offset from element center in feet
    TAG_LEADER = True


# ================================================================================================
# WPF WINDOW CLASS
# ================================================================================================

class DrawingPackageWindow(Window):
    """Main UI window for Drawing Package Automation"""
    
    def __init__(self, doc, config):
        # Load XAML
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        self.doc = doc
        self.config = config
        
        # Results
        self.selected_series = None
        self.series_config = None
        self.package_name = None
        self.legend_view = None
        self.create_legend = False
        
        # Initialize UI
        self.initialize()
        self.connect_events()
    
    def initialize(self):
        """Initialize UI elements"""
        
        # Load CSV configuration
        csv_path = self.config.get_csv_path()
        self.csv_config = load_series_config_from_csv(csv_path)
        
        # Populate series dropdown
        if self.csv_config:
            # Sort by series number (as integers for proper ordering)
            sorted_series = sorted(self.csv_config.items(), key=lambda x: int(x[0]))
            for series_num, series_data in sorted_series:
                # Format with leading zero for single digits
                display_num = series_num.zfill(2)
                display_text = "{} - {}".format(display_num, series_data['name'])
                self.UI_series_combo.Items.Add(display_text)
            
            if self.UI_series_combo.Items.Count > 0:
                self.UI_series_combo.SelectedIndex = 0
        else:
            # Fallback to hardcoded list with leading zeros
            hardcoded_series = [
                "00 - Proposed GA Layouts",
                "01 - Existing Layouts",
                "02 - Demolition Layouts",
                "03 - Site Layouts / Location Plans / Boundaries",
                "06 - Areas",
                "08 - Life Safety / Building Regulations / Fire Strategy",
                "20 - Existing External Facade",
                "22 - Internal Walls, Partitions",
                "23 - Floors, including beams",
                "24 - Cores",
                "25 - Stair and Ramps",
                "26 - Lift Lobby",
                "27 - Roofs",
                "30 - Internal Walls - Blockwork / Brickwork",
                "31 - External Windows, Doors, Louvres",
                "32 - Internal Doors",
                "33 - Raised Floors",
                "35 - Ceilings",
                "70 - Block Layouts / Stack Plans",
                "71 - Circulation fittings, signs, etc",
                "72 - Space planning / Furniture Layouts",
                "73 - Kitchens",
                "74 - Bathroom, WCs",
            ]
            
            for series in hardcoded_series:
                self.UI_series_combo.Items.Add(series)
            
            if self.UI_series_combo.Items.Count > 0:
                self.UI_series_combo.SelectedIndex = 0

        
        # Populate legend dropdown
        all_legends = get_all_legend_views(self.doc)
        if all_legends:
            legend_names = sorted([legend.Name for legend in all_legends])
            for name in legend_names:
                self.UI_legend_combo.Items.Add(name)
            
            if self.UI_legend_combo.Items.Count > 0:
                self.UI_legend_combo.SelectedIndex = 0
        else:
            self.UI_add_legend_check.IsEnabled = False
            self.UI_add_legend_check.Content = "Add legend to all sheets (No legends available)"
        
        # Set default package name from selected series
        self.update_package_name_from_series()
        
        # Update info panel with initial series info
        self.update_info_panel()

        # Connect batch mode handler
        self.UI_batch_mode_check.Checked += self.on_batch_mode_changed
        self.UI_batch_mode_check.Unchecked += self.on_batch_mode_changed
    
    def on_batch_mode_changed(self, sender, args):
        """Enable/disable series selection in batch mode"""
        is_batch = self.UI_batch_mode_check.IsChecked == True
        self.UI_series_combo.IsEnabled = not is_batch
        self.UI_package_text.IsEnabled = not is_batch
        
        if is_batch:
            self.UI_package_text.Text = "[AUTO - FROM CSV]"
        else:
            self.update_package_name_from_series()
    
    def connect_events(self):
        """Connect UI event handlers"""
        self.UI_generate_btn.Click += self.on_generate
        self.UI_cancel_btn.Click += self.on_cancel
        self.UI_add_legend_check.Checked += self.on_legend_check_changed
        self.UI_add_legend_check.Unchecked += self.on_legend_check_changed
        self.UI_series_combo.SelectionChanged += self.on_series_changed
    
    def on_series_changed(self, sender, args):
        """Update package name and info panel when series changes"""
        self.update_package_name_from_series()
        self.update_info_panel()
    
    def update_info_panel(self):
        """Update info panel based on selected series configuration with color formatting"""
        
        try:
            # Import WPF controls for formatting
            from System.Windows.Documents import Run
            from System.Windows.Media import Brushes, FontFamily
            from System.Windows import FontWeights, FontStyles
            
            # Clear existing content
            self.UI_info_text.Inlines.Clear()
            
            if not self.UI_series_combo.SelectedItem:
                run = Run("Select a drawing series to see tagging configuration details.")
                run.Foreground = Brushes.Gray
                self.UI_info_text.Inlines.Add(run)
                return
            
            # Get series number
            selected_text = str(self.UI_series_combo.SelectedItem)
            if " - " in selected_text:
                series_num = selected_text.split(" - ")[0].lstrip('0') or '0'  # Remove leading zeros for lookup
            else:
                series_num = selected_text.split()[0].lstrip('0') or '0'
            
            # Get series config
            if self.csv_config and series_num in self.csv_config:
                series_cfg = self.csv_config[series_num]
                
                # Check if custom guidance exists in CSV
                custom_guidance = series_cfg.get('user_guidance', '')
                if custom_guidance:
                    # Parse and format custom guidance
                    lines = custom_guidance.split('.')
                    for i, line in enumerate(lines):
                        line = line.strip()
                        if not line:
                            continue
                        
                        # First sentence is title (bold, blue)
                        if i == 0:
                            title_run = Run(line + ".\n\n")
                            title_run.FontWeight = FontWeights.Bold
                            title_run.Foreground = Brushes.DarkBlue
                            title_run.FontSize = 12
                            self.UI_info_text.Inlines.Add(title_run)
                        else:
                            # Other sentences (normal, dark gray)
                            content_run = Run(line + ". ")
                            content_run.Foreground = Brushes.DarkSlateGray
                            self.UI_info_text.Inlines.Add(content_run)
                    return
                
                # Generate default guidance with color formatting
                
                # Title
                title = Run("TAGGING CONFIGURATION\n")
                title.FontWeight = FontWeights.Bold
                title.Foreground = Brushes.DarkBlue
                title.FontSize = 12
                self.UI_info_text.Inlines.Add(title)
                
                self.UI_info_text.Inlines.Add(Run("\n"))
                
                # View type info
                view_type = series_cfg.get('view_type', 'FloorPlan')
                view_label = Run("View Type: ")
                view_label.FontWeight = FontWeights.SemiBold
                view_label.Foreground = Brushes.DarkSlateGray
                self.UI_info_text.Inlines.Add(view_label)
                
                if view_type == 'CeilingPlan':
                    view_value = Run("Reflected Ceiling Plans (RCP)\n")
                elif view_type == 'AreaPlan':
                    view_value = Run("Area Plans\n")
                else:
                    view_value = Run("Floor Plans\n")
                view_value.Foreground = Brushes.DarkGreen
                self.UI_info_text.Inlines.Add(view_value)
                
                self.UI_info_text.Inlines.Add(Run("\n"))
                
                # What will be tagged - header
                tag_header = Run("Elements to be tagged:\n")
                tag_header.FontWeight = FontWeights.SemiBold
                tag_header.Foreground = Brushes.DarkSlateGray
                self.UI_info_text.Inlines.Add(tag_header)
                
                has_tags = False
                
                # Rooms
                if series_cfg.get('tag_rooms'):
                    has_tags = True
                    bullet = Run("  • ")
                    bullet.Foreground = Brushes.Gray
                    self.UI_info_text.Inlines.Add(bullet)
                    
                    room_text = Run("Rooms: ")
                    room_text.FontWeight = FontWeights.SemiBold
                    room_text.Foreground = Brushes.DarkOrange
                    self.UI_info_text.Inlines.Add(room_text)
                    
                    room_detail = Run("All placed rooms\n")
                    room_detail.Foreground = Brushes.Black
                    self.UI_info_text.Inlines.Add(room_detail)
                
                # Walls
                if series_cfg.get('tag_walls'):
                    has_tags = True
                    wall_filter = series_cfg.get('wall_filter', '')
                    
                    bullet = Run("  • ")
                    bullet.Foreground = Brushes.Gray
                    self.UI_info_text.Inlines.Add(bullet)
                    
                    wall_text = Run("Walls: ")
                    wall_text.FontWeight = FontWeights.SemiBold
                    wall_text.Foreground = Brushes.DarkOrange
                    self.UI_info_text.Inlines.Add(wall_text)
                    
                    if wall_filter:
                        wall_detail = Run("Only types containing ")
                        wall_detail.Foreground = Brushes.Black
                        self.UI_info_text.Inlines.Add(wall_detail)
                        
                        filter_text = Run("'{}'".format(wall_filter))
                        filter_text.FontFamily = FontFamily("Consolas")
                        filter_text.Foreground = Brushes.DarkBlue
                        self.UI_info_text.Inlines.Add(filter_text)
                        
                        self.UI_info_text.Inlines.Add(Run("\n"))
                        
                        # Example - will tag
                        ex1_indent = Run("      ")
                        self.UI_info_text.Inlines.Add(ex1_indent)
                        ex1 = Run("'AUK_Wall_Partition_P2_122mm' ")
                        ex1.FontFamily = FontFamily("Consolas")
                        ex1.FontSize = 10
                        ex1.Foreground = Brushes.DarkGreen
                        self.UI_info_text.Inlines.Add(ex1)
                        ex1_label = Run("[TAGGED]\n")
                        ex1_label.FontSize = 10
                        ex1_label.Foreground = Brushes.DarkGreen
                        ex1_label.FontWeight = FontWeights.Bold
                        self.UI_info_text.Inlines.Add(ex1_label)
                        
                        # Example - will skip
                        ex2_indent = Run("      ")
                        self.UI_info_text.Inlines.Add(ex2_indent)
                        ex2 = Run("'AUK_Wall_External_Brick' ")
                        ex2.FontFamily = FontFamily("Consolas")
                        ex2.FontSize = 10
                        ex2.Foreground = Brushes.DarkRed
                        self.UI_info_text.Inlines.Add(ex2)
                        ex2_label = Run("[SKIPPED]\n")
                        ex2_label.FontSize = 10
                        ex2_label.Foreground = Brushes.DarkRed
                        ex2_label.FontWeight = FontWeights.Bold
                        self.UI_info_text.Inlines.Add(ex2_label)
                    else:
                        wall_detail = Run("All wall types\n")
                        wall_detail.Foreground = Brushes.Black
                        self.UI_info_text.Inlines.Add(wall_detail)
                
                # Doors
                if series_cfg.get('tag_doors'):
                    has_tags = True
                    door_filter = series_cfg.get('door_filter', '')
                    
                    bullet = Run("  • ")
                    bullet.Foreground = Brushes.Gray
                    self.UI_info_text.Inlines.Add(bullet)
                    
                    door_text = Run("Doors: ")
                    door_text.FontWeight = FontWeights.SemiBold
                    door_text.Foreground = Brushes.DarkOrange
                    self.UI_info_text.Inlines.Add(door_text)
                    
                    if door_filter:
                        door_detail = Run("Only families containing ")
                        door_detail.Foreground = Brushes.Black
                        self.UI_info_text.Inlines.Add(door_detail)
                        
                        filter_text = Run("'{}'".format(door_filter))
                        filter_text.FontFamily = FontFamily("Consolas")
                        filter_text.Foreground = Brushes.DarkBlue
                        self.UI_info_text.Inlines.Add(filter_text)
                        self.UI_info_text.Inlines.Add(Run("\n"))
                    else:
                        door_detail = Run("All door types\n")
                        door_detail.Foreground = Brushes.Black
                        self.UI_info_text.Inlines.Add(door_detail)
                
                # Windows
                if series_cfg.get('tag_windows'):
                    has_tags = True
                    window_filter = series_cfg.get('window_filter', '')
                    
                    bullet = Run("  • ")
                    bullet.Foreground = Brushes.Gray
                    self.UI_info_text.Inlines.Add(bullet)
                    
                    window_text = Run("Windows: ")
                    window_text.FontWeight = FontWeights.SemiBold
                    window_text.Foreground = Brushes.DarkOrange
                    self.UI_info_text.Inlines.Add(window_text)
                    
                    if window_filter:
                        window_detail = Run("Only families containing ")
                        window_detail.Foreground = Brushes.Black
                        self.UI_info_text.Inlines.Add(window_detail)
                        
                        filter_text = Run("'{}'".format(window_filter))
                        filter_text.FontFamily = FontFamily("Consolas")
                        filter_text.Foreground = Brushes.DarkBlue
                        self.UI_info_text.Inlines.Add(filter_text)
                        self.UI_info_text.Inlines.Add(Run("\n"))
                    else:
                        window_detail = Run("All window types\n")
                        window_detail.Foreground = Brushes.Black
                        self.UI_info_text.Inlines.Add(window_detail)
                
                # Ceilings
                if series_cfg.get('tag_ceilings'):
                    has_tags = True
                    bullet = Run("  • ")
                    bullet.Foreground = Brushes.Gray
                    self.UI_info_text.Inlines.Add(bullet)
                    
                    ceiling_text = Run("Ceilings: ")
                    ceiling_text.FontWeight = FontWeights.SemiBold
                    ceiling_text.Foreground = Brushes.DarkOrange
                    self.UI_info_text.Inlines.Add(ceiling_text)
                    
                    ceiling_detail = Run("All ceiling types\n")
                    ceiling_detail.Foreground = Brushes.Black
                    self.UI_info_text.Inlines.Add(ceiling_detail)
                
                if not has_tags:
                    no_tag = Run("No automatic tagging configured for this series.\n")
                    no_tag.Foreground = Brushes.Gray
                    no_tag.FontStyle = FontStyles.Italic
                    self.UI_info_text.Inlines.Add(no_tag)
                
                # Footer note
                self.UI_info_text.Inlines.Add(Run("\n"))
                note = Run("Note: Only elements matching the filters will be tagged.\nDetailed results shown in output window during generation.")
                note.Foreground = Brushes.Gray
                note.FontSize = 10
                note.FontStyle = FontStyles.Italic
                self.UI_info_text.Inlines.Add(note)
                
            else:
                # No CSV config available
                fallback = Run("Using default configuration. Tagging behavior controlled by PackageConfig class.")
                fallback.Foreground = Brushes.Gray
                fallback.FontStyle = FontStyles.Italic
                self.UI_info_text.Inlines.Add(fallback)
        
        except Exception as e:
            # Log error but don't crash
            logger.error("Error updating info panel: {}".format(str(e)))
            self.UI_info_text.Inlines.Clear()
            error_run = Run("Error loading configuration details. Check output window for details.")
            error_run.Foreground = Brushes.Red
            self.UI_info_text.Inlines.Add(error_run)
    
    def update_package_name_from_series(self):
        """Auto-populate package name from selected series"""
        if self.UI_series_combo.SelectedItem:
            selected_text = str(self.UI_series_combo.SelectedItem)
            # Extract name after " - "
            if " - " in selected_text:
                package_name = selected_text.split(" - ", 1)[1].upper()
                self.UI_package_text.Text = package_name
    
    def on_legend_check_changed(self, sender, args):
        """Enable/disable legend selection"""
        is_checked = self.UI_add_legend_check.IsChecked == True
        self.UI_legend_combo.IsEnabled = is_checked
    
    def on_generate(self, sender, args):
        """Generate button clicked"""
        
        # Validate inputs
        if not self.validate_inputs():
            return
        
        # Check if batch mode
        if self.UI_batch_mode_check.IsChecked == True:
            self.setup_batch_mode()
        else:
            self.setup_single_mode()
        
        # Close dialog with OK result
        self.DialogResult = True
        self.Close()
    
    def setup_batch_mode(self):
        """Setup for batch generation"""
        self.batch_mode = True
        self.batch_series = []
        
        # Collect all auto-generate series
        for series_num, series_cfg in self.csv_config.items():
            if series_cfg.get('auto_generate', False):
                self.batch_series.append({
                    'series_num': series_num,
                    'config': series_cfg,
                    'package_name': series_cfg['name'].upper()
                })
        
        # Sort by series number
        self.batch_series.sort(key=lambda x: int(x['series_num']))
        
        # Get legend if enabled
        self.create_legend = self.UI_add_legend_check.IsChecked == True
        if self.create_legend and self.UI_legend_combo.SelectedItem:
            legend_name = str(self.UI_legend_combo.SelectedItem)
            all_legends = get_all_legend_views(self.doc)
            for legend in all_legends:
                if legend.Name == legend_name:
                    self.legend_view = legend
                    break
    
    def setup_single_mode(self):
        """Setup for single series generation"""
        self.batch_mode = False
        
        # Extract selected series number
        selected_text = str(self.UI_series_combo.SelectedItem)
        if " - " in selected_text:
            series_display = selected_text.split(" - ")[0]
        else:
            series_display = selected_text.split()[0]
        
        self.selected_series = series_display.lstrip('0') or '0'
        
        # Get series config
        if self.csv_config and self.selected_series in self.csv_config:
            self.series_config = self.csv_config[self.selected_series]
        else:
            if self.csv_config and series_display in self.csv_config:
                self.selected_series = series_display
                self.series_config = self.csv_config[series_display]
        
        # Get package name
        self.package_name = self.UI_package_text.Text.strip()
        
        # Get legend if enabled
        self.create_legend = self.UI_add_legend_check.IsChecked == True
        if self.create_legend and self.UI_legend_combo.SelectedItem:
            legend_name = str(self.UI_legend_combo.SelectedItem)
            all_legends = get_all_legend_views(self.doc)
            for legend in all_legends:
                if legend.Name == legend_name:
                    self.legend_view = legend
                    break
    
    def on_cancel(self, sender, args):
        """Cancel button clicked"""
        self.DialogResult = False
        self.Close()
    
    def validate_inputs(self):
        """Validate user inputs"""
        
        # Check if batch mode
        if self.UI_batch_mode_check.IsChecked == True:
            return self.validate_batch_mode()
        
        # Check series selection
        if not self.UI_series_combo.SelectedItem:
            forms.alert("Please select a drawing series.", title="Input Required")
            return False
        
        # Check package name
        package_name = self.UI_package_text.Text.strip()
        if not package_name:
            forms.alert("Please enter a package name.", title="Input Required")
            return False
        
        # Check legend if enabled
        if self.UI_add_legend_check.IsChecked == True:
            if not self.UI_legend_combo.SelectedItem:
                forms.alert("Please select a legend or disable the legend option.", 
                           title="Input Required")
                return False
        
        # PRE-FLIGHT VALIDATION
        return self.validate_project_setup()
    
    def validate_batch_mode(self):
        """Validate batch mode setup"""
        if not self.csv_config:
            forms.alert("CSV configuration required for batch mode.", 
                       title="Configuration Missing")
            return False
        
        # Count auto-generate series
        auto_series = [s for s, cfg in self.csv_config.items() 
                      if cfg.get('auto_generate', False)]
        
        if not auto_series:
            forms.alert("No series marked 'AutoGenerate=TRUE' in CSV.\n\n"
                       "Enable batch mode by setting AutoGenerate=TRUE for desired series.",
                       title="No Series Selected")
            return False
        
        # Show confirmation
        series_list = "\n".join([" - Series {}: {}".format(s, self.csv_config[s]['name']) 
                                for s in sorted(auto_series, key=int)])
        
        confirm = forms.alert(
            "BATCH MODE: Will generate {} series:\n\n{}\n\n"
            "Legend option will apply to all series.\n\n"
            "Continue?".format(len(auto_series), series_list),
            title="Confirm Batch Generation",
            ok=False,
            yes=True,
            no=True
        )
        
        if not confirm:
            return False
        
        # Validate each series
        return self.validate_project_setup_batch(auto_series)
    
    def validate_project_setup(self):
        """Pre-flight checks for single series"""
        
        issues = []
        
        # Get series config
        selected_text = str(self.UI_series_combo.SelectedItem)
        if " - " in selected_text:
            series_num = selected_text.split(" - ")[0].lstrip('0') or '0'
        else:
            series_num = selected_text.split()[0].lstrip('0') or '0'
        
        series_config = None
        if self.csv_config and series_num in self.csv_config:
            series_config = self.csv_config[series_num]
        
        # Check 1: View template exists
        if series_config:
            template_pattern = series_config['template_pattern']
            template = get_view_template_by_pattern(self.doc, template_pattern)
            if not template:
                issues.append("View template not found: '{}'".format(template_pattern))
        
        # Check 2: Titleblock exists
        titleblock = get_titleblock_type(self.doc, self.config.TITLEBLOCK_FAMILY_NAME)
        if not titleblock:
            issues.append("Titleblock not found: '{}'".format(
                self.config.TITLEBLOCK_FAMILY_NAME))
        
        # Check 3: Tag families loaded (if tagging enabled)
        if series_config:
            if series_config.get('tag_rooms'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_RoomTags):
                    issues.append("Room tag family not loaded")
            
            if series_config.get('tag_walls'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_WallTags):
                    issues.append("Wall tag family not loaded")
            
            if series_config.get('tag_doors'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_DoorTags):
                    issues.append("Door tag family not loaded")
            
            if series_config.get('tag_windows'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_WindowTags):
                    issues.append("Window tag family not loaded")
            
            if series_config.get('tag_ceilings'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_CeilingTags):
                    issues.append("Ceiling tag family not loaded")
        
        # Check 4: Levels exist
        levels = get_all_levels(self.doc)
        if not levels:
            issues.append("No levels found in project")
        
        # Check 5: Area scheme exists (if area plan)
        if series_config and series_config.get('view_type') == 'AreaPlan':
            area_schemes = DB.FilteredElementCollector(self.doc)\
                .OfClass(DB.AreaScheme).ToElements()
            has_gia_gea = False
            for scheme in area_schemes:
                name = scheme.Name.upper()
                if 'GIA' in name or 'GEA' in name:
                    has_gia_gea = True
                    break
            if not has_gia_gea:
                issues.append("No GIA/GEA area scheme found (required for area plans)")
        
        # Report issues
        if issues:
            message = "PRE-FLIGHT VALIDATION FAILED:\n\n" + "\n".join(
                [" • {}".format(issue) for issue in issues]
            )
            message += "\n\nFix these issues before generating package."
            forms.alert(message, title="Validation Failed")
            return False
        
        return True
    
    def validate_project_setup_batch(self, auto_series):
        """Pre-flight checks for batch mode"""
        
        all_issues = {}
        
        for series_num in auto_series:
            series_config = self.csv_config[series_num]
            issues = []
            
            # Check 1: View template
            template_pattern = series_config['template_pattern']
            template = get_view_template_by_pattern(self.doc, template_pattern)
            if not template:
                issues.append("Template not found: '{}'".format(template_pattern))
            
            # Check 2: Tag families
            if series_config.get('tag_rooms'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_RoomTags):
                    issues.append("Room tags not loaded")
            
            if series_config.get('tag_walls'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_WallTags):
                    issues.append("Wall tags not loaded")
            
            if series_config.get('tag_doors'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_DoorTags):
                    issues.append("Door tags not loaded")
            
            if series_config.get('tag_windows'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_WindowTags):
                    issues.append("Window tags not loaded")
            
            if series_config.get('tag_ceilings'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_CeilingTags):
                    issues.append("Ceiling tags not loaded")
            
            # Check 3: Area scheme (if needed)
            if series_config.get('view_type') == 'AreaPlan':
                area_schemes = DB.FilteredElementCollector(self.doc)\
                    .OfClass(DB.AreaScheme).ToElements()
                has_gia_gea = False
                for scheme in area_schemes:
                    name = scheme.Name.upper()
                    if 'GIA' in name or 'GEA' in name:
                        has_gia_gea = True
                        break
                if not has_gia_gea:
                    issues.append("No GIA/GEA area scheme")
            
            if issues:
                all_issues[series_num] = issues
        
        # Global checks
        global_issues = []
        
        titleblock = get_titleblock_type(self.doc, self.config.TITLEBLOCK_FAMILY_NAME)
        if not titleblock:
            global_issues.append("Titleblock not found: '{}'".format(
                self.config.TITLEBLOCK_FAMILY_NAME))
        
        levels = get_all_levels(self.doc)
        if not levels:
            global_issues.append("No levels found in project")
        
        # Report
        if global_issues or all_issues:
            message = "PRE-FLIGHT VALIDATION FAILED:\n\n"
            
            if global_issues:
                message += "GLOBAL ISSUES:\n"
                for issue in global_issues:
                    message += " • {}\n".format(issue)
                message += "\n"
            
            if all_issues:
                message += "SERIES-SPECIFIC ISSUES:\n"
                for series_num, issues in all_issues.items():
                    series_name = self.csv_config[series_num]['name']
                    message += "Series {} ({}):\n".format(series_num, series_name)
                    for issue in issues:
                        message += "  • {}\n".format(issue)
                message += "\n"
            
            message += "Fix these issues before batch generation."
            forms.alert(message, title="Batch Validation Failed")
            return False
        
        return True
    
    def check_tag_family_loaded(self, category):
        """Check if tag family is loaded for category"""
        try:
            collector = DB.FilteredElementCollector(self.doc)
            tags = collector.OfCategory(category).WhereElementIsElementType().ToElements()
            return len(list(tags)) > 0
        except:
            return False


# ================================================================================================
# UTILITY FUNCTIONS
# ================================================================================================

logger = script.get_logger()


def load_series_config_from_csv(csv_path):
    """Load series configuration from CSV file
    
    Returns dict of series configurations, or None if file not found
    """
    if not os.path.exists(csv_path):
        logger.warning("CSV config not found: {}".format(csv_path))
        return None
    
    config = {}
    try:
        with open(csv_path, 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    series_num = row.get('SeriesNumber', '').strip()
                    if not series_num:
                        continue  # Skip rows without series number
                    
                    # Helper function to safely get and strip values
                    def safe_get(key, default=''):
                        value = row.get(key, default)
                        if value is None:
                            return default
                        return str(value).strip()
                    
                    # Helper function for boolean values
                    def safe_bool(key):
                        value = safe_get(key, 'FALSE')
                        return value.upper() == 'TRUE'
                    
                    config[series_num] = {
                    'name': safe_get('SeriesName'),
                    'view_type': safe_get('ViewType', 'FloorPlan'),
                    'template_pattern': safe_get('ViewTemplate'),
                    'tag_rooms': safe_bool('TagRooms'),
                    'tag_doors': safe_bool('TagDoors'),
                    'tag_windows': safe_bool('TagWindows'),
                    'tag_walls': safe_bool('TagWalls'),
                    'tag_ceilings': safe_bool('TagCeilings'),
                    'create_rcp': safe_bool('CreateRCP'),
                    'level_numbering': safe_bool('LevelNumbering'),  # ADD THIS LINE
                    'auto_generate': safe_bool('AutoGenerate'),  # ADD THIS
                    'wall_filter': safe_get('WallFilter'),
                    'door_filter': safe_get('DoorFilter'),
                    'window_filter': safe_get('WindowFilter'),
                    'user_guidance': safe_get('UserGuidance'),
                    }
                except Exception as row_error:
                    logger.warning("Error processing CSV row for series {}: {}".format(
                        row.get('SeriesNumber', 'unknown'), str(row_error)))
                    continue
        
        return config
    except Exception as e:
        logger.error("Failed to load CSV config: {}".format(str(e)))
        return None


def get_all_levels(doc):
    """Get all levels sorted by elevation"""
    collector = DB.FilteredElementCollector(doc)
    levels = collector.OfClass(DB.Level).WhereElementIsNotElementType().ToElements()
    sorted_levels = sorted(levels, key=lambda l: l.Elevation)
    return sorted_levels


def get_view_template_by_pattern(doc, pattern):
    """Find view template matching pattern"""
    collector = DB.FilteredElementCollector(doc)
    view_templates = collector.OfClass(DB.View).ToElements()
    
    for vt in view_templates:
        if vt.IsTemplate and pattern in vt.Name:
            return vt
    
    return None


def get_next_sheet_number(doc, series_number):
    """Calculate next available sheet number in series"""
    
    series_prefix = series_number.zfill(2)
    start_number_str = "{}100".format(series_prefix)
    start_number = int(start_number_str)
    
    range_start = start_number
    range_end = int("{}200".format(series_prefix))
    
    collector = DB.FilteredElementCollector(doc)
    sheets = collector.OfCategory(DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    
    highest = start_number - 1
    
    for sheet in sheets:
        sheet_num = sheet.SheetNumber
        
        if sheet_num.startswith(series_prefix):
            try:
                num = int(sheet_num)
                if num >= range_start and num < range_end and num > highest:
                    highest = num
            except:
                continue
    
    next_num = highest + 1
    return str(next_num).zfill(5)


def get_titleblock_type(doc, family_name):
    """Get titleblock type by family name"""
    collector = DB.FilteredElementCollector(doc)
    titleblocks = collector.OfCategory(DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
    
    for tb in titleblocks:
        try:
            family_param = tb.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            if family_param and family_name in family_param.AsString():
                return tb
        except:
            continue
    
    return None


def get_element_center_point(element):
    """Get approximate center point of an element"""
    try:
        bbox = element.get_BoundingBox(None)
        if bbox:
            center = (bbox.Min + bbox.Max) / 2.0
            return center
        
        if hasattr(element, 'Location') and isinstance(element.Location, DB.LocationCurve):
            curve = element.Location.Curve
            midpoint = curve.Evaluate(0.5, True)
            return midpoint
    except Exception as e:
        logger.debug("Error getting element center: {}".format(str(e)))
    
    return None


def get_all_legend_views(doc):
    """Get all non-template legend views"""
    all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    legends = [v for v in all_views if v.ViewType == DB.ViewType.Legend and not v.IsTemplate]
    return legends


# ================================================================================================
# MAIN AUTOMATION CLASS
# ================================================================================================

class DrawingPackageAutomation:
    """Main automation orchestrator"""
    
    def __init__(self, doc, config, selected_series, series_config, package_name, 
                 legend_view, create_legend):
        self.doc = doc
        self.config = config
        self.output = script.get_output()
        
        self.selected_series = selected_series
        self.series_config = series_config
        self.package_name = package_name
        self.legend_view = legend_view
        self.create_legend = create_legend
        
        self.created_views = []
        self.created_sheets = []
        self.total_tagged = 0
    
    def run(self):
        """Execute the complete automation workflow"""
        
        self.output.print_md("# DRAWING PACKAGE AUTOMATION")
        self.output.print_md("---")
        self.output.print_md("**Series:** {}".format(self.selected_series))
        self.output.print_md("**Package:** {}".format(self.package_name))
        self.output.print_md("")
        
        try:
            # TRANSACTION 1: Create views, sheets, place everything
            with revit.Transaction("Drawing Package - Create All - {}".format(self.selected_series)):
                if not self.step2_create_views():
                    return False
                
                if not self.step3_rename_views():
                    return False
                
                if not self.step4_apply_templates():
                    return False
                
                if not self.step6_create_sheets():
                    return False
                
                if not self.step7_place_views():
                    return False
            
            self.output.print_md("**Views and sheets created - document committed**")
            self.output.print_md("")
            
            # TRANSACTION 2: Tag elements
            with revit.Transaction("Drawing Package - Tag Elements - {}".format(self.selected_series)):
                if not self.step5_tag_elements():
                    return False
            
            # TRANSACTION 3: Place legend if selected
            if self.create_legend and self.legend_view:
                with revit.Transaction("Drawing Package - Legend - {}".format(self.selected_series)):
                    if not self.step9_place_legend():
                        return False
            
            self.print_summary()
            return True
            
        except Exception as e:
            self.output.print_md("## ERROR: {}".format(str(e)))
            import traceback
            self.output.print_html("<pre>{}</pre>".format(traceback.format_exc()))
            return False
    
    def step2_create_views(self):
        """Step 2: Create views"""
        
        self.output.print_md("## STEP 2: View Creation")
        
        levels = get_all_levels(self.doc)
        
        if not levels:
            forms.alert("No levels found")
            return False
        
        self.output.print_md("Found {} levels".format(len(levels)))
        
        # Determine view type from config
        view_type = 'PLAN'
        use_ceiling_plan = False
        use_area_plan = False
        area_scheme_id = None
        
        if self.series_config:
            config_view_type = self.series_config['view_type']
            if config_view_type == 'CeilingPlan':
                view_type = 'RCP'
                use_ceiling_plan = True
                self.output.print_md("*Creating Reflected Ceiling Plans*")
            elif config_view_type == 'FloorPlan':
                view_type = 'PLAN'
                self.output.print_md("*Creating Floor Plans*")
            elif config_view_type == 'AreaPlan':
                view_type = 'AREA'
                use_area_plan = True
                self.output.print_md("*Creating Area Plans*")
                
                area_schemes = DB.FilteredElementCollector(self.doc)\
                    .OfClass(DB.AreaScheme)\
                    .ToElements()
                
                for scheme in area_schemes:
                    scheme_name = scheme.Name.upper()
                    if 'GIA' in scheme_name:
                        area_scheme_id = scheme.Id
                        self.output.print_md("*Using area scheme: {}*".format(scheme.Name))
                        break
                
                if not area_scheme_id:
                    for scheme in area_schemes:
                        scheme_name = scheme.Name.upper()
                        if 'GEA' in scheme_name:
                            area_scheme_id = scheme.Id
                            self.output.print_md("*Using area scheme: {}*".format(scheme.Name))
                            break
                
                if not area_scheme_id:
                    self.output.print_md("**ERROR:** No GIA or GEA area scheme found")
                    return False
        
        for level in levels:
            try:
                if use_ceiling_plan:
                    view_family_type_id = None
                    for vft in DB.FilteredElementCollector(self.doc).OfClass(DB.ViewFamilyType):
                        if vft.ViewFamily == DB.ViewFamily.CeilingPlan:
                            view_family_type_id = vft.Id
                            break
                    
                    if not view_family_type_id:
                        self.output.print_md("  - WARNING: No ceiling plan view type found")
                        continue
                    
                    new_view = DB.ViewPlan.Create(
                        self.doc,
                        view_family_type_id,
                        level.Id
                    )
                elif use_area_plan:
                    view_family_type_id = None
                    for vft in DB.FilteredElementCollector(self.doc).OfClass(DB.ViewFamilyType):
                        if vft.ViewFamily == DB.ViewFamily.AreaPlan:
                            view_family_type_id = vft.Id
                            break
                    
                    if not view_family_type_id:
                        self.output.print_md("  - WARNING: No area plan view type found")
                        continue
                    
                    new_view = DB.ViewPlan.CreateAreaPlan(
                        self.doc,
                        area_scheme_id,
                        level.Id
                    )
                else:
                    view_family_type_id = self.doc.GetDefaultElementTypeId(
                        DB.ElementTypeGroup.ViewTypeFloorPlan
                    )
                    
                    new_view = DB.ViewPlan.Create(
                        self.doc,
                        view_family_type_id,
                        level.Id
                    )
                
                self.created_views.append({
                    'view': new_view,
                    'level': level,
                    'type': view_type
                })
                
                self.output.print_md("- Created: {}".format(level.Name))
                
            except Exception as e:
                self.output.print_md("  - WARNING: {}".format(str(e)))
        
        self.output.print_md("**Created {} views**".format(len(self.created_views)))
        self.output.print_md("")
        
        return len(self.created_views) > 0
    
    def step3_rename_views(self):
        """Step 3: Rename views"""
        
        self.output.print_md("## STEP 3: View Renaming")
        
        for view_data in self.created_views:
            view = view_data['view']
            level = view_data['level']
            view_type = view_data['type']
            
            new_name = "{} - {} - {} - {}".format(
                self.selected_series,
                self.package_name,
                level.Name.upper(),
                view_type
            )
            
            try:
                view.Name = new_name
                view_data['renamed'] = new_name
                self.output.print_md("- {}".format(new_name))
            except Exception as e:
                self.output.print_md("  - WARNING: {}".format(str(e)))
        
        self.output.print_md("")
        return True
    
    def step4_apply_templates(self):
        """Step 4: Apply templates"""
        
        self.output.print_md("## STEP 4: Apply View Templates")
        
        if self.series_config:
            template_pattern = self.series_config['template_pattern']
        else:
            template_pattern = "PUBLISHED - {} - PLANS".format(self.selected_series)
        
        template = get_view_template_by_pattern(self.doc, template_pattern)
        
        if not template:
            self.output.print_md("**WARNING:** No template found for pattern: {}".format(template_pattern))
            return True
        
        self.output.print_md("Template: {}".format(template.Name))
        
        for view_data in self.created_views:
            view = view_data['view']
            try:
                view.ViewTemplateId = template.Id
            except Exception as e:
                self.output.print_md("  - WARNING: {}".format(str(e)))
        
        self.output.print_md("")
        return True
    
    def step5_tag_elements(self):
        """Step 5: Tag elements - ROBUST VERSION with view regeneration and user guidance"""
        
        self.output.print_md("## STEP 5: Tag Elements")
        
        # Force document regeneration to ensure views are populated
        self.doc.Regenerate()
        
        total_tagged = 0
        
        # Build tagging configuration based on view type
        categories_to_tag = {}
        
        if self.series_config:
            view_type = self.created_views[0]['type'] if self.created_views else 'PLAN'
            
            self.output.print_md("**Tagging Configuration for View Type: {}**".format(view_type))
            self.output.print_md("")
            
            if view_type == 'RCP':
                if self.series_config['tag_ceilings']:
                    categories_to_tag[DB.BuiltInCategory.OST_Ceilings] = []
                    self.output.print_md("- Ceilings: ENABLED (all ceiling types)")
            elif view_type == 'AREA':
                pass
            else:
                if self.series_config['tag_walls']:
                    wall_filter = self.series_config.get('wall_filter', 'AUK_Wall_Partition_')
                    if wall_filter:
                        categories_to_tag[DB.BuiltInCategory.OST_Walls] = [wall_filter]
                        self.output.print_md("- Walls: ENABLED")
                        self.output.print_md("  - Only wall types containing: **'{}'**".format(wall_filter))
                        self.output.print_md("  - Example: 'AUK_Wall_Partition_P2_122mm' WILL be tagged")
                        self.output.print_md("  - Example: 'AUK_Wall_External_Brick' will be SKIPPED")
                    else:
                        categories_to_tag[DB.BuiltInCategory.OST_Walls] = []
                        self.output.print_md("- Walls: ENABLED (all wall types)")
                
                if self.series_config['tag_doors']:
                    door_filter = self.series_config.get('door_filter', 'AUK_Door')
                    if door_filter:
                        categories_to_tag[DB.BuiltInCategory.OST_Doors] = [door_filter]
                        self.output.print_md("- Doors: ENABLED")
                        self.output.print_md("  - Only door families containing: **'{}'**".format(door_filter))
                    else:
                        categories_to_tag[DB.BuiltInCategory.OST_Doors] = []
                        self.output.print_md("- Doors: ENABLED (all door types)")
                
                if self.series_config['tag_windows']:
                    window_filter = self.series_config.get('window_filter', 'AUK_Window')
                    if window_filter:
                        categories_to_tag[DB.BuiltInCategory.OST_Windows] = [window_filter]
                        self.output.print_md("- Windows: ENABLED")
                        self.output.print_md("  - Only window families containing: **'{}'**".format(window_filter))
                    else:
                        categories_to_tag[DB.BuiltInCategory.OST_Windows] = []
                        self.output.print_md("- Windows: ENABLED (all window types)")
            
            if self.series_config['tag_rooms']:
                self.output.print_md("- Rooms: ENABLED (all placed rooms)")
        else:
            categories_to_tag = self.config.ELEMENTS_TO_TAG
            self.output.print_md("**Using fallback configuration**")
        
        self.output.print_md("")
        self.output.print_md("---")
        self.output.print_md("")
        
        for view_data in self.created_views:
            view = view_data['view']
            view_type = view_data['type']
            view_tagged = 0
            view_name = view_data.get('renamed', view.Name)
            
            # Tag rooms if configured
            if self.series_config and self.series_config['tag_rooms']:
                all_rooms = DB.FilteredElementCollector(self.doc)\
                    .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                
                level_id = view_data['level'].Id
                rooms = []
                for room in all_rooms:
                    try:
                        if room.LevelId == level_id and room.Area > 0:
                            rooms.append(room)
                    except:
                        pass
                
                self.output.print_md("  - {} - Found {} rooms".format(view_name, len(rooms)))
                
                for room in rooms:
                    try:
                        location = room.Location
                        if location:
                            point = location.Point
                            uv = DB.UV(point.X, point.Y)
                            
                            room_tag = revit.doc.Create.NewRoomTag(
                                DB.LinkElementId(room.Id),
                                uv,
                                view.Id
                            )
                            
                            if room_tag:
                                view_tagged += 1
                                total_tagged += 1
                    except Exception as e:
                        logger.debug("Failed to tag room {}: {}".format(room.Id, str(e)))
                        pass
            
            # Tag other categories
            for category, type_patterns in categories_to_tag.items():
                
                # CRITICAL: Collect from document, filter by level
                collector = DB.FilteredElementCollector(self.doc)
                all_elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()
                
                level_id = view_data['level'].Id
                elements = []
                for elem in all_elements:
                    try:
                        elem_level_id = elem.LevelId
                        if elem_level_id == level_id:
                            elements.append(elem)
                    except:
                        pass
                
                category_name = "Unknown"
                if category == DB.BuiltInCategory.OST_Walls:
                    category_name = "Walls"
                elif category == DB.BuiltInCategory.OST_Doors:
                    category_name = "Doors"
                elif category == DB.BuiltInCategory.OST_Windows:
                    category_name = "Windows"
                elif category == DB.BuiltInCategory.OST_Ceilings:
                    category_name = "Ceilings"
                
                self.output.print_md("  - {} - Checking {} {} on this level".format(
                    view_name, len(elements), category_name
                ))
                
                tagged_in_category = 0
                skipped_no_match = 0
                failed_to_tag = 0
                
                for element in elements:
                    try:
                        # Type filter check
                        if type_patterns:
                            element_type = query.get_type(element)
                            if not element_type:
                                continue
                            
                            # Get name to check based on category
                            if category == DB.BuiltInCategory.OST_Walls:
                                name_to_check = database.get_name(element_type)
                            else:
                                try:
                                    name_to_check = element_type.FamilyName
                                except:
                                    name_to_check = database.get_name(element_type)
                            
                            # Check if matches any pattern
                            matches = False
                            for pattern in type_patterns:
                                if pattern in name_to_check:
                                    matches = True
                                    break
                            
                            if not matches:
                                skipped_no_match += 1
                                continue
                        
                        # Get element center for tag placement
                        center_point = get_element_center_point(element)
                        if not center_point:
                            continue
                        
                        tag_location = DB.XYZ(
                            center_point.X + self.config.TAG_OFFSET,
                            center_point.Y + self.config.TAG_OFFSET,
                            center_point.Z
                        )
                        
                        try:
                            new_tag = DB.IndependentTag.Create(
                                self.doc,
                                view.Id,
                                DB.Reference(element),
                                self.config.TAG_LEADER,
                                DB.TagMode.TM_ADDBY_CATEGORY,
                                DB.TagOrientation.Horizontal,
                                tag_location
                            )
                            
                            if new_tag:
                                view_tagged += 1
                                total_tagged += 1
                                tagged_in_category += 1
                            else:
                                failed_to_tag += 1
                        except Exception as tag_ex:
                            failed_to_tag += 1
                            if failed_to_tag <= 2:  # Show first 2 errors
                                self.output.print_md("    WARNING: Tag creation failed for {}: {}".format(
                                    element.Id, str(tag_ex)))
                    except Exception as elem_ex:
                        pass
                
                # Report results for this category with explanations
                if tagged_in_category > 0:
                    self.output.print_md("    **SUCCESS: Tagged {} {}**".format(
                        tagged_in_category, category_name
                    ))
                elif len(elements) > 0:
                    # Found elements but none tagged - explain why
                    if skipped_no_match > 0:
                        self.output.print_md("    **INFO: 0 tagged** - {} {} didn't match type filter **'{}'**".format(
                            skipped_no_match, category_name, type_patterns[0] if type_patterns else 'none'
                        ))
                        self.output.print_md("       Check that wall/door/window type names contain the filter text")
                    if failed_to_tag > 0:
                        self.output.print_md("    **WARNING: 0 tagged** - {} {} failed tag creation (see errors above)".format(
                            failed_to_tag, category_name
                        ))
                else:
                    self.output.print_md("    INFO: No {} found on this level".format(category_name))
            
            if view_tagged > 0:
                self.output.print_md("  **View Total: {} tags created**".format(view_tagged))
            self.output.print_md("")
        
        self.output.print_md("---")
        self.output.print_md("**TAGGING COMPLETE: {} elements tagged across all views**".format(total_tagged))
        self.total_tagged = total_tagged
        self.output.print_md("")
        
        return True
    
    def step6_create_sheets(self):
        """Step 6: Create sheets"""
    
        self.output.print_md("## STEP 6: Create Sheets")
    
        titleblock = get_titleblock_type(self.doc, self.config.TITLEBLOCK_FAMILY_NAME)
    
        if not titleblock:
            self.output.print_md("**ERROR:** Titleblock not found")
            return False
    
        # Check if level numbering is enabled
        use_level_numbering = False
        if self.series_config and self.series_config.get('level_numbering'):
            use_level_numbering = True
            self.output.print_md("*Level-based numbering enabled*")
    
        current_sheet_number = get_next_sheet_number(self.doc, self.selected_series)
    
        for view_data in self.created_views:
            view = view_data['view']
            level = view_data['level']
        
            sheet_name = "{} - {}".format(self.package_name, level.Name)
        
            # Apply level-based numbering if enabled
            if use_level_numbering:
                level_name = level.Name.upper()
            
                # Extract level number from name (handles "LEVEL 00", "LEVEL B1", etc)
                import re
            
                # Check for basement levels first (B1, B2, etc)
                basement_match = re.search(r'B(\d+)', level_name)
                if basement_match:
                    basement_num = int(basement_match.group(1))
                    level_suffix = str(100 - basement_num).zfill(3)  # B1=099, B2=098
                else:
                    # Standard levels (00, 01, 02, etc)
                    level_match = re.search(r'(\d+)', level_name)
                    if level_match:
                        level_num = int(level_match.group(1))
                        level_suffix = str(100 + level_num + 1).zfill(3)  # 00=101, 01=102
                    else:
                        # Fallback to sequential if pattern not found
                        level_suffix = current_sheet_number[-3:]
            
                # Replace last 3 digits with level suffix
                sheet_number = current_sheet_number[:-3] + level_suffix
            else:
                sheet_number = current_sheet_number
        
            try:
                existing_sheets = DB.FilteredElementCollector(self.doc).OfCategory(
                    DB.BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
            
                sheet_exists = False
                for existing_sheet in existing_sheets:
                    if existing_sheet.SheetNumber == sheet_number:
                        sheet_exists = True
                        self.output.print_md("- **INFO:** Sheet {} already exists, skipping".format(
                            sheet_number
                        ))
                        break
            
                if sheet_exists:
                    if not use_level_numbering:
                        current_sheet_number = self._increment_sheet_number(current_sheet_number)
                    continue
            
                sheet = DB.ViewSheet.Create(self.doc, titleblock.Id)
                sheet.SheetNumber = sheet_number
                sheet.Name = sheet_name
            
                self.created_sheets.append({
                    'sheet': sheet,
                    'view': view,
                    'number': sheet_number
                })
            
                self.output.print_md("- {} - {}".format(sheet_number, sheet_name))
            
                # Only increment if NOT using level numbering
                if not use_level_numbering:
                    current_sheet_number = self._increment_sheet_number(current_sheet_number)
            
            except Exception as e:
                self.output.print_md("  - **WARNING:** Failed to create sheet {}: {}".format(
                    sheet_number, str(e)
                ))
                if not use_level_numbering:
                    current_sheet_number = self._increment_sheet_number(current_sheet_number)
    
        self.output.print_md("**Created {} sheets**".format(len(self.created_sheets)))
        self.output.print_md("")
    
        return len(self.created_sheets) > 0
    
    def _increment_sheet_number(self, current_number):
        """Increment sheet number with proper zero padding"""
        try:
            num = int(current_number)
            num += 1
            return str(num).zfill(len(current_number))
        except:
            return current_number
    
    def step7_place_views(self):
        """Step 7: Place views"""
        
        self.output.print_md("## STEP 7: Place Views on Sheets")
        
        for sheet_data in self.created_sheets:
            sheet = sheet_data['sheet']
            view = sheet_data['view']
            
            try:
                center_point = DB.XYZ(
                    self.config.VIEW_CENTER_X,
                    self.config.VIEW_CENTER_Y,
                    0
                )
                
                DB.Viewport.Create(
                    self.doc,
                    sheet.Id,
                    view.Id,
                    center_point
                )
                
            except Exception as e:
                self.output.print_md("  - WARNING: {}".format(str(e)))
        
        self.output.print_md("")
        return True
    
    def step9_place_legend(self):
        """Step 9: Place legend on sheets"""
        
        self.output.print_md("## STEP 9: Place Legend on Sheets")
        
        if not self.legend_view:
            return True
        
        placed = 0
        failed = 0
        
        for sheet_data in self.created_sheets:
            try:
                # Use (0,0,0) for auto-centering - legend viewports behave differently than view viewports
                # The legend will be centered on the sheet, then you can manually adjust in Revit if needed
                legend_point = DB.XYZ(0, 0, 0)
                
                DB.Viewport.Create(
                    self.doc,
                    sheet_data['sheet'].Id,
                    self.legend_view.Id,
                    legend_point
                )
                placed += 1
                
            except Exception as e:
                failed += 1
                logger.debug("Failed to place legend on sheet {}: {}".format(
                    sheet_data['number'], str(e)
                ))
        
        self.output.print_md("**Placed legend on {} sheets**".format(placed))
        if failed > 0:
            self.output.print_md("**Failed: {} sheets**".format(failed))
        self.output.print_md("*Note: Legends are auto-centered. Adjust position manually in Revit if needed.*")
        self.output.print_md("")
        
        return True
    
    def print_summary(self):
        """Print summary"""
        
        self.output.print_md("---")
        self.output.print_md("# SUMMARY")
        self.output.print_md("")
        self.output.print_md("**Series:** {}".format(self.selected_series))
        self.output.print_md("**Package:** {}".format(self.package_name))
        self.output.print_md("")
        self.output.print_md("**Created:**")
        self.output.print_md("- {} Views".format(len(self.created_views)))
        self.output.print_md("- {} Elements Tagged".format(self.total_tagged))
        self.output.print_md("- {} Sheets".format(len(self.created_sheets)))
        
        if self.legend_view:
            self.output.print_md("- Legend: {}".format(self.legend_view.Name))
        
        if self.created_sheets:
            self.output.print_md("")
            self.output.print_md("**Sheet Numbers:** {} to {}".format(
                self.created_sheets[0]['number'],
                self.created_sheets[-1]['number']
            ))


# ================================================================================================
# SCRIPT ENTRY POINT
# ================================================================================================

if __name__ == '__main__':
    
    doc = revit.doc
    config = PackageConfig()
    output = script.get_output()
    
    if not doc or doc.IsReadOnly:
        forms.alert("Document must be open and editable", exitscript=True)
    
    # Show UI window
    window = DrawingPackageWindow(doc, config)
    result = window.ShowDialog()
    
    if not result:
        script.exit()
    
    # Check if batch mode
    if window.batch_mode:
        # BATCH MODE - Multiple series
        output.print_md("# BATCH MODE: DRAWING PACKAGE AUTOMATION")
        output.print_md("**Generating {} series**".format(len(window.batch_series)))
        output.print_md("---")
        output.print_md("")
        
        batch_success_count = 0
        batch_results = []
        
        for batch_item in window.batch_series:
            series_num = batch_item['series_num']
            series_cfg = batch_item['config']
            package_name = batch_item['package_name']
            
            output.print_md("## SERIES {}: {}".format(series_num, package_name))
            output.print_md("")
            
            automation = DrawingPackageAutomation(
                doc, 
                config,
                series_num,
                series_cfg,
                package_name,
                window.legend_view,
                window.create_legend
            )
            
            success = automation.run()
            
            batch_results.append({
                'series': series_num,
                'name': package_name,
                'success': success,
                'views': len(automation.created_views),
                'sheets': len(automation.created_sheets),
                'tags': automation.total_tagged
            })
            
            if success:
                batch_success_count += 1
            
            output.print_md("---")
            output.print_md("")
        
        # Print batch summary
        output.print_md("# BATCH SUMMARY")
        output.print_md("**Completed: {}/{}**".format(
            batch_success_count, len(window.batch_series)))
        output.print_md("")
        
        for result in batch_results:
            status = "✓ SUCCESS" if result['success'] else "✗ FAILED"
            output.print_md("**Series {}** - {} - {}".format(
                result['series'], result['name'], status))
            if result['success']:
                output.print_md("  - {} views, {} sheets, {} tags".format(
                    result['views'], result['sheets'], result['tags']))
        
        if batch_success_count == len(window.batch_series):
            forms.alert("All {} series completed successfully!".format(
                batch_success_count), title="Batch Complete")
        else:
            forms.alert("{}/{} series completed. Check output for details.".format(
                batch_success_count, len(window.batch_series)), 
                title="Batch Complete with Warnings")
    
    else:
        # SINGLE MODE - Original behavior
        automation = DrawingPackageAutomation(
            doc, 
            config,
            window.selected_series,
            window.series_config,
            window.package_name,
            window.legend_view,
            window.create_legend
        )
        
        success = automation.run()
        
        if success:
            forms.alert("Drawing package automation completed!", title="Success")
        else:
            forms.alert("Check output for details", title="Warning")
