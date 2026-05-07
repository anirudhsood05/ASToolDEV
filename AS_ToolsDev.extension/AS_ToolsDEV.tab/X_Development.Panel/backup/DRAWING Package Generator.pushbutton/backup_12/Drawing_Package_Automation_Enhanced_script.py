# -*- coding: utf-8 -*-
"""
DRAWING PACKAGE AUTOMATION TOOL
===============================

Author: Anirudh Sood
Compatible: Revit 2023-2025, Python 2.7
Enhanced with batch mode, live validation, and preview panel
"""

__title__ = "DRAWING Package Generator"
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
import System

def escape_xml(text):
    """Escape XML special characters for safe XAML injection"""
    if not text:
        return ''
    return str(text).replace('&', '&amp;') \
                    .replace('<', '&lt;') \
                    .replace('>', '&gt;') \
                    .replace('"', '&quot;') \
                    .replace("'", '&apos;')

# ================================================================================================
# XAML UI DEFINITIONS
# ================================================================================================

xaml_file = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Drawing Package Automation" Width="700" Height="700" 
    ShowInTaskbar="False" WindowStartupLocation="CenterScreen" 
    ResizeMode="NoResize" FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Area -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="Drawing Package Automation" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- Mode Selection -->
        <StackPanel Grid.Row="2" Orientation="Horizontal">
            <RadioButton x:Name="UI_single_mode" Content="Single Series" 
                         IsChecked="True" Margin="0,0,20,0" FontWeight="SemiBold"/>
            <RadioButton x:Name="UI_batch_mode" Content="Batch Mode (Select Multiple)" 
                         FontWeight="SemiBold"/>
        </StackPanel>

        <!-- Main Content Area -->
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="280"/>
            </Grid.ColumnDefinitions>
            
            <!-- LEFT: Series Selection -->
            <GroupBox Grid.Column="0" Header="Series Selection" Padding="8" Background="White">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Single Mode Panel -->
                    <Grid Grid.Row="0" x:Name="UI_single_panel">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="8"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <Label Grid.Row="0" Grid.Column="0" Content="Drawing Series:" 
                               FontWeight="Medium" VerticalAlignment="Center"/>
                        <ComboBox Grid.Row="0" Grid.Column="1" x:Name="UI_series_combo" 
                                  Height="24" VerticalAlignment="Center"/>
                        
                        <Label Grid.Row="2" Grid.Column="0" Content="Package Name:" 
                               FontWeight="Medium" VerticalAlignment="Center"/>
                        <TextBox Grid.Row="2" Grid.Column="1" x:Name="UI_package_text" 
                                 Height="24" VerticalAlignment="Center"
                                 ToolTip="Enter package name (e.g., PARTITIONS)"/>
                    </Grid>
                    
                    <!-- Batch Mode Panel -->
                    <Grid Grid.Row="1" x:Name="UI_batch_panel" Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                            <Button x:Name="UI_select_all_btn" Content="Select All" 
                                    Width="80" Height="24" Margin="0,0,8,0"/>
                            <Button x:Name="UI_clear_all_btn" Content="Clear All" 
                                    Width="80" Height="24"/>
                        </StackPanel>
                        
                        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" 
                                      MaxHeight="400" BorderBrush="#CCCCCC" BorderThickness="1">
                            <StackPanel x:Name="UI_series_checkboxes" Margin="4"/>
                        </ScrollViewer>
                        
                        <TextBlock Grid.Row="2" x:Name="UI_batch_count" 
                                   Text="0 series selected" 
                                   Foreground="#666666" FontSize="11" 
                                   Margin="0,4,0,0" FontStyle="Italic"/>
                    </Grid>
                </Grid>
            </GroupBox>
            
            <!-- RIGHT: Preview + Validation Panel -->
            <GroupBox Grid.Column="2" Header="Preview &amp; Validation" Padding="8" Background="#F8F8F8">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="8"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Preview Content -->
                    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="UI_preview_panel">
                            <!-- Validation Status -->
                            <Border x:Name="UI_validation_border" 
                                    Background="#E8F5E9" 
                                    Padding="8" CornerRadius="3" Margin="0,0,0,8">
                                <StackPanel>
                                    <TextBlock x:Name="UI_validation_icon" 
                                               Text="✓ Ready" 
                                               FontWeight="Bold" FontSize="13"
                                               Foreground="#2E7D32"/>
                                    <TextBlock x:Name="UI_validation_text" 
                                               Text="All checks passed"
                                               FontSize="11" Foreground="#666666"
                                               Margin="0,2,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                            </Border>
        
                            <!-- Requirements & Guidance (Collapsible) -->
                            <Expander x:Name="UI_guidance_expander" 
                                      Header="Requirements &amp; Guidance" 
                                      IsExpanded="False"
                                      Margin="0,0,0,8"
                                      BorderBrush="#CCCCCC" 
                                      BorderThickness="1"
                                      Background="White">
                                <Border Padding="8" Background="#FAFAFA">
                                    <ScrollViewer MaxHeight="200" VerticalScrollBarVisibility="Auto">
                                        <StackPanel x:Name="UI_guidance_content">
                                            <!-- Dynamic content populated by update_guidance_panel() -->
                                        </StackPanel>
                                    </ScrollViewer>
                                </Border>
                            </Expander>
        
                            <!-- Estimated Output -->
                            <TextBlock Text="Estimated Output:" 
                                       FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <StackPanel Margin="8,0,0,8">
                                <TextBlock x:Name="UI_estimate_views" 
                                           Text="• 0 views" 
                                           FontSize="11" Margin="0,2,0,2"/>
                                <TextBlock x:Name="UI_estimate_sheets" 
                                           Text="• 0 sheets" 
                                           FontSize="11" Margin="0,2,0,2"/>
                                <TextBlock x:Name="UI_estimate_tags" 
                                           Text="• ~0 tags" 
                                           FontSize="11" Margin="0,2,0,2"/>
                            </StackPanel>
                            
                            <!-- Sheet Ranges -->
                            <TextBlock x:Name="UI_sheet_range_title" 
                                       Text="Sheet Number Ranges:" 
                                       FontWeight="SemiBold" Margin="0,8,0,4" 
                                       Visibility="Collapsed"/>
                            <TextBlock x:Name="UI_sheet_ranges" 
                                       FontSize="11" FontFamily="Consolas" 
                                       Margin="8,0,0,8" TextWrapping="Wrap"/>
                            
                            <!-- Series Details -->
                            <TextBlock x:Name="UI_series_details_title" 
                                       Text="Series Details:" 
                                       FontWeight="SemiBold" Margin="0,8,0,4" 
                                       Visibility="Collapsed"/>
                            <StackPanel x:Name="UI_series_details" Margin="8,0,0,0"/>
                        </StackPanel>
                    </ScrollViewer>
                    
                    <!-- Refresh Button -->
                    <Button Grid.Row="2" x:Name="UI_refresh_preview" 
                            Content="↻ Refresh Preview" 
                            Height="24" FontSize="11"
                            HorizontalAlignment="Stretch"/>
                </Grid>
            </GroupBox>
        </Grid>

        <!-- Legend Options -->
        <!-- Legend Options -->
        <GroupBox Grid.Row="6" Header="Legend Options (Optional)" Padding="8" Background="White">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Existing Legend Placement -->
                <CheckBox Grid.Row="0" x:Name="UI_add_legend_check" 
                          Content="Add existing legend to all sheets" 
                          VerticalAlignment="Center" FontWeight="Medium"/>

                <Grid Grid.Row="2" x:Name="UI_legend_grid">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="120"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
            
                    <Label Grid.Column="0" Content="Select Legend:" 
                           FontWeight="Medium" VerticalAlignment="Center"/>
                    <ComboBox Grid.Column="1" x:Name="UI_legend_combo" 
                              Height="24" VerticalAlignment="Center" IsEnabled="False"/>
                </Grid>
        
                <!-- NEW: Filter Legend Generation -->
                <CheckBox Grid.Row="4" x:Name="UI_generate_filter_legends_check"
                          Content="Generate filter legends from view templates"
                          VerticalAlignment="Center" FontWeight="Medium"
                          ToolTip="Creates graphical legends showing filters applied via templates"/>
            </Grid>
        </GroupBox>

        <!-- Button Area -->
        <StackPanel Grid.Row="8" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="UI_generate_btn" Content="Generate Package" 
                    Width="130" Height="28" Margin="0,0,8,0" FontWeight="SemiBold"/>
            <Button x:Name="UI_cancel_btn" Content="Cancel" 
                    Width="90" Height="28"/>
        </StackPanel>
    </Grid>
</Window>
"""

progress_xaml = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Batch Generation Progress" Width="450" Height="200"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
    ShowInTaskbar="False" Background="#FFFFFF">
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" x:Name="UI_progress_title" 
                   Text="Generating Series 1/5..." 
                   FontSize="14" FontWeight="SemiBold"/>
        
        <ProgressBar Grid.Row="2" x:Name="UI_progress_bar" 
                     Height="24" Minimum="0" Maximum="100" Value="0"/>
        
        <TextBlock Grid.Row="4" x:Name="UI_progress_detail" 
                   Text="Current: Internal Walls Partitions"
                   FontSize="11" Foreground="#666666"/>
        
        <TextBlock Grid.Row="6" x:Name="UI_progress_status" 
                   Text="Initializing..."
                   FontSize="11" Foreground="#999999" 
                   HorizontalAlignment="Right"/>
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
        """Get CSV path with security validation"""
        try:
            from lib.secure_csv_config import SecureCSVConfig
            return SecureCSVConfig.get_safe_csv_path(__file__)
        except ImportError:
            # Fallback if secure loader not available
            script_dir = os.path.dirname(__file__)
            return os.path.join(script_dir, "drawing_series_config_enhanced.csv")
    
    # Titleblock
    TITLEBLOCK_FAMILY_NAME = "AUK_Titleblock_WorkingA1"
    
    # View placement (in feet) - coordinates from bottom-left of sheet
    VIEW_CENTER_X = 1.5  # 1.5 feet from left edge
    VIEW_CENTER_Y = 1.0  # 1.0 feet from bottom edge
    
    # Element types to tag (fallback if CSV not available)
    ELEMENTS_TO_TAG = {
        DB.BuiltInCategory.OST_Walls: ["AUK_Wall_Partition_"],
    }
    
    # Tag settings
    TAG_OFFSET = 0.2  # Offset from element center in feet
    TAG_LEADER = True

    # NEW: Security & Resource Limits
    MAX_BATCH_SERIES = 20  # Maximum series in batch mode
    MAX_BATCH_TIMEOUT_SECONDS = 3600  # 1 hour timeout
    MAX_VIEWS_PER_SERIES = 50  # Safety limit
    MAX_TAG_FAILURES_PER_VIEW = 20  # Abort if too many tag failures
    MAX_CHECKBOX_UI_COUNT = 50  # Prevent UI slowdown


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
        self.generate_filter_legends = False
        self.create_legend = False
        self.batch_mode = False
        self.batch_series = []
        
        # Initialize UI
        self.initialize()
        self.connect_events()
    
    def initialize(self):
        """Initialize UI elements"""
        
        # Load CSV configuration
        csv_path = self.config.get_csv_path()
        self.csv_config = load_series_config_from_csv(csv_path)
        
        # Populate series dropdown (single mode)
        if self.csv_config:
            sorted_series = sorted(self.csv_config.items(), key=lambda x: int(x[0]))
            for series_num, series_data in sorted_series:
                display_num = series_num.zfill(2)
                display_text = "{} - {}".format(display_num, series_data['name'])
                self.UI_series_combo.Items.Add(display_text)
            
            if self.UI_series_combo.Items.Count > 0:
                self.UI_series_combo.SelectedIndex = 0
        else:
            # Fallback to hardcoded list
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
        
        # Populate checkboxes (batch mode)
        self.populate_series_checkboxes()
        
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
        
        # Set default package name
        self.update_package_name_from_series()
        
        # Initial preview update
        self.update_preview_panel()
    
    def populate_series_checkboxes(self):
        """Create checkboxes for batch mode"""
        from System.Windows.Controls import CheckBox as WPFCheckBox
        from System.Windows import Thickness
        
        self.series_checkboxes = {}
        
        if self.csv_config:
            sorted_series = sorted(self.csv_config.items(), key=lambda x: int(x[0]))
            
            for series_num, series_data in sorted_series:
                display_num = series_num.zfill(2)
                display_text = "{} - {}".format(display_num, series_data['name'])
                
                checkbox = WPFCheckBox()
                checkbox.Content = display_text
                checkbox.Margin = Thickness(0, 2, 0, 2)
                checkbox.Tag = series_num
                checkbox.Checked += self.on_checkbox_changed
                checkbox.Unchecked += self.on_checkbox_changed
                
                self.UI_series_checkboxes.Children.Add(checkbox)
                self.series_checkboxes[series_num] = checkbox
        
        self.update_batch_count()
    
    def connect_events(self):
        """Connect UI event handlers"""
        self.UI_generate_btn.Click += self.on_generate
        self.UI_cancel_btn.Click += self.on_cancel
        self.UI_add_legend_check.Checked += self.on_legend_check_changed
        self.UI_add_legend_check.Unchecked += self.on_legend_check_changed
        self.UI_series_combo.SelectionChanged += self.on_series_changed
        # No event handler needed for filter legends checkbox - it's just a flag

        
        # Mode switching
        self.UI_single_mode.Checked += self.on_mode_changed
        self.UI_batch_mode.Checked += self.on_mode_changed
        
        # Batch controls
        self.UI_select_all_btn.Click += self.on_select_all
        self.UI_clear_all_btn.Click += self.on_clear_all
        
        # Preview refresh
        self.UI_refresh_preview.Click += self.on_refresh_preview
    
    def on_refresh_preview(self, sender, args):
        """Refresh preview panel"""
        self.update_preview_panel()
    
    def on_series_changed(self, sender, args):
        """Update package name and preview when series changes"""
        self.update_package_name_from_series()
        self.update_preview_panel()
    
    def on_checkbox_changed(self, sender, args):
        """Update count and preview when checkbox changes"""
        self.update_batch_count()
        self.update_preview_panel()
    
    def on_mode_changed(self, sender, args):
        """Switch between single and batch mode"""
        from System.Windows import Visibility
        
        if self.UI_single_mode.IsChecked:
            self.UI_single_panel.Visibility = Visibility.Visible
            self.UI_batch_panel.Visibility = Visibility.Collapsed
        else:
            self.UI_single_panel.Visibility = Visibility.Collapsed
            self.UI_batch_panel.Visibility = Visibility.Visible
        
        self.update_preview_panel()
    
    def on_select_all(self, sender, args):
        """Select all series checkboxes"""
        for checkbox in self.series_checkboxes.values():
            checkbox.IsChecked = True
    
    def on_clear_all(self, sender, args):
        """Clear all series checkboxes"""
        for checkbox in self.series_checkboxes.values():
            checkbox.IsChecked = False
    
    def update_batch_count(self):
        """Update selected series count"""
        count = sum(1 for cb in self.series_checkboxes.values() if cb.IsChecked)
        self.UI_batch_count.Text = "{} series selected".format(count)
    
    def update_preview_panel(self):
        """Update preview panel with validation and estimates"""
        from System.Windows import Visibility
        from System.Windows.Media import Brushes
        from System.Windows.Controls import TextBlock
        
        # Get selected series
        if self.UI_batch_mode.IsChecked:
            selected_series = [num for num, cb in self.series_checkboxes.items() if cb.IsChecked]
        else:
            if self.UI_series_combo.SelectedItem:
                selected_text = str(self.UI_series_combo.SelectedItem)
                series_num = selected_text.split(" - ")[0].lstrip('0') or '0'
                selected_series = [series_num]
            else:
                selected_series = []
        
        if not selected_series:
            self.show_empty_preview()
            return
        
        # Run validation
        validation_results = self.run_live_validation(selected_series)
        
        # Update validation status
        if validation_results['critical_issues']:
            self.UI_validation_border.Background = Brushes.MistyRose
            self.UI_validation_icon.Text = "✗ Cannot Generate"
            self.UI_validation_icon.Foreground = Brushes.DarkRed
            self.UI_validation_text.Text = "{} critical issue(s)".format(
                len(validation_results['critical_issues']))
        elif validation_results['warnings']:
            self.UI_validation_border.Background = Brushes.LightYellow
            self.UI_validation_icon.Text = "⚠ Warnings"
            self.UI_validation_icon.Foreground = Brushes.DarkOrange
            self.UI_validation_text.Text = "{} warning(s) - can continue".format(
                len(validation_results['warnings']))
        else:
            self.UI_validation_border.Background = Brushes.LightGreen
            self.UI_validation_icon.Text = "✓ Ready"
            self.UI_validation_icon.Foreground = Brushes.DarkGreen
            self.UI_validation_text.Text = "All checks passed"
        
        # Calculate estimates
        levels = get_all_levels(self.doc)
        level_count = len(levels) if levels else 0
        
        total_views = len(selected_series) * level_count
        total_sheets = total_views
        total_tags_estimate = 0
        
        for series_num in selected_series:
            series_cfg = self.csv_config.get(series_num, {})
            tags_per_view = 0
            if series_cfg.get('tag_rooms'): tags_per_view += 5
            if series_cfg.get('tag_walls'): tags_per_view += 15
            if series_cfg.get('tag_doors'): tags_per_view += 8
            if series_cfg.get('tag_windows'): tags_per_view += 6
            if series_cfg.get('tag_ceilings'): tags_per_view += 10
            total_tags_estimate += tags_per_view * level_count
        
        # Update estimates
        self.UI_estimate_views.Text = "• {} views".format(total_views)
        self.UI_estimate_sheets.Text = "• {} sheets".format(total_sheets)
        self.UI_estimate_tags.Text = "• ~{} tags (estimated)".format(total_tags_estimate)
        
        # Sheet ranges
        if len(selected_series) > 0:
            self.UI_sheet_range_title.Visibility = Visibility.Visible
            ranges = []
            for series_num in sorted(selected_series, key=int):
                start = get_next_sheet_number(self.doc, series_num)
                end_num = int(start) + level_count - 1
                end = str(end_num).zfill(5)
                ranges.append("{}-{}".format(start, end))
            self.UI_sheet_ranges.Text = "\n".join(ranges)
        else:
            self.UI_sheet_range_title.Visibility = Visibility.Collapsed
            self.UI_sheet_ranges.Text = ""
        
        # Series details with validation icons
        self.UI_series_details.Children.Clear()
        if len(selected_series) > 1:
            self.UI_series_details_title.Visibility = Visibility.Visible
            
            for series_num in sorted(selected_series, key=int):
                series_cfg = self.csv_config.get(series_num, {})
                series_name = series_cfg.get('name', 'Unknown')
                
                # Check validation for this series
                series_issues = validation_results['series_issues'].get(series_num, [])
                
                detail_text = TextBlock()
                detail_text.FontSize = 11
                detail_text.Margin = System.Windows.Thickness(0, 2, 0, 2)
                
                if series_issues:
                    detail_text.Text = "⚠ {} - {}".format(series_num.zfill(2), series_name)
                    detail_text.Foreground = Brushes.DarkOrange
                    detail_text.ToolTip = escape_xml("\n".join(series_issues))

                else:
                    detail_text.Text = "✓ {} - {}".format(series_num.zfill(2), series_name)
                    detail_text.Foreground = Brushes.DarkGreen
                
                self.UI_series_details.Children.Add(detail_text)
        else:
            self.UI_series_details_title.Visibility = Visibility.Collapsed

        # *** ADD THIS LINE AT THE END: ***
        self.update_guidance_panel(selected_series)
    
    def show_empty_preview(self):
        """Show empty state in preview panel"""
        from System.Windows.Media import Brushes
        from System.Windows import Visibility
        
        self.UI_validation_border.Background = Brushes.LightGray
        self.UI_validation_icon.Text = "○ No Selection"
        self.UI_validation_icon.Foreground = Brushes.Gray
        self.UI_validation_text.Text = "Select a series to see preview"
        
        self.UI_estimate_views.Text = "• 0 views"
        self.UI_estimate_sheets.Text = "• 0 sheets"
        self.UI_estimate_tags.Text = "• ~0 tags"
        
        self.UI_sheet_range_title.Visibility = Visibility.Collapsed
        self.UI_sheet_ranges.Text = ""
        self.UI_series_details_title.Visibility = Visibility.Collapsed
        self.UI_series_details.Children.Clear()
    
    def run_live_validation(self, selected_series):
        """Run validation checks and return results"""
        
        results = {
            'critical_issues': [],
            'warnings': [],
            'series_issues': {}
        }
        
        # Global checks
        titleblock = get_titleblock_type(self.doc, self.config.TITLEBLOCK_FAMILY_NAME)
        if not titleblock:
            results['critical_issues'].append("Titleblock not found")
        
        levels = get_all_levels(self.doc)
        if not levels:
            results['critical_issues'].append("No levels in project")
        
        # Series-specific checks
        for series_num in selected_series:
            series_cfg = self.csv_config.get(series_num, {})
            issues = []
            
            # Template check
            template_pattern = series_cfg.get('template_pattern', '')
            if template_pattern:
                template = get_view_template_by_pattern(self.doc, template_pattern)
                if not template:
                    issues.append("Template not found")
            
            # Tag family checks
            if series_cfg.get('tag_rooms'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_RoomTags):
                    issues.append("Room tags not loaded")
            
            if series_cfg.get('tag_walls'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_WallTags):
                    issues.append("Wall tags not loaded")
            
            if series_cfg.get('tag_doors'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_DoorTags):
                    issues.append("Door tags not loaded")
            
            if series_cfg.get('tag_windows'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_WindowTags):
                    issues.append("Window tags not loaded")
            
            if series_cfg.get('tag_ceilings'):
                if not self.check_tag_family_loaded(DB.BuiltInCategory.OST_CeilingTags):
                    issues.append("Ceiling tags not loaded")
            
            # Area scheme check
            if series_cfg.get('view_type') == 'AreaPlan':
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
                results['series_issues'][series_num] = issues
                results['warnings'].extend(issues)
        
        return results
    
    def check_tag_family_loaded(self, category):
        """Check if tag family is loaded for category"""
        try:
            collector = DB.FilteredElementCollector(self.doc)
            tags = collector.OfCategory(category).WhereElementIsElementType().ToElements()
            return len(list(tags)) > 0
        except:
            return False
    
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
        
        if not self.validate_inputs():
            return
        
        # Check mode
        if self.UI_batch_mode.IsChecked == True:
            self.setup_batch_mode()
        else:
            self.setup_single_mode()
        
        self.DialogResult = True
        self.Close()
    
    def setup_batch_mode(self):
        """Setup for batch generation"""
        self.batch_mode = True
        self.batch_series = []
    
        # Collect checked series
        for series_num, checkbox in self.series_checkboxes.items():
            if checkbox.IsChecked:
                series_cfg = self.csv_config[series_num]
                self.batch_series.append({
                    'series_num': series_num,
                    'config': series_cfg,
                    'package_name': series_cfg['name'].upper()
                })
    
        # Sort by series number
        self.batch_series.sort(key=lambda x: int(x['series_num']))
    
        # Get legend
        self.create_legend = self.UI_add_legend_check.IsChecked == True
        if self.create_legend and self.UI_legend_combo.SelectedItem:
            legend_name = str(self.UI_legend_combo.SelectedItem)
            all_legends = get_all_legend_views(self.doc)
            for legend in all_legends:
                if legend.Name == legend_name:
                    self.legend_view = legend
                    break

        # NEW: Filter legend generation flag
        self.generate_filter_legends = self.UI_generate_filter_legends_check.IsChecked == True
    
    def setup_single_mode(self):
        """Setup for single series generation"""
        self.batch_mode = False
        
        selected_text = str(self.UI_series_combo.SelectedItem)
        if " - " in selected_text:
            series_display = selected_text.split(" - ")[0]
        else:
            series_display = selected_text.split()[0]
        
        self.selected_series = series_display.lstrip('0') or '0'
        
        if self.csv_config and self.selected_series in self.csv_config:
            self.series_config = self.csv_config[self.selected_series]
        else:
            if self.csv_config and series_display in self.csv_config:
                self.selected_series = series_display
                self.series_config = self.csv_config[series_display]
        
        self.package_name = self.UI_package_text.Text.strip()
        
        self.create_legend = self.UI_add_legend_check.IsChecked == True
        if self.create_legend and self.UI_legend_combo.SelectedItem:
            legend_name = str(self.UI_legend_combo.SelectedItem)
            all_legends = get_all_legend_views(self.doc)
            for legend in all_legends:
                if legend.Name == legend_name:
                    self.legend_view = legend
                    break
        
        # NEW: Filter legend generation flag
        self.generate_filter_legends = self.UI_generate_filter_legends_check.IsChecked == True
    
    def on_cancel(self, sender, args):
        """Cancel button clicked"""
        self.DialogResult = False
        self.Close()
    
    def validate_inputs(self):
        """Validate user inputs"""
        
        # Check if batch mode
        if self.UI_batch_mode.IsChecked == True:
            return self.validate_batch_mode()
        
        # Single mode validation
        if not self.UI_series_combo.SelectedItem:
            forms.alert("Please select a drawing series.", title="Input Required")
            return False
        
        package_name = self.UI_package_text.Text.strip()
        if not package_name:
            forms.alert("Please enter a package name.", title="Input Required")
            return False
        
        if self.UI_add_legend_check.IsChecked == True:
            if not self.UI_legend_combo.SelectedItem:
                forms.alert("Please select a legend or disable the legend option.", 
                           title="Input Required")
                return False
        
        return self.validate_project_setup()
    
    def validate_batch_mode(self):
        """Validate batch mode setup with resource limits"""
    
        # Get selected series
        selected = [num for num, cb in self.series_checkboxes.items() if cb.IsChecked]
    
        if not selected:
            forms.alert("Please select at least one series.", title="No Series Selected")
            return False
    
        # NEW: Check maximum limit
        if len(selected) > PackageConfig.MAX_BATCH_SERIES:
            forms.alert(
                "Maximum {} series allowed in batch mode.\n\n"
                "You selected: {}\n\n"
                "Please reduce your selection.".format(
                    PackageConfig.MAX_BATCH_SERIES, len(selected)
                ),
                title="Too Many Series"
            )
            return False
    
        # Calculate estimated time
        levels = get_all_levels(self.doc)
        estimated_minutes = len(selected) * len(levels) * 0.5  # ~30 sec per view
    
        if estimated_minutes > 60:
            result = forms.alert(
                "BATCH MODE: Will generate {} series\n\n"
                "Estimated time: {} minutes\n\n"
                "This is a long operation. Continue?".format(
                    len(selected), int(estimated_minutes)
                ),
                title="Confirm Long Operation",
                ok=False,
                yes=True,
                no=True
            )
            if not result:
                return False
    
        # Show confirmation
        series_list = "\n".join([" - Series {}: {}".format(
            num, self.csv_config[num]['name']) for num in sorted(selected, key=int)])
    
        confirm = forms.alert(
            "BATCH MODE: Will generate {} series:\n\n{}\n\n"
            "Continue?".format(len(selected), series_list),
            title="Confirm Batch Generation",
            ok=False,
            yes=True,
            no=True
        )
    
        if not confirm:
            return False
    
        return self.validate_project_setup_batch(selected)
    
    def validate_project_setup(self):
        """Pre-flight checks for single series"""
        
        issues = []
        critical_issues = []
        
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
        
        # Check 2: Titleblock exists (CRITICAL)
        titleblock = get_titleblock_type(self.doc, self.config.TITLEBLOCK_FAMILY_NAME)
        if not titleblock:
            critical_issues.append("Titleblock not found: '{}'".format(
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
        
        # Check 4: Levels exist (CRITICAL)
        levels = get_all_levels(self.doc)
        if not levels:
            critical_issues.append("No levels found in project")
        
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
        if critical_issues:
            message = "PRE-FLIGHT VALIDATION FAILED:\n\n" + "\n".join(
                [" • {}".format(issue) for issue in critical_issues]
            )
            message += "\n\nCannot continue. Fix these critical issues first."
            forms.alert(message, title="Validation Failed")
            return False
        
        if issues:
            message = "PRE-FLIGHT VALIDATION WARNINGS:\n\n" + "\n".join(
                [" • {}".format(issue) for issue in issues]
            )
            message += "\n\nThese issues may cause problems but you can continue.\n\n"
            message += "Continue anyway?"
            
            result = forms.alert(
                message,
                title="Validation Warnings",
                ok=False,
                yes=True,
                no=True
            )
            
            return result
        
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
            message = "PRE-FLIGHT VALIDATION WARNINGS:\n\n"
            
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
            
            # CRITICAL vs NON-CRITICAL
            has_critical = len(global_issues) > 0
            
            if has_critical:
                message += "CRITICAL issues found. Cannot continue."
                forms.alert(message, title="Validation Failed")
                return False
            else:
                message += "These issues may cause problems but you can continue.\n\n"
                message += "Continue anyway?"
                
                result = forms.alert(
                    message,
                    title="Validation Warnings",
                    ok=False,
                    yes=True,
                    no=True
                )
                
                return result
        
        return True

    def update_guidance_panel(self, selected_series_list):
        """
        Update the guidance panel with requirements and tagging info for selected series
    
        Args:
            selected_series_list: List of series numbers (e.g., ['22', '35'])
        """
        from System.Windows.Controls import TextBlock, Separator
        from System.Windows.Media import Brushes
        from System.Windows import Thickness

        # Clear existing content
        self.UI_guidance_content.Children.Clear()

        if not selected_series_list:
            empty_text = TextBlock()
            empty_text.Text = "Select a series to see requirements"
            empty_text.Foreground = Brushes.Gray
            empty_text.FontStyle = System.Windows.FontStyles.Italic
            empty_text.FontSize = 11
            self.UI_guidance_content.Children.Add(empty_text)
            return

        # Get current validation state
        validation_results = self.run_live_validation(selected_series_list)

        # For single series - show detailed guidance
        if len(selected_series_list) == 1:
            self._add_single_series_guidance(selected_series_list[0], validation_results)
        else:
            # For multiple series - show summarized guidance
            self._add_batch_guidance(selected_series_list, validation_results)


    def _add_single_series_guidance(self, series_num, validation_results):
        """Add detailed guidance for single series - CSV driven"""
        from System.Windows.Controls import TextBlock, Separator
        from System.Windows.Media import Brushes
        from System.Windows import Thickness

        series_cfg = self.csv_config.get(series_num)
        if not series_cfg:
            no_config = TextBlock()
            no_config.Text = "No configuration available for this series"
            no_config.Foreground = Brushes.Gray
            no_config.FontSize = 11
            self.UI_guidance_content.Children.Add(no_config)
            return

        # View Type Description
        view_type_map = {
            'FloorPlan': 'Floor Plan views',
            'CeilingPlan': 'Reflected Ceiling Plan (RCP) views',
            'AreaPlan': 'Area Plan views'
        }
        view_type = series_cfg.get('view_type', 'FloorPlan')
        view_desc = TextBlock()
        view_desc.Text = "Creates: {}".format(view_type_map.get(view_type, view_type))
        view_desc.FontSize = 11
        view_desc.FontWeight = System.Windows.FontWeights.SemiBold
        view_desc.Margin = Thickness(0, 0, 0, 8)
        view_desc.TextWrapping = System.Windows.TextWrapping.Wrap
        self.UI_guidance_content.Children.Add(view_desc)

        # Prerequisites Section
        prereq_header = TextBlock()
        prereq_header.Text = "Prerequisites:"
        prereq_header.FontSize = 11
        prereq_header.FontWeight = System.Windows.FontWeights.SemiBold
        prereq_header.Margin = Thickness(0, 0, 0, 4)
        self.UI_guidance_content.Children.Add(prereq_header)

        # Build prerequisites dynamically from config
        prerequisites = self._build_prerequisites_from_config(series_num, series_cfg)
        series_issues = validation_results['series_issues'].get(series_num, [])
    
        for prereq_name, prereq_desc, is_critical in prerequisites:
            prereq_item = TextBlock()
            prereq_item.FontSize = 10
            prereq_item.Margin = Thickness(8, 0, 0, 2)
            prereq_item.TextWrapping = System.Windows.TextWrapping.Wrap
        
            # Check if this prerequisite has issues
            has_issue = self._check_prerequisite_has_issue(
                prereq_name, series_issues, validation_results)
        
            if has_issue:
                if is_critical:
                    prereq_item.Text = u"🚫 " + prereq_name + ": " + prereq_desc
                    prereq_item.Foreground = Brushes.DarkRed
                else:
                    prereq_item.Text = u"⚠ " + prereq_name + ": " + prereq_desc
                    prereq_item.Foreground = Brushes.DarkOrange
            else:
                prereq_item.Text = u"✓ " + prereq_name + ": " + prereq_desc
                prereq_item.Foreground = Brushes.DarkGreen
        
            self.UI_guidance_content.Children.Add(prereq_item)

        # Separator
        sep1 = Separator()
        sep1.Margin = Thickness(0, 8, 0, 8)
        self.UI_guidance_content.Children.Add(sep1)

        # Tagging Info Section (built from CSV config)
        tagging_header = TextBlock()
        tagging_header.Text = "What Gets Tagged:"
        tagging_header.FontSize = 11
        tagging_header.FontWeight = System.Windows.FontWeights.SemiBold
        tagging_header.Margin = Thickness(0, 0, 0, 4)
        self.UI_guidance_content.Children.Add(tagging_header)

        tagging_info = self._build_tagging_info_from_config(series_cfg)
        for tag_info in tagging_info:
            tag_item = TextBlock()
            tag_item.Text = "• " + tag_info
            tag_item.FontSize = 10
            tag_item.Margin = Thickness(8, 0, 0, 2)
            tag_item.Foreground = Brushes.DarkBlue
            tag_item.TextWrapping = System.Windows.TextWrapping.Wrap
            self.UI_guidance_content.Children.Add(tag_item)

        # User Guidance Notes (from CSV)
        user_guidance = series_cfg.get('user_guidance', '').strip()
        if user_guidance:
            sep2 = Separator()
            sep2.Margin = Thickness(0, 8, 0, 8)
            self.UI_guidance_content.Children.Add(sep2)
        
            notes_block = TextBlock()
            notes_block.Text = "Note: " + user_guidance
            notes_block.FontSize = 10
            notes_block.FontStyle = System.Windows.FontStyles.Italic
            notes_block.Foreground = Brushes.DarkSlateGray
            notes_block.TextWrapping = System.Windows.TextWrapping.Wrap
            self.UI_guidance_content.Children.Add(notes_block)

    def _build_prerequisites_from_config(self, series_num, series_cfg):
        """Build prerequisites list dynamically from series configuration
    
        Returns list of tuples: (name, description, is_critical)
        """
        prerequisites = []
    
        # Always required
        prerequisites.append((
            'Levels', 
            'Project must have levels defined', 
            True  # Critical
        ))
        prerequisites.append((
            'Titleblock', 
            'AUK_Titleblock_WorkingA1 must be loaded', 
            True  # Critical
        ))
    
        # View template
        template = series_cfg.get('template_pattern', '')
        if template:
            prerequisites.append((
                'View Template',
                "Template '{}' should exist".format(template),
                False  # Warning
            ))
    
        # Area scheme for area plans
        if series_cfg.get('view_type') == 'AreaPlan':
            prerequisites.append((
                'Area Scheme',
                'GIA or GEA area scheme MUST exist',
                True  # Critical
            ))
    
        # Tag families
        if series_cfg.get('tag_rooms'):
            prerequisites.append((
                'Room Tags',
                'Room tag family should be loaded',
                False
            ))
    
        if series_cfg.get('tag_walls'):
            prerequisites.append((
                'Wall Tags',
                'Wall tag family should be loaded',
                False
            ))
    
        if series_cfg.get('tag_doors'):
            prerequisites.append((
                'Door Tags',
                'Door tag family should be loaded',
                False
            ))
    
        if series_cfg.get('tag_windows'):
            prerequisites.append((
                'Window Tags',
                'Window tag family should be loaded',
                False
            ))
    
        if series_cfg.get('tag_ceilings'):
            prerequisites.append((
                'Ceiling Tags',
                'Ceiling tag family should be loaded',
                False
            ))
    
        return prerequisites
    def _build_tagging_info_from_config(self, series_cfg):
        """Build tagging information from series configuration
    
        Returns list of strings describing what gets tagged
        """
        tagging_info = []
    
        # Check view type first
        view_type = series_cfg.get('view_type', 'FloorPlan')
    
        if view_type == 'CeilingPlan':
            # RCP views - different tagging behavior
            if series_cfg.get('tag_rooms'):
                tagging_info.append('Rooms: All placed rooms')
            if series_cfg.get('tag_ceilings'):
                tagging_info.append('Ceilings: All ceiling types')
            tagging_info.append('NOTE: Walls, doors, windows NOT visible in RCP views')
            return tagging_info
    
        # Standard floor plans or area plans
        if series_cfg.get('tag_rooms'):
            tagging_info.append('Rooms: All placed rooms with area > 0')
    
        if series_cfg.get('tag_walls'):
            wall_filter = series_cfg.get('wall_filter', '')
            if wall_filter:
                tagging_info.append("Walls: Only types containing '{}'".format(wall_filter))
            else:
                tagging_info.append('Walls: All wall types')
    
        if series_cfg.get('tag_doors'):
            door_filter = series_cfg.get('door_filter', '')
            if door_filter:
                tagging_info.append("Doors: Only families containing '{}'".format(door_filter))
            else:
                tagging_info.append('Doors: All door types')
    
        if series_cfg.get('tag_windows'):
            window_filter = series_cfg.get('window_filter', '')
            if window_filter:
                tagging_info.append("Windows: Only families containing '{}'".format(window_filter))
            else:
                tagging_info.append('Windows: All window types')
    
        if series_cfg.get('tag_ceilings'):
            tagging_info.append('Ceilings: All ceiling types')
    
        # If nothing is tagged
        if not tagging_info:
            tagging_info.append('NO AUTOMATIC TAGGING - manually tag elements as needed')
    
        return tagging_info


    def _check_prerequisite_has_issue(self, prereq_name, series_issues, validation_results):
        """Check if a specific prerequisite has validation issues
    
        Returns True if there's an issue, False if OK
        """
        # Check global critical issues
        if prereq_name == 'Titleblock':
            return any('Titleblock' in issue for issue in validation_results['critical_issues'])
    
        if prereq_name == 'Levels':
            return any('level' in issue.lower() for issue in validation_results['critical_issues'])
    
        # Check series-specific issues
        if prereq_name == 'View Template':
            return any('Template' in issue for issue in series_issues)
    
        if prereq_name == 'Area Scheme':
            return any('area scheme' in issue.lower() for issue in series_issues)
    
        if 'Tags' in prereq_name:
            tag_type = prereq_name.replace(' Tags', '').lower()
            return any(tag_type in issue.lower() and 'tag' in issue.lower() 
                       for issue in series_issues)
    
        return False

    def _add_batch_guidance(self, selected_series_list, validation_results):
        """Add summarized guidance for multiple series"""
        from System.Windows.Controls import TextBlock, Separator
        from System.Windows.Media import Brushes
        from System.Windows import Thickness

        # Header
        header = TextBlock()
        header.Text = "Batch Mode - {} series selected".format(len(selected_series_list))
        header.FontSize = 11
        header.FontWeight = System.Windows.FontWeights.SemiBold
        header.Margin = Thickness(0, 0, 0, 8)
        self.UI_guidance_content.Children.Add(header)

        # Common prerequisites
        common_header = TextBlock()
        common_header.Text = "Common Prerequisites (all series):"
        common_header.FontSize = 11
        common_header.FontWeight = System.Windows.FontWeights.SemiBold
        common_header.Margin = Thickness(0, 0, 0, 4)
        self.UI_guidance_content.Children.Add(common_header)

        # Check global issues
        has_titleblock = len([i for i in validation_results['critical_issues'] 
                              if 'Titleblock' in i]) == 0
        has_levels = len([i for i in validation_results['critical_issues'] 
                          if 'levels' in i]) == 0

        prereq1 = TextBlock()
        prereq1.FontSize = 10
        prereq1.Margin = Thickness(8, 0, 0, 2)
        if has_titleblock:
            prereq1.Text = u"✓ Titleblock: AUK_Titleblock_WorkingA1 loaded"
            prereq1.Foreground = Brushes.DarkGreen
        else:
            prereq1.Text = u"🚫 Titleblock: AUK_Titleblock_WorkingA1 MISSING"
            prereq1.Foreground = Brushes.DarkRed
        self.UI_guidance_content.Children.Add(prereq1)

        prereq2 = TextBlock()
        prereq2.FontSize = 10
        prereq2.Margin = Thickness(8, 0, 0, 2)
        if has_levels:
            prereq2.Text = u"✓ Levels: Project has levels defined"
            prereq2.Foreground = Brushes.DarkGreen
        else:
            prereq2.Text = u"🚫 Levels: NO LEVELS FOUND IN PROJECT"
            prereq2.Foreground = Brushes.DarkRed
        self.UI_guidance_content.Children.Add(prereq2)

        sep = Separator()
        sep.Margin = Thickness(0, 8, 0, 8)
        self.UI_guidance_content.Children.Add(sep)

        # Per-series status summary
        series_header = TextBlock()
        series_header.Text = "Series Status:"
        series_header.FontSize = 11
        series_header.FontWeight = System.Windows.FontWeights.SemiBold
        series_header.Margin = Thickness(0, 0, 0, 4)
        self.UI_guidance_content.Children.Add(series_header)

        for series_num in sorted(selected_series_list, key=int):
            series_cfg = self.csv_config.get(series_num, {})
            series_name = series_cfg.get('name', 'Unknown')
            series_issues = validation_results['series_issues'].get(series_num, [])
        
            series_item = TextBlock()
            series_item.FontSize = 10
            series_item.Margin = Thickness(8, 0, 0, 2)
            series_item.TextWrapping = System.Windows.TextWrapping.Wrap
        
            if series_issues:
                series_item.Text = u"⚠ {} - {} ({} warnings)".format(
                    series_num.zfill(2), series_name, len(series_issues))
                series_item.Foreground = Brushes.DarkOrange
                series_item.ToolTip = escape_xml("\n".join(series_issues))
            else:
                series_item.Text = u"✓ {} - {}".format(
                    series_num.zfill(2), series_name)
                series_item.Foreground = Brushes.DarkGreen
        
            self.UI_guidance_content.Children.Add(series_item)

        # General note
        note = TextBlock()
        note.Text = "Click single series mode to see detailed guidance for each series"
        note.FontSize = 10
        note.FontStyle = System.Windows.FontStyles.Italic
        note.Foreground = Brushes.Gray
        note.Margin = Thickness(0, 8, 0, 0)
        note.TextWrapping = System.Windows.TextWrapping.Wrap
        self.UI_guidance_content.Children.Add(note)


    def _check_prerequisite_status(self, series_num, prereq_name, validation_results):
        """
        Check if a specific prerequisite has validation issues
    
        Returns True if there's an issue, False if OK
        """
        # Check global critical issues
        if prereq_name == 'Titleblock':
            return any('Titleblock' in issue for issue in validation_results['critical_issues'])
    
        if prereq_name == 'Levels':
            return any('level' in issue.lower() for issue in validation_results['critical_issues'])
    
        # Check series-specific issues
        series_issues = validation_results['series_issues'].get(series_num, [])
    
        if prereq_name == 'View Template':
            return any('Template' in issue for issue in series_issues)
    
        if prereq_name == 'Area Scheme':
            return any('area scheme' in issue.lower() for issue in series_issues)
    
        if 'Tags' in prereq_name:
            tag_type = prereq_name.replace(' Tags', '').lower()
            return any(tag_type in issue.lower() and 'tag' in issue.lower() 
                       for issue in series_issues)
    
        return False


class ProgressWindow(Window):
    """Progress window for batch operations"""
    
    def __init__(self):
        xaml_stream = StringReader(progress_xaml)
        wpf.LoadComponent(self, xaml_stream)
        
    def update_progress(self, current, total, series_name, status="Processing..."):
        """Update progress display"""
        percentage = int((float(current) / total) * 100) if total > 0 else 0
        
        self.UI_progress_title.Text = "Generating Series {}/{}...".format(current, total)
        self.UI_progress_bar.Value = percentage
        self.UI_progress_detail.Text = "Current: {}".format(series_name)
        self.UI_progress_status.Text = status
        
        # Force UI update
        from System.Windows.Threading import Dispatcher
        Dispatcher.CurrentDispatcher.Invoke(
            System.Action(lambda: None),
            System.Windows.Threading.DispatcherPriority.Background
        )


# ================================================================================================
# UTILITY FUNCTIONS
# ================================================================================================

logger = script.get_logger()

# REPLACE THIS FUNCTION (around line 440)
def load_series_config_from_csv(csv_path):
    """Load series configuration from CSV file with security validation"""
    
    # Import the secure loader
    try:
        from lib.secure_csv_config import SecureCSVConfig
    except ImportError:
        logger.error("Cannot import SecureCSVConfig - using unsafe fallback")
        # Fall back to old method if import fails
        return load_series_config_from_csv_UNSAFE(csv_path)
    
    try:
        return SecureCSVConfig.load_config(csv_path)
    except (ValueError, IOError) as e:
        logger.error("Failed to load CSV config: {}".format(str(e)))
        from pyrevit import forms
        forms.alert(
            "CSV Configuration Error:\n\n{}\n\n"
            "Using fallback hardcoded series list.".format(str(e)),
            title="Configuration Warning"
        )
        return None

def load_series_config_from_csv_UNSAFE(csv_path):
    """UNSAFE: Original CSV loader - kept as emergency fallback only"""
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
                        continue
                    
                    def safe_get(key, default=''):
                        value = row.get(key, default)
                        if value is None:
                            return default
                        return str(value).strip()
                    
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
                        'level_numbering': safe_bool('LevelNumbering'),
                        'auto_generate': safe_bool('AutoGenerate'),
                        'wall_filter': safe_get('WallFilter'),
                        'door_filter': safe_get('DoorFilter'),
                        'window_filter': safe_get('WindowFilter'),
                        'user_guidance': safe_get('UserGuidance'),  # NEW: Load guidance from CSV
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
                 legend_view, create_legend, generate_filter_legends=False):
        self.doc = doc
        self.config = config
        self.output = script.get_output()
    
        self.selected_series = selected_series
        self.series_config = series_config
        self.package_name = package_name
        self.legend_view = legend_view
        self.create_legend = create_legend
        self.generate_filter_legends = generate_filter_legends  # ADD THIS LINE
    
        self.created_views = []
        self.created_sheets = []
        self.total_tagged = 0
        self.generated_filter_legend = None  # ADD THIS

    def get_text_type_by_name(self, text_type_name):
        """Get TextNoteType by name with fallback"""
        all_text_types = DB.FilteredElementCollector(self.doc)\
            .OfClass(DB.TextNoteType)\
            .WhereElementIsElementType()\
            .ToElements()
    
        # Try exact match first
        for tt in all_text_types:
            type_name = tt.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            if type_name == text_type_name:
                self.output.print_md("*Using text type: {}*".format(text_type_name))
                return tt
    
        # Fallback to first available
        if all_text_types:
            fallback_name = all_text_types[0].get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            self.output.print_md("*Text type '{}' not found - using: {}*".format(
                text_type_name, fallback_name))
            return all_text_types[0]
    
        return None
    
    def run(self):
        """Execute with proper rollback on failure"""
    
        self.output.print_md("# DRAWING PACKAGE AUTOMATION")
        self.output.print_md("---")
        self.output.print_md("**Series:** {}".format(self.selected_series))
        self.output.print_md("**Package:** {}".format(self.package_name))
        self.output.print_md("")
    
        created_view_ids = []
        created_sheet_ids = []
        transaction_group = None
    
        try:
            # Use TransactionGroup for atomic rollback
            transaction_group = DB.TransactionGroup(
                self.doc,
                "Drawing Package - {}".format(self.selected_series)
            )
            transaction_group.Start()
        
            # Transaction 1: Create views and sheets
            with revit.Transaction("Drawing Package - Create All"):
                if not self.step2_create_views():
                    raise Exception("View creation failed")
                created_view_ids = [v['view'].Id for v in self.created_views]
            
                if not self.step3_rename_views():
                    raise Exception("View renaming failed")
            
                if not self.step4_apply_templates():
                    raise Exception("Template application failed")
            
                if not self.step6_create_sheets():
                    raise Exception("Sheet creation failed")
                created_sheet_ids = [s['sheet'].Id for s in self.created_sheets]
            
                if not self.step7_place_views():
                    raise Exception("View placement failed")
            
                # Regenerate inside transaction - CRITICAL FIX
                self.doc.Regenerate()
        
            self.output.print_md("**Views and sheets created**")
            self.output.print_md("")
        
            # Transaction 2: Tag elements
            with revit.Transaction("Drawing Package - Tag Elements"):
                if not self.step5_tag_elements():
                    raise Exception("Tagging failed")
        
            # Transaction 3: Legend (optional)
            if self.create_legend and self.legend_view:
                with revit.Transaction("Drawing Package - Legend"):
                    if not self.step9_place_legend():
                        raise Exception("Legend placement failed")

            # Transaction 4: Create Filter Legend (optional)
            if self.generate_filter_legends:
                with revit.Transaction("Drawing Package - Create Filter Legend"):
                    if not self.step10_generate_filter_legends():
                        raise Exception("Filter legend generation failed")
    
                # Transaction 5: Place Filter Legend on Sheets
                with revit.Transaction("Drawing Package - Place Filter Legend"):
                    # Regenerate INSIDE the transaction
                    self.doc.Regenerate()
        
                    if not self.step11_place_filter_legend_on_sheets():
                        raise Exception("Filter legend placement failed")
        
            # Commit all transactions
            transaction_group.Assimilate()
        
            self.print_summary()
            return True
        
        except Exception as e:
            # Rollback everything
            self.output.print_md("## ERROR: {}".format(str(e)))
        
            if transaction_group and transaction_group.HasStarted():
                self.output.print_md("*Rolling back all changes...*")
                try:
                    transaction_group.RollBack()
                    self.output.print_md("*Rollback successful*")
                except Exception as rollback_error:
                    self.output.print_md("*Rollback failed: {}*".format(str(rollback_error)))
                    self._emergency_cleanup(created_view_ids, created_sheet_ids)
            else:
                self._emergency_cleanup(created_view_ids, created_sheet_ids)
        
            import traceback
            self.output.print_html("<pre>{}</pre>".format(traceback.format_exc()))
            return False

    def _emergency_cleanup(self, view_ids, sheet_ids):
        """Manual cleanup when transaction group fails"""
        self.output.print_md("*Attempting emergency cleanup...*")
    
        deleted = 0
        try:
            with revit.Transaction("Emergency Cleanup"):
                for sheet_id in sheet_ids:
                    try:
                        self.doc.Delete(sheet_id)
                        deleted += 1
                    except:
                        pass
            
                for view_id in view_ids:
                    try:
                        self.doc.Delete(view_id)
                        deleted += 1
                    except:
                        pass
        except Exception as cleanup_error:
            self.output.print_md("*Cleanup error: {}*".format(str(cleanup_error)))
    
        self.output.print_md("*Cleaned up {} elements*".format(deleted))

    def _emergency_cleanup(self, view_ids, sheet_ids):
        """Manual cleanup when transaction group fails"""
        self.output.print_md("*Attempting emergency cleanup...*")
    
        deleted = 0
        try:
            with revit.Transaction("Emergency Cleanup"):
                for sheet_id in sheet_ids:
                    try:
                        self.doc.Delete(sheet_id)
                        deleted += 1
                    except:
                        pass
            
                for view_id in view_ids:
                    try:
                        self.doc.Delete(view_id)
                        deleted += 1
                    except:
                        pass
        except Exception as cleanup_error:
            self.output.print_md("*Cleanup error: {}*".format(str(cleanup_error)))
    
        self.output.print_md("*Cleaned up {} elements*".format(deleted))
    
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
        """Step 5: Tag elements"""
        
        self.output.print_md("## STEP 5: Tag Elements")
        
        # Force document regeneration
        self.doc.Regenerate()
        
        total_tagged = 0
        
        # Build tagging configuration
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
                            
                            # Get name to check
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
                        except (DB.Autodesk.Revit.Exceptions.ArgumentException,
                            DB.Autodesk.Revit.Exceptions.InvalidOperationException) as expected_ex:
                            # Expected failures (element not taggable, etc.)
                            failed_to_tag += 1
                            if failed_to_tag <= 3:
                                self.output.print_md("    Tag failed: {}".format(str(expected_ex)))
        
                        except Exception as unexpected_ex:
                            # Unexpected failures - log prominently
                            failed_to_tag += 1
                            self.output.print_md("    **UNEXPECTED ERROR: {}**".format(str(unexpected_ex)))
                            logger.error("Unexpected tag error on element {}".format(element.Id), exc_info=True)
    
                            # Abort if too many unexpected errors
                            if failed_to_tag > PackageConfig.MAX_TAG_FAILURES_PER_VIEW:
                                self.output.print_md("    **TOO MANY FAILURES - ABORTING TAGGING**")
                                break
                    except Exception as elem_ex:
                        pass
                
                # Report results
                if tagged_in_category > 0:
                    self.output.print_md("    **SUCCESS: Tagged {} {}**".format(
                        tagged_in_category, category_name
                    ))
                elif len(elements) > 0:
                    if skipped_no_match > 0:
                        self.output.print_md("    **INFO: 0 tagged** - {} {} didn't match type filter **'{}'**".format(
                            skipped_no_match, category_name, type_patterns[0] if type_patterns else 'none'
                        ))
                    if failed_to_tag > 0:
                        self.output.print_md("    **WARNING: 0 tagged** - {} {} failed tag creation".format(
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
        """Step 6: Create sheets with level numbering support"""
        
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
                
                import re
                
                # Check for basement levels (B1, B2, etc)
                basement_match = re.search(r'B(\d+)', level_name)
                if basement_match:
                    basement_num = int(basement_match.group(1))
                    level_suffix = str(100 - basement_num).zfill(3)
                else:
                    # Standard levels (00, 01, 02, etc)
                    level_match = re.search(r'(\d+)', level_name)
                    if level_match:
                        level_num = int(level_match.group(1))
                        level_suffix = str(100 + level_num + 1).zfill(3)
                    else:
                        # Fallback
                        level_suffix = current_sheet_number[-3:]
                
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

                # Validate sheet was created
                if sheet and sheet.Id != DB.ElementId.InvalidElementId:
                    self.created_sheets.append({
                        'sheet': sheet,
                        'view': view,
                        'number': sheet_number
                    })
    
                    self.output.print_md("- {} - {}".format(sheet_number, sheet_name))
                else:
                    self.output.print_md("- **ERROR:** Failed to create sheet {}".format(sheet_number))
                
                self.output.print_md("- {} - {}".format(sheet_number, sheet_name))
                
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
                legend_point = DB.XYZ(2.575459, 1.509186, 0)
                
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
        self.output.print_md("*Note: Legends are auto-centered. Adjust position manually if needed.*")
        self.output.print_md("")
        
        return True

    def step10_generate_filter_legends(self):
        """Step 10: Generate filter legends from templates"""
    
        if not self.generate_filter_legends:
            return True
    
        self.output.print_md("## STEP 10: Generate Filter Legends from Templates")
    
        try:
            # Import legend generator
            from legend_generator import generate_filter_legends
        
            # Get AUK standard text type using helper method
            text_type = self.get_text_type_by_name('AUK_Arial Bold_T_2.5mm')
        
            if not text_type:
                self.output.print_md("**WARNING:** No text types found - cannot create legends")
                return True
        
            # Filter views that have filters applied
            views_with_filters = []
            for view_data in self.created_views:
                view = view_data['view']
                if view.GetFilters():
                    views_with_filters.append(view)
        
            if not views_with_filters:
                self.output.print_md("**INFO:** No filters found in created views - skipping legend generation")
                self.output.print_md("*Make sure view templates have filters applied*")
                return True
        
            self.output.print_md("Found {} views with filters".format(len(views_with_filters)))
        
            # Use only the FIRST view to generate ONE legend
            representative_view = views_with_filters[0]
            self.output.print_md("*Using {} as representative view for legend*".format(
                representative_view.Name))
        
            # Generate single legend from first view
            created_legends = generate_filter_legends(
                self.doc,
                [representative_view],
                text_type
            )
        
            if created_legends:
                legend = created_legends[0]
            
                # Rename legend to series name (not view-specific)
                try:
                    legend.Name = 'Legend_{}_Filters'.format(self.selected_series)
                except:
                    pass  # Name conflict - keep default
            
                self.output.print_md("**Successfully created filter legend: {}**".format(legend.Name))
            
                # Store legend for placement on sheets
                self.generated_filter_legend = legend
            else:
                self.generated_filter_legend = None
        
            self.output.print_md("")
            return True
        
        except ImportError as import_error:
            self.output.print_md("**ERROR:** Cannot import legend_generator module")
            self.output.print_md("*Make sure legend_generator.py is in the same folder as this script*")
            self.output.print_md("*Error: {}*".format(str(import_error)))
            return True
        
        except Exception as e:
            self.output.print_md("**WARNING:** Legend generation failed: {}".format(str(e)))
            import traceback
            self.output.print_html("<pre>{}</pre>".format(traceback.format_exc()))
            return True
    
    def step11_place_filter_legend_on_sheets(self):
        """Step 11: Place generated filter legend on all sheets"""
    
        if not self.generate_filter_legends or not self.generated_filter_legend:
            return True
    
        self.output.print_md("## STEP 11: Place Filter Legend on All Sheets")
    
        # Get sheet numbers from created_sheets
        sheet_numbers = [data['number'] for data in self.created_sheets]
    
        # Re-fetch actual sheet objects from document
        all_sheets = DB.FilteredElementCollector(self.doc)\
            .OfCategory(DB.BuiltInCategory.OST_Sheets)\
            .WhereElementIsNotElementType()\
            .ToElements()
    
        # Filter to only our created sheets
        our_sheets = [s for s in all_sheets if s.SheetNumber in sheet_numbers]
    
        placed = 0
        failed = 0
    
        for sheet in our_sheets:
            try:
                # Place legend at origin
                legend_point = DB.XYZ(2.575459, 2.509186, 0)
            
                viewport = DB.Viewport.Create(
                    self.doc,
                    sheet.Id,
                    self.generated_filter_legend.Id,
                    legend_point
                )
            
                if viewport:
                    placed += 1
                else:
                    failed += 1
            
            except Exception as e:
                failed += 1
                self.output.print_md("- {} - Error: {}".format(sheet.SheetNumber, str(e)))
    
        self.output.print_md("**Placed filter legend on {} sheets**".format(placed))
        if failed > 0:
            self.output.print_md("**Failed: {}**".format(failed))
    
        self.output.print_md("*Legend placed at origin (0,0) - adjust position as needed*")
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
        # ==========================================================================
        # BATCH MODE - Simplified Version
        # ==========================================================================
    
        output.print_md("# BATCH MODE: DRAWING PACKAGE AUTOMATION")
        output.print_md("**Generating {} series**".format(len(window.batch_series)))
        output.print_md("---")
        output.print_md("")
    
        # Show progress window
        progress_win = ProgressWindow()
        progress_win.Show()
    
        # Track results
        successful = 0
        failed = 0
        total = len(window.batch_series)
    
        # Process each series
        for index, batch_item in enumerate(window.batch_series, start=1):
        
            # Extract data
            series_num = batch_item['series_num']
            series_config = batch_item['config']
            package_name = batch_item['package_name']
        
            # Update progress window
            progress_win.update_progress(
                index, 
                total, 
                package_name,
                "Processing..."
            )
        
            # Print header
            output.print_md("## SERIES {}: {}".format(series_num, package_name))
            output.print_md("")
        
            try:
                # Create automation instance
                automation = DrawingPackageAutomation(
                    doc=doc,
                    config=config,
                    selected_series=series_num,
                    series_config=series_config,
                    package_name=package_name,
                    legend_view=window.legend_view,
                    create_legend=window.create_legend,
                    generate_filter_legends=window.generate_filter_legends  # NEW
                )
            
                # Run automation
                success = automation.run()
            
                if success:
                    successful += 1
                    output.print_md("**SUCCESS**")
                else:
                    failed += 1
                    output.print_md("**FAILED**")
            
            except Exception as e:
                failed += 1
                output.print_md("**ERROR: {}**".format(str(e)))
                import traceback
                output.print_html("<pre>{}</pre>".format(traceback.format_exc()))
        
            output.print_md("---")
            output.print_md("")
    
        # Close progress window
        progress_win.Close()
    
        # Print summary
        output.print_md("# BATCH SUMMARY")
        output.print_md("**Total:** {} series".format(total))
        output.print_md("**Successful:** {}".format(successful))
        output.print_md("**Failed:** {}".format(failed))
        output.print_md("")
    
        # Alert user
        if successful == total:
            forms.alert(
                "All {} series completed successfully!".format(total),
                title="Batch Complete"
            )
        else:
            forms.alert(
                "{} of {} series completed.\n\nCheck output for details.".format(
                    successful, total
                ),
                title="Batch Complete"
            )

    else:
        # ==========================================================================
        # SINGLE MODE - Your existing code (keep as is)
        # ==========================================================================
    
        automation = DrawingPackageAutomation(
            doc, 
            config,
            window.selected_series,
            window.series_config,
            window.package_name,
            window.legend_view,
            window.create_legend,
            window.generate_filter_legends  # NEW
        )
    
        success = automation.run()
    
        if success:
            forms.alert("Drawing package automation completed!", title="Success")
        else:
            forms.alert("Check output for details", title="Warning")