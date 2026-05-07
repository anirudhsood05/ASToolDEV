# -*- coding: utf-8 -*-
"""Enhanced QA Check Recording System with Checkboxes - Standardized File Management."""

# Imports
from pyrevit import script, forms, revit, DB
import wpf
import clr
from System import Windows
from System.Windows import Controls
from System.IO import StringReader
import json
import os
import datetime
import getpass

# Import shared utilities (assume this is in the same directory)
# Note: In pyRevit, you might need to adjust the import path based on your structure
import sys
script_dir = os.path.dirname(__file__)
sys.path.append(script_dir)
try:
    from qa_file_utils import get_qa_data_file_path, migrate_qa_file_if_needed
except ImportError:
    # Fallback if shared utility is not available
    def get_qa_data_file_path():
        """Fallback implementation if shared utility not available."""
        doc = revit.doc
        if not doc:
            return None
        
        # Get project name
        project_name = None
        if doc.IsWorkshared:
            try:
                central_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(
                    doc.GetWorksharingCentralModelPath())
                project_name = os.path.splitext(os.path.basename(central_path))[0]
                project_dir = os.path.dirname(central_path)
            except:
                pass
        
        if not project_name and doc.PathName:
            project_name = os.path.splitext(os.path.basename(doc.PathName))[0]
            project_dir = os.path.dirname(doc.PathName)
        
        if not project_name:
            forms.alert("Please save the project before using QA tools.", 
                       title="Project Not Saved")
            return None
        
        # Create QA folder in model directory
        qa_folder = os.path.join(project_dir, "QA_Records")
        if not os.path.exists(qa_folder):
            try:
                os.makedirs(qa_folder)
            except:
                # Fallback to AppData
                appdata_path = os.path.expanduser("~\\AppData\\Roaming")
                qa_folder = os.path.join(appdata_path, "AukettSwanke", "QARecords")
                if not os.path.exists(qa_folder):
                    os.makedirs(qa_folder)
        
        filename = "{}_QAChecks.json".format(project_name)
        return os.path.join(qa_folder, filename)
    
    def migrate_qa_file_if_needed():
        """Placeholder for migration function."""
        return True

# Define XAML with improved layout and file path display
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK QA Check Recorder" Width="850" Height="750" ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize" MinWidth="600" MinHeight="600"
    FontFamily="Arial" FontSize="12" Background="#FFFFFF">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Area -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="AUK QA Check Recorder" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- File Info Panel -->
        <GroupBox Grid.Row="2" Header="File Information" Padding="5" Margin="0,0,0,8" Background="White">
            <StackPanel>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Label Grid.Column="0" Content="Data File:" FontWeight="Medium" VerticalAlignment="Center"/>
                    <TextBlock Grid.Column="1" x:Name="UI_file_path" 
                               Text="Initializing..." 
                               VerticalAlignment="Center" 
                               Margin="8,0,0,0"
                               TextWrapping="Wrap"
                               FontFamily="Consolas"
                               FontSize="10"/>
                </Grid>
                <TextBlock x:Name="UI_save_status" 
                           Text="Ready" 
                           Margin="0,5,0,0"
                           FontStyle="Italic"
                           FontSize="11"
                           Foreground="Green"/>
            </StackPanel>
        </GroupBox>

        <!-- Quick Actions -->
        <GroupBox Grid.Row="3" Header="Quick Actions" Padding="5" Margin="0,0,0,8" Background="White">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,5">
                <Button x:Name="UI_mark_daily" Content="✓ Complete Daily Tasks" 
                        Width="150" Height="28" Margin="0,0,8,0" 
                        Background="LightGreen" FontWeight="Medium"/>
                <Button x:Name="UI_mark_weekly" Content="✓ Complete Weekly Tasks" 
                        Width="150" Height="28" Margin="0,0,8,0"
                        Background="LightBlue" FontWeight="Medium"/>
                <Button x:Name="UI_generate_sheet" Content="📋 Generate Report" 
                        Width="130" Height="28" Margin="0,0,8,0"/>
                <Button x:Name="UI_save_close" Content="💾 Save &amp; Close" 
                        Width="120" Height="28"/>
            </StackPanel>
        </GroupBox>

        <!-- QA Checks ScrollViewer -->
        <ScrollViewer Grid.Row="4" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
            <StackPanel>
                
                <!-- Model Maintenance Section -->
                <GroupBox Header="Model Maintenance" Padding="8" Margin="0,0,0,8" Background="White">
                    <StackPanel>
                        
                        <!-- Daily Tasks -->
                        <Expander x:Name="UI_daily_expander" Header="Daily Tasks" IsExpanded="True" Margin="0,0,0,10">
                            <StackPanel Margin="20,5,0,0">
                                <CheckBox x:Name="UI_daily_audit" Content="Open with Audit checkbox ticked"
                                          Margin="0,3" FontSize="11"/>
                                <CheckBox x:Name="UI_daily_local" Content="Always create new local from the central model"
                                          Margin="0,3" FontSize="11"/>
                                <CheckBox x:Name="UI_daily_timestamp" Content="Append timestamp once per week; rest of days open as overwrite"
                                          Margin="0,3" FontSize="11"/>
                            </StackPanel>
                        </Expander>
                        
                        <!-- Weekly Tasks -->
                        <Expander x:Name="UI_weekly_expander" Header="Weekly Tasks" IsExpanded="True" Margin="0,0,0,10">
                            <StackPanel Margin="20,5,0,0">
                                <CheckBox x:Name="UI_weekly_warnings" Content="Review warnings and fix them in order of priority"
                                          Margin="0,3" FontSize="11"/>
                                <CheckBox x:Name="UI_weekly_rooms" Content="Delete redundant, unenclosed, not placed rooms"
                                          Margin="0,3" FontSize="11"/>
                                <CheckBox x:Name="UI_weekly_views" Content="Remove redundant views from the file"
                                          Margin="0,3" FontSize="11"/>
                                <TextBlock Text="  • Be mindful not to delete parent views that are not placed on sheets"
                                           Margin="20,0,0,3" FontSize="10" FontStyle="Italic" Foreground="Gray"/>
                                <CheckBox x:Name="UI_weekly_purge" Content="Purge to zero before compact sync"
                                          Margin="0,3" FontSize="11"/>
                            </StackPanel>
                        </Expander>
                        
                    </StackPanel>
                </GroupBox>
                
                <!-- Model Management Section -->
                <GroupBox Header="Model Management" Padding="8" Margin="0,0,0,8" Background="White">
                    <StackPanel>
                        
                        <!-- Design Options -->
                        <Expander x:Name="UI_design_expander" Header="Design Options" IsExpanded="True" Margin="0,0,0,10">
                            <StackPanel Margin="20,5,0,0">
                                <CheckBox x:Name="UI_design_remove" Content="Remove design options when no longer needed"
                                          Margin="0,3" FontSize="11"/>
                            </StackPanel>
                        </Expander>
                        
                        <!-- Groups and Families -->
                        <Expander x:Name="UI_groups_expander" Header="Groups &amp; Families" IsExpanded="True" Margin="0,0,0,10">
                            <StackPanel Margin="20,5,0,0">
                                <CheckBox x:Name="UI_groups_convert" Content="Convert frequent groups into families"
                                          Margin="0,3" FontSize="11"/>
                                <CheckBox x:Name="UI_families_size" Content="Identify and correct big sized families"
                                          Margin="0,3" FontSize="11"/>
                            </StackPanel>
                        </Expander>
                        
                    </StackPanel>
                </GroupBox>
                
                <!-- Comments Section -->
                <GroupBox Header="Additional Comments" Padding="8" Margin="0,0,0,8" Background="White">
                    <StackPanel>
                        <TextBlock Text="Add any additional notes about today's QA checks..." 
                                   FontStyle="Italic" 
                                   Foreground="Gray" 
                                   Margin="2,0,0,3"/>
                        <TextBox x:Name="UI_global_comments" 
                                 Height="60" 
                                 TextWrapping="Wrap" 
                                 AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto"/>
                    </StackPanel>
                </GroupBox>
                
            </StackPanel>
        </ScrollViewer>

        <!-- Button Area -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,5,0,0">
            <Button x:Name="UI_reset_btn" Content="🔄 Reset Form" 
                    Width="90" Height="25" Margin="0,0,8,0"/>
            <Button x:Name="UI_close_btn" Content="Close" 
                    Width="90" Height="25"/>
        </StackPanel>
    </Grid>
</Window>"""

# QA Check mapping
QA_CHECKS_MAP = {
    "Daily Tasks": {
        "UI_daily_audit": "Open with Audit checkbox ticked",
        "UI_daily_local": "Always create new local from the central model",
        "UI_daily_timestamp": "Append timestamp once per week; rest of days open as overwrite"
    },
    "Weekly Tasks": {
        "UI_weekly_warnings": "Review warnings and fix them in order of priority",
        "UI_weekly_rooms": "Delete redundant, unenclosed, not placed rooms",
        "UI_weekly_views": "Remove redundant views from the file",
        "UI_weekly_purge": "Purge to zero before compact sync"
    },
    "Design Options": {
        "UI_design_remove": "Remove design options when no longer needed"
    },
    "Groups & Families": {
        "UI_groups_convert": "Convert frequent groups into families",
        "UI_families_size": "Identify and correct big sized families"
    }
}


class QARecorderWindow(forms.WPFWindow):
    def __init__(self):
        # Load XAML
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
        
        # Initialize data file path using shared utility
        self.initialize_file_path()
        
        # Initialize
        self.initialize()
        
        # Connect events
        self.connect_events()
        
        # Load today's status
        self.load_todays_status()
    
    def initialize_file_path(self):
        """Initialize QA data file path and migrate if needed."""
        try:
            # Attempt to migrate old files if needed
            migrate_qa_file_if_needed()
            
            # Get standardized file path
            self.qa_data_file = get_qa_data_file_path()
            
            if self.qa_data_file:
                # Display file path in UI
                self.UI_file_path.Text = self.qa_data_file
                
                # Check if file exists
                if os.path.exists(self.qa_data_file):
                    self.update_save_status("File found - Ready to load", "Green")
                else:
                    self.update_save_status("New file will be created", "Orange")
            else:
                self.UI_file_path.Text = "ERROR: Could not determine file path"
                self.update_save_status("Error with file path", "Red")
                
        except Exception as e:
            self.UI_file_path.Text = "ERROR: {}".format(str(e))
            self.update_save_status("Error initializing file", "Red")
            forms.alert("Error initializing QA file: {}".format(str(e)), 
                       title="Initialization Error")
    
    def update_save_status(self, message, color="Green"):
        """Update the save status display."""
        self.UI_save_status.Text = message
        self.UI_save_status.Foreground = Windows.Media.Brushes.__getattr__(color)
    
    def initialize(self):
        """Initialize UI elements."""
        # Set expander styles
        for expander_name in ['UI_daily_expander', 'UI_weekly_expander', 
                             'UI_design_expander', 'UI_groups_expander']:
            expander = getattr(self, expander_name)
            expander.FontWeight = Windows.FontWeights.Medium
            expander.FontSize = 13
    
    def connect_events(self):
        """Connect UI events."""
        self.UI_mark_daily.Click += self.mark_daily_complete
        self.UI_mark_weekly.Click += self.mark_weekly_complete
        self.UI_generate_sheet.Click += self.generate_qa_sheet
        self.UI_save_close.Click += self.save_and_close
        self.UI_reset_btn.Click += self.reset_form
        self.UI_close_btn.Click += self.close_window
        
        # Connect checkbox events for auto-save
        for check_type, checkboxes in QA_CHECKS_MAP.items():
            for checkbox_name, task_name in checkboxes.items():
                checkbox = getattr(self, checkbox_name)
                checkbox.Checked += self.on_checkbox_changed
                checkbox.Unchecked += self.on_checkbox_changed
    
    def on_checkbox_changed(self, sender, args):
        """Handle checkbox state changes."""
        # Auto-save when checkbox state changes
        self.save_current_status()
        self.update_save_status("Auto-saved", "Green")
    
    def load_todays_status(self):
        """Load and apply today's checkbox status."""
        if not self.qa_data_file:
            return
            
        try:
            qa_data = self.load_qa_data()
            today = datetime.datetime.now().date()
            
            # Find today's records
            todays_records = []
            for r in qa_data:
                try:
                    # Parse timestamp manually for Python 2.7 compatibility
                    timestamp_str = r.get("timestamp", "")
                    if timestamp_str:
                        # Parse ISO format manually
                        date_part = timestamp_str.split('T')[0]
                        record_date = datetime.datetime.strptime(date_part, "%Y-%m-%d").date()
                        if record_date == today:
                            todays_records.append(r)
                except:
                    continue
            
            # Apply checkbox states based on today's records
            for record in todays_records:
                task = record.get("task", "")
                status = record.get("status", "")
                
                # Find corresponding checkbox
                for check_type, checkboxes in QA_CHECKS_MAP.items():
                    for checkbox_name, checkbox_task in checkboxes.items():
                        if checkbox_task == task:
                            checkbox = getattr(self, checkbox_name)
                            checkbox.IsChecked = (status == "Complete")
                            break
            
            # Load global comments from today's first record
            if todays_records:
                self.UI_global_comments.Text = todays_records[0].get("comments", "")
                self.update_save_status("Loaded today's data", "Green")
            else:
                self.update_save_status("No data for today", "Orange")
                
        except Exception as e:
            print("Error loading today's status: {}".format(str(e)))
            self.update_save_status("Error loading data", "Red")
    
    def save_current_status(self):
        """Save current checkbox status."""
        if not self.qa_data_file:
            return
            
        current_time = datetime.datetime.now()
        user = getpass.getuser()
        global_comments = self.UI_global_comments.Text
        
        # Get current checkbox states
        for check_type, checkboxes in QA_CHECKS_MAP.items():
            for checkbox_name, task_name in checkboxes.items():
                checkbox = getattr(self, checkbox_name)
                
                if checkbox.IsChecked is not None:
                    status = "Complete" if checkbox.IsChecked else "Incomplete"
                    comments = global_comments if global_comments else ""
                    
                    # Remove existing record for today
                    self.remove_todays_record(check_type, task_name)
                    
                    # Add new record only if checked or was previously saved
                    if checkbox.IsChecked or self.had_previous_record(check_type, task_name):
                        self.add_qa_record(check_type, task_name, status, user, current_time, comments)
    
    def had_previous_record(self, check_type, task):
        """Check if there was a previous record for this task today."""
        try:
            qa_data = self.load_qa_data()
            today = datetime.datetime.now().date()
            
            for record in qa_data:
                try:
                    # Parse timestamp manually for Python 2.7 compatibility
                    timestamp_str = record.get("timestamp", "")
                    if timestamp_str:
                        date_part = timestamp_str.split('T')[0]
                        record_date = datetime.datetime.strptime(date_part, "%Y-%m-%d").date()
                        if (record_date == today and 
                            record.get("check_type") == check_type and 
                            record.get("task") == task):
                            return True
                except:
                    continue
            return False
        except:
            return False
    
    def remove_todays_record(self, check_type, task):
        """Remove existing record for today."""
        try:
            qa_data = self.load_qa_data()
            today = datetime.datetime.now().date()
            
            # Filter out today's record for this specific task
            filtered_data = []
            for record in qa_data:
                try:
                    # Parse timestamp manually for Python 2.7 compatibility
                    timestamp_str = record.get("timestamp", "")
                    if timestamp_str:
                        date_part = timestamp_str.split('T')[0]
                        record_date = datetime.datetime.strptime(date_part, "%Y-%m-%d").date()
                        if not (record_date == today and 
                               record.get("check_type") == check_type and 
                               record.get("task") == task):
                            filtered_data.append(record)
                    else:
                        # Keep records with missing timestamps
                        filtered_data.append(record)
                except:
                    # Keep records with invalid dates
                    filtered_data.append(record)
            
            # Save filtered data
            with open(self.qa_data_file, 'w') as f:
                json.dump(filtered_data, f, indent=2)
        except Exception as e:
            print("Error removing today's record: {}".format(str(e)))
    
    def mark_daily_complete(self, sender, args):
        """Mark all daily tasks as complete."""
        for checkbox_name in QA_CHECKS_MAP["Daily Tasks"].keys():
            checkbox = getattr(self, checkbox_name)
            checkbox.IsChecked = True
        
        # Save status
        self.save_current_status()
        self.update_save_status("Daily tasks completed", "Green")
        forms.alert("All daily tasks marked as complete!", title="Daily Tasks Complete")
    
    def mark_weekly_complete(self, sender, args):
        """Mark all weekly tasks as complete."""
        for checkbox_name in QA_CHECKS_MAP["Weekly Tasks"].keys():
            checkbox = getattr(self, checkbox_name)
            checkbox.IsChecked = True
        
        # Save status
        self.save_current_status()
        self.update_save_status("Weekly tasks completed", "Green")
        forms.alert("All weekly tasks marked as complete!", title="Weekly Tasks Complete")
    
    def generate_qa_sheet(self, sender, args):
        """Generate QA report sheet."""
        # Save current status first
        self.save_current_status()
        
        # Check if we have any data
        qa_data = self.load_qa_data()
        if not qa_data:
            forms.alert("No QA data to report. Please complete some checks first.", 
                       title="No Data")
            return
        
        # Ask user if they want to keep the recorder open
        result = forms.alert("Generate QA report and close this window?", 
                           title="Generate Report", 
                           ok=False, 
                           cancel=True)
        
        if result:
            # Close this window
            self.Close()
            
            # Import and run the sheet generator
            try:
                # Execute the generate report script
                script_path = os.path.join(os.path.dirname(__file__), 
                                         '..', 'Generate Report.pushbutton', 'script.py')
                if os.path.exists(script_path):
                    execfile(script_path)
                else:
                    forms.alert("Generate Report script not found. Please ensure it's installed correctly.", 
                              title="Script Not Found")
            except Exception as e:
                forms.alert("Error generating report: {}".format(str(e)), title="Error")
    
    def save_and_close(self, sender, args):
        """Save current status and close."""
        self.save_current_status()
        self.update_save_status("All changes saved", "Green")
        forms.alert("QA checks saved successfully!", title="Saved")
        self.close_window(sender, args)
    
    def reset_form(self, sender, args):
        """Reset all checkboxes and comments."""
        result = forms.alert("Are you sure you want to reset all checkboxes? This will not delete saved records.", 
                           title="Reset Form", 
                           ok=False, 
                           cancel=True)
        
        if result:
            # Reset all checkboxes
            for check_type, checkboxes in QA_CHECKS_MAP.items():
                for checkbox_name in checkboxes.keys():
                    checkbox = getattr(self, checkbox_name)
                    checkbox.IsChecked = False
            
            # Clear comments
            self.UI_global_comments.Text = ""
            self.update_save_status("Form reset", "Orange")
    
    def add_qa_record(self, check_type, task, status, user, timestamp, comments):
        """Add a QA record to the data file."""
        qa_data = self.load_qa_data()
        
        record = {
            "check_type": check_type,
            "task": task,
            "status": status,
            "user": user,
            "timestamp": timestamp.isoformat(),
            "comments": comments
        }
        
        qa_data.append(record)
        
        # Save to file
        try:
            with open(self.qa_data_file, 'w') as f:
                json.dump(qa_data, f, indent=2)
        except Exception as e:
            self.update_save_status("Error saving", "Red")
            forms.alert("Error saving QA data: {}".format(str(e)), title="Save Error")
    
    def load_qa_data(self):
        """Load QA data from file."""
        if not self.qa_data_file:
            return []
            
        if os.path.exists(self.qa_data_file):
            try:
                with open(self.qa_data_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print("Error loading QA data: {}".format(str(e)))
                return []
        return []
    
    def close_window(self, sender, args):
        """Close the window."""
        self.Close()


# Run the tool
if __name__ == '__main__':
    QARecorderWindow().ShowDialog()