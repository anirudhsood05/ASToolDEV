# -*- coding: utf-8 -*-
"""AUK Technical Guidance Panel - Context-aware guidance system for Revit workflows."""

__title__ = "Guidance\nPanel"
__author__ = "Anirudh Sood"
__min_revit_ver__ = 2020
__max_revit_ver__ = 2025

from pyrevit import script, forms, revit, DB, UI
import wpf
import clr
from System.IO import StringReader
from System.Windows import Visibility, Thickness, HorizontalAlignment, VerticalAlignment, CornerRadius, FontWeights, TextWrapping
from System.Windows.Media import Brushes, Color
from System.Windows.Controls import ScrollBarVisibility, Border, TextBlock, Button, StackPanel, Orientation
from System.Windows.Interop import WindowInteropHelper
import os
import csv

# Configuration
GUIDANCE_CSV_PATH = r"I:\002_BIM\007_Addins\Aukett Swanke\Extensions\ASToolDEV\AS_ToolsDev.extension\AS_ToolsDEV.tab\X_Development.Panel\DRAWING Package Guidance.pushbutton\guidance_metadata.csv"
GUIDANCE_PDF_FOLDER = r"I:\006_Technical\Drawing Package Guidance\Standard_Packages"

# XAML Definition
xaml_file = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Technical Guidance" Width="400" Height="600" ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize" MinWidth="350" MinHeight="400"
    FontFamily="Arial" FontSize="12" Background="#FFFFFF" Topmost="True">
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Area -->
        <Border Grid.Row="0" Background="#F0F0F0" Padding="8" CornerRadius="3">
            <TextBlock Text="Aukett Swanke Technical Guidance" 
                       FontWeight="Bold" FontSize="14" 
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- Results Section -->
        <GroupBox Grid.Row="2" Header="Guidance Topics" Padding="8" Background="White">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel Name="UI_results_panel">
                    <TextBlock Name="UI_context_label" FontStyle="Italic" FontSize="11" 
                               Foreground="#666666" Margin="0,0,0,8" TextWrapping="Wrap"/>
                    <TextBlock Name="UI_no_results" Text="Loading guidance database..." 
                               FontStyle="Italic" Foreground="#999999" TextWrapping="Wrap"/>
                </StackPanel>
            </ScrollViewer>
        </GroupBox>

        <!-- Footer -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="UI_close_btn" Content="Close" Width="70" Height="25"/>
        </StackPanel>
    </Grid>
</Window>"""

def escape_xml(text):
        """Escape XML special characters for XAML safety"""
        if not text:
            return ''
        return str(text).replace('&', '&amp;') \
                        .replace('<', '&lt;') \
                        .replace('>', '&gt;') \
                        .replace('"', '&quot;') \
                        .replace("'", '&apos;')
class SecureCSVConfig:
    """Security validation for CSV configuration"""
    MAX_FILE_SIZE = 1048576  # 1 MB
    MAX_ROWS = 100
    MAX_FIELD_LENGTH = 500
    
    @staticmethod
    def sanitize_field(field_name, value):
        """Strip dangerous prefixes and validate"""
        if value is None:
            return ''
        value = str(value).strip()
        
        # Remove CSV injection prefixes
        dangerous_prefixes = ['=', '+', '-', '@', '\t', '\r', '\n']
        while value and value[0] in dangerous_prefixes:
            value = value[1:]
        
        # Length limit
        if len(value) > SecureCSVConfig.MAX_FIELD_LENGTH:
            value = value[:SecureCSVConfig.MAX_FIELD_LENGTH]
        
        return value


class GuidanceDatabase:
    """Manages guidance content loading and filtering."""
    
    def __init__(self):
        self.guidance_items = []
        self.last_load_time = None
    
    def load_guidance(self):
        """Load guidance from CSV with security validation."""
        items = []
    
        if not os.path.exists(GUIDANCE_CSV_PATH):
            print("CSV not found: {}".format(GUIDANCE_CSV_PATH))
            return items
    
        try:
            # File size check
            file_size = os.path.getsize(GUIDANCE_CSV_PATH)
            if file_size > SecureCSVConfig.MAX_FILE_SIZE:
                print("CSV too large: {} bytes (max: {})".format(
                    file_size, SecureCSVConfig.MAX_FILE_SIZE))
                return items
        
            row_count = 0
            with open(GUIDANCE_CSV_PATH, 'r') as csvfile:
                reader = csv.DictReader(csvfile)
            
                for row in reader:
                    row_count += 1
                    if row_count > SecureCSVConfig.MAX_ROWS:
                        print("CSV exceeds {} rows - stopping".format(SecureCSVConfig.MAX_ROWS))
                        break
                
                    # Sanitize all fields
                    sanitized_row = {
                        k: SecureCSVConfig.sanitize_field(k, v) 
                        for k, v in row.items()
                    }
                
                    # Build PDF path
                    pdf_file = sanitized_row.get('pdf_file', '').strip()
                    if pdf_file:
                        pdf_path = os.path.join(GUIDANCE_PDF_FOLDER, pdf_file)
                        sanitized_row['pdf_path'] = pdf_path if os.path.exists(pdf_path) else None
                    else:
                        sanitized_row['pdf_path'] = None
                
                    # Parse multi-value fields
                    sanitized_row['context_types'] = [
                        ct.strip() for ct in sanitized_row.get('context_type', '').split(';') if ct.strip()
                    ]
                    sanitized_row['related_list'] = [
                        r.strip() for r in sanitized_row.get('related_packages', '').split(';') if r.strip()
                    ]
                    sanitized_row['nbs_list'] = [
                        n.strip() for n in sanitized_row.get('nbs_specs', '').split(';') if n.strip()
                    ]
                
                    # Parse RIBA stage range
                    stage_str = sanitized_row.get('riba_stage', '').strip()
                    if '-' in stage_str and stage_str != 'All':
                        try:
                            start, end = stage_str.split('-')
                            sanitized_row['stage_list'] = [str(s) for s in range(int(start), int(end) + 1)]
                        except:
                            sanitized_row['stage_list'] = [stage_str]
                    else:
                        sanitized_row['stage_list'] = [stage_str]
                
                    items.append(sanitized_row)
        
            self.guidance_items = items
            
            return items
        
        except Exception as e:
            print("Error loading guidance: {}".format(str(e)))
            return items
    
    def filter_by_context(self, view_type=None, keywords=None):
        """Filter guidance by context or keywords."""
        if not self.guidance_items:
            return []
        
        filtered = self.guidance_items
        
        # Filter by view type context
        if view_type:
            view_type_str = str(view_type)
            filtered = [item for item in filtered 
                       if any(view_type_str.lower() in ct.lower() for ct in item.get('context_types', []))
                       or 'all' in [ct.lower() for ct in item.get('context_types', [])]]
        
        # Filter by keywords - search across multiple fields
        if keywords:
            keywords_lower = keywords.lower()
            filtered = [item for item in filtered
                       if any(keywords_lower in str(v).lower() 
                             for k, v in item.items() 
                             if k in ['title', 'keywords', 'brief_summary', 'detailed_summary', 
                                     'cisfb_code', 'discipline', 'nbs_specs', 'related_packages'])]
        
        return filtered
    
    def get_all(self):
        """Return all guidance items."""
        return self.guidance_items


class GuidancePanelWindow(forms.WPFWindow):
    """Main guidance panel window."""
    
    def __init__(self):
        # Load XAML
        xaml_stream = StringReader(xaml_file)
        wpf.LoadComponent(self, xaml_stream)
    
        # Initialize
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.database = GuidanceDatabase()
    
        # Load guidance ONCE on startup
        try:
            self.database.load_guidance()
        
            if not self.database.guidance_items:
                self.UI_no_results.Text = (
                    "No guidance items loaded.\n\n"
                    "Check CSV path:\n{}\n\n"
                    "Close and reopen panel to reload.".format(GUIDANCE_CSV_PATH)
                )
        except Exception as e:
            print("Error loading guidance: {}".format(str(e)))
            self.UI_no_results.Text = (
                "Error loading guidance database.\n"
                "See output window for details."
            )
    
        # Connect events and display
        self.connect_events()
        self.display_all_guidance()
        self.update_context_label()
    
    
    def connect_events(self):
        """Connect UI events."""
        self.UI_close_btn.Click += self.on_close
    
    def refresh_database(self):
        """Reload guidance database from CSV."""
        self.database.load_guidance()
    
    def update_context_label(self):
        """Update context label based on active view."""
        try:
            active_view = self.uidoc.ActiveView
            view_type = active_view.ViewType
            view_name = active_view.Name
            
            context_text = "Current context: {} - {}".format(
                str(view_type).replace("ViewType.", ""), 
                view_name
            )
            self.UI_context_label.Text = context_text
            self.UI_context_label.Visibility = Visibility.Visible
            
        except:
            self.UI_context_label.Visibility = Visibility.Collapsed
    
    def display_all_guidance(self):
        """Display all guidance items."""
        items = self.database.get_all()
        self.display_guidance_items(items)
    
    def display_guidance_items(self, items):
        """Display guidance items in panel."""
        # Clear existing
        self.UI_results_panel.Children.Clear()
        
        # Re-add context label
        self.UI_results_panel.Children.Add(self.UI_context_label)
        
        if not items:
            self.UI_no_results.Text = "No guidance found. Check CSV path or search terms."
            self.UI_no_results.Visibility = Visibility.Visible
            self.UI_results_panel.Children.Add(self.UI_no_results)
            return
        
        self.UI_no_results.Visibility = Visibility.Collapsed
        
        # Create UI for each item
        for item in items:
            self.create_guidance_item_ui(item)
    
    def create_guidance_item_ui(self, item):
        """Create UI elements with XML escaping"""
        
        # Container
        container = Border()
        container.Background = Brushes.WhiteSmoke
        container.CornerRadius = CornerRadius(3)
        container.Padding = Thickness(10)
        container.Margin = Thickness(0, 0, 0, 8)
        
        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        
        # Title with CI/SfB code
        title_stack = StackPanel()
        title_stack.Orientation = Orientation.Horizontal
        
        # CI/SfB code badge
        cisfb_code = item.get('cisfb_code', '').strip()
        if cisfb_code:
            code_label = TextBlock()
            code_label.Text = cisfb_code
            code_label.FontWeight = FontWeights.Bold
            code_label.FontSize = 11
            code_label.Foreground = Brushes.White
            code_label.Background = Brushes.DarkSlateGray
            code_label.Padding = Thickness(5, 2, 5, 2)
            code_label.Margin = Thickness(0, 0, 8, 0)
            title_stack.Children.Add(code_label)
        
        # Title
        title = TextBlock()
        title.Text = escape_xml(item.get('title', 'Untitled'))
        title.FontWeight = FontWeights.Bold
        title.FontSize = 13
        title.TextWrapping = TextWrapping.Wrap
        title_stack.Children.Add(title)
        
        stack.Children.Add(title_stack)
        
        # Metadata row: Discipline and RIBA Stage
        meta_stack = StackPanel()
        meta_stack.Orientation = Orientation.Horizontal
        meta_stack.Margin = Thickness(0, 5, 0, 5)
        
        discipline = item.get('discipline', '').strip()
        riba_stage = item.get('riba_stage', '').strip()
        
        if discipline or riba_stage:
            meta_text = TextBlock()
            meta_parts = []
            if discipline:
                meta_parts.append("Discipline: {}".format(discipline))
            if riba_stage:
                meta_parts.append("RIBA Stage: {}".format(riba_stage))
            meta_text.Text = escape_xml(" | ".join(meta_parts))
            meta_text.FontSize = 10
            meta_text.Foreground = Brushes.DarkBlue
            meta_stack.Children.Add(meta_text)
        
        stack.Children.Add(meta_stack)
        
        # Brief summary
        brief_summary = item.get('brief_summary', item.get('summary', '')).strip()
        if brief_summary:
            summary_text = TextBlock()
            summary_text.Text = escape_xml(brief_summary)
            summary_text.FontSize = 11
            summary_text.TextWrapping = TextWrapping.Wrap
            summary_text.Margin = Thickness(0, 0, 0, 8)
            stack.Children.Add(summary_text)
        
        # Related packages (if any)
        related_list = item.get('related_list', [])
        if related_list:
            related_label = TextBlock()
            related_label.Text = escape_xml("Related: {}".format(", ".join(related_list)))
            related_label.FontSize = 10
            related_label.Foreground = Brushes.DarkGreen
            related_label.TextWrapping = TextWrapping.Wrap
            related_label.Margin = Thickness(0, 0, 0, 3)
            stack.Children.Add(related_label)
        
        # NBS specs (if any)
        nbs_list = item.get('nbs_list', [])
        if nbs_list:
            nbs_label = TextBlock()
            nbs_label.Text = escape_xml("NBS: {}".format(", ".join(nbs_list)))
            nbs_label.FontSize = 10
            nbs_label.Foreground = Brushes.DarkOrange
            nbs_label.Margin = Thickness(0, 0, 0, 8)
            stack.Children.Add(nbs_label)
        
        # Button container
        btn_stack = StackPanel()
        btn_stack.Orientation = Orientation.Horizontal
        
        # Open PDF button - always show, but indicate if missing
        pdf_path = item.get('pdf_path')
        pdf_file = item.get('pdf_file', '').strip()
        
        if pdf_file:
            # Button creation - verify this section
            pdf_btn = Button()
            if pdf_path:
                pdf_btn.Content = "View Guide (PDF)"
                pdf_btn.Tag = pdf_path  # <-- Path must be stored here
                pdf_btn.Click += self.on_open_pdf  # <-- Must use +=
            else:
                pdf_btn.Content = "PDF Not Found"
                pdf_btn.IsEnabled = False
            
            pdf_btn.Height = 25
            pdf_btn.Padding = Thickness(10, 0, 10, 0)
            btn_stack.Children.Add(pdf_btn)
        
        stack.Children.Add(btn_stack)
        container.Child = stack
        
        self.UI_results_panel.Children.Add(container)
    
    def on_open_pdf(self, sender, args):
        """Open PDF with improved error handling"""
        import os  # <-- ADD THIS LINE
    
        pdf_path = sender.Tag
    
        if not pdf_path:
            forms.alert("No PDF linked to this guidance item.", title="No PDF")
            return
    
        try:
            # Validate path still exists
            if not os.path.exists(pdf_path):
                forms.alert(
                    "PDF file not found:\n\n{}\n\n"
                    "It may have been moved or deleted.".format(pdf_path), 
                    title="File Not Found"
                )
                return
        
            # Validate it's actually a PDF
            if not pdf_path.lower().endswith('.pdf'):
                forms.alert(
                    "File is not a PDF:\n{}".format(pdf_path),
                    title="Invalid File Type"
                )
                return
        
            # Open with default application
            os.startfile(pdf_path)
        
        except WindowsError as e:
            forms.alert(
                "Windows error opening PDF:\n{}\n\nPath: {}".format(str(e), pdf_path), 
                title="Windows Error"
            )
        except Exception as e:
            forms.alert(
                "Error opening PDF:\n{}\n\nPath: {}".format(str(e), pdf_path), 
                title="Error"
            )
    
    
    
      
    def on_close(self, sender, args):
        """Close window."""
        self.Close()


# Run the guidance panel
if __name__ == '__main__':
    try:
        # Create modeless window
        panel = GuidancePanelWindow()
        
        # Set Revit as owner window (enables keyboard input in modeless)
        try:
            from Autodesk.Revit.UI import UIApplication
            helper = WindowInteropHelper(panel)
            uiapp = UIApplication(revit.HOST_APP.app)
            helper.Owner = uiapp.MainWindowHandle
        except Exception as e:
            print("Could not set window owner: {}".format(str(e)))
            # Window will still open but keyboard might not work
        
        # Show modeless (allows Revit interaction)
        panel.Show()
        
    except Exception as e:
        forms.alert("Error launching guidance panel:\n{}".format(str(e)), 
                   title="Guidance Panel Error")
