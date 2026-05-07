# -*- coding: utf-8 -*-
"""
Script:   Tagless Shamelist
Desc:     Identifies untagged Walls, Floors, and Ceilings across selected views.
          Double-click a result row to navigate to the element in its view.
Author:   Anirudh Sood @ Aukett Swanke
Usage:    Run tool, filter/select views, click Analyse. Double-click rows to navigate.
Result:   A filterable list of all untagged elements across the selected views.
"""

__title__ = "AUK Tagless Shamelist"
__doc__ = "Identifies untagged Walls, Floors and Ceilings across views. Double-click to navigate."
__author__ = "Anirudh Sood"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from pyrevit import revit, DB, script, forms
from System.IO import StringReader
from System.Windows import Window, Visibility, Thickness, GridLength, GridUnitType
from System.Windows import HorizontalAlignment, VerticalAlignment, TextTrimming
from System.Windows.Controls import CheckBox, StackPanel, Border, TextBlock, ScrollViewer
from System.Windows.Controls import ColumnDefinition as WPFColumnDefinition
from System.Windows.Controls import Grid as WPFGrid
from System.Windows.Input import Keyboard, ModifierKeys
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Threading import Dispatcher, DispatcherPriority
import System
import wpf

logger = script.get_logger()
doc    = revit.doc
uidoc  = revit.uidoc


# ── Constants ──────────────────────────────────────────────────────────────────
MAX_VIEWS_WARNING = 100
MAX_RESULTS_SHOWN = 5000

ROW_BG_ODD     = "#FFFFFF"
ROW_BG_EVEN    = "#F7F9FB"
ROW_BG_CHECKED = "#D6E8F5"
ROW_BG_HOVER   = "#EAF3FB"
ROW_BORDER_CLR = "#EEEEEE"


# ── Compatibility helper ───────────────────────────────────────────────────────
def eid_int(element_id):
    """Return integer value of an ElementId — Revit 2023-2026 safe."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


INVALID_ID_INT = eid_int(DB.ElementId.InvalidElementId)


# ── Brush factory ──────────────────────────────────────────────────────────────
def _brush(hex_colour):
    """Return a SolidColorBrush from a #RRGGBB string."""
    h = hex_colour.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return SolidColorBrush(Color.FromRgb(r, g, b))


# ── XAML ───────────────────────────────────────────────────────────────────────
# The view list uses a ScrollViewer+StackPanel built in Python code (not a
# ListView DataTemplate) so that Keyboard.Modifiers is reliably accessible
# inside CheckBox.Click handlers — the only reliable way to detect Shift in
# IronPython WPF DataTemplates.
XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Tagless Shamelist"
    Width="920" Height="860"
    MinWidth="760" MinHeight="640"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize"
    ShowInTaskbar="False"
    FontFamily="Arial" FontSize="12"
    Background="#F5F5F5">

    <Window.Resources>

        <Style TargetType="Button" x:Key="AUKButton">
            <Setter Property="Background" Value="#143D5C"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#4A7A8A"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="Button" x:Key="CancelButton">
            <Setter Property="Background" Value="#CCCCCC"/>
            <Setter Property="Foreground" Value="#2C2C2C"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#BBBBBB"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="Button" x:Key="StopButton">
            <Setter Property="Background" Value="#E8A735"/>
            <Setter Property="Foreground" Value="#2C2C2C"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#D4972B"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="CheckBox" x:Key="CatCheckBox">
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,0,16,0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#2C2C2C"/>
        </Style>

        <Style TargetType="ListViewItem">
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Padding" Value="2"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#D6E4F7"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#B3CDE0"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ComboBox" x:Key="FilterCombo">
            <Setter Property="Height" Value="24"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="4,0"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>

    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="2*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="6"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Bar -->
        <Border Grid.Row="0" Background="#143D5C" Padding="12,10" CornerRadius="4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Vertical">
                    <TextBlock Text="AUK Tagless Shamelist" Foreground="#FFFFFF"
                               FontWeight="Bold" FontSize="15"/>
                    <TextBlock Text="Identify untagged Walls, Floors and Ceilings across views"
                               Foreground="#AECBD6" FontSize="10" Margin="0,2,0,0"/>
                </StackPanel>
                <Border Grid.Column="1" Background="#4A7A8A" CornerRadius="3" Padding="8,4">
                    <TextBlock x:Name="UI_result_badge" Text="--" Foreground="#FFFFFF"
                               FontSize="11" FontWeight="Bold" VerticalAlignment="Center"/>
                </Border>
            </Grid>
        </Border>

        <!-- View Selection Panel -->
        <GroupBox Grid.Row="2" Header=" View Selection  (Shift-click to range select) "
                  Padding="8" Background="#FFFFFF"
                  BorderBrush="#CCCCCC" Foreground="#2C2C2C" FontWeight="SemiBold">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="4"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="210"/>
                </Grid.RowDefinitions>

                <!-- Search + on-sheet + action buttons -->
                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Search:" VerticalAlignment="Center"
                               FontWeight="SemiBold" Foreground="#2C2C2C" Margin="0,0,8,0"/>
                    <TextBox Grid.Column="1" x:Name="UI_view_filter" Height="24"
                             VerticalContentAlignment="Center" Padding="4,0"
                             Background="#FFFFFF" Foreground="#2C2C2C"
                             BorderBrush="#CCCCCC" BorderThickness="1"/>
                    <CheckBox Grid.Column="2" x:Name="UI_chk_onsheet" Content="On Sheet Only"
                              VerticalAlignment="Center" Foreground="#2C2C2C"/>
                    <Button Grid.Column="4" x:Name="UI_btn_selall"      Content="Select All"
                            Style="{StaticResource AUKButton}" Width="85"/>
                    <Button Grid.Column="6" x:Name="UI_btn_selfiltered" Content="Sel. Filtered"
                            Style="{StaticResource AUKButton}" Width="90"/>
                    <Button Grid.Column="8" x:Name="UI_btn_selclear"    Content="Clear"
                            Style="{StaticResource CancelButton}" Width="60"/>
                </Grid>

                <!-- View type checkboxes -->
                <StackPanel Grid.Row="2" Orientation="Horizontal">
                    <TextBlock Text="View Types:" VerticalAlignment="Center"
                               FontWeight="SemiBold" Foreground="#2C2C2C" Margin="0,0,8,0"/>
                    <CheckBox x:Name="UI_chk_floorplan"   Content="Floor Plans"
                              Style="{StaticResource CatCheckBox}" IsChecked="True"/>
                    <CheckBox x:Name="UI_chk_ceilingplan" Content="Ceiling Plans"
                              Style="{StaticResource CatCheckBox}" IsChecked="True"/>
                    <CheckBox x:Name="UI_chk_section"     Content="Sections"
                              Style="{StaticResource CatCheckBox}" IsChecked="False"/>
                    <CheckBox x:Name="UI_chk_elevation"   Content="Elevations"
                              Style="{StaticResource CatCheckBox}" IsChecked="False"/>
                    <CheckBox x:Name="UI_chk_3d"          Content="3D Views"
                              Style="{StaticResource CatCheckBox}" IsChecked="False"/>
                </StackPanel>

                <!-- Scale + Template filters -->
                <Grid Grid.Row="4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="130"/>
                        <ColumnDefinition Width="16"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="180"/>
                        <ColumnDefinition Width="16"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Scale:" VerticalAlignment="Center"
                               FontWeight="SemiBold" Foreground="#2C2C2C" Margin="0,0,6,0"/>
                    <ComboBox Grid.Column="1" x:Name="UI_scale_filter"
                              Style="{StaticResource FilterCombo}"/>
                    <TextBlock Grid.Column="3" Text="Template:" VerticalAlignment="Center"
                               FontWeight="SemiBold" Foreground="#2C2C2C" Margin="0,0,6,0"/>
                    <ComboBox Grid.Column="4" x:Name="UI_template_filter"
                              Style="{StaticResource FilterCombo}"/>
                    <CheckBox Grid.Column="6" x:Name="UI_chk_no_template"
                              Content="No Template Only"
                              VerticalAlignment="Center" Foreground="#2C2C2C"/>
                </Grid>

                <!-- Column headers (manually positioned to match row grid columns) -->
                <Border Grid.Row="6" Background="#E8E8E8"
                        BorderBrush="#CCCCCC" BorderThickness="1,1,1,0">
                    <Grid Margin="2,2,2,2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="26"/>
                            <ColumnDefinition Width="290"/>
                            <ColumnDefinition Width="105"/>
                            <ColumnDefinition Width="75"/>
                            <ColumnDefinition Width="155"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="1" Text="View Name"
                                   FontWeight="SemiBold" FontSize="11" Foreground="#2C2C2C"/>
                        <TextBlock Grid.Column="2" Text="Type"
                                   FontWeight="SemiBold" FontSize="11" Foreground="#2C2C2C"/>
                        <TextBlock Grid.Column="3" Text="Scale"
                                   FontWeight="SemiBold" FontSize="11" Foreground="#2C2C2C"/>
                        <TextBlock Grid.Column="4" Text="Template"
                                   FontWeight="SemiBold" FontSize="11" Foreground="#2C2C2C"/>
                        <TextBlock Grid.Column="5" Text="Sheet"
                                   FontWeight="SemiBold" FontSize="11" Foreground="#2C2C2C"/>
                    </Grid>
                </Border>

                <!-- Scrollable StackPanel — rows built entirely in Python -->
                <ScrollViewer Grid.Row="7" x:Name="UI_view_scroll"
                              VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Disabled"
                              BorderBrush="#CCCCCC" BorderThickness="1,0,1,1"
                              Background="#FFFFFF">
                    <StackPanel x:Name="UI_view_panel" Background="#FFFFFF"/>
                </ScrollViewer>
            </Grid>
        </GroupBox>

        <!-- Category Filter -->
        <GroupBox Grid.Row="4" Header=" Categories to Check " Padding="8" Background="#FFFFFF"
                  BorderBrush="#CCCCCC" Foreground="#2C2C2C" FontWeight="SemiBold">
            <StackPanel Orientation="Horizontal">
                <CheckBox x:Name="UI_cat_walls"    Content="Walls"
                          Style="{StaticResource CatCheckBox}" IsChecked="True"/>
                <CheckBox x:Name="UI_cat_floors"   Content="Floors"
                          Style="{StaticResource CatCheckBox}" IsChecked="True"/>
                <CheckBox x:Name="UI_cat_ceilings" Content="Ceilings"
                          Style="{StaticResource CatCheckBox}" IsChecked="True"/>
            </StackPanel>
        </GroupBox>

        <!-- Results -->
        <GroupBox Grid.Row="6" Header=" Results  --  double-click a row to navigate "
                  Padding="6" Background="#FFFFFF" BorderBrush="#CCCCCC"
                  Foreground="#2C2C2C" FontWeight="SemiBold">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="200"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Filter Results:" VerticalAlignment="Center"
                               FontWeight="SemiBold" Foreground="#2C2C2C" Margin="0,0,8,0"/>
                    <TextBox Grid.Column="1" x:Name="UI_result_filter" Height="24"
                             VerticalContentAlignment="Center" Padding="4,0"
                             Background="#FFFFFF" Foreground="#2C2C2C"
                             BorderBrush="#CCCCCC" BorderThickness="1"/>
                    <TextBlock Grid.Column="2" x:Name="UI_status_text" Text=""
                               VerticalAlignment="Center" Foreground="#666666"
                               FontSize="11" Margin="12,0,0,0"/>
                </Grid>

                <ListView Grid.Row="2" x:Name="UI_results_list"
                          BorderBrush="#CCCCCC" BorderThickness="1" Background="#FFFFFF"
                          VirtualizingStackPanel.IsVirtualizing="True">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Header="Category"      Width="90"
                                            DisplayMemberBinding="{Binding Category}"/>
                            <GridViewColumn Header="Element ID"    Width="90"
                                            DisplayMemberBinding="{Binding ElementId}"/>
                            <GridViewColumn Header="Type / Family" Width="220"
                                            DisplayMemberBinding="{Binding TypeName}"/>
                            <GridViewColumn Header="Level"         Width="110"
                                            DisplayMemberBinding="{Binding Level}"/>
                            <GridViewColumn Header="View"          Width="230"
                                            DisplayMemberBinding="{Binding ViewName}"/>
                        </GridView>
                    </ListView.View>
                </ListView>
            </Grid>
        </GroupBox>

        <!-- Progress Bar -->
        <ProgressBar Grid.Row="8" x:Name="UI_progress_bar"
                     Height="6" Minimum="0" Maximum="100" Value="0"
                     Foreground="#143D5C" Background="#EEEEEE"
                     BorderThickness="0" Visibility="Collapsed"/>

        <!-- Footer -->
        <Grid Grid.Row="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" x:Name="UI_footer_text" Text=""
                       VerticalAlignment="Center" FontSize="11" Foreground="#666666"/>
            <Button Grid.Column="1" x:Name="UI_btn_analyse" Content="Analyse"
                    Style="{StaticResource AUKButton}" Width="100"/>
            <Button Grid.Column="3" x:Name="UI_btn_cancel_analysis" Content="Cancel"
                    Style="{StaticResource StopButton}" Width="80" Visibility="Collapsed"/>
            <Button Grid.Column="5" x:Name="UI_btn_close" Content="Close"
                    Style="{StaticResource CancelButton}" Width="80"/>
        </Grid>
    </Grid>
</Window>
"""


# ── Data models ────────────────────────────────────────────────────────────────

class ViewItem(object):
    """Represents one view in the selection list."""

    def __init__(self, view):
        self._view        = view
        self.IsSelected   = True
        self.ViewName     = view.Name or "(unnamed)"
        self.ViewType     = str(view.ViewType)
        self.ScaleLabel   = self._get_scale_label(view)
        self.TemplateName = self._get_template_name(view)
        self.SheetInfo    = self._get_sheet_info(view)
        # Set by _build_view_row; used for visual updates and shift-range
        self.display_index = 0
        self.row_border    = None
        self.row_checkbox  = None

    @staticmethod
    def _get_scale_label(view):
        try:
            s = view.Scale
            return "1:{}".format(s) if s and s > 0 else "-"
        except Exception:
            return "-"

    @staticmethod
    def _get_template_name(view):
        try:
            tid = view.ViewTemplateId
            if tid is not None and eid_int(tid) != INVALID_ID_INT:
                tmpl = doc.GetElement(tid)
                if tmpl:
                    return tmpl.Name or "(unnamed template)"
        except Exception:
            pass
        return "(none)"

    @staticmethod
    def _get_sheet_info(view):
        try:
            num_p  = view.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER)
            name_p = view.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NAME)
            if num_p and name_p:
                num  = num_p.AsString() or ""
                name = name_p.AsString() or ""
                if num:
                    return "{} - {}".format(num, name)
        except Exception:
            pass
        return "(not on sheet)"

    @property
    def RevitView(self):
        return self._view

    @property
    def HasTemplate(self):
        return self.TemplateName != "(none)"

    def update_row_visuals(self):
        """Sync WPF row background with current IsSelected state."""
        if self.row_border is None:
            return
        try:
            if self.IsSelected:
                self.row_border.Background = _brush(ROW_BG_CHECKED)
            else:
                bg = ROW_BG_ODD if self.display_index % 2 == 0 else ROW_BG_EVEN
                self.row_border.Background = _brush(bg)
        except Exception:
            pass


class ResultItem(object):
    """Bindable result row for a single untagged element."""

    def __init__(self, element, view, category_name, type_name, level_name):
        self.Category  = category_name
        self.ElementId = str(eid_int(element.Id))
        self.TypeName  = type_name
        self.Level     = level_name
        self.ViewName  = view.Name or "(unnamed)"
        self._elem_id  = element.Id
        self._view_id  = view.Id

    @property
    def ElemId(self):
        return self._elem_id

    @property
    def ViewId(self):
        return self._view_id


# ── Core analysis helpers ──────────────────────────────────────────────────────

def get_tagged_element_ids_in_view(view):
    """Return a set of integer element IDs that carry a tag in this view."""
    tagged_ids = set()
    try:
        tags = (
            DB.FilteredElementCollector(doc, view.Id)
            .OfClass(DB.IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        logger.debug("Tag collector failed for '{}': {}".format(view.Name, str(e)))
        return tagged_ids

    for tag in tags:
        try:
            # Revit 2023-2025: returns LinkElementId; 2026+: returns ElementId
            for item in tag.GetTaggedLocalElementIds():
                if hasattr(item, 'HostElementId'):
                    for eid in (item.HostElementId, item.LinkedElementId):
                        v = eid_int(eid)
                        if v != INVALID_ID_INT:
                            tagged_ids.add(v)
                else:
                    v = eid_int(item)
                    if v != INVALID_ID_INT:
                        tagged_ids.add(v)
        except Exception:
            try:
                eid = tag.TaggedLocalElementId
                if eid is not None and eid_int(eid) != INVALID_ID_INT:
                    tagged_ids.add(eid_int(eid))
            except Exception:
                pass

    return tagged_ids


def get_elements_in_view(view, bic):
    """Return model elements of the given BuiltInCategory visible in the view."""
    try:
        return list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        logger.debug("Element collector failed ({}/{}): {}".format(view.Name, bic, str(e)))
        return []


def get_element_type_name(element):
    """Return 'Family : Type' string, or '(unknown)'."""
    try:
        type_id = element.GetTypeId()
        if type_id is None or eid_int(type_id) == INVALID_ID_INT:
            return "(unknown type)"
        et = doc.GetElement(type_id)
        if et is None:
            return "(unknown type)"
        try:
            fp     = et.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            family = fp.AsString() if fp else ""
        except Exception:
            family = ""
        name = et.Name or ""
        if family and name:
            return "{} : {}".format(family, name)
        return name or family or "(unknown)"
    except Exception:
        return "(unknown)"


def get_element_level_name(element):
    """Return level name for an element, or '-'."""
    for bip in (DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM):
        try:
            p = element.get_Parameter(bip)
            if p and eid_int(p.AsElementId()) != INVALID_ID_INT:
                lvl = doc.GetElement(p.AsElementId())
                if lvl:
                    return lvl.Name
        except Exception:
            continue
    try:
        if hasattr(element, 'LevelId') and eid_int(element.LevelId) != INVALID_ID_INT:
            lvl = doc.GetElement(element.LevelId)
            if lvl:
                return lvl.Name
    except Exception:
        pass
    return "-"


def analyse_view(view, active_categories):
    """Return ResultItems for every untagged element visible in the view."""
    results    = []
    tagged_ids = get_tagged_element_ids_in_view(view)
    for cat_name, bic in active_categories:
        for el in get_elements_in_view(view, bic):
            try:
                if eid_int(el.Id) not in tagged_ids:
                    results.append(ResultItem(
                        el, view, cat_name,
                        get_element_type_name(el),
                        get_element_level_name(el),
                    ))
            except Exception as e:
                logger.debug("Skipping element {}: {}".format(el.Id, str(e)))
    return results


def collect_candidate_views():
    """Return all non-template views of relevant types from the document."""
    valid_types = {
        DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
        DB.ViewType.Section,   DB.ViewType.Elevation,
        DB.ViewType.ThreeD,
    }
    try:
        return [v for v in
                DB.FilteredElementCollector(doc)
                  .OfClass(DB.View)
                  .WhereElementIsNotElementType()
                  .ToElements()
                if not v.IsTemplate and v.ViewType in valid_types]
    except Exception as e:
        logger.error("Failed to collect views: {}".format(str(e)))
        return []


def navigate_to_element(element, view):
    """Open the view and zoom/select the element. Returns True on success."""
    try:
        from System.Collections.Generic import List as NetList
        ids = NetList[DB.ElementId]()
        ids.Add(element.Id)
        uidoc.RequestViewChange(view)
        uidoc.Selection.SetElementIds(ids)
        uidoc.ShowElements(ids)
        return True
    except Exception as e:
        logger.warning("Navigation failed: {}".format(str(e)))
        return False


def pump_dispatcher():
    """Pump WPF dispatcher to keep UI responsive during the analysis loop."""
    try:
        Dispatcher.CurrentDispatcher.Invoke(
            System.Action(lambda: None),
            DispatcherPriority.Background
        )
    except Exception:
        pass


# ── Main Window ────────────────────────────────────────────────────────────────

class TaglessShamelist(Window):

    def __init__(self):
        wpf.LoadComponent(self, StringReader(XAML))
        self._all_view_items   = []
        self._shown_view_items = []
        self._all_results      = []
        self._is_analysing     = False
        self._cancel_requested = False
        # Anchor index within _shown_view_items for shift-click range selection
        self._last_clicked_idx = None
        self._populate_views()
        self._populate_filter_dropdowns()
        self._connect_events()

    # ── Initialisation ─────────────────────────────────────────────────────────

    def _populate_views(self):
        try:
            self._all_view_items = [ViewItem(v) for v in collect_candidate_views()]
        except Exception as e:
            logger.error("populate_views error: {}".format(str(e)))
            self._all_view_items = []
        self._apply_view_filter()

    def _populate_filter_dropdowns(self):
        self.UI_scale_filter.Items.Clear()
        self.UI_scale_filter.Items.Add("All Scales")
        seen_scales = set()
        for vi in self._all_view_items:
            if vi.ScaleLabel and vi.ScaleLabel != "-" and vi.ScaleLabel not in seen_scales:
                seen_scales.add(vi.ScaleLabel)
                self.UI_scale_filter.Items.Add(vi.ScaleLabel)
        self.UI_scale_filter.SelectedIndex = 0

        self.UI_template_filter.Items.Clear()
        self.UI_template_filter.Items.Add("All Templates")
        seen_tmpl = set()
        for vi in self._all_view_items:
            if vi.TemplateName and vi.TemplateName not in seen_tmpl:
                seen_tmpl.add(vi.TemplateName)
                self.UI_template_filter.Items.Add(vi.TemplateName)
        self.UI_template_filter.SelectedIndex = 0

    def _connect_events(self):
        self.UI_btn_analyse.Click         += self._on_analyse
        self.UI_btn_cancel_analysis.Click += self._on_cancel_analysis
        self.UI_btn_close.Click           += self._on_close
        self.UI_btn_selall.Click          += self._on_select_all
        self.UI_btn_selfiltered.Click     += self._on_select_filtered
        self.UI_btn_selclear.Click        += self._on_clear_all
        self.UI_view_filter.TextChanged   += self._on_view_filter_changed
        self.UI_result_filter.TextChanged += self._on_result_filter_changed
        self.UI_scale_filter.SelectionChanged    += self._on_view_filter_changed
        self.UI_template_filter.SelectionChanged += self._on_view_filter_changed
        for ctrl in (self.UI_chk_onsheet, self.UI_chk_no_template,
                     self.UI_chk_floorplan, self.UI_chk_ceilingplan,
                     self.UI_chk_section, self.UI_chk_elevation, self.UI_chk_3d):
            ctrl.Checked   += self._on_view_type_filter_changed
            ctrl.Unchecked += self._on_view_type_filter_changed
        self.UI_results_list.MouseDoubleClick += self._on_result_double_click

    # ── Row builder — StackPanel pattern for reliable shift-click ──────────────

    def _build_view_row(self, vi, display_index):
        """
        Construct one WPF row for vi inside the StackPanel.

        Pattern: Border > Grid with a CheckBox and TextBlocks.
        CheckBox.Click is wired directly (not through DataTemplate binding)
        so Keyboard.Modifiers is accessible at event time — the only reliable
        way to detect Shift in IronPython WPF.
        """
        vi.display_index = display_index

        row = Border()
        row.BorderThickness = Thickness(0, 0, 0, 1)
        row.BorderBrush     = _brush(ROW_BORDER_CLR)
        row.Padding         = Thickness(2, 3, 2, 3)
        row.Cursor          = System.Windows.Input.Cursors.Hand

        inner = WPFGrid()
        # column widths match the XAML header row exactly
        for w in (26, 290, 105, 75, 155, 0):
            cd = WPFColumnDefinition()
            cd.Width = (GridLength(w) if w > 0
                        else GridLength(1, GridUnitType.Star))
            inner.ColumnDefinitions.Add(cd)

        chk = CheckBox()
        chk.VerticalAlignment   = VerticalAlignment.Center
        chk.HorizontalAlignment = HorizontalAlignment.Center
        chk.IsChecked           = vi.IsSelected
        WPFGrid.SetColumn(chk, 0)

        def _tb(text, col):
            tb = TextBlock()
            tb.Text          = text
            tb.FontSize      = 11
            tb.Foreground    = _brush("#2C2C2C")
            tb.VerticalAlignment  = VerticalAlignment.Center
            tb.TextTrimming  = TextTrimming.CharacterEllipsis
            tb.Margin        = Thickness(2, 0, 4, 0)
            WPFGrid.SetColumn(tb, col)
            return tb

        inner.Children.Add(chk)
        inner.Children.Add(_tb(vi.ViewName,     1))
        inner.Children.Add(_tb(vi.ViewType,     2))
        inner.Children.Add(_tb(vi.ScaleLabel,   3))
        inner.Children.Add(_tb(vi.TemplateName, 4))
        inner.Children.Add(_tb(vi.SheetInfo,    5))

        row.Child = inner

        # Store WPF references back onto the ViewItem for later updates
        vi.row_border   = row
        vi.row_checkbox = chk

        # Hover highlight
        def _hover_on(s, a, _vi=vi):
            if not _vi.IsSelected:
                _vi.row_border.Background = _brush(ROW_BG_HOVER)

        def _hover_off(s, a, _vi=vi):
            _vi.update_row_visuals()

        row.MouseEnter += _hover_on
        row.MouseLeave += _hover_off

        # Both the checkbox click and a bare row click route through the same handler.
        # CheckBox.Click fires before MouseLeftButtonUp so we suppress double-firing
        # by checking the source in the row handler.
        def _chk_click(s, a, _vi=vi):
            # Revert the automatic toggle WPF applies before we process shift logic
            _vi.row_checkbox.IsChecked = _vi.IsSelected
            self._handle_row_click(_vi)

        def _row_click(s, a, _vi=vi):
            # Only fire when clicking the row background, not the checkbox itself
            src = a.OriginalSource
            if isinstance(src, CheckBox):
                return
            # Also skip if source is the checkbox's internal visuals
            try:
                parent = System.Windows.Media.VisualTreeHelper.GetParent(src)
                while parent is not None:
                    if isinstance(parent, CheckBox):
                        return
                    parent = System.Windows.Media.VisualTreeHelper.GetParent(parent)
            except Exception:
                pass
            self._handle_row_click(_vi)

        chk.Click                += _chk_click
        row.MouseLeftButtonUp    += _row_click

        vi.update_row_visuals()
        return row

    def _rebuild_view_panel(self):
        """Repopulate the StackPanel from current _shown_view_items."""
        self.UI_view_panel.Children.Clear()
        self._last_clicked_idx = None
        for idx, vi in enumerate(self._shown_view_items):
            self.UI_view_panel.Children.Add(self._build_view_row(vi, idx))
        self._update_footer()

    # ── Shift-click range selection ────────────────────────────────────────────

    def _handle_row_click(self, clicked_vi):
        """
        Core selection logic called by both checkbox and row clicks.

        Shift held + previous anchor  -> range select using anchor's state
        No shift (or no anchor)       -> single toggle, update anchor
        """
        try:
            clicked_idx = self._shown_view_items.index(clicked_vi)
        except ValueError:
            return

        shift_held = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift

        if shift_held and self._last_clicked_idx is not None:
            # Apply the anchor row's current selection state across the range
            anchor_state = self._shown_view_items[self._last_clicked_idx].IsSelected
            lo = min(self._last_clicked_idx, clicked_idx)
            hi = max(self._last_clicked_idx, clicked_idx)
            for vi in self._shown_view_items[lo:hi + 1]:
                vi.IsSelected = anchor_state
                vi.update_row_visuals()
                if vi.row_checkbox is not None:
                    vi.row_checkbox.IsChecked = anchor_state
        else:
            # Plain click — toggle and set as new anchor
            clicked_vi.IsSelected = not clicked_vi.IsSelected
            clicked_vi.update_row_visuals()
            if clicked_vi.row_checkbox is not None:
                clicked_vi.row_checkbox.IsChecked = clicked_vi.IsSelected
            self._last_clicked_idx = clicked_idx

        self._update_footer()

    # ── View list filtering ─────────────────────────────────────────────────────

    def _get_allowed_view_types(self):
        allowed = set()
        if self.UI_chk_floorplan.IsChecked:   allowed.add("FloorPlan")
        if self.UI_chk_ceilingplan.IsChecked: allowed.add("CeilingPlan")
        if self.UI_chk_section.IsChecked:     allowed.add("Section")
        if self.UI_chk_elevation.IsChecked:   allowed.add("Elevation")
        if self.UI_chk_3d.IsChecked:          allowed.add("ThreeD")
        return allowed

    def _apply_view_filter(self):
        text        = (self.UI_view_filter.Text or "").strip().lower()
        on_sheet    = bool(self.UI_chk_onsheet.IsChecked)
        no_template = bool(self.UI_chk_no_template.IsChecked)
        allowed_vt  = self._get_allowed_view_types()
        sel_scale   = None
        sel_tmpl    = None
        try:
            if self.UI_scale_filter.SelectedIndex > 0:
                sel_scale = str(self.UI_scale_filter.SelectedItem)
        except Exception:
            pass
        try:
            if self.UI_template_filter.SelectedIndex > 0:
                sel_tmpl = str(self.UI_template_filter.SelectedItem)
        except Exception:
            pass

        filtered = []
        for vi in self._all_view_items:
            if vi.ViewType not in allowed_vt:
                continue
            if on_sheet and "(not on sheet)" in vi.SheetInfo:
                continue
            if no_template and vi.HasTemplate:
                continue
            if sel_scale and vi.ScaleLabel != sel_scale:
                continue
            if sel_tmpl and vi.TemplateName != sel_tmpl:
                continue
            if text and text not in vi.ViewName.lower():
                continue
            filtered.append(vi)

        self._shown_view_items = filtered
        self._rebuild_view_panel()

    def _on_view_filter_changed(self, sender, args):
        self._apply_view_filter()

    def _on_view_type_filter_changed(self, sender, args):
        self._apply_view_filter()

    # ── Selection helpers ──────────────────────────────────────────────────────

    def _on_select_all(self, sender, args):
        for vi in self._all_view_items:
            vi.IsSelected = True
            vi.update_row_visuals()
            if vi.row_checkbox is not None:
                vi.row_checkbox.IsChecked = True
        self._update_footer()

    def _on_select_filtered(self, sender, args):
        for vi in self._shown_view_items:
            vi.IsSelected = True
            vi.update_row_visuals()
            if vi.row_checkbox is not None:
                vi.row_checkbox.IsChecked = True
        self._update_footer()

    def _on_clear_all(self, sender, args):
        for vi in self._all_view_items:
            vi.IsSelected = False
            vi.update_row_visuals()
            if vi.row_checkbox is not None:
                vi.row_checkbox.IsChecked = False
        self._update_footer()

    # ── Results filtering ──────────────────────────────────────────────────────

    def _apply_result_filter(self):
        text = (self.UI_result_filter.Text or "").strip().lower()
        shown = (
            self._all_results if not text else
            [r for r in self._all_results
             if text in r.Category.lower()
             or text in r.ElementId
             or text in r.TypeName.lower()
             or text in r.Level.lower()
             or text in r.ViewName.lower()]
        )
        capped = shown[:MAX_RESULTS_SHOWN]
        self.UI_results_list.ItemsSource = None
        self.UI_results_list.ItemsSource = capped
        self._update_status(len(shown), len(capped))

    def _on_result_filter_changed(self, sender, args):
        self._apply_result_filter()

    # ── Analysis state ─────────────────────────────────────────────────────────

    def _set_analysing(self, is_analysing):
        self._is_analysing = is_analysing
        if is_analysing:
            self.UI_btn_analyse.IsEnabled          = False
            self.UI_btn_cancel_analysis.Visibility = Visibility.Visible
            self.UI_btn_cancel_analysis.IsEnabled  = True
            self.UI_btn_close.IsEnabled            = False
            self.UI_progress_bar.Visibility        = Visibility.Visible
            self.UI_progress_bar.Value             = 0
            self.UI_btn_selall.IsEnabled           = False
            self.UI_btn_selfiltered.IsEnabled      = False
            self.UI_btn_selclear.IsEnabled         = False
        else:
            self.UI_btn_analyse.IsEnabled          = True
            self.UI_btn_cancel_analysis.Visibility = Visibility.Collapsed
            self.UI_btn_close.IsEnabled            = True
            self.UI_progress_bar.Visibility        = Visibility.Collapsed
            self.UI_btn_selall.IsEnabled           = True
            self.UI_btn_selfiltered.IsEnabled      = True
            self.UI_btn_selclear.IsEnabled         = True

    def _on_cancel_analysis(self, sender, args):
        self._cancel_requested = True
        self.UI_btn_cancel_analysis.IsEnabled = False
        self.UI_status_text.Text = "Cancelling after current view..."

    # ── Analysis ───────────────────────────────────────────────────────────────

    def _get_active_categories(self):
        cats = []
        if self.UI_cat_walls.IsChecked:
            cats.append(("Walls",    DB.BuiltInCategory.OST_Walls))
        if self.UI_cat_floors.IsChecked:
            cats.append(("Floors",   DB.BuiltInCategory.OST_Floors))
        if self.UI_cat_ceilings.IsChecked:
            cats.append(("Ceilings", DB.BuiltInCategory.OST_Ceilings))
        return cats

    def _on_analyse(self, sender, args):
        active_cats = self._get_active_categories()
        if not active_cats:
            forms.alert("Select at least one category to check.")
            return

        selected_views = [vi for vi in self._all_view_items if vi.IsSelected]
        if not selected_views:
            forms.alert("No views selected. Use the checkboxes to select views.")
            return

        if len(selected_views) > MAX_VIEWS_WARNING:
            msg = ("You have selected {} views. This may take some time.\n\n"
                   "Use the Cancel button to stop mid-run.\n\nContinue?").format(
                len(selected_views))
            if not forms.alert(msg, yes=True, no=True):
                return

        self._cancel_requested = False
        self._all_results      = []
        self._set_analysing(True)
        self.UI_status_text.Text = "Starting analysis..."

        total  = len(selected_views)
        failed = []

        try:
            for idx, vi in enumerate(selected_views):
                if self._cancel_requested:
                    break
                self.UI_progress_bar.Value = int((idx / float(total)) * 100)
                self.UI_status_text.Text   = "Analysing view {}/{}: {}".format(
                    idx + 1, total, vi.ViewName[:45])
                pump_dispatcher()
                try:
                    self._all_results.extend(analyse_view(vi.RevitView, active_cats))
                except Exception as e:
                    logger.warning("Failed on view '{}': {}".format(vi.ViewName, str(e)))
                    failed.append(vi.ViewName)
        finally:
            self._set_analysing(False)

        # De-duplicate: same element + view + category once only
        seen, deduped = set(), []
        for r in self._all_results:
            key = (r.ElementId, r.ViewName, r.Category)
            if key not in seen:
                seen.add(key)
                deduped.append(r)
        self._all_results = deduped

        self.UI_result_badge.Text = "{} untagged".format(len(self._all_results))
        if failed:
            logger.warning("Views skipped: {}".format(", ".join(failed)))
        self._apply_result_filter()

    # ── Navigation ─────────────────────────────────────────────────────────────

    def _on_result_double_click(self, sender, args):
        item = self.UI_results_list.SelectedItem
        if not isinstance(item, ResultItem):
            return
        try:
            el = doc.GetElement(item.ElemId)
            vw = doc.GetElement(item.ViewId)
            if el is None:
                forms.alert("Element {} no longer exists.".format(item.ElementId))
                return
            if vw is None:
                forms.alert("View '{}' no longer exists.".format(item.ViewName))
                return
            if not navigate_to_element(el, vw):
                forms.alert(
                    "Could not navigate automatically.\n"
                    "Element ID: {}  View: {}".format(item.ElementId, item.ViewName)
                )
        except Exception as e:
            logger.error("Navigation error: {}".format(str(e)))
            forms.alert("Navigation error: {}".format(str(e)))

    # ── UI helpers ─────────────────────────────────────────────────────────────

    def _update_footer(self):
        total    = len(self._all_view_items)
        shown    = len(self._shown_view_items)
        selected = sum(1 for vi in self._all_view_items if vi.IsSelected)
        self.UI_footer_text.Text = (
            "Showing {} of {} views  |  {} selected".format(shown, total, selected)
        )

    def _update_status(self, total_matched, shown_count):
        if not self._all_results:
            self.UI_status_text.Text = "No untagged elements found."
            return
        if shown_count < total_matched:
            self.UI_status_text.Text = (
                "Showing {} of {} matched (cap: {})".format(
                    shown_count, total_matched, MAX_RESULTS_SHOWN)
            )
        elif total_matched < len(self._all_results):
            self.UI_status_text.Text = "Showing {} of {} results".format(
                total_matched, len(self._all_results))
        else:
            self.UI_status_text.Text = "Showing all {} results".format(total_matched)

    def _on_close(self, sender, args):
        if self._is_analysing:
            self._cancel_requested = True
        self.Close()


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    if not doc:
        forms.alert("No active Revit document found.", exitscript=True)
        return
    try:
        TaglessShamelist().ShowDialog()
    except Exception as e:
        logger.error("Tool launch error: {}".format(str(e)))
        forms.alert("Failed to launch Tagless Shamelist: {}".format(str(e)))


main()
