# -*- coding: utf-8 -*-
"""
Script:   Tagless Shamelist
Desc:     Identifies untagged Walls, Floors, and Ceilings across selected views.
          Double-click a result row to navigate to the element in its view.
Author:   Anirudh Sood @ Aukett Swanke
Usage:    Run tool, select views, click Analyse. Double-click rows to navigate.
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
from System.Windows import Window
from System.Windows.Threading import Dispatcher, DispatcherPriority
import System
import wpf

logger = script.get_logger()
doc    = revit.doc
uidoc  = revit.uidoc


# ── Constants ──────────────────────────────────────────────────────────────────
MAX_VIEWS_WARNING  = 100   # Warn user before analysing very large view sets
MAX_RESULTS_SHOWN  = 5000  # Safety cap: avoid UI freeze on extreme result counts


# ── Revit 2026 compatibility ───────────────────────────────────────────────────
def eid_int(element_id):
    """Return the integer value of an ElementId.  Compatible with Revit 2023-2026."""
    try:
        return element_id.Value          # Revit 2026+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023-2025


INVALID_ID_INT = eid_int(DB.ElementId.InvalidElementId)


# ── XAML ───────────────────────────────────────────────────────────────────────
XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Tagless Shamelist"
    Width="860" Height="700"
    MinWidth="700" MinHeight="520"
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

        <!-- Amber stop button shown during analysis -->
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

    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="*"/>
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
                               FontWeight="Bold" FontSize="15" FontFamily="Arial"/>
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
        <GroupBox Grid.Row="2" Header=" View Selection " Padding="8" Background="#FFFFFF"
                  BorderBrush="#CCCCCC" Foreground="#2C2C2C" FontWeight="SemiBold">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="130"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Filter Views:" VerticalAlignment="Center"
                               FontWeight="SemiBold" Foreground="#2C2C2C" Margin="0,0,8,0"/>
                    <TextBox Grid.Column="1" x:Name="UI_view_filter" Height="24"
                             VerticalContentAlignment="Center" Padding="4,0"
                             Background="#FFFFFF" Foreground="#2C2C2C"
                             BorderBrush="#CCCCCC" BorderThickness="1"/>
                    <CheckBox Grid.Column="2" x:Name="UI_chk_onsheet" Content="On Sheet Only"
                              VerticalAlignment="Center" Foreground="#2C2C2C"/>
                    <Button Grid.Column="4" x:Name="UI_btn_selall" Content="Select All"
                            Style="{StaticResource AUKButton}" Width="85"/>
                    <Button Grid.Column="6" x:Name="UI_btn_selclear" Content="Clear"
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

                <!-- View list -->
                <ListView Grid.Row="4" x:Name="UI_view_list"
                          BorderBrush="#CCCCCC" BorderThickness="1"
                          Background="#FFFFFF"
                          SelectionMode="Extended"
                          VirtualizingStackPanel.IsVirtualizing="True">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Width="28">
                                <GridViewColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox IsChecked="{Binding IsSelected}"
                                                  VerticalAlignment="Center"
                                                  HorizontalAlignment="Center"/>
                                    </DataTemplate>
                                </GridViewColumn.CellTemplate>
                            </GridViewColumn>
                            <GridViewColumn Header="View Name" Width="320"
                                            DisplayMemberBinding="{Binding ViewName}"/>
                            <GridViewColumn Header="Type"      Width="120"
                                            DisplayMemberBinding="{Binding ViewType}"/>
                            <GridViewColumn Header="Sheet"     Width="200"
                                            DisplayMemberBinding="{Binding SheetInfo}"/>
                        </GridView>
                    </ListView.View>
                </ListView>
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
                          BorderBrush="#CCCCCC" BorderThickness="1"
                          Background="#FFFFFF"
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

        <!-- Progress Bar (hidden until analysis starts) -->
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
            <!-- Cancel button: visible only during analysis -->
            <Button Grid.Column="3" x:Name="UI_btn_cancel_analysis" Content="Cancel"
                    Style="{StaticResource StopButton}" Width="80"
                    Visibility="Collapsed"/>
            <Button Grid.Column="5" x:Name="UI_btn_close" Content="Close"
                    Style="{StaticResource CancelButton}" Width="80"/>
        </Grid>
    </Grid>
</Window>
"""


# ── Data models ────────────────────────────────────────────────────────────────

class ViewItem(object):
    """Bindable wrapper for a Revit view used in the view list."""

    def __init__(self, view):
        self._view      = view
        self.IsSelected = True
        self.ViewName   = view.Name or "(unnamed)"
        self.ViewType   = str(view.ViewType)
        self.SheetInfo  = self._get_sheet_info(view)

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


class ResultItem(object):
    """Bindable result row for an untagged element."""

    def __init__(self, element, view, category_name, type_name, level_name):
        self.Category  = category_name
        self.ElementId = str(eid_int(element.Id))
        self.TypeName  = type_name
        self.Level     = level_name
        self.ViewName  = view.Name or "(unnamed)"
        self._elem_id  = element.Id   # Store Id only; re-fetch element at navigation time
        self._view_id  = view.Id

    @property
    def ElemId(self):
        return self._elem_id

    @property
    def ViewId(self):
        return self._view_id


# ── Core helpers ───────────────────────────────────────────────────────────────

def get_tagged_element_ids_in_view(view):
    """Return a set of integer element IDs that have a tag in this view."""
    tagged_ids = set()
    try:
        tags = (
            DB.FilteredElementCollector(doc, view.Id)
            .OfClass(DB.IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        logger.debug("Tag collector failed for view '{}': {}".format(view.Name, str(e)))
        return tagged_ids

    for tag in tags:
        try:
            # Revit 2023-2025 returns ICollection<LinkElementId>; 2026+ returns ICollection<ElementId>
            multi_ids = tag.GetTaggedLocalElementIds()
            for item in multi_ids:
                if hasattr(item, 'HostElementId'):
                    host   = eid_int(item.HostElementId)
                    linked = eid_int(item.LinkedElementId)
                    if host   != INVALID_ID_INT: tagged_ids.add(host)
                    if linked != INVALID_ID_INT: tagged_ids.add(linked)
                else:
                    val = eid_int(item)
                    if val != INVALID_ID_INT:
                        tagged_ids.add(val)
        except Exception:
            # Fallback for older API versions
            try:
                eid = tag.TaggedLocalElementId
                if eid is not None and eid_int(eid) != INVALID_ID_INT:
                    tagged_ids.add(eid_int(eid))
            except Exception:
                pass

    return tagged_ids


def get_elements_in_view(view, bic):
    """Collect model elements of the given BuiltInCategory visible in the view."""
    try:
        return list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        logger.debug("Element collector failed ({} / {}): {}".format(
            view.Name, bic, str(e)))
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
            fp = et.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            family = fp.AsString() if fp else ""
        except Exception:
            family = ""
        type_name = et.Name or ""
        if family and type_name:
            return "{} : {}".format(family, type_name)
        return type_name or family or "(unknown)"
    except Exception:
        return "(unknown)"


def get_element_level_name(element):
    """Return the level name for an element, or '-' if unavailable."""
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
    """Return ResultItem list for every untagged element visible in the view."""
    results = []
    tagged_ids = get_tagged_element_ids_in_view(view)

    for cat_name, bic in active_categories:
        elements = get_elements_in_view(view, bic)
        for el in elements:
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
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.Section,
        DB.ViewType.Elevation,
        DB.ViewType.ThreeD,
    }
    try:
        all_views = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.View)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        return [v for v in all_views
                if not v.IsTemplate and v.ViewType in valid_types]
    except Exception as e:
        logger.error("Failed to collect views: {}".format(str(e)))
        return []


def navigate_to_element(element, view):
    """Open the view and zoom to / select the element."""
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
    """
    Process pending WPF dispatcher messages so the UI stays responsive
    during a long synchronous loop.  This is safe in IronPython on the
    UI thread — it processes only background/render priority messages.
    """
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
        self._shown_results    = []
        self._is_analysing     = False
        self._cancel_requested = False
        self._populate_views()
        self._connect_events()

    # -- Initialisation --------------------------------------------------------

    def _populate_views(self):
        try:
            self._all_view_items = [ViewItem(v) for v in collect_candidate_views()]
        except Exception as e:
            logger.error("populate_views error: {}".format(str(e)))
            self._all_view_items = []
        self._apply_view_filter()

    def _connect_events(self):
        self.UI_btn_analyse.Click         += self._on_analyse
        self.UI_btn_cancel_analysis.Click += self._on_cancel_analysis
        self.UI_btn_close.Click           += self._on_close
        self.UI_btn_selall.Click          += self._on_select_all
        self.UI_btn_selclear.Click        += self._on_clear_all

        self.UI_view_filter.TextChanged   += self._on_view_filter_changed
        self.UI_result_filter.TextChanged += self._on_result_filter_changed

        for ctrl in (self.UI_chk_onsheet, self.UI_chk_floorplan,
                     self.UI_chk_ceilingplan, self.UI_chk_section,
                     self.UI_chk_elevation, self.UI_chk_3d):
            ctrl.Checked   += self._on_view_type_filter_changed
            ctrl.Unchecked += self._on_view_type_filter_changed

        self.UI_results_list.MouseDoubleClick += self._on_result_double_click

    # -- View list filtering ---------------------------------------------------

    def _get_allowed_view_types(self):
        allowed = set()
        if self.UI_chk_floorplan.IsChecked:    allowed.add("FloorPlan")
        if self.UI_chk_ceilingplan.IsChecked:  allowed.add("CeilingPlan")
        if self.UI_chk_section.IsChecked:      allowed.add("Section")
        if self.UI_chk_elevation.IsChecked:    allowed.add("Elevation")
        if self.UI_chk_3d.IsChecked:           allowed.add("ThreeD")
        return allowed

    def _apply_view_filter(self):
        text       = (self.UI_view_filter.Text or "").strip().lower()
        on_sheet   = bool(self.UI_chk_onsheet.IsChecked)
        allowed_vt = self._get_allowed_view_types()

        filtered = []
        for vi in self._all_view_items:
            if vi.ViewType not in allowed_vt:
                continue
            if on_sheet and "(not on sheet)" in vi.SheetInfo:
                continue
            if text and text not in vi.ViewName.lower():
                continue
            filtered.append(vi)

        self._shown_view_items        = filtered
        self.UI_view_list.ItemsSource = None
        self.UI_view_list.ItemsSource = filtered
        self._update_footer()

    def _on_view_filter_changed(self, sender, args):
        self._apply_view_filter()

    def _on_view_type_filter_changed(self, sender, args):
        self._apply_view_filter()

    def _on_select_all(self, sender, args):
        for vi in self._shown_view_items:
            vi.IsSelected = True
        self._refresh_view_list()

    def _on_clear_all(self, sender, args):
        for vi in self._shown_view_items:
            vi.IsSelected = False
        self._refresh_view_list()

    def _refresh_view_list(self):
        src = self.UI_view_list.ItemsSource
        self.UI_view_list.ItemsSource = None
        self.UI_view_list.ItemsSource = src

    # -- Results filtering -----------------------------------------------------

    def _apply_result_filter(self):
        text = (self.UI_result_filter.Text or "").strip().lower()
        if not text:
            self._shown_results = list(self._all_results)
        else:
            self._shown_results = [
                r for r in self._all_results
                if text in r.Category.lower()
                or text in r.ElementId
                or text in r.TypeName.lower()
                or text in r.Level.lower()
                or text in r.ViewName.lower()
            ]
        # Safety cap to prevent UI freeze on enormous result sets
        capped = self._shown_results[:MAX_RESULTS_SHOWN]
        self.UI_results_list.ItemsSource = None
        self.UI_results_list.ItemsSource = capped
        self._update_status(len(self._shown_results), len(capped))

    def _on_result_filter_changed(self, sender, args):
        self._apply_result_filter()

    # -- Analysis state management ---------------------------------------------

    def _set_analysing(self, is_analysing):
        """Toggle UI between idle and analysing states."""
        self._is_analysing = is_analysing
        from System.Windows import Visibility
        if is_analysing:
            self.UI_btn_analyse.IsEnabled          = False
            self.UI_btn_cancel_analysis.Visibility = Visibility.Visible
            self.UI_btn_close.IsEnabled            = False
            self.UI_progress_bar.Visibility        = Visibility.Visible
            self.UI_progress_bar.Value             = 0
            self.UI_btn_selall.IsEnabled           = False
            self.UI_btn_selclear.IsEnabled         = False
        else:
            self.UI_btn_analyse.IsEnabled          = True
            self.UI_btn_cancel_analysis.Visibility = Visibility.Collapsed
            self.UI_btn_close.IsEnabled            = True
            self.UI_progress_bar.Visibility        = Visibility.Collapsed
            self.UI_btn_selall.IsEnabled           = True
            self.UI_btn_selclear.IsEnabled         = True

    def _on_cancel_analysis(self, sender, args):
        """Signal the running analysis loop to stop after the current view."""
        self._cancel_requested = True
        self.UI_btn_cancel_analysis.IsEnabled = False
        self.UI_status_text.Text = "Cancelling after current view..."

    # -- Analysis --------------------------------------------------------------

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

        selected_views = [vi for vi in self._shown_view_items if vi.IsSelected]
        if not selected_views:
            forms.alert("No views selected. Tick the checkboxes to select views.")
            return

        # Warn on large view sets but allow the user to proceed
        if len(selected_views) > MAX_VIEWS_WARNING:
            msg = ("You have selected {} views, which may take some time.\n\n"
                   "You can use the Cancel button to stop mid-run.\n\n"
                   "Continue?").format(len(selected_views))
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

                # Check cancellation flag (set by Cancel button click)
                if self._cancel_requested:
                    break

                # Update progress bar and status label
                progress = int((idx / float(total)) * 100)
                self.UI_progress_bar.Value = progress
                self.UI_status_text.Text   = "Analysing view {}/{}: {}".format(
                    idx + 1, total, vi.ViewName[:40])

                # Pump the dispatcher so the UI repaints and the Cancel
                # button click event can be processed between views
                pump_dispatcher()

                try:
                    self._all_results.extend(analyse_view(vi.RevitView, active_cats))
                except Exception as e:
                    logger.warning("Failed on view '{}': {}".format(vi.ViewName, str(e)))
                    failed.append(vi.ViewName)

        finally:
            self._set_analysing(False)
            self.UI_btn_cancel_analysis.IsEnabled = True

        # De-duplicate: same element in same view and category reported once
        seen    = set()
        deduped = []
        for r in self._all_results:
            key = (r.ElementId, r.ViewName, r.Category)
            if key not in seen:
                seen.add(key)
                deduped.append(r)
        self._all_results = deduped

        count = len(self._all_results)
        self.UI_result_badge.Text = "{} untagged".format(count)

        # Build status suffix
        suffix_parts = []
        if self._cancel_requested:
            suffix_parts.append("cancelled")
        if failed:
            suffix_parts.append("{} view(s) skipped due to errors".format(len(failed)))
        self._apply_result_filter()

        # Show a summary if something went wrong
        if failed:
            logger.warning("Views that failed: {}".format(", ".join(failed)))

    # -- Navigation ------------------------------------------------------------

    def _on_result_double_click(self, sender, args):
        item = self.UI_results_list.SelectedItem
        if not isinstance(item, ResultItem):
            return
        try:
            el = doc.GetElement(item.ElemId)
            vw = doc.GetElement(item.ViewId)
            if el is None:
                forms.alert("Element {} no longer exists in the model.".format(item.ElementId))
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

    # -- UI helpers ------------------------------------------------------------

    def _update_footer(self):
        total    = len(self._all_view_items)
        shown    = len(self._shown_view_items)
        selected = sum(1 for vi in self._shown_view_items if vi.IsSelected)
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
        forms.alert("Failed to launch: {}".format(str(e)))


main()
