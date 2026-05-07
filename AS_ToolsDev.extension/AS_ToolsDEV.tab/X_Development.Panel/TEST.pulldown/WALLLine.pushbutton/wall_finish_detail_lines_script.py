# -*- coding: utf-8 -*-
"""
Script:   Wall Finish Detail Lines
Desc:     Places line-based detail items on wall faces with Mark and auto-tag.
          Single WPF window with wall source, type filter, family/tag pickers,
          and Mark value options. AUK branded UI.
Author:   AUK BIM Team
Usage:    Click button, configure in single window, click Place.
Result:   Detail items on both wall faces, Mark set, tagged.
"""
__title__ = "Finish\nLines"
__author__ = "Anirudh Sood"

from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms
from pyrevit.framework import wpf

import clr
clr.AddReference("RevitAPI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

import System
from System.IO import StringReader
from System.Windows import Controls, Visibility, MessageBox, Thickness, FontWeights
from System.Windows.Input import Keyboard, ModifierKeys

logger = script.get_logger()
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

# ── Constants ────────────────────────────────────────────────────────────────
TOOL_NAME = "AUK Wall Finish Lines"
INVALID_ID_INT = -1
TAG_OFFSET_FEET = 0.15
MAX_BATCH_WALLS = 500

# ── XAML ─────────────────────────────────────────────────────────────────────
XAML_STR = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Wall Finish Lines" Width="520" Height="640"
    ShowInTaskbar="False"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize" MinWidth="480" MinHeight="580"
    FontFamily="Arial" FontSize="12" Background="#F5F5F5">

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title Bar -->
        <Border Grid.Row="0" Background="#143D5C" Padding="10" CornerRadius="3">
            <TextBlock Text="Wall Finish Detail Lines"
                       FontWeight="Bold" FontSize="14"
                       Foreground="#FFFFFF"
                       HorizontalAlignment="Center"/>
        </Border>

        <!-- Wall Source -->
        <GroupBox Grid.Row="2" Header="Wall Source" Padding="5"
                  Foreground="#2C2C2C" FontWeight="SemiBold"
                  Background="#FFFFFF">
            <StackPanel>
                <RadioButton x:Name="UI_rb_selection" Content="Current selection"
                             Margin="4" IsChecked="True" FontWeight="Normal"/>
                <RadioButton x:Name="UI_rb_batch" Content="All walls in active view"
                             Margin="4" FontWeight="Normal"/>
                <TextBlock x:Name="UI_wall_count" Text=""
                           Foreground="#666666" FontWeight="Normal"
                           Margin="4,4,4,0" FontSize="11"/>
            </StackPanel>
        </GroupBox>

        <!-- Wall Type Filter -->
        <GroupBox Grid.Row="4" Header="Wall Type Filter" Padding="5"
                  Foreground="#2C2C2C" FontWeight="SemiBold"
                  Background="#FFFFFF">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid Grid.Row="0" Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0"
                               Text="Shift+click for range select"
                               Foreground="#666666" FontWeight="Normal"
                               FontSize="11" VerticalAlignment="Center"/>
                    <Button x:Name="UI_btn_sel_all" Content="All"
                            Grid.Column="1" Width="42" Height="20"
                            Margin="0,0,4,0" FontSize="10"
                            Background="#F5F5F5" Foreground="#2C2C2C"
                            BorderBrush="#CCCCCC" BorderThickness="1"/>
                    <Button x:Name="UI_btn_sel_none" Content="None"
                            Grid.Column="2" Width="42" Height="20"
                            FontSize="10"
                            Background="#F5F5F5" Foreground="#2C2C2C"
                            BorderBrush="#CCCCCC" BorderThickness="1"/>
                </Grid>
                <ScrollViewer Grid.Row="1" MaxHeight="110"
                              VerticalScrollBarVisibility="Auto"
                              BorderBrush="#CCCCCC" BorderThickness="1"
                              Background="#FFFFFF">
                    <StackPanel x:Name="UI_type_panel"/>
                </ScrollViewer>
            </Grid>
        </GroupBox>

        <!-- Families -->
        <GroupBox Grid.Row="6" Header="Families" Padding="5"
                  Foreground="#2C2C2C" FontWeight="SemiBold"
                  Background="#FFFFFF">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Label Grid.Row="0" Grid.Column="0" Content="Detail Item:"
                       FontWeight="Medium" VerticalAlignment="Center"
                       Foreground="#2C2C2C" Margin="0,0,8,0"/>
                <ComboBox Grid.Row="0" Grid.Column="1"
                          x:Name="UI_cb_detail" Height="24"
                          VerticalAlignment="Center" Margin="0,3"
                          Background="#FFFFFF" Foreground="#2C2C2C"
                          BorderBrush="#CCCCCC" BorderThickness="1"/>

                <Label Grid.Row="1" Grid.Column="0" Content="Tag Type:"
                       FontWeight="Medium" VerticalAlignment="Center"
                       Foreground="#2C2C2C" Margin="0,0,8,0"/>
                <ComboBox Grid.Row="1" Grid.Column="1"
                          x:Name="UI_cb_tag" Height="24"
                          VerticalAlignment="Center" Margin="0,3"
                          Background="#FFFFFF" Foreground="#2C2C2C"
                          BorderBrush="#CCCCCC" BorderThickness="1"/>
            </Grid>
        </GroupBox>

        <!-- Mark Value -->
        <GroupBox Grid.Row="8" Header="Mark Parameter" Padding="5"
                  Foreground="#2C2C2C" FontWeight="SemiBold"
                  Background="#FFFFFF">
            <StackPanel>
                <RadioButton x:Name="UI_rb_mark_type" Content="Wall type name (per wall)"
                             Margin="4" IsChecked="True" FontWeight="Normal"/>
                <RadioButton x:Name="UI_rb_mark_custom" Content="Custom value:"
                             Margin="4" FontWeight="Normal"/>
                <TextBox x:Name="UI_tb_mark" Height="24" Margin="24,2,4,2"
                         IsEnabled="False"
                         Background="#FFFFFF" Foreground="#2C2C2C"
                         BorderBrush="#CCCCCC" BorderThickness="1"/>
                <RadioButton x:Name="UI_rb_mark_none" Content="No Mark value"
                             Margin="4" FontWeight="Normal"/>
            </StackPanel>
        </GroupBox>

        <!-- Buttons -->
        <StackPanel Grid.Row="10" Orientation="Horizontal"
                    HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="UI_btn_place" Content="Place"
                    Width="100" Height="28" Margin="0,0,8,0"
                    Background="#143D5C" Foreground="#FFFFFF"
                    FontWeight="SemiBold"/>
            <Button x:Name="UI_btn_cancel" Content="Cancel"
                    Width="90" Height="28"
                    Background="#CCCCCC" Foreground="#2C2C2C"/>
        </StackPanel>
    </Grid>
</Window>"""


# ── Helpers ──────────────────────────────────────────────────────────────────
def eid_int(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def safe_name(element):
    try:
        return DB.Element.Name.GetValue(element)
    except Exception:
        try:
            return element.Name
        except Exception:
            return "<unnamed>"


def get_wall_type_name(wall):
    try:
        wt = doc.GetElement(wall.GetTypeId())
        if wt is not None:
            return safe_name(wt)
    except Exception:
        pass
    return "<unknown type>"


def set_mark_parameter(instance, value):
    try:
        param = instance.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        if param is not None and not param.IsReadOnly:
            param.Set(str(value))
            return True
    except Exception as e:
        logger.warning("Could not set Mark: {}".format(str(e)))
    return False


def get_line_based_detail_families():
    result = []
    try:
        collector = DB.FilteredElementCollector(doc) \
            .OfCategory(DB.BuiltInCategory.OST_DetailComponents) \
            .OfClass(DB.FamilySymbol).ToElements()
        for sym in collector:
            try:
                fam = sym.Family
                if fam is None:
                    continue
                pt = fam.FamilyPlacementType
                if pt == DB.FamilyPlacementType.CurveBasedDetail or \
                   pt == DB.FamilyPlacementType.CurveBased:
                    result.append(sym)
            except Exception:
                continue
    except Exception:
        pass
    return result


def get_detail_component_tag_types():
    result = []
    try:
        target_id = eid_int(DB.ElementId(
            DB.BuiltInCategory.OST_DetailComponentTags))
        collector = DB.FilteredElementCollector(doc) \
            .OfClass(DB.FamilySymbol).ToElements()
        for sym in collector:
            try:
                cat = sym.Category
                if cat is not None and eid_int(cat.Id) == target_id:
                    result.append(sym)
            except Exception:
                continue
    except Exception:
        pass
    return result


def get_wall_face_lines(wall):
    try:
        loc = wall.Location
        if loc is None or not isinstance(loc, DB.LocationCurve):
            return None, None
        wall_curve = loc.Curve
        if wall_curve is None:
            return None, None
        p0 = wall_curve.GetEndPoint(0)
        p1 = wall_curve.GetEndPoint(1)
        if p0.DistanceTo(p1) < 0.001:
            return None, None
        direction = (p1 - p0).Normalize()
        normal = DB.XYZ(-direction.Y, direction.X, 0).Normalize()
        z = p0.Z
        wall_type = doc.GetElement(wall.GetTypeId())
        if wall_type is None:
            return None, None
        cs = None
        try:
            cs = wall_type.GetCompoundStructure()
        except Exception:
            pass
        if cs is not None:
            try:
                offset_ext = cs.GetOffsetForLocationLine(
                    DB.WallLocationLine.FinishFaceExterior)
                offset_int = cs.GetOffsetForLocationLine(
                    DB.WallLocationLine.FinishFaceInterior)
            except Exception:
                half = wall.Width / 2.0
                offset_ext = half
                offset_int = -half
        else:
            half = wall.Width / 2.0
            offset_ext = half
            offset_int = -half

        def make_offset_line(offset_dist):
            off = normal * offset_dist
            ns = DB.XYZ(p0.X + off.X, p0.Y + off.Y, z)
            ne = DB.XYZ(p1.X + off.X, p1.Y + off.Y, z)
            if ns.DistanceTo(ne) < 0.001:
                return None
            return DB.Line.CreateBound(ns, ne)

        return make_offset_line(offset_ext), make_offset_line(offset_int)
    except Exception as e:
        logger.warning("get_wall_face_lines error: {}".format(str(e)))
        return None, None


def place_line_based_instance(view, line, symbol):
    try:
        return doc.Create.NewFamilyInstance(line, symbol, view)
    except Exception as e:
        logger.warning("Failed to place instance: {}".format(str(e)))
        return None


def tag_instance(view, instance, tag_type_id, offset_normal):
    try:
        mid = None
        try:
            bbox = instance.get_BoundingBox(view)
            if bbox is not None:
                mid = DB.XYZ(
                    (bbox.Min.X + bbox.Max.X) / 2.0,
                    (bbox.Min.Y + bbox.Max.Y) / 2.0,
                    (bbox.Min.Z + bbox.Max.Z) / 2.0)
        except Exception:
            pass
        if mid is None:
            try:
                loc = instance.Location
                if isinstance(loc, DB.LocationCurve):
                    mid = loc.Curve.Evaluate(0.5, True)
                elif isinstance(loc, DB.LocationPoint):
                    mid = loc.Point
            except Exception:
                pass
        if mid is None:
            return None
        tag_pt = DB.XYZ(
            mid.X + offset_normal.X * TAG_OFFSET_FEET,
            mid.Y + offset_normal.Y * TAG_OFFSET_FEET,
            mid.Z)
        ref = DB.Reference(instance)
        tag = None
        try:
            tag = DB.IndependentTag.Create(
                doc, view.Id, ref, False,
                DB.TagMode.TM_ADDBY_CATEGORY,
                DB.TagOrientation.Horizontal, tag_pt)
        except Exception:
            try:
                tag = doc.Create.NewTag(
                    view, instance, False,
                    DB.TagMode.TM_ADDBY_CATEGORY,
                    DB.TagOrientation.Horizontal, tag_pt)
            except Exception:
                return None
        if tag is not None and tag_type_id is not None:
            try:
                tag.ChangeTypeId(tag_type_id)
            except Exception:
                pass
        return tag
    except Exception:
        return None


# ── Safe Transaction Helpers ─────────────────────────────────────────────────
def safe_start(name):
    t = None
    try:
        t = DB.Transaction(doc, name)
        status = t.Start()
        if status == DB.TransactionStatus.Started:
            return t, True
    except Exception:
        pass
    if t is not None:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
    return None, False


def safe_commit(t):
    if t is None:
        return False
    try:
        if t.HasStarted() and not t.HasEnded():
            return t.Commit() == DB.TransactionStatus.Committed
    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
    return False


def safe_rollback(t):
    if t is None:
        return
    try:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
    except Exception:
        pass


# ── WPF Window ───────────────────────────────────────────────────────────────
class FinishLinesWindow(forms.WPFWindow):
    def __init__(self, view):
        xaml_stream = StringReader(XAML_STR)
        wpf.LoadComponent(self, xaml_stream)

        self.view = view
        self.result = None  # set to dict on Place click

        # Data maps: display string -> API object
        self._detail_map = {}
        self._tag_map = {}
        self._type_checkboxes = []  # list of (checkbox, type_name)
        self._last_checked_index = -1  # for shift-click range selection

        self.initialize()
        self.connect_events()

    def initialize(self):
        """Populate all controls."""
        # ── Wall count preview ───────────────────────────────────────
        sel_walls = self._get_selection_walls()
        view_walls = self._get_view_walls()
        self.UI_wall_count.Text = "{} selected  |  {} in active view".format(
            len(sel_walls), len(view_walls))

        # ── Populate wall type checkboxes ────────────────────────────
        all_walls = list(set(sel_walls + view_walls))
        type_names = sorted(set(get_wall_type_name(w) for w in all_walls))
        for tn in type_names:
            cb = Controls.CheckBox()
            cb.Content = tn
            cb.IsChecked = True
            cb.Margin = Thickness(2)
            cb.FontWeight = FontWeights.Normal
            cb.Click += self.on_type_cb_click
            self.UI_type_panel.Children.Add(cb)
            self._type_checkboxes.append((cb, tn))

        # ── Detail family combo ──────────────────────────────────────
        symbols = get_line_based_detail_families()
        for sym in symbols:
            display = "{} : {}".format(safe_name(sym.Family), safe_name(sym))
            self.UI_cb_detail.Items.Add(display)
            self._detail_map[display] = sym
        if self.UI_cb_detail.Items.Count > 0:
            self.UI_cb_detail.SelectedIndex = 0

        # ── Tag type combo ───────────────────────────────────────────
        self.UI_cb_tag.Items.Add("(No tagging)")
        self._tag_map["(No tagging)"] = None
        tags = get_detail_component_tag_types()
        for tt in tags:
            display = "{} : {}".format(safe_name(tt.Family), safe_name(tt))
            self.UI_cb_tag.Items.Add(display)
            self._tag_map[display] = tt
        self.UI_cb_tag.SelectedIndex = 0

    def connect_events(self):
        self.UI_btn_place.Click += self.on_place
        self.UI_btn_cancel.Click += self.on_cancel
        self.UI_rb_selection.Checked += self.on_source_changed
        self.UI_rb_batch.Checked += self.on_source_changed
        self.UI_rb_mark_custom.Checked += self.on_mark_mode_changed
        self.UI_rb_mark_type.Checked += self.on_mark_mode_changed
        self.UI_rb_mark_none.Checked += self.on_mark_mode_changed
        self.UI_btn_sel_all.Click += self.on_select_all
        self.UI_btn_sel_none.Click += self.on_select_none

    # ── Event Handlers ───────────────────────────────────────────────
    def on_type_cb_click(self, sender, args):
        """Handle shift-click range selection on type checkboxes."""
        # Find index of clicked checkbox
        clicked_idx = -1
        for i, (cb, tn) in enumerate(self._type_checkboxes):
            if cb is sender:
                clicked_idx = i
                break

        if clicked_idx < 0:
            return

        # Shift held — select/deselect range from last click
        if Keyboard.Modifiers == ModifierKeys.Shift and \
           self._last_checked_index >= 0:
            start = min(self._last_checked_index, clicked_idx)
            end = max(self._last_checked_index, clicked_idx)
            new_state = sender.IsChecked
            for i in range(start, end + 1):
                cb_i, _ = self._type_checkboxes[i]
                if cb_i.IsEnabled:
                    cb_i.IsChecked = new_state

        self._last_checked_index = clicked_idx

    def on_select_all(self, sender, args):
        """Check all enabled type checkboxes."""
        for cb, tn in self._type_checkboxes:
            if cb.IsEnabled:
                cb.IsChecked = True

    def on_select_none(self, sender, args):
        """Uncheck all type checkboxes."""
        for cb, tn in self._type_checkboxes:
            if cb.IsEnabled:
                cb.IsChecked = False

    def on_source_changed(self, sender, args):
        """Update type checkboxes when source changes."""
        walls = self._get_active_walls()
        type_names = set(get_wall_type_name(w) for w in walls)
        for cb, tn in self._type_checkboxes:
            if tn in type_names:
                cb.IsChecked = True
                cb.IsEnabled = True
            else:
                cb.IsChecked = False
                cb.IsEnabled = False

    def on_mark_mode_changed(self, sender, args):
        """Enable/disable custom Mark textbox."""
        self.UI_tb_mark.IsEnabled = self.UI_rb_mark_custom.IsChecked

    def on_cancel(self, sender, args):
        self.result = None
        self.Close()

    def on_place(self, sender, args):
        """Validate and collect all settings, close window."""
        # Validate detail family
        if self.UI_cb_detail.SelectedIndex < 0 or \
           not self.UI_cb_detail.Items.Count:
            forms.alert("No line-based Detail Item family available.\n"
                        "Load one and try again.",
                        title=TOOL_NAME)
            return

        # Validate custom mark
        if self.UI_rb_mark_custom.IsChecked:
            if not self.UI_tb_mark.Text.strip():
                forms.alert("Please enter a custom Mark value.",
                            title=TOOL_NAME)
                return

        # Collect selected type names
        selected_types = set()
        for cb, tn in self._type_checkboxes:
            if cb.IsChecked and cb.IsEnabled:
                selected_types.add(tn)

        if not selected_types:
            forms.alert("No wall types selected.", title=TOOL_NAME)
            return

        # Build result
        detail_key = str(self.UI_cb_detail.SelectedItem)
        tag_key = str(self.UI_cb_tag.SelectedItem)

        self.result = {
            "walls": self._get_filtered_walls(selected_types),
            "detail_symbol": self._detail_map.get(detail_key),
            "tag_type": self._tag_map.get(tag_key),
            "mark_mode": self._get_mark_mode(),
            "custom_mark": self.UI_tb_mark.Text.strip(),
        }

        if not self.result["walls"]:
            forms.alert("No walls match the selected types.",
                        title=TOOL_NAME)
            self.result = None
            return

        if not self.result["detail_symbol"]:
            forms.alert("Please select a Detail Item family.",
                        title=TOOL_NAME)
            self.result = None
            return

        self.Close()

    # ── Internal Helpers ─────────────────────────────────────────────
    def _get_selection_walls(self):
        try:
            sel = revit.get_selection()
            return [e for e in sel.elements if isinstance(e, DB.Wall)]
        except Exception:
            return []

    def _get_view_walls(self):
        try:
            collector = DB.FilteredElementCollector(doc, self.view.Id) \
                .OfCategory(DB.BuiltInCategory.OST_Walls) \
                .WhereElementIsNotElementType().ToElements()
            return [e for e in collector if isinstance(e, DB.Wall)]
        except Exception:
            return []

    def _get_active_walls(self):
        if self.UI_rb_batch.IsChecked:
            return self._get_view_walls()
        else:
            return self._get_selection_walls()

    def _get_filtered_walls(self, selected_types):
        walls = self._get_active_walls()
        return [w for w in walls if get_wall_type_name(w) in selected_types]

    def _get_mark_mode(self):
        if self.UI_rb_mark_type.IsChecked:
            return "type"
        elif self.UI_rb_mark_custom.IsChecked:
            return "custom"
        else:
            return "none"


# ── Validation ───────────────────────────────────────────────────────────────
def validate():
    view = doc.ActiveView
    if view is None:
        forms.alert("No active view.", title=TOOL_NAME, exitscript=True)
    allowed = [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.EngineeringPlan,
        DB.ViewType.AreaPlan,
    ]
    if view.ViewType not in allowed:
        forms.alert(
            "This tool works in plan views only.\n\n"
            "Current view type: {}".format(view.ViewType),
            title=TOOL_NAME, exitscript=True)
    return view


# ── Main Logic ───────────────────────────────────────────────────────────────
def run_placement(settings, view):
    """Execute placement, mark, and tagging based on window settings."""
    walls = settings["walls"]
    detail_symbol = settings["detail_symbol"]
    tag_type = settings["tag_type"]
    tag_type_id = tag_type.Id if tag_type else None
    mark_mode = settings["mark_mode"]
    custom_mark = settings["custom_mark"]

    walls_ok = 0
    walls_skipped = []
    instances_placed = 0
    marks_set = 0
    tags_placed = 0
    placed_data = []

    # ── Pass 1: Place + Mark ─────────────────────────────────────────
    t1, t1_ok = safe_start("Place Finish Detail Items")
    if not t1_ok:
        forms.alert("Could not start transaction.", title=TOOL_NAME)
        return

    try:
        try:
            if not detail_symbol.IsActive:
                detail_symbol.Activate()
                doc.Regenerate()
        except Exception:
            pass

        for wall in walls:
            try:
                ext_line, int_line = get_wall_face_lines(wall)
                if ext_line is None and int_line is None:
                    walls_skipped.append(eid_int(wall.Id))
                    continue

                # Resolve mark value
                mark_value = None
                if mark_mode == "type":
                    mark_value = get_wall_type_name(wall)
                elif mark_mode == "custom":
                    mark_value = custom_mark

                wc = wall.Location.Curve
                w_dir = (wc.GetEndPoint(1) - wc.GetEndPoint(0)).Normalize()
                w_normal = DB.XYZ(-w_dir.Y, w_dir.X, 0).Normalize()

                wall_count = 0
                face_configs = [
                    (ext_line, w_normal),
                    (int_line, DB.XYZ(-w_normal.X, -w_normal.Y, 0)),
                ]

                for face_line, tag_normal in face_configs:
                    if face_line is None:
                        continue
                    try:
                        inst = place_line_based_instance(
                            view, face_line, detail_symbol)
                        if inst is not None:
                            if mark_value is not None:
                                if set_mark_parameter(inst, mark_value):
                                    marks_set += 1
                            placed_data.append((inst, tag_normal))
                            instances_placed += 1
                            wall_count += 1
                    except Exception:
                        pass

                if wall_count > 0:
                    walls_ok += 1
                else:
                    walls_skipped.append(eid_int(wall.Id))

            except Exception as we:
                logger.warning("Wall {}: {}".format(
                    eid_int(wall.Id), str(we)))
                walls_skipped.append(eid_int(wall.Id))

        if not safe_commit(t1):
            forms.alert("Placement transaction failed.", title=TOOL_NAME)
            return

    except Exception as e1:
        safe_rollback(t1)
        forms.alert("Placement error:\n{}".format(str(e1)), title=TOOL_NAME)
        return

    # ── Pass 2: Tag ──────────────────────────────────────────────────
    if tag_type_id is not None and placed_data:
        t2, t2_ok = safe_start("Tag Finish Detail Items")
        if t2_ok:
            try:
                for inst, t_normal in placed_data:
                    try:
                        if inst is not None and inst.IsValidObject:
                            tag = tag_instance(
                                view, inst, tag_type_id, t_normal)
                            if tag is not None:
                                tags_placed += 1
                    except Exception:
                        pass
                safe_commit(t2)
            except Exception:
                safe_rollback(t2)

    # ── Summary ──────────────────────────────────────────────────────
    lines = ["{} detail items placed on {} wall(s).".format(
        instances_placed, walls_ok)]
    if marks_set > 0:
        lines.append("{} Mark values set.".format(marks_set))
    if tags_placed > 0:
        lines.append("{} tags placed.".format(tags_placed))
    elif tag_type_id is not None:
        lines.append("No tags could be placed.")
    if walls_skipped:
        lines.append("{} wall(s) skipped.".format(len(walls_skipped)))
    forms.alert("\n".join(lines), title=TOOL_NAME)


# ── Entry Point ──────────────────────────────────────────────────────────────
def main():
    view = validate()

    # Show WPF window
    window = FinishLinesWindow(view)
    window.ShowDialog()

    if window.result is None:
        return  # user cancelled

    if len(window.result["walls"]) > MAX_BATCH_WALLS:
        proceed = forms.alert(
            "{} walls selected. This may take a while.\n\nContinue?".format(
                len(window.result["walls"])),
            title=TOOL_NAME, yes=True, no=True)
        if not proceed:
            return

    run_placement(window.result, view)


main()
