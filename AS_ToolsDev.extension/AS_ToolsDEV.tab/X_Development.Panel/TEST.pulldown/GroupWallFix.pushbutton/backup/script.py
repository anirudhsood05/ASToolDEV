# -*- coding: utf-8 -*-
"""
Script:   Fix Wall Top Constraints in Groups
Desc:     Resets wall top constraints inside model groups to "Unconnected",
          preserving current wall height. Best practice: walls in groups
          should use base constraint only with an unconnected height.
Author:   Anirudh Sood
Usage:    Click button from pyRevit toolbar - no pre-selection needed.
Result:   WPF window showing walls with top constraints inside groups;
          preview and batch-fix actions.
"""
__title__ = "Fix Wall\nTop Constraints"
__author__ = "Anirudh Sood"

import clr
import traceback
from System.Collections.Generic import List

# pyRevit imports
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import script

# WPF imports
from pyrevit.framework import wpf
from System import Windows
from System.IO import StringReader

# ── Setup ────────────────────────────────────────────────────────────────────
logger = script.get_logger()
output = script.get_output()
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ── Constants ────────────────────────────────────────────────────────────────
MAX_WALLS = 5000
MAX_HIGHLIGHT = 200
FEET_PER_MM = 1.0 / 304.8


# ── Version-Safe Helpers ─────────────────────────────────────────────────────
def eid_int(element_id):
    """Return integer value of ElementId. Revit 2023-2026 compatible."""
    try:
        return element_id.Value           # Revit 2026+
    except AttributeError:
        return element_id.IntegerValue    # Revit 2023-2025


INVALID_ID = eid_int(DB.ElementId.InvalidElementId)


def safe_element_name(elem):
    """Safely get element name via GetValue pattern."""
    try:
        if elem is not None:
            return DB.Element.Name.GetValue(elem) or "Unnamed"
    except Exception:
        pass
    return "Unnamed"


def safe_param_double(elem, bip):
    """Safely read a double parameter value, returning 0.0 on failure."""
    try:
        p = elem.get_Parameter(bip)
        if p is not None and p.HasValue:
            return p.AsDouble()
    except Exception:
        pass
    return 0.0


def safe_param_element_id(elem, bip):
    """Safely read an ElementId parameter, returning InvalidElementId on failure."""
    try:
        p = elem.get_Parameter(bip)
        if p is not None and p.HasValue:
            return p.AsElementId()
    except Exception:
        pass
    return DB.ElementId.InvalidElementId


def get_level_elevation(level_id):
    """Get the elevation of a level by its ElementId."""
    try:
        level = doc.GetElement(level_id)
        if level is not None and isinstance(level, DB.Level):
            return level.Elevation
    except Exception:
        pass
    return 0.0


# ── Data Collection ──────────────────────────────────────────────────────────
class WallConstraintInfo(object):
    """Stores info about a wall with a top constraint inside a group."""

    def __init__(self, wall, group):
        self.wall_id = eid_int(wall.Id)
        self.group_id = eid_int(group.Id)
        self.group_name = safe_element_name(group)
        self.wall_type_name = safe_element_name(wall.WallType) if wall.WallType else "Unknown"

        # Read current constraint values
        self.top_constraint_id = safe_param_element_id(wall, DB.BuiltInParameter.WALL_HEIGHT_TYPE)
        self.base_constraint_id = safe_param_element_id(wall, DB.BuiltInParameter.WALL_BASE_CONSTRAINT)
        self.top_offset = safe_param_double(wall, DB.BuiltInParameter.WALL_TOP_OFFSET)
        self.base_offset = safe_param_double(wall, DB.BuiltInParameter.WALL_BASE_OFFSET)
        self.unconnected_height = safe_param_double(wall, DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)

        # Resolve top constraint name
        top_elem = doc.GetElement(self.top_constraint_id) if self.top_constraint_id else None
        self.top_constraint_name = safe_element_name(top_elem) if top_elem else "Unconnected"

        # Calculate the actual wall height (internal feet)
        self.computed_height = self._compute_height()

    def _compute_height(self):
        """Compute the wall's actual unconnected height from its constraints."""
        top_id = self.top_constraint_id
        base_id = self.base_constraint_id

        # If already unconnected, use existing unconnected height
        if top_id is None or eid_int(top_id) == INVALID_ID:
            return self.unconnected_height if self.unconnected_height > 0 else 3000.0 * FEET_PER_MM

        # Compute from level elevations + offsets
        top_elev = get_level_elevation(top_id)
        base_elev = get_level_elevation(base_id)
        height = (top_elev + self.top_offset) - (base_elev + self.base_offset)

        # Fallback: if computed height is invalid, use existing or default
        if height <= 0:
            if self.unconnected_height > 0:
                return self.unconnected_height
            return 3000.0 * FEET_PER_MM  # 3m fallback

        return height

    @property
    def computed_height_mm(self):
        return self.computed_height / FEET_PER_MM

    @property
    def has_top_constraint(self):
        return (self.top_constraint_id is not None
                and eid_int(self.top_constraint_id) != INVALID_ID)


def collect_walls_in_groups():
    """Find all walls inside model groups that have a top constraint set."""
    results = []
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Group)
        for group in collector:
            try:
                # Filter to model groups only
                gt = group.GroupType
                if gt is None or gt.Category is None:
                    continue
                if eid_int(gt.Category.Id) != int(DB.BuiltInCategory.OST_IOSModelGroups):
                    continue

                member_ids = group.GetMemberIds()
                if not member_ids:
                    continue

                for mid in member_ids:
                    try:
                        elem = doc.GetElement(mid)
                        if elem is None or not isinstance(elem, DB.Wall):
                            continue
                        info = WallConstraintInfo(elem, group)
                        if info.has_top_constraint:
                            results.append(info)
                            if len(results) >= MAX_WALLS:
                                return results
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception as e:
        logger.error("Collection error: {}".format(str(e)))
    return results


# ── Fix Logic ────────────────────────────────────────────────────────────────
def fix_wall_top_constraint(wall_info):
    """Set a wall's top constraint to Unconnected with computed height.
    Must be called inside a transaction.
    Returns (success, message)."""
    try:
        wall = doc.GetElement(DB.ElementId(wall_info.wall_id))
        if wall is None:
            return False, "Wall no longer exists"

        # Set top constraint to "Unconnected" (InvalidElementId)
        top_param = wall.get_Parameter(DB.BuiltInParameter.WALL_HEIGHT_TYPE)
        if top_param is None or top_param.IsReadOnly:
            return False, "Top constraint parameter is read-only"
        top_param.Set(DB.ElementId.InvalidElementId)

        # Set unconnected height to preserve actual wall height
        height_param = wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        if height_param is not None and not height_param.IsReadOnly:
            height_param.Set(wall_info.computed_height)

        # Reset top offset to 0 (no longer relevant with unconnected height)
        offset_param = wall.get_Parameter(DB.BuiltInParameter.WALL_TOP_OFFSET)
        if offset_param is not None and not offset_param.IsReadOnly:
            offset_param.Set(0.0)

        return True, "Set to Unconnected ({:.0f} mm)".format(wall_info.computed_height_mm)
    except Exception as e:
        return False, str(e)


# ── WPF ListView Item ────────────────────────────────────────────────────────
class ListItem(object):
    """Data carrier for WPF ListView binding."""
    def __init__(self, **kwargs):
        for k, v in kwargs.items():
            setattr(self, k, v)


# ── XAML ─────────────────────────────────────────────────────────────────────
XAML_MAIN = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AUK Fix Wall Top Constraints in Groups" Width="750" Height="600"
    ShowInTaskbar="True" WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize" MinWidth="700" MinHeight="550"
    FontFamily="Arial" FontSize="12" Background="#F5F5F5" Topmost="False">

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#375D6D" Padding="12" CornerRadius="3">
            <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="Fix Wall Top Constraints in Groups"
                           FontWeight="Bold" FontSize="14"
                           HorizontalAlignment="Center" Foreground="#FFFFFF"/>
                <TextBlock Text="Set walls to Unconnected height, preserving actual height"
                           FontSize="10" HorizontalAlignment="Center" Foreground="#CCCCCC"
                           Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Status -->
        <Border Grid.Row="2" Background="#FFFFFF" Padding="10" CornerRadius="3"
                BorderBrush="#CCCCCC" BorderThickness="1">
            <StackPanel>
                <TextBlock x:Name="UI_status" Text="Click 'Scan' to find walls with top constraints..."
                           Foreground="#2C2C2C" FontWeight="Medium"/>
                <ProgressBar x:Name="UI_progress" Height="8" Margin="0,6,0,0"
                             Visibility="Collapsed" Background="#CCCCCC" Foreground="#375D6D"/>
            </StackPanel>
        </Border>

        <!-- Walls List -->
        <GroupBox Grid.Row="4" Header="Walls with Top Constraints in Groups" Padding="8"
                  Background="#FFFFFF" Foreground="#2C2C2C" FontWeight="SemiBold"
                  BorderBrush="#CCCCCC" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <ListView x:Name="UI_wall_list" Grid.Row="0" SelectionMode="Extended"
                          Background="#FFFFFF" Foreground="#2C2C2C"
                          BorderBrush="#CCCCCC" BorderThickness="1" FontWeight="Normal">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Header="Wall Type" Width="180"
                                            DisplayMemberBinding="{Binding WallType}"/>
                            <GridViewColumn Header="Top Constraint" Width="130"
                                            DisplayMemberBinding="{Binding TopConstraint}"/>
                            <GridViewColumn Header="Computed Height (mm)" Width="140"
                                            DisplayMemberBinding="{Binding HeightMM}"/>
                            <GridViewColumn Header="Group Name" Width="150"
                                            DisplayMemberBinding="{Binding GroupName}"/>
                            <GridViewColumn Header="Wall ID" Width="70"
                                            DisplayMemberBinding="{Binding WallID}"/>
                        </GridView>
                    </ListView.View>
                </ListView>

                <!-- Selection controls -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button x:Name="UI_select_all" Content="Select All"
                            Width="90" Height="25" Margin="0,0,8,0"
                            Background="#375D6D" Foreground="#FFFFFF" FontWeight="Medium"/>
                    <Button x:Name="UI_select_none" Content="Deselect All"
                            Width="90" Height="25" Margin="0,0,16,0"
                            Background="#CCCCCC" Foreground="#2C2C2C" FontWeight="Medium"/>
                    <Button x:Name="UI_highlight" Content="Highlight Selected"
                            Width="120" Height="25"
                            Background="#375D6D" Foreground="#FFFFFF" FontWeight="Medium"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- Footer Buttons -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="UI_scan" Content="Scan" IsDefault="True"
                    Width="100" Height="30" Margin="0,0,8,0"
                    Background="#375D6D" Foreground="#FFFFFF" FontWeight="Medium"/>
            <Button x:Name="UI_fix_selected" Content="Fix Selected"
                    Width="120" Height="30" Margin="0,0,8,0" IsEnabled="False"
                    Background="#E8A735" Foreground="#2C2C2C" FontWeight="Medium"/>
            <Button x:Name="UI_fix_all" Content="Fix All"
                    Width="100" Height="30" Margin="0,0,8,0" IsEnabled="False"
                    Background="#E8A735" Foreground="#2C2C2C" FontWeight="Medium"/>
            <Button x:Name="UI_close" Content="Close"
                    Width="90" Height="30"
                    Background="#CCCCCC" Foreground="#2C2C2C" FontWeight="Medium"/>
        </StackPanel>
    </Grid>
</Window>"""


# ── Main Window ──────────────────────────────────────────────────────────────
class FixWallConstraintsWindow(forms.WPFWindow):
    """Primary UI for fixing wall top constraints in model groups."""

    def __init__(self):
        xaml_stream = StringReader(XAML_MAIN)
        wpf.LoadComponent(self, xaml_stream)
        self.wall_infos = []
        self._connect_events()

    def _connect_events(self):
        self.UI_scan.Click += self._on_scan
        self.UI_fix_selected.Click += self._on_fix_selected
        self.UI_fix_all.Click += self._on_fix_all
        self.UI_close.Click += self._on_close
        self.UI_select_all.Click += self._on_select_all
        self.UI_select_none.Click += self._on_select_none
        self.UI_highlight.Click += self._on_highlight
        self.UI_wall_list.SelectionChanged += self._on_selection_changed

    # ── Status helpers ───────────────────────────────────────────────────
    def _set_status(self, msg, show_progress=False):
        self.UI_status.Text = msg
        self.UI_progress.Visibility = (
            Windows.Visibility.Visible if show_progress
            else Windows.Visibility.Collapsed
        )

    def _update_fix_buttons(self):
        has_items = self.UI_wall_list.Items.Count > 0
        has_selection = self.UI_wall_list.SelectedItems.Count > 0
        self.UI_fix_all.IsEnabled = has_items
        self.UI_fix_selected.IsEnabled = has_selection

    # ── Scan ─────────────────────────────────────────────────────────────
    def _on_scan(self, sender, args):
        try:
            self._set_status("Scanning model groups for walls with top constraints...", True)
            self.UI_scan.IsEnabled = False
            self.UI_wall_list.Items.Clear()
            self.wall_infos = []

            self.wall_infos = collect_walls_in_groups()

            if not self.wall_infos:
                self._set_status("No walls with top constraints found in model groups.")
                self._update_fix_buttons()
                return

            for info in self.wall_infos:
                self.UI_wall_list.Items.Add(ListItem(
                    WallType=info.wall_type_name,
                    TopConstraint=info.top_constraint_name,
                    HeightMM="{:.0f}".format(info.computed_height_mm),
                    GroupName=info.group_name,
                    WallID=str(info.wall_id),
                    Info=info,
                ))

            self._set_status("Found {} walls with top constraints in groups.".format(
                len(self.wall_infos)))
            self._update_fix_buttons()

        except Exception as e:
            logger.error("Scan error: {}\n{}".format(str(e), traceback.format_exc()))
            self._set_status("Error during scan.")
            forms.alert("Scan error:\n{}".format(str(e)), title="Scan Error")
        finally:
            self.UI_scan.IsEnabled = True

    # ── Fix operations ───────────────────────────────────────────────────
    def _do_fix(self, infos_to_fix):
        """Apply the fix to a list of WallConstraintInfo objects."""
        if not infos_to_fix:
            forms.alert("No walls to fix.", title="Nothing Selected")
            return

        count = len(infos_to_fix)
        if not forms.alert(
            "This will set {} wall(s) to Unconnected top constraint, "
            "preserving their current height.\n\n"
            "Walls inside groups will be modified. Revit may ungroup "
            "and regroup affected groups automatically.\n\n"
            "Continue?".format(count),
            title="Confirm Fix",
            yes=True, no=True
        ):
            return

        succeeded = 0
        failed = []

        try:
            with revit.TransactionGroup("Fix Wall Top Constraints"):
                with revit.Transaction("Set Walls to Unconnected"):
                    for info in infos_to_fix:
                        try:
                            ok, msg = fix_wall_top_constraint(info)
                            if ok:
                                succeeded += 1
                            else:
                                failed.append((info.wall_id, msg))
                        except Exception as e:
                            failed.append((info.wall_id, str(e)))

            # Report results
            summary = "Fixed {} of {} walls.".format(succeeded, count)
            if failed:
                summary += "\n\nFailed ({}):\n".format(len(failed))
                for wid, msg in failed[:20]:
                    summary += "  Wall {}: {}\n".format(wid, msg)
                if len(failed) > 20:
                    summary += "  ... and {} more\n".format(len(failed) - 20)

            self._set_status("Fix complete: {} succeeded, {} failed.".format(
                succeeded, len(failed)))
            forms.alert(summary, title="Fix Complete")

            # Re-scan to refresh the list
            if succeeded > 0:
                self._on_scan(None, None)

        except Exception as e:
            logger.error("Fix error: {}\n{}".format(str(e), traceback.format_exc()))
            forms.alert("Error during fix:\n{}".format(str(e)), title="Fix Error")

    def _on_fix_selected(self, sender, args):
        selected = []
        for item in self.UI_wall_list.SelectedItems:
            if hasattr(item, "Info") and item.Info is not None:
                selected.append(item.Info)
        self._do_fix(selected)

    def _on_fix_all(self, sender, args):
        self._do_fix(self.wall_infos)

    # ── Selection & Highlight ────────────────────────────────────────────
    def _on_selection_changed(self, sender, args):
        self._update_fix_buttons()

    def _on_select_all(self, sender, args):
        self.UI_wall_list.SelectAll()

    def _on_select_none(self, sender, args):
        self.UI_wall_list.UnselectAll()

    def _on_highlight(self, sender, args):
        ids = []
        for item in self.UI_wall_list.SelectedItems:
            if hasattr(item, "Info") and item.Info is not None:
                ids.append(item.Info.wall_id)
        if not ids:
            forms.alert("Select walls to highlight first.", title="No Selection")
            return

        try:
            revit_ids = List[DB.ElementId]()
            for id_int in ids[:MAX_HIGHLIGHT]:
                elem = doc.GetElement(DB.ElementId(id_int))
                if elem is not None:
                    revit_ids.Add(DB.ElementId(id_int))

            if revit_ids.Count > 0:
                uidoc.Selection.SetElementIds(revit_ids)
                t = DB.Transaction(doc, "Highlight Walls")
                t.Start()
                try:
                    uidoc.ShowElements(revit_ids)
                    t.Commit()
                except Exception:
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()
        except Exception as e:
            logger.warning("Highlight failed: {}".format(str(e)))

    def _on_close(self, sender, args):
        self.Close()


# ── Entry Point ──────────────────────────────────────────────────────────────
def main():
    if doc is None:
        forms.alert("No active Revit document.", exitscript=True)
        return

    try:
        window = FixWallConstraintsWindow()
        window.ShowDialog()
    except Exception as e:
        logger.error("Fatal: {}\n{}".format(str(e), traceback.format_exc()))
        forms.alert("Fatal error:\n{}".format(str(e)), title="Error")


if __name__ == "__main__":
    main()
