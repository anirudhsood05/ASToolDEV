# -*- coding: utf-8 -*-
"""
Script:   Fix Wall Top Constraints in Groups
Desc:     Resets wall top constraints inside model groups to "Unconnected",
          preserving current wall height. Uses ungroup-modify-regroup pattern
          on a single instance per group type, then re-places all other
          instances. Groups are fully preserved.
Author:   Anirudh Sood
Usage:    Click button from pyRevit toolbar - no pre-selection needed.
Result:   WPF window showing affected group types; preview and batch-fix.
"""
__title__ = "Fix Wall\nTop Constraints"
__author__ = "Anirudh Sood"

import clr
import traceback
from collections import OrderedDict
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


def get_group_type_name(group_type):
    """Get group type name via SYMBOL_NAME_PARAM (reliable across versions)."""
    try:
        p = group_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p is not None and p.HasValue:
            return p.AsString() or "Unnamed"
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
    """Safely read an ElementId parameter, returning InvalidElementId."""
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


# ── Wall Analysis ────────────────────────────────────────────────────────────
class WallConstraintInfo(object):
    """Stores info about a wall with a top constraint inside a group."""

    def __init__(self, wall, group_instance):
        self.wall_id = eid_int(wall.Id)
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

        # Calculate actual wall height (internal feet)
        self.computed_height = self._compute_height()

    def _compute_height(self):
        """Compute wall's actual unconnected height from constraints."""
        top_id = self.top_constraint_id
        base_id = self.base_constraint_id

        # If already unconnected, use existing unconnected height
        if top_id is None or eid_int(top_id) == INVALID_ID:
            return self.unconnected_height if self.unconnected_height > 0 else 3000.0 * FEET_PER_MM

        # Compute from level elevations + offsets
        top_elev = get_level_elevation(top_id)
        base_elev = get_level_elevation(base_id)
        height = (top_elev + self.top_offset) - (base_elev + self.base_offset)

        if height <= 0:
            if self.unconnected_height > 0:
                return self.unconnected_height
            return 3000.0 * FEET_PER_MM
        return height

    @property
    def computed_height_mm(self):
        return self.computed_height / FEET_PER_MM

    @property
    def has_top_constraint(self):
        return (self.top_constraint_id is not None
                and eid_int(self.top_constraint_id) != INVALID_ID)


class GroupTypeInfo(object):
    """Aggregates data for one group type: all its instances and affected walls."""

    def __init__(self, group_type_id, group_type_name):
        self.group_type_id = group_type_id
        self.group_type_name = group_type_name
        self.instance_ids = []          # list of int (group instance element ids)
        self.affected_walls = []        # list of WallConstraintInfo
        self.instance_count = 0
        self._seen_wall_types = set()
        self.has_rotated = False
        self.has_mirrored = False

    def add_instance(self, group_instance):
        gid = eid_int(group_instance.Id)
        if gid not in self.instance_ids:
            self.instance_ids.append(gid)
            self.instance_count = len(self.instance_ids)
            # Check rotation/mirror state
            try:
                loc = group_instance.Location
                if loc is not None and hasattr(loc, "Rotation"):
                    if abs(loc.Rotation) > 1e-6:
                        self.has_rotated = True
            except Exception:
                pass
            try:
                if group_instance.Mirrored:
                    self.has_mirrored = True
            except Exception:
                pass

    def add_wall(self, wall_info):
        """Add wall info, deduplicating by wall type name (same type
        definition means same constraint setup across all instances)."""
        key = wall_info.wall_type_name
        if key not in self._seen_wall_types:
            self._seen_wall_types.add(key)
            self.affected_walls.append(wall_info)

    @property
    def wall_count(self):
        return len(self.affected_walls)


def collect_affected_group_types():
    """Find all group types containing walls with top constraints.
    Returns OrderedDict of group_type_id -> GroupTypeInfo."""
    result = OrderedDict()
    seen_types = set()  # track which types already had walls scanned

    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Group)
        for group in collector:
            try:
                gt = group.GroupType
                if gt is None or gt.Category is None:
                    continue
                if eid_int(gt.Category.Id) != int(DB.BuiltInCategory.OST_IOSModelGroups):
                    continue

                gt_id = eid_int(gt.Id)

                # Ensure GroupTypeInfo exists
                if gt_id not in result:
                    result[gt_id] = GroupTypeInfo(gt_id, get_group_type_name(gt))

                gti = result[gt_id]
                gti.add_instance(group)

                # Only scan walls from the first instance of each type
                if gt_id not in seen_types:
                    seen_types.add(gt_id)
                    member_ids = group.GetMemberIds()
                    for mid in member_ids:
                        try:
                            elem = doc.GetElement(mid)
                            if elem is None or not isinstance(elem, DB.Wall):
                                continue
                            info = WallConstraintInfo(elem, group)
                            if info.has_top_constraint:
                                gti.add_wall(info)
                        except Exception:
                            continue
            except Exception:
                continue
    except Exception as e:
        logger.error("Collection error: {}".format(str(e)))

    # Filter to only types that have affected walls
    filtered = OrderedDict()
    for gt_id, gti in result.items():
        if gti.wall_count > 0:
            filtered[gt_id] = gti
    return filtered


# ── Fix Logic (Ungroup-Modify-Regroup per GroupType) ─────────────────────────
def fix_group_type(gti):
    """Fix all walls with top constraints in a single group type.

    Strategy:
      1. Collect all instances of this group type
      2. Record locations of ALL other instances, then delete them
      3. Ungroup the first instance -> loose elements
      4. Fix wall top constraints on the loose walls
      5. Regroup the loose elements -> new group (type)
      6. Re-place instances at all recorded locations

    Must be called inside a TransactionGroup (caller manages).
    Returns (success, message).
    """
    group_type_id_int = gti.group_type_id
    group_type_eid = DB.ElementId(group_type_id_int)
    group_type = doc.GetElement(group_type_eid)

    if group_type is None:
        return False, "Group type no longer exists"

    original_name = gti.group_type_name

    # ── 1. Collect all instances of this group type ──────────────────────
    all_instances = []
    instance_collector = DB.FilteredElementCollector(doc).OfClass(DB.Group)
    for g in instance_collector:
        try:
            if g.GroupType is not None and eid_int(g.GroupType.Id) == group_type_id_int:
                all_instances.append(g)
        except Exception:
            continue

    if not all_instances:
        return False, "No instances found"

    # ── 2. Pick the first instance to modify; record locations of others ─
    #    LocationPoint.Rotation is unreliable for groups. Instead we record
    #    the direction vector of a reference wall in each instance. After
    #    regroup we compare placed vs. target wall directions to compute
    #    the exact rotation needed.
    primary = all_instances[0]
    others = all_instances[1:]

    def get_ref_wall_direction(group_inst):
        """Get direction vector of the first wall in a group instance."""
        try:
            for mid in group_inst.GetMemberIds():
                elem = doc.GetElement(mid)
                if elem is not None and isinstance(elem, DB.Wall):
                    loc = elem.Location
                    if loc is not None and hasattr(loc, "Curve"):
                        curve = loc.Curve
                        start = curve.GetEndPoint(0)
                        end = curve.GetEndPoint(1)
                        direction = end.Subtract(start)
                        if direction.GetLength() > 1e-6:
                            return direction.Normalize()
        except Exception:
            pass
        return None

    # Record primary's reference wall direction and mirror state
    primary_wall_dir = get_ref_wall_direction(primary)
    primary_mirrored = False
    try:
        primary_mirrored = primary.Mirrored
    except Exception:
        pass

    other_data = []
    for inst in others:
        try:
            loc = inst.Location
            if loc is None or not hasattr(loc, "Point"):
                continue
            point = loc.Point
            if point is None:
                continue

            # Mirror state
            is_mirrored = False
            try:
                is_mirrored = inst.Mirrored
            except Exception:
                pass

            # Record reference wall direction for this instance
            wall_dir = get_ref_wall_direction(inst)

            # Delta mirror: only mirror if state differs from primary
            needs_mirror = (is_mirrored != primary_mirrored)

            other_data.append({
                "point": DB.XYZ(point.X, point.Y, point.Z),
                "level_id": inst.LevelId,
                "wall_dir": wall_dir,
                "needs_mirror": needs_mirror,
            })
        except Exception:
            continue

    # ── 3. Delete all OTHER instances (Transaction 1) ───────────────────
    with revit.Transaction("Remove duplicate group instances"):
        for inst in others:
            try:
                doc.Delete(inst.Id)
            except Exception:
                pass
        doc.Regenerate()

    # ── 4. Ungroup the primary instance (Transaction 2) ─────────────────
    ungrouped_ids = None
    with revit.Transaction("Ungroup for editing"):
        try:
            ungrouped_ids = primary.UngroupMembers()
            doc.Regenerate()
        except Exception as e:
            return False, "Ungroup failed: {}".format(str(e))

    if not ungrouped_ids:
        return False, "Ungroup returned no elements"

    # ── 5. Fix wall top constraints on loose elements (Transaction 3) ───
    walls_fixed = 0
    walls_failed = 0
    with revit.Transaction("Fix wall constraints"):
        for eid in ungrouped_ids:
            try:
                elem = doc.GetElement(eid)
                if elem is None or not isinstance(elem, DB.Wall):
                    continue

                # Check if wall has a top constraint
                top_param = elem.get_Parameter(DB.BuiltInParameter.WALL_HEIGHT_TYPE)
                if top_param is None:
                    continue
                top_id = top_param.AsElementId()
                if top_id is None or eid_int(top_id) == INVALID_ID:
                    continue

                # Compute height before clearing constraint
                base_id = safe_param_element_id(elem, DB.BuiltInParameter.WALL_BASE_CONSTRAINT)
                top_offset = safe_param_double(elem, DB.BuiltInParameter.WALL_TOP_OFFSET)
                base_offset = safe_param_double(elem, DB.BuiltInParameter.WALL_BASE_OFFSET)
                existing_height = safe_param_double(elem, DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)

                top_elev = get_level_elevation(top_id)
                base_elev = get_level_elevation(base_id)
                computed = (top_elev + top_offset) - (base_elev + base_offset)

                if computed <= 0:
                    computed = existing_height if existing_height > 0 else 3000.0 * FEET_PER_MM

                # Set top constraint to Unconnected
                if not top_param.IsReadOnly:
                    top_param.Set(DB.ElementId.InvalidElementId)

                # Set unconnected height to preserve actual wall height
                height_param = elem.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
                if height_param is not None and not height_param.IsReadOnly:
                    height_param.Set(computed)

                # Reset top offset to 0 (no longer relevant)
                offset_param = elem.get_Parameter(DB.BuiltInParameter.WALL_TOP_OFFSET)
                if offset_param is not None and not offset_param.IsReadOnly:
                    offset_param.Set(0.0)

                walls_fixed += 1
            except Exception as e:
                walls_failed += 1
                logger.warning("Wall fix failed: {}".format(str(e)))
        doc.Regenerate()

    # ── 6. Regroup the loose elements (Transaction 4) ───────────────────
    new_group = None
    with revit.Transaction("Regroup elements"):
        try:
            id_collection = List[DB.ElementId](ungrouped_ids)
            new_group = doc.Create.NewGroup(id_collection)
            doc.Regenerate()
        except Exception as e:
            return False, "Regroup failed: {}".format(str(e))

    if new_group is None:
        return False, "Regroup returned None"

    # ── 7. Rename new group type to match original name (Transaction 5) ─
    new_group_type = new_group.GroupType
    with revit.Transaction("Rename group type"):
        try:
            new_group_type.Name = original_name
        except Exception:
            # Name collision is acceptable - Revit may have kept the name
            pass
        doc.Regenerate()

    # ── 8. Re-place all other instances (Transaction 6) ─────────────────
    #    After placing each new instance, compare the reference wall
    #    direction in the placed group vs. the recorded target direction
    #    and rotate to match. This is far more reliable than using
    #    LocationPoint.Rotation values.
    import math

    def get_placed_wall_direction(group_inst):
        """Get direction of first wall in a newly placed group instance."""
        try:
            for mid in group_inst.GetMemberIds():
                elem = doc.GetElement(mid)
                if elem is not None and isinstance(elem, DB.Wall):
                    loc = elem.Location
                    if loc is not None and hasattr(loc, "Curve"):
                        curve = loc.Curve
                        s = curve.GetEndPoint(0)
                        e = curve.GetEndPoint(1)
                        d = e.Subtract(s)
                        if d.GetLength() > 1e-6:
                            return d.Normalize()
        except Exception:
            pass
        return None

    def angle_between_vectors_2d(v1, v2):
        """Signed angle from v1 to v2 in the XY plane (radians)."""
        return math.atan2(
            v1.X * v2.Y - v1.Y * v2.X,
            v1.X * v2.X + v1.Y * v2.Y
        )

    placed_count = 0
    mirror_count = 0
    rotate_count = 0
    with revit.Transaction("Re-place group instances"):
        for data in other_data:
            try:
                pt = data["point"]
                needs_mirror = data.get("needs_mirror", False)
                target_dir = data.get("wall_dir", None)

                if not needs_mirror:
                    # ── Non-mirrored: place, measure, rotate ─────────
                    new_inst = doc.Create.PlaceGroup(pt, new_group_type)
                    if new_inst is None:
                        continue

                    # Measure actual wall direction in placed group
                    if target_dir is not None:
                        placed_dir = get_placed_wall_direction(new_inst)
                        if placed_dir is not None:
                            angle = angle_between_vectors_2d(placed_dir, target_dir)
                            if abs(angle) > 1e-4:
                                axis = DB.Line.CreateBound(
                                    pt, DB.XYZ(pt.X, pt.Y, pt.Z + 1.0))
                                DB.ElementTransformUtils.RotateElement(
                                    doc, new_inst.Id, axis, angle)
                                rotate_count += 1

                    placed_count += 1

                else:
                    # ── Mirrored: place, mirror, move back, rotate ───
                    new_inst = doc.Create.PlaceGroup(pt, new_group_type)
                    if new_inst is None:
                        continue

                    # Snapshot group IDs before mirror
                    ids_before = set()
                    for g in DB.FilteredElementCollector(doc).OfClass(DB.Group):
                        ids_before.add(eid_int(g.Id))

                    # Mirror about X-axis plane through insertion point
                    mirror_plane = DB.Plane.CreateByNormalAndOrigin(
                        DB.XYZ(1, 0, 0), pt)
                    try:
                        DB.ElementTransformUtils.MirrorElement(
                            doc, new_inst.Id, mirror_plane)
                    except Exception as e:
                        logger.warning(
                            "Mirror failed at ({:.1f}, {:.1f}): {}".format(
                                pt.X, pt.Y, str(e)))
                        placed_count += 1
                        continue

                    # Delete un-mirrored original
                    doc.Delete(new_inst.Id)

                    # Find mirrored copy by diffing IDs
                    mirrored_id = None
                    for g in DB.FilteredElementCollector(doc).OfClass(DB.Group):
                        gid = eid_int(g.Id)
                        if gid not in ids_before:
                            mirrored_id = g.Id
                            break

                    if mirrored_id is not None:
                        mirrored_elem = doc.GetElement(mirrored_id)
                        if mirrored_elem is not None:
                            # Move mirrored copy to intended point
                            mloc = mirrored_elem.Location
                            if mloc is not None and hasattr(mloc, "Point"):
                                actual_pt = mloc.Point
                                delta_move = DB.XYZ(
                                    pt.X - actual_pt.X,
                                    pt.Y - actual_pt.Y,
                                    pt.Z - actual_pt.Z)
                                if delta_move.GetLength() > 1e-6:
                                    DB.ElementTransformUtils.MoveElement(
                                        doc, mirrored_id, delta_move)

                            # Rotate to match target wall direction
                            if target_dir is not None:
                                placed_dir = get_placed_wall_direction(mirrored_elem)
                                if placed_dir is not None:
                                    angle = angle_between_vectors_2d(
                                        placed_dir, target_dir)
                                    if abs(angle) > 1e-4:
                                        axis = DB.Line.CreateBound(
                                            pt, DB.XYZ(pt.X, pt.Y, pt.Z + 1.0))
                                        DB.ElementTransformUtils.RotateElement(
                                            doc, mirrored_id, axis, angle)
                                        rotate_count += 1

                    mirror_count += 1
                    placed_count += 1

            except Exception as e:
                logger.warning("Failed to place instance: {}".format(str(e)))
        doc.Regenerate()

    msg = "Fixed {} walls. Re-placed {} of {} instances".format(
        walls_fixed, placed_count, len(other_data))
    if rotate_count > 0:
        msg += " ({} rotated".format(rotate_count)
        if mirror_count > 0:
            msg += ", {} mirrored".format(mirror_count)
        msg += ")"
    elif mirror_count > 0:
        msg += " ({} mirrored)".format(mirror_count)
    else:
        msg += "."
    if walls_failed > 0:
        msg += " {} walls failed.".format(walls_failed)

    return True, msg


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
    Title="AUK Fix Wall Top Constraints in Groups" Width="860" Height="620"
    ShowInTaskbar="True" WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize" MinWidth="800" MinHeight="560"
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
            <RowDefinition Height="8"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#375D6D" Padding="12" CornerRadius="3">
            <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="Fix Wall Top Constraints in Groups"
                           FontWeight="Bold" FontSize="14"
                           HorizontalAlignment="Center" Foreground="#FFFFFF"/>
                <TextBlock Text="Sets walls to Unconnected height — preserves groups and all instances"
                           FontSize="10" HorizontalAlignment="Center" Foreground="#CCCCCC"
                           Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Status -->
        <Border Grid.Row="2" Background="#FFFFFF" Padding="10" CornerRadius="3"
                BorderBrush="#CCCCCC" BorderThickness="1">
            <StackPanel>
                <TextBlock x:Name="UI_status" Text="Click 'Scan' to find group types with wall top constraints..."
                           Foreground="#2C2C2C" FontWeight="Medium"/>
                <ProgressBar x:Name="UI_progress" Height="8" Margin="0,6,0,0"
                             Visibility="Collapsed" Background="#CCCCCC" Foreground="#375D6D"/>
            </StackPanel>
        </Border>

        <!-- Group Types List -->
        <GroupBox Grid.Row="4" Header="Group Types with Affected Walls" Padding="8"
                  Background="#FFFFFF" Foreground="#2C2C2C" FontWeight="SemiBold"
                  BorderBrush="#CCCCCC" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <ListView x:Name="UI_group_list" Grid.Row="0" SelectionMode="Extended"
                          Background="#FFFFFF" Foreground="#2C2C2C"
                          BorderBrush="#CCCCCC" BorderThickness="1" FontWeight="Normal">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Header="Group Type Name" Width="220"
                                            DisplayMemberBinding="{Binding GroupTypeName}"/>
                            <GridViewColumn Header="Instances" Width="80"
                                            DisplayMemberBinding="{Binding InstanceCount}"/>
                            <GridViewColumn Header="Walls to Fix" Width="90"
                                            DisplayMemberBinding="{Binding WallCount}"/>
                            <GridViewColumn Header="Wall Types Affected" Width="250"
                                            DisplayMemberBinding="{Binding WallTypes}"/>
                            <GridViewColumn Header="Type ID" Width="70"
                                            DisplayMemberBinding="{Binding TypeID}"/>
                            <GridViewColumn Header="Transforms" Width="80"
                                            DisplayMemberBinding="{Binding Transforms}"/>
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
                    <Button x:Name="UI_highlight" Content="Highlight Instances"
                            Width="130" Height="25"
                            Background="#375D6D" Foreground="#FFFFFF" FontWeight="Medium"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- Warning / Info -->
        <Border Grid.Row="6" Background="#FFF8E1" Padding="8" CornerRadius="3"
                BorderBrush="#E8A735" BorderThickness="1">
            <TextBlock TextWrapping="Wrap" Foreground="#2C2C2C" FontSize="11">
                <Run FontWeight="SemiBold">How it works:</Run>
                <Run> For each group type, one instance is ungrouped, walls are fixed, elements are regrouped, and all other instances are re-placed at their original locations. Attached detail groups and hosted elements on walls within groups may need manual review after fixing. This operation can be undone (Ctrl+Z).</Run>
            </TextBlock>
        </Border>

        <!-- Footer Buttons -->
        <StackPanel Grid.Row="8" Orientation="Horizontal" HorizontalAlignment="Right">
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
        self.group_type_infos = OrderedDict()
        self._connect_events()

    def _connect_events(self):
        self.UI_scan.Click += self._on_scan
        self.UI_fix_selected.Click += self._on_fix_selected
        self.UI_fix_all.Click += self._on_fix_all
        self.UI_close.Click += self._on_close
        self.UI_select_all.Click += self._on_select_all
        self.UI_select_none.Click += self._on_select_none
        self.UI_highlight.Click += self._on_highlight
        self.UI_group_list.SelectionChanged += self._on_selection_changed

    # ── Status helpers ───────────────────────────────────────────────────
    def _set_status(self, msg, show_progress=False):
        self.UI_status.Text = msg
        self.UI_progress.Visibility = (
            Windows.Visibility.Visible if show_progress
            else Windows.Visibility.Collapsed
        )

    def _update_fix_buttons(self):
        has_items = self.UI_group_list.Items.Count > 0
        has_selection = self.UI_group_list.SelectedItems.Count > 0
        self.UI_fix_all.IsEnabled = has_items
        self.UI_fix_selected.IsEnabled = has_selection

    # ── Scan ─────────────────────────────────────────────────────────────
    def _on_scan(self, sender, args):
        try:
            self._set_status("Scanning model groups...", True)
            self.UI_scan.IsEnabled = False
            self.UI_group_list.Items.Clear()
            self.group_type_infos = OrderedDict()

            self.group_type_infos = collect_affected_group_types()

            if not self.group_type_infos:
                self._set_status("No walls with top constraints found in any model groups.")
                self._update_fix_buttons()
                return

            for gt_id, gti in self.group_type_infos.items():
                wall_types_str = ", ".join(w.wall_type_name for w in gti.affected_walls)
                transforms = []
                if gti.has_rotated:
                    transforms.append("R")
                if gti.has_mirrored:
                    transforms.append("M")
                transforms_str = "+".join(transforms) if transforms else "-"
                self.UI_group_list.Items.Add(ListItem(
                    GroupTypeName=gti.group_type_name,
                    InstanceCount=str(gti.instance_count),
                    WallCount=str(gti.wall_count),
                    WallTypes=wall_types_str,
                    TypeID=str(gt_id),
                    Transforms=transforms_str,
                    Info=gti,
                ))

            total_walls = sum(gti.wall_count for gti in self.group_type_infos.values())
            total_instances = sum(gti.instance_count for gti in self.group_type_infos.values())
            self._set_status("Found {} group types ({} instances) with {} unique wall constraints.".format(
                len(self.group_type_infos), total_instances, total_walls))
            self._update_fix_buttons()

        except Exception as e:
            logger.error("Scan error: {}\n{}".format(str(e), traceback.format_exc()))
            self._set_status("Error during scan.")
            forms.alert("Scan error:\n{}".format(str(e)), title="Scan Error")
        finally:
            self.UI_scan.IsEnabled = True

    # ── Fix operations ───────────────────────────────────────────────────
    def _get_selected_infos(self):
        selected = []
        for item in self.UI_group_list.SelectedItems:
            if hasattr(item, "Info") and item.Info is not None:
                selected.append(item.Info)
        return selected

    def _do_fix(self, infos_to_fix):
        """Fix wall constraints for a list of GroupTypeInfo objects."""
        if not infos_to_fix:
            forms.alert("No group types selected.", title="Nothing Selected")
            return

        total_instances = sum(gti.instance_count for gti in infos_to_fix)
        total_walls = sum(gti.wall_count for gti in infos_to_fix)

        if not forms.alert(
            "This will fix {} group type(s) containing {} wall constraint(s) "
            "across {} instance(s).\n\n"
            "Process per group type:\n"
            "  1. Delete extra instances (locations recorded)\n"
            "  2. Ungroup the remaining instance\n"
            "  3. Fix wall top constraints to Unconnected\n"
            "  4. Regroup elements\n"
            "  5. Re-place all instances at original locations\n\n"
            "This operation can be undone (Ctrl+Z). Continue?".format(
                len(infos_to_fix), total_walls, total_instances),
            title="Confirm Fix",
            yes=True, no=True
        ):
            return

        results = []
        tg = DB.TransactionGroup(doc, "Fix Wall Top Constraints in Groups")
        tg.Start()

        try:
            for i, gti in enumerate(infos_to_fix):
                self._set_status("Fixing group type {} of {}: '{}'...".format(
                    i + 1, len(infos_to_fix), gti.group_type_name), True)
                try:
                    ok, msg = fix_group_type(gti)
                    results.append((gti.group_type_name, ok, msg))
                except Exception as e:
                    results.append((gti.group_type_name, False, str(e)))
                    logger.error("Fix failed for '{}': {}\n{}".format(
                        gti.group_type_name, str(e), traceback.format_exc()))

            tg.Assimilate()
        except Exception as e:
            if tg.HasStarted():
                tg.RollBack()
            logger.error("TransactionGroup failed: {}".format(str(e)))
            forms.alert("Operation failed and was rolled back:\n{}".format(str(e)),
                        title="Fix Error")
            self._set_status("Fix failed - rolled back.")
            return

        # Report results
        succeeded = sum(1 for _, ok, _ in results if ok)
        failed = sum(1 for _, ok, _ in results if not ok)

        summary = "Completed: {} succeeded, {} failed.\n\n".format(succeeded, failed)
        for name, ok, msg in results:
            status = "OK" if ok else "FAILED"
            summary += "  [{}] {}: {}\n".format(status, name, msg)

        self._set_status("Fix complete: {} succeeded, {} failed.".format(succeeded, failed))
        forms.alert(summary, title="Fix Complete")

        # Re-scan to refresh
        if succeeded > 0:
            self._on_scan(None, None)

    def _on_fix_selected(self, sender, args):
        self._do_fix(self._get_selected_infos())

    def _on_fix_all(self, sender, args):
        self._do_fix(list(self.group_type_infos.values()))

    # ── Selection & Highlight ────────────────────────────────────────────
    def _on_selection_changed(self, sender, args):
        self._update_fix_buttons()

    def _on_select_all(self, sender, args):
        self.UI_group_list.SelectAll()

    def _on_select_none(self, sender, args):
        self.UI_group_list.UnselectAll()

    def _on_highlight(self, sender, args):
        selected = self._get_selected_infos()
        if not selected:
            forms.alert("Select group types to highlight.", title="No Selection")
            return

        ids = []
        for gti in selected:
            ids.extend(gti.instance_ids)

        try:
            revit_ids = List[DB.ElementId]()
            for id_int in ids[:MAX_HIGHLIGHT]:
                elem = doc.GetElement(DB.ElementId(id_int))
                if elem is not None:
                    revit_ids.Add(DB.ElementId(id_int))

            if revit_ids.Count > 0:
                uidoc.Selection.SetElementIds(revit_ids)
                t = DB.Transaction(doc, "Highlight Group Instances")
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
