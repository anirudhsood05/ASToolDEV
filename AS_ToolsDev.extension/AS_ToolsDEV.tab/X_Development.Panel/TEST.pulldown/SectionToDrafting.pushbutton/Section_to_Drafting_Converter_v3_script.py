# -*- coding: utf-8 -*-
"""
Script:   Section to Drafting View Converter v3
Desc:     Recreates all 2D annotation/detail items from selected section views
          into new corresponding drafting views. Uses element recreation (not
          CopyElements) because the view-to-view overload silently fails when
          copying between different view types (Section -> Drafting).
Author:   AUK BIM Team
Usage:    Select section views in project browser or pre-select in canvas,
          then click button. Or run with no selection to pick from list.
Result:   One new drafting view per section, containing recreated 2D content.
          Summary report in pyRevit output window.

v3 CHANGES:
  - FilledRegion: validate typeId, fallback to default FilledRegionType
  - Group member failures logged to output (visible) not just debug
  - Improved arc projection with better degenerate detection
  - NewFamilyInstance for line-based: try Reference overload as second fallback
"""
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms
from System.Collections.Generic import List as CsList
import math

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc
uidoc  = revit.uidoc

# -- Constants -----------------------------------------------------------------

DRAFTING_SUFFIX = " (Drafting)"

SKIP_CATEGORIES = set([
    int(DB.BuiltInCategory.OST_SpotElevations),
    int(DB.BuiltInCategory.OST_SpotCoordinates),
    int(DB.BuiltInCategory.OST_SpotSlopes),
    int(DB.BuiltInCategory.OST_Dimensions),
    int(DB.BuiltInCategory.OST_KeynoteTags),
])


# -- Version-safe helpers ------------------------------------------------------

def eid_int(eid):
    """Version-safe ElementId integer value (Revit 2023-2026)."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def get_cat_int(element):
    """Return integer value of element's category Id, or None."""
    if element.Category is None:
        return None
    try:
        return int(element.Category.Id.IntegerValue)
    except AttributeError:
        try:
            return int(element.Category.Id.Value)
        except Exception:
            return None


# -- View name / type helpers --------------------------------------------------

def get_unique_view_name(base_name, existing_names):
    if base_name not in existing_names:
        return base_name
    counter = 1
    while True:
        candidate = "{} ({})".format(base_name, counter)
        if candidate not in existing_names:
            return candidate
        counter += 1


def get_drafting_view_family_type():
    vfts = DB.FilteredElementCollector(doc)\
             .OfClass(DB.ViewFamilyType).ToElements()
    for vft in vfts:
        if vft.ViewFamily == DB.ViewFamily.Drafting:
            return vft
    return None


def get_default_filled_region_type_id():
    """
    Get a valid FilledRegionType Id. Tries the document default first,
    then falls back to the first available FilledRegionType.
    """
    # Try document default
    try:
        default_id = doc.GetDefaultElementTypeId(
            DB.ElementTypeGroup.FilledRegionType)
        if default_id and default_id != DB.ElementId.InvalidElementId:
            return default_id
    except Exception:
        pass

    # Fallback: first available
    frt_collector = DB.FilteredElementCollector(doc)\
                      .OfClass(DB.FilledRegionType).ToElements()
    for frt in frt_collector:
        return frt.Id

    return None


# -- Coordinate projection ----------------------------------------------------

def get_view_transform(section_view):
    return (section_view.Origin,
            section_view.RightDirection,
            section_view.UpDirection)


def project_point(pt, origin, right, up):
    delta = pt - origin
    return DB.XYZ(delta.DotProduct(right), delta.DotProduct(up), 0)


def project_line(line, origin, right, up):
    p0 = project_point(line.GetEndPoint(0), origin, right, up)
    p1 = project_point(line.GetEndPoint(1), origin, right, up)
    if p0.DistanceTo(p1) < 1e-9:
        return None
    return DB.Line.CreateBound(p0, p1)


def project_arc(arc, origin, right, up):
    """
    Project an Arc to 2D. Returns list of curves.
    Handles both bounded arcs and unbound full circles.
    Full circles are split into two semicircular arcs because
    NewDetailCurve rejects unbound curves.
    """
    # Check if arc is unbound (full circle) -- this is the cause of
    # "The input curve is not bound" error on 64 elements
    if not arc.IsBound:
        return _project_full_circle(arc, origin, right, up)

    p_start = project_point(arc.GetEndPoint(0), origin, right, up)
    p_end   = project_point(arc.GetEndPoint(1), origin, right, up)

    param0 = arc.GetEndParameter(0)
    param1 = arc.GetEndParameter(1)
    mid_param = (param0 + param1) / 2.0
    mid_3d = arc.Evaluate(mid_param, False)
    p_mid = project_point(mid_3d, origin, right, up)

    # All three points coincide -- degenerate
    if (p_start.DistanceTo(p_end) < 1e-9
            and p_start.DistanceTo(p_mid) < 1e-9):
        return []

    # Try arc from 3 points
    try:
        v1 = p_mid - p_start
        v2 = p_end - p_start
        cross_z = v1.X * v2.Y - v1.Y * v2.X
        if abs(cross_z) > 1e-9:
            new_arc = DB.Arc.Create(p_start, p_end, p_mid)
            return [new_arc]
    except Exception:
        pass

    # Collinear or arc creation failed -- try as line
    if p_start.DistanceTo(p_end) > 1e-9:
        return [DB.Line.CreateBound(p_start, p_end)]

    # Fallback: multi-segment approximation
    segments = []
    num_segs = max(4, min(int(arc.Length * 10), 24))
    prev_pt = p_start
    for i in range(1, num_segs + 1):
        t = param0 + (param1 - param0) * (float(i) / num_segs)
        pt_3d = arc.Evaluate(t, False)
        pt_2d = project_point(pt_3d, origin, right, up)
        if prev_pt.DistanceTo(pt_2d) > 1e-9:
            try:
                segments.append(DB.Line.CreateBound(prev_pt, pt_2d))
                prev_pt = pt_2d
            except Exception:
                continue
    return segments


def _project_full_circle(arc, origin, right, up):
    """
    Handle unbound (full circle) arcs by splitting into two semicircles.
    Full circles have no distinct start/end points and NewDetailCurve
    rejects them. We sample 4 points at 0, PI/2, PI, 3PI/2 to create
    two bounded semicircular arcs.
    """
    # Get circle center and radius from the arc
    center_3d = arc.Center
    radius = arc.Radius

    # Sample 4 points around the circle using the arc's own coordinate system
    # Arc.Evaluate with normalized parameter: 0.0 = start, 1.0 = end
    # For unbound arcs, use raw parameter (angle in radians)
    try:
        # Sample at 0, PI/2, PI, 3*PI/2 radians
        p0_3d = arc.Evaluate(0, False)
        p1_3d = arc.Evaluate(math.pi / 2.0, False)
        p2_3d = arc.Evaluate(math.pi, False)
        p3_3d = arc.Evaluate(3.0 * math.pi / 2.0, False)
    except Exception:
        # If raw parameter fails, try normalized
        try:
            p0_3d = arc.Evaluate(0.0, True)
            p1_3d = arc.Evaluate(0.25, True)
            p2_3d = arc.Evaluate(0.5, True)
            p3_3d = arc.Evaluate(0.75, True)
        except Exception:
            return []

    # Project all 4 points to 2D
    p0 = project_point(p0_3d, origin, right, up)
    p1 = project_point(p1_3d, origin, right, up)
    p2 = project_point(p2_3d, origin, right, up)
    p3 = project_point(p3_3d, origin, right, up)

    result = []

    # First semicircle: p0 -> p2 through p1
    try:
        if (p0.DistanceTo(p1) > 1e-9 and p0.DistanceTo(p2) > 1e-9
                and p1.DistanceTo(p2) > 1e-9):
            # Check non-collinear
            v1 = p1 - p0
            v2 = p2 - p0
            cross_z = v1.X * v2.Y - v1.Y * v2.X
            if abs(cross_z) > 1e-9:
                arc1 = DB.Arc.Create(p0, p2, p1)
                result.append(arc1)
    except Exception:
        pass

    # Second semicircle: p2 -> p0 through p3
    try:
        if (p2.DistanceTo(p3) > 1e-9 and p2.DistanceTo(p0) > 1e-9
                and p3.DistanceTo(p0) > 1e-9):
            v1 = p3 - p2
            v2 = p0 - p2
            cross_z = v1.X * v2.Y - v1.Y * v2.X
            if abs(cross_z) > 1e-9:
                arc2 = DB.Arc.Create(p2, p0, p3)
                result.append(arc2)
    except Exception:
        pass

    if result:
        return result

    # Final fallback: tessellate the full circle into line segments
    segments = []
    num_segs = max(8, min(int(2.0 * math.pi * radius * 10), 36))
    points_2d = []
    for i in range(num_segs):
        angle = 2.0 * math.pi * float(i) / num_segs
        try:
            pt_3d = arc.Evaluate(angle, False)
            points_2d.append(project_point(pt_3d, origin, right, up))
        except Exception:
            continue

    if len(points_2d) >= 3:
        # Close the loop
        points_2d.append(points_2d[0])
        for i in range(len(points_2d) - 1):
            if points_2d[i].DistanceTo(points_2d[i + 1]) > 1e-9:
                try:
                    segments.append(
                        DB.Line.CreateBound(points_2d[i], points_2d[i + 1]))
                except Exception:
                    continue

    return segments


def project_curve(curve, origin, right, up):
    """Project any curve to 2D. Returns list of curves."""
    if isinstance(curve, DB.Line):
        result = project_line(curve, origin, right, up)
        return [result] if result else []

    if isinstance(curve, DB.Arc):
        return project_arc(curve, origin, right, up)

    # NurbSpline, HermiteSpline, Ellipse -- tessellate
    try:
        tess = curve.Tessellate()
        if tess:
            pts = [project_point(p, origin, right, up) for p in tess]
            segs = []
            for i in range(len(pts) - 1):
                if pts[i].DistanceTo(pts[i + 1]) > 1e-9:
                    try:
                        segs.append(DB.Line.CreateBound(pts[i], pts[i + 1]))
                    except Exception:
                        continue
            return segs
    except Exception:
        pass

    # Last resort
    p0 = project_point(curve.GetEndPoint(0), origin, right, up)
    p1 = project_point(curve.GetEndPoint(1), origin, right, up)
    if p0.DistanceTo(p1) > 1e-9:
        return [DB.Line.CreateBound(p0, p1)]
    return []


# -- Element Recreation --------------------------------------------------------

def recreate_detail_curve(el, dest_view, origin, right, up):
    """Recreate DetailLine / DetailArc / any CurveElement subclass.
    Handles unbound (full circle) curves by splitting into bounded segments."""
    geom_curve = el.GeometryCurve
    if geom_curve is None:
        return 0, 1, "No geometry curve"

    # For unbound curves that aren't Arc type, make them bound first
    if not geom_curve.IsBound and not isinstance(geom_curve, DB.Arc):
        try:
            geom_curve.MakeBound(0, geom_curve.Period if hasattr(geom_curve, 'Period') else 2.0 * math.pi)
        except Exception:
            pass  # project_curve will handle via tessellation

    projected = project_curve(geom_curve, origin, right, up)
    if not projected:
        return 0, 1, "Degenerate after projection"

    source_style = None
    try:
        source_style = el.LineStyle
    except Exception:
        pass

    created = 0
    failed  = 0
    for crv in projected:
        try:
            new_el = doc.Create.NewDetailCurve(dest_view, crv)
            if source_style is not None:
                try:
                    new_el.LineStyle = source_style
                except Exception:
                    pass
            created += 1
        except Exception as ex:
            failed += 1

    return created, failed, None


def recreate_detail_component_point(el, dest_view, origin, right, up):
    """Recreate a point-based detail component."""
    type_id = el.GetTypeId()
    if type_id == DB.ElementId.InvalidElementId:
        return 0, 1, "No valid type"

    loc = el.Location
    if not isinstance(loc, DB.LocationPoint):
        return 0, 1, "Expected LocationPoint"

    pt_2d = project_point(loc.Point, origin, right, up)

    try:
        new_inst = doc.Create.NewFamilyInstance(
            pt_2d, doc.GetElement(type_id), dest_view)
    except Exception as e:
        return 0, 1, "NewFamilyInstance: {}".format(str(e))

    # Rotation
    try:
        if hasattr(loc, 'Rotation') and abs(loc.Rotation) > 1e-6:
            axis = DB.Line.CreateBound(
                pt_2d, DB.XYZ(pt_2d.X, pt_2d.Y, 1))
            DB.ElementTransformUtils.RotateElement(
                doc, new_inst.Id, axis, loc.Rotation)
    except Exception:
        pass

    # Mirror
    try:
        if el.Mirrored:
            mirror_plane = DB.Plane.CreateByNormalAndOrigin(
                DB.XYZ(1, 0, 0), pt_2d)
            DB.ElementTransformUtils.MirrorElement(
                doc, new_inst.Id, mirror_plane)
    except Exception:
        pass

    return 1, 0, None


def recreate_detail_component_line(el, dest_view, origin, right, up):
    """Recreate a line-based detail component."""
    type_id = el.GetTypeId()
    if type_id == DB.ElementId.InvalidElementId:
        return 0, 1, "No valid type"

    loc = el.Location
    if not isinstance(loc, DB.LocationCurve):
        return 0, 1, "Expected LocationCurve"

    source_curve = loc.Curve
    projected = project_curve(source_curve, origin, right, up)
    if not projected:
        return 0, 1, "Degenerate curve"

    placement_line = projected[0]
    family_symbol = doc.GetElement(type_id)

    # Attempt 1: Curve-based placement (line-based families)
    try:
        new_inst = doc.Create.NewFamilyInstance(
            placement_line, family_symbol, dest_view)
        return 1, 0, None
    except Exception:
        pass

    # Attempt 2: Two-endpoint placement using reference line
    try:
        p0 = placement_line.GetEndPoint(0)
        p1 = placement_line.GetEndPoint(1)
        new_inst = doc.Create.NewFamilyInstance(
            DB.Line.CreateBound(p0, p1), family_symbol, dest_view)
        return 1, 0, None
    except Exception:
        pass

    # Attempt 3: Point-based fallback at midpoint
    try:
        p0 = placement_line.GetEndPoint(0)
        p1 = placement_line.GetEndPoint(1)
        mid = DB.XYZ((p0.X + p1.X) / 2.0,
                     (p0.Y + p1.Y) / 2.0, 0)
        new_inst = doc.Create.NewFamilyInstance(
            mid, family_symbol, dest_view)
        return 1, 0, "Placed at midpoint (line placement failed)"
    except Exception as e:
        return 0, 1, "All placement methods failed: {}".format(str(e))


def recreate_text_note(el, dest_view, origin, right, up):
    """Recreate a TextNote with content, type, position, rotation."""
    if not isinstance(el, DB.TextNote):
        return 0, 1, "Not a TextNote"

    text = el.Text
    if not text or text.strip() == "":
        text = " "

    type_id = el.GetTypeId()
    pt_2d = project_point(el.Coord, origin, right, up)

    try:
        new_note = DB.TextNote.Create(
            doc, dest_view.Id, pt_2d, text, type_id)
    except Exception as e:
        return 0, 1, "TextNote.Create: {}".format(str(e))

    # Rotation
    try:
        angle_param = el.get_Parameter(DB.BuiltInParameter.TEXT_ANGLE)
        if angle_param and angle_param.HasValue:
            angle = angle_param.AsDouble()
            if abs(angle) > 1e-6:
                new_angle = new_note.get_Parameter(
                    DB.BuiltInParameter.TEXT_ANGLE)
                if new_angle:
                    new_angle.Set(angle)
    except Exception:
        pass

    return 1, 0, None


def recreate_filled_region(el, dest_view, origin, right, up,
                           default_fr_type_id=None):
    """
    Recreate a FilledRegion. Validates type ID and falls back to
    document default if the source type is invalid in this context.
    """
    if not isinstance(el, DB.FilledRegion):
        return 0, 1, "Not a FilledRegion"

    # Validate type ID -- this was causing "The Id typeId is invalid"
    type_id = el.GetTypeId()
    type_valid = False
    if type_id and type_id != DB.ElementId.InvalidElementId:
        type_el = doc.GetElement(type_id)
        if type_el is not None and isinstance(type_el, DB.FilledRegionType):
            type_valid = True

    if not type_valid:
        if default_fr_type_id:
            type_id = default_fr_type_id
            logger.info("  FilledRegion type invalid, using default")
        else:
            return 0, 1, "Invalid type and no default available"

    try:
        boundary_loops = el.GetBoundaries()
    except Exception as e:
        return 0, 1, "GetBoundaries: {}".format(str(e))

    if not boundary_loops or boundary_loops.Count == 0:
        return 0, 1, "No boundary loops"

    new_loops = CsList[DB.CurveLoop]()

    for loop in boundary_loops:
        new_loop = DB.CurveLoop()
        valid = True

        for curve in loop:
            projected = project_curve(curve, origin, right, up)
            if not projected:
                valid = False
                break
            for crv in projected:
                try:
                    new_loop.Append(crv)
                except Exception:
                    valid = False
                    break
            if not valid:
                break

        if valid:
            try:
                # Validate loop has curves
                has_curves = False
                it = new_loop.GetCurveLoopIterator()
                if it.MoveNext():
                    has_curves = True
                if has_curves:
                    new_loops.Add(new_loop)
            except Exception:
                pass

    if new_loops.Count == 0:
        return 0, 1, "No valid loops after projection"

    try:
        new_region = DB.FilledRegion.Create(
            doc, type_id, dest_view.Id, new_loops)
        return 1, 0, None
    except Exception as e:
        # Second attempt with default type
        if default_fr_type_id and type_id != default_fr_type_id:
            try:
                new_region = DB.FilledRegion.Create(
                    doc, default_fr_type_id, dest_view.Id, new_loops)
                return 1, 0, "Used default filled region type"
            except Exception:
                pass
        return 0, 1, "FilledRegion.Create: {}".format(str(e))


# -- Dispatcher ----------------------------------------------------------------

def recreate_element(el, dest_view, origin, right, up,
                     default_fr_type_id=None):
    """
    Dispatch element to handler by .NET type.
    Returns (created_count, failed_count, reason_or_None).
    """
    # FilledRegion (may appear under "Detail Items" category)
    if isinstance(el, DB.FilledRegion):
        return recreate_filled_region(
            el, dest_view, origin, right, up, default_fr_type_id)

    # CurveElement (DetailLine, DetailArc, DetailNurbSpline, etc.)
    if isinstance(el, DB.CurveElement):
        return recreate_detail_curve(el, dest_view, origin, right, up)

    # TextNote
    if isinstance(el, DB.TextNote):
        return recreate_text_note(el, dest_view, origin, right, up)

    # FamilyInstance (detail components)
    if isinstance(el, DB.FamilyInstance):
        loc = el.Location
        if isinstance(loc, DB.LocationCurve):
            return recreate_detail_component_line(
                el, dest_view, origin, right, up)
        elif isinstance(loc, DB.LocationPoint):
            return recreate_detail_component_point(
                el, dest_view, origin, right, up)
        else:
            return 0, 1, "FamilyInstance: unsupported location"

    cat_name = el.Category.Name if el.Category else "No Category"
    return 0, 1, "No handler: {} ({})".format(
        type(el).__name__, cat_name)


def recreate_group_members(group_el, dest_view, origin, right, up,
                           default_fr_type_id=None):
    """
    Recreate individual members of a Detail Group (ungrouped).
    Returns (created_count, failed_count, failure_details_dict).
    """
    if not isinstance(group_el, DB.Group):
        return 0, 1, {}

    member_ids = group_el.GetMemberIds()
    if not member_ids:
        return 0, 0, {}

    total_created = 0
    total_failed  = 0
    failure_details = {}  # {reason: count}

    for mid in member_ids:
        member = doc.GetElement(mid)
        if member is None:
            total_failed += 1
            failure_details["Null element"] = \
                failure_details.get("Null element", 0) + 1
            continue

        c, f, reason = recreate_element(
            member, dest_view, origin, right, up, default_fr_type_id)
        total_created += c
        total_failed  += f
        if reason and f > 0:
            failure_details[reason] = failure_details.get(reason, 0) + 1

    return total_created, total_failed, failure_details


# -- Collection ----------------------------------------------------------------

def collect_elements(section_view):
    all_in_view = DB.FilteredElementCollector(doc, section_view.Id)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

    to_recreate = []
    groups      = []
    skipped     = []

    for el in all_in_view:
        if not el.ViewSpecific:
            continue

        cat_int = get_cat_int(el)

        if cat_int is not None and cat_int in SKIP_CATEGORIES:
            cat_name = el.Category.Name if el.Category else "Unknown"
            skipped.append((el, cat_name, "References model geometry"))
            continue

        if el.Category is None:
            skipped.append((el, "No Category", "Internal reference"))
            continue

        if isinstance(el, DB.Group):
            groups.append(el)
            continue

        if el.GroupId != DB.ElementId.InvalidElementId:
            continue

        to_recreate.append(el)

    return to_recreate, groups, skipped


# -- View selection ------------------------------------------------------------

def get_section_views():
    selection = revit.get_selection()

    if selection.elements:
        sections = [el for el in selection.elements
                    if isinstance(el, DB.View)
                    and el.ViewType == DB.ViewType.Section
                    and not el.IsTemplate]
        if sections:
            return sections

    all_views = DB.FilteredElementCollector(doc)\
                  .OfClass(DB.View).ToElements()
    candidates = [v for v in all_views
                  if v.ViewType == DB.ViewType.Section
                  and not v.IsTemplate]

    if not candidates:
        forms.alert("No section views found.", exitscript=True)

    name_map = {v.Name: v for v in candidates}
    chosen = forms.SelectFromList.show(
        sorted(name_map.keys()),
        title="Select Section Views to Convert",
        button_name="Convert to Drafting Views",
        multiselect=True
    )
    if not chosen:
        script.exit()

    return [name_map[n] for n in chosen]


# -- Main ----------------------------------------------------------------------

def main():
    section_views = get_section_views()

    drafting_vft = get_drafting_view_family_type()
    if drafting_vft is None:
        forms.alert("No Drafting View Family Type found.", exitscript=True)

    # Get default filled region type for fallback
    default_fr_type = get_default_filled_region_type_id()

    all_views      = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    existing_names = set(v.Name for v in all_views)

    # Confirmation
    msg = "Convert {} section view(s) to drafting views?\n\n".format(
        len(section_views))
    msg += "\n".join("  - {}".format(v.Name)
                     for v in section_views[:10])
    if len(section_views) > 10:
        msg += "\n  ... and {} more".format(len(section_views) - 10)
    msg += "\n\nElements will be recreated (not copied)."
    msg += "\nDetail Groups will be ungrouped."
    msg += "\nDimensions, Spots, Tags require manual placement."

    if not forms.alert(msg, yes=True, no=True, title="Confirm Conversion"):
        script.exit()

    results      = []
    failed_views = []

    with DB.TransactionGroup(doc, "Section to Drafting Conversion") as tg:
        tg.Start()

        for section in section_views:
            section_name = section.Name
            output.print_md("### Processing: {}".format(section_name))

            try:
                # Step 1: Collect
                elements, groups, skipped = collect_elements(section)

                total_source = len(elements) + len(groups) + len(skipped)
                output.print_md(
                    "  Total: {} | Recreate: {} | Groups: {} | Skip: {}".format(
                        total_source, len(elements),
                        len(groups), len(skipped)))

                # Log skips
                skip_summary = {}
                for _, cat_name, reason in skipped:
                    key = "{} ({})".format(cat_name, reason)
                    skip_summary[key] = skip_summary.get(key, 0) + 1
                for key, count in skip_summary.iteritems():
                    output.print_md("  Skip: {} x {}".format(count, key))

                # Log .NET types
                type_counts = {}
                for el in elements:
                    tn = type(el).__name__
                    type_counts[tn] = type_counts.get(tn, 0) + 1
                for tn, count in type_counts.iteritems():
                    output.print_md("  Type: {} x {}".format(count, tn))

                # Count group members
                group_member_count = 0
                group_member_types = {}
                for grp in groups:
                    try:
                        mids = grp.GetMemberIds()
                        if mids:
                            for mid in mids:
                                m = doc.GetElement(mid)
                                if m:
                                    group_member_count += 1
                                    tn = type(m).__name__
                                    group_member_types[tn] = \
                                        group_member_types.get(tn, 0) + 1
                    except Exception:
                        pass
                if group_member_count:
                    parts = ["{} x {}".format(c, t)
                             for t, c in group_member_types.iteritems()]
                    output.print_md(
                        "  Group members: {} ({})".format(
                            group_member_count, ", ".join(parts)))

                # Step 2: Create drafting view
                new_view = None
                with DB.Transaction(doc, "Create Drafting: {}".format(
                        section_name)) as t:
                    t.Start()
                    safe_name = get_unique_view_name(
                        section_name + DRAFTING_SUFFIX, existing_names)
                    new_view = DB.ViewDrafting.Create(doc, drafting_vft.Id)
                    new_view.Name = safe_name
                    try:
                        new_view.Scale = section.Scale
                    except Exception:
                        pass
                    existing_names.add(safe_name)
                    t.Commit()

                if new_view is None:
                    raise Exception("Failed to create drafting view")

                # Step 3: View transform
                origin, right_dir, up_dir = get_view_transform(section)

                # Step 4: Recreate elements
                total_created = 0
                total_failed  = 0
                all_failure_reasons = {}

                if elements or groups:
                    with DB.Transaction(doc, "Recreate: {}".format(
                            section_name)) as t:
                        t.Start()

                        # Individual elements
                        for el in elements:
                            try:
                                c, f, reason = recreate_element(
                                    el, new_view,
                                    origin, right_dir, up_dir,
                                    default_fr_type)
                                total_created += c
                                total_failed  += f
                                if reason and f > 0:
                                    all_failure_reasons[reason] = \
                                        all_failure_reasons.get(reason, 0) + 1
                            except Exception as ex:
                                total_failed += 1
                                r = str(ex)
                                all_failure_reasons[r] = \
                                    all_failure_reasons.get(r, 0) + 1

                        # Group members
                        for grp in groups:
                            try:
                                c, f, details = recreate_group_members(
                                    grp, new_view,
                                    origin, right_dir, up_dir,
                                    default_fr_type)
                                total_created += c
                                total_failed  += f
                                for r, cnt in details.iteritems():
                                    all_failure_reasons[r] = \
                                        all_failure_reasons.get(r, 0) + cnt
                            except Exception as ex:
                                total_failed += 1

                        t.Commit()

                # Log failure reasons (visible in output)
                if all_failure_reasons:
                    output.print_md("  **Failure breakdown:**")
                    for reason, count in sorted(
                            all_failure_reasons.iteritems(),
                            key=lambda x: -x[1]):
                        output.print_md(
                            "    {} x {}".format(count, reason))

                skip_count = len(skipped)
                results.append({
                    "section":  section_name,
                    "drafting": new_view.Name,
                    "created":  total_created,
                    "failed":   total_failed,
                    "skipped":  skip_count,
                })

                output.print_md(
                    "  **OK** -> **{}** | Created: {} | Failed: {} "
                    "| Skipped: {}".format(
                        new_view.Name, total_created,
                        total_failed, skip_count))

            except Exception as e:
                logger.error("Failed: '{}': {}".format(
                    section_name, str(e)))
                output.print_md("  **FAILED**: {}".format(str(e)))
                failed_views.append(section_name)

        tg.Assimilate()

    # --- Summary ---
    output.print_md("---")
    output.print_md("## Conversion Complete")
    output.print_md("**Views processed:** {}".format(len(section_views)))
    output.print_md("**Successful:** {}".format(len(results)))
    output.print_md("**Failed:** {}".format(len(failed_views)))

    if results:
        output.print_table(
            table_data=[[r["section"], r["drafting"],
                         r["created"], r["failed"], r["skipped"]]
                        for r in results],
            columns=["Source Section", "New Drafting View",
                     "Created", "Failed", "Skipped"]
        )

    if failed_views:
        output.print_md("**Failed views:**")
        for name in failed_views:
            output.print_md("  - {}".format(name))

    if results:
        total_c = sum(r["created"] for r in results)
        total_f = sum(r["failed"] for r in results)
        forms.alert(
            "{} drafting view(s) created.\n\n"
            "Created: {} elements\n"
            "Failed: {} elements\n\n"
            "Detail Groups were ungrouped.\n"
            "Dimensions/Tags require manual placement.\n"
            "See pyRevit output for details.".format(
                len(results), total_c, total_f),
            title="Conversion Complete")


main()
